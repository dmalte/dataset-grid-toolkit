import io
import json
import re
import sys
import tempfile
from contextlib import redirect_stdout
from dataclasses import replace
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable, TextIO

try:
    from item_manager.core import (
        get_item_config, atomic_write_text,
        ActionItem, BuyItem, ConsumeItem, IdeaItem,
        PRIORITY_EMOJIS, extract_tags, parse_line, render_item,
        ItemRepository,
    )
except ImportError:
    from ._vendor import (
        get_item_config, atomic_write_text,
        ActionItem, BuyItem, ConsumeItem, IdeaItem,
        PRIORITY_EMOJIS, extract_tags, parse_line, render_item,
        ItemRepository,
    )
from .export_grid import FIELD_SCHEMA, _compute_baseline_hash, build_export


CHANGE_VERSION = "1"
SUPPORTED_ACTIONS = frozenset({"update", "create", "delete"})
SUPPORTED_STATUS_VALUES = frozenset({"open", "done", "new", "ongoing", "waiting"})
SUPPORTED_PRIORITY_VALUES = frozenset({"highest", "high", "medium", "low", "lowest"})
SUPPORTED_URGENCY_VALUES = frozenset({"highest", "high", "medium", "low"})
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
EDITABLE_FIELDS = frozenset(field_name for field_name, meta in FIELD_SCHEMA.items() if meta.get("editable"))
ACTION_ITEM_TYPE = "action"
BUY_ITEM_TYPE = "buy"
CONSUME_ITEM_TYPE = "consume"
IDEA_ITEM_TYPE = "idea"
ITEM_TYPE_PREFIXES = {
    "ai-": ACTION_ITEM_TYPE,
    "bi-": BUY_ITEM_TYPE,
    "ci-": CONSUME_ITEM_TYPE,
    "ii-": IDEA_ITEM_TYPE,
}
ITEM_TYPE_FIELDS = {
    ACTION_ITEM_TYPE: frozenset({"text", "status", "created", "completed", "due", "area", "priority", "urgency", "tags"}),
    BUY_ITEM_TYPE: frozenset({"text", "status", "created", "completed", "area", "priority", "urgency", "shop", "tags"}),
    CONSUME_ITEM_TYPE: frozenset({"text", "status", "created", "completed", "area", "priority", "urgency", "tags"}),
    IDEA_ITEM_TYPE: frozenset({"text", "created", "area", "priority", "urgency", "tags"}),
}


def _load_dataset(input_path: str) -> dict:
    path = Path(input_path)
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError("dataset root must be an object")
    return payload


def _item_id_from_change(change_row: dict) -> str | None:
    target = change_row.get("target") or {}
    proposed = change_row.get("proposed") or {}
    baseline = change_row.get("baseline") or {}
    return target.get("itemId") or target.get("id") or proposed.get("id") or baseline.get("id")


def _infer_item_type(item_id: str | None, row_data: dict | None = None) -> str | None:
    if row_data:
        item_type = row_data.get("itemType")
        if item_type:
            return item_type
    if not item_id:
        return None
    lowered = item_id.lower()
    for prefix, item_type in ITEM_TYPE_PREFIXES.items():
        if lowered.startswith(prefix):
            return item_type
    return None


def _normalize_field_value(field_name: str, value):
    if isinstance(value, str):
        stripped = value.strip()
        if field_name == "text":
            return stripped
        return stripped or None
    return value


def _validate_date(field_name: str, value) -> None:
    if value is None:
        return
    if not isinstance(value, str) or not DATE_RE.match(value):
        raise ValueError(f"{field_name} must be YYYY-MM-DD")


def _validate_priority(value) -> None:
    if value is None:
        return
    if value not in SUPPORTED_PRIORITY_VALUES:
        raise ValueError("priority must be one of highest, high, medium, low, lowest")


def _validate_urgency(value) -> None:
    if value is None:
        return
    if value not in SUPPORTED_URGENCY_VALUES:
        raise ValueError("urgency must be one of highest, high, medium, low")


def _validate_status(item_type: str, value) -> None:
    if value is None:
        return
    if item_type == IDEA_ITEM_TYPE:
        raise ValueError("idea items do not support status")
    if value not in SUPPORTED_STATUS_VALUES:
        raise ValueError(f"status must be one of {', '.join(sorted(SUPPORTED_STATUS_VALUES))}")


def _normalize_proposed_fields(item_type: str, proposed: dict, allow_create_fields: bool = False) -> dict:
    if not isinstance(proposed, dict):
        raise ValueError("proposed must be an object")

    allowed_fields = ITEM_TYPE_FIELDS[item_type]
    normalized = {}
    field_tags = {}
    for field_name, raw_value in proposed.items():
        if field_name in {"id", "itemType"}:
            if not allow_create_fields:
                raise ValueError(f"{field_name} cannot be modified")
            normalized[field_name] = raw_value
            continue

        # Known editable field
        if field_name in EDITABLE_FIELDS:
            if field_name not in allowed_fields:
                raise ValueError(f"field is not supported for {item_type}: {field_name}")

            if field_name == "tags":
                if not isinstance(raw_value, list):
                    raise ValueError("tags must be a list")
                normalized_tags = []
                for tag in raw_value:
                    if not isinstance(tag, str):
                        raise ValueError("each tag must be a string")
                    clean = tag.strip().lstrip("#")
                    if clean:
                        normalized_tags.append(clean)
                normalized[field_name] = normalized_tags
                continue

            normalized_value = _normalize_field_value(field_name, raw_value)
            if field_name == "text":
                if not normalized_value:
                    raise ValueError("text must not be empty")
            elif field_name in {"created", "completed", "due"}:
                _validate_date(field_name, normalized_value)
            elif field_name == "priority":
                _validate_priority(normalized_value)
            elif field_name == "urgency":
                _validate_urgency(normalized_value)
            elif field_name == "status":
                _validate_status(item_type, normalized_value)

            normalized[field_name] = normalized_value
            continue

        # Non-editable base field — reject
        if field_name in FIELD_SCHEMA:
            raise ValueError(f"field is not editable: {field_name}")

        # Dynamic field-tag (e.g. project, person, category)
        if isinstance(raw_value, str):
            stripped = raw_value.strip()
            field_tags[field_name] = stripped or None
        elif raw_value is None:
            field_tags[field_name] = None
        else:
            raise ValueError(f"field-tag {field_name} must be a string or null")

    if field_tags:
        normalized["_field_tags"] = field_tags
    return normalized


def _baseline_index(dataset: dict) -> dict:
    data_rows = dataset.get("data") or []
    if not isinstance(data_rows, list):
        raise ValueError("data must be an array")
    index = {}
    for row in data_rows:
        if not isinstance(row, dict):
            continue
        item_id = row.get("id")
        if not item_id:
            continue
        index[item_id] = row
    return index


def _load_repo(cfg: dict, item_type: str) -> ItemRepository:
    return ItemRepository(cfg["vault"], cfg["search_paths"], item_type)


def _find_item_by_id(cfg: dict, item_type: str, item_id: str):
    repo = _load_repo(cfg, item_type)
    found = repo.find_by_id(item_id)
    if not found:
        return repo, None
    if len(found) > 1:
        raise ValueError(f"duplicate id found in vault: {item_id}")
    return repo, found[0]


def _current_baseline_hash(item, item_type: str) -> str:
    row = item.to_dict()
    row["itemType"] = item_type
    row["tags"] = extract_tags(item.raw, getattr(item, "shop", None))
    field_tags = row.pop("field_tags", {}) or {}
    for prefix, value in field_tags.items():
        row[prefix] = value
    return _compute_baseline_hash(row)


def _resolve_baseline_hash(baseline_row: dict) -> str | None:
    sync_block = baseline_row.get("sync") if isinstance(baseline_row, dict) else None
    if not isinstance(sync_block, dict):
        return None
    return sync_block.get("baselineHash")


def _ensure_no_duplicate_targets(change_rows: list[dict]):
    seen = set()
    duplicates = set()
    for change_row in change_rows:
        item_id = _item_id_from_change(change_row)
        if not item_id:
            continue
        if item_id in seen:
            duplicates.add(item_id)
        seen.add(item_id)
    return duplicates


def _build_change_result(change_row: dict, item_id: str | None, action: str, status: str, message: str, extra: dict | None = None):
    result = {
        "changeId": change_row.get("changeId"),
        "itemId": item_id,
        "action": action,
        "status": status,
        "message": message,
    }
    if extra:
        result.update(extra)
    return result


def _coerce_item_for_update(item, item_type: str, proposed_fields: dict):
    updated_item = replace(item)
    updated_item.tags = item.tags  # preserve (replace loses init=False fields)
    updated_item.field_tags = dict(item.field_tags)
    incoming_field_tags = proposed_fields.pop("_field_tags", None)
    for field_name, value in proposed_fields.items():
        if field_name == "tags":
            updated_item.tags = value
            continue
        setattr(updated_item, field_name, value)
    if incoming_field_tags:
        for key, value in incoming_field_tags.items():
            if value is None:
                updated_item.field_tags.pop(key, None)
            else:
                updated_item.field_tags[key] = value
    if item_type != IDEA_ITEM_TYPE and "status" in proposed_fields:
        if updated_item.status == "done" and "completed" not in proposed_fields and not getattr(updated_item, "completed", None):
            updated_item.completed = date.today().isoformat()
        if updated_item.status != "done" and "completed" not in proposed_fields:
            updated_item.completed = None
    return updated_item


def _find_and_update_line(path: Path, item_id: str, updater, item_type: str):
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines(True)
    file_rel = path.name
    try:
        file_rel = str(path.relative_to(Path(path.anchor)))
    except ValueError:
        pass
    for index, line in enumerate(lines):
        item = parse_line(line, file_rel, index + 1, item_type)
        if item and item.id == item_id:
            updated = updater(item)
            lines[index] = render_item(updated, item_type)
            return lines, updated
    return None, None


def _insert_under_date_heading(lines, line: str, created: str) -> None:
    created_date = datetime.fromisoformat(created)
    day_name = created_date.strftime("%A")
    heading = f"## {created}: {day_name}\n"
    heading_found = False
    insertion_point = None
    for index, existing in enumerate(lines):
        stripped = existing.strip()
        if stripped.startswith(f"## {created}:") or stripped.startswith(f"{created}:"):
            heading_found = True
            insertion_point = index + 1
            for scan_index in range(index + 1, len(lines)):
                next_line = lines[scan_index].strip()
                if next_line.startswith("##") or (next_line and ":" in next_line and len(next_line.split(":")[0]) == 10):
                    insertion_point = scan_index
                    break
                if next_line:
                    insertion_point = scan_index + 1
            break
    if not heading_found:
        if lines and not lines[-1].endswith("\n"):
            lines.append("\n")
        lines.append(heading)
        insertion_point = len(lines)
    if insertion_point is None or insertion_point >= len(lines):
        lines.append(line)
    else:
        lines.insert(insertion_point, line)


def _insert_action_item(target: Path, item, cfg: dict):
    lines = target.read_text(encoding="utf-8").splitlines(True) if target.exists() else ["# Inbox\n"]
    insertion_point = None
    insertion_indent = ""
    area = item.area
    found_matches = 0

    if area:
        normalized_area = area.strip().lower()
        for index, line in enumerate(lines):
            stripped = line.strip()
            if stripped.startswith("- #") or stripped.startswith("#"):
                bullet_area = stripped[3:].strip().lower() if stripped.startswith("- #") else stripped[1:].strip().lower()
                line_indent = len(line) - len(line.lstrip())
                if bullet_area == normalized_area:
                    found_matches += 1
                    if found_matches == 1:
                        base_indent = line_indent
                        insertion_point = index + 1
                        insertion_indent = " " * (base_indent + 1)
                        scan_index = index + 1
                        while scan_index < len(lines):
                            next_line = lines[scan_index]
                            next_stripped = next_line.strip()
                            if not next_stripped:
                                scan_index += 1
                                continue
                            next_indent = len(next_line) - len(next_line.lstrip())
                            if next_indent <= base_indent:
                                insertion_point = scan_index
                                break
                            insertion_point = scan_index + 1
                            scan_index += 1
                    else:
                        break
        if found_matches == 0:
            area = None

    if not area or insertion_point is None:
        for index, line in enumerate(lines):
            if line.strip().lower() in ("# inbox", "inbox"):
                insertion_point = index + 1
                insertion_indent = ""
                break
        else:
            lines.insert(0, "# Inbox\n")
            insertion_point = 1
            insertion_indent = ""

    rendered_item = replace(item, indent=len(insertion_indent))
    rendered_item.tags = item.tags
    rendered_item.field_tags = dict(item.field_tags)
    rendered_line = render_item(rendered_item, ACTION_ITEM_TYPE)
    formatted_line = insertion_indent + rendered_line.lstrip()

    if insertion_point >= len(lines):
        if lines and not lines[-1].endswith("\n"):
            lines[-1] += "\n"
        lines.append(formatted_line)
    else:
        lines.insert(insertion_point, formatted_line)

    atomic_write_text(target, "".join(lines))
    return insertion_point + 1, formatted_line.rstrip("\n")


def _insert_dated_item(target: Path, item, item_type: str):
    lines = target.read_text(encoding="utf-8").splitlines(True) if target.exists() else []
    replaced = replace(item, indent=0)
    replaced.tags = item.tags
    replaced.field_tags = dict(item.field_tags)
    rendered_line = render_item(replaced, item_type)
    created = item.created or date.today().isoformat()
    _insert_under_date_heading(lines, rendered_line, created)
    line_index = lines.index(rendered_line) + 1
    atomic_write_text(target, "".join(lines))
    return line_index, rendered_line.rstrip("\n")


def _build_new_item(item_type: str, item_id: str, proposed_fields: dict, file_path: str):
    text = proposed_fields.get("text")
    if not text:
        raise ValueError("create requires text")
    created = proposed_fields.get("created") or date.today().isoformat()
    priority = proposed_fields.get("priority")
    area = proposed_fields.get("area")
    incoming_field_tags = proposed_fields.get("_field_tags") or {}
    common_fields = {
        "id": item_id,
        "text": text,
        "created": created,
        "area": area,
        "priority": priority,
        "file": file_path,
        "line_no": 0,
        "indent": 0,
        "raw": "",
        "managed": True,
    }
    tags = proposed_fields.get("tags")
    if item_type == ACTION_ITEM_TYPE:
        status = proposed_fields.get("status") or "open"
        completed = proposed_fields.get("completed")
        if status == "done" and not completed:
            completed = date.today().isoformat()
        if status != "done":
            completed = None
        item = ActionItem(status=status, due=proposed_fields.get("due"), completed=completed, **common_fields)
        item.field_tags = incoming_field_tags
        item.tags = tags
        return item
    if item_type == BUY_ITEM_TYPE:
        status = proposed_fields.get("status") or "open"
        completed = proposed_fields.get("completed")
        if status == "done" and not completed:
            completed = date.today().isoformat()
        if status != "done":
            completed = None
        item = BuyItem(status=status, completed=completed, shop=proposed_fields.get("shop"), **common_fields)
        item.field_tags = incoming_field_tags
        item.tags = tags
        return item
    if item_type == CONSUME_ITEM_TYPE:
        status = proposed_fields.get("status") or "open"
        completed = proposed_fields.get("completed")
        if status == "done" and not completed:
            completed = date.today().isoformat()
        if status != "done":
            completed = None
        item = ConsumeItem(status=status, completed=completed, **common_fields)
        item.field_tags = incoming_field_tags
        item.tags = tags
        return item
    item = IdeaItem(**common_fields)
    item.field_tags = incoming_field_tags
    item.tags = tags
    return item


def _ensure_target_file(cfg: dict, item_type: str) -> Path:
    item_cfg = get_item_config(cfg, item_type)
    target = Path(cfg["vault"]) / item_cfg["creation"]["file"]
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        if item_type == ACTION_ITEM_TYPE:
            target.write_text("# Inbox\n", encoding="utf-8")
        else:
            target.write_text("", encoding="utf-8")
    return target


def _apply_update(cfg: dict, item_type: str, item_id: str, baseline_row: dict, proposed: dict, dry_run: bool):
    repo, found = _find_item_by_id(cfg, item_type, item_id)
    if not found:
        raise ValueError("item not found in vault")
    item, path, _lines = found
    exported_hash = _resolve_baseline_hash(baseline_row)
    if not exported_hash:
        raise ValueError("baseline row is missing sync.baselineHash")
    current_hash = _current_baseline_hash(item, item_type)
    if current_hash != exported_hash:
        return None, "conflict", f"baseline hash mismatch (current {current_hash})"
    updated_item = _coerce_item_for_update(item, item_type, proposed)
    if dry_run:
        return updated_item, "dry-run", "validated update"
    new_lines, _ = _find_and_update_line(path, item_id, lambda _item: updated_item, item_type)
    if new_lines is None:
        raise ValueError("line not found during update")
    repo.write_file(path, new_lines)
    return updated_item, "applied", f"updated {path}"


def _apply_delete(cfg: dict, item_type: str, item_id: str, baseline_row: dict, dry_run: bool):
    repo, found = _find_item_by_id(cfg, item_type, item_id)
    if not found:
        raise ValueError("item not found in vault")
    item, path, _lines = found
    exported_hash = _resolve_baseline_hash(baseline_row)
    if not exported_hash:
        raise ValueError("baseline row is missing sync.baselineHash")
    current_hash = _current_baseline_hash(item, item_type)
    if current_hash != exported_hash:
        return "conflict", f"baseline hash mismatch (current {current_hash})"
    if dry_run:
        return "dry-run", "validated delete"
    lines = path.read_text(encoding="utf-8").splitlines(True)
    del lines[item.line_no - 1]
    if lines and not lines[-1].endswith("\n"):
        lines[-1] += "\n"
    repo.write_file(path, lines)
    return "applied", f"deleted from {path}"


def _apply_create(cfg: dict, item_type: str, item_id: str, proposed: dict, dry_run: bool):
    repo, found = _find_item_by_id(cfg, item_type, item_id)
    if found:
        return None, "conflict", "item id already exists in vault"
    target = _ensure_target_file(cfg, item_type)
    relative_target = str(target.relative_to(Path(cfg["vault"])))
    new_item = _build_new_item(item_type, item_id, proposed, relative_target)
    if dry_run:
        return new_item, "dry-run", f"validated create in {target}"
    if item_type == ACTION_ITEM_TYPE:
        line_no, raw_line = _insert_action_item(target, new_item, cfg)
    elif item_type == BUY_ITEM_TYPE:
        line_no, raw_line = _insert_dated_item(target, new_item, BUY_ITEM_TYPE)
    elif item_type == CONSUME_ITEM_TYPE:
        line_no, raw_line = _insert_dated_item(target, new_item, CONSUME_ITEM_TYPE)
    else:
        line_no, raw_line = _insert_dated_item(target, new_item, IDEA_ITEM_TYPE)
    parsed_item = parse_line(raw_line + "\n", relative_target, line_no, item_type)
    if not parsed_item:
        raise ValueError("created line could not be parsed")
    return parsed_item, "applied", f"created in {target}"


def import_grid_changes(input_path: str, cfg: dict, *, dry_run: bool = False, report_path: str | None = None, reexport_path: str | None = None) -> dict:
    dataset = _load_dataset(input_path)
    changes = dataset.get("changes") or {}
    change_rows = changes.get("rows") or []
    if not isinstance(change_rows, list):
        raise ValueError("changes.rows must be an array")
    if changes.get("version", CHANGE_VERSION) != CHANGE_VERSION:
        raise ValueError(f"unsupported change version: {changes.get('version')}")

    duplicate_targets = _ensure_no_duplicate_targets(change_rows)
    baseline_rows = _baseline_index(dataset)
    results = []
    counts = {"total": len(change_rows), "applied": 0, "skippedConflict": 0, "skippedInvalid": 0, "failed": 0, "dryRunValidated": 0}

    for change_row in change_rows:
        action = str(change_row.get("action") or "update").strip().lower()
        item_id = _item_id_from_change(change_row)
        try:
            if action not in SUPPORTED_ACTIONS:
                raise ValueError(f"unsupported action: {action}")
            if not item_id:
                raise ValueError("missing target item id")
            if item_id in duplicate_targets:
                raise ValueError(f"duplicate target item id in change set: {item_id}")

            baseline_row = baseline_rows.get(item_id)
            row_data = baseline_row or (change_row.get("proposed") if isinstance(change_row.get("proposed"), dict) else None)
            item_type = _infer_item_type(item_id, row_data)
            if item_type not in ITEM_TYPE_FIELDS:
                raise ValueError(f"could not infer item type for {item_id}")
            expected_item_type = _infer_item_type(item_id)
            if expected_item_type and item_type != expected_item_type:
                raise ValueError(f"item type does not match id prefix for {item_id}")

            proposed = change_row.get("proposed") or {}
            normalized_proposed = _normalize_proposed_fields(item_type, proposed, allow_create_fields=(action == "create"))

            if action == "create":
                proposed_id = normalized_proposed.get("id")
                if proposed_id and proposed_id != item_id:
                    raise ValueError("proposed.id must match target item id")
                proposed_item_type = normalized_proposed.get("itemType")
                if proposed_item_type and proposed_item_type != item_type:
                    raise ValueError("proposed.itemType must match target item type")

            if action in {"update", "delete"} and not baseline_row:
                raise ValueError("baseline row is required for update/delete")

            if action == "update":
                if not isinstance(baseline_row, dict):
                    raise ValueError("baseline row is required for update/delete")
                updated_item, result_status, message = _apply_update(cfg, item_type, item_id, baseline_row, normalized_proposed, dry_run)
                if result_status == "conflict":
                    counts["skippedConflict"] += 1
                    results.append(_build_change_result(change_row, item_id, action, "skipped-conflict", message))
                    continue
                if result_status == "dry-run":
                    if updated_item is None:
                        raise ValueError("update produced no item during dry-run")
                    counts["dryRunValidated"] += 1
                    results.append(_build_change_result(change_row, item_id, action, "dry-run", message, {"item": updated_item.to_dict()}))
                    continue
                if updated_item is None:
                    raise ValueError("update produced no item during apply")
                counts["applied"] += 1
                results.append(_build_change_result(change_row, item_id, action, "applied", message, {"item": updated_item.to_dict()}))
                continue

            if action == "delete":
                if not isinstance(baseline_row, dict):
                    raise ValueError("baseline row is required for update/delete")
                result_status, message = _apply_delete(cfg, item_type, item_id, baseline_row, dry_run)
                if result_status == "conflict":
                    counts["skippedConflict"] += 1
                    results.append(_build_change_result(change_row, item_id, action, "skipped-conflict", message))
                    continue
                if result_status == "dry-run":
                    counts["dryRunValidated"] += 1
                    results.append(_build_change_result(change_row, item_id, action, "dry-run", message))
                    continue
                counts["applied"] += 1
                results.append(_build_change_result(change_row, item_id, action, "applied", message))
                continue

            created_item, result_status, message = _apply_create(cfg, item_type, item_id, normalized_proposed, dry_run)
            if result_status == "conflict":
                counts["skippedConflict"] += 1
                results.append(_build_change_result(change_row, item_id, action, "skipped-conflict", message))
                continue
            if result_status == "dry-run":
                if created_item is None:
                    raise ValueError("create produced no item during dry-run")
                counts["dryRunValidated"] += 1
                results.append(_build_change_result(change_row, item_id, action, "dry-run", message, {"item": created_item.to_dict()}))
                continue
            if created_item is None:
                raise ValueError("create produced no item during apply")
            counts["applied"] += 1
            results.append(_build_change_result(change_row, item_id, action, "applied", message, {"item": created_item.to_dict()}))
        except ValueError as exc:
            counts["skippedInvalid"] += 1
            results.append(_build_change_result(change_row, item_id, action, "skipped-invalid", str(exc)))
        except Exception as exc:
            counts["failed"] += 1
            results.append(_build_change_result(change_row, item_id, action, "failed", str(exc)))

    report = {
        "ok": counts["skippedConflict"] == 0 and counts["skippedInvalid"] == 0 and counts["failed"] == 0,
        "command": "import-grid",
        "input": str(Path(input_path)),
        "dryRun": bool(dry_run),
        "processedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "counts": counts,
        "results": results,
    }

    if report_path:
        atomic_write_text(Path(report_path), json.dumps(report, ensure_ascii=False, indent=2) + "\n")

    if reexport_path and not dry_run and counts["failed"] == 0:
        with io.StringIO() as export_stdout, redirect_stdout(export_stdout):
            build_export(cfg, output_path=reexport_path, vault_root_hint=(dataset.get("sync") or {}).get("vaultRootHint"))
        report["reexportedTo"] = str(Path(reexport_path))

    return report


def import_grid_changes_confirm_cli(
    input_path: str,
    cfg: dict,
    *,
    reexport_path: str | None = None,
    input_fn: Callable[[str], str] = input,
    output_stream: TextIO | None = None,
) -> dict:
    """Interactive CLI review: dry-run, prompt per change, apply approved."""
    dataset = _load_dataset(input_path)
    changes = dataset.get("changes") or {}
    change_rows = changes.get("rows") or []
    stream = output_stream if output_stream is not None else sys.stdout

    dry_report = import_grid_changes(input_path, cfg, dry_run=True)
    results = dry_report.get("results") or []
    counts = dry_report.get("counts") or {}

    print("Obsidian Review", file=stream)
    print("=" * 40, file=stream)
    print(f"Dataset: {input_path}", file=stream)
    print(f"Total changes: {counts.get('total', 0)}", file=stream)
    print("", file=stream)

    if not results:
        print("No pending changes found.", file=stream)
        # Still re-export if requested (gives a fresh vault snapshot)
        if reexport_path:
            vault_hint = (dataset.get("sync") or {}).get("vaultRootHint")
            with io.StringIO() as _out, redirect_stdout(_out):
                build_export(cfg, output_path=reexport_path, vault_root_hint=vault_hint)
            dry_report["reexportedTo"] = str(Path(reexport_path))
        return dry_report

    change_lookup = {}
    for cr in change_rows:
        cid = cr.get("changeId")
        if cid:
            change_lookup[cid] = cr

    approved_ids: list[str] = []

    for idx, result in enumerate(results, 1):
        change_id = result.get("changeId") or f"(unknown-{idx})"
        item_id = result.get("itemId") or "(unknown)"
        action = result.get("action") or "update"
        status = result.get("status") or "unknown"
        message = result.get("message") or ""
        item = result.get("item") or {}

        print(f"[{idx}/{len(results)}] {change_id} ({action})", file=stream)
        print(f"  itemId: {item_id}", file=stream)
        print(f"  status: {status}", file=stream)
        if message:
            print(f"  message: {message}", file=stream)

        change_row = change_lookup.get(change_id, {})
        proposed = change_row.get("proposed") or {}
        if proposed:
            for field, value in proposed.items():
                current = item.get(field)
                cur_s = _fmt_value(current)
                new_s = _fmt_value(value)
                marker = " *" if str(current) != str(value) else ""
                print(f"    {field}: {cur_s} \u2192 {new_s}{marker}", file=stream)

        if status != "dry-run":
            print(f"  \u2192 auto-skipped ({status})", file=stream)
            print("", file=stream)
            continue

        stream.flush()
        answer = input_fn(f"Approve {change_id}? [y/N]: ").strip().lower()
        if answer in {"y", "yes"}:
            approved_ids.append(change_id)
            print("  \u2192 approved", file=stream)
        else:
            print("  \u2192 skipped", file=stream)
        print("", file=stream)

    print(f"Approved {len(approved_ids)} of {len(results)} change(s).", file=stream)

    if not approved_ids:
        print("No changes approved \u2014 vault unchanged.", file=stream)
        if reexport_path:
            vault_hint = (dataset.get("sync") or {}).get("vaultRootHint")
            with io.StringIO() as _out, redirect_stdout(_out):
                build_export(cfg, output_path=reexport_path, vault_root_hint=vault_hint)
            dry_report["reexportedTo"] = str(Path(reexport_path))
        return dry_report

    approved_set = set(approved_ids)
    filtered_rows = [r for r in change_rows if r.get("changeId") in approved_set]
    dataset.setdefault("changes", {})["rows"] = filtered_rows

    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as f:
        temp_path = Path(f.name)
        json.dump(dataset, f, indent=2, ensure_ascii=True)
        f.write("\n")

    try:
        report = import_grid_changes(
            str(temp_path),
            cfg,
            dry_run=False,
            reexport_path=reexport_path,
        )
    finally:
        temp_path.unlink(missing_ok=True)

    applied = (report.get("counts") or {}).get("applied", 0)
    print(f"\nApplied {applied} change(s) to Obsidian vault.", file=stream)
    if reexport_path:
        print(f"Re-exported dataset to: {reexport_path}", file=stream)

    return report


def _fmt_value(value) -> str:
    if value is None or value == "":
        return "<empty>"
    if isinstance(value, str):
        return value
    return json.dumps(value, ensure_ascii=True)
