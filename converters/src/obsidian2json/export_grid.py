import hashlib
import json
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

try:
    from item_manager.core import atomic_write_text, extract_tags, ItemRepository
except ImportError:
    from ._vendor import atomic_write_text, extract_tags, ItemRepository


ITEM_TYPES = ("action", "buy", "consume", "idea")
EXPORT_VERSION = "1"

TYPE_GROUPS = [
    {"id": "grp-action", "name": "Action Items", "color": "#3498db", "itemType": "action"},
    {"id": "grp-buy", "name": "Buy Items", "color": "#2ecc71", "itemType": "buy"},
    {"id": "grp-consume", "name": "Consume Items", "color": "#e67e22", "itemType": "consume"},
    {"id": "grp-idea", "name": "Idea Items", "color": "#9b59b6", "itemType": "idea"},
]

RELATION_TYPES = [
    {
        "name": "parent",
        "inwardName": "is parent of",
        "outwardName": "is parent of",
        "sourceType": "obsidian",
        "sourceLinkType": "Parent",
    },
    {
        "name": "child",
        "inwardName": "is child of",
        "outwardName": "is child of",
        "sourceType": "obsidian",
        "sourceLinkType": "Parent",
    },
]

FIELD_SCHEMA = {
    "text": {"type": "string", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "child_md": {"type": "string", "kind": "data", "editable": False, "selectable": True, "visible": True},
    "status": {"type": "enum", "kind": "data", "editable": True, "selectable": True, "visible": True, "validValues": ["new", "done", "ongoing", "open", "waiting"]},
    "created": {"type": "date", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "completed": {"type": "date", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "due": {"type": "date", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "area": {"type": "string", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "priority": {"type": "enum", "kind": "data", "editable": True, "selectable": True, "visible": True, "validValues": ["low", "medium", "high", "highest"]},
    "urgency": {"type": "enum", "kind": "data", "editable": True, "selectable": True, "visible": True, "validValues": ["low", "medium", "high", "highest"]},
    "shop": {"type": "string", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "tags": {"type": "multi-value", "kind": "data", "editable": True, "selectable": True, "visible": True},
    "id": {"type": "string", "kind": "data", "editable": False, "selectable": True, "visible": True},
    "itemType": {"type": "enum", "kind": "data", "editable": False, "selectable": True, "visible": True},
    "file": {"type": "string", "kind": "data", "editable": False, "selectable": True, "visible": False},
    "line_no": {"type": "number", "kind": "internal", "editable": False, "selectable": False, "visible": False},
    "indent": {"type": "number", "kind": "internal", "editable": False, "selectable": False, "visible": False},
    "raw": {"type": "string", "kind": "internal", "editable": False, "selectable": False, "visible": False},
    "managed": {"type": "boolean", "kind": "internal", "editable": False, "selectable": False, "visible": False},
    "url": {"type": "string", "kind": "internal", "editable": False, "selectable": False, "visible": False},
    "sourceRef": {"type": "structured", "kind": "system", "editable": False, "selectable": False, "visible": False},
    "sync": {"type": "structured", "kind": "system", "editable": False, "selectable": False, "visible": False},
    "relations": {"type": "structured", "kind": "system", "editable": False, "selectable": False, "visible": False},
    "created_week": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "created_month": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "completed_week": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "completed_month": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "due_week": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "due_month": {"type": "string", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "is_overdue": {"type": "boolean", "kind": "derived", "editable": False, "selectable": True, "visible": True},
    "_hasParent": {"type": "scalar", "subtype": "boolean", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
    "_hasChildren": {"type": "scalar", "subtype": "boolean", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
    "_childrenCount": {"type": "scalar", "subtype": "number", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
    "_blocksCount": {"type": "scalar", "subtype": "number", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
    "_isBlocked": {"type": "scalar", "subtype": "boolean", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
    "_relationsCount": {"type": "scalar", "subtype": "number", "kind": "relationship", "editable": False, "selectable": True, "visible": False},
}

SCHEMA_LISTS = {
    "itemTypes": ["action", "buy", "consume", "idea"],
    "writableFields": ["text", "status", "created", "completed", "due", "area", "priority", "urgency", "shop", "tags"],
    "readOnlyFields": ["id", "itemType", "file", "child_md", "line_no", "indent", "raw", "managed", "url", "sourceRef", "sync"],
}

HASH_EXCLUDE_KEYS = frozenset(["child_md", "relations", "sourceRef", "sync", "url"] + [key for key, value in FIELD_SCHEMA.items() if value["kind"] == "derived"])


def _iso_week(date_str: str | None) -> str | None:
    if not date_str:
        return None
    try:
        parsed = datetime.strptime(date_str, "%Y-%m-%d")
        iso_year, iso_week, _ = parsed.isocalendar()
        return f"{iso_year}-W{iso_week:02d}"
    except (TypeError, ValueError):
        return None


def _iso_month(date_str: str | None) -> str | None:
    if not date_str:
        return None
    try:
        return date_str[:7] if len(date_str) >= 7 else None
    except (TypeError, IndexError):
        return None


def _is_overdue(item_dict: dict, today_str: str) -> bool | None:
    due = item_dict.get("due")
    if not due:
        return None
    if item_dict.get("status") == "done":
        return False
    return due < today_str


def _encode_obsidian_vault_name(vault_path: str) -> str:
    vault_name = Path(vault_path).name if vault_path else ""
    return quote(vault_name, safe="")


def _build_obsidian_file_url(vault_name: str, vault_relative_path: str) -> str | None:
    if not vault_name or not vault_relative_path:
        return None
    normalized_path = str(vault_relative_path).replace("\\", "/")
    return f"obsidian://open?vault={vault_name}&file={quote(normalized_path, safe='/')}"


def _build_card_click_config(vault_name: str) -> dict:
    return {
        "mode": "open-url",
        "openInNewTab": False,
        "openUrl": {
            "baseUrl": f"obsidian://search?vault={vault_name}&query=",
            "suffix": "",
            "useUrlField": True,
            "argumentGenerator": {"type": "template", "template": "${id}", "field": "", "source": ""},
        },
    }


def _compute_baseline_hash(item_dict: dict) -> str:
    filtered = {key: value for key, value in item_dict.items() if key not in HASH_EXCLUDE_KEYS}
    canonical = json.dumps(filtered, sort_keys=True, ensure_ascii=True, separators=(",", ":"))
    digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    return f"sha256:{digest}"


def _build_field_types() -> dict:
    return _build_field_types_from(FIELD_SCHEMA)


def _build_field_types_from(schema: dict) -> dict:
    field_types = {}
    for field_name, meta in schema.items():
        field_type = meta["type"]
        if field_type == "multi-value":
            field_types[field_name] = "multi-value"
        elif field_type == "structured":
            field_types[field_name] = "structured"
        else:
            field_types[field_name] = "scalar"
    return field_types


def _leading_indent(line: str) -> int:
    return len(line) - len(line.lstrip())


def _is_blank(line: str) -> bool:
    return not line.strip()


def _find_block_end(line_no: int, indent: int, lines: list[str]) -> int:
    for index in range(line_no, len(lines)):
        line = lines[index]
        if _is_blank(line):
            continue
        if _leading_indent(line) <= indent:
            return index
    return len(lines)


def _build_child_md(line_no: int, indent: int, block_end: int, lines: list[str]) -> str | None:
    nested_lines = lines[line_no:block_end]
    child_indent = None
    for line in nested_lines:
        if _is_blank(line):
            continue
        line_indent = _leading_indent(line)
        if line_indent > indent and (child_indent is None or line_indent < child_indent):
            child_indent = line_indent
    if child_indent is None:
        return None
    child_lines = [line.rstrip("\n") for line in nested_lines if not _is_blank(line) and _leading_indent(line) == child_indent]
    if not child_lines:
        return None
    return "\n".join(child_lines)


def _build_parent_relation(child_record: dict, parent_record: dict) -> dict:
    child_id = child_record["item"].id
    parent_item = parent_record["item"]
    return {
        "relationId": f"rel-parent-{child_id}-{parent_item.id}",
        "type": "parent",
        "direction": "outward",
        "target": {
            "itemId": parent_item.id,
            "sourceRef": {
                "sourceType": "obsidian",
                "vaultRelativePath": parent_item.file,
                "managedId": parent_item.id,
            },
        },
        "meta": {"sourceType": "obsidian", "sourceLinkType": "Parent"},
    }


def _annotate_hierarchy(scan_records: list[dict]) -> None:
    records_by_file = defaultdict(list)
    for record in scan_records:
        records_by_file[record["item"].file].append(record)

    for file_records in records_by_file.values():
        file_records.sort(key=lambda record: record["item"].line_no)
        lines = file_records[0]["lines"]

        for record in file_records:
            item = record["item"]
            block_end = _find_block_end(item.line_no, item.indent, lines)
            record["block_end"] = block_end
            record["child_md"] = _build_child_md(item.line_no, item.indent, block_end, lines)
            record["relations"] = []

        stack = []
        for record in file_records:
            item = record["item"]
            while stack and item.line_no > stack[-1]["block_end"]:
                stack.pop()
            while stack and item.indent <= stack[-1]["item"].indent:
                stack.pop()
            if stack:
                parent_record = stack[-1]
                record["relations"].append(_build_parent_relation(record, parent_record))
            stack.append(record)


def serialize_item(item, item_type: str, export_timestamp: str, today_str: str, vault_name: str, child_md: str | None = None, relations: list[dict] | None = None) -> dict:
    data = item.to_dict()
    data["itemType"] = item_type
    data["child_md"] = child_md
    data["tags"] = extract_tags(item.raw, getattr(item, "shop", None))
    field_tags = data.pop("field_tags", {}) or {}
    for prefix, value in field_tags.items():
        data[prefix] = value
    data["url"] = _build_obsidian_file_url(vault_name, item.file)
    data["sourceRef"] = {
        "sourceType": "obsidian",
        "vaultRelativePath": item.file,
        "managedId": item.id,
    }
    data["sync"] = {
        "baselineHash": _compute_baseline_hash(data),
        "exportedAt": export_timestamp,
    }
    if relations:
        data["relations"] = relations
    data["created_week"] = _iso_week(data.get("created"))
    data["created_month"] = _iso_month(data.get("created"))
    data["completed_week"] = _iso_week(data.get("completed"))
    data["completed_month"] = _iso_month(data.get("completed"))
    data["due_week"] = _iso_week(data.get("due"))
    data["due_month"] = _iso_month(data.get("due"))
    data["is_overdue"] = _is_overdue(data, today_str)
    return data


def _build_type_groups() -> list[dict]:
    return [
        {
            "id": group["id"],
            "name": group["name"],
            "color": group["color"],
            "enabled": True,
            "rule": {"field": "itemType", "operator": "equals", "values": [group["itemType"]]},
            "manualMembers": [],
        }
        for group in TYPE_GROUPS
    ]


def build_export(cfg: dict, output_path: str, vault_root_hint: str | None = None) -> dict:
    vault = cfg["vault"]
    vault_name = _encode_obsidian_vault_name(vault)
    search_paths = cfg.get("search_paths", [])
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    card_click_config = _build_card_click_config(vault_name)

    seen_ids: set[str] = set()
    scan_records: list[dict] = []

    for item_type in ITEM_TYPES:
        repo = ItemRepository(vault, search_paths, item_type)
        for item, path, lines in repo.scan_items(include_unmanaged=False):
            if item.id in seen_ids:
                print(f"Warning: duplicate id {item.id} skipped (type={item_type})", file=sys.stderr)
                continue
            seen_ids.add(item.id)
            scan_records.append({"item": item, "item_type": item_type, "path": path, "lines": lines})

    _annotate_hierarchy(scan_records)
    data = [
        serialize_item(
            record["item"],
            record["item_type"],
            now,
            today_str,
            vault_name,
            child_md=record.get("child_md"),
            relations=record.get("relations"),
        )
        for record in scan_records
    ]

    # Discover dynamic field-tag keys and register them in the schema
    dynamic_keys: set[str] = set()
    for row in data:
        for key in row:
            if key not in FIELD_SCHEMA:
                dynamic_keys.add(key)
    dynamic_field_schema = dict(FIELD_SCHEMA)
    dynamic_writable = list(SCHEMA_LISTS["writableFields"])
    for key in sorted(dynamic_keys):
        dynamic_field_schema[key] = {"type": "string", "kind": "data", "editable": True, "selectable": True, "visible": True}
        dynamic_writable.append(key)

    export = {
        "meta": {
            "datasetType": "obsidian-item-grid",
            "sourceSystem": "obsidian2json",
            "version": EXPORT_VERSION,
            "exportedAt": now,
            "cardClick": card_click_config,
            "supportsRelations": True,
        },
        "schema": {
            "fields": dynamic_field_schema,
            "relationTypes": RELATION_TYPES,
            "fieldTypes": _build_field_types_from(dynamic_field_schema),
            **{k: v for k, v in SCHEMA_LISTS.items() if k != "writableFields"},
            "writableFields": dynamic_writable,
        },
        "sync": {
            "exportedAt": now,
            "vaultRootHint": vault_root_hint or vault,
            "exportMode": "full",
        },
        "data": data,
        "changes": {"version": EXPORT_VERSION, "rows": []},
        "view": {
            "groups": _build_type_groups(),
            "relations": {
                "focusMode": False,
                "focusRootId": None,
                "focusDepth": 1,
                "focusTypes": None,
            },
            "showDataFields": True,
            "showRelationshipFields": False,
            "showDerivedFields": False,
            "cardClick": card_click_config,
            "urlConfig": {
                "prefix": f"obsidian://search?vault={vault_name}&query=",
                "suffix": "",
                "useUrlField": True,
            },
        },
    }

    out = Path(output_path)
    atomic_write_text(out, json.dumps(export, ensure_ascii=False, indent=2) + "\n")
    return export
