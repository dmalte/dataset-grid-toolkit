# Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

"""YAML-driven tool loader, template resolver, and subprocess args builder.

Reads ``tools.yaml`` from the project root, resolves ``${...}`` template
expressions against control values / settings / datetime, and builds the
final subprocess argument list for execution by the desktop bridge.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

_TEMPLATE_RE = re.compile(r"\$\{([^}]+)\}")


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def load_tools_yaml(yaml_path: Path) -> dict[str, Any]:
    """Load and return the full ``tools.yaml`` configuration."""
    with open(yaml_path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def get_defaults(cfg: dict[str, Any]) -> dict[str, Any]:
    return dict(cfg.get("defaults") or {})


def get_settings_schema(cfg: dict[str, Any]) -> dict[str, Any]:
    return dict(cfg.get("settings") or {})


def get_tools(cfg: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return dict(cfg.get("tools") or {})


def get_allowed_executables(cfg: dict[str, Any]) -> frozenset[str]:
    """Return the set of executable names declared across all tools."""
    executables: set[str] = set()
    for tdef in (cfg.get("tools") or {}).values():
        cmd = tdef.get("command") or {}
        exe = cmd.get("executable")
        if exe:
            executables.add(str(exe))
    return frozenset(executables)


# ---------------------------------------------------------------------------
# Frontend-serialisable summary
# ---------------------------------------------------------------------------

def tools_summary(cfg: dict[str, Any]) -> dict[str, Any]:
    """Build a JSON-serialisable summary sent to the frontend.

    Returns ``{settings: {...}, tools: [...]}``.
    """
    settings_schema = get_settings_schema(cfg)
    defaults = get_defaults(cfg)
    tools = get_tools(cfg)

    tool_list = []
    for tid, tdef in tools.items():
        ui = tdef.get("ui") or {}
        controls = tdef.get("controls") or []
        result = tdef.get("result") or {}

        # Pre-resolve default values that reference ${defaults.*}
        resolved_controls = []
        for ctrl in controls:
            c = dict(ctrl)
            if "default" in c and isinstance(c["default"], str):
                c["default"] = _resolve_defaults_only(c["default"], defaults)
            resolved_controls.append(c)

        tool_list.append({
            "id": tid,
            "label": tdef.get("label", tid),
            "description": tdef.get("description", ""),
            "type": tdef.get("type", "action"),
            "ui": {
                "order": int(ui.get("order", 0)),
                "tab": str(ui.get("tab", "General")),
                "group": str(ui.get("group", "Actions")),
            },
            "controls": resolved_controls,
            "requires": tdef.get("requires") or [],
            "result": {
                "auto_load": bool(result.get("auto_load")),
                "save_before_run": bool(result.get("save_before_run")),
                "promote_after": bool(result.get("promote_after")),
                "success_message": str(result.get("success_message") or ""),
            },
        })

    return {
        "settings": {
            key: {
                "label": spec.get("label", key),
                "type": spec.get("type", "text"),
                "placeholder": spec.get("placeholder", ""),
                "tab": spec.get("tab", ""),
            }
            for key, spec in settings_schema.items()
        },
        "tools": sorted(tool_list, key=lambda t: t["ui"]["order"]),
    }


def _resolve_defaults_only(template: str, defaults: dict[str, Any]) -> str:
    """Resolve only ``${defaults.*}`` expressions (for control defaults)."""
    def replace(match: re.Match[str]) -> str:
        expr = match.group(1)
        if expr.startswith("defaults."):
            return str(defaults.get(expr.split(".", 1)[1], ""))
        return match.group(0)  # leave unresolved
    return _TEMPLATE_RE.sub(replace, template)


# ---------------------------------------------------------------------------
# Template resolution
# ---------------------------------------------------------------------------

def render_template(
    template: str,
    *,
    defaults: dict[str, Any],
    controls: dict[str, str],
    settings: dict[str, str],
    paths: dict[str, str],
    outputs: dict[str, str],
    now: datetime | None = None,
) -> str:
    """Resolve ``${...}`` template expressions in a string.

    Supported expressions:
        ${now:<strftime>}         — formatted current datetime
        ${defaults.<key>}         — from defaults dict
        ${settings.<key>}         — from settings dict
        ${paths.<key>}            — from paths dict
        ${outputs.<key>}          — from outputs dict
        ${<control_id>}           — control value (as-is)
        ${<control_id>.stem}      — Path().stem of control value
        ${<control_id>.name}      — Path().name of control value
    """
    if now is None:
        now = datetime.now().astimezone()

    def replace(match: re.Match[str]) -> str:
        expr = match.group(1)

        # ${now:<format>}
        if expr.startswith("now:"):
            return now.strftime(expr.split(":", 1)[1])

        # ${defaults.<key>}
        if expr.startswith("defaults."):
            key = expr.split(".", 1)[1]
            if key not in defaults:
                raise ValueError(f"Unknown default: {key}")
            return str(defaults[key])

        # ${settings.<key>}
        if expr.startswith("settings."):
            key = expr.split(".", 1)[1]
            if key not in settings:
                raise ValueError(f"Setting not configured: {key}")
            return str(settings[key])

        # ${paths.<key>}
        if expr.startswith("paths."):
            key = expr.split(".", 1)[1]
            if key not in paths:
                raise ValueError(f"Unknown path: {key}")
            return str(paths[key])

        # ${outputs.<key>}
        if expr.startswith("outputs."):
            key = expr.split(".", 1)[1]
            if key not in outputs:
                raise ValueError(f"Unknown output: {key}")
            return str(outputs[key])

        # ${control_id.stem} or ${control_id.name}
        if "." in expr:
            ctrl_id, attr = expr.rsplit(".", 1)
            if ctrl_id in controls:
                p = Path(controls[ctrl_id])
                if attr == "stem":
                    return p.stem
                if attr == "name":
                    return p.name
                raise ValueError(f"Unknown path attribute: {attr}")

        # ${control_id} — plain control value
        if expr in controls:
            return str(controls[expr])

        raise ValueError(f"Unresolved template expression: ${{{expr}}}")

    return _TEMPLATE_RE.sub(replace, template)


# ---------------------------------------------------------------------------
# Tool resolution and execution
# ---------------------------------------------------------------------------

def resolve_tool(
    cfg: dict[str, Any],
    tool_id: str,
    *,
    control_values: dict[str, str],
    settings_values: dict[str, str],
    project_root: Path,
    dataset_path: str | None = None,
    source_info: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Resolve a tool's command args and output paths.

    Returns ``{args: [...], outputs: {...}, result: {...}}``.
    """
    tools = get_tools(cfg)
    if tool_id not in tools:
        raise ValueError(f"Unknown tool: {tool_id}")

    tdef = tools[tool_id]
    defaults = get_defaults(cfg)
    source_info = source_info or {}

    # Resolve hidden controls from their sources
    controls: dict[str, str] = {}
    for ctrl in tdef.get("controls") or []:
        cid = ctrl.get("id", "")
        ctype = ctrl.get("type", "")

        if ctype == "action_button":
            continue

        if ctype == "hidden":
            source = ctrl.get("source", "")
            if source == "current-dataset":
                controls[cid] = dataset_path or ""
            elif source == "tracked-source":
                source_key = ctrl.get("source_key", "source_path")
                controls[cid] = str(source_info.get(source_key, ""))
            elif source == "source-info":
                source_key = ctrl.get("source_key", cid)
                value = source_info.get(source_key, "")
                if not value:
                    fallback = ctrl.get("fallback", "")
                    if fallback:
                        value = render_template(
                            fallback,
                            defaults=defaults,
                            controls={},
                            settings=settings_values,
                            paths={},
                            outputs={},
                        )
                controls[cid] = str(value)
            elif source == "setting":
                setting_key = ctrl.get("setting_key", cid)
                controls[cid] = settings_values.get(setting_key, "")
            elif source == "default":
                controls[cid] = str(ctrl.get("default", ""))
            else:
                controls[cid] = ""
        else:
            # User-provided value from the frontend
            controls[cid] = control_values.get(cid, ctrl.get("default") or "")

    # Resolve paths relative to project_root
    raw_paths = tdef.get("paths") or {}
    resolved_paths: dict[str, str] = {}
    for key, value in raw_paths.items():
        p = project_root / value
        p.mkdir(parents=True, exist_ok=True)
        resolved_paths[key] = str(p)

    # Resolve outputs (need paths + controls + now)
    now = datetime.now().astimezone()
    raw_outputs = (tdef.get("result") or {}).get("outputs") or {}
    resolved_outputs: dict[str, str] = {}
    for oname, otemplate in raw_outputs.items():
        resolved_outputs[oname] = render_template(
            otemplate,
            defaults=defaults,
            controls=controls,
            settings=settings_values,
            paths=resolved_paths,
            outputs={},
            now=now,
        )

    # Resolve command args
    command = tdef.get("command") or {}
    executable = command.get("executable", "")
    raw_args = command.get("args") or []
    resolved_args = [executable]
    for arg_template in raw_args:
        resolved_args.append(
            render_template(
                arg_template,
                defaults=defaults,
                controls=controls,
                settings=settings_values,
                paths=resolved_paths,
                outputs=resolved_outputs,
                now=now,
            )
        )

    result_spec = tdef.get("result") or {}
    track_source = result_spec.get("track_source")

    # Build source tracking info if configured
    source_tracking = None
    if track_source:
        from_ctrl = track_source.get("from", "")
        source_tracking = {
            "source_path": controls.get(from_ctrl, ""),
            "kind": track_source.get("kind", ""),
            "converted_at": now.strftime("%Y-%m-%dT%H:%M:%S"),
        }
        sheet_key = track_source.get("sheet_key")
        if sheet_key and sheet_key in controls:
            source_tracking["sheet"] = controls[sheet_key]

    return {
        "args": resolved_args,
        "outputs": resolved_outputs,
        "track_source": source_tracking,
        "result": {
            "auto_load": bool(result_spec.get("auto_load")),
            "save_before_run": bool(result_spec.get("save_before_run")),
            "promote_after": bool(result_spec.get("promote_after")),
            "success_message": str(result_spec.get("success_message") or ""),
        },
    }
