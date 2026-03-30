from __future__ import annotations

import argparse
import io
import json
import sys
from pathlib import Path

try:
    from item_manager.core import load_config
except ImportError:
    from ._vendor import load_config
from .export_grid import build_export
from .import_grid import import_grid_changes, import_grid_changes_confirm_cli


def _ensure_utf8_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    else:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


def _warn(message: str) -> None:
    print(f"Warning: {message}", file=sys.stderr)


def _validate_data_source(cfg: dict) -> bool:
    meta = cfg.get("_meta", {})
    for warning in meta.get("warnings", []):
        _warn(warning)

    config_path = meta.get("config_path", "the config file")
    vault = (cfg.get("vault") or "").strip()
    if not vault:
        _warn(f"Obsidian vault path is not configured. Update 'vault' in {config_path}.")
        return False

    vault_path = Path(vault).expanduser()
    if not vault_path.exists():
        _warn(f"Configured Obsidian vault was not found: {vault_path}. Update 'vault' in {config_path}.")
        return False
    if not vault_path.is_dir():
        _warn(f"Configured Obsidian vault is not a directory: {vault_path}. Update 'vault' in {config_path}.")
        return False

    search_paths = cfg.get("search_paths") or []
    if search_paths and not any((vault_path / search_path).exists() for search_path in search_paths):
        _warn(
            f"None of the configured search_paths were found under {vault_path}. "
            "Commands may return no data until those folders exist."
        )
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="obsidian2json")
    parser.add_argument("--config", help="Optional path to the Obsidian sync config YAML")
    subparsers = parser.add_subparsers(dest="command", required=True)

    export_parser = subparsers.add_parser("export-grid", help="Export all managed Obsidian items to one grid JSON dataset")
    export_parser.add_argument("--output", required=True, help="Output JSON file path")
    export_parser.add_argument("--vault-root-hint", default=None, help="Override vault path hint in sync metadata")

    import_parser = subparsers.add_parser("import-grid", help="Apply grid JSON changes back to the Obsidian vault")
    import_parser.add_argument("--input", required=True, help="Input dataset JSON path")
    import_parser.add_argument("--dry-run", action="store_true", help="Validate changes without writing vault files")
    import_parser.add_argument("--confirm", action="store_true", help="Interactive review: dry-run, prompt per change, apply approved")
    import_parser.add_argument("--report", default=None, help="Optional report JSON output path")
    import_parser.add_argument("--reexport", default=None, help="Optional fresh export path after successful apply")

    args = parser.parse_args(argv)

    cfg = load_config(Path(args.config) if args.config else None)
    if not _validate_data_source(cfg):
        print(json.dumps({"ok": False, "error": "obsidian2json data source not found"}))
        return 1

    _ensure_utf8_stdout()

    if args.command == "export-grid":
        export = build_export(cfg, output_path=args.output, vault_root_hint=getattr(args, "vault_root_hint", None))
        print(json.dumps({"ok": True, "command": "export-grid", "output": str(Path(args.output)), "itemCount": len(export.get("data", []))}, ensure_ascii=False))
        return 0

    if getattr(args, "confirm", False):
        report = import_grid_changes_confirm_cli(
            args.input,
            cfg,
            reexport_path=args.reexport,
        )
        return 0 if report.get("ok", True) else 1

    report = import_grid_changes(
        args.input,
        cfg,
        dry_run=bool(args.dry_run),
        report_path=args.report,
        reexport_path=args.reexport,
    )
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report.get("ok") else 1
