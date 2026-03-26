# Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .converter import (
    build_default_excel_output_path,
    build_default_json_output_path,
    build_default_review_html_output_path,
    build_default_review_json_output_path,
    apply_approved_review,
    excel_to_json,
    json_to_excel,
    json_to_excel_changes,
    json_to_excel_confirm_cli,
    json_to_excel_confirm_html,
)
from .confirm_ui import json_to_excel_confirm_ui


def _write_warnings(warnings: list[str]) -> None:
    for warning in warnings:
        print(f"warning: {warning}", file=sys.stderr)


def _write_debug(enabled: bool, label: str, value: Path | str) -> None:
    if enabled:
        print(f"{label}: {value}", file=sys.stderr)


def _normalize_mode(mode: str) -> str:
    aliases = {
        "full": "restore",
        "confirm": "confirm-cli",
    }
    return aliases.get(mode, mode)


def excel_to_json_main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="excel-to-json")
    parser.add_argument("--excel", required=True, help="Path to the input .xlsx file")
    parser.add_argument("--json", dest="json_path", help="Path to the output .json file")
    parser.add_argument("--sheet", default="data", help="Worksheet name to read")
    parser.add_argument("--debug", action="store_true", help="Print resolved input/output paths to stderr")
    args = parser.parse_args(argv)

    excel_path = Path(args.excel)
    json_path = Path(args.json_path) if args.json_path else build_default_json_output_path(excel_path)
    _, warnings = excel_to_json(excel_path, json_path=json_path, sheet_name=args.sheet)
    _write_debug(args.debug, "input_excel", excel_path.resolve())
    _write_debug(args.debug, "output_json", json_path.resolve())
    _write_warnings(warnings)
    print(str(json_path))
    return 0


def json_to_excel_main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="json-to-excel")
    parser.add_argument("--json", required=True, help="Path to the input .json file")
    parser.add_argument("--excel", dest="excel_path", help="Path to the output .xlsx file")
    parser.add_argument("--sheet", default="data", help="Worksheet name to write")
    parser.add_argument("--debug", action="store_true", help="Print resolved input/output paths to stderr")
    parser.add_argument(
        "--mode",
        choices=("restore", "changes", "confirm", "confirm-cli", "confirm-html", "confirm-ui", "full"),
        default="confirm-cli",
        help=(
            "Conversion mode. 'restore' writes baseline data, 'changes' overlays pending changes directly, "
            "'confirm-cli' prompts in the console, 'confirm-html' generates or applies HTML review artifacts, "
            "and 'confirm-ui' opens a native desktop review window."
        ),
    )
    parser.add_argument(
        "--include-source-ref",
        action="store_true",
        help="Write sourceRef fields to Excel columns when present",
    )
    parser.add_argument("--review-json", dest="review_json_path", help="Path to the changes review .json file")
    parser.add_argument("--review-html", dest="review_html_path", help="Path to the changes review .html file")
    parser.add_argument(
        "--apply-approved",
        dest="apply_approved_path",
        help="Apply approved rows from an existing changes review .json file",
    )
    args = parser.parse_args(argv)
    mode = _normalize_mode(args.mode)

    json_path = Path(args.json)
    requested_excel_path = Path(args.excel_path) if args.excel_path else None
    excel_path = requested_excel_path or build_default_excel_output_path(json_path)

    if mode not in {"confirm-html", "confirm-ui"} and (
        args.review_json_path is not None or args.review_html_path is not None or args.apply_approved_path is not None
    ):
        parser.error("--review-json, --review-html, and --apply-approved require --mode confirm-html or --mode confirm-ui")

    if mode == "confirm-ui" and (args.review_html_path is not None or args.apply_approved_path is not None):
        parser.error("--review-html and --apply-approved are only supported with --mode confirm-html")

    if mode == "confirm-html":
        if args.apply_approved_path:
            review_artifact = json.loads(Path(args.apply_approved_path).read_text(encoding="utf-8"))
            artifact_excel_path = review_artifact.get("excelPath")
            resolved_excel_path = requested_excel_path or (
                Path(artifact_excel_path) if isinstance(artifact_excel_path, str) else None
            )
            warnings = apply_approved_review(
                Path(args.apply_approved_path),
                excel_path=resolved_excel_path,
                sheet_name=args.sheet,
            )
            _write_debug(args.debug, "review_json", Path(args.apply_approved_path).resolve())
            _write_debug(args.debug, "output_excel", (resolved_excel_path or excel_path).resolve())
            _write_warnings(warnings)
            print(str(resolved_excel_path or excel_path))
            return 0

        review_json_path = Path(args.review_json_path) if args.review_json_path else None
        review_html_path = Path(args.review_html_path) if args.review_html_path else build_default_review_html_output_path(json_path)
        result, warnings = json_to_excel_confirm_html(
            json_path,
            excel_path=excel_path,
            sheet_name=args.sheet,
            include_source_ref=args.include_source_ref,
            review_json_path=review_json_path,
            review_html_path=review_html_path,
        )
        _write_debug(args.debug, "input_json", json_path.resolve())
        _write_debug(args.debug, "review_json", Path(result["review_json_path"]).resolve())
        _write_debug(args.debug, "review_html", Path(result["review_html_path"]).resolve())
        _write_warnings(warnings)
        print(str(result["review_json_path"]))
        print(str(result["review_html_path"]))
        return 0

    if mode == "confirm-ui":
        review_json_path = Path(args.review_json_path) if args.review_json_path else build_default_review_json_output_path(json_path)
        result, warnings = json_to_excel_confirm_ui(
            json_path,
            excel_path=excel_path,
            sheet_name=args.sheet,
            include_source_ref=args.include_source_ref,
            review_json_path=review_json_path,
        )
        _write_debug(args.debug, "input_json", json_path.resolve())
        _write_debug(args.debug, "review_json", Path(result["review_json_path"]).resolve())
        _write_debug(args.debug, "output_excel", excel_path.resolve())
        _write_warnings(warnings)
        print(str(result["review_json_path"]))
        if result.get("applied"):
            print(str(excel_path))
        return 0

    if mode == "restore":
        warnings = json_to_excel(
            json_path,
            excel_path=excel_path,
            sheet_name=args.sheet,
            include_source_ref=args.include_source_ref,
        )
    elif mode == "changes":
        warnings = json_to_excel_changes(
            json_path,
            excel_path=excel_path,
            sheet_name=args.sheet,
            include_source_ref=args.include_source_ref,
        )
    else:
        warnings = json_to_excel_confirm_cli(
            json_path,
            excel_path=excel_path,
            sheet_name=args.sheet,
            include_source_ref=args.include_source_ref,
        )

    _write_debug(args.debug, "input_json", json_path.resolve())
    _write_debug(args.debug, "output_excel", excel_path.resolve())
    _write_warnings(warnings)
    print(str(excel_path))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="data2json-converters")
    subparsers = parser.add_subparsers(dest="command", required=True)

    excel_parser = subparsers.add_parser("excel-to-json")
    excel_parser.add_argument("--excel", required=True)
    excel_parser.add_argument("--json", dest="json_path")
    excel_parser.add_argument("--sheet", default="data")
    excel_parser.add_argument("--debug", action="store_true")

    json_parser = subparsers.add_parser("json-to-excel")
    json_parser.add_argument("--json", required=True)
    json_parser.add_argument("--excel", dest="excel_path")
    json_parser.add_argument("--sheet", default="data")
    json_parser.add_argument("--debug", action="store_true")
    json_parser.add_argument(
        "--mode",
        choices=("restore", "changes", "confirm", "confirm-cli", "confirm-html", "confirm-ui", "full"),
        default="confirm-cli",
    )
    json_parser.add_argument("--include-source-ref", action="store_true")
    json_parser.add_argument("--review-json", dest="review_json_path")
    json_parser.add_argument("--review-html", dest="review_html_path")
    json_parser.add_argument("--apply-approved", dest="apply_approved_path")

    args = parser.parse_args(argv)
    if args.command == "excel-to-json":
        forward_args = ["--excel", args.excel, "--sheet", args.sheet]
        if args.json_path:
            forward_args.extend(["--json", args.json_path])
        if args.debug:
            forward_args.append("--debug")
        return excel_to_json_main(forward_args)

    forward_args = ["--json", args.json, "--sheet", args.sheet, "--mode", args.mode]
    if args.excel_path:
        forward_args.extend(["--excel", args.excel_path])
    if args.debug:
        forward_args.append("--debug")
    if args.include_source_ref:
        forward_args.append("--include-source-ref")
    if args.review_json_path:
        forward_args.extend(["--review-json", args.review_json_path])
    if args.review_html_path:
        forward_args.extend(["--review-html", args.review_html_path])
    if args.apply_approved_path:
        forward_args.extend(["--apply-approved", args.apply_approved_path])
    return json_to_excel_main(forward_args)


if __name__ == "__main__":
    raise SystemExit(main())