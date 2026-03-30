# Structured Data Grid Toolkit

## Overview

This public bundle contains a PySide6 desktop shell and Python CLI converters that share the same JSON dataset format. It is intended for local exploration, lightweight editing, Excel round-tripping, and CLI-driven import/export workflows.

## Features

- Native desktop shell with card and table views, filters, sticky headers, and pending-change overlays
- YAML-driven tool execution through local CLI entry points
- Excel import and export via `excel-to-json` and `json-to-excel`
- Obsidian import and export via `obsidian2json`
- Sample JSON and Excel files for local testing
- Exported bundle includes a generated `.gitignore`, bundle reports, and minimal package metadata

## Architecture

- `grid/src/`: vanilla JS/HTML/CSS frontend
- `app/`: PySide6 desktop shell and QWebChannel bridge
- `converters/src/data2json_converters/`: Excel/JSON conversion CLI
- `converters/src/obsidian2json/`: Obsidian/grid JSON conversion CLI
- `grid/data/samples/` and `converters/data/samples/`: generated public sample files

Data flow:
- Load or generate JSON data
- Explore and edit it in the grid
- Keep proposed edits in `changes` until promotion or export
- Write approved changes back to Excel or Obsidian through the CLI tools

## Installation

Python 3.10–3.13 is required (PySide6 does not yet support 3.14+).

Install the root package extras for the desktop shell:

```bash
pip install -e converters
pip install -e ".[desktop]"
```

If you only need the converters:

```bash
cd converters
pip install -e .
```

## Usage

Desktop mode:

```bash
python app/main.py
```

On Windows you can also use:

```bash
start-desktop.bat
```

Converter examples:

```bash
excel-to-json --excel converters/data/samples/sample-data.xlsx --json converters/data/samples/sample-data.json --sheet data

json-to-excel --mode restore --json converters/data/samples/sample-data.json --excel restored.xlsx --sheet data

json-to-excel --mode changes --json converters/data/samples/sample-data.json --excel updated.xlsx --sheet data

json-to-excel --mode confirm-html --json converters/data/samples/sample-data.json --excel review.xlsx --sheet data
```

Desktop tools are declared in `app/tools.yaml`. The shipped configuration includes Excel and Obsidian actions and can be extended with additional CLI-backed tools by editing that YAML file.

## Configuration

- `app/tools.yaml`: desktop tool definitions, tabs, controls, commands, and result handling
- `pyproject.toml`: optional dependency groups such as `desktop`
- `grid/data/samples/` and `converters/data/samples/`: sample input/output locations used by the exported bundle

The shared dataset format uses a top-level `data` array and may also include `schema`, `view`, `meta`, and `changes` sections. Pending edits remain in `changes` until promoted or written back through a converter.

## Development

- Grid source lives in `grid/src/`
- Desktop bridge and shell live in `app/`
- Converter source lives in `converters/src/`
- Public bundle export is produced by `tools/export_public_bundle.py`

To regenerate a clean public bundle:

```bash
python tools/export_public_bundle.py D:/dev/_export/data-visualization-grid --force
```

## Limitations

- This bundle does not include the internal FastAPI server mode
- Desktop mode requires the local CLI entry points to be installed in the active Python environment
- Some frontend libraries are still loaded from CDNs, so full offline use is not guaranteed

## Roadmap

- Vendor frontend CDN dependencies for stronger offline support
- Unify all runtime modes around one YAML tool schema
- Add richer long-running CLI progress and streaming output in desktop mode

## License

MIT
