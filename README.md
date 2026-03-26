# Structured Data Grid Toolkit

## Overview

This repository contains a lightweight toolkit for exploring and transforming structured datasets. It includes a browser-based grid app for viewing JSON data in configurable 2D layouts, and a Python converter package for moving data between Excel workbooks and the shared JSON dataset format.

## Components

- `grid/` - browser-based grid viewer with card and table modes, filters, tags, and pending-change support
- `converters/` - Python CLI for Excel-to-JSON export and JSON-to-Excel restore, merge, and review flows

## Features

### Grid

- Card and table rendering modes
- Sticky row and column headers
- Per-field slicers with search
- Tag display and simple grouping support
- Pending-change overlay model for edits and creates

### Converters

- `excel-to-json` and `json-to-excel` CLI commands
- Restore, merge, and review-oriented export modes
- Structured field handling for comments and multi-value fields
- Sample workbook and JSON files for local testing

## Repository Contents

- Core source files for both projects
- Minimal package metadata and license files
- Sample data for trying the grid and converter flows locally

## Installation

### Grid

Open `grid/src/index.html` directly in a browser for basic local use. If browser file restrictions interfere with loading local JSON files, serve the repository through a small local HTTP server.

### Converters

```bash
cd converters
pip install -e .
```

## Quick Start

### Grid

Open `grid/src/index.html` in a modern browser.

Use `grid/data/samples/sample-data.json` for a small example dataset.

### Converters

```bash
cd converters
excel-to-json --excel data/samples/sample-data.xlsx --json data/samples/sample-data.json --sheet data

# Restore baseline rows into a workbook
json-to-excel --mode restore --json data/samples/sample-data.json --excel restored.xlsx --sheet data

# Apply pending changes from the dataset
json-to-excel --mode changes --json data/samples/sample-data.json --excel updated.xlsx --sheet data

# Review changes before applying them
json-to-excel --mode confirm-html --json data/samples/sample-data.json --excel review.xlsx --sheet data
```

## JSON Format

The shared dataset format is a JSON object with a top-level `data` array and optional `schema`, `view`, `meta`, and `changes` sections.

Simple example:

```json
{
	"data": [
		{
			"id": "HOME-001",
			"title": "Build seed trays",
			"status": "In Progress",
			"category": "Gardening",
			"tags": ["outdoor", "weekend"]
		}
	],
	"schema": {
		"fields": {
			"id": {"type": "scalar", "required": true},
			"title": {"type": "scalar", "required": true},
			"status": {"type": "scalar", "required": false},
			"category": {"type": "scalar", "required": false},
			"tags": {"type": "multi-value", "required": false}
		}
	},
	"view": {
		"axisSelections": {
			"x": "status",
			"y": "category",
			"title": "title"
		}
	},
	"meta": {
		"datasetName": "Sample Planning Board"
	},
	"changes": {
		"version": "1",
		"rows": [
			{
				"changeId": "chg-001",
				"action": "update",
				"target": {
					"itemId": "HOME-001"
				},
				"baseline": {
					"status": "Planned"
				},
				"proposed": {
					"status": "In Progress"
				}
			}
		]
	}
}
```

Notes:

- `data` contains the baseline items rendered by the grid
- `schema` can describe field types and constraints when needed
- `view` stores UI state such as axis selections and other saved layout settings
- `meta` stores lightweight dataset-level information
- `changes` can store pending edits or creates without modifying the baseline rows

## Development

- Grid source: `grid/src/`
- Converter source: `converters/src/data2json_converters/`
- Export script: `tools/export_public_bundle.py`

## License

MIT
