from pathlib import Path
from typing import List
from .parsers import parse_line, render_item
from .files import atomic_write_text
from .models import BaseItem


class ItemRepository:
    """Shared repository for all item types — scans vault markdown files."""

    def __init__(self, vault: str, search_paths: List[str], item_type: str):
        self.vault = Path(vault)
        self.search_paths = search_paths
        self.item_type = item_type

    def scan_items(self, include_unmanaged=False):
        items = []
        for sp in self.search_paths:
            folder = self.vault / sp
            if not folder.exists():
                continue
            for p in folder.rglob('*.md'):
                rel = str(p.relative_to(self.vault))
                with open(p, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                for i, l in enumerate(lines, start=1):
                    item = parse_line(l, rel, i, self.item_type)
                    if not item:
                        continue
                    if not include_unmanaged and not item.managed:
                        continue
                    items.append((item, p, lines))
        return items

    def find_by_id(self, item_id: str):
        found = []
        for item, p, lines in self.scan_items(include_unmanaged=True):
            if item.id == item_id:
                found.append((item, p, lines))
        return found

    def write_file(self, path: Path, lines):
        atomic_write_text(path, ''.join(lines))

    def render_item(self, item: BaseItem) -> str:
        return render_item(item, self.item_type)
