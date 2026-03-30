import re
from typing import Optional
from .models import ActionItem, BuyItem, IdeaItem, ConsumeItem, BaseItem

PRIORITY_EMOJIS = {
    'highest': '🔺',
    'high': '⏫',
    'medium': '⏺️',
    'low': '⏬',
    'lowest': '🔽',
}

EMOJI_TO_PRIORITY = {v: k for k, v in PRIORITY_EMOJIS.items()}

URGENCY_EMOJIS = {
    'highest': '🚨',
    'high': '🔴',
    'medium': '🟡',
    'low': '🟢',
}

EMOJI_TO_URGENCY = {v: k for k, v in URGENCY_EMOJIS.items()}

PATTERNS = {
    'action': re.compile(r"^(?P<indent>\s*)- \[(?P<state>[ xX])\] (?P<body>.*?)\s*$"),
    'buy': re.compile(r"^(?P<indent>\s*)- \[(?P<state>[ xX])\] (?P<body>.*?)\s*$"),
    'idea': re.compile(r"^(?P<indent>\s*)- (?P<body>.*?)\s*$"),
    'consume': re.compile(r"^(?P<indent>\s*)- \[(?P<state>[ xX])\] (?P<body>.*?)\s*$"),
}

ID_PATTERNS = {
    'action': re.compile(r'🔖\s*(ai-\d+)'),
    'buy': re.compile(r'🔖\s*(bi-\d+)'),
    'idea': re.compile(r'🔖\s*(ii-\d+)'),
    'consume': re.compile(r'🔖\s*(ci-\d+)'),
}

CREATED_RE = re.compile(r'➕\s*(\d{4}-\d{2}-\d{2})')
DUE_RE = re.compile(r'📅\s*(\d{4}-\d{2}-\d{2})')
DONE_RE = re.compile(r'✅\s*(\d{4}-\d{2}-\d{2})')
AREA_RE = re.compile(r'🧭\s*(\[\[([^\]]+)\]\]|([^🔖🧭📅➕✅⏫⏺️⏬🔺🔽🆔#\s]+))')
SHOP_RE = re.compile(r'@([A-Za-z0-9_-]+)')
PRIORITY_RE = re.compile(r'([🔺⏫⏺️⏬🔽])')
URGENCY_RE = re.compile(r'([🚨🔴🟡🟢])')
TRAILING_TAGS_RE = re.compile(r'(?:\s+[#@][^\s]+)+\s*$')
TAG_RE = re.compile(r'(?<!\S)([#@][A-Za-z0-9_-]+(?:/[A-Za-z0-9_-]+)*)\b')
FIELD_TAG_RE = re.compile(r'(?<!\S)#([A-Za-z0-9_-]+)/((?:[A-Za-z0-9_-]+)(?:/[A-Za-z0-9_-]+)*)\b')


def _search_group(pattern: re.Pattern, text: str, group: int = 1) -> Optional[str]:
    match = pattern.search(text)
    if not match:
        return None
    return match.group(group)


def _extract_area(text: str) -> Optional[str]:
    match = AREA_RE.search(text)
    if not match:
        return None
    return (match.group(2) or match.group(3) or '').strip() or None


def _clean_body(body: str, item_type: str, item_id: Optional[str], shop: Optional[str]) -> str:
    clean_text = body
    clean_text = PRIORITY_RE.sub('', clean_text)
    clean_text = URGENCY_RE.sub('', clean_text)
    clean_text = CREATED_RE.sub('', clean_text)
    clean_text = DUE_RE.sub('', clean_text)
    clean_text = DONE_RE.sub('', clean_text)
    clean_text = AREA_RE.sub('', clean_text)
    clean_text = FIELD_TAG_RE.sub('', clean_text)
    clean_text = clean_text.replace('🧭', '')
    if shop and item_type == 'buy':
        clean_text = re.sub(r'\s*@' + re.escape(shop) + r'\b', '', clean_text)
    if item_id:
        clean_text = re.sub(r'\s*🔖\s*' + re.escape(item_id) + r'(?:\s+[#@][^\s]+)*', '', clean_text)
    clean_text = TRAILING_TAGS_RE.sub('', clean_text)
    clean_text = re.sub(r'\s+', ' ', clean_text).strip()
    return clean_text


def extract_tags(text: str, shop: Optional[str] = None):
    """Extract #tag/@tag tokens from a source line, excluding buy-item shop handles."""
    tags = []
    seen = set()
    for tag in TAG_RE.findall(text or ''):
        if shop and tag.lower() == f'@{shop}'.lower():
            continue
        stripped = tag[1:]  # remove leading # or @
        lowered = stripped.lower()
        if '/' in lowered:
            continue
        if lowered in seen:
            continue
        seen.add(lowered)
        tags.append(stripped)
    return tags

ITEM_CLASSES = {
    'action': ActionItem,
    'buy': BuyItem,
    'idea': IdeaItem,
    'consume': ConsumeItem,
}


def parse_line(line: str, file: str, line_no: int, item_type: str) -> Optional[BaseItem]:
    """Parse a markdown line into an item of the given type, or None."""
    pattern = PATTERNS.get(item_type)
    item_class = ITEM_CLASSES.get(item_type)
    if not pattern or not item_class:
        raise ValueError(f"Unknown item type: {item_type}")

    m = pattern.match(line.rstrip('\n'))
    if not m:
        return None

    indent = len(m.group('indent') or '')
    body = (m.group('body') or '').strip()
    raw = line.rstrip('\n')

    created = _search_group(CREATED_RE, raw)
    raw_area = _extract_area(raw)
    priority_emoji = _search_group(PRIORITY_RE, raw)
    priority = EMOJI_TO_PRIORITY.get(priority_emoji) if priority_emoji else None
    urgency_emoji = _search_group(URGENCY_RE, raw)
    urgency = EMOJI_TO_URGENCY.get(urgency_emoji) if urgency_emoji else None
    item_id = _search_group(ID_PATTERNS[item_type], raw)
    managed = bool(item_id)

    field_tags = {prefix.lower(): value for prefix, value in FIELD_TAG_RE.findall(raw)}
    area = field_tags.pop('area', None) or raw_area

    # Extract shop for buy items
    shop = None
    if item_type == 'buy':
        shop = _search_group(SHOP_RE, raw)

    clean_text = _clean_body(body, item_type, item_id, shop)

    base_fields = dict(
        id=item_id, text=clean_text, created=created, area=area,
        priority=priority, urgency=urgency, file=file, line_no=line_no,
        indent=indent, raw=raw, managed=managed,
    )
    parsed_tags = extract_tags(raw, shop)

    if item_type in ('action', 'buy', 'consume'):
        state = m.group('state')
        status = 'done' if state.lower() == 'x' else 'open'
        done = _search_group(DONE_RE, raw)
        if item_type == 'action':
            item = ActionItem(status=status, due=_search_group(DUE_RE, raw), completed=done, **base_fields)
            item.field_tags = field_tags
            item.tags = parsed_tags
            return item
        elif item_type == 'buy':
            item = BuyItem(status=status, completed=done, shop=shop, **base_fields)
            item.field_tags = field_tags
            item.tags = parsed_tags
            return item
        else:
            item = ConsumeItem(status=status, completed=done, **base_fields)
            item.field_tags = field_tags
            item.tags = parsed_tags
            return item
    elif item_type == 'idea':
        item = IdeaItem(**base_fields)
        item.field_tags = field_tags
        item.tags = parsed_tags
        return item

    return None


def render_item(item: BaseItem, item_type: str) -> str:
    """Render an item back to a markdown line."""
    indent = ' ' * item.indent
    parts = []

    if item_type in ('action', 'buy', 'consume'):
        chk = '[x]' if item.status == 'done' else '[ ]'
        parts.append(f"{indent}- {chk} {item.text}")
    else:
        parts.append(f"{indent}- {item.text}")

    if item.priority:
        parts.append(PRIORITY_EMOJIS.get(item.priority, ''))
    if item.urgency:
        parts.append(URGENCY_EMOJIS.get(item.urgency, ''))
    if item.created:
        parts.append(f"➕{item.created}")
    if item_type == 'action' and hasattr(item, 'due') and item.due:
        parts.append(f"📅 {item.due}")
    if hasattr(item, 'status') and item.status == 'done' and hasattr(item, 'completed') and item.completed:
        parts.append(f"✅ {item.completed}")
    if item.area:
        parts.append(f"#area/{item.area}")
    if item.field_tags:
        for prefix, value in item.field_tags.items():
            parts.append(f"#{prefix}/{value}")
    if item_type == 'buy' and hasattr(item, 'shop') and item.shop:
        parts.append(f"@{item.shop}")
    if item.id:
        parts.append(f"🔖 {item.id}")
    if item.tags:
        for tag in item.tags:
            parts.append(f"#{tag}")

    return ' '.join(parts) + '\n'
