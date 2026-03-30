from copy import deepcopy
from pathlib import Path

import yaml


CONFIG_FILENAME = '.item_manager_config.yaml'

DEFAULT_CONFIG = {
    'vault': '',
    'search_paths': ['0_Inbox', '1_Projects', '2_Areas', '3_Resources', '4_Archive', '5_Log', '6_Reviews'],

    'action': {
        'creation': {'file': '0_Inbox/action items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ai', 'start': 1000},
    },
    'buy': {
        'creation': {'file': '0_Inbox/buy items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'bi', 'start': 1000},
    },
    'idea': {
        'creation': {'file': '0_Inbox/idea items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ii', 'start': 1000},
    },
    'consume': {
        'creation': {'file': '0_Inbox/consume items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ci', 'start': 1000},
    },
}

TEMPLATE_CONFIG = {
    'vault': '/path/to/obsidian/vault',
    'search_paths': ['0_Inbox', '1_Projects', '2_Areas', '3_Resources', '4_Archive', '5_Log', '6_Reviews'],
    'action': {
        'creation': {'file': '0_Inbox/action items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ai', 'start': 1000},
    },
    'buy': {
        'creation': {'file': '0_Inbox/buy items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'bi', 'start': 1000},
    },
    'idea': {
        'creation': {'file': '0_Inbox/idea items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ii', 'start': 1000},
    },
    'consume': {
        'creation': {'file': '0_Inbox/consume items.md', 'inbox_header': '# Inbox'},
        'id': {'prefix': 'ci', 'start': 1000},
    },
}


def _merge_config(base: dict, override: dict) -> dict:
    result = deepcopy(base)
    result.update(override)
    for item_type in ('action', 'buy', 'idea', 'consume'):
        if item_type in override and item_type in result:
            for section in ('creation', 'id'):
                if section in override[item_type]:
                    result[item_type][section].update(override[item_type][section])
    return result


def _write_template_config(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        yaml.safe_dump(TEMPLATE_CONFIG, f, sort_keys=False, allow_unicode=True)


def load_config(path: Path = None):
    if path is None:
        path = Path.home() / CONFIG_FILENAME

    created_template = False
    warnings = []

    if path.exists():
        try:
            with open(path, 'r', encoding='utf-8') as f:
                cfg = yaml.safe_load(f) or {}
        except Exception:
            cfg = {}
    else:
        _write_template_config(path)
        created_template = True
        warnings.append(f'Created template config at {path}')
        cfg = {}

    result = _merge_config(DEFAULT_CONFIG, cfg) if cfg else deepcopy(DEFAULT_CONFIG)
    result['_meta'] = {
        'config_path': str(path),
        'created_template': created_template,
        'warnings': warnings,
    }
    return result


def get_item_config(cfg, item_type: str):
    if item_type in cfg:
        return cfg[item_type]
    return {
        'creation': cfg.get('creation', {}),
        'id': cfg.get('id', {}),
    }
