from dataclasses import dataclass, asdict, field
from typing import Optional


@dataclass
class BaseItem:
    """Base class for all item types."""
    id: Optional[str]
    text: str
    created: Optional[str]
    area: Optional[str]
    priority: Optional[str]
    urgency: Optional[str]
    file: str
    line_no: int
    indent: int
    raw: str
    managed: bool
    field_tags: dict = field(init=False, default_factory=dict)
    tags: Optional[list] = field(init=False, default=None)

    def to_dict(self):
        return asdict(self)


@dataclass
class ActionItem(BaseItem):
    """Action item with checkbox status and due/done dates."""
    status: str
    due: Optional[str]
    completed: Optional[str]


@dataclass
class BuyItem(BaseItem):
    """Buy item with checkbox status and optional shop."""
    status: str
    completed: Optional[str]
    shop: Optional[str] = None


@dataclass
class IdeaItem(BaseItem):
    """Idea item — no checkbox, no status."""
    pass


@dataclass
class ConsumeItem(BaseItem):
    """Consume item with checkbox status."""
    status: str
    completed: Optional[str]
