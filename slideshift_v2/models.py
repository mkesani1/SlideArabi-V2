"""
SlideShift v2 — Data Models

Frozen dataclasses representing the fully-resolved OOXML presentation structure.
Every effective property has a concrete value — no None for properties that should
always be resolvable (font size, font name, bold, italic, alignment, RTL, level).

Design principles:
1. Immutability: All resolved models are frozen dataclasses with tuple collections.
2. No None for effective values: The PropertyResolver guarantees concrete values
   by walking the 7-level OOXML inheritance chain.
3. Provenance: source_level / source_font_size_level tracks where each value
   was resolved from, enabling debugging and auditing.
4. Separation: Resolved models (read-only snapshots) are separate from
   TransformPlan/TransformAction (mutable planning models for Phase 2/3).

OOXML Constants:
- Font sizes in hundredths of a point (e.g., 1800 = 18pt)
- EMU (English Metric Units): 1 inch = 914400 EMU, 1 pt = 12700 EMU
- Placeholder types from ST_PlaceholderType
- Layout types from ST_SlideLayoutType
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

A_NS = NSMAP['a']
P_NS = NSMAP['p']
R_NS = NSMAP['r']

DEFAULT_FONT_SIZE_PT = 18.0
DEFAULT_FONT_NAME = 'Calibri'
DEFAULT_BOLD = False
DEFAULT_ITALIC = False
DEFAULT_UNDERLINE = False
DEFAULT_ALIGNMENT = 'l'
DEFAULT_RTL = False
DEFAULT_LEVEL = 0
DEFAULT_ROTATION = 0.0

VALID_ALIGNMENTS = frozenset({'l', 'r', 'ctr', 'just', 'dist'})
VALID_SHAPE_TYPES = frozenset({
    'placeholder', 'textbox', 'picture', 'chart', 'table',
    'group', 'connector', 'freeform', 'ole', 'smartart', 'media',
})
VALID_PLACEHOLDER_TYPES = frozenset({
    'title', 'body', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum',
    'pic', 'chart', 'tbl', 'dgm', 'media', 'clipArt', 'obj',
})
VALID_LAYOUT_TYPES = frozenset({
    'title', 'tx', 'twoColTx', 'obj', 'secHead', 'blank', 'tbl',
    'chart', 'txAndChart', 'picTx', 'cust', 'titleOnly', 'twoObj',
    'objTx', 'txAndObj', 'dgm', 'txOverObj', 'objOverTx',
    'fourObj', 'objAndTx', 'vertTx', 'vertTitleAndTx', 'clipArtAndTx',
    'txAndClipArt', 'mediaAndTx', 'txAndMedia', 'objAndTwoObj',
    'twoObjAndObj', 'twoObjOverTx', 'txAndTwoObj', 'twoTxTwoObj',
    'txOverObj2', 'objOverTx2',
})
VALID_SOURCE_LEVELS = frozenset({'master', 'layout', 'slide'})
INHERITANCE_LEVELS = (
    'run', 'paragraph', 'textframe', 'shape', 'layout', 'master', 'txstyles', 'default',
)


@dataclass(frozen=True)
class ResolvedRun:
    text: str
    effective_font_size_pt: float
    effective_font_name: str
    effective_bold: bool
    effective_italic: bool
    effective_color: Optional[str]
    effective_underline: bool
    source_font_size_level: str


@dataclass(frozen=True)
class ResolvedParagraph:
    runs: Tuple[ResolvedRun, ...]
    effective_alignment: str
    effective_rtl: bool
    effective_level: int
    effective_bullet_type: Optional[str]
    effective_line_spacing: Optional[float]
    effective_space_before: Optional[float]
    effective_space_after: Optional[float]


@dataclass(frozen=True)
class ResolvedShape:
    shape_id: int
    shape_name: str
    shape_type: str
    placeholder_type: Optional[str]
    placeholder_idx: Optional[int]
    x_emu: int
    y_emu: int
    width_emu: int
    height_emu: int
    rotation_degrees: float
    paragraphs: Tuple[ResolvedParagraph, ...]
    is_master_inherited: bool
    source_level: str
    has_local_position_override: bool
    has_text: bool
    original_xml_element: Any = field(compare=False, hash=False, default=None)

    @property
    def full_text(self) -> str:
        parts = []
        for para in self.paragraphs:
            para_text = ''.join(r.text for r in para.runs)
            parts.append(para_text)
        return '\n'.join(parts)

    @property
    def is_placeholder(self) -> bool:
        return self.shape_type == 'placeholder'


@dataclass(frozen=True)
class ResolvedLayout:
    layout_name: str
    layout_type: str
    master_index: int
    placeholders: Tuple[ResolvedShape, ...]
    freeform_shapes: Tuple[ResolvedShape, ...]


@dataclass(frozen=True)
class ResolvedMaster:
    master_name: str
    master_index: int
    placeholders: Tuple[ResolvedShape, ...]
    freeform_shapes: Tuple[ResolvedShape, ...]
    tx_styles: Dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ResolvedSlide:
    slide_number: int
    layout_name: str
    layout_type: str
    layout_index: int
    master_index: int
    shapes: Tuple[ResolvedShape, ...]


@dataclass(frozen=True)
class ResolvedPresentation:
    slide_width_emu: int
    slide_height_emu: int
    masters: Tuple[ResolvedMaster, ...]
    layouts: Tuple[ResolvedLayout, ...]
    slides: Tuple[ResolvedSlide, ...]

    @property
    def total_shapes(self) -> int:
        return sum(len(s.shapes) for s in self.slides)

    @property
    def total_slides(self) -> int:
        return len(self.slides)


VALID_ACTION_TYPES = frozenset({
    'mirror', 'swap', 'keep', 'right_align', 'center_align', 'set_rtl',
    'set_font', 'resize_font', 'reverse_columns', 'reverse_axes',
    'set_language', 'remove_position',
})


@dataclass
class TransformAction:
    shape_id: int
    action_type: str
    params: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if self.action_type not in VALID_ACTION_TYPES:
            raise ValueError(
                f"Invalid action_type '{self.action_type}'. "
                f"Must be one of: {sorted(VALID_ACTION_TYPES)}"
            )


@dataclass
class TransformPlan:
    slide_actions: Dict[int, List[TransformAction]] = field(default_factory=dict)
    master_actions: Dict[int, List[TransformAction]] = field(default_factory=dict)
    layout_actions: Dict[Tuple[int, int], List[TransformAction]] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def add_slide_action(self, slide_number: int, action: TransformAction) -> None:
        if slide_number not in self.slide_actions:
            self.slide_actions[slide_number] = []
        self.slide_actions[slide_number].append(action)

    def add_master_action(self, master_index: int, action: TransformAction) -> None:
        if master_index not in self.master_actions:
            self.master_actions[master_index] = []
        self.master_actions[master_index].append(action)

    def add_layout_action(self, master_index: int, layout_index: int, action: TransformAction) -> None:
        key = (master_index, layout_index)
        if key not in self.layout_actions:
            self.layout_actions[key] = []
        self.layout_actions[key].append(action)

    @property
    def total_actions(self) -> int:
        count = sum(len(v) for v in self.slide_actions.values())
        count += sum(len(v) for v in self.master_actions.values())
        count += sum(len(v) for v in self.layout_actions.values())
        return count


VALID_SEVERITIES = frozenset({'error', 'warning', 'info'})


@dataclass(frozen=True)
class ValidationIssue:
    severity: str
    slide_number: int
    shape_id: Optional[int]
    issue_type: str
    message: str
    expected_value: Any = None
    actual_value: Any = None


@dataclass(frozen=True)
class ValidationReport:
    issues: Tuple[ValidationIssue, ...]
    total_shapes_checked: int = 0
    total_slides_checked: int = 0

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == 'error')

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == 'warning')

    @property
    def info_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == 'info')

    @property
    def has_errors(self) -> bool:
        return self.error_count > 0

    @property
    def passed(self) -> bool:
        return not self.has_errors
