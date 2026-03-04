"""
layout_analyzer.py — Slide layout type classifier.

SlideShift v2: Template-First Deterministic RTL Transformation Engine.

Classifies each slide's layout into a canonical ST_SlideLayoutType for
deterministic transformation by the TemplateRegistry.

Classification strategy (priority order):
1. Read explicit `type` attribute from slideLayout XML element.
2. Infer from placeholder configuration using heuristic rules.
3. Fall back to spatial analysis (e.g., detect two-column by geometry).
4. Flag as 'cust' (custom) if no confident match — may require AI.

All 36 standard OOXML ST_SlideLayoutType values are supported.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class SlideLayoutType(str, Enum):
    """All 36 standard OOXML ST_SlideLayoutType values."""
    TITLE = 'title'
    TX = 'tx'
    TWO_COL_TX = 'twoColTx'
    TBL = 'tbl'
    TX_AND_CHART = 'txAndChart'
    CHART_AND_TX = 'chartAndTx'
    DGM = 'dgm'
    CHART = 'chart'
    TX_AND_CLIP_ART = 'txAndClipArt'
    CLIP_ART_AND_TX = 'clipArtAndTx'
    TITLE_ONLY = 'titleOnly'
    BLANK = 'blank'
    TX_AND_OBJ = 'txAndObj'
    OBJ_AND_TX = 'objAndTx'
    OBJ_ONLY = 'objOnly'
    OBJ = 'obj'
    TX_AND_MEDIA = 'txAndMedia'
    MEDIA_AND_TX = 'mediaAndTx'
    OBJ_TX = 'objTx'
    TX_OBJ = 'txObj'
    OBJ_OVER_TX = 'objOverTx'
    TX_OVER_OBJ = 'txOverObj'
    TX_AND_TWO_OBJ = 'txAndTwoObj'
    TWO_OBJ_AND_TX = 'twoObjAndTx'
    TWO_OBJ_OVER_TX = 'twoObjOverTx'
    FOUR_OBJ = 'fourObj'
    TWO_TX_TWO_OBJ = 'twoTxTwoObj'
    TWO_OBJ_AND_OBJ = 'twoObjAndObj'
    SEC_HEAD = 'secHead'
    TWO_OBJ = 'twoObj'
    OBJ_AND_TWO_OBJ = 'objAndTwoObj'
    PIC_TX = 'picTx'
    VERT_TX = 'vertTx'
    VERT_TITLE_AND_TX = 'vertTitleAndTx'
    VERT_TITLE_AND_TX_OVER_CHART = 'vertTitleAndTxOverChart'
    CUST = 'cust'


_VALID_LAYOUT_TYPES = frozenset(lt.value for lt in SlideLayoutType)

_PH_TYPE_MAP = {
    'title': 'title',
    'center_title': 'ctrTitle',
    'subtitle': 'subTitle',
    'body': 'body',
    'object': 'obj',
    'chart': 'chart',
    'table': 'tbl',
    'org_chart': 'dgm',
    'slide_number': 'sldNum',
    'date': 'dt',
    'footer': 'ftr',
    'slide_image': 'pic',
    'media_clip': 'media',
    'clip_art': 'clipArt',
    'bitmap': 'pic',
    'picture': 'pic',
    'ctrtitle': 'ctrTitle',
    'sldnum': 'sldNum',
    'dt': 'dt',
    'ftr': 'ftr',
    'tbl': 'tbl',
    'dgm': 'dgm',
    'pic': 'pic',
    'obj': 'obj',
    'media': 'media',
    'clipart': 'clipArt',
}

_DECORATIVE_PH_TYPES = frozenset({'dt', 'ftr', 'sldNum'})


def _normalise_ph_type(raw_type) -> str:
    if raw_type is None:
        return 'body'
    s = str(raw_type).split('.')[-1].lower()
    return _PH_TYPE_MAP.get(s, s)


@dataclass
class LayoutClassification:
    slide_number: int
    layout_name: str
    explicit_type: Optional[str]
    resolved_type: str
    confidence: float
    placeholder_summary: Dict[str, int]
    requires_ai_classification: bool


_AI_CONFIDENCE_THRESHOLD = 0.7


class LayoutAnalyzer:
    """Classify each slide's layout into a canonical ST_SlideLayoutType."""

    def __init__(self, presentation):
        self._prs = presentation
        self._slide_width = presentation.slide_width
        self._layout_cache: Dict[int, Tuple[str, float]] = {}

    def analyze_all(self) -> Dict[int, LayoutClassification]:
        results: Dict[int, LayoutClassification] = {}
        for idx, slide in enumerate(self._prs.slides):
            slide_number = idx + 1
            classification = self.classify_slide(slide, slide_number)
            results[slide_number] = classification
            logger.debug(
                'Slide %d: layout=%r type=%s (%.0f%% confidence)',
                slide_number, classification.layout_name,
                classification.resolved_type, classification.confidence * 100,
            )
        logger.info('LayoutAnalyzer: classified %d slides', len(results))
        return results

    def classify_slide(self, slide, slide_number: int) -> LayoutClassification:
        layout = slide.slide_layout
        layout_name = layout.name or ''
        resolved_type, confidence = self.classify_layout(layout)
        ph_summary = self._get_placeholder_summary(slide)
        explicit_type = self._get_explicit_type(layout)
        requires_ai = confidence < _AI_CONFIDENCE_THRESHOLD
        return LayoutClassification(
            slide_number=slide_number,
            layout_name=layout_name,
            explicit_type=explicit_type,
            resolved_type=resolved_type,
            confidence=confidence,
            placeholder_summary=ph_summary,
            requires_ai_classification=requires_ai,
        )

    def classify_layout(self, layout) -> Tuple[str, float]:
        cache_key = id(layout)
        if cache_key in self._layout_cache:
            return self._layout_cache[cache_key]
        explicit = self._get_explicit_type(layout)
        if explicit is not None and explicit != 'cust':
            result = (explicit, 1.0)
            self._layout_cache[cache_key] = result
            return result
        inferred_type, confidence = self._infer_type_from_placeholders(layout)
        result = (inferred_type, confidence)
        self._layout_cache[cache_key] = result
        return result

    def _get_explicit_type(self, layout) -> Optional[str]:
        try:
            raw = layout._element.get('type')
            if raw is None:
                return None
            if raw in _VALID_LAYOUT_TYPES:
                return raw
            logger.debug('Unknown layout type attribute: %r', raw)
            return raw
        except Exception:
            return None

    def _infer_type_from_placeholders(self, layout) -> Tuple[str, float]:
        ph_summary = self._get_placeholder_summary(layout)
        ctr_title = ph_summary.get('ctrTitle', 0)
        sub_title = ph_summary.get('subTitle', 0)
        title_count = ph_summary.get('title', 0)
        body_count = ph_summary.get('body', 0)
        obj_count = ph_summary.get('obj', 0)
        chart_count = ph_summary.get('chart', 0)
        tbl_count = ph_summary.get('tbl', 0)
        pic_count = ph_summary.get('pic', 0)
        dgm_count = ph_summary.get('dgm', 0)
        media_count = ph_summary.get('media', 0)
        clip_art_count = ph_summary.get('clipArt', 0)
        total_structural = sum(v for k, v in ph_summary.items() if k not in _DECORATIVE_PH_TYPES)

        if total_structural == 0:
            return ('blank', 0.95)
        if ctr_title >= 1 and sub_title >= 1:
            return ('title', 0.95)
        if (title_count >= 1 and body_count == 0 and obj_count == 0
                and chart_count == 0 and tbl_count == 0 and pic_count == 0
                and dgm_count == 0 and media_count == 0 and clip_art_count == 0
                and ctr_title == 0):
            return ('titleOnly', 0.9)
        if title_count >= 1 and tbl_count >= 1:
            return ('tbl', 0.9)
        if title_count >= 1 and chart_count >= 1 and body_count == 0:
            return ('chart', 0.9)
        if title_count >= 1 and dgm_count >= 1:
            return ('dgm', 0.85)
        if title_count >= 1 and body_count >= 1 and chart_count >= 1:
            return ('txAndChart', 0.85)
        if title_count >= 1 and body_count >= 1 and media_count >= 1:
            return ('txAndMedia', 0.8)
        if title_count >= 1 and body_count >= 1 and clip_art_count >= 1:
            return ('txAndClipArt', 0.8)
        if title_count >= 1 and pic_count >= 1:
            return ('picTx', 0.85)
        if title_count >= 1 and body_count == 2:
            placeholders = self._collect_body_placeholders(layout)
            if self._detect_two_column_spatial(placeholders, self._slide_width):
                return ('twoColTx', 0.85)
            return ('twoColTx', 0.75)
        if title_count >= 1 and obj_count == 2 and body_count == 0:
            return ('twoObj', 0.85)
        if title_count >= 1 and body_count >= 1 and obj_count == 2:
            return ('txAndTwoObj', 0.8)
        if title_count >= 1 and body_count >= 1 and obj_count >= 1:
            return ('txAndObj', 0.85)
        if obj_count == 4:
            return ('fourObj', 0.85)
        if title_count >= 1 and body_count == 1:
            return ('tx', 0.9)
        if title_count >= 1 and obj_count == 1:
            return ('obj', 0.85)
        if obj_count >= 1 and title_count == 0 and ctr_title == 0:
            return ('objOnly', 0.8)
        if body_count >= 1 and title_count == 0 and ctr_title == 0:
            return ('tx', 0.6)
        logger.debug('Could not confidently classify layout. Placeholder summary: %s', ph_summary)
        return ('cust', 0.4)

    def _get_placeholder_summary(self, layout_or_slide) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        try:
            placeholders = layout_or_slide.placeholders
        except Exception:
            return counts
        for ph in placeholders:
            try:
                raw_type = ph.placeholder_format.type
                ph_type = _normalise_ph_type(raw_type)
                counts[ph_type] = counts.get(ph_type, 0) + 1
            except Exception:
                logger.debug('Could not read placeholder type for shape %r', getattr(ph, 'name', '?'))
        return counts

    def _collect_body_placeholders(self, layout) -> List[Any]:
        bodies = []
        try:
            for ph in layout.placeholders:
                try:
                    raw_type = ph.placeholder_format.type
                    ph_type = _normalise_ph_type(raw_type)
                    if ph_type == 'body':
                        bodies.append(ph)
                except Exception:
                    continue
        except Exception:
            pass
        return bodies

    def _detect_two_column_spatial(self, placeholders: List[Any], slide_width: int) -> bool:
        if len(placeholders) < 2:
            return False
        try:
            midpoint = slide_width // 2
            left_found = False
            right_found = False
            for ph in placeholders:
                ph_left = ph.left
                ph_width = ph.width
                if ph_left is None or ph_width is None:
                    continue
                ph_center = ph_left + ph_width // 2
                if ph_center < midpoint:
                    left_found = True
                else:
                    right_found = True
            return left_found and right_found
        except Exception:
            return False

    def get_layout_type_for_slide(self, slide_number: int) -> Optional[str]:
        try:
            slides_list = list(self._prs.slides)
            if slide_number < 1 or slide_number > len(slides_list):
                return None
            slide = slides_list[slide_number - 1]
            layout = slide.slide_layout
            resolved_type, _ = self.classify_layout(layout)
            return resolved_type
        except Exception:
            return None

    def get_all_layout_types(self) -> Dict[str, str]:
        result: Dict[str, str] = {}
        try:
            for master in self._prs.slide_masters:
                for layout in master.slide_layouts:
                    name = layout.name or '(unnamed)'
                    lt, _ = self.classify_layout(layout)
                    result[name] = lt
        except Exception:
            pass
        return result
