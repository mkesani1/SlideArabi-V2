"""
utils.py — OOXML namespace helpers, coordinate math, text direction utilities.

SlideShift v2: Template-First Deterministic RTL Transformation Engine.
"""

from __future__ import annotations

import re
from typing import Dict, Optional, Tuple
from lxml import etree


NSMAP: Dict[str, str] = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
}

A_NS = NSMAP['a']
R_NS = NSMAP['r']
P_NS = NSMAP['p']
C_NS = NSMAP['c']

EMU_PER_INCH: int = 914400
EMU_PER_PT: int = 12700
HUNDREDTHS_PER_PT: int = 100

_ARABIC_RANGES = [
    (0x0600, 0x06FF),
    (0x0750, 0x077F),
    (0xFB50, 0xFDFF),
    (0xFE70, 0xFEFF),
]


def emu_to_inches(emu: int) -> float:
    return emu / EMU_PER_INCH

def emu_to_pt(emu: int) -> float:
    return emu / EMU_PER_PT

def pt_to_emu(pt: float) -> int:
    return int(pt * EMU_PER_PT)

def inches_to_emu(inches: float) -> int:
    return int(inches * EMU_PER_INCH)

def hundredths_pt_to_pt(val: int) -> float:
    return val / HUNDREDTHS_PER_PT

def pt_to_hundredths_pt(pt: float) -> int:
    return int(round(pt * HUNDREDTHS_PER_PT))


def mirror_x(x_emu: int, width_emu: int, slide_width_emu: int) -> int:
    """
    Compute the mirrored X position for RTL layout.
    new_x = S - (x + w)
    """
    return slide_width_emu - (x_emu + width_emu)


def swap_positions(
    shape1_x: int, shape1_w: int,
    shape2_x: int, shape2_w: int,
    slide_width: int,
) -> Tuple[int, int]:
    new_x1 = mirror_x(shape2_x, shape2_w, slide_width)
    new_x2 = mirror_x(shape1_x, shape1_w, slide_width)
    return new_x1, new_x2


def _is_arabic_char(ch: str) -> bool:
    cp = ord(ch)
    return any(lo <= cp <= hi for lo, hi in _ARABIC_RANGES)

def _is_latin_char(ch: str) -> bool:
    return ('A' <= ch <= 'Z') or ('a' <= ch <= 'z')

def has_arabic(text: str) -> bool:
    return any(_is_arabic_char(ch) for ch in text)

def has_latin(text: str) -> bool:
    return any(_is_latin_char(ch) for ch in text)

def is_bidi_text(text: str) -> bool:
    return has_arabic(text) and has_latin(text)

def compute_script_ratio(text: str) -> Dict[str, float]:
    counts = {'arabic': 0, 'latin': 0, 'numeric': 0, 'other': 0}
    total = 0
    for ch in text:
        if ch.isspace():
            continue
        total += 1
        if _is_arabic_char(ch):
            counts['arabic'] += 1
        elif _is_latin_char(ch):
            counts['latin'] += 1
        elif ch.isdigit():
            counts['numeric'] += 1
        else:
            counts['other'] += 1
    if total == 0:
        return {k: 0.0 for k in counts}
    return {k: v / total for k, v in counts.items()}


def qn(tag: str) -> str:
    prefix, local = tag.split(':', 1)
    return f'{{{NSMAP[prefix]}}}{local}'


def ensure_pPr(paragraph_element) -> etree._Element:
    pPr_tag = qn('a:pPr')
    pPr = paragraph_element.find(pPr_tag)
    if pPr is None:
        pPr = etree.Element(pPr_tag)
        paragraph_element.insert(0, pPr)
    return pPr


def set_rtl_on_paragraph(paragraph_element) -> None:
    pPr = ensure_pPr(paragraph_element)
    pPr.set('rtl', '1')


def set_alignment_on_paragraph(paragraph_element, alignment: str) -> None:
    pPr = ensure_pPr(paragraph_element)
    pPr.set('algn', alignment)


def get_placeholder_info(shape) -> Optional[Tuple[str, int]]:
    try:
        if not getattr(shape, 'is_placeholder', False):
            return None
        ph_fmt = shape.placeholder_format
        if ph_fmt is None:
            return None
        ph_type = str(ph_fmt.type).split('.')[-1].lower()
        ph_idx = ph_fmt.idx if ph_fmt.idx is not None else 0
        return ph_type, ph_idx
    except Exception:
        return None


def get_placeholder_info_from_xml(shape_element) -> Optional[Tuple[str, int]]:
    try:
        nv_sp_pr = shape_element.find(qn('p:nvSpPr'))
        if nv_sp_pr is None:
            nv_sp_pr = shape_element.find(qn('p:nvPicPr'))
        if nv_sp_pr is None:
            return None
        ph = nv_sp_pr.find(f'.//{qn("p:ph")}')
        if ph is None:
            return None
        ph_type = ph.get('type', 'body')
        try:
            ph_idx = int(ph.get('idx', '0'))
        except (ValueError, TypeError):
            ph_idx = 0
        return ph_type, ph_idx
    except Exception:
        return None


def set_body_pr_rtl_col(txBody_element) -> None:
    body_pr = txBody_element.find(qn('a:bodyPr'))
    if body_pr is not None:
        body_pr.set('rtlCol', '1')


def set_defRPr_lang(txBody_element, lang: str = 'ar-SA') -> None:
    for defRPr in txBody_element.iter(qn('a:defRPr')):
        defRPr.set('lang', lang)


def iter_paragraphs(txBody_element):
    yield from txBody_element.findall(qn('a:p'))


def iter_runs(paragraph_element):
    yield from paragraph_element.findall(qn('a:r'))


def get_run_text(run_element) -> str:
    t_elem = run_element.find(qn('a:t'))
    if t_elem is None:
        return ''
    return t_elem.text or ''


def set_run_text(run_element, text: str) -> None:
    t_elem = run_element.find(qn('a:t'))
    if t_elem is None:
        t_elem = etree.SubElement(run_element, qn('a:t'))
    t_elem.text = text


def get_or_create_rPr(run_element) -> etree._Element:
    rPr = run_element.find(qn('a:rPr'))
    if rPr is None:
        rPr = etree.Element(qn('a:rPr'))
        run_element.insert(0, rPr)
    return rPr


def set_run_language(run_element, lang: str = 'ar-SA') -> None:
    rPr = get_or_create_rPr(run_element)
    rPr.set('lang', lang)


def bounds_check_emu(value: int, slide_dimension: int, label: str = '') -> bool:
    lower = -200_000
    upper = slide_dimension + 500_000
    return lower <= value <= upper


def clamp_emu(value: int, slide_dimension: int) -> int:
    lower = -200_000
    upper = slide_dimension + 500_000
    return max(lower, min(upper, value))
