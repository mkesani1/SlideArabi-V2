#!/usr/bin/env python3
"""
Process a deck through the SlideShift V2 pipeline AND embedded Excel handler.
This wrapper calls the standard pipeline then applies embedded_excel processing.
"""
import json
import logging
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent / "skills" / "pptx" / "scripts"))

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# Monkey-patch python-pptx for ZIP_STORED
try:
    import pptx.opc.phys_pkg as _phys_pkg
except ModuleNotFoundError:
    _phys_pkg = None
import pptx.opc.serialized as _ser

class _ZipStoredPatch:
    @staticmethod
    def apply():
        from pptx.util import lazyproperty
        def _zipf_stored(self):
            return zipfile.ZipFile(self._pkg_file, 'w', zipfile.ZIP_STORED)
        _ser._ZipPkgWriter._zipf = lazyproperty(_zipf_stored)
        print("  [Patch] ZIP_STORED monkey-patch applied")

_ZipStoredPatch.apply()

from pptx import Presentation
from slideshift_v2.rtl_transforms import MasterLayoutTransformer, SlideContentTransformer
from slideshift_v2.layout_analyzer import LayoutAnalyzer
from slideshift_v2.template_registry import TemplateRegistry
from slideshift_v2.typography import TypographyNormalizer
from slideshift_v2.embedded_excel import EmbeddedExcelHandler


def recompress_pptx(pptx_path: Path) -> None:
    tmp = pptx_path.with_suffix('.pptx.tmp')
    try:
        with zipfile.ZipFile(str(pptx_path), 'r') as zin:
            with zipfile.ZipFile(str(tmp), 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    zout.writestr(item, data)
        shutil.move(str(tmp), str(pptx_path))
        print(f"  [Recompress] Saved compressed: {pptx_path.stat().st_size / 1024:.0f} KB")
    except Exception as e:
        print(f"  [WARN] Recompress failed: {e}")
        if tmp.exists():
            tmp.unlink()


def process_with_excel(input_path, output_path, translations_json):
    input_p = Path(input_path)
    output_p = Path(output_path)
    
    print(f"\n{'='*60}")
    print(f"  Processing: {input_p.name}")
    print(f"  WITH embedded Excel handler")
    print(f"{'='*60}")
    
    # Load translations
    translations = {}
    trans_p = Path(translations_json)
    if trans_p.exists():
        with open(trans_p, 'r', encoding='utf-8') as f:
            translations = json.load(f)
        print(f"  Loaded {len(translations)} translations from {trans_p.name}")
    
    # Copy input to output path
    shutil.copy2(str(input_p), str(output_p))
    
    # Load presentation
    prs = Presentation(str(output_p))
    slide_width = int(prs.slide_width)
    slide_height = int(prs.slide_height)
    print(f"  Slide dimensions: {slide_width}x{slide_height} EMU, {len(prs.slides)} slides")
    
    # Phase 0: Layout analysis
    layout_classifications = {}
    try:
        analyzer = LayoutAnalyzer(prs)
        layout_classifications = analyzer.analyze_all()
        print(f"  Phase 0: Analyzed {len(layout_classifications)} slide layouts")
    except Exception as e:
        print(f"  [WARN] Layout analysis: {e}")
    
    # Build template registry
    registry = None
    try:
        registry = TemplateRegistry(slide_width, slide_height)
    except Exception as e:
        print(f"  [WARN] Template registry: {e}")
    
    # Phase 2: Master & Layout transform
    try:
        ml_transformer = MasterLayoutTransformer(prs, registry)
        master_report = ml_transformer.transform_all_masters()
        layout_report = ml_transformer.transform_all_layouts()
        print(f"  Phase 2: {master_report.total_changes} master + {layout_report.total_changes} layout changes")
    except Exception as e:
        print(f"  [ERROR] Phase 2: {e}")
    
    # Phase 3: Slide content transform WITH translations
    try:
        content_transformer = SlideContentTransformer(
            prs,
            template_registry=registry,
            layout_classifications=layout_classifications,
            translations=translations,
        )
        slide_report = content_transformer.transform_all_slides()
        print(f"  Phase 3: {slide_report.total_changes} slide changes")
        if slide_report.errors:
            for err in slide_report.errors[:5]:
                print(f"    ERROR: {err}")
    except Exception as e:
        print(f"  [ERROR] Phase 3: {e}")
        import traceback; traceback.print_exc()
    
    # Phase 4: Typography normalization
    try:
        normalizer = TypographyNormalizer(prs)
        typo_report = normalizer.normalize_all()
        print(f"  Phase 4: {typo_report.total_changes} typography changes")
    except Exception as e:
        print(f"  [WARN] Phase 4: {e}")
    
    # Phase 4b: Embedded Excel Handler
    print(f"\n  Phase 4b: Embedded Excel Handler...")
    try:
        excel_handler = EmbeddedExcelHandler()
        
        # Step 1: Detect all embedded Excel objects
        embedded_objects = excel_handler.detect_embedded_excel(prs)
        ole_count = sum(1 for o in embedded_objects if o.object_type == 'ole_excel_table')
        chart_count = sum(1 for o in embedded_objects if 'chart' in o.object_type)
        native_table_count = sum(1 for o in embedded_objects if o.object_type == 'native_table')
        print(f"    Detected: {ole_count} OLE Excel tables, {chart_count} charts, {native_table_count} native tables")
        
        for obj in embedded_objects:
            print(f"    -> Slide {obj.slide_number}: {obj.shape_name} ({obj.object_type})"
                  f"{' progId=' + obj.prog_id if obj.prog_id else ''}"
                  f"{' embedded=' + str(obj.is_embedded) if obj.object_type == 'ole_excel_table' else ''}")
        
        # Step 2: Create a translate_fn from the translations dict
        def translate_fn(text):
            if not text or not text.strip():
                return text
            t = text.strip()
            if t in translations:
                return translations[t]
            t_lower = t.lower()
            for k, v in translations.items():
                if k.lower() == t_lower:
                    return v
            return text
        
        # Step 3: Process the presentation with embedded Excel handler
        excel_handler.process_presentation(prs, translate_fn)
        
        # Get report
        report = excel_handler.report
        print(f"    OLE tables found: {report.total_ole_tables_found}, translated: {report.total_ole_tables_translated}")
        print(f"    Charts found: {report.total_charts_found}, translated: {report.total_charts_translated}")
        print(f"    Total cells translated: {report.total_cells_translated}")
        if report.errors:
            print(f"    Errors ({len(report.errors)}):")
            for err in report.errors[:10]:
                print(f"      ERROR: {err}")
        if report.warnings:
            print(f"    Warnings ({len(report.warnings)}):")
            for warn in report.warnings[:10]:
                print(f"      WARN: {warn}")
    except Exception as e:
        print(f"  [ERROR] Phase 4b Embedded Excel: {e}")
        import traceback; traceback.print_exc()
    
    # Save
    try:
        prs.save(str(output_p))
        size_kb = output_p.stat().st_size / 1024
        print(f"\n  Saved: {output_p.name} ({size_kb:.0f} KB)")
    except Exception as e:
        print(f"  [ERROR] Save failed: {e}")
        return False
    
    # Recompress
    try:
        recompress_pptx(output_p)
    except Exception as e:
        print(f"  [WARN] Recompress: {e}")
    
    print(f"  DONE: {input_p.name}")
    return True


if __name__ == '__main__':
    if len(sys.argv) != 4:
        print(f"Usage: {sys.argv[0]} <input.pptx> <output.pptx> <translations.json>")
        sys.exit(1)
    
    success = process_with_excel(sys.argv[1], sys.argv[2], sys.argv[3])
    sys.exit(0 if success else 1)
