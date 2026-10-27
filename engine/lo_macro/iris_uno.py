# -*- coding: utf-8 -*-
"""
iris_uno.py — LibreOffice 안에서 도는 UNO 작업 모음(macOS 엔진, IRIS 문서 MCP 공용).

왜 매크로인가: LibreOffice 26 의 내장 Python 은 실행 제약(Launch Constraint)으로 soffice 만 띄울 수 있다 —
밖에서 `python` 으로 부르면 커널이 죽인다. 그래서 이 파일을 사용자 프로필의 Scripts/python 에 두고
`soffice --headless "vnd.sun.star.script:iris_uno.py$run?language=Python&location=user"` 로 soffice 가
스스로 부르게 한다(engine/lo_bridge.py). 작업 내용은 환경변수 IRIS_LO_JOB(JSON 파일)로 받고, 결과는
IRIS_LO_RESULT(JSON 파일)에 쓴다. 무슨 일이 있어도 끝에 데스크톱을 닫는다(soffice 가 남지 않게).

이 파일은 엑셀·PPT·워드 MCP 세 저장소에 **글자까지 같은** 사본으로 들어 있다(각 저장소는 따로 배포된다).
"""
import json
import os
import traceback

import uno
from com.sun.star.beans import PropertyValue

# 확장자 → 저장 필터(문서 종류별)
CALC_FILTERS = {
    ".xlsx": "Calc Office Open XML", ".xlsm": "Calc Office Open XML", ".xls": "MS Excel 97",
    ".ods": "calc8", ".csv": "Text - txt - csv (StarCalc)", ".pdf": "calc_pdf_Export",
}
WRITER_FILTERS = {
    ".docx": "MS Word 2007 XML", ".docm": "MS Word 2007 XML", ".doc": "MS Word 97", ".odt": "writer8",
    ".rtf": "Rich Text Format", ".txt": "Text (encoded)", ".html": "HTML (StarWriter)", ".pdf": "writer_pdf_Export",
}
IMPRESS_FILTERS = {
    ".pptx": "Impress Office Open XML", ".ppt": "MS PowerPoint 97", ".odp": "impress8", ".pdf": "impress_pdf_Export",
}


def _props(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name = k
        p.Value = v
        out.append(p)
    return tuple(out)


def _url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


class Office:
    def __init__(self):
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def load(self, path, read_only=False):
        if not os.path.exists(path):
            raise FileNotFoundError("원본을 찾을 수 없습니다: %s" % path)
        # Hidden: 창 없음 · UpdateDocMode 0: 외부 링크 갱신 물음 없이 · MacroExecutionMode 0: 문서 매크로는 돌리지 않음
        doc = self.desktop.loadComponentFromURL(_url(path), "_blank", 0, _props(
            Hidden=True, ReadOnly=bool(read_only), UpdateDocMode=0, MacroExecutionMode=0))
        if doc is None:
            raise RuntimeError("LibreOffice 가 문서를 열지 못했습니다: %s" % path)
        return doc

    def store(self, doc, out_path, filters, filter_data=None):
        ext = os.path.splitext(out_path)[1].lower()
        name = filters.get(ext)
        if not name:
            raise ValueError("지원하지 않는 저장 형식입니다: %s" % ext)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        kw = {"FilterName": name, "Overwrite": True}
        if filter_data:
            kw["FilterData"] = uno.Any("[]com.sun.star.beans.PropertyValue", _props(**filter_data))
        doc.storeToURL(_url(out_path), _props(**kw))
        if not os.path.exists(out_path):
            raise RuntimeError("저장 실패: %s" % out_path)
        return os.path.abspath(out_path)

    def dispatch(self, doc, command, **args):
        helper = self.smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", self.ctx)
        frame = doc.getCurrentController().getFrame()
        return helper.executeDispatch(frame, command, "", 0, _props(**args))


def _close(doc):
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            pass


# ---------------------------------------------------------------- Calc

def _col_name(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _a1(addr):
    a = "$%s$%d" % (_col_name(addr.StartColumn), addr.StartRow + 1)
    if addr.EndColumn != addr.StartColumn or addr.EndRow != addr.StartRow:
        a += ":$%s$%d" % (_col_name(addr.EndColumn), addr.EndRow + 1)
    return a


def _pick_sheet(doc, sheet):
    sheets = doc.getSheets()
    if sheet is None or str(sheet).strip() == "":
        return sheets.getByIndex(0)
    s = str(sheet).strip()
    if s.isdigit():
        return sheets.getByIndex(int(s) - 1)
    if not sheets.hasByName(s):
        raise ValueError("시트를 찾을 수 없습니다: %s" % s)
    return sheets.getByName(s)


def _cell_value(cell):
    from com.sun.star.table.CellContentType import EMPTY, VALUE, TEXT, FORMULA
    t = cell.getType()
    if t == EMPTY:
        return None
    if t == VALUE:
        return cell.getValue()
    if t == TEXT:
        return cell.getString()
    # FORMULA: 결과의 종류(1=값, 2=글자, 4=오류 — com.sun.star.sheet.FormulaResult)
    if cell.getError():
        return cell.getString() or "#ERR%d" % cell.getError()
    rt = cell.FormulaResultType2
    return cell.getString() if rt == 2 else cell.getValue()


def calc_recalc(office, job):
    doc = office.load(job["src"])
    try:
        doc.calculateAll()
        return {"out_path": office.store(doc, job["out"], CALC_FILTERS)}
    finally:
        _close(doc)


def calc_read(office, job):
    doc = office.load(job["path"], read_only=True)
    try:
        if job.get("recalc", True):
            doc.calculateAll()
        ws = _pick_sheet(doc, job.get("sheet"))
        if job.get("range"):
            rng = ws.getCellRangeByName(job["range"])
        else:
            cur = ws.createCursor()
            cur.gotoStartOfUsedArea(False)
            cur.gotoEndOfUsedArea(True)
            rng = cur
        addr = rng.getRangeAddress()
        rows = addr.EndRow - addr.StartRow + 1
        cols = addr.EndColumn - addr.StartColumn + 1
        grid = [[_cell_value(rng.getCellByPosition(c, r)) for c in range(cols)] for r in range(rows)]
        text = rng.getCellByPosition(0, 0).getString() if rows == 1 and cols == 1 else None
        return {"sheet": ws.getName(), "address": _a1(addr), "rows": rows, "cols": cols, "values": grid, "text": text}
    finally:
        _close(doc)


def calc_pdf(office, job):
    doc = office.load(job["src"])
    try:
        if job.get("recalc", True):
            doc.calculateAll()
        sheet = job.get("sheet")
        if sheet not in (None, ""):
            keep = _pick_sheet(doc, sheet).getName()
            sheets = doc.getSheets()
            for name in list(sheets.getElementNames()):
                if name != keep:
                    sheets.removeByName(name)  # 저장하지 않는 사본 문서에서만 — 그 시트만 내보낸다
        return {"out_path": office.store(doc, job["out"], CALC_FILTERS)}
    finally:
        _close(doc)


# ---------------------------------------------------------------- Writer

def writer_update_fields(office, job):
    doc = office.load(job["src"])
    try:
        counts = _refresh_writer(doc)
        return {"out_path": office.store(doc, job["out"], WRITER_FILTERS), **counts}
    finally:
        _close(doc)


def _refresh_writer(doc):
    fields = doc.getTextFields()
    n_fields = 0
    enum = fields.createEnumeration()
    while enum.hasMoreElements():
        enum.nextElement()
        n_fields += 1
    fields.refresh()
    indexes = doc.getDocumentIndexes()
    for i in range(indexes.getCount()):
        indexes.getByIndex(i).update()
    try:
        doc.refresh()
    except Exception:
        pass
    return {"fields": n_fields, "indexes": indexes.getCount()}


def writer_layout(office, job):
    doc = office.load(job["path"], read_only=True)
    try:
        if job.get("update_fields", True):
            _refresh_writer(doc)
        ctrl = doc.getCurrentController()
        pages = ctrl.getPropertyValue("PageCount")
        stats = {}
        try:
            for nv in doc.getDocumentProperties().DocumentStatistics:
                stats[nv.Name] = nv.Value
        except Exception:
            pass
        styles = doc.getStyleFamilies().getByName("PageStyles")
        cursor = ctrl.getViewCursor()
        page_info = []
        try:
            cursor.jumpToFirstPage()
            for _ in range(pages):
                style = styles.getByName(cursor.PageStyleName)
                page_info.append({
                    "page": cursor.getPage(),
                    "width_mm": round(style.Width / 100.0, 1),
                    "height_mm": round(style.Height / 100.0, 1),
                    "landscape": bool(style.IsLandscape),
                })
                if not cursor.jumpToNextPage():
                    break
        except Exception:
            pass
        return {
            "pages": pages,
            "words": stats.get("WordCount"),
            "characters": stats.get("CharacterCount"),
            "paragraphs": stats.get("ParagraphCount"),
            "tables": doc.getTextTables().getCount(),
            "images": doc.getGraphicObjects().getCount(),
            "sections": doc.getTextSections().getCount(),
            "page_info": page_info,
        }
    finally:
        _close(doc)


def writer_store(office, job):
    """변환·PDF 공용: 열어서(필요하면 필드 갱신) 다른 형식으로 저장."""
    doc = office.load(job["src"], read_only=not job.get("update_fields"))
    try:
        if job.get("update_fields"):
            _refresh_writer(doc)
        return {"out_path": office.store(doc, job["out"], WRITER_FILTERS)}
    finally:
        _close(doc)


def writer_compare(office, job):
    """수정본을 열고 원본과 비교해 변경 추적(빨간 줄) 문서로 저장한다(Word 의 비교와 같은 결과물)."""
    doc = office.load(job["revised"])
    try:
        office.dispatch(doc, ".uno:CompareDocuments", URL=_url(job["original"]), NoAcceptDialog=True)
        redlines = doc.getPropertyValue("Redlines") if hasattr(doc, "getPropertyValue") else None
        n = 0
        if redlines is not None:
            enum = redlines.createEnumeration()
            while enum.hasMoreElements():
                enum.nextElement()
                n += 1
        return {"out_path": office.store(doc, job["out"], WRITER_FILTERS), "revisions": n}
    finally:
        _close(doc)


# ---------------------------------------------------------------- Impress

def impress_pdf(office, job):
    doc = office.load(job["src"], read_only=True)
    try:
        return {"out_path": office.store(doc, job["out"], IMPRESS_FILTERS), "slides": doc.getDrawPages().getCount()}
    finally:
        _close(doc)


IMAGE_TYPES = {"PNG": ("png", "image/png"), "JPG": ("jpg", "image/jpeg"), "JPEG": ("jpg", "image/jpeg"),
               "GIF": ("gif", "image/gif"), "BMP": ("bmp", "image/bmp"), "TIF": ("tif", "image/tiff")}


def impress_images(office, job):
    """슬라이드마다 이미지 한 장: out_dir/slide_001.<ext> … (PowerPoint 판 export_images 와 같은 이름·같은 반환).
    width·height 를 둘 다 주면 그 크기, 아니면 슬라이드 크기를 96dpi 로(PowerPoint 기본과 같은 기준)."""
    fmt = (job.get("fmt") or "PNG").upper()
    if fmt not in IMAGE_TYPES:
        raise ValueError("지원하지 않는 형식: %s (가능: %s)" % (fmt, ", ".join(IMAGE_TYPES)))
    ext, media = IMAGE_TYPES[fmt]
    doc = office.load(job["src"], read_only=True)
    try:
        pages = doc.getDrawPages()
        out_dir = os.path.abspath(job["out_dir"])
        os.makedirs(out_dir, exist_ok=True)
        width, height = job.get("width"), job.get("height")
        exporter = office.smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", office.ctx)
        files = []
        for i in range(pages.getCount()):
            page = pages.getByIndex(i)
            if width and height:
                pw, ph = int(width), int(height)
            else:  # 1/100 mm → 96dpi 픽셀
                pw, ph = int(round(page.Width / 2540.0 * 96)), int(round(page.Height / 2540.0 * 96))
            out = os.path.join(out_dir, "slide_%03d.%s" % (i + 1, ext))
            exporter.setSourceDocument(page)
            exporter.filter(_props(URL=_url(out), MediaType=media, FilterData=uno.Any(
                "[]com.sun.star.beans.PropertyValue", _props(PixelWidth=pw, PixelHeight=ph))))
            files.append(out)
        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            raise RuntimeError("이미지 일부 미생성: %s ..." % missing[:3])
        return {"out_dir": out_dir, "count": len(files), "files": files}
    finally:
        _close(doc)


OPS = {
    "calc_recalc": calc_recalc, "calc_read": calc_read, "calc_pdf": calc_pdf,
    "writer_update_fields": writer_update_fields, "writer_layout": writer_layout,
    "writer_store": writer_store, "writer_compare": writer_compare,
    "impress_pdf": impress_pdf, "impress_images": impress_images,
}


def run(*_args):
    result_path = os.environ.get("IRIS_LO_RESULT")
    office = None
    out = None
    try:
        with open(os.environ["IRIS_LO_JOB"], "r", encoding="utf-8") as f:
            job = json.load(f)
        office = Office()
        fn = OPS.get(job.get("op"))
        if fn is None:
            raise ValueError("모르는 작업: %r" % job.get("op"))
        out = {"ok": True, "result": fn(office, job)}
    except Exception as e:  # noqa: BLE001
        out = {"ok": False, "error": "%s: %s" % (type(e).__name__, e), "trace": traceback.format_exc()[-2000:]}
    finally:
        if result_path:
            try:
                with open(result_path, "w", encoding="utf-8") as f:
                    json.dump(out or {"ok": False, "error": "no result"}, f, ensure_ascii=False)
            except Exception:
                pass
        try:
            (office.desktop if office else uno.getComponentContext().ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", uno.getComponentContext())).terminate()
        except Exception:
            pass


g_exportedScripts = (run,)
