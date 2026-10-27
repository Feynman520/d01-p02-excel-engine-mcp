"""
macOS 엔진 검증(LibreOffice) — engine/lo_bridge.py + engine/lo_workbook.py 를 진짜 soffice 로 돌린다.

수식의 저장값(캐시)이 틀린 xlsx 를 표준 라이브러리로 직접 만들어(`=A1+A2` 인데 캐시 999),
재계산을 켜면 3, 끄면 999 가 읽히는지로 "계산 엔진이 실제로 돌았다"를 증명한다.

실행: python tests/lo_engine.py      (LibreOffice 가 없으면 SKIP)
"""
import os
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine.lo_bridge import LoSession, find_soffice  # noqa: E402
from engine import lo_workbook as wb  # noqa: E402

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="계산" sheetId="1" r:id="rId1"/><sheet name="둘째" sheetId="2" r:id="rId2"/></sheets>
</workbook>"""
WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"""
SHEET1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>한글</t></is></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><f>A1+A2</f><v>999</v></c></row>
</sheetData></worksheet>"""
SHEET2 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>second</t></is></c></row>
</sheetData></worksheet>"""


def make_stale_xlsx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("xl/workbook.xml", WB)
        z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
        z.writestr("xl/worksheets/sheet1.xml", SHEET1)
        z.writestr("xl/worksheets/sheet2.xml", SHEET2)


def main():
    if not find_soffice():
        print("SKIP: LibreOffice 없음")
        return 0
    tmp = tempfile.mkdtemp(prefix="iris-lo-test-")
    fails = 0

    def check(cond, name):
        nonlocal fails
        print(("PASS" if cond else "FAIL"), name, flush=True)
        if not cond:
            fails += 1

    try:
        src = os.path.join(tmp, "stale.xlsx")
        make_stale_xlsx(src)
        s = LoSession(profile_dir=os.path.join(tmp, "profile"))
        s.start()
        stale = s.run(lambda a: wb.read_values(a, src, cell_range="A3", recalc=False))
        check(stale["values"] == [[999.0]], f"재계산 끔 → 저장값 그대로 999 ({stale['values']})")
        fresh = s.run(lambda a: wb.read_values(a, src, cell_range="A3", recalc=True))
        check(fresh["values"] == [[3.0]] and fresh["text"] == "3", f"재계산 켬 → 3 ({fresh['values']}, text={fresh['text']!r})")
        whole = s.run(lambda a: wb.read_values(a, src))
        check(whole["sheet"] == "계산" and whole["address"] == "$A$1:$B$3" and whole["values"][0] == [1.0, "한글"] and whole["values"][1][1] is None,
              f"범위 생략 → 사용 영역 전체·한글·빈 칸 None ({whole['address']}, {whole['values']})")
        second = s.run(lambda a: wb.read_values(a, src, sheet="2"))
        check(second["sheet"] == "둘째" and second["values"] == [["second"]], "시트 번호(1-기반)")
        out = os.path.join(tmp, "out", "recalc.xlsx")
        s.run(lambda a: wb.recalculate(a, src, out))
        after = s.run(lambda a: wb.read_values(a, out, cell_range="A3", recalc=False))
        check(after["values"] == [[3.0]], "excel_recalculate 결과 파일에는 새 값이 박혀 있다(다시 읽을 때 재계산 없이도 3)")
        pdf = os.path.join(tmp, "out", "one.pdf")
        s.run(lambda a: wb.export_pdf(a, src, pdf, sheet="둘째"))
        with open(pdf, "rb") as f:
            head = f.read(5)
        check(head == b"%PDF-", "export_pdf(시트 하나) → PDF")
        try:
            s.run(lambda a: wb.export_pdf(a, src, os.path.join(tmp, "x.png")))
            check(False, "PDF 아닌 확장자는 거부")
        except ValueError:
            check(True, "PDF 아닌 확장자는 거부")
        try:
            s.run(lambda a: wb.read_values(a, os.path.join(tmp, "없음.xlsx")))
            check(False, "없는 파일은 FileNotFoundError")
        except FileNotFoundError:
            check(True, "없는 파일은 FileNotFoundError")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"---- FAILS: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
