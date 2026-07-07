"""
스모크 테스트 — 본 머신에서 Excel COM 엔진 자동화가 실제로 동작하는지 검증한다.
MCP 계층 없이 engine(session+workbook)을 직접 호출한다.

실행:  .venv\\Scripts\\python.exe tests\\smoke_com.py

통과 기준: 예외/팝업 없이 끝까지 진행하고 recalc.xlsx / out.pdf 가 생성되며
재계산 값(SUM=42, 곱=200)이 정확히 읽힌다.
"""
import os
import sys
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine import workbook  # noqa: E402
from engine.session import ExcelSession  # noqa: E402

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out")
os.makedirs(OUT_DIR, exist_ok=True)


def log(msg):
    print(f"[smoke] {msg}", flush=True)


def _make_input(app, path):
    """수식이 든 입력 통합문서를 COM으로 생성한다(openpyxl이 만든 입력 파일 대역)."""
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Range("A1").Value = 10
    ws.Range("A2").Value = 20
    ws.Range("A3").Value = 12
    ws.Range("B1").Formula = "=SUM(A1:A3)"   # 42
    ws.Range("B2").Formula = "=A1*A2"        # 200
    if os.path.exists(path):
        os.remove(path)
    wb.SaveAs(path, 51)  # xlOpenXMLWorkbook
    wb.Close(SaveChanges=False)
    return path


def main():
    inp = os.path.join(OUT_DIR, "input.xlsx")
    recalc = os.path.join(OUT_DIR, "recalc.xlsx")
    pdf = os.path.join(OUT_DIR, "out.pdf")

    log("ExcelSession 시작 ...")
    s = ExcelSession(visible=False)
    s.start()
    try:
        log(f"Excel Version = {s.run(lambda app: str(app.Version))}")

        log("입력 통합문서 생성 ...")
        s.run(lambda app: _make_input(app, inp))
        assert os.path.exists(inp), "입력 파일 미생성"

        log("① recalculate ...")
        s.run(lambda app: workbook.recalculate(app, inp, recalc))
        assert os.path.exists(recalc), "재계산 파일 미생성"

        log("② read_values B1:B2 ...")
        res = s.run(lambda app: workbook.read_values(app, recalc, None, "B1:B2", True))
        log(f"   읽은 값 = {res['values']}")
        flat = [res["values"][0][0], res["values"][1][0]]
        assert flat == [42, 200], f"재계산 값 불일치: {flat}"

        log("③ export_pdf ...")
        s.run(lambda app: workbook.export_pdf(app, recalc, pdf, None, True))
        assert os.path.exists(pdf), "PDF 미생성"

        log("=== SMOKE PASS: Excel COM 재계산 + 실측값 + PDF OK ===")
        return 0
    finally:
        log("ExcelSession 종료 ...")
        s.stop()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
