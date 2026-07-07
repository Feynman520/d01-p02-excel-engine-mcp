"""
engine/workbook.py — Excel 엔진 전용 작업: ① 재계산  ② 실측값 추출  ③ PDF 렌더링.

모든 함수는 `app`(Excel.Application COM)을 첫 인자로 받고 전용 워커 스레드에서 호출된다
(session.run). 각 함수는 open → 작업 → Close(저장 안 함)로 상태를 정리해 다음 호출 오염을
막는다(원본 보존). 라이브러리(openpyxl 등)가 못 하는 **계산 엔진/렌더 엔진** 작업만 노출한다.
"""
from __future__ import annotations

import os
from typing import Any, Optional

# --- 후기바인딩용 숫자 enum 상수 ------------------------------------------------
XL_TYPE_PDF = 0          # xlTypePDF
XL_CALC_AUTOMATIC = -4105  # xlCalculationAutomatic
# 확장자 → FileFormat(SaveAs)
_FMT_BY_EXT = {
    ".xlsx": 51,   # xlOpenXMLWorkbook
    ".xlsm": 52,   # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,   # xlExcel12
    ".xls": 56,    # xlExcel8
    ".csv": 6,     # xlCSV
}


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _ensure_parent(path: str) -> None:
    os.makedirs(os.path.dirname(_abspath(path)), exist_ok=True)


def _coerce(v: Any) -> Any:
    """COM 셀 값을 JSON 직렬화 가능한 값으로 변환한다."""
    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    # pywintypes 날짜/시간 → ISO 문자열
    iso = getattr(v, "isoformat", None)
    if callable(iso):
        try:
            return iso()
        except Exception:
            pass
    return str(v)


def _to_grid(values: Any) -> list:
    """Range.Value2 결과(스칼라 | 튜플의 튜플)를 2차원 리스트로 정규화한다."""
    if values is None:
        return [[None]]
    if not isinstance(values, tuple):  # 단일 셀
        return [[_coerce(values)]]
    grid = []
    for row in values:
        if isinstance(row, tuple):
            grid.append([_coerce(c) for c in row])
        else:  # 단일 행/열
            grid.append([_coerce(row)])
    return grid


def _open(app, path: str, read_only: bool):
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    # UpdateLinks=0 → 외부링크 갱신 프롬프트 차단
    return app.Workbooks.Open(src, 0, read_only)


def _pick_sheet(wb, sheet: Optional[str]):
    if sheet is None or sheet == "":
        return wb.Worksheets(1)
    s = sheet.strip()
    if s.isdigit():
        return wb.Worksheets(int(s))
    return wb.Worksheets(s)


# ============================================================================
# ① 재계산 — 의존성 그래프 전체 재계산 후 실측값이 박힌 파일로 저장
# ============================================================================
def recalculate(app, src_path: str, out_path: str) -> str:
    """수식 의존성 전체를 강제 재계산(CalculateFullRebuild)하고 결과를 저장한다.

    openpyxl 등 라이브러리는 수식 문자열만 저장하고 값을 계산하지 못한다(읽으면 None).
    이 함수는 진짜 Excel 계산 엔진을 돌려 **최종 숫자가 박힌** 파일을 만든다.
    """
    _ensure_parent(out_path)
    wb = _open(app, src_path, read_only=False)
    try:
        try:
            app.Calculation = XL_CALC_AUTOMATIC
        except Exception:
            pass
        app.CalculateFullRebuild()
        out = _abspath(out_path)
        ext = os.path.splitext(out)[1].lower()
        fmt = _FMT_BY_EXT.get(ext, 51)
        wb.SaveAs(out, fmt)
    finally:
        wb.Close(SaveChanges=False)
    if not os.path.exists(_abspath(out_path)):
        raise RuntimeError(f"저장 실패: {out_path}")
    return _abspath(out_path)


# ============================================================================
# ② 실측값 추출 — 재계산된 셀의 실제 값/표시텍스트를 읽어 반환
# ============================================================================
def read_values(
    app,
    path: str,
    sheet: Optional[str] = None,
    cell_range: Optional[str] = None,
    recalc: bool = True,
) -> dict:
    """범위의 **계산된 실제 값**을 읽어 2차원 배열로 반환한다(라이브러리가 못 읽는 값).

    cell_range 미지정 시 시트의 UsedRange 전체를 읽는다.
    """
    wb = _open(app, path, read_only=True)
    try:
        if recalc:
            try:
                app.Calculation = XL_CALC_AUTOMATIC
            except Exception:
                pass
            app.CalculateFullRebuild()
        ws = _pick_sheet(wb, sheet)
        rng = ws.UsedRange if not cell_range else ws.Range(cell_range)
        addr = str(rng.Address)  # 후기바인딩: Address는 인자 없는 속성($B$1:$B$2 형태)
        grid = _to_grid(rng.Value2)
        # 단일 셀이면 표시 텍스트(NumberFormat 적용)도 함께 제공
        text = None
        if len(grid) == 1 and len(grid[0]) == 1:
            try:
                text = rng.Text
            except Exception:
                text = None
        return {
            "sheet": ws.Name,
            "address": addr,
            "rows": len(grid),
            "cols": len(grid[0]) if grid else 0,
            "values": grid,
            "text": text,
        }
    finally:
        wb.Close(SaveChanges=False)


# ============================================================================
# ③ PDF 렌더링 — Excel 렌더 엔진으로 정밀 PDF 출력
# ============================================================================
def export_pdf(
    app,
    src_path: str,
    out_path: str,
    sheet: Optional[str] = None,
    recalc: bool = True,
) -> str:
    """통합문서(또는 특정 시트)를 Excel 렌더 엔진으로 PDF 출력한다(인쇄영역·페이지맞춤 반영).

    sheet 지정 시 해당 시트만, 미지정 시 인쇄영역이 설정된 전체를 출력한다.
    """
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    _ensure_parent(out_path)
    wb = _open(app, src_path, read_only=False)
    try:
        if recalc:
            try:
                app.Calculation = XL_CALC_AUTOMATIC
            except Exception:
                pass
            app.CalculateFullRebuild()
        target = wb if (sheet is None or sheet == "") else _pick_sheet(wb, sheet)
        target.ExportAsFixedFormat(XL_TYPE_PDF, _abspath(out_path))
    finally:
        wb.Close(SaveChanges=False)
    if not os.path.exists(_abspath(out_path)):
        raise RuntimeError(f"PDF 저장 실패: {out_path}")
    return _abspath(out_path)
