"""
engine/lo_workbook.py — macOS 엔진의 ① 재계산 ② 실측값 추출 ③ PDF 렌더링(engine/workbook.py 와 같은 함수·같은 반환).

첫 인자 `app` 은 COM 의 Excel.Application 대신 engine/lo_bridge.LoBridge 다. 계산은 LibreOffice Calc 엔진이
한다 — Excel 고유 함수(예: LAMBDA·일부 동적 배열)는 LibreOffice 가 지원하는 범위에서만 계산된다.
"""
from __future__ import annotations

import os
from typing import Optional


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _require(path: str) -> str:
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    return src


def recalculate(app, src_path: str, out_path: str) -> str:
    """수식 전체를 다시 계산해 **결과값이 박힌** 파일로 저장한다(확장자로 형식 결정)."""
    r = app.run("calc_recalc", src=_require(src_path), out=_abspath(out_path))
    return r["out_path"]


def read_values(app, path: str, sheet: Optional[str] = None, cell_range: Optional[str] = None, recalc: bool = True) -> dict:
    """범위(미지정이면 사용 영역 전체)의 계산된 값을 2차원 배열로. 단일 셀이면 표시 텍스트도."""
    return app.run("calc_read", path=_require(path), sheet=sheet, range=cell_range or None, recalc=bool(recalc))


def export_pdf(app, src_path: str, out_path: str, sheet: Optional[str] = None, recalc: bool = True) -> str:
    """통합문서(또는 한 시트)를 PDF 로. 인쇄 영역·페이지 설정은 LibreOffice 가 읽는 만큼 반영된다."""
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    r = app.run("calc_pdf", src=_require(src_path), out=_abspath(out_path), sheet=sheet, recalc=bool(recalc))
    return r["out_path"]
