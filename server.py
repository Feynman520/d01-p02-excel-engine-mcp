"""
server.py — Excel 엔진 자동화 MCP 서버 (stdio).

Excel.Application 데스크톱을 COM으로 자동화해, 라이브러리(openpyxl 등)가 못 하는
**계산 엔진/렌더 엔진** 작업을 핵심 도구로 노출한다:
  ① excel_recalculate   수식 전체 재계산 → 실측값이 박힌 파일 저장
  ② excel_read_values   계산된 실제 셀 값 읽기
  ③ excel_export_pdf    Excel 렌더 엔진으로 PDF 출력

모든 Excel 호출은 단일 STA 워커 스레드(engine.session)에서 직렬화되며, 도구는 async로
그 블로킹 작업을 워커 스레드에 위임해 이벤트 루프를 막지 않는다. 세션은 **지연 초기화**:
첫 도구 호출 시 Excel을 띄우고, 서버 종료 시 닫는다.

실행:  python server.py        (Claude Code가 stdio로 기동)
"""
from __future__ import annotations

import threading
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any, Optional

import anyio
from mcp.server.fastmcp import FastMCP

import sys

if sys.platform == "win32":
    from engine import workbook
    from engine.session import ExcelSession
else:
    # macOS: Excel COM 이 없다 — 같은 도구를 LibreOffice(무료) 계산·렌더 엔진으로 처리한다(engine/lo_bridge.py).
    from engine import lo_workbook as workbook
    from engine.lo_bridge import LoSession as ExcelSession

# ---- 단일 세션 관리(지연 초기화) -------------------------------------------
_session: Optional[ExcelSession] = None
_lock = threading.Lock()


def _get_session() -> ExcelSession:
    global _session
    with _lock:
        if _session is None or not _session.alive:
            s = ExcelSession(visible=False)
            s.start()
            _session = s
        return _session


def _stop_session() -> None:
    global _session
    with _lock:
        if _session is not None:
            _session.stop()
            _session = None


@asynccontextmanager
async def _lifespan(_server: FastMCP) -> AsyncIterator[dict]:
    try:
        yield {}
    finally:
        await anyio.to_thread.run_sync(_stop_session)


mcp = FastMCP(
    "excel-automation",
    lifespan=_lifespan,
    instructions=(
        "Excel.Application COM 엔진 자동화. 라이브러리가 못 하는 계산·렌더 작업 전용. "
        "일반 흐름: (openpyxl 등으로 입력값을 쓴 파일을) excel_recalculate로 재계산 → "
        "excel_read_values로 실측값 확인 → excel_export_pdf로 PDF 출력. "
        "경로는 절대경로 권장. 원본은 보존되며 결과는 out_path로 저장된다."
        + ("" if sys.platform == "win32" else
           " (macOS: Excel 대신 LibreOffice Calc 엔진으로 계산·출력한다. Excel 전용 함수는 LibreOffice 가 지원하는 범위에서만 계산된다.)")
    ),
)


def _run(fn) -> Any:
    return _get_session().run(fn)


# ============================================================================
# 핵심 3개 도구 (+ 진단 2)
# ============================================================================
@mcp.tool()
async def excel_recalculate(src_path: str, out_path: str) -> dict:
    """① 수식 의존성 전체를 강제 재계산하고 **실측값이 박힌** 파일로 저장한다.

    openpyxl 등은 수식만 쓰고 결과를 못 채운다(읽으면 None). 이 도구는 진짜 Excel 계산
    엔진으로 NOW/RAND/XLOOKUP/동적배열/LAMBDA까지 계산해 최종 숫자를 확정한다.

    Args:
        src_path: 입력 통합문서 절대경로(.xlsx/.xlsm/.xlsb/.xls)
        out_path: 재계산 결과 저장 경로(확장자로 형식 결정)
    Returns: {"out_path"}
    """
    def work():
        return {"out_path": _run(lambda app: workbook.recalculate(app, src_path, out_path))}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def excel_read_values(
    path: str,
    sheet: Optional[str] = None,
    cell_range: Optional[str] = None,
    recalc: bool = True,
) -> dict:
    """② 셀의 **계산된 실제 값**을 2차원 배열로 읽는다(라이브러리가 못 읽는 값).

    Args:
        path: 통합문서 절대경로
        sheet: 시트 이름 또는 1-기반 인덱스(미지정=첫 시트)
        cell_range: A1 표기 범위(예: "B2:D10"). 미지정 시 UsedRange 전체.
        recalc: 읽기 전 전체 재계산 여부(기본 True)
    Returns: {"sheet","address","rows","cols","values"(2D), "text"(단일셀 표시텍스트)}
    """
    def work():
        return _run(lambda app: workbook.read_values(app, path, sheet, cell_range, recalc))
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def excel_export_pdf(
    src_path: str,
    out_path: str,
    sheet: Optional[str] = None,
    recalc: bool = True,
) -> dict:
    """③ 통합문서(또는 특정 시트)를 Excel 렌더 엔진으로 PDF 출력한다.

    인쇄영역·페이지 맞춤·머리글/바닥글이 반영된 WYSIWYG 출력. 라이브러리에는 렌더 엔진이 없다.

    Args:
        src_path: 입력 통합문서 절대경로
        out_path: 출력 .pdf 절대경로
        sheet: 특정 시트만 출력하려면 이름/인덱스(미지정=전체)
        recalc: 출력 전 재계산 여부(기본 True)
    Returns: {"out_path"}
    """
    def work():
        return {"out_path": _run(lambda app: workbook.export_pdf(app, src_path, out_path, sheet, recalc))}
    return await anyio.to_thread.run_sync(work)


# ---- 진단/복구 -------------------------------------------------------------
@mcp.tool()
async def excel_health() -> dict:
    """엔진 상태를 점검한다(세션 기동/Excel 버전). 첫 호출 시 Excel을 띄운다."""
    def work():
        s = _get_session()
        try:
            ver = s.run(lambda app: str(app.Version))
        except Exception as e:  # noqa: BLE001
            ver = f"버전 조회 실패: {e}"
        return {"alive": s.alive, "excel_version": ver}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def excel_restart() -> dict:
    """Excel 세션을 종료 후 재기동한다(COM 오류 복구용)."""
    def work():
        _stop_session()
        s = _get_session()
        return {"alive": s.alive}
    return await anyio.to_thread.run_sync(work)


if __name__ == "__main__":
    mcp.run(transport="stdio")
