"""
MCP 계층 검증 — Excel을 띄우지 않고(지연 초기화) server.py가 정상 임포트되고
5개 도구가 등록됐는지 확인한다.

실행: .venv\\Scripts\\python.exe tests\\server_tools.py
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import anyio  # noqa: E402
import server  # noqa: E402

EXPECTED = {
    "excel_recalculate",
    "excel_read_values",
    "excel_export_pdf",
    "excel_health",
    "excel_restart",
}


async def _list():
    tools = await server.mcp.list_tools()
    return {t.name for t in tools}


def main():
    names = anyio.run(_list)
    print(f"[srv] 등록된 도구({len(names)}): {sorted(names)}", flush=True)
    missing = EXPECTED - names
    extra = names - EXPECTED
    assert not missing, f"누락된 도구: {missing}"
    print(f"[srv] === MCP 계층 OK: 기대한 5개 도구 모두 등록됨 (추가:{extra or '없음'}) ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
