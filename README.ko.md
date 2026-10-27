# excel-engine-mcp — 한국어 안내

[English → README.md](README.md)

Excel.Application 데스크톱을 COM으로 자동화해, **라이브러리(openpyxl 등)가 못 하는 계산
엔진·렌더 엔진 작업**을 핵심 도구로 노출하는 로컬 MCP 서버.

> 설계 사상: 정적 셀 쓰기·대량 데이터는 라이브러리가 빠르고, **수식 재계산·실측값·PDF 렌더**는
> 엔진만 가능하다. 이 MCP는 후자(엔진 전용)만 담당한다 — 그래서 도구가 5개뿐이다.

## 요구사항
- Windows 10+ / **로그인된 인터랙티브 데스크톱 세션** (Excel은 진짜 headless가 없음)
- Microsoft Office(Excel) 설치 — 검증 환경: **Office 2016+ (Excel 16.0)**
- Python 3.10+ — 검증: **3.12**
- [Claude Code](https://claude.com/claude-code) 또는 임의의 MCP 클라이언트

## macOS (LibreOffice 엔진)

macOS 에는 COM 이 없어 같은 도구 5개를 **LibreOffice**(무료, MPL-2.0)를 화면 없이 띄워 처리한다 —
`server.py` 가 COM 세션 대신 `engine/lo_workbook.py` + `engine/lo_bridge.py` 를 고른다.

- LibreOffice 설치(`brew install --cask libreoffice` 또는 libreoffice.org). 다른 `soffice` 는 `IRIS_SOFFICE` 로 지정.
- 호출마다 전용 프로필(`.lo-profile/` 또는 `IRIS_LO_PROFILE_DIR`)로 `soffice` 를 한 번 띄우고 그 안에서 UNO 매크로
  (`engine/lo_macro/iris_uno.py`)를 돌린다 — LibreOffice 26 의 내장 Python 은 `soffice` 만 띄울 수 있다.
- 계산은 LibreOffice Calc 엔진이 한다: Excel 전용 함수는 LibreOffice 가 지원하는 범위에서만 계산된다.
- 검사: `python tests/lo_engine.py`(저장값이 틀린 xlsx 를 만들어 재계산이 실제로 돌았음을 증명).

## 설치

```powershell
git clone https://github.com/Feynman520/d01-p02-excel-engine-mcp.git
cd d01-p02-excel-engine-mcp
py -3.12 -m venv .venv          # 또는: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Claude Code 등록

클론한 폴더 안에서 아래 한 줄 실행 (절대경로가 박히므로 이후 어디서든 동작):

```powershell
claude mcp add excel-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user`는 모든 프로젝트에서 사용 가능. 한 프로젝트에서만 쓰려면 `--scope project`.

### 동작 검증

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # COM 재계산·실측값·PDF·좀비정리 검증
& $py tests\server_tools.py   # MCP 5개 도구 등록 검증(Excel 미기동)
```

## 도구 (핵심 3 + 진단 2)

| # | 도구 | 입력 → 출력 | 엔진 전용 이유 |
|---|---|---|---|
| ① | `excel_recalculate` | `src_path, out_path` → `{out_path}` | 수식 의존성 전체 재계산(라이브러리는 값 못 채움) |
| ② | `excel_read_values` | `path, sheet?, cell_range?, recalc?` → `{sheet,address,rows,cols,values,text}` | 계산된 실제 셀 값(openpyxl은 수식만 보유) |
| ③ | `excel_export_pdf` | `src_path, out_path, sheet?, recalc?` → `{out_path}` | Excel 렌더 엔진 PDF(인쇄영역·페이지맞춤·머리글/바닥글 반영) |
| — | `excel_health` | → `{alive, excel_version}` | 세션 점검(첫 호출 시 Excel 기동) |
| — | `excel_restart` | → `{alive}` | COM 오류 복구 |

전형적 흐름: openpyxl 등으로 입력값을 쓴 `.xlsx` → `excel_recalculate`로 실측값 확정 →
`excel_read_values`로 결과 확인 / `excel_export_pdf`로 제출용 PDF.
원본은 절대 수정하지 않고 결과는 항상 `out_path`로 저장한다.

## 아키텍처 핵심
- **단일 STA 워커 스레드**(`engine/session.py`): win32com COM은 STA에 묶이고 FastMCP는 async
  스레드를 옮겨다니므로, 모든 Excel 호출을 전용 스레드 1개에 직렬화한다.
- **세션 지연 초기화**: 첫 도구 호출 시 Excel 기동, 서버 종료 시 닫음. 1개 인스턴스 재사용.
- **`DispatchEx` + `gencache.EnsureDispatch`(조기바인딩)**: 전용 인스턴스를 새로 띄우고(사용자
  Excel·좀비에 안 붙음), 조기바인딩으로 호출한다.
  ⚠️ `ExportAsFixedFormat`은 **후기바인딩에서 1004(0x800A03EC)로 실패**한다(선택적 인자
  마샬링 오류). 조기바인딩이 필수.
- **좀비 방지**: 종료 시 `Quit()` 후에도 남는 전용 Excel 프로세스를 PID(메인창 Hwnd로 식별)로
  강제 종료한다.
- **RPC 거부 재시도**: 기동 직후 `RPC_E_CALL_REJECTED`를 잠깐 쉬고 재시도.
- **stdout 보호**: 첫 makepy 캐시 생성 출력을 fd 우회로 stderr로 보냄(stdout=JSON-RPC 채널).

## 한계
- 무인/서비스 세션 부적합(데스크톱 세션 필요).
- 대량 정적 쓰기는 라이브러리가 빠름 — 의도된 역할 분담이다.

## 라이선스

[MIT](LICENSE)
