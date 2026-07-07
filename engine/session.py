"""
engine/session.py — Excel.Application COM 세션을 전용 STA 워커 스레드 1개에 고정한다.

win32com COM 객체는 STA(단일 스레드 아파트)에 묶여 생성한 그 스레드에서만 안전하게
호출된다. FastMCP는 asyncio 이벤트 루프에서 도구를 실행하며 스레드를 옮겨다닐 수 있으므로,
모든 Excel 호출을 여기 정의한 전용 스레드의 작업 큐로 직렬화한다.

후기바인딩(`win32com.client.Dispatch`)을 쓴다 — makepy/gen_py 캐시를 생성하지 않으므로
import-time stdout 오염(MCP의 JSON-RPC 채널)이 없고, 필요한 enum은 숫자 상수로 직접 쓴다.

사용:
    session = ExcelSession()
    session.start()                          # Excel 띄움 (블로킹)
    val = session.run(lambda app: app.Version)
    session.stop()                           # Excel 종료
"""
from __future__ import annotations

import contextlib
import os
import queue
import sys
import threading
import time
import traceback
from concurrent.futures import Future
from typing import Any, Callable, Optional


@contextlib.contextmanager
def _fd1_to_stderr():
    """파일디스크립터 1(stdout)을 잠시 stderr로 우회한다.

    win32com `gencache.EnsureDispatch`는 첫 호출 시 Excel 타입라이브러리의 makepy 캐시를
    생성하며 stdout으로 진행상황을 print할 수 있는데, MCP stdio 서버에서 stdout은 JSON-RPC
    채널(파이프)이다. 그 1회성 생성 구간 동안만 fd1을 stderr로 돌려 프로토콜 오염을 막는다.
    """
    try:
        sys.stdout.flush()
    except Exception:
        pass
    saved = os.dup(1)
    try:
        os.dup2(2, 1)
        yield
    finally:
        try:
            sys.stdout.flush()
        except Exception:
            pass
        os.dup2(saved, 1)
        os.close(saved)

# COM "서버 사용 중" HRESULT — 시작 직후 Excel이 초기화 중일 때 흔히 발생, 재시도로 해소.
_RPC_RETRY_HRESULTS = (
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED ("피호출자가 호출을 거부했습니다")
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
)


def _com_retry(fn: Callable[[], Any], *, tries: int = 24, delay: float = 0.25) -> Any:
    """COM 호출이 '서버 사용 중'으로 거부되면 잠깐 쉬고 재시도한다(그 외 오류는 즉시 전파)."""
    import pythoncom  # noqa: F401  (com_error 식별용)

    last: Optional[BaseException] = None
    for _ in range(tries):
        try:
            return fn()
        except pythoncom.com_error as e:  # type: ignore[attr-defined]
            hr = e.args[0] if e.args else None
            if hr in _RPC_RETRY_HRESULTS:
                last = e
                time.sleep(delay)
                continue
            raise
    if last is not None:
        raise last


def _app_pid(app) -> Optional[int]:
    """Excel.Application 메인 창 핸들로 그 프로세스 PID를 얻는다(전용 인스턴스 식별용)."""
    try:
        import win32process

        hwnd = int(app.Hwnd)
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid)
    except Exception as e:  # noqa: BLE001
        _dbg(f"PID 조회 실패(무시): {e}")
    return None


def _force_kill(pid: Optional[int]) -> None:
    """Quit 후에도 살아있는 전용 Excel 프로세스를 PID로 강제 종료한다(좀비 방지 안전망)."""
    if not pid:
        return
    try:
        import win32api
        import win32con

        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 0)
        win32api.CloseHandle(h)
        _dbg(f"잔존 Excel PID {pid} 강제종료")
    except Exception:  # 이미 종료됐거나 접근 불가 → 무시
        pass


def _dbg(msg: str) -> None:
    """진단 로그를 stderr와(설정 시) 파일에 쓴다. stdout은 MCP JSON-RPC 채널이라 절대 쓰지 않는다."""
    line = f"[excel-session {time.strftime('%H:%M:%S')}] {msg}"
    try:
        print(line, file=sys.stderr, flush=True)
    except Exception:
        pass
    path = os.environ.get("EXCEL_MCP_DEBUG")
    if path:
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


class ExcelSession:
    """전용 워커 스레드가 소유하는 단일 Excel.Application 인스턴스."""

    def __init__(self, *, visible: bool = False):
        self._visible = visible
        self._tasks: "queue.Queue[Optional[tuple[Callable, Future]]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._init_error: Optional[BaseException] = None
        self._app: Any = None
        self._pid: Optional[int] = None

    # ---- lifecycle ---------------------------------------------------------
    def start(self, timeout: float = 90.0) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(
            target=self._run, name="excel-com-worker", daemon=True
        )
        self._thread.start()
        if not self._ready.wait(timeout):
            raise TimeoutError("Excel COM 세션 초기화가 시간 초과되었습니다.")
        if self._init_error is not None:
            raise self._init_error

    def stop(self, timeout: float = 30.0) -> None:
        if self._thread is None:
            return
        self._tasks.put(None)  # sentinel → 워커 종료
        self._thread.join(timeout)
        self._thread = None

    @property
    def alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    # ---- 작업 제출 ---------------------------------------------------------
    def run(self, fn: Callable[[Any], Any], timeout: Optional[float] = 300.0) -> Any:
        """`fn(app)`을 워커 스레드에서 실행하고 결과를 반환(블로킹)."""
        if not self.alive:
            raise RuntimeError("Excel 세션이 살아있지 않습니다. start()를 먼저 호출하세요.")
        fut: Future = Future()
        self._tasks.put((fn, fut))
        return fut.result(timeout)

    # ---- 워커 본체 ---------------------------------------------------------
    def _run(self) -> None:
        import pythoncom

        _dbg("worker thread 시작")
        try:
            pythoncom.CoInitialize()
            _dbg("CoInitialize 완료")
        except Exception as e:  # 이미 초기화된 경우 등은 무시
            _dbg(f"CoInitialize 예외(무시): {e}")

        try:
            import win32com.client

            _dbg("Excel.Application DispatchEx 시작 ...")
            t0 = time.time()
            # DispatchEx → 사용자가 띄운 Excel/좀비에 붙지 않고 전용 인스턴스를 새로 띄운다.
            raw = _com_retry(lambda: win32com.client.DispatchEx("Excel.Application"))
            # 조기바인딩으로 래핑 — 후기바인딩은 ExportAsFixedFormat 등 선택적 인자 마샬링을
            # 잘못해 Excel이 1004(0x800A03EC)를 던진다. EnsureDispatch는 타입라이브러리로
            # 정확히 호출한다. 첫 호출의 makepy 생성 출력은 fd 가드로 stderr로 보낸다.
            with _fd1_to_stderr():
                app = win32com.client.gencache.EnsureDispatch(raw)
            # 워밍업: Excel이 호출을 받을 준비가 될 때까지 사소한 속성을 폴링.
            _com_retry(lambda: app.Version)
            # 무인 자동화용 환경설정 — 모달 팝업/링크갱신/이벤트 차단
            for prop, val in (
                ("Visible", self._visible),
                ("DisplayAlerts", False),
                ("ScreenUpdating", False),
                ("EnableEvents", False),
                ("AskToUpdateLinks", False),
                ("AlertBeforeOverwriting", False),
            ):
                try:
                    _com_retry(lambda p=prop, v=val: setattr(app, p, v))
                except Exception as e:  # noqa: BLE001
                    _dbg(f"{prop} 설정 실패(무시): {e}")
            self._app = app
            self._pid = _app_pid(app)
            _dbg(f"Excel 생성 완료 ({time.time() - t0:.1f}s, pid={self._pid})")
        except BaseException as e:  # noqa: BLE001 — 초기화 실패를 메인 스레드로 전달
            _dbg(f"Excel 생성 실패: {e!r}\n{traceback.format_exc()}")
            self._init_error = e
            self._ready.set()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return

        self._ready.set()

        try:
            while True:
                item = self._tasks.get()
                if item is None:
                    break
                fn, fut = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn(self._app))
                except BaseException as e:  # noqa: BLE001
                    fut.set_exception(e)
        finally:
            try:
                if self._app is not None:
                    # 남은 통합문서를 저장 없이 닫고 Excel 종료
                    try:
                        for wb in list(self._app.Workbooks):
                            try:
                                wb.Close(SaveChanges=False)
                            except Exception:
                                pass
                    except Exception:
                        pass
                    try:
                        _com_retry(lambda: self._app.Quit(), tries=12)
                    except Exception as e:
                        _dbg(f"Quit 실패(무시): {e}")
            finally:
                # COM 참조를 모두 끊고 gc → Quit이 프로세스를 끝낼 수 있게 한다.
                self._app = None
                app = None  # type: ignore[assignment]  # noqa: F841
                raw = None  # type: ignore[assignment]  # noqa: F841
                import gc

                gc.collect()
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                time.sleep(0.4)  # Quit의 정상 종료에 잠깐 여유
                _force_kill(self._pid)  # 그래도 살아있으면 PID로 강제 종료(전용 인스턴스)
                self._pid = None
            _dbg("worker thread 종료")
