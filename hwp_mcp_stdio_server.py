#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import logging
import ssl
import tempfile
import atexit
import ctypes
import queue
import socket
import subprocess
import threading
import time
import uuid

# Configure logging
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hwp_mcp_stdio_server.log")

logging.basicConfig(
    level=logging.INFO,
    filename=log_path,
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger("hwp-mcp-stdio-server")

try:
    sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')
except Exception:
    pass

if not any(isinstance(handler, logging.StreamHandler) for handler in logger.handlers):
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
    logger.addHandler(stderr_handler)

# Optional: Disable SSL certificate validation for development
ssl._create_default_https_context = ssl._create_unverified_context

# Set up paths
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    # Import FastMCP library
    from mcp.server.fastmcp import FastMCP
    logger.info("FastMCP successfully imported")
except ImportError as e:
    logger.error(f"Failed to import FastMCP: {str(e)}")
    print(f"Error: Failed to import FastMCP. Please install with 'pip install mcp'", file=sys.stderr)
    sys.exit(1)

# Try to import HwpController
try:
    from src.tools.hwp_controller import HwpController
    logger.info("HwpController imported successfully")
except ImportError as e:
    logger.error(f"Failed to import HwpController: {str(e)}")
    # Try alternate paths
    try:
        sys.path.append(os.path.join(current_dir, "src"))
        sys.path.append(os.path.join(current_dir, "src", "tools"))
        from hwp_controller import HwpController
        logger.info("HwpController imported from alternate path")
    except ImportError as e2:
        logger.error(f"Could not find HwpController in any path: {str(e2)}")
        print(f"Error: Could not find HwpController module", file=sys.stderr)
        sys.exit(1)

# Try to import HwpTableTools
try:
    from src.tools.hwp_table_tools import HwpTableTools
    logger.info("HwpTableTools imported successfully")
except ImportError as e:
    logger.error(f"Failed to import HwpTableTools: {str(e)}")
    # Try alternate paths
    try:
        from hwp_table_tools import HwpTableTools
        logger.info("HwpTableTools imported from alternate path")
    except ImportError as e2:
        logger.error(f"Could not find HwpTableTools in any path: {str(e2)}")
        print(f"Error: Could not find HwpTableTools module", file=sys.stderr)
        sys.exit(1)

# Initialize FastMCP server
mcp = FastMCP(
    "hwp-mcp",
    instructions="HWP MCP Server for controlling Hangul Word Processor",
    dependencies=["pywin32>=305"]
)

_SIMPLE_RETURN_TYPES = (str, int, float, bool, bytes, type(None))
ROOT_ENSURE_TIMEOUT_SECONDS = 75.0
BROKER_CONNECT_TIMEOUT_SECONDS = 5.0


def _broker_state_path() -> str:
    temp_root = os.path.join(tempfile.gettempdir(), "hwp-mcp")
    os.makedirs(temp_root, exist_ok=True)
    return os.path.join(temp_root, "broker_state.json")


def _is_marshaled_value(value) -> bool:
    if isinstance(value, _SIMPLE_RETURN_TYPES):
        return True
    if isinstance(value, tuple):
        return all(_is_marshaled_value(item) for item in value)
    if isinstance(value, list):
        return all(_is_marshaled_value(item) for item in value)
    if isinstance(value, dict):
        return all(
            isinstance(key, _SIMPLE_RETURN_TYPES) and _is_marshaled_value(item)
            for key, item in value.items()
        )
    return False


class HwpComWorker:
    """Persistent helper-process client for all HWP automation."""

    def __init__(self):
        self._lock = threading.Lock()
        self._pending_lock = threading.Lock()
        self._pending = {}
        self._process = None
        self._socket = None
        self._socket_reader = None
        self._reader_thread = None
        self._last_hwp_error = None
        self._helper_script = os.path.join(current_dir, "hwp_mcp_helper.py")
        self._startup_cooldown_until = 0.0
        self._startup_cooldown_message = None

    def _set_last_hwp_error(self, message) -> None:
        self._last_hwp_error = message

    def _set_startup_cooldown(self, message: str, seconds: float = 90.0) -> None:
        self._startup_cooldown_until = time.time() + seconds
        self._startup_cooldown_message = message
        self._set_last_hwp_error(message)

    def _clear_startup_cooldown(self) -> None:
        self._startup_cooldown_until = 0.0
        self._startup_cooldown_message = None

    def _in_startup_cooldown(self) -> bool:
        return time.time() < self._startup_cooldown_until

    def last_error_message(self) -> str:
        return self._last_hwp_error or "Failed to connect to HWP program"

    def _cleanup_hwp_processes(self) -> None:
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    "$targets = Get-Process Hwp,HwpApi -ErrorAction SilentlyContinue; "
                    "if ($targets) { $targets | Stop-Process -Force }",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=15,
                check=False,
            )
            if result.returncode != 0 and (result.stderr or "").strip():
                logger.warning(f"HWP cleanup stderr: {result.stderr.strip()}")
        except Exception as exc:
            logger.warning(f"Failed to cleanup HWP processes: {exc}")

    @staticmethod
    def _is_root_ensure_request(command: str, payload: dict) -> bool:
        return command == "ensure_root" and payload.get("root") in {"controller", "table_tools"}

    @staticmethod
    def _running_in_job() -> bool:
        try:
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(0x1000, False, os.getpid())
            if not handle:
                return False
            try:
                in_job = ctypes.c_int()
                if not kernel32.IsProcessInJob(handle, 0, ctypes.byref(in_job)):
                    return False
                return bool(in_job.value)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return False

    def _broker_unavailable_message(self) -> str:
        return (
            "외부 HWP broker가 실행 중이 아닙니다. "
            "start_hwp_broker.ps1로 broker를 먼저 실행해 주세요."
        )

    def _connect_to_external_broker(self) -> bool:
        state_path = _broker_state_path()
        if not os.path.exists(state_path):
            self._set_last_hwp_error(self._broker_unavailable_message())
            return False

        try:
            with open(state_path, "r", encoding="utf-8") as fp:
                state = json.load(fp)
            port = int(state["port"])
        except Exception as exc:
            self._set_last_hwp_error(f"외부 HWP broker 상태 파일을 읽지 못했습니다: {exc}")
            return False

        try:
            sock = socket.create_connection(
                ("127.0.0.1", port),
                timeout=BROKER_CONNECT_TIMEOUT_SECONDS,
            )
            sock.settimeout(None)
            reader = sock.makefile("r", encoding="utf-8", errors="replace")
        except Exception as exc:
            self._set_last_hwp_error(f"외부 HWP broker 연결 실패: {exc}")
            return False

        self._socket = sock
        self._socket_reader = reader
        self._reader_thread = threading.Thread(
            target=self._reader_loop,
            args=(reader,),
            name="hwp-broker-reader",
            daemon=True,
        )
        self._reader_thread.start()
        return True

    def _start_process(self) -> None:
        if self._socket is not None:
            return

        if self._process is not None and self._process.poll() is None:
            return

        if self._running_in_job():
            if self._connect_to_external_broker():
                return
            raise RuntimeError(self.last_error_message())

        self._process = subprocess.Popen(
            [sys.executable, self._helper_script],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=current_dir,
            bufsize=1,
        )
        self._reader_thread = threading.Thread(
            target=self._reader_loop,
            args=(self._process.stdout,),
            name="hwp-helper-reader",
            daemon=True,
        )
        self._reader_thread.start()

    def _bootstrap_hwp_process(self) -> bool:
        bootstrap_script = (
            "import sys, time, win32com.client; "
            "win32com.client.Dispatch('HWPFrame.HwpObject'); "
            "sys.stdout.write('BOOTSTRAP_OK\\n'); "
            "sys.stdout.flush(); "
            "time.sleep(3)"
        )
        command = [sys.executable, "-X", "utf8", "-c", bootstrap_script]

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
                check=False,
                cwd=current_dir,
            )
        except Exception as exc:
            logger.error(f"Failed to bootstrap HWP process: {exc}", exc_info=True)
            self._set_last_hwp_error(f"Failed to bootstrap HWP process: {exc}")
            return False

        stdout = (result.stdout or "").strip()
        stderr = (result.stderr or "").strip()
        if stdout:
            logger.info(f"HWP bootstrap stdout: {stdout}")
        if stderr:
            logger.warning(f"HWP bootstrap stderr: {stderr}")

        if result.returncode != 0:
            message = f"HWP bootstrap process failed with exit code {result.returncode}"
            logger.error(message)
            self._set_last_hwp_error(message)
            return False

        return True

    @staticmethod
    def _should_retry_with_bootstrap(command: str, payload: dict, last_error: str) -> bool:
        if command != "ensure_root":
            return False
        if payload.get("root") not in {"controller", "table_tools"}:
            return False
        return bool(last_error)

    def _perform_request_locked(self, command: str, timeout: float, payload: dict):
        if (
            command == "ensure_root"
            and payload.get("root") in {"controller", "table_tools"}
            and self._in_startup_cooldown()
        ):
            message = self._startup_cooldown_message or "최근 한글 시작 실패로 잠시 재시도를 중단합니다."
            return {
                "ok": True,
                "exists": False,
                "last_error": message,
            }

        self._start_process()

        if self._socket is None and (self._process is None or self._process.stdin is None):
            raise RuntimeError("Failed to start HWP helper process")

        request_id = uuid.uuid4().hex
        waiter = queue.Queue(maxsize=1)
        with self._pending_lock:
            self._pending[request_id] = waiter

        message = {"id": request_id, "command": command}
        message.update(payload)

        try:
            self._send_request_message(message)
        except Exception as exc:
            with self._pending_lock:
                self._pending.pop(request_id, None)
            self.shutdown()
            raise RuntimeError(f"Failed to send request to HWP helper: {exc}") from exc

        try:
            response = waiter.get(timeout=timeout)
        except queue.Empty as exc:
            with self._pending_lock:
                self._pending.pop(request_id, None)
            raise TimeoutError("HWP 작업이 제한 시간 내에 완료되지 않았습니다.") from exc

        with self._pending_lock:
            self._pending.pop(request_id, None)

        last_error = response.get("last_error")
        self._set_last_hwp_error(last_error)
        return response

    def _send_request_message(self, message: dict) -> None:
        payload = json.dumps(message, ensure_ascii=False) + "\n"

        if self._socket is not None:
            self._socket.sendall(payload.encode("utf-8"))
            return

        if self._process is not None and self._process.stdin is not None:
            self._process.stdin.write(payload)
            self._process.stdin.flush()
            return

        raise RuntimeError("HWP helper transport is not available")

    def _reader_loop(self, stream) -> None:
        try:
            while stream is not None:
                line = stream.readline()
                if not line:
                    break

                try:
                    response = json.loads(line)
                except json.JSONDecodeError:
                    logger.error(f"Invalid helper response: {line!r}")
                    continue

                request_id = response.get("id")
                with self._pending_lock:
                    waiter = self._pending.get(request_id)

                if waiter is not None:
                    waiter.put(response)
        finally:
            self._fail_pending_requests("HWP helper process closed the transport")

    def _fail_pending_requests(self, message: str) -> None:
        with self._pending_lock:
            pending = list(self._pending.items())
            self._pending.clear()

        for _, waiter in pending:
            waiter.put({"ok": False, "error": message, "last_error": self._last_hwp_error})

    def request(self, command: str, timeout: float = 120.0, _allow_bootstrap_retry: bool = True, **payload):
        with self._lock:
            try:
                response = self._perform_request_locked(command, timeout, payload)
            except TimeoutError:
                if _allow_bootstrap_retry and self._is_root_ensure_request(command, payload) and timeout < ROOT_ENSURE_TIMEOUT_SECONDS:
                    logger.warning("Retrying HWP helper connection after timeout")
                    self.shutdown()
                    self._cleanup_hwp_processes()
                    time.sleep(1.0)
                    try:
                        response = self._perform_request_locked(command, min(timeout, 30.0), payload)
                    except TimeoutError:
                        self._set_startup_cooldown("한글 시작이 제한 시간 내에 완료되지 않아 잠시 재시도를 중단합니다.")
                        raise
                else:
                    if self._is_root_ensure_request(command, payload):
                        self.shutdown()
                        self._cleanup_hwp_processes()
                        self._set_startup_cooldown("한글 시작이 제한 시간 내에 완료되지 않아 잠시 재시도를 중단합니다.")
                    raise

            last_error = response.get("last_error") or response.get("error") or ""
            needs_bootstrap_retry = (
                _allow_bootstrap_retry
                and self._should_retry_with_bootstrap(command, payload, last_error)
                and (
                    not response.get("ok", False)
                    or not response.get("exists", True)
                )
            )

            if needs_bootstrap_retry:
                logger.warning("Retrying HWP helper connection after helper error")
                self.shutdown()
                self._cleanup_hwp_processes()
                time.sleep(1.0)
                retry_response = self._perform_request_locked(command, min(timeout, 30.0), payload)
                retry_last_error = retry_response.get("last_error") or retry_response.get("error") or ""
                self._set_last_hwp_error(retry_last_error)
                if not retry_response.get("ok", False):
                    self._set_startup_cooldown(retry_last_error or "한글 시작 실패로 잠시 재시도를 중단합니다.")
                    raise RuntimeError(retry_response.get("error", "Unknown HWP helper error"))
                if not retry_response.get("exists", True):
                    self._set_startup_cooldown(retry_last_error or "한글 시작 실패로 잠시 재시도를 중단합니다.")
                return retry_response

            if not response.get("ok", False):
                raise RuntimeError(response.get("error", "Unknown HWP helper error"))

            if response.get("exists", True):
                self._clear_startup_cooldown()

            return response

    def clear_state(self) -> None:
        try:
            self.request("clear_state", timeout=10)
        except Exception:
            self._set_last_hwp_error(None)

    def shutdown(self) -> None:
        process = self._process
        self._process = None
        sock = self._socket
        self._socket = None
        reader = self._socket_reader
        self._socket_reader = None

        if reader is not None:
            try:
                reader.close()
            except Exception:
                pass

        if sock is not None:
            try:
                sock.close()
            except Exception:
                pass

        if process is not None:
            try:
                if process.stdin:
                    process.stdin.close()
            except Exception:
                pass

            try:
                process.terminate()
                process.wait(timeout=5)
            except Exception:
                try:
                    process.kill()
                except Exception:
                    pass

        self._fail_pending_requests("HWP helper process was terminated")


class RemoteObjectProxy:
    """Proxy that marshals every access through the helper process."""

    def __init__(self, worker: HwpComWorker, object_id: str, label: str, is_root: bool = False):
        object.__setattr__(self, "_worker", worker)
        object.__setattr__(self, "_object_id", object_id)
        object.__setattr__(self, "_label", label)
        object.__setattr__(self, "_is_root", is_root)

    def __bool__(self) -> bool:
        try:
            if self._is_root:
                response = self._worker.request(
                    "ensure_root",
                    timeout=ROOT_ENSURE_TIMEOUT_SECONDS,
                    root=self._object_id,
                )
            else:
                response = self._worker.request("object_exists", timeout=30, object_id=self._object_id)
            return bool(response.get("exists"))
        except Exception:
            return False

    def __getattr__(self, name: str):
        response = self._worker.request(
            "inspect_attr",
            timeout=30,
            object_id=self._object_id,
            name=name,
        )
        kind = response.get("kind")

        if kind == "value":
            return response.get("value")

        if kind == "callable":
            def _call(*args, **kwargs):
                result = self._worker.request(
                    "call_method",
                    timeout=120,
                    object_id=self._object_id,
                    name=name,
                    args=args,
                    kwargs=kwargs,
                )
                if result.get("kind") == "value":
                    return result.get("value")

                return RemoteObjectProxy(
                    self._worker,
                    result["object_id"],
                    f"{self._label}.{name}()",
                )

            return _call

        return RemoteObjectProxy(
            self._worker,
            response["object_id"],
            f"{self._label}.{name}",
        )


hwp_worker = HwpComWorker()
atexit.register(hwp_worker.shutdown)


def _clear_hwp_thread_state():
    hwp_worker.clear_state()


def _set_last_hwp_error(message):
    hwp_worker._set_last_hwp_error(message)


def _connection_error_message() -> str:
    return hwp_worker.last_error_message()


def _connection_error_response() -> str:
    return f"Error: {_connection_error_message()}"


def _connection_error_payload() -> dict:
    return {"status": "error", "message": _connection_error_message()}


def _default_temp_document_path() -> str:
    temp_root = os.path.join(tempfile.gettempdir(), "hwp-mcp")
    os.makedirs(temp_root, exist_ok=True)
    return os.path.join(temp_root, f"document-{int(time.time() * 1000)}.hwp")


def get_hwp_controller():
    """Return a proxy backed by the helper process."""
    return RemoteObjectProxy(
        hwp_worker,
        "controller",
        "HwpController",
        is_root=True,
    )


def get_hwp_table_tools():
    """Return a proxy backed by the helper process."""
    return RemoteObjectProxy(
        hwp_worker,
        "table_tools",
        "HwpTableTools",
        is_root=True,
    )

@mcp.tool()
def hwp_create() -> str:
    """Create a new HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        if hwp.create_new_document():
            logger.info("Successfully created new document")
            return "New document created successfully"
        else:
            return "Error: Failed to create new document"
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_list_tabs() -> str:
    """
    현재 HWP 창에서 열려있는 탭(문서) 목록을 반환합니다.

    Returns:
        str: 탭 목록 또는 에러 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, documents = hwp.get_open_documents()
        if not success:
            return "Error: Failed to get document list"

        if not documents:
            return "열려있는 문서가 없습니다."

        result = f"열려있는 문서 ({len(documents)}개):\n"
        for doc in documents:
            marker = "→ " if doc["is_current"] else "  "
            result += f"{marker}[{doc['index']}] {doc['path']}\n"

        return result
    except Exception as e:
        logger.error(f"Error listing documents: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_switch_tab(index: int) -> str:
    """
    현재 HWP 창에서 특정 탭으로 전환합니다.

    Args:
        index: 탭 인덱스 (hwp_list_tabs로 확인)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, message = hwp.switch_document(index)
        if success:
            logger.info(message)
            return message
        else:
            return f"Error: {message}"
    except Exception as e:
        logger.error(f"Error switching document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_list_windows() -> str:
    """
    실행 중인 모든 HWP 창을 찾습니다.
    다른 창에 열려있는 HWP도 포함됩니다.

    Returns:
        str: 창 목록 또는 에러 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, instances = hwp.get_all_hwp_instances()
        if not success:
            return "Error: Failed to get HWP instances"

        if not instances:
            return "실행 중인 HWP 인스턴스가 없습니다."

        result = f"HWP 인스턴스 ({len(instances)}개):\n"
        for inst in instances:
            marker = "→ " if inst["is_current"] else "  "
            result += f"{marker}[{inst['index']}] {inst['title']} (hwnd: {inst['hwnd']})\n"

        return result
    except Exception as e:
        logger.error(f"Error listing HWP instances: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_switch_window(hwnd: int) -> str:
    """
    다른 HWP 창으로 전환합니다.

    Args:
        hwnd: 윈도우 핸들 (hwp_list_windows로 확인)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, message = hwp.connect_to_hwp_instance(hwnd)
        if success:
            logger.info(message)
            return message
        else:
            return f"Error: {message}"
    except Exception as e:
        logger.error(f"Error connecting to HWP instance: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_close_window(hwnd: int) -> str:
    """
    HWP 창을 닫습니다 (저장 안 함).

    Args:
        hwnd: 윈도우 핸들 (hwp_list_all_instances로 확인)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, message = hwp.close_hwp_window(hwnd)
        if success:
            logger.info(message)
            return message
        else:
            return f"Error: {message}"
    except Exception as e:
        logger.error(f"Error closing HWP window: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_open(path: str) -> str:
    """Open an existing HWP document."""
    try:
        if not path:
            return "Error: File path is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        if hwp.open_document(path):
            logger.info(f"Successfully opened document: {path}")
            return f"Document opened: {path}"
        else:
            return "Error: Failed to open document"
    except Exception as e:
        logger.error(f"Error opening document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_save(path: str = None) -> str:
    """Save the current HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        if path:
            if hwp.save_document(path):
                logger.info(f"Successfully saved document to: {path}")
                return f"Document saved to: {path}"
            else:
                return "Error: Failed to save document"
        else:
            temp_path = _default_temp_document_path()
            if hwp.save_document(temp_path):
                logger.info(f"Successfully saved document to temporary location: {temp_path}")
                return f"Document saved to: {temp_path}"
            else:
                return "Error: Failed to save document"
    except Exception as e:
        logger.error(f"Error saving document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_text(text: str, preserve_linebreaks: bool = True) -> str:
    """Insert text at the current cursor position."""
    try:
        if not text:
            return "Error: Text is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        # 현재 커서가 표 안에 있는지 확인
        is_in_table = False
        try:
            hwp.hwp.Run("TableCellBlock")
            hwp.hwp.Run("Cancel")
            is_in_table = True
        except Exception as e:
            logger.debug(f"표 셀 확인 실패 (무시): {e}")
            is_in_table = False

        # 줄바꿈 문자 처리
        if preserve_linebreaks and ('\n' in text or '\\n' in text):
            # 이스케이프된 줄바꿈 문자(\n)와 실제 줄바꿈 문자 모두 처리
            processed_text = text.replace('\\n', '\n')
            lines = processed_text.split('\n')
            
            success = True
            for i, line in enumerate(lines):
                if not hwp.insert_text(line):
                    success = False
                    break
                # 마지막 줄이 아니면 줄바꿈 삽입
                if i < len(lines) - 1:
                    hwp.insert_paragraph()
            
            if success:
                logger.info("Successfully inserted text with line breaks")
                return "Text with line breaks inserted successfully"
            else:
                return "Error: Failed to insert text with line breaks"
        else:
            if hwp.insert_text(text):
                # 표 안이 아닐 경우에만 커서를 오른쪽으로 이동
                if not is_in_table:
                    # 현재 위치 저장
                    current_pos = hwp.hwp.GetPos()
                    if current_pos:
                        # 텍스트 길이만큼 오른쪽으로 이동
                        for _ in range(len(text)):
                            hwp.hwp.Run("CharRight")
                logger.info("Successfully inserted text")
                return "Text inserted successfully"
            else:
                return "Error: Failed to insert text"
    except Exception as e:
        logger.error(f"Error inserting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_set_font(
    name: str = None, 
    size: int = None, 
    bold: bool = False, 
    italic: bool = False, 
    underline: bool = False,
    select_previous_text: bool = False
) -> str:
    """Set font properties for selected text."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        # 현재 선택된 텍스트에 대해 글자 모양 설정
        if hwp.set_font_style(
            font_name=name,
            font_size=size,
            bold=bold,
            italic=italic,
            underline=underline,
            select_previous_text=select_previous_text
        ):
            logger.info("Successfully set font")
            return "Font set successfully"
        else:
            return "Error: Failed to set font"
    
    except Exception as e:
        logger.error(f"Error setting font: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_table(rows: int, cols: int) -> str:
    """Insert a table at the current cursor position."""
    try:
        # HwpTableTools 인스턴스 가져오기
        table_tools = get_hwp_table_tools()
        if not table_tools:
            return _connection_error_response()
        
        return table_tools.insert_table(rows, cols)
    except Exception as e:
        logger.error(f"Error inserting table: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_paragraph() -> str:
    """Insert a new paragraph."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        if hwp.insert_paragraph():
            logger.info("Successfully inserted paragraph")
            return "Paragraph inserted successfully"
        else:
            return "Error: Failed to insert paragraph"
    except Exception as e:
        logger.error(f"Error inserting paragraph: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_get_text() -> str:
    """Get the text content of the current document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        text = hwp.get_text()
        if text is not None:
            logger.info("Successfully retrieved document text")
            return text
        else:
            return "Error: Failed to get document text"
    except Exception as e:
        logger.error(f"Error getting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_close_document(save: bool = False, suppress_dialog: bool = True) -> str:
    """
    현재 문서를 닫습니다.

    Args:
        save: 저장 후 닫을지 여부 (기본값: False)
        suppress_dialog: 저장 확인 팝업창 표시 안함 (기본값: True)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        if hwp.close_document(save, suppress_dialog):
            logger.info(f"Successfully closed document (save={save}, suppress_dialog={suppress_dialog})")
            return "Document closed successfully"
        else:
            return "Error: Failed to close document"
    except Exception as e:
        logger.error(f"Error closing document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_close_all_documents(save: bool = False, suppress_dialog: bool = True) -> str:
    """
    모든 문서를 닫습니다.

    Args:
        save: 저장 후 닫을지 여부 (기본값: False)
        suppress_dialog: 저장 확인 팝업창 표시 안함 (기본값: True)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        if hwp.close_all_documents(save, suppress_dialog):
            logger.info(f"Successfully closed all documents (save={save}, suppress_dialog={suppress_dialog})")
            return "All documents closed successfully"
        else:
            return "Error: Failed to close all documents"
    except Exception as e:
        logger.error(f"Error closing all documents: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_undo(count: int = 1) -> str:
    """
    실행 취소(Undo)를 수행합니다.

    Args:
        count: 취소할 횟수 (기본값: 1)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, message = hwp.undo(count)
        if success:
            logger.info(message)
            return message
        else:
            return f"Error: {message}"
    except Exception as e:
        logger.error(f"Error in undo: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_redo(count: int = 1) -> str:
    """
    다시 실행(Redo)을 수행합니다.

    Args:
        count: 다시 실행할 횟수 (기본값: 1)

    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, message = hwp.redo(count)
        if success:
            logger.info(message)
            return message
        else:
            return f"Error: {message}"
    except Exception as e:
        logger.error(f"Error in redo: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_find_text(text: str) -> str:
    """
    문서에서 텍스트를 찾습니다.

    Args:
        text: 찾을 텍스트

    Returns:
        str: 결과 메시지
    """
    try:
        if not text:
            return "Error: Text is required"

        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        if hwp.find_text(text):
            logger.info(f"Found text: {text}")
            return f"Text found: {text}"
        else:
            return f"Text not found: {text}"
    except Exception as e:
        logger.error(f"Error finding text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_replace_text(find: str, replace: str, replace_all: bool = True) -> str:
    """
    문서에서 텍스트를 찾아 바꿉니다.

    Args:
        find: 찾을 텍스트
        replace: 바꿀 텍스트
        replace_all: 모두 바꾸기 여부 (기본값: True)

    Returns:
        str: 결과 메시지
    """
    try:
        if not find:
            return "Error: Find text is required"

        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        if hwp.replace_text(find, replace, replace_all):
            logger.info(f"Replaced text: '{find}' -> '{replace}' (replace_all={replace_all})")
            return f"Text replaced: '{find}' -> '{replace}'"
        else:
            return f"Text not found or replace failed: {find}"
    except Exception as e:
        logger.error(f"Error replacing text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_ping_pong(message: str = "핑") -> str:
    """
    핑퐁 테스트용 함수입니다. 핑을 보내면 퐁을 응답하고, 퐁을 보내면 핑을 응답합니다.

    Args:
        message: 테스트 메시지 (기본값: "핑")

    Returns:
        str: 응답 메시지
    """
    try:
        logger.info(f"핑퐁 테스트 함수 호출됨: 메시지 - {message}")

        if message == "핑":
            response = "퐁"
        elif message == "퐁":
            response = "핑"
        else:
            response = f"모르는 메시지입니다: {message} (핑 또는 퐁을 보내주세요)"

        current_time = time.strftime("%Y-%m-%d %H:%M:%S")

        result = {
            "response": response,
            "original_message": message,
            "timestamp": current_time
        }

        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        logger.error(f"핑퐁 테스트 함수 오류: {str(e)}", exc_info=True)
        return f"테스트 오류 발생: {str(e)}"

@mcp.tool()
def hwp_create_table_with_data(rows: int, cols: int, data = None, has_header: bool = False) -> str:
    """
    pywin32를 사용하여 현재 커서 위치에 표를 생성하고 데이터를 채웁니다.
    
    Args:
        rows: 표의 행 수
        cols: 표의 열 수
        data: 표에 채울 데이터 (JSON 문자열 또는 파이썬 리스트)
        has_header: 첫 번째 행을 헤더로 처리할지 여부
        
    Returns:
        str: 결과 메시지
    """
    try:
        # HwpTableTools 인스턴스 가져오기
        table_tools = get_hwp_table_tools()
        if not table_tools:
            return _connection_error_response()
        
        # 현재 커서가 표 안에 있는지 확인
        hwp = get_hwp_controller()
        is_in_table = False
        try:
            hwp.hwp.Run("TableCellBlock")
            hwp.hwp.Run("Cancel")
            is_in_table = True
        except Exception as e:
            logger.debug(f"표 셀 확인 실패 (무시): {e}")
            is_in_table = False

        # 표 안에 있지 않은 경우에만 새 표 생성
        if not is_in_table:
            # 표 생성
            if not table_tools.insert_table(rows, cols):
                return "Error: Failed to create table"
        
        # 데이터가 있는 경우 표 채우기
        if data is not None:
            # 데이터 형식 로깅
            logger.info(f"Create table with data type: {type(data)}, data: {str(data)[:100]}...")
            
            # 데이터가 이미 리스트 형태인 경우
            if isinstance(data, list):
                logger.info("Data is already a list, using directly")
                processed_data = data
            # 데이터가 문자열인 경우 JSON 파싱 시도
            elif isinstance(data, str):
                try:
                    import json
                    try:
                        processed_data = json.loads(data)
                        logger.info(f"Successfully parsed JSON data with {len(processed_data)} rows")
                    except json.JSONDecodeError as e:
                        logger.error(f"JSON 파싱 오류: {str(e)}")
                        try:
                            import ast
                            processed_data = ast.literal_eval(data)
                            logger.info(f"Successfully parsed data with literal_eval")
                        except Exception as e2:
                            logger.error(f"리터럴 평가 오류: {str(e2)}")
                            return f"표는 생성되었으나 JSON 데이터 파싱 오류: {str(e)}"
                except Exception as e:
                    logger.error(f"데이터 파싱 오류: {str(e)}", exc_info=True)
                    return f"표는 생성되었으나 데이터 파싱 오류: {str(e)}"
            else:
                return f"표는 생성되었으나 지원되지 않는 데이터 유형: {type(data)}"
            
            # 데이터 구조 유효성 검사
            if not isinstance(processed_data, list):
                return f"표는 생성되었으나 데이터가 리스트 형식이 아닙니다: {type(processed_data)}"
            
            if len(processed_data) == 0:
                return "표는 생성되었으나 데이터 리스트가 비어 있습니다."
            
            # 모든 행이 리스트인지 확인 및 변환
            for i, row in enumerate(processed_data):
                if not isinstance(row, list):
                    logger.warning(f"Row {i} is not a list, converting: {row}")
                    processed_data[i] = [row]
            
            # 모든 데이터를 문자열로 변환
            string_data = []
            for row in processed_data:
                string_row = [str(cell) if cell is not None else "" for cell in row]
                string_data.append(string_row)
            
            # 표에 데이터 채우기
            if table_tools.fill_table_with_data(string_data, 1, 1, has_header):
                return f"표 생성 및 데이터 입력 완료 ({rows}x{cols})"
            else:
                return "표는 생성되었으나 데이터 입력에 실패했습니다."
        
        return f"표 생성 완료 ({rows}x{cols})"
    except Exception as e:
        logger.error(f"표 생성 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_create_complete_document(document_spec: dict) -> dict:
    """
    전체 문서를 한 번의 호출로 작성합니다. 문서 구조, 내용 및 서식을 JSON으로 정의하여 전달합니다.
    
    Args:
        document_spec (dict): 문서 사양을 담은 딕셔너리. 다음과 같은 구조를 가집니다:
            {
                "title": "문서 제목",             # 선택 사항
                "filename": "저장할_파일명.hwp",  # 선택 사항, 저장할 경우
                "elements": [                    # 필수: 문서를 구성하는 요소 배열
                    {
                        "type": "heading",       # 요소 유형 (heading, text, paragraph, table, etc.)
                        "content": "제목",        # 요소 내용
                        "properties": {          # 요소 속성 (선택 사항)
                            "font_size": 16,
                            "bold": true,
                            ...
                        }
                    },
                    ...
                ],
                "special_type": {               # 특수 문서 유형 (선택 사항)
                    "type": "report",           # 보고서 등 특수 문서 유형
                    "params": { ... }           # 특수 문서에 필요한 매개변수
                },
                "save": true                    # 저장 여부 (선택 사항)
            }
    
    Returns:
        dict: 문서 생성 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_payload()
        
        # 새 문서 생성
        if not hwp.create_new_document():
            return {"status": "error", "message": "Failed to create new document"}
        
        # 문서 사양 유효성 검사
        if not document_spec:
            return {"status": "error", "message": "Document specification is required"}
        
        if "special_type" in document_spec:
            # 특수 문서 유형 처리 (보고서 등)
            special_type = document_spec["special_type"]
            special_type_name = special_type.get("type", "")
            special_params = special_type.get("params", {})
            
            # 보고서 처리
            if special_type_name == "report":
                return _create_report(hwp, special_params, document_spec)
            
            # 편지 처리
            elif special_type_name == "letter":
                return _create_letter(hwp, special_params, document_spec)
            
            else:
                return {"status": "error", "message": f"Unknown special document type: {special_type_name}"}
        
        # 일반 문서 처리
        elif "elements" in document_spec:
            elements = document_spec.get("elements", [])
            
            # 문서 요소 처리
            for element in elements:
                element_type = element.get("type", "")
                content = element.get("content", "")
                properties = element.get("properties", {})
                
                # 요소 유형에 따른 처리
                if element_type == "heading":
                    # 제목 스타일 설정
                    font_size = properties.get("font_size", 16)
                    bold = properties.get("bold", True)
                    hwp.set_font(None, font_size, bold, False)
                    hwp.insert_text(content)
                    hwp.insert_paragraph()
                
                elif element_type == "text":
                    # 텍스트 스타일 설정
                    font_size = properties.get("font_size", 10)
                    bold = properties.get("bold", False)
                    italic = properties.get("italic", False)
                    hwp.set_font(None, font_size, bold, italic)
                    hwp.insert_text(content)
                
                elif element_type == "paragraph":
                    hwp.insert_paragraph()
                
                elif element_type == "table":
                    rows = properties.get("rows", 0)
                    cols = properties.get("cols", 0)
                    data = properties.get("data", [])
                    
                    if rows > 0 and cols > 0:
                        hwp.insert_table(rows, cols)
                        
                        # 테이블 데이터 채우기 (구현 필요)
                        # 현재는 표만 생성하고 데이터는 처리하지 않음
                
                else:
                    logger.warning(f"Unknown element type: {element_type}")
        
        else:
            return {"status": "error", "message": "Document must contain 'elements' or 'special_type'"}
        
        # 문서 저장
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "generated_document.hwp")
            if hwp.save_document(filename):
                return {
                    "status": "success", 
                    "message": "Document created and saved successfully",
                    "saved_path": filename
                }
            else:
                return {
                    "status": "partial_success", 
                    "message": "Document created but failed to save"
                }
        
        return {"status": "success", "message": "Document created successfully"}
    
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

def _create_report(hwp, params, document_spec):
    """보고서 문서를 생성합니다."""
    try:
        title = params.get("title", "보고서 제목")
        author = params.get("author", "작성자")
        date = params.get("date", time.strftime("%Y년 %m월 %d일"))
        sections = params.get("sections", [{"title": "섹션 제목", "content": "섹션 내용"}])
        
        # 제목 페이지
        hwp.set_font(None, 22, True, False)
        hwp.insert_text(title)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        hwp.set_font(None, 14, False, False)
        hwp.insert_text(f"작성자: {author}")
        hwp.insert_paragraph()
        hwp.insert_text(f"작성일: {date}")
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 각 섹션
        for section in sections:
            section_title = section.get("title", "")
            section_content = section.get("content", "")
            
            hwp.set_font(None, 16, True, False)
            hwp.insert_text(section_title)
            hwp.insert_paragraph()
            
            hwp.set_font(None, 12, False, False)
            hwp.insert_text(section_content)
            hwp.insert_paragraph()
            hwp.insert_paragraph()
        
        # 문서 저장
        result = {"status": "success", "message": "Report created successfully"}
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "report.hwp")
            if hwp.save_document(filename):
                result["saved_path"] = filename
            else:
                result["message"] = "Report created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating report: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

def _create_letter(hwp, params, document_spec):
    """편지 문서를 생성합니다."""
    try:
        title = params.get("title", "제목 없음")
        recipient = params.get("recipient", "받는 사람")
        content = params.get("content", "내용을 입력하세요.")
        sender = params.get("sender", "보내는 사람")
        date = params.get("date", time.strftime("%Y년 %m월 %d일"))
        
        # 제목 (굵게, 크게)
        hwp.set_font(None, 16, True, False)
        hwp.insert_text(title)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 받는 사람
        hwp.set_font(None, 12, False, False)
        hwp.insert_text(f"받는 사람: {recipient}")
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 내용
        hwp.set_font(None, 12, False, False)
        hwp.insert_text(content)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 날짜 (오른쪽 정렬)
        # 오른쪽 정렬은 현재 구현되어 있지 않으므로 공백으로 대체
        hwp.set_font(None, 12, False, False)
        hwp.insert_text("".ljust(40) + date)
        hwp.insert_paragraph()
        
        # 보내는 사람 (오른쪽 정렬, 굵게)
        hwp.set_font(None, 12, True, False)
        hwp.insert_text("".ljust(40) + sender)
        
        # 문서 저장
        result = {"status": "success", "message": "Letter created successfully"}
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "letter.hwp")
            if hwp.save_document(filename):
                result["saved_path"] = filename
            else:
                result["message"] = "Letter created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating letter: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

@mcp.tool()
def hwp_create_document_from_text(content: str, title: str = None, format_content: bool = True, save_filename: str = None, preserve_linebreaks: bool = True) -> dict:
    """
    단일 문자열로 된 텍스트 내용으로 문서를 생성합니다.
    
    Args:
        content (str): 문서 내용 (형식을 자동으로 감지하고 처리)
        title (str, optional): 문서 제목. 없으면 첫 줄을 제목으로 사용.
        format_content (bool): 내용 자동 포맷팅 여부 (줄바꿈, 문단 구분 등)
        save_filename (str, optional): 저장할 파일 이름. 제공되지 않으면 저장하지 않음.
        preserve_linebreaks (bool): 줄바꿈 유지 여부. True이면 원본 텍스트의 모든 줄바꿈 유지.
        
    Returns:
        dict: 문서 생성 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_payload()
        
        # 새 문서 생성
        if not hwp.create_new_document():
            return {"status": "error", "message": "Failed to create new document"}
        
        # 내용이 없는 경우
        if not content:
            return {"status": "error", "message": "Document content is required"}
        
        # 내용을 줄로 분리
        lines = content.split('\n')
        
        # 빈 줄을 기준으로 블록 구분
        blocks = []
        current_block = []
        
        for line in lines:
            if line.strip():  # 빈 줄이 아닌 경우
                current_block.append(line)
            else:  # 빈 줄인 경우 블록 구분
                if current_block:
                    blocks.append(current_block)
                    current_block = []
        
        # 마지막 블록 추가
        if current_block:
            blocks.append(current_block)
        
        # 제목 처리
        if not title and blocks:
            # 첫 번째 블록의 첫 번째 줄을 제목으로 사용
            title = blocks[0][0]
            if len(blocks[0]) > 1:
                blocks[0] = blocks[0][1:]  # 첫 번째 줄 제거
            else:
                blocks = blocks[1:]  # 첫 번째 블록 제거
        
        # 제목 추가
        if title:
            # 먼저 폰트 설정 후 텍스트 입력 (수정된 방식)
            hwp.set_font(None, 16, True, False)
            hwp.insert_text(title)
            hwp.insert_paragraph()
            hwp.insert_paragraph()
        
        # 내용 자동 포맷팅
        if format_content:
            # 블록 단위로 처리
            for block in blocks:
                # 블록 내 첫 번째 줄로 블록 유형 판단
                first_line = block[0].strip() if block else ""
                
                # 제목 형식 감지 (예: #으로 시작하면 제목)
                if first_line.startswith('#'):
                    level = 0
                    for char in first_line:
                        if char == '#':
                            level += 1
                        else:
                            break
                    
                    heading_text = first_line[level:].strip()
                    font_size = max(11, 16 - (level - 1))  # 제목 레벨에 따라 글자 크기 조정
                    
                    # 먼저 폰트 설정 후 텍스트 입력 (수정된 방식)
                    hwp.set_font(None, font_size, True, False)
                    hwp.insert_text(heading_text)
                    hwp.insert_paragraph()
                    
                    # 제목 이후의 줄들 처리 (있을 경우)
                    if len(block) > 1:
                        hwp.set_font(None, 11, False, False)
                        for line in block[1:]:
                            hwp.insert_text(line)
                            hwp.insert_paragraph()
                
                # 글머리 기호 감지 (예: - 또는 * 으로 시작하면 글머리 기호)
                elif first_line.startswith(('-', '*', '•')):
                    hwp.set_font(None, 11, False, False)
                    for line in block:
                        line_stripped = line.strip()
                        if line_stripped.startswith(('-', '*', '•')):
                            content_text = line_stripped[1:].strip()
                            hwp.insert_text(f"• {content_text}")
                        else:
                            hwp.insert_text(line_stripped)
                        hwp.insert_paragraph()
                
                # 시 또는 줄바꿈이 중요한 텍스트 (각 줄을 개별적으로 처리)
                elif preserve_linebreaks:
                    hwp.set_font(None, 11, False, False)
                    for line in block:
                        hwp.insert_text(line)
                        hwp.insert_paragraph()
                
                # 일반 텍스트 (블록 전체를 하나의 단락으로 처리)
                else:
                    hwp.set_font(None, 11, False, False)
                    block_text = '\n'.join(block)
                    hwp.insert_text(block_text)
                    hwp.insert_paragraph()
                
                # 블록 사이에 추가 줄바꿈
                hwp.insert_paragraph()
        
        # 자동 포맷팅 없이 그대로 삽입 (줄바꿈 보존)
        else:
            hwp.set_font(None, 11, False, False)
            for line in lines:
                if line.strip():  # 내용이 있는 줄
                    hwp.insert_text(line)
                hwp.insert_paragraph()  # 빈 줄이든 내용이 있는 줄이든 항상 줄바꿈
        
        # 문서 저장
        result = {"status": "success", "message": "Document created from text successfully"}
        if save_filename:
            if hwp.save_document(save_filename):
                result["saved_path"] = save_filename
            else:
                result["message"] = "Document created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating document from text: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

@mcp.tool()
def hwp_batch_operations(operations: list) -> dict:
    """
    여러 HWP 작업을 한 번의 호출로 일괄 처리합니다.
    
    Args:
        operations (list): 실행할 작업 목록. 각 작업은 다음 형식의 딕셔너리입니다:
            {
                "operation": "작업명", # 예: "create", "set_font", "insert_text" 등
                "params": {파라미터 딕셔너리}  # 해당 작업에 필요한 파라미터
            }
    
    Returns:
        dict: 각 작업의 실행 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_payload()
        
        results = []
        
        for op in operations:
            operation = op.get("operation", "")
            params = op.get("params", {})
            
            result = {"operation": operation, "status": "success", "message": ""}
            
            try:
                if operation == "create":
                    if hwp.create_new_document():
                        result["message"] = "New document created successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to create new document"
                
                elif operation == "open":
                    path = params.get("path", "")
                    if not path:
                        result["status"] = "error"
                        result["message"] = "File path is required"
                    elif hwp.open_document(path):
                        result["message"] = f"Document opened: {path}"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to open document"
                
                elif operation == "save":
                    path = params.get("path", None)
                    if path and hwp.save_document(path):
                        result["message"] = f"Document saved to: {path}"
                    elif not path:
                        temp_path = _default_temp_document_path()
                        if hwp.save_document(temp_path):
                            result["message"] = f"Document saved to: {temp_path}"
                            result["path"] = temp_path
                        else:
                            result["status"] = "error"
                            result["message"] = "Failed to save document"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to save document"
                
                elif operation == "insert_text":
                    text = params.get("text", "")
                    preserve_linebreaks = params.get("preserve_linebreaks", True)
                    
                    if not text:
                        result["status"] = "error"
                        result["message"] = "Text is required"
                    elif preserve_linebreaks and ('\n' in text or '\\n' in text):
                        # 줄바꿈 보존 처리 개선
                        # 이스케이프된 줄바꿈 문자(\n)와 실제 줄바꿈 문자 모두 처리
                        # 먼저 이스케이프된 줄바꿈 문자를 실제 줄바꿈으로 변환
                        processed_text = text.replace('\\n', '\n')
                        lines = processed_text.split('\n')
                        
                        success = True
                        for i, line in enumerate(lines):
                            if not hwp.insert_text(line):
                                success = False
                                break
                            # 마지막 줄이 아니면 줄바꿈 삽입
                            if i < len(lines) - 1:
                                hwp.insert_paragraph()
                        
                        if success:
                            result["message"] = "Text with line breaks inserted successfully"
                        else:
                            result["status"] = "error"
                            result["message"] = "Failed to insert text with line breaks"
                    elif hwp.insert_text(text):
                        result["message"] = "Text inserted successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to insert text"
                
                elif operation == "set_font":
                    name = params.get("name", None)
                    size = params.get("size", None)
                    bold = params.get("bold", False)
                    italic = params.get("italic", False)
                    underline = params.get("underline", False)
                    select_previous_text = params.get("select_previous_text", False)
                    
                    if hwp.set_font_style(font_name=name, font_size=size, bold=bold, italic=italic, underline=underline, select_previous_text=select_previous_text):
                        result["message"] = "Font set successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to set font"
                
                elif operation == "insert_paragraph":
                    count = params.get("count", 1)  # 여러 줄 삽입 가능
                    success = True
                    for _ in range(count):
                        if not hwp.insert_paragraph():
                            success = False
                            break
                    
                    if success:
                        result["message"] = f"{count} paragraph(s) inserted successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to insert paragraph"
                
                elif operation == "insert_table":
                    rows = params.get("rows", 0)
                    cols = params.get("cols", 0)
                    data = params.get("data", [])
                    has_header = params.get("has_header", False)
                    
                    table_tools = get_hwp_table_tools()
                    if not table_tools:
                        result["status"] = "error"
                        result["message"] = _connection_error_message()
                    elif rows <= 0 or cols <= 0:
                        result["status"] = "error"
                        result["message"] = "Valid rows and cols are required"
                    else:
                        # 데이터가 있으면 테이블 생성 후 데이터 채우기
                        if data:
                            resp = table_tools.create_table_with_data(rows, cols, json.dumps(data) if isinstance(data, list) else data, has_header)
                            result["message"] = resp
                            if resp.startswith("Error"):
                                result["status"] = "error"
                        else:
                            resp = table_tools.insert_table(rows, cols)
                            result["message"] = resp
                            if resp.startswith("Error"):
                                result["status"] = "error"
                
                elif operation == "set_table_cell_text":
                    row = params.get("row", 0)
                    col = params.get("col", 0)
                    text = params.get("text", "")
                    
                    table_tools = get_hwp_table_tools()
                    if not table_tools:
                        result["status"] = "error"
                        result["message"] = _connection_error_message()
                    elif row <= 0 or col <= 0:
                        result["status"] = "error"
                        result["message"] = "Valid row and col are required"
                    else:
                        resp = table_tools.set_cell_text(row, col, text)
                        result["message"] = resp
                        if resp.startswith("Error"):
                            result["status"] = "error"
                
                elif operation == "merge_table_cells":
                    start_row = params.get("start_row", 0)
                    start_col = params.get("start_col", 0)
                    end_row = params.get("end_row", 0)
                    end_col = params.get("end_col", 0)
                    
                    table_tools = get_hwp_table_tools()
                    if not table_tools:
                        result["status"] = "error"
                        result["message"] = _connection_error_message()
                    elif start_row <= 0 or start_col <= 0 or end_row <= 0 or end_col <= 0:
                        result["status"] = "error"
                        result["message"] = "Valid cell coordinates are required"
                    else:
                        resp = table_tools.merge_cells(start_row, start_col, end_row, end_col)
                        result["message"] = resp
                        if resp.startswith("Error"):
                            result["status"] = "error"
                
                elif operation == "get_text":
                    text = hwp.get_text()
                    if text is not None:
                        result["message"] = "Text retrieved successfully"
                        result["text"] = text
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to retrieve text"
                
                elif operation == "close":
                    save = params.get("save", True)
                    if hwp.disconnect():
                        result["message"] = "Document closed successfully"
                        _clear_hwp_thread_state()
                        _set_last_hwp_error(None)
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to close document"
                
                # 새로 추가: 문서 한 번에 생성
                elif operation == "create_document_from_text":
                    content = params.get("content", "")
                    title = params.get("title", None)
                    format_content = params.get("format_content", True)
                    save_filename = params.get("save_filename", None)
                    preserve_linebreaks = params.get("preserve_linebreaks", True)
                    
                    if not content:
                        result["status"] = "error"
                        result["message"] = "Document content is required"
                    else:
                        # 내부적으로 기존 함수 호출
                        doc_result = hwp_create_document_from_text(
                            content=content,
                            title=title,
                            format_content=format_content,
                            save_filename=save_filename,
                            preserve_linebreaks=preserve_linebreaks
                        )
                        
                        result["status"] = doc_result.get("status", "error")
                        result["message"] = doc_result.get("message", "Unknown error")
                        if "saved_path" in doc_result:
                            result["saved_path"] = doc_result["saved_path"]
                
                else:
                    result["status"] = "error"
                    result["message"] = f"Unknown operation: {operation}"
            
            except Exception as e:
                result["status"] = "error"
                result["message"] = f"Error in operation '{operation}': {str(e)}"
            
            results.append(result)
        
        return {"status": "success", "results": results}
    
    except Exception as e:
        logger.error(f"Error in batch operations: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

@mcp.tool()
def hwp_fill_table_with_data(data, start_row: int = 1, start_col: int = 1, has_header: bool = False) -> str:
    """
    이미 존재하는 표에 데이터를 채웁니다.
    
    Args:
        data: 표에 채울 데이터 (JSON 문자열 또는 2차원 리스트)
        start_row: 시작 행 번호 (1부터 시작)
        start_col: 시작 열 번호 (1부터 시작)
        has_header: 첫 번째 행을 헤더로 처리할지 여부
        
    Returns:
        str: 결과 메시지
    """
    try:
        table_tools = get_hwp_table_tools()
        if not table_tools:
            return _connection_error_response()
        
        # 데이터 형식 로깅
        logger.info(f"Received data type: {type(data)}, data: {str(data)[:100]}...")
        
        # 데이터 처리
        processed_data = []
        
        # 이미 리스트 형태인 경우
        if isinstance(data, list):
            logger.info("Data is already a list, processing directly")
            processed_data = data
        # 문자열인 경우 JSON 파싱 시도
        elif isinstance(data, str):
            try:
                import json
                
                # JSON 파싱 시도
                try:
                    processed_data = json.loads(data)
                    logger.info(f"Successfully parsed JSON data with {len(processed_data)} rows")
                except json.JSONDecodeError as e:
                    logger.error(f"JSON 디코딩 오류: {str(e)}")
                    
                    # 특수 케이스: 1부터 10까지 세로로 채우는 요청인 경우
                    if "1부터 10까지" in data and "세로" in data:
                        logger.info("특수 케이스 감지: 1부터 10까지 세로로 채우기")
                        processed_data = []
                        for i in range(1, 11):
                            processed_data.append([str(i)])
                    else:
                        # 마지막 시도: 리터럴 평가
                        try:
                            import ast
                            processed_data = ast.literal_eval(data)
                            logger.info(f"Successfully parsed data with literal_eval: {len(processed_data)} rows")
                        except Exception as e:
                            logger.debug(f"literal_eval 파싱 실패: {e}")
                            # 단순 문자열을 직접 파싱
                            try:
                                # 문자열에서 쉼표로 구분된 항목 추출 시도
                                if "," in data:
                                    items = [item.strip() for item in data.split(",")]
                                    processed_data = [[item] for item in items]
                                else:
                                    # 단일 값인 경우
                                    processed_data = [[data]]
                            except Exception as parse_err:
                                return f"Error: Failed to parse string data - {str(parse_err)}"
            except Exception as e:
                logger.error(f"데이터 파싱 오류: {str(e)}", exc_info=True)
                return f"Error: Failed to parse data - {str(e)}"
        else:
            return f"Error: Unsupported data type: {type(data)}"
        
        # 데이터 구조 유효성 검사
        if not isinstance(processed_data, list):
            logger.error(f"Processed data is not a list: {type(processed_data)}")
            return f"Error: Data must be a list, got {type(processed_data)}"
        
        if len(processed_data) == 0:
            logger.error("Empty data list")
            return "Error: Empty data list"
        
        # 모든 행이 리스트인지 확인 및 변환
        for i, row in enumerate(processed_data):
            if not isinstance(row, list):
                logger.warning(f"Row {i} is not a list, converting to list: {row}")
                processed_data[i] = [row]  # 리스트가 아닌 항목을 리스트로 변환
        
        # 모든 데이터를 문자열로 변환
        final_data = []
        for row in processed_data:
            final_row = [str(cell) if cell is not None else "" for cell in row]
            final_data.append(final_row)
        
        logger.info(f"Final processed data has {len(final_data)} rows")
        
        # 표에 데이터 채우기
        result = table_tools.fill_table_with_data(final_data, start_row, start_col, has_header)
        logger.info(f"Table filling result: {result}")
        return result
        
    except Exception as e:
        logger.error(f"표 데이터 입력 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_navigate(direction: str) -> str:
    """
    표에서 지정된 방향으로 이동하고 현재 셀의 내용을 반환합니다.
    표 구조를 실시간으로 탐색할 때 유용합니다.

    Args:
        direction: 이동 방향 ("left", "right", "up", "down")

    Returns:
        str: 이동 후 현재 셀의 내용
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, dir_used, cell_text = hwp.navigate_and_get_cell(direction)

        if success:
            return f"[{dir_used}] → 현재 셀: 「{cell_text}」"
        else:
            return f"Error: {cell_text}"

    except Exception as e:
        logger.error(f"네비게이션 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"


@mcp.tool()
def hwp_find_and_show_cell(text: str) -> str:
    """
    텍스트를 찾고 해당 셀의 내용을 반환합니다.
    표 구조 탐색의 시작점으로 유용합니다.

    Args:
        text: 찾을 텍스트

    Returns:
        str: 찾은 셀의 내용
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        success, cell_text = hwp.find_and_get_cell(text)

        if success:
            return f"'{text}' 찾음 → 현재 셀: 「{cell_text}」"
        else:
            return f"Error: {cell_text}"

    except Exception as e:
        logger.error(f"찾기 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"


@mcp.tool()
def hwp_table_view(depth: int = 1) -> dict:
    """
    현재 위치 기준으로 주변 셀들의 내용을 가져옵니다.
    표 구조를 파악할 때 유용합니다.

    **사용 예시:**
    ```
    # 상하좌우 1칸씩 보기
    hwp_table_view(depth=1)
    # 결과: {"center": "현재", "up_1": "위", "down_1": "아래", ...}

    # 2칸씩 보기
    hwp_table_view(depth=2)
    # 결과: {"center": "현재", "up_1": "위1", "up_2": "위2", ...}
    ```

    **hwp_navigate, hwp_find_and_show_cell과 함께 사용:**
    1. hwp_find_and_show_cell("레이블")로 시작점 찾기
    2. hwp_table_view()로 주변 구조 파악
    3. hwp_navigate("방향")으로 이동
    4. hwp_fill_cells()로 값 입력

    Args:
        depth: 탐색 깊이 (기본값: 1, 최대 권장: 3)

    Returns:
        dict: 셀 내용 딕셔너리
            - "center": 현재 셀
            - "up_1", "up_2", ...: 위쪽 셀들
            - "down_1", "down_2", ...: 아래쪽 셀들
            - "left_1", "left_2", ...: 왼쪽 셀들
            - "right_1", "right_2", ...: 오른쪽 셀들
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": _connection_error_message()}

        # depth 제한
        depth = min(max(depth, 1), 5)

        success, result = hwp.get_table_view(depth)

        if success:
            logger.info(f"테이블 뷰 가져오기 성공 (depth={depth})")
            return result
        else:
            return result

    except Exception as e:
        logger.error(f"테이블 뷰 오류: {str(e)}", exc_info=True)
        return {"error": str(e)}


@mcp.tool()
def hwp_fill_cells(
    path_value_map: dict,
    mode: str = "replace"
) -> str:
    """
    표에서 경로를 따라 셀에 값을 입력합니다. 단일/배치 자동 인식.

    **⚠️ 중요: 반드시 표 구조를 먼저 파악하세요!**
    1. hwp_find_and_show_cell("레이블")로 시작점 찾기
    2. hwp_table_view(depth=2)로 주변 셀 구조 확인
    3. 구조를 이해한 후 hwp_fill_cells 호출

    **방향 키워드:** <left>, <right>, <up>, <down>
    - 텍스트 찾기 후 방향 이동을 조합하여 정확한 셀 탐색

    **사용 예시:**
    ```
    # 레이블 오른쪽 셀에 값 입력
    hwp_fill_cells({"이름 > <right>": "홍길동"})

    # 헤더 아래 셀에 값 입력
    hwp_fill_cells({"항목1 > <down>": "값1"})

    # 여러 셀 배치 입력
    hwp_fill_cells({
        "이름 > <right>": "홍길동",
        "연락처 > <right>": "010-1234-5678",
        "주소 > <right>": "서울시"
    })

    # 복잡한 경로 (여러 번 이동)
    hwp_fill_cells({
        "합계 > <down> > <down>": "100",
        "비고 > <right> > <down>": "완료"
    })

    # 기존 내용 앞에 추가 (예: "원" → "1000원")
    hwp_fill_cells({"금액 > <right>": "1000"}, mode="prepend")
    ```

    **경로 형식:**
    - "텍스트 > <방향> > <방향> > ..."
    - 구분자: " > " 또는 "/"
    - 첫 번째 항목: 찾을 텍스트 (표 안의 셀)
    - 이후 항목: 방향 키워드로 이동

    Args:
        path_value_map: 경로와 값의 딕셔너리
            - 키: 경로 문자열 (텍스트와 방향 키워드 조합)
            - 값: 입력할 값
        mode: 입력 모드
            - "replace": 기존 내용 삭제 후 입력 (기본값)
            - "prepend": 기존 내용 앞에 추가 (예: "명" → "3명")
            - "append": 기존 내용 뒤에 추가

    Returns:
        str: 처리 결과 메시지
    """
    try:
        if not path_value_map:
            return "Error: path_value_map이 필요합니다."

        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()

        # 배치 처리 (direction은 경로에서 결정되므로 "right"를 기본값으로)
        results = hwp.fill_cells_by_path_batch(path_value_map, "right", mode)

        # 결과 정리
        success_count = 0
        fail_count = 0
        messages = []

        for path_str, (success, message) in results.items():
            if success:
                success_count += 1
                messages.append(f"✓ {path_str}: {message}")
            else:
                fail_count += 1
                messages.append(f"✗ {path_str}: {message}")

        summary = f"\n총 {success_count}개 성공, {fail_count}개 실패"
        result_message = "\n".join(messages) + summary

        logger.info(f"셀 채우기 완료: {success_count}개 성공, {fail_count}개 실패")
        return result_message

    except Exception as e:
        logger.error(f"셀 채우기 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"


@mcp.tool()
def hwp_fill_column_numbers(start: int = 1, end: int = 10, column: int = 1, from_first_cell: bool = True) -> str:
    """
    표의 특정 열에 시작 숫자부터 끝 숫자까지 세로로 채웁니다.
    
    Args:
        start: 시작 숫자 (기본값: 1)
        end: 끝 숫자 (기본값: 10)
        column: 숫자를 채울 열 번호 (1부터 시작, 기본값: 1)
        from_first_cell: 정확히 표의 첫 번째 셀부터 시작할지 여부 (기본값: True)
    
    Returns:
        str: 결과 메시지
    """
    try:
        # HWP 컨트롤러 가져오기
        hwp = get_hwp_controller()
        if not hwp:
            return _connection_error_response()
        
        # 표 선택 (현재 커서 위치에 표가 있어야 함)
        logger.info(f"테이블 열에 숫자 채우기: 열 {column}, {start}부터 {end}까지")
        
        # 표의 첫 번째 셀로 이동 (문서의 표 맨 앞)
        hwp.hwp.Run("TableColBegin")
        
        # from_first_cell이 False인 경우에만 아래로 이동
        if not from_first_cell:
            hwp.hwp.Run("TableLowerCell")
        
        # 지정된 열로 이동
        for _ in range(column - 1):
            hwp.hwp.Run("TableRightCell")
        
        # 각 행에 숫자 채우기
        for num in range(start, end + 1):
            # 셀 선택 및 내용 지우기
            hwp.hwp.Run("Select")
            hwp.hwp.Run("Delete")
            
            # 셀에 숫자 입력
            hwp.insert_text(str(num))
            
            # 다음 행으로 이동 (마지막 행이 아닌 경우)
            if num < end:
                hwp.hwp.Run("TableLowerCell")
        
        logger.info(f"테이블 열({column})에 숫자 {start}~{end} 입력 완료")
        return f"테이블 열({column})에 숫자 {start}~{end} 입력 완료"
        
    except Exception as e:
        logger.error(f"테이블 숫자 채우기 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

if __name__ == "__main__":
    logger.info("Starting HWP MCP stdio server")
    try:
        # Run the FastMCP server with stdio transport
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Error running server: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1) 
