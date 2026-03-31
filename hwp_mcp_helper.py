#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import logging
import os
import sys
import traceback


current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

from src.tools.hwp_controller import HwpController
from src.tools.hwp_table_tools import HwpTableTools


log_path = os.path.join(current_dir, "hwp_mcp_helper.log")
logging.basicConfig(
    level=logging.INFO,
    filename=log_path,
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("hwp-mcp-helper")

try:
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

_SIMPLE_RETURN_TYPES = (str, int, float, bool, type(None))


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


class HelperState:
    def __init__(self):
        self.controller = None
        self.table_tools = None
        self.last_error = None
        self.objects = {}
        self._next_object_id = 1

    def _set_last_error(self, message):
        self.last_error = message

    def _reset_object_store(self):
        self.objects = {}
        if self.controller is not None:
            self.objects["controller"] = self.controller
        if self.table_tools is not None:
            self.objects["table_tools"] = self.table_tools

    def clear_state(self):
        self.controller = None
        self.table_tools = None
        self._set_last_error(None)
        self._reset_object_store()

    def _connection_alive(self) -> bool:
        if self.controller is None:
            return False

        try:
            _ = self.controller.hwp.XHwpWindows.Count
            return True
        except Exception as exc:
            logger.warning(f"HWP connection lost ({exc}), resetting helper state")
            self.clear_state()
            return False

    def ensure_controller(self):
        if self.controller is not None and not self._connection_alive():
            self.controller = None
            self.table_tools = None

        if self.controller is None:
            logger.info("Creating HwpController instance inside helper")
            controller = HwpController()
            if not controller.connect(visible=True):
                message = controller.last_error or "Failed to connect to HWP program"
                self._set_last_error(message)
                logger.error(f"Failed to connect to HWP program: {message}")
                return None

            self.controller = controller
            self.table_tools = HwpTableTools(controller)
            self._set_last_error(None)
            self._reset_object_store()
            logger.info("Successfully connected to HWP program")

        return self.controller

    def ensure_table_tools(self):
        controller = self.ensure_controller()
        if controller is None:
            return None

        if self.table_tools is None:
            self.table_tools = HwpTableTools(controller)
            self._reset_object_store()

        return self.table_tools

    def register_object(self, value):
        object_id = f"obj-{self._next_object_id}"
        self._next_object_id += 1
        self.objects[object_id] = value
        return object_id

    def resolve_object(self, object_id):
        if object_id == "controller":
            return self.ensure_controller()
        if object_id == "table_tools":
            return self.ensure_table_tools()
        return self.objects.get(object_id)


state = HelperState()


def _success(request_id, **payload):
    response = {"id": request_id, "ok": True}
    response.update(payload)
    response["last_error"] = state.last_error
    return response


def _failure(request_id, message):
    return {
        "id": request_id,
        "ok": False,
        "error": message,
        "last_error": state.last_error,
    }


def handle_request(request):
    request_id = request.get("id")
    command = request.get("command")

    if command == "ping":
        return _success(request_id, result="pong")

    if command == "clear_state":
        state.clear_state()
        return _success(request_id, cleared=True)

    if command == "ensure_root":
        root = request.get("root")
        if root == "controller":
            exists = state.ensure_controller() is not None
        elif root == "table_tools":
            exists = state.ensure_table_tools() is not None
        else:
            return _failure(request_id, f"Unknown root object: {root}")

        return _success(request_id, exists=exists)

    if command == "object_exists":
        object_id = request.get("object_id")
        exists = state.resolve_object(object_id) is not None
        return _success(request_id, exists=exists)

    if command == "inspect_attr":
        object_id = request.get("object_id")
        name = request.get("name")
        target = state.resolve_object(object_id)
        if target is None:
            return _failure(request_id, f"Remote object is not available: {object_id}")

        attr = getattr(target, name)
        if callable(attr):
            return _success(request_id, kind="callable")
        if _is_marshaled_value(attr):
            return _success(request_id, kind="value", value=attr)

        return _success(
            request_id,
            kind="object",
            object_id=state.register_object(attr),
        )

    if command == "call_method":
        object_id = request.get("object_id")
        name = request.get("name")
        args = request.get("args", [])
        kwargs = request.get("kwargs", {})
        target = state.resolve_object(object_id)
        if target is None:
            return _failure(request_id, f"Remote object is not available: {object_id}")

        result = getattr(target, name)(*args, **kwargs)
        if _is_marshaled_value(result):
            return _success(request_id, kind="value", value=result)

        return _success(
            request_id,
            kind="object",
            object_id=state.register_object(result),
        )

    return _failure(request_id, f"Unknown command: {command}")


def main():
    logger.info("Starting HWP MCP helper process")

    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue

        try:
            request = json.loads(line)
            response = handle_request(request)
        except Exception as exc:
            logger.error(f"Helper request failed: {exc}", exc_info=True)
            request_id = None
            try:
                request_id = request.get("id")  # type: ignore[name-defined]
            except Exception:
                pass
            response = _failure(
                request_id,
                f"{exc}\n{traceback.format_exc()}",
            )

        try:
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()
        except OSError:
            break


if __name__ == "__main__":
    main()
