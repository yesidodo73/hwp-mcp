#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Tests for the HWP MCP server's helper-backed proxy behavior."""

from unittest.mock import patch

import hwp_mcp_stdio_server as server


class FakeHwpObject:
    def Run(self, command):
        return f"Run:{command}"


class FakeController:
    def __init__(self):
        self.hwp = FakeHwpObject()

    def get_text(self):
        return "controller-text"


class FakeTableTools:
    def describe(self):
        return "table-tools"


class FakeHelperClient:
    def __init__(self, include_roots=True):
        self._next_id = 1
        self.last_error = "연결 실패"
        self.objects = {}
        if include_roots:
            self.objects["controller"] = FakeController()
            self.objects["table_tools"] = FakeTableTools()

    def _register(self, value):
        object_id = f"obj-{self._next_id}"
        self._next_id += 1
        self.objects[object_id] = value
        return object_id

    def request(self, command, timeout=120, **payload):
        if command == "ensure_root":
            root = payload["root"]
            return {"ok": True, "exists": root in self.objects, "last_error": self.last_error}

        if command == "object_exists":
            object_id = payload["object_id"]
            return {"ok": True, "exists": object_id in self.objects, "last_error": self.last_error}

        if command == "inspect_attr":
            target = self.objects[payload["object_id"]]
            attr = getattr(target, payload["name"])
            if callable(attr):
                return {"ok": True, "kind": "callable", "last_error": self.last_error}
            if isinstance(attr, (str, int, float, bool, type(None))):
                return {"ok": True, "kind": "value", "value": attr, "last_error": self.last_error}
            return {
                "ok": True,
                "kind": "object",
                "object_id": self._register(attr),
                "last_error": self.last_error,
            }

        if command == "call_method":
            target = self.objects[payload["object_id"]]
            result = getattr(target, payload["name"])(*payload.get("args", []), **payload.get("kwargs", {}))
            if isinstance(result, (str, int, float, bool, type(None))):
                return {"ok": True, "kind": "value", "value": result, "last_error": self.last_error}
            return {
                "ok": True,
                "kind": "object",
                "object_id": self._register(result),
                "last_error": self.last_error,
            }

        if command == "clear_state":
            self.objects = {}
            return {"ok": True, "cleared": True, "last_error": None}

        raise AssertionError(f"Unexpected command: {command}")

    def clear_state(self):
        self.objects = {}

    def last_error_message(self):
        return self.last_error

    def shutdown(self):
        pass


def test_remote_proxy_routes_nested_calls_through_helper_client():
    helper = FakeHelperClient()

    with patch.object(server, "hwp_worker", helper):
        controller = server.get_hwp_controller()
        table_tools = server.get_hwp_table_tools()

        assert controller.get_text() == "controller-text"
        assert controller.hwp.Run("Ping") == "Run:Ping"
        assert table_tools.describe() == "table-tools"


def test_connection_error_message_uses_helper_last_error():
    helper = FakeHelperClient(include_roots=False)

    with patch.object(server, "hwp_worker", helper):
        controller = server.get_hwp_controller()

        assert not controller
        assert server._connection_error_message() == "연결 실패"


def test_root_proxy_uses_extended_timeout_for_initial_connection():
    helper = FakeHelperClient()
    recorded = []
    original_request = helper.request

    def _request(command, timeout=120, **payload):
        recorded.append((command, timeout, payload))
        return original_request(command, timeout=timeout, **payload)

    helper.request = _request

    with patch.object(server, "hwp_worker", helper):
        controller = server.get_hwp_controller()

        assert bool(controller) is True

    assert recorded[0][0] == "ensure_root"
    assert recorded[0][1] == server.ROOT_ENSURE_TIMEOUT_SECONDS


def test_worker_retries_when_helper_reports_connection_failure():
    worker = server.HwpComWorker()
    first_response = {
        "ok": True,
        "exists": False,
        "last_error": "한글 프로그램 시작/연결 실패: unavailable",
    }
    second_response = {
        "ok": True,
        "exists": True,
        "last_error": None,
    }

    with patch.object(worker, "_perform_request_locked", side_effect=[first_response, second_response]) as mock_request:
        with patch.object(worker, "_cleanup_hwp_processes") as mock_cleanup:
            with patch.object(worker, "shutdown") as mock_shutdown:
                with patch("hwp_mcp_stdio_server.time.sleep") as mock_sleep:
                    response = worker.request("ensure_root", root="controller")

    assert response["exists"] is True
    assert mock_request.call_count == 2
    mock_cleanup.assert_called_once()
    mock_shutdown.assert_called_once()
    mock_sleep.assert_called_once()


def test_worker_cleans_up_and_sets_cooldown_after_repeated_timeout():
    worker = server.HwpComWorker()

    with patch.object(worker, "_perform_request_locked", side_effect=[TimeoutError("x"), TimeoutError("y")]):
        with patch.object(worker, "_cleanup_hwp_processes") as mock_cleanup:
            with patch.object(worker, "shutdown") as mock_shutdown:
                with patch("hwp_mcp_stdio_server.time.sleep") as mock_sleep:
                    try:
                        worker.request("ensure_root", timeout=30, root="controller")
                    except TimeoutError:
                        pass
                    else:
                        raise AssertionError("TimeoutError expected")

    mock_cleanup.assert_called_once()
    mock_shutdown.assert_called_once()
    mock_sleep.assert_called_once()
    assert "제한 시간" in worker.last_error_message()


def test_worker_sets_cooldown_without_retry_after_long_root_timeout():
    worker = server.HwpComWorker()

    with patch.object(worker, "_perform_request_locked", side_effect=TimeoutError("slow")) as mock_request:
        with patch.object(worker, "_cleanup_hwp_processes") as mock_cleanup:
            with patch.object(worker, "shutdown") as mock_shutdown:
                try:
                    worker.request(
                        "ensure_root",
                        timeout=server.ROOT_ENSURE_TIMEOUT_SECONDS,
                        root="controller",
                    )
                except TimeoutError:
                    pass
                else:
                    raise AssertionError("TimeoutError expected")

    mock_request.assert_called_once()
    mock_cleanup.assert_called_once()
    mock_shutdown.assert_called_once()
    assert "제한 시간" in worker.last_error_message()


def test_worker_returns_fast_during_startup_cooldown():
    worker = server.HwpComWorker()
    worker._set_startup_cooldown("recent startup failure", seconds=60)

    with patch.object(worker, "_start_process") as mock_start:
        response = worker._perform_request_locked("ensure_root", 30, {"root": "controller"})

    assert response["ok"] is True
    assert response["exists"] is False
    assert response["last_error"] == "recent startup failure"
    mock_start.assert_not_called()


def test_worker_requires_external_broker_when_running_in_job():
    worker = server.HwpComWorker()

    def _missing_broker():
        worker._set_last_hwp_error(worker._broker_unavailable_message())
        return False

    with patch.object(worker, "_running_in_job", return_value=True):
        with patch.object(worker, "_connect_to_external_broker", side_effect=_missing_broker):
            try:
                worker.request("ensure_root", timeout=1, root="controller")
            except RuntimeError as exc:
                assert "broker" in str(exc)
            else:
                raise AssertionError("RuntimeError expected")


def test_worker_prefers_external_broker_when_running_in_job():
    worker = server.HwpComWorker()

    with patch.object(worker, "_running_in_job", return_value=True):
        with patch.object(worker, "_connect_to_external_broker", return_value=True) as mock_connect:
            with patch("hwp_mcp_stdio_server.subprocess.Popen") as mock_popen:
                worker._start_process()

    mock_connect.assert_called_once()
    mock_popen.assert_not_called()
