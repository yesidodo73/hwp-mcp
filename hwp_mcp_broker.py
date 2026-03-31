#!/usr/bin/env python
# -*- coding: utf-8 -*-

import atexit
import json
import os
import socketserver
import sys
import tempfile


current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

import hwp_mcp_helper as helper


def _state_path() -> str:
    temp_root = os.path.join(tempfile.gettempdir(), "hwp-mcp")
    os.makedirs(temp_root, exist_ok=True)
    return os.path.join(temp_root, "broker_state.json")


def _write_state(port: int) -> None:
    with open(_state_path(), "w", encoding="utf-8") as fp:
        json.dump(
            {
                "pid": os.getpid(),
                "port": port,
            },
            fp,
        )


def _remove_state() -> None:
    state_path = _state_path()
    if not os.path.exists(state_path):
        return

    try:
        with open(state_path, "r", encoding="utf-8") as fp:
            state = json.load(fp)
    except Exception:
        state = {}

    if state.get("pid") == os.getpid():
        try:
            os.remove(state_path)
        except OSError:
            pass


class ThreadedBrokerServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


class BrokerRequestHandler(socketserver.StreamRequestHandler):
    def handle(self):
        for raw_line in self.rfile:
            line = raw_line.decode("utf-8", errors="replace").strip()
            if not line:
                continue

            request = {}
            try:
                request = json.loads(line)
                response = helper.handle_request(request)
            except Exception as exc:
                helper.logger.error(f"Broker request failed: {exc}", exc_info=True)
                response = helper._failure(request.get("id"), str(exc))

            payload = json.dumps(response, ensure_ascii=False) + "\n"
            self.wfile.write(payload.encode("utf-8"))
            self.wfile.flush()


def main():
    helper.logger.info("Starting HWP MCP broker process")
    server = ThreadedBrokerServer(("127.0.0.1", 0), BrokerRequestHandler)
    _write_state(server.server_address[1])
    atexit.register(_remove_state)

    try:
        server.serve_forever(poll_interval=0.5)
    finally:
        _remove_state()
        server.server_close()


if __name__ == "__main__":
    main()
