#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Tests for the current HWP controller behavior."""

from unittest.mock import MagicMock, patch

from src.tools.hwp_controller import HwpController


def _build_mock_hwp():
    mock_hwp = MagicMock()
    mock_hwp.XHwpWindows.Count = 1
    mock_hwp.XHwpWindows.Item.return_value = MagicMock()
    mock_hwp.XHwpDocuments.Item.return_value = MagicMock(Path="")
    mock_hwp.CurDocIndex = 0
    mock_hwp.Path = ""
    return mock_hwp


class TestHwpController:
    """Focused tests for controller connection and basic document operations."""

    @patch("src.tools.hwp_controller.os.path.exists", return_value=True)
    @patch("src.tools.hwp_controller.pythoncom.CoInitialize")
    @patch("src.tools.hwp_controller.win32com.client.GetActiveObject")
    def test_connect_uses_existing_instance(self, mock_get_active, mock_coinit, mock_exists):
        mock_hwp = _build_mock_hwp()
        mock_get_active.return_value = mock_hwp

        controller = HwpController()

        assert controller.connect() is True
        mock_coinit.assert_called_once()
        mock_get_active.assert_called_once_with("HWPFrame.HwpObject")
        mock_hwp.RegisterModule.assert_called_once()
        assert controller.is_hwp_running is True
        assert controller.last_error is None

    @patch("src.tools.hwp_controller.os.path.exists", return_value=True)
    @patch("src.tools.hwp_controller.pythoncom.CoInitialize")
    @patch("src.tools.hwp_controller.win32com.client.Dispatch")
    @patch("src.tools.hwp_controller.win32com.client.GetActiveObject")
    def test_connect_dispatches_when_hwp_is_not_running(
        self,
        mock_get_active,
        mock_dispatch,
        mock_coinit,
        mock_exists,
    ):
        mock_get_active.side_effect = [Exception("unavailable")] * 3
        mock_dispatch.return_value = _build_mock_hwp()

        controller = HwpController()

        with patch.object(controller, "_list_visible_hwp_windows", return_value=[]):
            assert controller.connect() is True

        mock_coinit.assert_called_once()
        mock_dispatch.assert_called_once_with("HWPFrame.HwpObject")
        assert controller.is_hwp_running is True

    @patch("src.tools.hwp_controller.pythoncom.CoInitialize")
    @patch("src.tools.hwp_controller.win32com.client.Dispatch")
    @patch("src.tools.hwp_controller.win32com.client.GetActiveObject")
    def test_connect_fails_when_direct_dispatch_is_disabled(
        self,
        mock_get_active,
        mock_dispatch,
        mock_coinit,
    ):
        mock_get_active.side_effect = [Exception("unavailable")] * 3

        controller = HwpController()

        with patch.object(controller, "_list_visible_hwp_windows", return_value=[]):
            assert controller.connect(allow_direct_dispatch=False) is False

        mock_coinit.assert_called_once()
        mock_dispatch.assert_not_called()
        assert "자동 시작이 필요합니다" in controller.last_error
        assert controller.is_hwp_running is False

    @patch("src.tools.hwp_controller.pythoncom.CoInitialize")
    @patch("src.tools.hwp_controller.win32com.client.Dispatch")
    @patch("src.tools.hwp_controller.win32com.client.GetActiveObject")
    def test_connect_fails_safely_when_hwp_window_exists_but_com_is_unavailable(
        self,
        mock_get_active,
        mock_dispatch,
        mock_coinit,
    ):
        mock_get_active.side_effect = [Exception("busy")] * 3

        controller = HwpController()

        with patch.object(
            controller,
            "_list_visible_hwp_windows",
            return_value=[{"hwnd": 100, "title": "example.hwp"}],
        ):
            assert controller.connect() is False

        mock_coinit.assert_called_once()
        mock_dispatch.assert_not_called()
        assert "COM 연결에 실패했습니다" in controller.last_error
        assert controller.is_hwp_running is False

    @patch("src.tools.hwp_controller.os.path.exists", return_value=False)
    def test_open_document_rejects_missing_path(self, mock_exists):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()

        assert controller.open_document("missing.hwp") is False
        assert "문서 파일을 찾을 수 없습니다" in controller.last_error
        controller.hwp.HAction.Execute.assert_not_called()

    @patch("src.tools.hwp_controller.os.makedirs")
    def test_save_document_creates_parent_directory(self, mock_makedirs):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()

        assert controller.save_document(r"nested\folder\document.hwp") is True
        mock_makedirs.assert_called_once()
        controller.hwp.SaveAs.assert_called_once()
        assert controller.current_document_path.endswith(r"nested\folder\document.hwp")

    def test_save_document_without_path_uses_active_document_path(self):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()
        controller.hwp.Path = r"C:\docs\current.hwp"

        assert controller.save_document() is True
        controller.hwp.Save.assert_called_once()
        assert controller.current_document_path == r"C:\docs\current.hwp"

    def test_insert_text_preserves_linebreaks(self):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()

        with patch.object(controller, "_insert_text_direct", return_value=True) as mock_insert:
            with patch.object(controller, "insert_paragraph", return_value=True) as mock_paragraph:
                assert controller.insert_text("Hello\nWorld") is True

        assert mock_insert.call_count == 2
        mock_insert.assert_any_call("Hello")
        mock_insert.assert_any_call("World")
        mock_paragraph.assert_called_once()

    def test_get_text_repairs_mojibake_korean_output(self):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()
        controller.hwp.GetTextFile.return_value = "¼­½Ä1 Âü°¡½ÅÃ»¼­"

        assert controller.get_text() == "서식1 참가신청서"
        assert controller.last_error is None

    def test_get_text_keeps_plain_ascii_output(self):
        controller = HwpController()
        controller.is_hwp_running = True
        controller.hwp = _build_mock_hwp()
        controller.hwp.GetTextFile.return_value = "plain ascii text"

        assert controller.get_text() == "plain ascii text"
