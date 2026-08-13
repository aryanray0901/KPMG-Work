#!/usr/bin/env python3
"""Exercise every non-layout editor control through its browser route."""

from __future__ import annotations

import io
import os
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import app as deck_app
from builder_ops import create_blank_deck


SOURCE = ROOT / "sample_files" / "kpmg_advisory_30_slide_original.pptx"


def fake_render(sess_dir: str, version: int) -> dict:
    count = len(Presentation(deck_app._editor_version_path(sess_dir, version)).slides)
    return {
        "slide_count": count,
        "rendering_ok": False,
        "render_engine": None,
        "render_help": "Test renderer disabled.",
        "preview_revision": time.time_ns(),
    }


def assert_ok(response, label: str) -> dict:
    payload = response.get_json(silent=True) or {}
    assert response.status_code == 200 and payload.get("ok"), (label, response.status_code, payload)
    assert payload.get("message"), (label, payload)
    return payload


def verify_button_hooks() -> int:
    template = (ROOT / "templates" / "editor.html").read_text(encoding="utf-8")
    script = (ROOT / "static" / "editor.js").read_text(encoding="utf-8")
    buttons = re.findall(r"<button\b[^>]*>", template, flags=re.I)
    known_ids = {
        "new-slide-button", "previous-slide", "next-slide", "inspector-toggle",
        "apply-theme-slide", "apply-theme-deck", "send-chat", "builder-close",
        "builder-cancel", "guided-close", "guided-cancel", "undo-btn", "redo-btn",
    }
    for tag in buttons:
        hooked = any(token in tag for token in (
            "data-action=", "data-prompt=", "data-editor-form=", "data-chart-type=", "data-right-tab=",
        ))
        hooked = hooked or bool(re.search(r'type="submit"', tag))
        identifier = re.search(r'id="([^"]+)"', tag)
        if identifier and identifier.group(1) in known_ids:
            hooked = True
            assert identifier.group(1) in script, identifier.group(1)
        assert hooked, f"Button has no handler: {tag}"
    return len(buttons)


def verify_routes() -> None:
    original_renderer = deck_app._render_editor_version
    deck_app._render_editor_version = fake_render
    deck_app.app.testing = True
    client = deck_app.app.test_client()
    session_ids: list[str] = []
    try:
        sid = deck_app._create_editor_session(str(SOURCE), SOURCE.name)
        session_ids.append(sid)

        editor_page = client.get(f"/editor/{sid}")
        assert editor_page.status_code == 200
        rendered_html = editor_page.get_data(as_text=True)
        assert "<summary>Chart editor</summary>" in rendered_html
        assert "<summary>Add chart</summary>" not in rendered_html
        assert rendered_html.count("data-chart-type=") == 7, "The HUD must expose seven native chart types"
        assert "data-new-slide-layout=" not in rendered_html and 'class="layout-choice"' not in rendered_html
        assert "data-editor-form=" not in rendered_html
        assert "Change company name" not in rendered_html and "rebrand_company" not in rendered_html
        assert "<summary>Slide actions</summary>" not in rendered_html, "Slide actions must not be repeated in the side panel"
        assert rendered_html.count("data-right-tab=") == 2, "Design and AI must be separate tabs"
        assert 'data-right-pane="design"' in rendered_html and 'data-right-pane="ai"' in rendered_html
        assert "apply-theme-slide" in rendered_html and "apply-theme-deck" in rendered_html and "theme-primary" in rendered_html
        assert "builder-chart-type" in rendered_html and "Blank chart" in rendered_html and "Use Excel data" in rendered_html
        assert "<summary>Chart tools</summary>" not in rendered_html
        assert all(f"Blank {name}" in rendered_html for name in ("Column", "Bar", "Line", "Pie", "Area", "Waterfall", "Scatter"))
        assert 'name="placement" value="new"' in rendered_html and 'name="placement" value="current"' in rendered_html
        assert "Convert the chart on this slide" not in rendered_html and "Replace the chart on this slide" not in rendered_html
        assert rendered_html.count('name="data_mode"') == 2
        assert "<summary>Tables</summary>" not in rendered_html and "Smart Insert" not in rendered_html
        assert "McKinsey" not in rendered_html and "Bain" not in rendered_html and "Change font" not in rendered_html
        assert "activateRightTab" in (ROOT / "static" / "editor.js").read_text(encoding="utf-8")

        payload = assert_ok(client.post(f"/editor/action/{sid}", json={"action": "duplicate_slide", "selected_slide": 30}), "duplicate")
        assert payload["state"]["slide_count"] == 31 and payload["state"]["selected_slide"] == 31
        payload = assert_ok(client.post(f"/editor/action/{sid}", json={"action": "move_left", "selected_slide": 31}), "move left")
        assert payload["state"]["selected_slide"] == 30
        payload = assert_ok(client.post(f"/editor/action/{sid}", json={"action": "move_right", "selected_slide": 30}), "move right")
        assert payload["state"]["selected_slide"] == 31
        assert_ok(client.post(f"/editor/action/{sid}", json={"action": "undo", "selected_slide": 31}), "undo")
        assert_ok(client.post(f"/editor/action/{sid}", json={"action": "redo", "selected_slide": 31}), "redo")
        payload = assert_ok(client.post(f"/editor/action/{sid}", json={"action": "delete_slide", "selected_slide": 31}), "delete")
        assert payload["state"]["slide_count"] == 30
        assert_ok(client.post(f"/editor/action/{sid}", json={"action": "retry_preview", "selected_slide": 30}), "retry preview")

        assert_ok(client.post(f"/editor/quick/{sid}", json={
            "command": "Rewrite this slide for a senior executive audience.", "selected_slide": 30,
        }), "quick edit")
        assert_ok(client.post(
            f"/editor/chat/{sid}",
            data={"message": "Add an editable native chart to this slide using data already in the deck.", "selected_slide": "30"},
            content_type="multipart/form-data",
        ), "chat edit")

        wizard_response = client.post(f"/editor/wizard/{sid}", json={
            "message": "I need to compare Q2 and Q3 sales.",
        })
        wizard = wizard_response.get_json(silent=True) or {}
        assert wizard_response.status_code == 200 and wizard.get("ok"), wizard
        assert wizard["recommendation"]["layout"] == "comparison"

        payload = assert_ok(client.post(f"/editor/build/{sid}", data={
            "layout": "bar_chart", "chart_type": "column", "data_mode": "blank",
            "selected_slide": "30", "placement": "new", "smart": "false",
        }), "blank chart builder")
        assert payload["state"]["slide_count"] == 31

        payload = assert_ok(client.post(f"/editor/build/{sid}", data={
            "layout": "bar_chart", "chart_type": "bar", "data_mode": "blank",
            "selected_slide": "1", "placement": "current", "smart": "false",
        }), "blank chart on current slide")
        assert payload["state"]["slide_count"] == 31 and payload["state"]["selected_slide"] == 1
        assert "slide 1" in payload["message"] and "blank" in payload["message"]

        csv_bytes = b"Month,Revenue,Margin\nJan,100,20%\nFeb,112,22%\nMar,125,24%\n"
        payload = assert_ok(client.post(
            f"/editor/build/{sid}",
            data={
                "layout": "line_chart", "chart_type": "line", "data_mode": "upload",
                "selected_slide": "31", "placement": "new", "smart": "false",
                "data_file": (io.BytesIO(csv_bytes), "monthly_revenue.csv"),
            },
            content_type="multipart/form-data",
        ), "Excel chart builder")
        assert payload["state"]["slide_count"] == 32 and payload.get("profile")

        download = client.get(f"/editor/download/{sid}")
        assert download.status_code == 200 and download.data.startswith(b"PK"), "download"

        handle, blank = tempfile.mkstemp(suffix=".pptx")
        os.close(handle)
        try:
            create_blank_deck(blank)
            single_sid = deck_app._create_editor_session(blank, "Untitled presentation.pptx")
            session_ids.append(single_sid)
            payload = assert_ok(client.post(f"/editor/action/{single_sid}", json={
                "action": "delete_slide", "selected_slide": 1,
            }), "single-slide delete")
            assert payload["state"]["slide_count"] == 1
        finally:
            os.remove(blank)
    finally:
        deck_app._render_editor_version = original_renderer
        for sid in session_ids:
            shutil.rmtree(os.path.join(deck_app.SESSIONS_DIR, sid), ignore_errors=True)


def main() -> None:
    button_count = verify_button_hooks()
    verify_routes()
    print(f"PASS: {button_count} button hooks and all toolbar, chart builder, theme, chat, and download routes.")


if __name__ == "__main__":
    main()
