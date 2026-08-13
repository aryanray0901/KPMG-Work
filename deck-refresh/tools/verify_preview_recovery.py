"""Regression test for preview refresh and recovery after a failed render."""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# Keep this test runnable before the launcher's first dependency install.
try:
    import openai  # noqa: F401
except ImportError:
    module = types.ModuleType("openai")
    module.OpenAI = type("OpenAI", (), {"__init__": lambda self, *args, **kwargs: None})
    sys.modules["openai"] = module

try:
    import flask  # noqa: F401
except ImportError:
    module = types.ModuleType("flask")

    class DummyFlask:
        def __init__(self, *args, **kwargs):
            self.config = {}
            self.secret_key = None

        def route(self, *args, **kwargs):
            return lambda function: function

        def after_request(self, *args, **kwargs):
            return lambda function: function

    module.Flask = DummyFlask
    module.request = types.SimpleNamespace()
    for name in ("render_template", "send_file", "redirect", "url_for", "flash", "abort", "jsonify"):
        setattr(module, name, lambda *args, **kwargs: None)
    sys.modules["flask"] = module

try:
    import werkzeug.utils  # noqa: F401
except ImportError:
    werkzeug = types.ModuleType("werkzeug")
    utils = types.ModuleType("werkzeug.utils")
    utils.secure_filename = lambda value: value
    sys.modules["werkzeug"] = werkzeug
    sys.modules["werkzeug.utils"] = utils

import app
from builtin_preview import render_pptx


def hashes(sess_dir, state):
    render_dir = app._editor_render_dir(sess_dir, state["version"])
    result = []
    for number in range(1, state["slide_count"] + 1):
        path = os.path.join(render_dir, f"slide_{number}.png")
        assert os.path.exists(path), path
        result.append(hashlib.sha256(Path(path).read_bytes()).hexdigest())
    return result


def main():
    # Force the built-in renderer to prove recovery works without Office.
    app.render_pptx_to_images = lambda pptx_path, out_dir, prefix: (
        render_pptx(pptx_path, out_dir, prefix),
        "Built-in PowerPoint preview",
    )
    source = ROOT / "sample_files" / "kpmg_advisory_q3_original.pptx"
    sid = app._create_editor_session(str(source), "sample.pptx")
    sess_dir = app._session_dir(sid)
    try:
        initial = app._editor_state(sess_dir, 1)
        assert initial["rendering_ok"] and initial["slide_count"] == 9
        initial_hashes = hashes(sess_dir, initial)

        app._commit_editor_operations(sess_dir, [{"op": "delete_slide", "slide": 9}])
        deleted = app._editor_state(sess_dir, 8)
        assert deleted["rendering_ok"] and deleted["slide_count"] == 8
        assert hashes(sess_dir, deleted) == initial_hashes[:8]

        app._commit_editor_operations(
            sess_dir,
            [{"op": "duplicate_slide", "slide": 4, "position": 5}],
        )
        duplicated = app._editor_state(sess_dir, 5)
        assert duplicated["rendering_ok"] and duplicated["slide_count"] == 9
        duplicate_hashes = hashes(sess_dir, duplicated)

        # Simulate the exact reported failure: the preview directory and state
        # are broken, then a later valid command succeeds.
        shutil.rmtree(app._editor_render_dir(sess_dir, duplicated["version"]), ignore_errors=True)
        app._save_session_meta(
            sess_dir,
            rendering_ok=False,
            render_engine=None,
            render_help="simulated failed preview",
        )
        assert not app._editor_state(sess_dir, 5)["rendering_ok"]

        app._commit_editor_operations(
            sess_dir,
            [{"op": "move_slide", "from_slide": 5, "to_slide": 9}],
        )
        recovered = app._editor_state(sess_dir, 9)
        recovered_hashes = hashes(sess_dir, recovered)
        assert recovered["rendering_ok"]
        assert recovered["slide_count"] == 9
        assert recovered["preview_revision"] > duplicated["preview_revision"]
        assert recovered_hashes[-1] == duplicate_hashes[4]
        print("PASS: a successful edit regenerated the preview after a prior failure")
    finally:
        shutil.rmtree(sess_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
