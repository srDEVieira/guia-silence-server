from __future__ import annotations

import atexit
from pathlib import Path
import threading

import webview


def _cleanup_pythonnet_shutdown() -> None:
    try:
        import pythonnet
    except Exception:
        return

    try:
        atexit.unregister(pythonnet.unload)
    except Exception:
        pass

    try:
        pythonnet.unload()
    except KeyboardInterrupt:
        pass
    except Exception:
        pass


class _LazyGuideAppApi:
    def __init__(self, base_dir: Path) -> None:
        self._base_dir = base_dir
        self._api = None
        self._lock = threading.Lock()

    def _get_api(self):
        with self._lock:
            if self._api is None:
                from src.web_backend import GuideAppApi

                self._api = GuideAppApi(base_dir=self._base_dir)
            return self._api

    def get_initial_state(self):
        return self._get_api().get_initial_state()

    def refresh_printers(self):
        return self._get_api().refresh_printers()

    def lookup_item(self, patrimony):
        return self._get_api().lookup_item(patrimony)

    def print_guides(self, payload):
        return self._get_api().print_guides(payload)


def run_webview_app() -> None:
    base_dir = Path.cwd()
    api = _LazyGuideAppApi(base_dir=base_dir)
    index_path = (base_dir / "web" / "index.html").resolve()

    webview.create_window(
        title="Central de Guias Patrimoniais",
        url=index_path.as_uri(),
        js_api=api,
        width=1460,
        height=920,
        min_size=(1220, 760),
        background_color="#e9eef4",
    )
    try:
        webview.start(debug=False, private_mode=False, http_server=True)
    finally:
        _cleanup_pythonnet_shutdown()
