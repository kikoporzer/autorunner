import threading
import time
import subprocess
import sys
from dataclasses import dataclass
from typing import Callable


@dataclass
class RecorderStatus:
    available: bool
    message: str


class GlobalClickRecorder:
    """Cross-platform global click recorder using pynput when available."""

    def __init__(self):
        self._mouse_listener = None
        self._keyboard_listener = None
        self._thread = None
        self._stop_event = threading.Event()
        self._lock = threading.Lock()

        self.steps: list[dict] = []
        self.is_running = False
        self.last_error = ""

        self._last_event_time = 0.0
        self._last_left_click_time = 0.0
        self._last_left_click_xy: tuple[int, int] | None = None
        self._last_left_click_index: int | None = None

        self._stop_key_name = "f8"
        self._on_finished: Callable[[list[dict], str], None] | None = None
        self._click_filter: Callable[[int, int, str], bool] | None = None

    @staticmethod
    def availability() -> RecorderStatus:
        ok, detail = GlobalClickRecorder._safe_probe_pynput()
        if not ok:
            return RecorderStatus(
                available=False,
                message=(
                    "Cross-platform recorder requires pynput. "
                    f"Install/fix with: python3 -m pip install pynput ({detail})"
                ),
            )

        return RecorderStatus(available=True, message="Recorder backend is available.")

    def start(self, *, stop_key_name: str = "f8", on_finished: Callable[[list[dict], str], None] | None = None) -> RecorderStatus:
        if self.is_running:
            return RecorderStatus(False, "Recorder is already running.")

        ok, detail = self._safe_probe_pynput()
        if not ok:
            return RecorderStatus(
                available=False,
                message=f"Unable to start recorder. Install/repair pynput first. ({detail})",
            )

        try:
            from pynput import keyboard, mouse
        except Exception as exc:
            return RecorderStatus(
                available=False,
                message=f"Unable to start recorder. Install/repair pynput first. ({exc})",
            )

        self.steps = []
        self.last_error = ""
        self._stop_event.clear()
        self.is_running = True
        self._stop_key_name = stop_key_name.lower().strip()
        self._on_finished = on_finished

        self._last_event_time = time.time()
        self._last_left_click_time = 0.0
        self._last_left_click_xy = None
        self._last_left_click_index = None

        click_filter = getattr(self, "_click_filter", None)

        def on_click(x, y, button, pressed):
            if not self.is_running or pressed:
                return
            if callable(click_filter):
                try:
                    if not click_filter(int(x), int(y), str(button)):
                        return
                except Exception:
                    pass

            now = time.time()
            elapsed = now - self._last_event_time
            if elapsed >= 0.05:
                self._append_step({"type": "wait", "seconds": round(elapsed, 3), "enabled": True})

            button_name = str(button).lower()
            if "left" in button_name:
                self._record_left_click(now, int(x), int(y))
            elif "right" in button_name:
                self._append_step({"type": "right_click", "x": int(x), "y": int(y), "enabled": True})
            else:
                self._append_step({"type": "click_xy", "x": int(x), "y": int(y), "enabled": True})

            self._last_event_time = now

        def on_press(key):
            if self._matches_stop_key(key):
                self.stop()
                return False
            return True

        def worker():
            try:
                self._mouse_listener = mouse.Listener(on_click=on_click)
                self._keyboard_listener = keyboard.Listener(on_press=on_press)

                self._mouse_listener.start()
                self._keyboard_listener.start()

                while self.is_running and not self._stop_event.is_set():
                    time.sleep(0.03)

            except Exception as exc:
                self.last_error = str(exc)
                self._stop_event.set()
            finally:
                self._close_listeners()
                self.is_running = False
                if self._on_finished:
                    self._on_finished(list(self.steps), self.last_error)

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()
        return RecorderStatus(True, "Recorder started.")

    def set_click_filter(self, callback: Callable[[int, int, str], bool] | None) -> None:
        self._click_filter = callback

    @staticmethod
    def _safe_probe_pynput() -> tuple[bool, str]:
        cmd = [
            sys.executable,
            "-c",
            "from pynput import keyboard, mouse; print('ok')",
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
        except Exception as exc:
            return False, str(exc)
        if proc.returncode == 0:
            return True, "ok"
        err = (proc.stderr or proc.stdout or "").strip()
        if not err:
            err = f"probe exited with code {proc.returncode}"
        return False, err

    def stop(self) -> None:
        self._stop_event.set()
        self.is_running = False
        self._close_listeners()

    def _close_listeners(self) -> None:
        for listener in (self._mouse_listener, self._keyboard_listener):
            if listener is not None:
                try:
                    listener.stop()
                except Exception:
                    pass
        self._mouse_listener = None
        self._keyboard_listener = None

    def _append_step(self, step: dict) -> None:
        with self._lock:
            self.steps.append(step)

    def _record_left_click(self, now: float, x: int, y: int) -> None:
        is_double = False
        if self._last_left_click_xy is not None and self._last_left_click_time > 0 and self._last_left_click_index is not None:
            dt_ms = (now - self._last_left_click_time) * 1000
            dist = ((x - self._last_left_click_xy[0]) ** 2 + (y - self._last_left_click_xy[1]) ** 2) ** 0.5
            if dt_ms <= 500 and dist <= 8:
                is_double = True

        if is_double and self._last_left_click_index is not None:
            with self._lock:
                if 0 <= self._last_left_click_index < len(self.steps):
                    self.steps[self._last_left_click_index] = {
                        "type": "double_click",
                        "x": x,
                        "y": y,
                        "enabled": True,
                    }
            self._last_left_click_index = None
            self._last_left_click_time = 0.0
            self._last_left_click_xy = None
            return

        self._append_step({"type": "click_xy", "x": x, "y": y, "enabled": True})
        self._last_left_click_index = len(self.steps) - 1
        self._last_left_click_time = now
        self._last_left_click_xy = (x, y)

    def _matches_stop_key(self, key) -> bool:
        # Supports F-keys and single-character keys.
        try:
            from pynput.keyboard import Key

            if self._stop_key_name.startswith("f") and self._stop_key_name[1:].isdigit():
                fn = getattr(Key, self._stop_key_name, None)
                return key == fn

            key_char = getattr(key, "char", None)
            if key_char:
                return str(key_char).lower() == self._stop_key_name

            key_name = getattr(key, "name", "")
            return str(key_name).lower() == self._stop_key_name

        except Exception:
            return False
