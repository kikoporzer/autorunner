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
        self._record_typing = False
        self._record_hotkeys = False
        self._typing_buffer = ""
        self._active_modifiers: set[str] = set()
        self._hotkeys_in_progress: set[str] = set()

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

    def start(
        self,
        *,
        stop_key_name: str = "f8",
        on_finished: Callable[[list[dict], str], None] | None = None,
        record_typing: bool = False,
        record_hotkeys: bool = False,
    ) -> RecorderStatus:
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
        self._record_typing = bool(record_typing)
        self._record_hotkeys = bool(record_hotkeys)
        self._typing_buffer = ""
        self._active_modifiers = set()
        self._hotkeys_in_progress = set()

        self._last_event_time = time.time()
        self._last_left_click_time = 0.0
        self._last_left_click_xy = None
        self._last_left_click_index = None

        click_filter = getattr(self, "_click_filter", None)

        def on_click(x, y, button, pressed):
            if not self.is_running or pressed:
                return
            self._flush_typing_buffer()
            if callable(click_filter):
                try:
                    if not click_filter(int(x), int(y), str(button)):
                        return
                except Exception:
                    pass

            now = time.time()
            self._append_wait_since_last_event(now)

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
                self._flush_typing_buffer()
                self.stop()
                return False
            if not self.is_running:
                return True

            now = time.time()
            key_name = self._key_to_name(key)
            if key_name in {"ctrl", "alt", "shift", "cmd"}:
                self._active_modifiers.add(key_name)
                self._flush_typing_buffer()
                return True

            hotkey_recorded = False
            if self._record_hotkeys and self._active_modifiers:
                keys = sorted(self._active_modifiers) + [key_name]
                signature = "+".join(keys)
                if signature not in self._hotkeys_in_progress:
                    self._append_wait_since_last_event(now)
                    self._flush_typing_buffer()
                    self._append_step({"type": "hotkey", "keys": keys, "enabled": True})
                    self._last_event_time = now
                    self._hotkeys_in_progress.add(signature)
                hotkey_recorded = True

            if hotkey_recorded:
                return True

            if self._record_typing:
                char = self._key_to_char(key)
                if char:
                    self._append_wait_since_last_event(now)
                    self._typing_buffer += char
                    self._last_event_time = now
                    return True

            special = self._special_press_key_name(key)
            if special:
                self._append_wait_since_last_event(now)
                self._flush_typing_buffer()
                self._append_step({"type": "press_key", "key": special, "enabled": True})
                self._last_event_time = now
            return True

        def on_release(key):
            key_name = self._key_to_name(key)
            if key_name in self._active_modifiers:
                self._active_modifiers.discard(key_name)
            if not self._active_modifiers:
                self._hotkeys_in_progress.clear()
            return True

        def worker():
            try:
                self._mouse_listener = mouse.Listener(on_click=on_click)
                self._keyboard_listener = keyboard.Listener(on_press=on_press, on_release=on_release)

                self._mouse_listener.start()
                self._keyboard_listener.start()

                while self.is_running and not self._stop_event.is_set():
                    time.sleep(0.03)

            except Exception as exc:
                self.last_error = str(exc)
                self._stop_event.set()
            finally:
                self._flush_typing_buffer()
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

    def _append_wait_since_last_event(self, now: float) -> None:
        elapsed = now - self._last_event_time
        if elapsed >= 0.05:
            self._append_step({"type": "wait", "seconds": round(elapsed, 3), "enabled": True})

    def _flush_typing_buffer(self) -> None:
        if not self._typing_buffer:
            return
        self._append_step({"type": "type_text", "value": self._typing_buffer, "enabled": True})
        self._typing_buffer = ""

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

    @staticmethod
    def _key_to_name(key) -> str:
        key_char = getattr(key, "char", None)
        if isinstance(key_char, str) and key_char:
            return key_char.lower()
        key_name = getattr(key, "name", "")
        if key_name:
            return str(key_name).lower()
        text = str(key).lower()
        if "." in text:
            return text.split(".")[-1]
        return text

    @staticmethod
    def _key_to_char(key) -> str:
        key_char = getattr(key, "char", None)
        if isinstance(key_char, str) and key_char and key_char.isprintable():
            return key_char
        if str(getattr(key, "name", "")).lower() == "space":
            return " "
        return ""

    @staticmethod
    def _special_press_key_name(key) -> str:
        mapped = {
            "enter": "enter",
            "tab": "tab",
            "esc": "esc",
            "backspace": "backspace",
            "delete": "delete",
            "up": "up",
            "down": "down",
            "left": "left",
            "right": "right",
        }
        name = GlobalClickRecorder._key_to_name(key)
        return mapped.get(name, "")
