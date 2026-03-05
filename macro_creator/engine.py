from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .models import MacroDocument, Task


@dataclass
class RunState:
    current_loop: int = 0
    current_task_index: int = 0
    running: bool = False
    paused: bool = False


class MacroEngine:
    def __init__(self) -> None:
        self.state = RunState()
        self._thread: Optional[threading.Thread] = None
        self._pause_event = threading.Event()
        self._pause_event.set()
        self._stop_requested = False
        self._pyautogui = None

    def _lazy_pyautogui(self):
        if self._pyautogui is None:
            import pyautogui as p  # lazy import for headless/dev safety

            self._pyautogui = p
        return self._pyautogui

    def run_async(
        self,
        doc: MacroDocument,
        update: Callable[[str], None],
        start_loop: int = 0,
        start_task_index: int = 0,
    ) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_requested = False
        self._pause_event.set()
        self.state.running = True
        self.state.paused = False

        def worker():
            try:
                self._run(doc, update, start_loop, start_task_index)
            finally:
                self.state.running = False
                self.state.paused = False

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()

    def pause(self) -> None:
        self.state.paused = True
        self._pause_event.clear()

    def resume(self) -> None:
        self.state.paused = False
        self._pause_event.set()

    def stop(self) -> None:
        self._stop_requested = True
        self._pause_event.set()

    def _run(self, doc: MacroDocument, update: Callable[[str], None], start_loop: int, start_task_index: int) -> None:
        for loop_idx in range(start_loop, max(1, doc.loop_count)):
            self.state.current_loop = loop_idx
            first_task = start_task_index if loop_idx == start_loop else 0
            for task_idx in range(first_task, len(doc.tasks)):
                if self._stop_requested:
                    update("Stopped")
                    return
                self._pause_event.wait()
                self.state.current_task_index = task_idx
                task = doc.tasks[task_idx]
                one_off = str(task.params.get("one_off", "0")).lower() in {"1", "true", "yes", "on"}
                if one_off and loop_idx > 0:
                    continue
                update(f"Loop {loop_idx + 1}/{doc.loop_count} | Task {task_idx + 1}: {task.task_type}")
                self._execute_task(task, doc, loop_idx, task_idx)
                delay = max(0.0, float(getattr(doc, "inter_task_delay", 0.0)))
                if delay > 0:
                    end_delay = time.time() + delay
                    while time.time() < end_delay:
                        if self._stop_requested:
                            return
                        self._pause_event.wait()
                        time.sleep(0.05)
        update("Completed")

    def _execute_task(self, task: Task, doc: MacroDocument, loop_idx: int, task_idx: int) -> None:
        p = self._lazy_pyautogui()
        params = task.params
        if task.task_type == "Click":
            p.click(
                x=int(params.get("x", 0)),
                y=int(params.get("y", 0)),
                clicks=2 if params.get("click_type", "single").lower() == "double" else 1,
                button=params.get("button", "left").lower(),
            )
        elif task.task_type == "Move":
            p.moveTo(x=int(params.get("x", 0)), y=int(params.get("y", 0)))
        elif task.task_type == "String":
            text = self._resolve_value(str(params.get("text", "")), doc, loop_idx)
            p.write(text)
        elif task.task_type == "Keypress":
            p.press(str(params.get("key", "enter")))
        elif task.task_type == "Hotkey":
            keys = [k.strip() for k in str(params.get("keys", "ctrl,c")).split(",") if k.strip()]
            p.hotkey(*keys)
        elif task.task_type == "Condition":
            x = int(params.get("x", 0))
            y = int(params.get("y", 0))
            rgb = tuple(int(v.strip()) for v in str(params.get("rgb", "0,0,0")).split(","))
            timeout = float(params.get("timeout", 30))
            end = time.time() + timeout
            while time.time() < end:
                current = p.screenshot().getpixel((x, y))
                if tuple(current[:3]) == rgb:
                    break
                if self._stop_requested:
                    return
                self._pause_event.wait()
                time.sleep(0.2)
        elif task.task_type == "Wait":
            seconds = float(params.get("seconds", 1))
            end = time.time() + seconds
            while time.time() < end:
                if self._stop_requested:
                    return
                self._pause_event.wait()
                time.sleep(0.1)
        elif task.task_type == "Variable":
            # Variable task acts as a marker for loop data usage in String tasks.
            return
        elif task.task_type == "Screenshot":
            folder = Path(params.get("folder", "screenshots"))
            folder.mkdir(parents=True, exist_ok=True)
            base_name = str(params.get("base_name", "shot")).strip() or "shot"
            file_name = f"{base_name}_loop_{loop_idx + 1:03d}_pass_{task_idx + 1:03d}.png"
            p.screenshot(str(folder / file_name))

    def _resolve_value(self, value: str, doc: MacroDocument, loop_idx: int) -> str:
        out = value
        for key, values in doc.variables.items():
            token = "{{" + key + "}}"
            if token in out and values:
                out = out.replace(token, values[loop_idx % len(values)])
        return out
