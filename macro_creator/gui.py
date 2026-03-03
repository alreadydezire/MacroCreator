from __future__ import annotations

import time
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Dict

from .engine import MacroEngine
from .models import MacroDocument, Task
from .storage import MacroStorage

try:
    import keyboard  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    keyboard = None

try:
    import pyautogui as _pyautogui
except Exception:  # pragma: no cover - optional dependency
    _pyautogui = None


TASK_FIELDS: Dict[str, Dict[str, str]] = {
    "Click": {"x": "0", "y": "0", "click_type": "Single", "button": "Left"},
    "String": {"text": ""},
    "Keypress": {"key": "enter"},
    "Hotkey": {"keys": "ctrl,c"},
    "Condition": {"x": "0", "y": "0", "rgb": "255,255,255", "timeout": "30"},
    "Wait": {"seconds": "1"},
    "Variable": {"name": ""},
    "Screenshot": {"folder": "screenshots"},
}


class TaskEditorDialog:
    def __init__(self, parent: tk.Tk, task_type: str, initial: Dict[str, str], variable_sources: list[str] | None = None) -> None:
        self.parent = parent
        self.task_type = task_type
        self.variable_sources = variable_sources or []
        self.window = tk.Toplevel(parent)
        self.window.title(f"{task_type} Task")
        self.window.transient(parent)
        self.window.grab_set()

        self.result: Dict[str, str] | None = None
        self.entries: Dict[str, tk.StringVar] = {}

        row = 0
        if task_type == "Click":
            self._position_row(row, "Mouse Position", "x", "y")
            row += 1
            self._dropdown_row(row, "Click Type", "click_type", ["Single", "Double"])
            row += 1
            self._dropdown_row(row, "Mouse Button", "button", ["Left", "Right", "Middle"])
            row += 1
        elif task_type == "Condition":
            self._position_row(row, "Mouse Position", "x", "y")
            row += 1
            self._color_row(row, "Color (RGB)", "rgb")
            row += 1
            self._entry_row(row, "Timeout (sec)", "timeout")
            row += 1
        elif task_type == "Keypress":
            self._key_capture_row(row, "Target Key", "key")
            row += 1
        elif task_type == "Wait":
            self._spinbox_row(row, "Time (seconds)", "seconds")
            row += 1
        elif task_type == "Variable":
            self._variable_source_row(row, "Variable Source", "name")
            row += 1
        else:
            for key in initial:
                self._entry_row(row, key, key)
                row += 1

        for key, value in initial.items():
            if key in self.entries and value != "":
                self.entries[key].set(str(value))

        btns = ttk.Frame(self.window)
        btns.grid(row=row, column=0, columnspan=4, pady=8)
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Cancel", command=self.window.destroy).pack(side=tk.LEFT, padx=4)

    def _entry_row(self, row: int, label: str, key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar())
        ttk.Entry(self.window, textvariable=var, width=42).grid(row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4)

    def _spinbox_row(self, row: int, label: str, key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="1"))
        ttk.Spinbox(self.window, from_=0, to=999999, increment=0.1, textvariable=var, width=15).grid(
            row=row, column=1, sticky=tk.W, padx=8, pady=4
        )

    def _dropdown_row(self, row: int, label: str, key: str, options: list[str]) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value=options[0]))
        ttk.Combobox(self.window, textvariable=var, values=options, state="readonly", width=39).grid(
            row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4
        )

    def _variable_source_row(self, row: int, label: str, key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        options = self.variable_sources if self.variable_sources else ["No Sources"]
        var = self.entries.setdefault(key, tk.StringVar(value=options[0]))
        state = "readonly" if self.variable_sources else "disabled"
        ttk.Combobox(self.window, textvariable=var, values=options, state=state, width=39).grid(
            row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4
        )

    def _key_capture_row(self, row: int, label: str, key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="enter"))
        ttk.Entry(self.window, textvariable=var, width=24).grid(row=row, column=1, sticky=tk.W, padx=8, pady=4)
        ttk.Button(self.window, text="Capture", command=lambda: self._capture_key(var)).grid(
            row=row, column=2, sticky=tk.W, padx=8, pady=4
        )

    def _position_row(self, row: int, label: str, x_key: str, y_key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        x_var = self.entries.setdefault(x_key, tk.StringVar(value="0"))
        y_var = self.entries.setdefault(y_key, tk.StringVar(value="0"))
        ttk.Entry(self.window, textvariable=x_var, width=10).grid(row=row, column=1, sticky=tk.W, padx=(8, 2), pady=4)
        ttk.Entry(self.window, textvariable=y_var, width=10).grid(row=row, column=2, sticky=tk.W, padx=(2, 8), pady=4)
        ttk.Button(self.window, text="Capture", command=lambda: self._capture_position(x_var, y_var)).grid(
            row=row, column=3, sticky=tk.W, padx=8, pady=4
        )

    def _color_row(self, row: int, label: str, key: str) -> None:
        ttk.Label(self.window, text=label).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="255,255,255"))
        ttk.Entry(self.window, textvariable=var, width=26).grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=8, pady=4)
        ttk.Button(self.window, text="Dropper", command=lambda: self._capture_color(var)).grid(
            row=row, column=3, sticky=tk.W, padx=8, pady=4
        )

    def _capture_key(self, key_var: tk.StringVar) -> None:
        messagebox.showinfo("Capture Key", "Press any key to capture for this task.", parent=self.window)
        captured = _capture_keypress(self.window)
        if captured:
            key_var.set(captured)

    def _capture_position(self, x_var: tk.StringVar, y_var: tk.StringVar) -> None:
        messagebox.showinfo(
            "Capture Position",
            "Move mouse to desired location, then press CTRL to capture. Press ESC to cancel.",
            parent=self.window,
        )
        point = _capture_with_hotkey(self.window, capture_type="position")
        if point:
            x_var.set(str(point[0]))
            y_var.set(str(point[1]))

    def _capture_color(self, rgb_var: tk.StringVar) -> None:
        messagebox.showinfo(
            "Capture Color",
            "Move mouse to desired location, then press CTRL to capture color. Press ESC to cancel.",
            parent=self.window,
        )
        color = _capture_with_hotkey(self.window, capture_type="color")
        if color:
            rgb_var.set(",".join(str(c) for c in color))

    def _save(self) -> None:
        self.result = {k: v.get() for k, v in self.entries.items()}
        self.window.destroy()


def _capture_keypress(window: tk.Toplevel) -> str | None:
    if keyboard is not None:
        key = keyboard.read_key()
        return None if key == "esc" else str(key)

    result: dict[str, Any] = {"value": None, "cancel": False}

    def on_key(event: tk.Event) -> None:
        if event.keysym == "Escape":
            result["cancel"] = True
            return
        result["value"] = event.keysym.lower()

    window.bind("<Key>", on_key)
    window.focus_force()
    while result["value"] is None and not result["cancel"]:
        window.update()
        time.sleep(0.03)
    window.unbind("<Key>")
    return result["value"]


def _capture_with_hotkey(window: tk.Toplevel, capture_type: str) -> tuple[int, int] | tuple[int, int, int] | None:
    if _pyautogui is None:
        messagebox.showerror("Unavailable", "pyautogui is required for capture actions.", parent=window)
        return None

    if keyboard is not None:
        while True:
            key = keyboard.read_key()
            if key == "ctrl":
                pos = _pyautogui.position()
                if capture_type == "position":
                    return int(pos.x), int(pos.y)
                pixel = _pyautogui.screenshot().getpixel((pos.x, pos.y))
                return int(pixel[0]), int(pixel[1]), int(pixel[2])
            if key == "esc":
                return None

    result: dict[str, Any] = {"value": None, "cancel": False}

    def on_ctrl(_event=None):
        pos = _pyautogui.position()
        if capture_type == "position":
            result["value"] = (int(pos.x), int(pos.y))
        else:
            pixel = _pyautogui.screenshot().getpixel((pos.x, pos.y))
            result["value"] = (int(pixel[0]), int(pixel[1]), int(pixel[2]))

    def on_esc(_event=None):
        result["cancel"] = True

    window.bind("<Control_L>", on_ctrl)
    window.bind("<Control_R>", on_ctrl)
    window.bind("<Escape>", on_esc)

    while result["value"] is None and not result["cancel"]:
        window.update()
        time.sleep(0.03)

    window.unbind("<Control_L>")
    window.unbind("<Control_R>")
    window.unbind("<Escape>")
    return result["value"]


class MacroCreatorApp:
    def __init__(self) -> None:
        self.doc = MacroDocument()
        self.engine = MacroEngine()
        self.file_path: str | None = None
        self.resume_loop = 0
        self.resume_task = 0

        self.root = tk.Tk()
        self.root.title("MacroCreator")
        self.root.geometry("980x620")

        self.status = tk.StringVar(value="Idle")
        self.loop_var = tk.IntVar(value=1)

        self._build_menu()
        self._build_layout()
        self._refresh_list()
        self._tick_status()

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="New", command=self.new_macro)
        file_menu.add_command(label="Open", command=self.open_macro)
        file_menu.add_command(label="Save", command=self.save_macro)
        file_menu.add_command(label="Save As", command=self.save_as_macro)
        file_menu.add_separator()
        file_menu.add_command(label="Import CSV Variables", command=self.import_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Edit Selected", command=self.edit_selected_task)
        edit_menu.add_command(label="Duplicate Selected", command=self.duplicate_selected)
        edit_menu.add_command(label="Delete Selected", command=self.delete_selected)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)

        macro_menu = tk.Menu(menu_bar, tearoff=0)
        macro_menu.add_command(label="Run", command=self.run_macro)
        macro_menu.add_command(label="Pause", command=self.pause_macro)
        macro_menu.add_command(label="Resume", command=self.resume_macro)
        macro_menu.add_command(label="Stop", command=self.stop_macro)
        macro_menu.add_command(label="Continue From Stop Point", command=self.continue_macro)
        menu_bar.add_cascade(label="Macro", menu=macro_menu)

        other_menu = tk.Menu(menu_bar, tearoff=0)
        other_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "MacroCreator prototype"))
        menu_bar.add_cascade(label="Other", menu=other_menu)

        self.root.config(menu=menu_bar)

    def _build_layout(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right = ttk.Frame(main)
        right.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(left, columns=("type", "summary"), show="headings", selectmode="browse")
        self.tree.heading("type", text="Task Type")
        self.tree.heading("summary", text="Parameters")
        self.tree.column("type", width=160, stretch=False)
        self.tree.column("summary", width=620)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", lambda _: self.edit_selected_task())

        controls = [
            ("Add Task", self.add_task),
            ("Edit Task", self.edit_selected_task),
            ("Duplicate", self.duplicate_selected),
            ("Delete", self.delete_selected),
            ("Move Up", self.move_up),
            ("Move Down", self.move_down),
        ]
        for text, cmd in controls:
            ttk.Button(right, text=text, command=cmd, width=26).pack(pady=4)

        ttk.Separator(right, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)
        ttk.Label(right, text="Loop count:").pack(anchor=tk.W)
        ttk.Spinbox(right, from_=1, to=999999, textvariable=self.loop_var, width=10).pack(anchor=tk.W)

        ttk.Button(right, text="Run", command=self.run_macro, width=26).pack(pady=4)
        ttk.Button(right, text="Pause", command=self.pause_macro, width=26).pack(pady=4)
        ttk.Button(right, text="Resume", command=self.resume_macro, width=26).pack(pady=4)
        ttk.Button(right, text="Stop", command=self.stop_macro, width=26).pack(pady=4)
        ttk.Button(right, text="Continue", command=self.continue_macro, width=26).pack(pady=4)

        status_bar = ttk.Label(self.root, textvariable=self.status, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _tick_status(self) -> None:
        if self.engine.state.running:
            label = "Paused" if self.engine.state.paused else "Running"
            self.status.set(
                f"{label} | Loop {self.engine.state.current_loop + 1}/{self.doc.loop_count} | Task {self.engine.state.current_task_index + 1}/{len(self.doc.tasks)}"
            )
            self.resume_loop = self.engine.state.current_loop
            self.resume_task = self.engine.state.current_task_index
        self.root.after(200, self._tick_status)

    def _refresh_list(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for idx, task in enumerate(self.doc.tasks):
            self.tree.insert("", tk.END, iid=str(idx), values=(task.task_type, str(task.params)))

    def _selected_index(self) -> int | None:
        selected = self.tree.selection()
        if not selected:
            return None
        return int(selected[0])

    def _open_task_editor(self, task_type: str, initial: Dict[str, str]) -> Dict[str, str] | None:
        dialog = TaskEditorDialog(self.root, task_type, initial, variable_sources=list(self.doc.variables.keys()))
        self.root.wait_window(dialog.window)
        return dialog.result

    def add_task(self) -> None:
        kind = simpledialog.askstring(
            "Task Type",
            "Enter task type:\nClick, String, Keypress, Hotkey, Condition, Wait, Variable, Screenshot",
        )
        if not kind:
            return
        kind = kind.strip()
        if kind not in TASK_FIELDS:
            messagebox.showerror("Error", "Unsupported task type")
            return
        params = self._open_task_editor(kind, TASK_FIELDS[kind].copy())
        if params is None:
            return
        self.doc.tasks.append(Task(kind, params))
        self._refresh_list()

    def edit_selected_task(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        task = self.doc.tasks[idx]
        params = self._open_task_editor(task.task_type, dict(task.params))
        if params is None:
            return
        task.params = params
        self._refresh_list()

    def duplicate_selected(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        t = self.doc.tasks[idx]
        self.doc.tasks.insert(idx + 1, Task(t.task_type, dict(t.params)))
        self._refresh_list()

    def delete_selected(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        del self.doc.tasks[idx]
        self._refresh_list()

    def move_up(self) -> None:
        idx = self._selected_index()
        if idx is None or idx == 0:
            return
        self.doc.tasks[idx - 1], self.doc.tasks[idx] = self.doc.tasks[idx], self.doc.tasks[idx - 1]
        self._refresh_list()
        self.tree.selection_set(str(idx - 1))

    def move_down(self) -> None:
        idx = self._selected_index()
        if idx is None or idx >= len(self.doc.tasks) - 1:
            return
        self.doc.tasks[idx + 1], self.doc.tasks[idx] = self.doc.tasks[idx], self.doc.tasks[idx + 1]
        self._refresh_list()
        self.tree.selection_set(str(idx + 1))

    def run_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        self.engine.run_async(self.doc, self.status.set)

    def pause_macro(self) -> None:
        self.engine.pause()

    def resume_macro(self) -> None:
        self.engine.resume()

    def stop_macro(self) -> None:
        self.engine.stop()

    def continue_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        self.engine.run_async(self.doc, self.status.set, self.resume_loop, self.resume_task)

    def import_csv(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path:
            return
        self.doc.variables = MacroStorage.load_csv_variables(path)
        messagebox.showinfo("Imported", f"Loaded columns: {', '.join(self.doc.variables.keys())}")

    def new_macro(self) -> None:
        self.doc = MacroDocument()
        self.file_path = None
        self.loop_var.set(1)
        self._refresh_list()

    def save_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        if not self.file_path:
            self.save_as_macro()
            return
        MacroStorage.save(self.file_path, self.doc)
        self.status.set(f"Saved: {self.file_path}")

    def save_as_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Macro JSON", "*.json")])
        if not path:
            return
        self.file_path = path
        MacroStorage.save(path, self.doc)
        self.status.set(f"Saved: {path}")

    def open_macro(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Macro JSON", "*.json")])
        if not path:
            return
        self.file_path = path
        self.doc = MacroStorage.load(path)
        self.loop_var.set(self.doc.loop_count)
        self._refresh_list()
        self.status.set(f"Opened: {path}")

    def run(self) -> None:
        self.root.mainloop()
