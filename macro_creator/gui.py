from __future__ import annotations

import csv
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Dict

try:
    from PIL import Image, ImageTk
except Exception:  # pragma: no cover
    Image = None
    ImageTk = None

from .engine import MacroEngine
from .models import MacroDocument, Task
from .storage import MacroStorage

PROTO_DIR = Path(__file__).resolve().parent.parent / "ProtoType"
if str(PROTO_DIR) not in sys.path:
    sys.path.append(str(PROTO_DIR))

try:
    from ToolTip import CreateToolTip
except Exception:  # pragma: no cover
    def CreateToolTip(_widget, _text):
        return None

try:
    import keyboard  # type: ignore
except Exception:  # pragma: no cover
    keyboard = None

try:
    import pyautogui as _pyautogui
except Exception:  # pragma: no cover
    _pyautogui = None

TASK_FIELDS: Dict[str, Dict[str, str]] = {
    "Click": {"x": "0", "y": "0", "click_type": "Single", "button": "Left", "comment": ""},
    "String": {"text": "", "comment": ""},
    "Keypress": {"key": "enter", "comment": ""},
    "Hotkey": {"keys": "ctrl,c", "comment": ""},
    "Condition": {"x": "0", "y": "0", "rgb": "255,255,255", "timeout": "30", "comment": ""},
    "Wait": {"seconds": "1", "comment": ""},
    "Variable": {"name": "", "comment": ""},
    "Screenshot": {"folder": "screenshots", "base_name": "shot", "comment": ""},
}
TASKS_WITH_POSITION = {"Click", "Condition"}

TASK_ICON_MAP = {
    "Click": "Images/Tasks/click-ico.png",
    "String": "Images/Tasks/type-ico.png",
    "Keypress": "Images/Tasks/press-ico.png",
    "Hotkey": "Images/Tasks/hotkey-ico.png",
    "Condition": "Images/Tasks/condition-ico.png",
    "Wait": "Images/Tasks/wait-ico.png",
    "Variable": "Images/Tasks/variable-ico.png",
    "Screenshot": "Images/Menu/Other/about.png",
}
MENU_ICON_MAP = {
    "new": "Images/Menu/File/new.png",
    "open": "Images/Menu/File/open.png",
    "save": "Images/Menu/File/save.png",
    "saveas": "Images/Menu/File/saveas.png",
    "exit": "Images/Menu/File/exit.png",
    "copy": "Images/Menu/Edit/copy.png",
    "paste": "Images/Menu/Edit/paste.png",
    "delete": "Images/Menu/Edit/delete.png",
    "greenplay": "Images/Menu/Macro/greenplay-ico.png",
    "task": "Images/Menu/Macro/task-ico.png",
    "varsett": "Images/Menu/Macro/variable-setting.png",
    "help": "Images/Menu/Other/help.png",
    "about": "Images/Menu/Other/about.png",
    "settings": "Images/sett-ico.png",
}

TASKLIST_ICON_MAP = {
    "show": "Images/TaskList/show.png",
    "up": "Images/TaskList/up-arrow.png",
    "down": "Images/TaskList/down-arrow.png",
    "edit": "Images/TaskList/edit.png",
    "duplicate": "Images/TaskList/duplicate.png",
}



def _hex_from_rgb_string(rgb_text: str) -> str:
    try:
        r, g, b = [max(0, min(255, int(v.strip()))) for v in rgb_text.split(",")[:3]]
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return "#ffffff"


def _capture_with_hotkey(capture_type: str) -> tuple[int, int] | tuple[int, int, int] | None:
    if _pyautogui is None or keyboard is None:
        return None
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


class TaskEditorDialog:
    def __init__(self, parent: tk.Tk, task_type: str, initial: Dict[str, str], variable_sources: list[str], icon=None) -> None:
        self.parent = parent
        self.task_type = task_type
        self.variable_sources = variable_sources
        self.result: Dict[str, str] | None = None
        self.entries: Dict[str, tk.StringVar] = {}

        self.window = tk.Toplevel(parent)
        self.window.title(f"{task_type} Task")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg="white")

        row = 0
        if icon is not None:
            tk.Label(self.window, image=icon, bg="white").grid(row=row, column=0, padx=8, pady=8, sticky="w")
            tk.Label(self.window, text=f"{task_type} Task", bg="white", font=("Yu Gothic UI Semibold", 12)).grid(row=row, column=1, columnspan=4, sticky="w")
            row += 1

        if task_type == "Click":
            self._position_row(row, "Mouse Position", "x", "y"); row += 1
            self._dropdown_row(row, "Click Type", "click_type", ["Single", "Double"]); row += 1
            self._dropdown_row(row, "Mouse Button", "button", ["Left", "Right", "Middle"]); row += 1
        elif task_type == "Condition":
            self._position_row(row, "Mouse Position", "x", "y"); row += 1
            self._color_row(row, "Color (RGB)", "rgb"); row += 1
            self._entry_row(row, "Timeout (sec)", "timeout"); row += 1
        elif task_type == "Keypress":
            self._keypress_row(row, "Target Key", "key"); row += 1
        elif task_type == "Wait":
            self._spinbox_row(row, "Time (seconds)", "seconds"); row += 1
        elif task_type == "Variable":
            self._variable_row(row, "Variable Source", "name"); row += 1
        elif task_type == "Screenshot":
            self._screenshot_row(row); row += 1
        else:
            for k in initial:
                if k != "comment":
                    self._entry_row(row, k, k)
                    row += 1

        self._entry_row(row, "Comment (optional)", "comment")
        CreateToolTip(self.window.grid_slaves(row=row, column=1)[0], "Optional label to help identify this task in the task list.")
        row += 1

        for k, v in initial.items():
            if k in self.entries and v != "":
                self.entries[k].set(str(v))
        self._sync_color_preview()

        btns = tk.Frame(self.window, bg="white")
        btns.grid(row=row, column=0, columnspan=5, pady=8)
        tk.Button(btns, text="Save", command=self._save, relief="ridge", bg="white").pack(side=tk.LEFT, padx=4)
        tk.Button(btns, text="Cancel", command=self.window.destroy, relief="ridge", bg="white").pack(side=tk.LEFT, padx=4)

        self.window.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.window.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{max(0,x)}+{max(0,y)}")

    def _entry_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar())
        tk.Entry(self.window, textvariable=var, bg="white", relief="ridge", width=45).grid(row=row, column=1, columnspan=4, sticky="w", padx=8, pady=4)

    def _spinbox_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="1"))
        tk.Spinbox(self.window, from_=0, to=999999, increment=0.1, textvariable=var, bg="white", relief="ridge", width=16).grid(row=row, column=1, sticky="w", padx=8, pady=4)

    def _dropdown_row(self, row: int, label: str, key: str, options: list[str]) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value=options[0]))
        ttk.Combobox(self.window, textvariable=var, values=options, state="readonly", width=42).grid(row=row, column=1, columnspan=4, sticky="w", padx=8, pady=4)

    def _position_row(self, row: int, label: str, x_key: str, y_key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        x_var = self.entries.setdefault(x_key, tk.StringVar(value="0"))
        y_var = self.entries.setdefault(y_key, tk.StringVar(value="0"))
        tk.Entry(self.window, textvariable=x_var, bg="white", relief="ridge", width=10).grid(row=row, column=1, sticky="w", padx=(8, 2), pady=4)
        tk.Entry(self.window, textvariable=y_var, bg="white", relief="ridge", width=10).grid(row=row, column=2, sticky="w", padx=(2, 8), pady=4)
        b = tk.Button(self.window, text="📍", relief="ridge", bg="white", width=3, command=lambda: self._capture_position(x_var, y_var))
        b.grid(row=row, column=3, sticky="w", padx=4)
        CreateToolTip(b, "Move mouse then press CTRL to capture position. ESC to cancel.")

    def _color_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="255,255,255"))
        var.trace_add("write", lambda *_: self._sync_color_preview())
        tk.Entry(self.window, textvariable=var, bg="white", relief="ridge", width=20).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        self.color_preview = tk.Label(self.window, text="   ", bg="#ffffff", relief="ridge", width=4)
        self.color_preview.grid(row=row, column=2, sticky="w", padx=4)
        b = tk.Button(self.window, text="🧪", relief="ridge", bg="white", width=3, command=lambda: self._capture_color(var))
        b.grid(row=row, column=3, sticky="w", padx=4)
        CreateToolTip(b, "Move mouse then press CTRL to capture RGB at cursor. ESC to cancel.")

    def _sync_color_preview(self) -> None:
        if hasattr(self, "color_preview"):
            rgb = self.entries.get("rgb", tk.StringVar(value="255,255,255")).get()
            self.color_preview.configure(bg=_hex_from_rgb_string(rgb))

    def _keypress_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="enter"))
        tk.Entry(self.window, textvariable=var, bg="white", relief="ridge", width=20).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        b = tk.Button(self.window, text="⌨", relief="ridge", bg="white", width=3, command=lambda: self._capture_key(var))
        b.grid(row=row, column=2, sticky="w", padx=4)
        CreateToolTip(b, "Press any key to capture.")

    def _variable_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        opts = self.variable_sources if self.variable_sources else ["No Sources"]
        var = self.entries.setdefault(key, tk.StringVar(value=opts[0]))
        c = ttk.Combobox(self.window, textvariable=var, values=opts, state="readonly" if self.variable_sources else "disabled", width=42)
        c.grid(row=row, column=1, columnspan=4, sticky="w", padx=8, pady=4)

    def _screenshot_row(self, row: int) -> None:
        tk.Label(self.window, text="Folder", bg="white").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        folder_var = self.entries.setdefault("folder", tk.StringVar(value="screenshots"))
        tk.Entry(self.window, textvariable=folder_var, bg="white", relief="ridge", width=30).grid(row=row, column=1, sticky="w", padx=8)
        btn = tk.Button(self.window, text="📁", relief="ridge", bg="white", width=3, command=lambda: self._pick_folder(folder_var))
        btn.grid(row=row, column=2, sticky="w")
        CreateToolTip(btn, "Browse for screenshot folder")
        tk.Label(self.window, text="Base Name", bg="white").grid(row=row, column=3, sticky="e", padx=4)
        base_var = self.entries.setdefault("base_name", tk.StringVar(value="shot"))
        tk.Entry(self.window, textvariable=base_var, bg="white", relief="ridge", width=14).grid(row=row, column=4, sticky="w", padx=4)

    def _pick_folder(self, var: tk.StringVar) -> None:
        p = filedialog.askdirectory(parent=self.window)
        if p:
            var.set(p)

    def _capture_key(self, key_var: tk.StringVar) -> None:
        if keyboard is None:
            return
        k = keyboard.read_key()
        if k != "esc":
            key_var.set(k)

    def _capture_position(self, x_var: tk.StringVar, y_var: tk.StringVar) -> None:
        point = _capture_with_hotkey("position")
        if point:
            x_var.set(str(point[0])); y_var.set(str(point[1]))

    def _capture_color(self, rgb_var: tk.StringVar) -> None:
        color = _capture_with_hotkey("color")
        if color:
            rgb_var.set(f"{color[0]},{color[1]},{color[2]}")

    def _save(self) -> None:
        self.result = {k: v.get() for k, v in self.entries.items()}
        self.window.destroy()


class MacroCreatorApp:
    def __init__(self) -> None:
        self.doc = MacroDocument()
        self.engine = MacroEngine()
        self.file_path: str | None = None
        self.selected_index: int | None = None
        self.icons: Dict[str, Any] = {}
        self.menu_icons: Dict[str, Any] = {}
        self.tasklist_icons: Dict[str, Any] = {}
        self.csv_sources: Dict[str, str] = {}
        self.dirty = False

        self.root = tk.Tk()
        self.root.title("MacroCreator")
        self.root.geometry("1180x760")
        self.root.configure(bg="black")

        self.state_var = tk.StringVar(value="State: Off")
        self.task_info_var = tk.StringVar(value="Current Task/Loop: -")
        self.loop_info_var = tk.StringVar(value="")
        self.action_var = tk.StringVar(value="Details: Idle")

        self.loop_mode_var = tk.StringVar(value="count")
        self.loop_var = tk.IntVar(value=1)
        self.loop_csv_var = tk.StringVar(value="")
        self.quick_delay = tk.StringVar(value="0")
        self.quick_fail_safe = tk.BooleanVar(value=True)

        self._load_images()
        self._build_menu()
        self._build_layout()
        self._refresh_loop_csv_options()
        self._refresh_task_rows()
        self._refresh_macro_info()
        self._tick_status()

    def _load_image(self, rel: str, size=(18, 18)):
        if Image is None or ImageTk is None:
            return None
        p = PROTO_DIR / rel
        if not p.exists():
            return None
        return ImageTk.PhotoImage(Image.open(p).resize(size, Image.LANCZOS))

    def _load_images(self) -> None:
        for t, rel in TASK_ICON_MAP.items():
            self.icons[t] = self._load_image(rel, (20, 20))
        self.play_icon = self._load_image("Images/play-ico.png", (20, 20))
        self.pause_icon = self._load_image("Images/pause-ico.png", (20, 20))
        self.settings_icon = self._load_image("Images/sett-ico.png", (20, 20))
        for k, rel in MENU_ICON_MAP.items():
            self.menu_icons[k] = self._load_image(rel, (18, 18))
        for k, rel in TASKLIST_ICON_MAP.items():
            self.tasklist_icons[k] = self._load_image(rel, (14, 14))

    def _build_menu(self) -> None:
        m = tk.Menu(self.root)
        file_menu = tk.Menu(m, tearoff=0)
        file_menu.add_command(label="New", command=self.new_macro, image=self.menu_icons.get("new"), compound=tk.LEFT)
        file_menu.add_command(label="Open", command=self.open_macro, image=self.menu_icons.get("open"), compound=tk.LEFT)
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=self.save_macro, image=self.menu_icons.get("save"), compound=tk.LEFT)
        file_menu.add_command(label="Save As", command=self.save_as_macro, image=self.menu_icons.get("saveas"), compound=tk.LEFT)
        file_menu.add_separator()
        file_menu.add_command(label="Manage CSVs", command=self.open_manage_csvs_window, image=self.menu_icons.get("varsett"), compound=tk.LEFT)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, image=self.menu_icons.get("exit"), compound=tk.LEFT)
        m.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(m, tearoff=0)
        edit_menu.add_command(label="Edit Selected", command=self.edit_selected_task, image=self.menu_icons.get("copy"), compound=tk.LEFT)
        edit_menu.add_command(label="Duplicate Selected", command=self.duplicate_selected, image=self.menu_icons.get("paste"), compound=tk.LEFT)
        edit_menu.add_command(label="Delete Selected", command=self.delete_selected, image=self.menu_icons.get("delete"), compound=tk.LEFT)
        m.add_cascade(label="Edit", menu=edit_menu)

        macro_menu = tk.Menu(m, tearoff=0)
        macro_menu.add_command(label="Run/Pause", command=self.run_pause_continue, image=self.menu_icons.get("greenplay"), compound=tk.LEFT)
        macro_menu.add_command(label="Cancel", command=self.stop_macro, image=self.menu_icons.get("delete"), compound=tk.LEFT)
        create = tk.Menu(macro_menu, tearoff=0)
        for task in TASK_FIELDS:
            create.add_command(label=task, command=lambda t=task: self.create_task(t), image=self.icons.get(task), compound=tk.LEFT)
        macro_menu.add_cascade(label="Create Task", menu=create, image=self.menu_icons.get("task"), compound=tk.LEFT)
        m.add_cascade(label="Macro", menu=macro_menu)

        settings_menu = tk.Menu(m, tearoff=0)
        settings_menu.add_command(label="All Settings", command=self.open_settings_window, image=self.menu_icons.get("settings"), compound=tk.LEFT)
        settings_menu.add_separator()
        settings_menu.add_command(label="CSV Handling Settings", command=lambda: self.open_settings_window("csv"))
        settings_menu.add_command(label="Loop Behaviour Settings", command=lambda: self.open_settings_window("loop"))
        settings_menu.add_command(label="Safety / Environment Options", command=lambda: self.open_settings_window("safety"))
        settings_menu.add_command(label="Advanced Settings", command=lambda: self.open_settings_window("advanced"))
        m.add_cascade(label="Settings", menu=settings_menu)

        other = tk.Menu(m, tearoff=0)
        other.add_command(label="Help", command=lambda: messagebox.showinfo("Help", "Use tooltips on controls."), image=self.menu_icons.get("help"), compound=tk.LEFT)
        other.add_command(label="About", command=lambda: messagebox.showinfo("About", "MacroCreator"), image=self.menu_icons.get("about"), compound=tk.LEFT)
        m.add_cascade(label="Other", menu=other)
        self.root.config(menu=m)

    def _build_layout(self) -> None:
        outer = tk.Frame(self.root, bg="black")
        outer.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        tk.Label(outer, text="Macro Tasks", bg="white", relief="ridge", font=("Yu Gothic UI Semibold", 16)).pack(fill=tk.X, pady=(0, 4))

        info_frame = tk.Frame(outer, bg="white", relief="ridge", bd=1)
        info_frame.pack(fill=tk.X, pady=(0, 4))
        tk.Label(info_frame, text="Current File:", bg="white", font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=(6,2))
        self.current_file_var = tk.StringVar(value="Untitled")
        tk.Label(info_frame, textvariable=self.current_file_var, bg="white", anchor="w").pack(side=tk.LEFT)
        self.unsaved_var = tk.StringVar(value="")
        tk.Label(info_frame, textvariable=self.unsaved_var, bg="white", fg="red", font=("TkDefaultFont", 12, "bold")).pack(side=tk.LEFT, padx=(2,8))
        self.task_count_var = tk.StringVar(value="Tasks: 0")
        tk.Label(info_frame, textvariable=self.task_count_var, bg="white").pack(side=tk.RIGHT, padx=8)

        state_frame = tk.Frame(outer, bg="white", relief="ridge", bd=1)
        state_frame.pack(fill=tk.X, pady=(0, 4))
        tk.Label(state_frame, textvariable=self.state_var, bg="white", anchor="w", font=("Yu Gothic UI Semibold", 11)).pack(fill=tk.X)
        tk.Label(state_frame, textvariable=self.task_info_var, bg="white", anchor="w").pack(fill=tk.X)
        tk.Label(state_frame, textvariable=self.action_var, bg="white", anchor="w").pack(fill=tk.X)

        list_wrap = tk.Frame(outer, bg="white", relief="sunken", bd=1)
        list_wrap.pack(fill=tk.BOTH, expand=True)
        self.task_canvas = tk.Canvas(list_wrap, bg="white", highlightthickness=0)
        y = ttk.Scrollbar(list_wrap, orient="vertical", command=self.task_canvas.yview)
        self.task_inner = tk.Frame(self.task_canvas, bg="white")
        self.task_window = self.task_canvas.create_window((0, 0), window=self.task_inner, anchor="nw")
        self.task_canvas.configure(yscrollcommand=y.set)
        self.task_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_inner.bind("<Configure>", lambda _e: self.task_canvas.configure(scrollregion=self.task_canvas.bbox("all")))
        self.task_canvas.bind("<Configure>", lambda e: self.task_canvas.itemconfig(self.task_window, width=e.width))

        bottom = tk.Frame(outer, bg="white", relief="ridge", bd=1)
        bottom.pack(fill=tk.X, pady=(4, 0))

        task_box = tk.LabelFrame(bottom, text="Create a Task", bg="white")
        task_box.pack(fill=tk.X, padx=4, pady=4)
        for i, t in enumerate(TASK_FIELDS.keys()):
            task_box.grid_columnconfigure(i, weight=1)
            b = tk.Button(task_box, text=t, image=self.icons.get(t), compound="left", relief="ridge", bg="white", command=lambda x=t: self.create_task(x))
            b.grid(row=0, column=i, sticky="ew", padx=2, pady=2)

        quick = tk.LabelFrame(bottom, text="Quick Settings", bg="white")
        quick.pack(fill=tk.X, padx=4, pady=4)
        tk.Label(quick, text="Inter-task Delay (sec):", bg="white").grid(row=0, column=0, sticky="w", padx=4)
        tk.Entry(quick, textvariable=self.quick_delay, width=8, relief="ridge", bg="white").grid(row=0, column=1, sticky="w")
        tk.Checkbutton(quick, text="Fail-safe enabled", variable=self.quick_fail_safe, bg="white").grid(row=0, column=2, padx=10, sticky="w")

        tk.Radiobutton(quick, text="Loop Count", variable=self.loop_mode_var, value="count", bg="white").grid(row=1, column=0, sticky="w", padx=4)
        self.loop_count_spin = ttk.Spinbox(quick, from_=1, to=999999, textvariable=self.loop_var, width=8)
        self.loop_count_spin.grid(row=1, column=1, sticky="w")
        tk.Radiobutton(quick, text="Loop till CSV end", variable=self.loop_mode_var, value="csv_end", bg="white").grid(row=1, column=2, sticky="w")
        self.loop_csv_combo = ttk.Combobox(quick, textvariable=self.loop_csv_var, state="readonly", values=[], width=20)
        self.loop_csv_combo.grid(row=1, column=3, sticky="w", padx=4)
        tk.Radiobutton(quick, text="Loop till stopped", variable=self.loop_mode_var, value="stopped", bg="white").grid(row=1, column=4, sticky="w")
        self.loop_mode_var.trace_add("write", lambda *_: self._update_loop_mode_controls())
        self._update_loop_mode_controls()

        ctrl = tk.Frame(bottom, bg="white")
        ctrl.pack(fill=tk.X, padx=4, pady=4)
        settings_btn = tk.Button(ctrl, image=self.settings_icon, width=42, height=42, bg="#dddddd", relief="ridge", command=self.open_settings_window)
        settings_btn.pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(ctrl, text="Run", image=self.play_icon, compound="left", bg="lime green", fg="white", relief="ridge", command=self.run_pause_continue).pack_forget()
        self.run_btn = tk.Button(ctrl, text="Run", image=self.play_icon, compound="left", bg="lime green", fg="white", relief="ridge", font=("TkDefaultFont", 11, "bold"), padx=10, pady=8, command=self.run_pause_continue)
        self.run_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_btn = tk.Button(ctrl, text="Cancel", bg="red", fg="white", relief="ridge", state="disabled", width=12, font=("TkDefaultFont", 11, "bold"), padx=10, pady=8, command=self.stop_macro)
        self.cancel_btn.pack(side=tk.LEFT, padx=(6, 0))

    def _update_loop_mode_controls(self) -> None:
        mode = self.loop_mode_var.get()
        self.loop_count_spin.configure(state="normal" if mode == "count" else "disabled")
        self.loop_csv_combo.configure(state="readonly" if mode == "csv_end" else "disabled")

    def _bind_row_click(self, widget, idx: int) -> None:
        if isinstance(widget, tk.Button):
            return
        widget.bind("<Button-1>", lambda _e, i=idx: self._select(i))
        for child in widget.winfo_children():
            self._bind_row_click(child, idx)

    def _mark_dirty(self) -> None:
        self.dirty = True
        self._refresh_macro_info()

    def _refresh_macro_info(self) -> None:
        name = Path(self.file_path).name if self.file_path else "Untitled"
        self.current_file_var.set(name)
        self.unsaved_var.set("*" if self.dirty else "")
        self.task_count_var.set(f"Tasks: {len(self.doc.tasks)}")

    def _summary(self, task: Task) -> str:
        p = task.params
        if task.task_type == "Variable":
            return f"name: {p.get('name','')}"
        if task.task_type == "Keypress":
            return f"key: {p.get('key','')}"
        if task.task_type == "Click":
            return f"{p.get('click_type','Single')} {p.get('button','Left')} @ {p.get('x','0')},{p.get('y','0')}"
        if task.task_type == "Condition":
            return f"pos {p.get('x','0')},{p.get('y','0')} rgb {p.get('rgb','')}"
        text = ", ".join(f"{k}: {v}" for k, v in p.items() if k != "comment")
        return text

    def _short(self, text: str, n: int = 68) -> str:
        text = str(text)
        return text if len(text) <= n else text[: n - 1] + "…"

    def _refresh_task_rows(self) -> None:
        for c in self.task_inner.winfo_children():
            c.destroy()
        if not self.doc.tasks:
            tk.Label(self.task_inner, text="There are currently no tasks, use the options below to start now.", bg="white").pack(anchor="w", padx=8, pady=8)
            self._refresh_macro_info()
            return

        for idx, task in enumerate(self.doc.tasks):
            bg = "#dff0ff" if idx == self.selected_index else "white"
            row = tk.Frame(self.task_inner, bg=bg, relief="ridge", bd=1)
            row.pack(fill=tk.X, padx=2, pady=2)
            row.grid_columnconfigure(1, weight=1)
            row.grid_columnconfigure(2, minsize=230)

            tk.Label(row, image=self.icons.get(task.task_type), bg=bg).grid(row=0, column=0, padx=4, sticky="w")

            text_wrap = tk.Frame(row, bg=bg)
            text_wrap.grid(row=0, column=1, sticky="w")
            tk.Label(text_wrap, text=task.task_type, bg=bg, font=("TkDefaultFont", 9, "bold"), anchor="w").pack(side=tk.LEFT)
            tk.Label(text_wrap, text=" - " + self._short(self._summary(task), 90), bg=bg, font=("TkDefaultFont", 9, "italic"), anchor="w").pack(side=tk.LEFT)

            if task.task_type == "Condition":
                tk.Label(row, text="  ", bg=_hex_from_rgb_string(str(task.params.get("rgb", "255,255,255"))), relief="ridge", width=2).grid(row=0, column=1, sticky="e", padx=(0, 6))

            comment = self._short(str(task.params.get("comment", "")).strip(), 36)
            tk.Label(row, text=comment, bg=bg, fg="#444", width=28, anchor="w").grid(row=0, column=2, sticky="w")

            btns = tk.Frame(row, bg=bg)
            btns.grid(row=0, column=3, sticky="e", padx=2)
            action_btn_w = 42
            action_btn_h = 42
            show_btn = tk.Button(
                btns,
                image=self.tasklist_icons.get("show"),
                text="◎" if not self.tasklist_icons.get("show") else "",
                width=action_btn_w,
                height=action_btn_h,
                relief="ridge",
                bg="white",
                command=lambda i=idx: self.show_position(i),
                state=("normal" if task.task_type in TASKS_WITH_POSITION else "disabled"),
            )
            show_btn.pack(side=tk.LEFT, padx=1)
            CreateToolTip(show_btn, "Show task position")
            button_defs = [
                ("up", "↑", lambda i=idx: self.move_up(i), "black", "Move up"),
                ("down", "↓", lambda i=idx: self.move_down(i), "black", "Move down"),
                ("edit", "✎", lambda i=idx: self.edit_task_at(i), "black", "Edit"),
                ("duplicate", "⧉", lambda i=idx: self.duplicate_at(i), "black", "Duplicate"),
                (None, "✖", lambda i=idx: self.delete_at(i), "red", "Delete"),
            ]
            for key, fallback, cmd, fg, tip in button_defs:
                b = tk.Button(
                    btns,
                    image=self.tasklist_icons.get(key) if key else None,
                    text=fallback if (not key or not self.tasklist_icons.get(key)) else "",
                    width=action_btn_w,
                    height=action_btn_h,
                    relief="ridge",
                    bg="white",
                    fg=fg,
                    command=cmd,
                )
                b.pack(side=tk.LEFT, padx=1)
                CreateToolTip(b, tip)

            self._bind_row_click(row, idx)
        self._refresh_macro_info()

    def _select(self, i: int) -> None:
        self.selected_index = None if self.selected_index == i else i
        self._refresh_task_rows()

    def _edit_dialog(self, task_type: str, initial: Dict[str, str]) -> Dict[str, str] | None:
        if task_type == "Screenshot" and not initial.get("base_name"):
            initial["base_name"] = (Path(self.file_path).stem if self.file_path else "Untitled")
        d = TaskEditorDialog(self.root, task_type, initial, self._variable_source_options(), icon=self.icons.get(task_type))
        self.root.wait_window(d.window)
        return d.result

    def create_task(self, task_type: str) -> None:
        params = self._edit_dialog(task_type, TASK_FIELDS[task_type].copy())
        if params is None:
            return
        self.doc.tasks.append(Task(task_type, params)); self._mark_dirty(); self._refresh_task_rows()

    def edit_task_at(self, i: int) -> None:
        t = self.doc.tasks[i]
        p = self._edit_dialog(t.task_type, dict(t.params))
        if p is None:
            return
        t.params = p; self._mark_dirty(); self._refresh_task_rows()

    def duplicate_at(self, i: int) -> None:
        t = self.doc.tasks[i]
        self.doc.tasks.insert(i + 1, Task(t.task_type, dict(t.params))); self._mark_dirty(); self._refresh_task_rows()

    def move_up(self, i: int) -> None:
        if i == 0: return
        self.doc.tasks[i - 1], self.doc.tasks[i] = self.doc.tasks[i], self.doc.tasks[i - 1]
        self.selected_index = i - 1; self._mark_dirty(); self._refresh_task_rows()

    def move_down(self, i: int) -> None:
        if i >= len(self.doc.tasks) - 1: return
        self.doc.tasks[i + 1], self.doc.tasks[i] = self.doc.tasks[i], self.doc.tasks[i + 1]
        self.selected_index = i + 1; self._mark_dirty(); self._refresh_task_rows()

    def delete_at(self, i: int) -> None:
        del self.doc.tasks[i]; self.selected_index = None; self._mark_dirty(); self._refresh_task_rows()

    def show_position(self, i: int) -> None:
        t = self.doc.tasks[i]
        x = int(float(t.params.get("x", 0))); y = int(float(t.params.get("y", 0)))
        marker = tk.Toplevel(self.root)
        marker.overrideredirect(True)
        marker.attributes("-topmost", True)
        marker.geometry(f"30x30+{x-15}+{y-15}")
        tk.Label(marker, text="◎", fg="red", bg="yellow", font=("TkDefaultFont", 18)).pack(fill=tk.BOTH, expand=True)
        if _pyautogui is not None:
            _pyautogui.moveTo(x, y)

        def close(_e=None):
            marker.destroy()
        marker.bind("<Key>", close)
        marker.bind("<Button-1>", close)
        marker.focus_force()

    def edit_selected_task(self) -> None:
        if self.selected_index is not None: self.edit_task_at(self.selected_index)

    def duplicate_selected(self) -> None:
        if self.selected_index is not None: self.duplicate_at(self.selected_index)

    def delete_selected(self) -> None:
        if self.selected_index is not None: self.delete_at(self.selected_index)

    def run_pause_continue(self) -> None:
        if self.loop_mode_var.get() == "count":
            self.doc.loop_count = max(1, int(self.loop_var.get()))
        elif self.loop_mode_var.get() == "csv_end" and self.loop_csv_var.get() in self.csv_sources:
            csv_name = self.loop_csv_var.get()
            lengths = [len(v) for k,v in self.doc.variables.items() if k.startswith(csv_name + "::")]
            self.doc.loop_count = max(1, max(lengths) if lengths else 1)
        else:
            self.doc.loop_count = 999999
        if not self.engine.state.running:
            self.engine.run_async(self.doc, lambda s: self.action_var.set(f"Details: {s}"))
            self.run_btn.config(text="Pause", image=self.pause_icon, bg="orange")
            self.cancel_btn.config(state="normal")
        elif not self.engine.state.paused:
            self.engine.pause(); self.run_btn.config(text="Continue", image=self.play_icon, bg="dodger blue")
        else:
            self.engine.resume(); self.run_btn.config(text="Pause", image=self.pause_icon, bg="orange")

    def stop_macro(self) -> None:
        self.engine.stop(); self.cancel_btn.config(state="disabled"); self.run_btn.config(text="Run", image=self.play_icon, bg="lime green")
        self.state_var.set("State: Cancelled")

    def _tick_status(self) -> None:
        if self.engine.state.running:
            self.state_var.set(f"State: {'Paused' if self.engine.state.paused else 'Running'}")
            self.task_info_var.set(f"Current Task: {self.engine.state.current_task_index + 1}/{max(1, len(self.doc.tasks))}   |   Current Loop: {self.engine.state.current_loop + 1}/{self.doc.loop_count}")
            self.selected_index = self.engine.state.current_task_index
            self._refresh_task_rows()
        else:
            if "Cancelled" not in self.state_var.get():
                self.state_var.set("State: Off")
            self.cancel_btn.config(state="disabled")
            self.run_btn.config(text="Run", image=self.play_icon, bg="lime green")
        self.root.after(300, self._tick_status)

    def open_settings_window(self, section: str = "all") -> None:
        w = tk.Toplevel(self.root); w.title("Settings"); w.configure(bg="white")
        tk.Label(w, text="Macro Settings", bg="white", font=("Yu Gothic UI Semibold", 13)).pack(anchor="w", padx=10, pady=10)
        sections = ["CSV Handling Settings", "Loop Behaviour Settings", "Safety / Environment Options", "Advanced Settings"]
        for s in sections:
            f = tk.LabelFrame(w, text=s, bg="white"); f.pack(fill=tk.X, padx=10, pady=4)
            tk.Label(f, text="Configuration options placeholder", bg="white", fg="#555").pack(anchor="w", padx=8, pady=6)
        if section != "all":
            messagebox.showinfo("Settings", f"Quick link opened section: {section}", parent=w)


    def _variable_source_options(self) -> list[str]:
        options: list[str] = []
        for csv_name in self.csv_sources.keys():
            for key in self.doc.variables.keys():
                if key.startswith(csv_name + "::"):
                    options.append(f"{csv_name} -> {key.split('::',1)[1]}")
        return options

    def _load_csv_source(self, path: str) -> dict[str, list[str]]:
        name = Path(path).name
        rows = []
        with open(path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))
        if not rows:
            return {}
        first = rows[0]
        has_header = any(c.strip() and not c.strip().isdigit() for c in first)
        headers = first if has_header else [f"column{i+1}" for i in range(len(first))]
        data_rows = rows[1:] if has_header else rows
        out: dict[str, list[str]] = {f"{name}::{h}": [] for h in headers}
        for r in data_rows:
            for i, h in enumerate(headers):
                out[f"{name}::{h}"].append(r[i] if i < len(r) else "")
        return out

    def _refresh_loop_csv_options(self) -> None:
        values = list(self.csv_sources.keys())
        if hasattr(self, "loop_csv_combo"):
            self.loop_csv_combo.configure(values=values)
            if values and not self.loop_csv_var.get():
                self.loop_csv_var.set(values[0])

    def open_manage_csvs_window(self) -> None:
        w = tk.Toplevel(self.root); w.title("Manage CSVs"); w.configure(bg="white")
        container = tk.Frame(w, bg="white"); container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        list_frame = tk.Frame(container, bg="white"); list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lb = tk.Listbox(list_frame, relief="ridge", bg="white")
        lb.pack(fill=tk.BOTH, expand=True)
        for name in self.csv_sources.keys(): lb.insert(tk.END, name)

        btns = tk.Frame(container, bg="white"); btns.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        def do_import():
            path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")], parent=w)
            if not path: return
            key = Path(path).name
            self.csv_sources[key] = path
            self.doc.variables.update(self._load_csv_source(path))
            lb.insert(tk.END, key)
            self._refresh_loop_csv_options()
            self._mark_dirty()

        def do_delete():
            if not lb.curselection(): return
            i = lb.curselection()[0]; key = lb.get(i)
            lb.delete(i); self.csv_sources.pop(key, None)
            for k in [vk for vk in list(self.doc.variables.keys()) if vk.startswith(key + "::")]:
                self.doc.variables.pop(k, None)
            self._refresh_loop_csv_options()
            self._mark_dirty()

        def do_edit():
            messagebox.showinfo("Edit CSV", "CSV Editor not implemented yet.", parent=w)

        def do_configure():
            if not lb.curselection(): return
            self.open_configure_csv_window(lb.get(lb.curselection()[0]))

        for text, cmd in [("Import", do_import), ("Delete", do_delete), ("Edit CSV", do_edit), ("Configure CSV", do_configure)]:
            tk.Button(btns, text=text, command=cmd, relief="ridge", bg="white", width=14).pack(pady=3)

    def open_configure_csv_window(self, csv_name: str) -> None:
        w = tk.Toplevel(self.root); w.title(f"Configure CSV - {csv_name}"); w.configure(bg="white")
        tk.Label(w, text=csv_name, bg="white", font=("Yu Gothic UI Semibold", 12)).pack(anchor="w", padx=10, pady=8)

        preview = ttk.Treeview(w, columns=("c1", "c2", "c3", "c4", "c5"), show="headings", height=6)
        for i in range(1, 6):
            preview.heading(f"c{i}", text=f"Col {i}")
            preview.column(f"c{i}", width=120)
        preview.pack(fill=tk.X, padx=10)

        path = self.csv_sources.get(csv_name)
        if path and Path(path).exists():
            with open(path, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                for ridx, row in enumerate(reader):
                    if ridx >= 6: break
                    values = row[:5] + [""] * (5 - len(row[:5]))
                    preview.insert("", tk.END, values=values)

        cfg = tk.LabelFrame(w, text="Basic Config", bg="white"); cfg.pack(fill=tk.X, padx=10, pady=10)
        v1 = tk.BooleanVar(value=True); v2 = tk.BooleanVar(value=False); v3 = tk.BooleanVar(value=False)
        tk.Checkbutton(cfg, text="First Row Contains Headers", variable=v1, bg="white").pack(anchor="w")
        tk.Label(cfg, text="Default delimiter", bg="white").pack(anchor="w")
        delim = tk.StringVar(value=",")
        tk.Entry(cfg, textvariable=delim, width=4, relief="ridge", bg="white").pack(anchor="w")
        tk.Button(cfg, text="Save", relief="ridge", bg="white", command=lambda: messagebox.showinfo("Saved", "CSV configuration saved.", parent=w)).pack(anchor="w", pady=6)
        tk.Checkbutton(cfg, text="Auto-trim whitespace", variable=v2, bg="white").pack(anchor="w")
        tk.Checkbutton(cfg, text="Auto-skip blank rows", variable=v3, bg="white").pack(anchor="w")

    def new_macro(self) -> None:
        self.doc = MacroDocument(); self.file_path = None; self.selected_index = None; self.loop_var.set(1); self.csv_sources = {}; self.dirty = False
        self._refresh_loop_csv_options(); self._refresh_task_rows(); self._refresh_macro_info()

    def save_macro(self) -> None:
        if self.loop_mode_var.get() == "count":
            self.doc.loop_count = max(1, int(self.loop_var.get()))
        elif self.loop_mode_var.get() == "csv_end" and self.loop_csv_var.get() in self.csv_sources:
            csv_name = self.loop_csv_var.get()
            lengths = [len(v) for k,v in self.doc.variables.items() if k.startswith(csv_name + "::")]
            self.doc.loop_count = max(1, max(lengths) if lengths else 1)
        else:
            self.doc.loop_count = 999999
        if not self.file_path: self.save_as_macro(); return
        self.doc.csv_sources = dict(self.csv_sources)
        MacroStorage.save(self.file_path, self.doc); self.dirty = False; self._refresh_macro_info()

    def save_as_macro(self) -> None:
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Macro JSON", "*.json")])
        if not p: return
        self.file_path = p
        self.doc.csv_sources = dict(self.csv_sources)
        MacroStorage.save(p, self.doc); self.dirty = False; self._refresh_macro_info()

    def open_macro(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Macro JSON", "*.json")])
        if not p: return
        self.file_path = p; self.doc = MacroStorage.load(p); self.csv_sources = dict(getattr(self.doc, "csv_sources", {})); self.loop_var.set(self.doc.loop_count); self.selected_index = None; self.dirty = False
        self._refresh_loop_csv_options(); self._refresh_task_rows(); self._refresh_macro_info()

    def run(self) -> None:
        self.root.mainloop()
