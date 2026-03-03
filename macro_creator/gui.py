from __future__ import annotations

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
    "Screenshot": {"folder": "screenshots", "comment": ""},
}

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
    "config": "Images/Menu/Macro/config.png",
    "runsett": "Images/Menu/Macro/runsett-ico.png",
    "varsett": "Images/Menu/Macro/variable-setting.png",
    "help": "Images/Menu/Other/help.png",
    "about": "Images/Menu/Other/about.png",
}


class TaskEditorDialog:
    def __init__(
        self,
        parent: tk.Tk,
        task_type: str,
        initial: Dict[str, str],
        variable_sources: list[str] | None = None,
        icon=None,
    ) -> None:
        self.parent = parent
        self.task_type = task_type
        self.variable_sources = variable_sources or []
        self.window = tk.Toplevel(parent)
        self.window.title(f"{task_type} Task")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg="white")

        self.result: Dict[str, str] | None = None
        self.entries: Dict[str, tk.StringVar] = {}

        row = 0
        if icon is not None:
            tk.Label(self.window, image=icon, bg="white").grid(row=row, column=0, padx=8, pady=8, sticky="w")
            tk.Label(self.window, text=f"{task_type} Task", bg="white", font=("Yu Gothic UI Semibold", 12)).grid(
                row=row, column=1, columnspan=3, sticky="w"
            )
            row += 1

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
                if key != "comment":
                    self._entry_row(row, key, key)
                    row += 1

        self._entry_row(row, "Comment (optional)", "comment")
        comment_entry = self.window.grid_slaves(row=row, column=1)[0]
        CreateToolTip(comment_entry, "Optional label to help identify key tasks in the task list.")
        row += 1

        for key, value in initial.items():
            if key in self.entries and value != "":
                self.entries[key].set(str(value))

        btns = tk.Frame(self.window, bg="white")
        btns.grid(row=row, column=0, columnspan=4, pady=8)
        save_btn = tk.Button(btns, text="Save", command=self._save, relief="ridge", bg="white")
        save_btn.pack(side=tk.LEFT, padx=4)
        cancel_btn = tk.Button(btns, text="Cancel", command=self.window.destroy, relief="ridge", bg="white")
        cancel_btn.pack(side=tk.LEFT, padx=4)
        CreateToolTip(save_btn, "Save task")
        CreateToolTip(cancel_btn, "Discard and close")

        self._center_on_parent()

    def _center_on_parent(self) -> None:
        self.window.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.window.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{max(0,x)}+{max(0,y)}")

    def _entry_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar())
        tk.Entry(self.window, textvariable=var, width=42, relief="ridge", bg="white").grid(
            row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4
        )

    def _spinbox_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="1"))
        tk.Spinbox(
            self.window,
            from_=0,
            to=999999,
            increment=0.1,
            textvariable=var,
            width=15,
            relief="ridge",
            bg="white",
        ).grid(row=row, column=1, sticky=tk.W, padx=8, pady=4)

    def _dropdown_row(self, row: int, label: str, key: str, options: list[str]) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value=options[0]))
        ttk.Combobox(self.window, textvariable=var, values=options, state="readonly", width=39).grid(
            row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4
        )

    def _variable_source_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        options = self.variable_sources if self.variable_sources else ["No Sources"]
        var = self.entries.setdefault(key, tk.StringVar(value=options[0]))
        state = "readonly" if self.variable_sources else "disabled"
        c = ttk.Combobox(self.window, textvariable=var, values=options, state=state, width=39)
        c.grid(row=row, column=1, columnspan=3, sticky=tk.W, padx=8, pady=4)
        if not self.variable_sources:
            CreateToolTip(c, "Import CSV variables to enable selecting a source.")

    def _key_capture_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="enter"))
        tk.Entry(self.window, textvariable=var, width=24, relief="ridge", bg="white").grid(
            row=row, column=1, sticky=tk.W, padx=8, pady=4
        )
        btn = tk.Button(self.window, text="⌨", command=lambda: self._capture_key(var), relief="ridge", bg="white", width=3)
        btn.grid(row=row, column=2, sticky=tk.W, padx=8, pady=4)
        CreateToolTip(btn, "Capture key press.")

    def _position_row(self, row: int, label: str, x_key: str, y_key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        x_var = self.entries.setdefault(x_key, tk.StringVar(value="0"))
        y_var = self.entries.setdefault(y_key, tk.StringVar(value="0"))
        tk.Entry(self.window, textvariable=x_var, width=10, relief="ridge", bg="white").grid(
            row=row, column=1, sticky=tk.W, padx=(8, 2), pady=4
        )
        tk.Entry(self.window, textvariable=y_var, width=10, relief="ridge", bg="white").grid(
            row=row, column=2, sticky=tk.W, padx=(2, 8), pady=4
        )
        btn = tk.Button(self.window, text="📍", command=lambda: self._capture_position(x_var, y_var), relief="ridge", bg="white", width=3)
        btn.grid(row=row, column=3, sticky=tk.W, padx=8, pady=4)
        CreateToolTip(btn, "Move mouse then press CTRL to capture position. Press ESC to cancel.")

    def _color_row(self, row: int, label: str, key: str) -> None:
        tk.Label(self.window, text=label, bg="white").grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
        var = self.entries.setdefault(key, tk.StringVar(value="255,255,255"))
        tk.Entry(self.window, textvariable=var, width=26, relief="ridge", bg="white").grid(
            row=row, column=1, columnspan=2, sticky=tk.W, padx=8, pady=4
        )
        btn = tk.Button(self.window, text="🧪", command=lambda: self._capture_color(var), relief="ridge", bg="white", width=3)
        btn.grid(row=row, column=3, sticky=tk.W, padx=8, pady=4)
        CreateToolTip(btn, "Move mouse then press CTRL to capture RGB at cursor. Press ESC to cancel.")

    def _capture_key(self, key_var: tk.StringVar) -> None:
        captured = _capture_keypress(self.window)
        if captured:
            key_var.set(captured)

    def _capture_position(self, x_var: tk.StringVar, y_var: tk.StringVar) -> None:
        point = _capture_with_hotkey("position")
        if point:
            x_var.set(str(point[0]))
            y_var.set(str(point[1]))

    def _capture_color(self, rgb_var: tk.StringVar) -> None:
        color = _capture_with_hotkey("color")
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
        else:
            result["value"] = event.keysym.lower()

    window.bind("<Key>", on_key)
    window.focus_force()
    while result["value"] is None and not result["cancel"]:
        window.update()
        time.sleep(0.03)
    window.unbind("<Key>")
    return result["value"]


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


class MacroCreatorApp:
    def __init__(self) -> None:
        self.doc = MacroDocument()
        self.engine = MacroEngine()
        self.file_path: str | None = None
        self.resume_loop = 0
        self.resume_task = 0
        self.selected_index: int | None = None
        self.icons: Dict[str, Any] = {}
        self.menu_icons: Dict[str, Any] = {}

        self.root = tk.Tk()
        self.root.title("MacroCreator")
        self.root.geometry("1120x720")
        self.root.configure(bg="black")

        self.status = tk.StringVar(value="State: Off")
        self.loop_var = tk.IntVar(value=1)

        self._load_images()
        self._build_menu()
        self._build_layout()
        self._refresh_task_rows()
        self._tick_status()

    def _load_image(self, rel_path: str, size: tuple[int, int] = (18, 18)):
        if Image is None or ImageTk is None:
            return None
        p = PROTO_DIR / rel_path
        if not p.exists():
            return None
        img = Image.open(p).resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def _load_images(self) -> None:
        for task, rel in TASK_ICON_MAP.items():
            ico = self._load_image(rel, (20, 20))
            if ico:
                self.icons[task] = ico
        self.play_icon = self._load_image("Images/play-ico.png", (20, 20))
        self.pause_icon = self._load_image("Images/pause-ico.png", (20, 20))
        for key, rel in MENU_ICON_MAP.items():
            ico = self._load_image(rel, (18, 18))
            if ico:
                self.menu_icons[key] = ico

    def _build_menu(self) -> None:
        m = tk.Menu(self.root)
        file_menu = tk.Menu(m, tearoff=0)
        file_menu.add_command(label="New", command=self.new_macro, image=self.menu_icons.get("new"), compound=tk.LEFT)
        file_menu.add_command(label="Open", command=self.open_macro, image=self.menu_icons.get("open"), compound=tk.LEFT)
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=self.save_macro, image=self.menu_icons.get("save"), compound=tk.LEFT)
        file_menu.add_command(label="Save As", command=self.save_as_macro, image=self.menu_icons.get("saveas"), compound=tk.LEFT)
        file_menu.add_separator()
        file_menu.add_command(label="Import CSV Variables", command=self.import_csv, image=self.menu_icons.get("varsett"), compound=tk.LEFT)
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
        create_menu = tk.Menu(macro_menu, tearoff=0)
        for task in TASK_FIELDS:
            create_menu.add_command(label=task, command=lambda t=task: self.create_task_button(t), image=self.icons.get(task), compound=tk.LEFT)
        macro_menu.add_cascade(label="Create Task", menu=create_menu, image=self.menu_icons.get("task"), compound=tk.LEFT)
        m.add_cascade(label="Macro", menu=macro_menu)

        other_menu = tk.Menu(m, tearoff=0)
        other_menu.add_command(label="Help", command=lambda: messagebox.showinfo("Help", "See tooltips on controls."), image=self.menu_icons.get("help"), compound=tk.LEFT)
        other_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "MacroCreator"), image=self.menu_icons.get("about"), compound=tk.LEFT)
        m.add_cascade(label="Other", menu=other_menu)
        self.root.config(menu=m)

    def _build_layout(self) -> None:
        outer = tk.Frame(self.root, bg="black")
        outer.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        title = tk.Label(outer, text="Macro Tasks", bg="white", fg="black", font=("Yu Gothic UI Semibold", 16), relief="ridge")
        title.pack(fill=tk.X, pady=(0, 4))

        state_box = tk.Label(outer, textvariable=self.status, anchor="w", bg="white", fg="black", relief="ridge")
        state_box.pack(fill=tk.X, pady=(0, 4))

        list_container = tk.Frame(outer, bg="white", bd=1, relief="sunken")
        list_container.pack(fill=tk.BOTH, expand=True)

        self.task_canvas = tk.Canvas(list_container, bg="white", highlightthickness=0)
        yscroll = ttk.Scrollbar(list_container, orient="vertical", command=self.task_canvas.yview)
        self.task_inner = tk.Frame(self.task_canvas, bg="white")
        self.task_window = self.task_canvas.create_window((0, 0), window=self.task_inner, anchor="nw")
        self.task_inner.bind("<Configure>", lambda _e: self.task_canvas.configure(scrollregion=self.task_canvas.bbox("all")))
        self.task_canvas.bind("<Configure>", self._on_canvas_resize)
        self.task_canvas.configure(yscrollcommand=yscroll.set)
        self.task_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        bottom = tk.Frame(outer, bg="white", relief="ridge", bd=1)
        bottom.pack(fill=tk.X, pady=(4, 0))

        task_box = tk.LabelFrame(bottom, text="Create a Task", bg="white")
        task_box.pack(fill=tk.X, padx=4, pady=4)
        for i, task in enumerate(TASK_FIELDS.keys()):
            btn = tk.Button(
                task_box,
                text=task,
                image=self.icons.get(task),
                compound="left",
                command=lambda t=task: self.create_task_button(t),
                relief="ridge",
                bg="white",
            )
            btn.grid(row=0, column=i, padx=2, pady=2, sticky="ew")
            CreateToolTip(btn, f"Create {task} task")

        ctrl = tk.Frame(bottom, bg="white")
        ctrl.pack(fill=tk.X, padx=4, pady=(0, 4))
        self.run_btn = tk.Button(
            ctrl,
            text="Run",
            image=self.play_icon,
            compound="left",
            bg="lime green",
            fg="white",
            command=self.run_pause_continue,
            relief="ridge",
        )
        self.run_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
        CreateToolTip(self.run_btn, "Run / Pause / Continue macro")

        self.cancel_btn = tk.Button(
            ctrl,
            text="Cancel",
            bg="red",
            fg="white",
            width=10,
            command=self.stop_macro,
            state="disabled",
            relief="ridge",
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=(6, 0))
        CreateToolTip(self.cancel_btn, "Cancel running macro")

        loop_box = tk.Frame(bottom, bg="white")
        loop_box.pack(fill=tk.X, padx=4, pady=(0, 4))
        tk.Label(loop_box, text="Loop count:", bg="white").pack(side=tk.LEFT)
        ttk.Spinbox(loop_box, from_=1, to=999999, textvariable=self.loop_var, width=10).pack(side=tk.LEFT, padx=4)

    def _on_canvas_resize(self, event) -> None:
        self.task_canvas.itemconfig(self.task_window, width=event.width)

    def _task_summary(self, task: Task) -> str:
        p = task.params
        if task.task_type == "Click":
            return f"{p.get('click_type','Single')} {p.get('button','Left')} | Position: {p.get('x','0')},{p.get('y','0')}"
        if task.task_type == "Condition":
            return f"Position: {p.get('x','0')},{p.get('y','0')} | RGB: {p.get('rgb','')}"
        if task.task_type == "Wait":
            return f"Seconds: {p.get('seconds','1')}"
        filtered = {k: v for k, v in p.items() if k != "comment"}
        return ", ".join(f"{k}:{v}" for k, v in filtered.items())

    def _select_task(self, idx: int) -> None:
        self.selected_index = idx
        self._refresh_task_rows()

    def _refresh_task_rows(self) -> None:
        for c in self.task_inner.winfo_children():
            c.destroy()
        if not self.doc.tasks:
            tk.Label(self.task_inner, text="There are currently no tasks, use the options below to start now.", bg="white").pack(anchor="w", padx=8, pady=8)
            return

        for idx, task in enumerate(self.doc.tasks):
            bg = "#dff0ff" if idx == self.selected_index else "white"
            row = tk.Frame(self.task_inner, bg=bg, bd=1, relief="ridge")
            row.pack(fill=tk.X, padx=2, pady=2)
            row.bind("<Button-1>", lambda _e, i=idx: self._select_task(i))

            left = tk.Frame(row, bg=bg)
            left.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(left, image=self.icons.get(task.task_type), bg=bg).pack(side=tk.LEFT, padx=6)
            txt = tk.Label(left, text=f"[{idx}] {task.task_type} - {self._task_summary(task)}", bg=bg, anchor="w")
            txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
            txt.bind("<Double-1>", lambda _e, i=idx: self.edit_task_at(i))

            comment = str(task.params.get("comment", "")).strip()
            comment_lbl = tk.Label(row, text=comment if comment else "", bg=bg, fg="#444", width=24, anchor="w")
            comment_lbl.pack(side=tk.LEFT, padx=(4, 2))
            CreateToolTip(comment_lbl, "Optional comment label for this task")

            btn_frame = tk.Frame(row, bg=bg)
            btn_frame.pack(side=tk.RIGHT, padx=2)
            for label, cmd, tip in [
                ("↑", lambda i=idx: self.move_up_at(i), "Move up"),
                ("↓", lambda i=idx: self.move_down_at(i), "Move down"),
                ("✎", lambda i=idx: self.edit_task_at(i), "Edit"),
                ("⧉", lambda i=idx: self.duplicate_at(i), "Duplicate"),
                ("✖", lambda i=idx: self.delete_at(i), "Delete"),
            ]:
                b = tk.Button(btn_frame, text=label, width=2, command=cmd, relief="ridge", bg="white")
                b.pack(side=tk.LEFT, padx=1, pady=2)
                CreateToolTip(b, tip)

    def _open_task_editor(self, task_type: str, initial: Dict[str, str]) -> Dict[str, str] | None:
        dialog = TaskEditorDialog(
            self.root,
            task_type,
            initial,
            variable_sources=list(self.doc.variables.keys()),
            icon=self.icons.get(task_type),
        )
        self.root.wait_window(dialog.window)
        return dialog.result

    def create_task_button(self, task_type: str) -> None:
        params = self._open_task_editor(task_type, TASK_FIELDS[task_type].copy())
        if params is None:
            return
        self.doc.tasks.append(Task(task_type, params))
        self._refresh_task_rows()

    def edit_task_at(self, idx: int) -> None:
        task = self.doc.tasks[idx]
        params = self._open_task_editor(task.task_type, dict(task.params))
        if params is None:
            return
        task.params = params
        self._refresh_task_rows()

    def duplicate_at(self, idx: int) -> None:
        t = self.doc.tasks[idx]
        self.doc.tasks.insert(idx + 1, Task(t.task_type, dict(t.params)))
        self._refresh_task_rows()

    def move_up_at(self, idx: int) -> None:
        if idx == 0:
            return
        self.doc.tasks[idx - 1], self.doc.tasks[idx] = self.doc.tasks[idx], self.doc.tasks[idx - 1]
        self.selected_index = idx - 1
        self._refresh_task_rows()

    def move_down_at(self, idx: int) -> None:
        if idx >= len(self.doc.tasks) - 1:
            return
        self.doc.tasks[idx + 1], self.doc.tasks[idx] = self.doc.tasks[idx], self.doc.tasks[idx + 1]
        self.selected_index = idx + 1
        self._refresh_task_rows()

    def delete_at(self, idx: int) -> None:
        del self.doc.tasks[idx]
        self.selected_index = None
        self._refresh_task_rows()

    def edit_selected_task(self) -> None:
        if self.selected_index is not None:
            self.edit_task_at(self.selected_index)

    def duplicate_selected(self) -> None:
        if self.selected_index is not None:
            self.duplicate_at(self.selected_index)

    def delete_selected(self) -> None:
        if self.selected_index is not None:
            self.delete_at(self.selected_index)

    def run_pause_continue(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        if not self.engine.state.running:
            self.engine.run_async(self.doc, self.status.set)
            self.cancel_btn.config(state="normal")
            self.run_btn.config(text="Pause", image=self.pause_icon, bg="orange")
        elif not self.engine.state.paused:
            self.engine.pause()
            self.run_btn.config(text="Continue", image=self.play_icon, bg="dodger blue")
        else:
            self.engine.resume()
            self.run_btn.config(text="Pause", image=self.pause_icon, bg="orange")

    def stop_macro(self) -> None:
        self.engine.stop()
        self.cancel_btn.config(state="disabled")
        self.run_btn.config(text="Run", image=self.play_icon, bg="lime green")
        self.status.set("State: Cancelled")

    def _tick_status(self) -> None:
        if self.engine.state.running:
            st = "Paused" if self.engine.state.paused else "Running"
            self.status.set(
                f"State: {st} | Task {self.engine.state.current_task_index + 1}/{len(self.doc.tasks)} | Loop {self.engine.state.current_loop + 1}/{self.doc.loop_count}"
            )
            self.resume_loop = self.engine.state.current_loop
            self.resume_task = self.engine.state.current_task_index
        else:
            if "Cancelled" not in self.status.get() and "Completed" not in self.status.get():
                self.status.set("State: Off")
            self.cancel_btn.config(state="disabled")
            self.run_btn.config(text="Run", image=self.play_icon, bg="lime green")
        self.root.after(200, self._tick_status)

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
        self.selected_index = None
        self._refresh_task_rows()

    def save_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        if not self.file_path:
            self.save_as_macro()
            return
        MacroStorage.save(self.file_path, self.doc)

    def save_as_macro(self) -> None:
        self.doc.loop_count = max(1, int(self.loop_var.get()))
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Macro JSON", "*.json")])
        if not path:
            return
        self.file_path = path
        MacroStorage.save(path, self.doc)

    def open_macro(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Macro JSON", "*.json")])
        if not path:
            return
        self.file_path = path
        self.doc = MacroStorage.load(path)
        self.loop_var.set(self.doc.loop_count)
        self.selected_index = None
        self._refresh_task_rows()

    def run(self) -> None:
        self.root.mainloop()
