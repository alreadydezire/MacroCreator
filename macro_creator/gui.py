from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict

from .engine import MacroEngine
from .models import MacroDocument, Task
from .storage import MacroStorage


TASK_FIELDS: Dict[str, Dict[str, str]] = {
    "Click": {"x": "0", "y": "0", "click_type": "single", "button": "left"},
    "String": {"text": ""},
    "Keypress": {"key": "enter"},
    "Hotkey": {"keys": "ctrl,c"},
    "Condition": {"x": "0", "y": "0", "rgb": "255,255,255", "timeout": "30"},
    "Wait": {"seconds": "1"},
    "Variable": {"name": "column_name"},
    "Screenshot": {"folder": "screenshots"},
}


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

    def add_task(self) -> None:
        kind = simpledialog.askstring("Task Type", "Enter task type:\nClick, String, Keypress, Hotkey, Condition, Wait, Variable, Screenshot")
        if not kind:
            return
        kind = kind.strip()
        if kind not in TASK_FIELDS:
            messagebox.showerror("Error", "Unsupported task type")
            return
        params = self._edit_params_dialog(kind, TASK_FIELDS[kind].copy())
        if params is None:
            return
        self.doc.tasks.append(Task(kind, params))
        self._refresh_list()

    def edit_selected_task(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        task = self.doc.tasks[idx]
        params = self._edit_params_dialog(task.task_type, dict(task.params))
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

    def _edit_params_dialog(self, kind: str, params: Dict[str, str]):
        win = tk.Toplevel(self.root)
        win.title(f"Edit {kind} Task")
        vars_map: Dict[str, tk.StringVar] = {}
        for row, (key, value) in enumerate(params.items()):
            ttk.Label(win, text=key).grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
            var = tk.StringVar(value=str(value))
            vars_map[key] = var
            ttk.Entry(win, textvariable=var, width=40).grid(row=row, column=1, padx=8, pady=4)

        out = {"ok": False}

        def save():
            out["ok"] = True
            win.destroy()

        ttk.Button(win, text="Save", command=save).grid(row=len(params), column=0, columnspan=2, pady=8)
        win.grab_set()
        self.root.wait_window(win)
        if not out["ok"]:
            return None
        return {k: v.get() for k, v in vars_map.items()}

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
