import json
import tkinter as tk
from tkinter import filedialog, ttk
from typing import Callable


class RunLogPanel(ttk.Frame):
    def __init__(self, master, *, title: str = "Run Log"):
        super().__init__(master, padding=8)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=title, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="Clear", command=self.clear).grid(row=0, column=1, padx=4)
        ttk.Button(header, text="Save", command=self.save_to_file).grid(row=0, column=2, padx=4)

        self.text = tk.Text(self, height=8, state="disabled", wrap="word")
        self.text.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

    def append(self, line: str) -> None:
        self.text.configure(state="normal")
        self.text.insert(tk.END, line + "\n")
        self.text.see(tk.END)
        self.text.configure(state="disabled")

    def clear(self) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", tk.END)
        self.text.configure(state="disabled")

    def save_to_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save run log",
            defaultextension=".log",
            filetypes=[("Log", "*.log"), ("Text", "*.txt"), ("All", "*.*")],
        )
        if not path:
            return
        data = self.text.get("1.0", tk.END)
        with open(path, "w", encoding="utf-8") as f:
            f.write(data)


class StepTable(ttk.Frame):
    def __init__(self, master, *, on_select: Callable[[int | None], None]):
        super().__init__(master)
        self.on_select = on_select

        cols = ("index", "enabled", "type", "target", "value", "wait", "description")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=20)
        self.tree.heading("index", text="#")
        self.tree.heading("enabled", text="Enabled")
        self.tree.heading("type", text="Type")
        self.tree.heading("target", text="Target/Flow")
        self.tree.heading("value", text="Value")
        self.tree.heading("wait", text="Wait")
        self.tree.heading("description", text="Description")

        self.tree.column("index", width=40, anchor="center")
        self.tree.column("enabled", width=70, anchor="center")
        self.tree.column("type", width=160)
        self.tree.column("target", width=160)
        self.tree.column("value", width=220)
        self.tree.column("wait", width=70, anchor="center")
        self.tree.column("description", width=260)

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def set_steps(self, steps: list[dict]) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, step in enumerate(steps, start=1):
            step_type = str(step.get("type", ""))
            target = step.get("target", "") or step.get("flow", "")
            value = step.get("value", "") or step.get("path", "") or step.get("key", "")
            wait = step.get("seconds", "") if "seconds" in step else ""
            desc = step.get("description", "") or step.get("text", "")

            self.tree.insert(
                "",
                tk.END,
                iid=str(i - 1),
                values=(
                    i,
                    "Yes" if step.get("enabled", True) else "No",
                    step_type,
                    str(target),
                    str(value),
                    str(wait),
                    str(desc),
                ),
            )

    def selected_index(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _on_select(self, _event=None):
        self.on_select(self.selected_index())


class StepInspector(ttk.Frame):
    def __init__(self, master, *, on_apply: Callable[[dict], None]):
        super().__init__(master, padding=8)
        self.on_apply = on_apply

        self.current_step: dict | None = None

        ttk.Label(self, text="Step Inspector", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.form = ttk.Frame(self)
        self.form.pack(fill="x", pady=(8, 0))

        self.step_type = tk.StringVar(value="")
        self.enabled = tk.BooleanVar(value=True)
        self.target = tk.StringVar(value="")
        self.value = tk.StringVar(value="")
        self.seconds = tk.StringVar(value="")
        self.description = tk.StringVar(value="")

        self._row("Type", self.step_type)
        ttk.Checkbutton(self.form, text="Enabled", variable=self.enabled).pack(fill="x", pady=4)
        self._row("Target/Flow", self.target)
        self._row("Value/Path/Key", self.value)
        self._row("Seconds", self.seconds)
        self._row("Description", self.description)

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="Apply", command=self.apply).pack(side="left")

        ttk.Separator(self).pack(fill="x", pady=8)
        ttk.Label(self, text="Raw JSON", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.raw = tk.Text(self, height=18)
        self.raw.pack(fill="both", expand=True)

    def _row(self, label: str, var: tk.StringVar) -> None:
        row = ttk.Frame(self.form)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=14).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

    def load_step(self, step: dict | None) -> None:
        self.current_step = step
        if step is None:
            self.step_type.set("")
            self.enabled.set(True)
            self.target.set("")
            self.value.set("")
            self.seconds.set("")
            self.description.set("")
            self.raw.delete("1.0", tk.END)
            return

        self.step_type.set(str(step.get("type", "")))
        self.enabled.set(bool(step.get("enabled", True)))
        self.target.set(str(step.get("target", step.get("flow", ""))))

        value = step.get("value", "")
        if value == "":
            value = step.get("path", "")
        if value == "":
            value = step.get("key", "")
        self.value.set(str(value))

        self.seconds.set(str(step.get("seconds", "")))
        desc = step.get("description", "") or step.get("text", "")
        self.description.set(str(desc))

        self.raw.delete("1.0", tk.END)
        self.raw.insert("1.0", json.dumps(step, indent=2))

    def apply(self) -> None:
        if self.current_step is None:
            return

        # Start from raw JSON if valid so advanced fields survive edits.
        try:
            parsed = json.loads(self.raw.get("1.0", tk.END).strip() or "{}")
            if not isinstance(parsed, dict):
                parsed = {}
        except Exception:
            parsed = dict(self.current_step)

        parsed["type"] = self.step_type.get().strip()
        parsed["enabled"] = bool(self.enabled.get())

        target_like = self.target.get().strip()
        if parsed.get("type") == "run_flow":
            if target_like:
                parsed["flow"] = target_like
            elif "flow" in parsed:
                parsed.pop("flow", None)
            parsed.pop("target", None)
        else:
            if target_like:
                parsed["target"] = target_like
            elif "target" in parsed:
                parsed.pop("target", None)

        raw_value = self.value.get()
        if "seconds" in parsed or self.seconds.get().strip():
            try:
                parsed["seconds"] = float(self.seconds.get())
            except Exception:
                parsed["seconds"] = self.seconds.get().strip()

        if raw_value:
            # Common field mapping: leave one canonical field by step type.
            if parsed.get("type") in {"press_key"}:
                parsed["key"] = raw_value
                parsed.pop("value", None)
                parsed.pop("path", None)
            elif parsed.get("type") in {"assert_file_exists"}:
                parsed["path"] = raw_value
                parsed.pop("value", None)
                parsed.pop("key", None)
            else:
                parsed["value"] = raw_value
        else:
            parsed.pop("value", None)

        description = self.description.get().strip()
        if description:
            if parsed.get("type") == "comment":
                parsed["text"] = description
            else:
                parsed["description"] = description

        self.on_apply(parsed)
