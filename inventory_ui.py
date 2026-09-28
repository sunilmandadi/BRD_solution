"""Tkinter UI for extracting BRD and solution slide content to separate JSON."""

import json
import tkinter as tk
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from copilot_prompts import build_copilot_prompt
from inventory_core import extract_brd, extract_pptx


class InventoryTab:
    def __init__(self, parent):
        self.root = parent.winfo_toplevel()
        self.brd_payload = None
        self.solution_payload = None
        self.brd_json = ""
        self.solution_json = ""
        self.brd_prompt = ""
        self.solution_prompt = ""

        inputs = ttk.LabelFrame(parent, text="Input files", padding=12)
        inputs.pack(fill="x", padx=12, pady=(12, 8))
        self.brd_path = tk.StringVar()
        self.solution_path = tk.StringVar()
        ttk.Label(inputs, text="BRD (.docx or .pdf):").grid(row=0, column=0, sticky="w")
        ttk.Entry(inputs, textvariable=self.brd_path).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(inputs, text="Browse…", command=self.browse_brd).grid(row=0, column=2)
        ttk.Label(inputs, text="Solution slides (.pptx):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(inputs, textvariable=self.solution_path).grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))
        ttk.Button(inputs, text="Browse…", command=self.browse_solution).grid(row=1, column=2, pady=(8, 0))
        inputs.columnconfigure(1, weight=1)

        actions = ttk.Frame(parent, padding=(12, 0, 12, 8))
        actions.pack(fill="x")
        ttk.Button(actions, text="Extract both to JSON", command=self.extract).pack(side="left")
        self.save_button = ttk.Button(actions, text="Save single bundle ZIP…", command=self.save_bundle, state="disabled")
        self.save_button.pack(side="right")

        copy_actions = ttk.Frame(parent, padding=(12, 0, 12, 8))
        copy_actions.pack(fill="x")
        self.brd_copy = ttk.Button(copy_actions, text="Copy BRD JSON", command=lambda: self.copy_json("brd"), state="disabled")
        self.brd_copy.pack(side="left", padx=(0, 6))
        self.solution_copy = ttk.Button(copy_actions, text="Copy solution JSON", command=lambda: self.copy_json("solution"), state="disabled")
        self.solution_copy.pack(side="left", padx=(0, 18))
        self.brd_prompt_copy = ttk.Button(copy_actions, text="Copy BRD Copilot prompt", command=lambda: self.copy_prompt("brd"), state="disabled")
        self.brd_prompt_copy.pack(side="left", padx=(0, 6))
        self.solution_prompt_copy = ttk.Button(copy_actions, text="Copy solution Copilot prompt", command=lambda: self.copy_prompt("solution"), state="disabled")
        self.solution_prompt_copy.pack(side="left")

        self.status = tk.StringVar(value="Choose both files, then click Extract both to JSON.")
        ttk.Label(parent, textvariable=self.status, padding=(12, 3)).pack(fill="x")

        self.tabs = ttk.Notebook(parent)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.brd_text = self._make_text_tab("BRD JSON")
        self.solution_text = self._make_text_tab("Solution JSON")
        self.brd_prompt_text = self._make_text_tab("BRD Copilot prompt")
        self.solution_prompt_text = self._make_text_tab("Solution Copilot prompt")

        # Editing either file path invalidates the previous output. This avoids
        # accidentally copying JSON from a different document.
        self.brd_path.trace_add("write", self._invalidate)
        self.solution_path.trace_add("write", self._invalidate)

    def _make_text_tab(self, title):
        frame = ttk.Frame(self.tabs, padding=5)
        self.tabs.add(frame, text=title)
        output = tk.Text(frame, wrap="none", font=("Consolas", 10))
        output.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=output.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=output.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        output.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        return output

    def _invalidate(self, *_args):
        self.brd_payload = None
        self.solution_payload = None
        self.brd_json = ""
        self.solution_json = ""
        self.brd_prompt = ""
        self.solution_prompt = ""
        self.brd_text.delete("1.0", "end")
        self.solution_text.delete("1.0", "end")
        self.brd_prompt_text.delete("1.0", "end")
        self.solution_prompt_text.delete("1.0", "end")
        self.brd_copy.configure(state="disabled")
        self.solution_copy.configure(state="disabled")
        self.brd_prompt_copy.configure(state="disabled")
        self.solution_prompt_copy.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.status.set("Input changed. Click Extract both to JSON.")

    def browse_brd(self):
        path = filedialog.askopenfilename(title="Select BRD", filetypes=[
            ("BRD document", "*.docx *.pdf"), ("All files", "*.*")])
        if path:
            self.brd_path.set(path)

    def browse_solution(self):
        path = filedialog.askopenfilename(title="Select solution slides", filetypes=[
            ("PowerPoint", "*.pptx"), ("All files", "*.*")])
        if path:
            self.solution_path.set(path)

    def extract(self):
        brd_path = self.brd_path.get().strip()
        solution_path = self.solution_path.get().strip()
        if not brd_path or not solution_path:
            messagebox.showwarning("Select both files", "Choose a BRD DOCX/PDF and a solution PPTX.")
            return
        self._invalidate()
        try:
            brd = extract_brd(brd_path)
            solution = extract_pptx(solution_path)
            brd_json = json.dumps(brd, indent=2, ensure_ascii=False)
            solution_json = json.dumps(solution, indent=2, ensure_ascii=False)
            brd_prompt = build_copilot_prompt("brd", brd)
            solution_prompt = build_copilot_prompt("solution", solution)
        except Exception as exc:
            messagebox.showerror("Extraction error", str(exc))
            return

        self.brd_payload, self.solution_payload = brd, solution
        self.brd_json, self.solution_json = brd_json, solution_json
        self.brd_prompt, self.solution_prompt = brd_prompt, solution_prompt
        self.brd_text.insert("1.0", brd_json)
        self.solution_text.insert("1.0", solution_json)
        self.brd_prompt_text.insert("1.0", brd_prompt)
        self.solution_prompt_text.insert("1.0", solution_prompt)
        self.brd_copy.configure(state="normal")
        self.solution_copy.configure(state="normal")
        self.brd_prompt_copy.configure(state="normal")
        self.solution_prompt_copy.configure(state="normal")
        self.save_button.configure(state="normal")
        self.tabs.select(0)
        warnings = len(brd["warnings"]) + len(solution["warnings"])
        self.status.set("BRD: %d blocks. Solution: %d blocks. %d extraction warning(s); review JSON before copying." %
                        (brd["block_count"], solution["block_count"], warnings))

    def copy_json(self, source):
        value = self.brd_json if source == "brd" else self.solution_json
        if not value:
            messagebox.showinfo("Extract first", "Extract both files before copying JSON.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.root.update()
        self.tabs.select(0 if source == "brd" else 1)
        self.status.set(("BRD" if source == "brd" else "Solution") + " JSON copied. Paste it separately into Copilot.")

    def copy_prompt(self, source):
        value = self.brd_prompt if source == "brd" else self.solution_prompt
        if not value:
            messagebox.showinfo("Extract first", "Extract both files before copying a Copilot prompt.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.root.update()
        self.tabs.select(2 if source == "brd" else 3)
        self.status.set(("BRD requirement" if source == "brd" else "Solution inventory") +
                        " Copilot prompt copied with its JSON. Paste it into Copilot.")

    def save_bundle(self):
        if not (self.brd_payload and self.solution_payload):
            messagebox.showinfo("Extract first", "Extract both files before saving the bundle.")
            return
        path = filedialog.asksaveasfilename(
            title="Save BRD and solution JSON bundle", initialfile="BRD_Solution_JSON_Bundle.zip",
            defaultextension=".zip", filetypes=[("ZIP archive", "*.zip")])
        if not path:
            return
        combined = {
            "schema_version": 1,
            "brd": self.brd_payload,
            "solution": self.solution_payload,
        }
        try:
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("brd_content.json", self.brd_json.encode("utf-8"))
                zf.writestr("solution_content.json", self.solution_json.encode("utf-8"))
                zf.writestr("brd_solution_bundle.json", json.dumps(combined, indent=2, ensure_ascii=False).encode("utf-8"))
                zf.writestr("brd_requirement_copilot_prompt.txt", self.brd_prompt.encode("utf-8"))
                zf.writestr("solution_inventory_copilot_prompt.txt", self.solution_prompt.encode("utf-8"))
            self.status.set("Saved both JSON outputs and Copilot prompts: " + str(Path(path)))
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))
