# -*- coding: utf-8 -*-
# ============================================================
# BrailleSpeakGR Benchmark Analyzer GUI v2.0
#
# Put this file in the same folder as:
#   benchmark_analyzer_v2.py
#   Benchmark_Master_BrailleSpeakGR.csv
#
# Run:
#   python gui_benchmark_analyzer_v2.py
# ============================================================

from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


class BenchmarkAnalyzerGUIV2:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("BrailleSpeakGR Benchmark Analyzer v2.0")

        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry("1300x850")

        self.folder = Path(__file__).resolve().parent
        self.script_path = self.folder / "benchmark_analyzer_v2.py"
        self.csv_path = self.folder / "Benchmark_Master_BrailleSpeakGR.csv"
        self.output_dir = self.folder / "benchmark_outputs_v2"

        self.root.configure(bg="#f0f0f0")

        self._build_ui()

    def _build_ui(self):
        title = tk.Label(
            self.root,
            text="BrailleSpeakGR Benchmark Analyzer v2.0",
            font=("Arial", 24, "bold"),
            fg="darkblue",
            bg="#f0f0f0"
        )
        title.pack(pady=10)

        subtitle = tk.Label(
            self.root,
            text="Professional OCR/Braille/TTS Benchmark Engine - Tables, Statistics, 16 Figures, Word, PDF, Excel",
            font=("Arial", 13, "bold"),
            fg="black",
            bg="#f0f0f0"
        )
        subtitle.pack(pady=4)

        author = tk.Label(
            self.root,
            text="BrailleSpeakGR Research Prototype | Alexandros Angelis | University of Ioannina",
            font=("Arial", 11),
            fg="gray20",
            bg="#f0f0f0"
        )
        author.pack(pady=2)

        self.file_label = tk.Label(
            self.root,
            text=f"CSV: {self.csv_path}",
            font=("Arial", 10),
            fg="black",
            bg="#f0f0f0",
            wraplength=1200
        )
        self.file_label.pack(pady=6)

        btn_frame = tk.Frame(self.root, bg="#f0f0f0")
        btn_frame.pack(pady=8)

        self.select_btn = tk.Button(
            btn_frame,
            text="SELECT CSV",
            command=self.select_csv,
            font=("Arial", 12, "bold"),
            width=16,
            height=2,
            bg="#555555",
            fg="white"
        )
        self.select_btn.grid(row=0, column=0, padx=6)

        self.run_btn = tk.Button(
            btn_frame,
            text="RUN BENCHMARK v2.0",
            command=self.run_analysis_thread,
            font=("Arial", 12, "bold"),
            width=24,
            height=2,
            bg="#003366",
            fg="white"
        )
        self.run_btn.grid(row=0, column=1, padx=6)

        self.open_btn = tk.Button(
            btn_frame,
            text="OPEN OUTPUT FOLDER",
            command=self.open_output_folder,
            font=("Arial", 12, "bold"),
            width=22,
            height=2,
            bg="#006633",
            fg="white"
        )
        self.open_btn.grid(row=0, column=2, padx=6)

        self.open_word_btn = tk.Button(
            btn_frame,
            text="OPEN WORD REPORT",
            command=self.open_word_report,
            font=("Arial", 12, "bold"),
            width=20,
            height=2,
            bg="#6A3D9A",
            fg="white"
        )
        self.open_word_btn.grid(row=0, column=3, padx=6)

        self.open_pdf_btn = tk.Button(
            btn_frame,
            text="OPEN PDF REPORT",
            command=self.open_pdf_report,
            font=("Arial", 12, "bold"),
            width=18,
            height=2,
            bg="#8B4513",
            fg="white"
        )
        self.open_pdf_btn.grid(row=0, column=4, padx=6)

        self.clear_btn = tk.Button(
            btn_frame,
            text="CLEAR",
            command=self.clear_output,
            font=("Arial", 12, "bold"),
            width=12,
            height=2,
            bg="#990000",
            fg="white"
        )
        self.clear_btn.grid(row=0, column=5, padx=6)

        self.status = tk.Label(
            self.root,
            text="Status: Ready",
            font=("Arial", 12, "bold"),
            fg="green",
            bg="#f0f0f0"
        )
        self.status.pack(pady=4)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=850)
        self.progress.pack(pady=5)

        frame = tk.Frame(self.root)
        frame.pack(padx=15, pady=8, fill="both", expand=True)

        self.output = tk.Text(
            frame,
            width=160,
            height=38,
            font=("Consolas", 10),
            wrap="word",
            bg="black",
            fg="#00FF66"
        )
        self.output.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(frame, command=self.output.yview)
        scrollbar.pack(side="right", fill="y")
        self.output.configure(yscrollcommand=scrollbar.set)

        footer = tk.Label(
            self.root,
            text="Outputs: CSV summaries | article text | 16 PNG figures | DOCX | PDF | XLSX | JSON",
            font=("Arial", 10, "italic"),
            fg="gray",
            bg="#f0f0f0"
        )
        footer.pack(pady=5)

    def select_csv(self):
        selected = filedialog.askopenfilename(
            title="Select Benchmark Master CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if selected:
            self.csv_path = Path(selected)
            self.file_label.configure(text=f"CSV: {self.csv_path}")

    def clear_output(self):
        self.output.delete("1.0", tk.END)
        self.status.configure(text="Status: Ready", fg="green")

    def run_analysis_thread(self):
        t = threading.Thread(target=self.run_analysis)
        t.daemon = True
        t.start()

    def run_analysis(self):
        if not self.script_path.exists():
            messagebox.showerror("Missing script", f"Cannot find:\n{self.script_path}")
            return
        if not self.csv_path.exists():
            messagebox.showerror("Missing CSV", f"Cannot find:\n{self.csv_path}")
            return

        self.run_btn.configure(state="disabled")
        self.status.configure(text="Status: Running benchmark analysis v2.0...", fg="orange")
        self.progress.start(10)
        self.output.delete("1.0", tk.END)

        self.output.insert(tk.END, "Running BrailleSpeakGR Benchmark Analyzer v2.0...\n\n")
        self.output.insert(tk.END, f"Script: {self.script_path}\n")
        self.output.insert(tk.END, f"CSV:    {self.csv_path}\n\n")
        self.output.see(tk.END)

        try:
            process = subprocess.Popen(
                [sys.executable, str(self.script_path), str(self.csv_path)],
                cwd=str(self.folder),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )

            for line in process.stdout:
                self.output.insert(tk.END, line)
                self.output.see(tk.END)
                self.root.update_idletasks()

            process.wait()

            if process.returncode == 0:
                self.status.configure(text="Status: Completed successfully", fg="green")
                self.output.insert(tk.END, "\nBenchmark Analyzer v2.0 completed successfully.\n")
            else:
                self.status.configure(text="Status: Execution error", fg="red")
                self.output.insert(tk.END, "\nERROR: Benchmark Analyzer v2.0 failed.\n")

        except Exception as e:
            self.status.configure(text="Status: Error", fg="red")
            messagebox.showerror("Execution error", str(e))

        finally:
            self.progress.stop()
            self.run_btn.configure(state="normal")

    def open_output_folder(self):
        self.output_dir.mkdir(exist_ok=True)
        self._open_path(self.output_dir)

    def open_word_report(self):
        path = self.output_dir / "04_reports" / "benchmark_report.docx"
        if not path.exists():
            messagebox.showwarning("Missing report", "Run the benchmark first. Word report not found.")
            return
        self._open_path(path)

    def open_pdf_report(self):
        path = self.output_dir / "04_reports" / "benchmark_report.pdf"
        if not path.exists():
            messagebox.showwarning("Missing report", "Run the benchmark first. PDF report not found.")
            return
        self._open_path(path)

    def _open_path(self, path: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Open error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = BenchmarkAnalyzerGUIV2(root)
    root.mainloop()
