from __future__ import annotations

import json
import logging
import os
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

APP_NAME = "Raio-X de Investimentos"
DEFAULT_INPUT = "WORD"
DEFAULT_OUTPUT = "PDF"


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config() -> dict:
    """Load configuration defensively; never let malformed JSON break startup."""
    path = app_root() / "config.json"
    defaults = {"input_dir": DEFAULT_INPUT, "output_dir": DEFAULT_OUTPUT}
    try:
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8-sig"))
            if isinstance(data, dict):
                for key in defaults:
                    if isinstance(data.get(key), str) and data[key].strip():
                        defaults[key] = data[key].strip()
    except (OSError, ValueError, TypeError):
        pass
    return defaults


def setup_logging() -> Path:
    log_dir = app_root() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / f"novoraiox_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        filename=str(path), level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s", encoding="utf-8"
    )
    return path


def convert_docx_to_pdf(source: Path, destination: Path) -> None:
    """Export through installed Microsoft Word to preserve document layout."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(source.resolve()), ReadOnly=True, AddToRecentFiles=False,
            Visible=False, ConfirmConversions=False
        )
        # wdExportFormatPDF = 17
        document.ExportAsFixedFormat(
            OutputFileName=str(destination.resolve()),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                logging.exception("Could not close Word document: %s", source.name)
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                logging.exception("Could not quit Word for: %s", source.name)
        pythoncom.CoUninitialize()


def process_folder(input_dir: Path, output_dir: Path, progress_callback=None, log_callback=None):
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"A pasta de entrada não existe: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$"))
    if not files:
        raise ValueError(f"Nenhum arquivo .docx encontrado em {input_dir}")
    results = []
    for index, source in enumerate(files, 1):
        destination = output_dir / (source.stem + ".pdf")
        try:
            if destination.exists():
                destination.unlink()
            convert_docx_to_pdf(source, destination)
            if not destination.exists() or destination.stat().st_size < 100:
                raise RuntimeError("O Word não gerou um PDF válido.")
            msg = f"OK  | {source.name} → {destination.name}"
            logging.info(msg)
            results.append((source.name, True, msg))
        except Exception as exc:
            msg = f"ERRO | {source.name} | {exc}"
            logging.exception("Failed to convert %s", source)
            results.append((source.name, False, msg))
        if log_callback:
            log_callback(msg)
        if progress_callback:
            progress_callback(index, len(files))
    return results


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME + " — Word para PDF")
        self.geometry("760x540")
        self.minsize(650, 460)
        self.configure(bg="#f4f6f9")
        self.config_data = load_config()
        self.input_var = tk.StringVar(value=str(app_root() / self.config_data["input_dir"]))
        self.output_var = tk.StringVar(value=str(app_root() / self.config_data["output_dir"]))
        self.status_var = tk.StringVar(value="Pronto para processar documentos Word.")
        self._build()

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f4f6f9")
        style.configure("TLabel", background="#f4f6f9", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#17365d")
        style.configure("TButton", font=("Segoe UI", 10), padding=7)
        frame = ttk.Frame(self, padding=22)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Raio-X de Investimentos", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Conversor de documentos Word (.docx) para PDF, preservando o layout via Microsoft Word.").pack(anchor="w", pady=(5, 20))
        self._folder_row(frame, "Pasta com documentos Word", self.input_var, self.pick_input)
        self._folder_row(frame, "Pasta de saída dos PDFs", self.output_var, self.pick_output)
        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(20, 8))
        ttk.Label(frame, textvariable=self.status_var).pack(anchor="w", pady=(0, 8))
        self.log = tk.Text(frame, height=12, wrap="word", font=("Consolas", 9), bg="#ffffff", relief="solid", bd=1)
        self.log.pack(fill="both", expand=True)
        bottom = ttk.Frame(frame)
        bottom.pack(fill="x", pady=(12, 0))
        self.run_button = ttk.Button(bottom, text="Converter documentos", command=self.start)
        self.run_button.pack(side="left")
        ttk.Button(bottom, text="Abrir pasta de saída", command=self.open_output).pack(side="left", padx=8)
        ttk.Label(bottom, text="Requer Windows + Microsoft Word instalado").pack(side="right")

    def _folder_row(self, parent, label, variable, command):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=25).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Selecionar…", command=command).pack(side="right")

    def pick_input(self):
        path = filedialog.askdirectory(title="Selecione a pasta com documentos Word")
        if path:
            self.input_var.set(path)

    def pick_output(self):
        path = filedialog.askdirectory(title="Selecione a pasta de saída dos PDFs")
        if path:
            self.output_var.set(path)

    def add_log(self, message):
        self.after(0, lambda: (self.log.insert("end", message + "\n"), self.log.see("end")))

    def update_progress(self, current, total):
        self.after(0, lambda: (self.progress.configure(value=current * 100 / max(total, 1)), self.status_var.set(f"Processando {current} de {total}…")))

    def start(self):
        self.run_button.configure(state="disabled")
        self.log.delete("1.0", "end")
        self.progress.configure(value=0)
        self.status_var.set("Iniciando…")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            log_path = setup_logging()
            self.add_log(f"Log: {log_path}")
            results = process_folder(Path(self.input_var.get().strip()), Path(self.output_var.get().strip()), self.update_progress, self.add_log)
            ok = sum(1 for _, success, _ in results if success)
            failed = len(results) - ok
            summary = f"Concluído: {ok} PDF(s) gerado(s); {failed} erro(s)."
            logging.info(summary)
            self.after(0, lambda: self.status_var.set(summary))
            self.after(0, lambda: messagebox.showinfo("Processamento concluído", summary))
        except Exception as exc:
            logging.exception("Processing could not start")
            self.add_log("ERRO | " + str(exc))
            self.after(0, lambda: self.status_var.set("Não foi possível iniciar o processamento."))
            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
        finally:
            self.after(0, lambda: self.run_button.configure(state="normal"))

    def open_output(self):
        path = Path(self.output_var.get().strip())
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path))
        except Exception as exc:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {exc}")


if __name__ == "__main__":
    App().mainloop()
