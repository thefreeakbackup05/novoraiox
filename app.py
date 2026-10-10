from __future__ import annotations

import json
import logging
import os
import re
import shutil
import sys
import tempfile
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Raio-X de Investimentos"
VERSION = "1.0.0"
DEFAULTS = {
    "input_dir": "WORD",
    "output_dir": "PDF",
    "filename_pattern": "[NOME DO PLANO] - Raio X de [MÊS DO RAIO X]",
    "plan_name_override": "",
    "competency_override": "",
    "link_url": "",
    "video_indicators": "nesse vídeo; neste vídeo;nesse vídeo - aqui;nesse vídeo — aqui;neste vídeo - aqui;neste vídeo — aqui",
    "clickable_text": "aqui",
    "footer_brand": "Vivest • Raio-X de Investimentos",
    "footer_date_label": "Documento fechado em:",
    "footer_date": datetime.now().strftime("%d/%m/%Y"),
    "apply_footer": True,
    "overwrite": False,
}
MONTHS = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
          "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def root_dir() -> Path:
    return Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent


def load_config() -> dict:
    cfg = DEFAULTS.copy()
    p = root_dir() / "config.json"
    try:
        saved = json.loads(p.read_text(encoding="utf-8-sig"))
        if isinstance(saved, dict):
            cfg.update({k: v for k, v in saved.items() if k in cfg})
    except (OSError, ValueError, TypeError):
        pass
    return cfg


def save_config(cfg: dict) -> None:
    (root_dir() / "config.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def setup_logging() -> Path:
    log_dir = root_dir() / "logs"
    log_dir.mkdir(exist_ok=True)
    p = log_dir / f"raiox_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(filename=str(p), level=logging.INFO,
                        format="%(asctime)s | %(levelname)s | %(message)s", encoding="utf-8")
    return p


def docx_text(path: Path) -> str:
    from docx import Document
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(parts)


def infer_plan_and_month(path: Path, cfg: dict) -> tuple[str, str]:
    text = docx_text(path)
    flat = re.sub(r"[\u00a0\t]+", " ", text)
    plan = str(cfg.get("plan_name_override", "")).strip()
    month = str(cfg.get("competency_override", "")).strip()

    # DOCX text may contain cover information split across paragraphs, text boxes,
    # or line breaks. Search flexibly instead of requiring a perfect whole line.
    if not plan:
        patterns = (
            r"(?im)\bPlano\s+([^\r\n|]+)",
            r"(?im)\bNome do plano\s*[:\-]\s*([^\r\n|]+)",
        )
        for pattern in patterns:
            match = re.search(pattern, flat)
            if match:
                candidate = match.group(1).strip().strip('"“” ')
                candidate = re.sub(r"\s{2,}.*$", "", candidate).strip()
                if candidate:
                    plan = candidate
                    break

    if not month:
        months_pattern = "|".join(MONTHS)
        match = re.search(rf"(?i)\b({months_pattern})\s*(?:de\s*)?(\d{{4}})\b", flat)
        if match:
            month = f"{match.group(1).lower()} de {match.group(2)}"
        else:
            # Handle common file stems such as FORD_8_2026_Raio-X_OK.docx,
            # AUREN_CD_8_2026_OK.docx and PSAP_Auren_ago_26.docx.
            stem = path.stem.casefold()
            match = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[_ .-]+(20\d{2})(?!\d)", stem)
            if match:
                month = f"{MONTHS[int(match.group(1)) - 1]} de {match.group(2)}"
            else:
                short = {"jan":0,"fev":1,"mar":2,"abr":3,"mai":4,"jun":5,
                         "jul":6,"ago":7,"set":8,"out":9,"nov":10,"dez":11}
                match = re.search(r"(?i)\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[._ -]?(\d{2,4})\b", stem)
                if match:
                    year = match.group(2)
                    month = f"{MONTHS[short[match.group(1).lower()]]} de {('20' + year) if len(year) == 2 else year}"

    if not plan:
        # Fall back to the source filename after removing technical prefixes,
        # competency markers and processing suffixes.
        candidate = path.stem
        candidate = re.sub(r"(?i)(?:^|[_ -])(?:raio[- _]?x|ok|final|aprovado)(?:$|[_ -].*)", "", candidate)
        candidate = re.sub(r"(?i)^(?:\d{1,2}[_ .-]+20\d{2}[_ .-]*)", "", candidate)
        candidate = re.sub(r"(?i)(?:[_ .-]+(?:0?[1-9]|1[0-2])[_ .-]+20\d{2})", "", candidate)
        candidate = re.sub(r"[_]+", " ", candidate).strip(" _-.")
        plan = candidate or path.stem

    if not month:
        month = f"{MONTHS[datetime.now().month - 1]} de {datetime.now().year}"
    return plan, month

def safe_filename(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", s)
    s = re.sub(r"\s+", " ", s).strip().rstrip(". ")
    return s or "Raio X"


def output_filename(path: Path, cfg: dict) -> str:
    plan, month = infer_plan_and_month(path, cfg)
    pattern = str(cfg.get("filename_pattern", DEFAULTS["filename_pattern"]))
    month_only = re.sub(r"\s+de\s+\d{4}$", "", month, flags=re.I).strip().capitalize()
    year_match = re.search(r"\d{4}", month)
    tokens = {
        "[NOME DO PLANO]": plan,
        "[NOME DO PLANO NO WORD]": plan,
        "[MÊS DO RAIO X]": month_only,
        "[MES DO RAIO X]": month_only,
        "[COMPETÊNCIA]": month,
        "[COMPETENCIA]": month,
        "[ARQUIVO DE ORIGEM]": path.stem,
        "[ANO]": year_match.group(0) if year_match else str(datetime.now().year),
    }
    for token, value in tokens.items():
        pattern = pattern.replace(token, value)
    result = safe_filename(pattern)
    return result if result.lower().endswith(".pdf") else result + ".pdf"


def apply_typography_rules(doc) -> list[str]:
    """Apply agreed typography to editable Word text while preserving layout objects."""
    warnings = []
    try:
        # Text in the document body and tables; embedded charts/images remain untouched.
        body = doc.Content
        body.Font.Name = "Open Sans"
        body.Font.Bold = 0
        body.Font.Italic = 0
        body.Font.Size = 10
    except Exception as exc:
        warnings.append(f"não foi possível normalizar o corpo em Open Sans: {exc}")

    section_titles = (
        "introdução", "informações gerais", "histórico de rentabilidade",
        "rentabilidade do plano", "comparativo com o mercado",
        "composição da carteira",
        "evolução do patrimônio líquido", "análise da equipe de investimentos",
    )
    try:
        for paragraph in doc.Paragraphs:
            text = (paragraph.Range.Text or "").strip().replace("\r", "").replace("\x07", "")
            if not text:
                continue
            low = text.casefold()
            font = paragraph.Range.Font
            if "raio" in low and "investimentos" in low and len(text) < 90:
                font.Name = "Campton SemiBold"
                font.Bold = -1
                font.Size = 22
            elif low.startswith("ativos que compõem uma carteira"):
                font.Name = "Open Sans"
                font.Bold = -1
                font.Italic = -1
                font.Size = 10
            elif any(low.startswith(title) for title in section_titles):
                font.Name = "Campton SemiBold"
                font.Bold = -1
                font.Size = 15
            elif low.startswith("rentabilidade acumulada") or low.startswith("dados em mil") or low.startswith("plano na modalidade"):
                font.Name = "Campton Light"
                font.Bold = 0
                font.Size = 11
            elif re.match(r"(?i)^plano\s+.+", text) and len(text) < 100:
                font.Name = "Campton SemiBold"
                font.Bold = -1
                font.Size = 14
            elif re.match(r"(?i)^(?:\*{1,2}|fonte de informações|fonte indicadores)", text):
                font.Name = "Open Sans"
                font.Size = 8
            # The six asset descriptions commonly place the heading before a colon
            # in the same paragraph. Format the heading run separately.
            elif ":" in text and any(key in low for key in (
                "renda fixa", "multimercado", "empréstimos", "renda variável", "exterior", "imobiliário"
            )):
                font.Name = "Open Sans Light"
                font.Italic = -1
                font.Bold = 0
                colon = text.find(":")
                try:
                    heading = paragraph.Range.Duplicate
                    heading.End = heading.Start + colon
                    heading.Font.Name = "Open Sans"
                    heading.Font.Bold = -1
                    heading.Font.Italic = -1
                except Exception:
                    warnings.append(f"não foi possível destacar o título do bloco: {text[:45]}")
        # Table contents often hold the return history and indicator comparison.
        for table in doc.Tables:
            for row in table.Rows:
                for cell in row.Cells:
                    try:
                        cell.Range.Font.Name = "Open Sans"
                    except Exception:
                        pass
    except Exception as exc:
        warnings.append(f"aplicação tipográfica parcial: {exc}")
    return warnings


def export_word_to_pdf(source: Path, destination: Path, cfg: dict) -> tuple[bool, str]:
    """Export with Word COM; isolate hyperlink failures and retry PDF export once."""
    import pythoncom
    import win32com.client

    indicators = [x.strip().casefold() for x in str(cfg.get("video_indicators", "")).split(";") if x.strip()]
    click_text = str(cfg.get("clickable_text", "aqui")).strip()
    url = str(cfg.get("link_url", "")).strip()
    tmpdir = Path(tempfile.mkdtemp(prefix="raiox_word_"))
    workdoc = tmpdir / source.name
    shutil.copy2(source, workdoc)
    pythoncom.CoInitialize()
    word = None
    doc = None
    video_found = False
    link_added = False
    stage = "iniciar Word"
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.Options.ConfirmConversions = False
        stage = "abrir documento"
        try:
            doc = word.Documents.Open(
                FileName=str(workdoc.resolve()), ConfirmConversions=False,
                ReadOnly=False, AddToRecentFiles=False, Visible=False,
                OpenAndRepair=True, NoEncodingDialog=True,
            )
        except Exception:
            logging.exception("Primeira tentativa de abertura falhou para %s; tentando abertura simples", source.name)
            doc = word.Documents.Open(str(workdoc.resolve()), ReadOnly=False, AddToRecentFiles=False)

        stage = "aplicar tipografia"
        typography_warnings = apply_typography_rules(doc)
        for warning in typography_warnings:
            logging.warning("%s | %s", source.name, warning)
        stage = "detectar chamada de vídeo"
        fulltext = doc.Content.Text or ""
        low = fulltext.casefold()
        indicator_pos = -1
        for indicator in indicators:
            pos = low.find(indicator)
            if pos >= 0:
                video_found = True
                indicator_pos = pos
                break

        if video_found and url and click_text:
            # Word's range offsets use carriage returns, not the literal two-character
            # sequence "\\r". Limit search to the callout paragraph to avoid unrelated links.
            paragraph_start = low.rfind("\r", 0, indicator_pos)
            paragraph_end = low.find("\r", indicator_pos)
            if paragraph_start < 0:
                paragraph_start = max(0, indicator_pos - 100)
            if paragraph_end < 0:
                paragraph_end = min(len(low), indicator_pos + 220)
            local_click = low.find(click_text.casefold(), paragraph_start, min(paragraph_end, indicator_pos + 220))
            if local_click < 0:
                local_click = low.find(click_text.casefold(), indicator_pos, min(len(low), indicator_pos + 120))
            if local_click >= 0:
                stage = "inserir hiperlink"
                try:
                    anchor = doc.Range(local_click, local_click + len(click_text))
                    doc.Hyperlinks.Add(Anchor=anchor, Address=url, ScreenTip="Saiba mais")
                    link_added = True
                except Exception:
                    logging.exception("Falha ao inserir hiperlink em %s; a exportação continuará sem link", source.name)

        stage = "exportar PDF"
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(destination.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                From=1,
                To=1,
                Item=0,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        except Exception:
            logging.exception("ExportAsFixedFormat falhou para %s; tentando SaveAs2 PDF", source.name)
            try:
                if destination.exists():
                    destination.unlink()
                doc.SaveAs2(FileName=str(destination.resolve()), FileFormat=17, AddToRecentFiles=False)
            except Exception as fallback_exc:
                raise RuntimeError(f"Falha na etapa '{stage}' via ExportAsFixedFormat e SaveAs2: {fallback_exc}") from fallback_exc
    except Exception as exc:
        raise RuntimeError(f"Falha na etapa '{stage}' para '{source.name}': {exc}") from exc
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                logging.exception("Could not close temporary Word document")
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                logging.exception("Could not quit Word")
        pythoncom.CoUninitialize()
        shutil.rmtree(tmpdir, ignore_errors=True)
    return video_found, ("link inserido" if link_added else ("vídeo detectado, mas link não inserido" if video_found else "sem chamada de vídeo")) + (f" | {len(typography_warnings)} aviso(s) de tipografia" if typography_warnings else " | tipografia aplicada")

def apply_pdf_footer(pdf_path: Path, cfg: dict) -> None:
    """Draw a configurable footer on the last PDF page only."""
    if not cfg.get("apply_footer", True):
        return
    import fitz
    doc = fitz.open(str(pdf_path))
    if not len(doc):
        doc.close()
        raise RuntimeError("O PDF não contém páginas.")
    page = doc[-1]
    rect = page.rect
    # Replace an existing copy of the configurable footer on the last page,
    # avoiding duplicate footer text when the source DOCX already has one.
    brand = str(cfg.get("footer_brand", DEFAULTS["footer_brand"])).strip()
    label = str(cfg.get("footer_date_label", DEFAULTS["footer_date_label"])).strip()
    date_text = str(cfg.get("footer_date", "")).strip()
    for old_text in (brand, label):
        if old_text:
            for hit in page.search_for(old_text):
                page.add_redact_annot(hit, fill=(1, 1, 1))
    if brand or label:
        page.apply_redactions(images=0, graphics=0)
    # Reserve a subtle footer in the bottom margin.
    y = rect.height - 22
    left = brand
    right = f"{label} {date_text}".strip()
    font_size = 7.5
    margin = 34
    page.insert_text((margin, y), left, fontsize=font_size, fontname="helv", color=(0.35, 0.35, 0.35), overlay=True)
    width = fitz.get_text_length(right, fontname="helv", fontsize=font_size)
    page.insert_text((rect.width - margin - width, y), right, fontsize=font_size, fontname="helv", color=(0.35, 0.35, 0.35), overlay=True)
    tmp = pdf_path.with_name(pdf_path.stem + ".tmp.pdf")
    doc.save(str(tmp), garbage=4, deflate=True)
    doc.close()
    os.replace(tmp, pdf_path)


@dataclass
class Result:
    source: str
    output: str = ""
    status: str = "ERRO"
    detail: str = ""
    pages: int = 0
    video: bool = False
    link: bool = False


def process_folder(cfg: dict, report):
    import fitz
    input_dir = Path(str(cfg["input_dir"])).expanduser()
    output_dir = Path(str(cfg["output_dir"])).expanduser()
    if not input_dir.is_dir():
        raise ValueError(f"A pasta de entrada não existe: {input_dir}")
    files = sorted((p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")), key=lambda p: p.name.casefold())
    if not files:
        raise ValueError(f"Nenhum documento .docx encontrado em {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for i, source in enumerate(files, 1):
        r = Result(source=source.name)
        destination = output_dir / output_filename(source, cfg)
        try:
            if destination.exists() and not cfg.get("overwrite", False):
                raise FileExistsError(f"O PDF já existe: {destination.name}. Ative 'Substituir PDFs existentes' nas configurações para sobrescrever.")
            temp_pdf = output_dir / (".raiox_temp_" + destination.name)
            if temp_pdf.exists():
                temp_pdf.unlink()
            r.video, r.detail = export_word_to_pdf(source, temp_pdf, cfg)
            if not temp_pdf.exists() or temp_pdf.stat().st_size < 100:
                raise RuntimeError("O Microsoft Word não gerou um PDF válido.")
            apply_pdf_footer(temp_pdf, cfg)
            with fitz.open(str(temp_pdf)) as pdf:
                r.pages = len(pdf)
                if r.pages == 0:
                    raise RuntimeError("PDF sem páginas.")
                extracted = "\n".join(page.get_text() for page in pdf).strip()
                if not extracted:
                    raise RuntimeError("O PDF não contém texto extraível; exige revisão visual.")
            r.link = "link inserido" in r.detail
            r.detail += f" | {r.pages} página(s); texto extraível validado"
            os.replace(temp_pdf, destination)
            r.output = destination.name
            r.status = "OK" if (not r.video or r.link or not cfg.get("link_url", "").strip()) else "REVISAR"
        except Exception as exc:
            r.status = "ERRO"
            r.detail = str(exc)
            logging.exception("Erro processando %s", source.name)
        results.append(r)
        report(i, len(files), r)
    return results


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} — Gerador Word v{VERSION}")
        self.geometry("1040x800")
        self.minsize(900, 700)
        self.cfg = load_config()
        self.vars = {}
        self._build()
        self._load()

    def _build(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#17365d")
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="RAIO-X DE INVESTIMENTOS", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Geração a partir de Word (.docx), com nome, hiperlink e rodapé configuráveis.").pack(anchor="w", pady=(2, 12))
        canvas = tk.Canvas(outer, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas, padding=(2, 2, 10, 2))
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._section(body, "Pastas")
        self._field(body, "Pasta dos documentos Word", "input_dir", browse=True)
        self._field(body, "Pasta de saída dos PDFs", "output_dir", browse=True)
        self._section(body, "Nome do arquivo")
        self._field(body, "Padrão do nome do PDF", "filename_pattern")
        ttk.Label(body, text="Tokens: [NOME DO PLANO], [NOME DO PLANO NO WORD], [MÊS DO RAIO X], [ANO], [ARQUIVO DE ORIGEM]").pack(anchor="w", padx=8)
        self._field(body, "Nome do plano (opcional; vazio = detectar no Word)", "plan_name_override")
        self._field(body, "Competência/mês (opcional; vazio = detectar no Word)", "competency_override")
        self._section(body, "Hiperlink condicional")
        self._field(body, "URL do vídeo", "link_url")
        self._field(body, "Textos que identificam vídeo (separados por ;)", "video_indicators")
        self._field(body, "Palavra que recebe o hiperlink", "clickable_text")
        self._section(body, "Rodapé da última página")
        self._field(body, "Texto institucional", "footer_brand")
        self._field(body, "Rótulo da data", "footer_date_label")
        self._field(body, "Data de fechamento (dd/mm/aaaa)", "footer_date")
        self.apply_footer = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=False)
        ttk.Checkbutton(body, text="Aplicar rodapé somente à última página", variable=self.apply_footer).pack(anchor="w", padx=8, pady=4)
        ttk.Checkbutton(body, text="Substituir PDFs existentes", variable=self.overwrite).pack(anchor="w", padx=8, pady=4)
        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=10)
        self.run_btn = ttk.Button(buttons, text="GERAR PDFs", command=self.start)
        self.run_btn.pack(side="left")
        ttk.Button(buttons, text="Salvar configurações", command=self.save).pack(side="left", padx=8)
        ttk.Button(buttons, text="Abrir saída", command=self.open_output).pack(side="left")
        self.progress = ttk.Progressbar(body, mode="determinate")
        self.progress.pack(fill="x", pady=(4, 4))
        self.status = tk.StringVar(value="Pronto.")
        ttk.Label(body, textvariable=self.status).pack(anchor="w", pady=(0, 6))
        self.log = tk.Text(body, height=10, wrap="word", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

    def _section(self, parent, title):
        ttk.Label(parent, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(12, 3))

    def _field(self, parent, label, key, browse=False):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", padx=6, pady=3)
        ttk.Label(frame, text=label, width=42).pack(side="left")
        var = tk.StringVar()
        self.vars[key] = var
        ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True)
        if browse:
            ttk.Button(frame, text="Selecionar…", command=lambda k=key: self.pick(k)).pack(side="left", padx=(6, 0))

    def _load(self):
        for key, var in self.vars.items():
            value = self.cfg.get(key, DEFAULTS.get(key, ""))
            if key in ("input_dir", "output_dir") and not Path(str(value)).is_absolute():
                value = str(root_dir() / value)
            var.set(str(value))
        self.apply_footer.set(bool(self.cfg.get("apply_footer", True)))
        self.overwrite.set(bool(self.cfg.get("overwrite", False)))

    def current_config(self):
        cfg = self.cfg.copy()
        for key, var in self.vars.items():
            cfg[key] = var.get().strip()
        cfg["apply_footer"] = self.apply_footer.get()
        cfg["overwrite"] = self.overwrite.get()
        try:
            datetime.strptime(cfg["footer_date"], "%d/%m/%Y")
        except ValueError:
            raise ValueError("A data do rodapé deve estar no formato dd/mm/aaaa.")
        if cfg["link_url"] and not re.match(r"^https?://", cfg["link_url"], re.I):
            raise ValueError("A URL do hiperlink deve começar com http:// ou https://.")
        return cfg

    def save(self):
        try:
            self.cfg = self.current_config()
            save_config(self.cfg)
            self.status.set("Configurações salvas.")
        except Exception as exc:
            messagebox.showerror("Configurações", str(exc))

    def pick(self, key):
        path = filedialog.askdirectory(title="Selecione a pasta")
        if path:
            self.vars[key].set(path)

    def start(self):
        try:
            cfg = self.current_config()
            save_config(cfg)
        except Exception as exc:
            messagebox.showerror("Configurações", str(exc))
            return
        self.run_btn.configure(state="disabled")
        self.log.delete("1.0", "end")
        self.progress.configure(value=0, maximum=1)
        self.status.set("Iniciando…")
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _worker(self, cfg):
        try:
            log_path = setup_logging()
            self.add_log(f"Log: {log_path}")
            def report(i, total, r):
                self.after(0, lambda: self.progress.configure(maximum=total, value=i))
                self.add_log(f"{r.status} | {r.source} → {r.output or '-'} | {r.detail}")
                self.after(0, lambda: self.status.set(f"Processados {i}/{total}"))
            results = process_folder(cfg, report)
            ok = sum(r.status == "OK" for r in results)
            review = sum(r.status == "REVISAR" for r in results)
            errors = sum(r.status == "ERRO" for r in results)
            summary = f"Concluído: {ok} OK, {review} para revisão, {errors} erro(s)."
            self.add_log(summary)
            self.after(0, lambda: messagebox.showinfo("Geração concluída", summary))
            self.after(0, lambda: self.status.set(summary))
        except Exception as exc:
            logging.exception("Falha ao iniciar geração")
            self.add_log("ERRO | " + str(exc))
            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
            self.after(0, lambda: self.status.set("Falha na geração."))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))

    def add_log(self, text):
        self.after(0, lambda: (self.log.insert("end", text + "\n"), self.log.see("end")))

    def open_output(self):
        try:
            p = Path(self.vars["output_dir"].get())
            p.mkdir(parents=True, exist_ok=True)
            os.startfile(str(p))
        except Exception as exc:
            messagebox.showerror("Saída", str(exc))


if __name__ == "__main__":
    App().mainloop()
