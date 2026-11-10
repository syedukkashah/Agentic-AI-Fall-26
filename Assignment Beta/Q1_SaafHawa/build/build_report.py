"""
Build the Q1 Word report and the AI-tools declaration from the notebook.

Every number in the report comes from executing the notebook's own code cells, so the report
always matches the notebook. After re-running the latency logger on day 2:

    jupyter nbconvert --to notebook --execute --inplace Assignment_Beta_Question_no_1_23K-0055.ipynb
    python build/build_report.py

Run from the Q1_SaafHawa folder. Needs: pip install python-docx
Convert to PDF afterwards with Word (File > Save As > PDF) or build/docx_to_pdf.ps1.
"""
import contextlib
import io
import json
import math
import re
import sys
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent.parent
NOTEBOOK = HERE / "Assignment_Beta_Question_no_1_23K-0055.ipynb"
REPORT = HERE / "Assignment_Beta_Q1_Report_23K-0055.docx"
DECLARATION = HERE / "AI_Tools_Declaration_23K-0055.docx"
STUDENT, ROLL_NO = "Syed Ukkashah", "23K-0055"
CONFIG_URL = "https://huggingface.co/mistralai/Ministral-3-14B-Instruct-2512/blob/main/config.json"

ACCENT = RGBColor(0x1F, 0x4E, 0x79)
HEADER_FILL, BAND_FILL, WORK_FILL, NOTE_FILL = "1F4E79", "DCE6F1", "F2F2F2", "FFF2CC"


# ------------------------------------------------------------------ run the notebook code
def run_notebook():
    nb = json.loads(NOTEBOOK.read_text(encoding="utf-8"))
    ns = {"display": lambda *a, **k: None}
    import os
    os.chdir(HERE)
    with contextlib.redirect_stdout(io.StringIO()):
        for cell in nb["cells"]:
            if cell["cell_type"] == "code":
                exec(compile("".join(cell["source"]), "<nb>", "exec"), ns)
    return ns


# ------------------------------------------------------------------ docx helpers
def set_cell_fill(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_table_cell_margins(table, top=40, bottom=40, left=70, right=70):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}"); el.set(qn("w:w"), str(v)); el.set(qn("w:type"), "dxa"); mar.append(el)
    tblPr.append(mar)


def fixed_layout(table):
    layout = OxmlElement("w:tblLayout"); layout.set(qn("w:type"), "fixed")
    table._tbl.tblPr.append(layout)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trPr.append(el)


def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit"); el.set(qn("w:val"), "true"); trPr.append(el)


def add_runs(par, text, size=None, color=None, font=None):
    """Add text to a paragraph; **bold**, *italic* and `code` markers are honoured."""
    for part in re.split(r"(\*\*[^*]+\*\*|(?<!\w)\*(?=\S)[^*]+?(?<=\S)\*(?!\w)|`[^`]+`)", text):
        if not part:
            continue
        if part.startswith("**"):
            run = par.add_run(part[2:-2]); run.bold = True
        elif part.startswith("`"):
            run = par.add_run(part[1:-1]); run.font.name = "Consolas"
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = par.add_run(part[1:-1]); run.italic = True
        else:
            run = par.add_run(part)
        if size: run.font.size = Pt(size)
        if color: run.font.color.rgb = color
        if font: run.font.name = font
    return par


class Report:
    def __init__(self):
        self.doc = Document()
        self.fig_no = 0
        self.tab_no = 0
        st = self.doc.styles
        normal = st["Normal"]
        normal.font.name = "Calibri"; normal.font.size = Pt(10.5)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        normal.paragraph_format.space_after = Pt(4)
        normal.paragraph_format.line_spacing = 1.1
        for name, size, before in (("Title", 20, 0), ("Heading 1", 15, 14), ("Heading 2", 12.5, 10), ("Heading 3", 11, 8)):
            s = st[name]
            s.font.name = "Calibri"; s.font.size = Pt(size); s.font.bold = True
            s.font.color.rgb = ACCENT
            rfonts = s.element.rPr.find(qn("w:rFonts"))
            if rfonts is not None:
                for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                    rfonts.attrib.pop(qn(attr), None)
                rfonts.set(qn("w:ascii"), "Calibri"); rfonts.set(qn("w:hAnsi"), "Calibri")
            s.paragraph_format.space_before = Pt(before); s.paragraph_format.space_after = Pt(4)
            s.paragraph_format.keep_with_next = True
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        for side in ("left_margin", "right_margin"):
            setattr(sec, side, Cm(2.0))
        sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
        self._header_footer(sec)

    # -- page furniture
    def _header_footer(self, sec):
        hp = sec.header.paragraphs[0]
        hp.text = ""
        add_runs(hp, f"Agentic AI (AI4015) · Assignment Beta · Question 1 · {STUDENT} ({ROLL_NO})", size=8, color=RGBColor(0x80, 0x80, 0x80))
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = fp.add_run(); run.font.size = Pt(8)
        for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
            if tag:
                el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag); run._r.append(el)
            else:
                el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = text; run._r.append(el)

    def new_section(self, landscape):
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        if landscape:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
            sec.left_margin = sec.right_margin = Cm(1.3)
            sec.top_margin = sec.bottom_margin = Cm(1.4)
        else:
            sec.orientation = WD_ORIENT.PORTRAIT
            sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
            sec.left_margin = sec.right_margin = Cm(2.0)
            sec.top_margin = sec.bottom_margin = Cm(1.8)
        return sec

    def page_break(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # -- text
    def h(self, text, level=1):
        return self.doc.add_heading(text, level=level)

    def p(self, text, size=None, align=None, space_after=None, italic=False):
        par = add_runs(self.doc.add_paragraph(), text, size=size)
        if italic:
            for r in par.runs: r.italic = True
        if align: par.alignment = align
        if space_after is not None: par.paragraph_format.space_after = Pt(space_after)
        return par

    def bullets(self, items, numbered=False):
        style = "List Number" if numbered else "List Bullet"
        for it in items:
            par = add_runs(self.doc.add_paragraph(style=style), it)
            par.paragraph_format.space_after = Pt(2)

    def working(self, lines, width_cm=17.0):
        """Monospace block for step-by-step arithmetic."""
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = t.rows[0].cells[0]; cell.width = Cm(width_cm)
        set_cell_fill(cell, WORK_FILL)
        set_table_cell_margins(t, 60, 60, 120, 120)
        first = True
        for line in lines:
            par = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
            run = par.add_run(line); run.font.name = "Consolas"; run.font.size = Pt(8.5)
            run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Consolas")
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def note(self, text, fill=NOTE_FILL, width_cm=17.0):
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = t.rows[0].cells[0]; cell.width = Cm(width_cm)
        set_cell_fill(cell, fill)
        set_table_cell_margins(t, 80, 80, 140, 140)
        paras = text if isinstance(text, list) else [text]
        for i, txt in enumerate(paras):
            par = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            par.paragraph_format.space_after = Pt(2)
            add_runs(par, txt, size=10)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def caption(self, text, kind="Table"):
        if kind == "Table":
            self.tab_no += 1; label = f"Table {self.tab_no}. "
        else:
            self.fig_no += 1; label = f"Figure {self.fig_no}. "
        par = self.doc.add_paragraph()
        r = par.add_run(label); r.bold = True; r.font.size = Pt(9); r.font.color.rgb = ACCENT
        add_runs(par, text, size=9)
        par.paragraph_format.space_after = Pt(3)
        if kind == "Table":
            par.paragraph_format.keep_with_next = True
        else:
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return par

    def table(self, header, rows, widths_cm, size=8.5, align=None, bands=None, bold_first_col=False, highlight=None):
        """rows: list of lists of str. bands: {row_index: label} inserts a shaded full-width band row before that row.
        highlight: set of row indexes to shade."""
        bands = bands or {}
        highlight = highlight or set()
        n_rows = 1 + len(rows) + len(bands)
        t = self.doc.add_table(rows=n_rows, cols=len(header))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        fixed_layout(t)
        set_table_cell_margins(t)
        for j, txt in enumerate(header):
            c = t.rows[0].cells[j]
            c.width = Cm(widths_cm[j]); set_cell_fill(c, HEADER_FILL)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            par = c.paragraphs[0]; par.paragraph_format.space_after = Pt(0)
            run = par.add_run(txt); run.bold = True; run.font.size = Pt(size); run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        repeat_header(t.rows[0])
        r_i = 1
        for i, row in enumerate(rows):
            if i in bands:
                band = t.rows[r_i]; cant_split(band)
                merged = band.cells[0].merge(band.cells[-1])
                set_cell_fill(merged, BAND_FILL)
                par = merged.paragraphs[0]; par.paragraph_format.space_after = Pt(0)
                run = par.add_run(bands[i]); run.bold = True; run.font.size = Pt(size)
                r_i += 1
            tr = t.rows[r_i]; cant_split(tr)
            for j, txt in enumerate(row):
                c = tr.cells[j]; c.width = Cm(widths_cm[j])
                if i in highlight: set_cell_fill(c, NOTE_FILL)
                lines = str(txt).split("\n")
                for k, line in enumerate(lines):
                    par = c.paragraphs[0] if k == 0 else c.add_paragraph()
                    par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
                    if align and align[j]: par.alignment = align[j]
                    add_runs(par, line, size=size)
                    if bold_first_col and j == 0:
                        for r in par.runs: r.bold = True
            r_i += 1
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)
        return t

    def figure(self, path, caption, width_cm=16.5):
        par = self.doc.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.keep_with_next = True
        par.add_run().add_picture(str(path), width=Cm(width_cm))
        self.caption(caption, kind="Figure")

    def save(self, path):
        self.doc.save(path)


# ------------------------------------------------------------------ formatting
R_ = WD_ALIGN_PARAGRAPH.RIGHT
C_ = WD_ALIGN_PARAGRAPH.CENTER


def fnum(v, nd=None):
    """Thousands separators for numbers; strings pass through."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if math.isnan(f):
        return "not reported"
    if nd is not None:
        return f"{f:,.{nd}f}"
    return f"{int(f):,}" if f.is_integer() else f"{f:g}"


def tag(text, src):
    return f"{text} [{src}]"


def usd(x, nd=0):
    return f"${x:,.{nd}f}"


# ------------------------------------------------------------------ the report
def build_report(ns):
    g = ns  # shorthand
    R, D, Q, N, p, H, B, T = (g[k] for k in "R D Q N p H B T".split())
    models, sources, costs, rate, options, cand = g["models"], g["sources"], g["costs"], g["rate"], g["options"], g["cand"]
    log, ok, per_tool = g["log"], g["ok"], g["per_tool"]
    IN_Q, OUT_Q, CALLS_Q = g["IN_Q"], g["OUT_Q"], g["CALLS_Q"]
    mq = g["monthly_queries"]
    days = sorted(ok["date"].unique())
    c_h = costs.set_index("model")
    best_api = g["api_name"]
    best_cost = c_h.loc[best_api, "monthly $ (cache)"]
    fallback = [m for m in cand.model if m != best_api][0]
    fb_cost = c_h.loc[fallback, "monthly $ (cache)"]
    sh = g["selfhost_monthly"]
    peak_tpm = g["peak_tpm_in"] + g["peak_tpm_out"]
    ratio = g["peak_qpm"] / g["avg_qpm"]
    sh_per_q = g["GPU_PRICE_HR"] * 720 / g["max_seqs"] * g["peak_concurrency"](1) / (30 + 4 * H)
    api_per_q = g["api_monthly"](1) / (30 + 4 * H)
    w = g["weights"]; npar = g["n_params"]; GB = 1e9
    sl, kvb = g["seq_len"], g["kv_bytes"]
    rep = Report()

    # ============================================================ page 1
    t = rep.doc.add_paragraph(style="Title"); add_runs(t, "SaafHawa: closed-source API or self-hosted open-weight model?")
    t.paragraph_format.space_after = Pt(2)
    rep.p("Agentic AI (AI4015), Assignment Beta, Question 1", size=11.5, space_after=0)
    rep.p(f"{STUDENT}, roll no. {ROLL_NO}, FAST-NUCES Karachi, Fall 2026, {date.today():%d %B %Y}", size=10.5, space_after=8)

    rep.h("Parameters", 1)
    digits = g["digits"]
    rep.p(f"My roll number is {ROLL_NO}. Ignoring the letter and the hyphen, its digits are {', '.join(map(str, digits[:-1]))} and {digits[-1]}, "
          "which give the parameters below.", space_after=4)
    rep.table(["Symbol", "Parameter", "Formula", "Working", "Value"], [
        ["R", "Last two digits of roll number", "", f"…{ROLL_NO[-2:]}", f"{R}"],
        ["D", "Sum of digits mod 10", "Σ digits mod 10", f"({' + '.join(map(str, digits))}) = {sum(digits)}; {sum(digits)} mod 10", f"{D}"],
        ["Q", "Normal daily queries", "1,500 + 25 × R", f"1,500 + 25 × {R} = 1,500 + {25 * R:,}", f"{Q:,}"],
        ["N", "LLM calls per query (first attempt)", "5 + (R mod 4)", f"5 + ({R} mod 4) = 5 + {R % 4}", f"{N}"],
        ["p", "Tool-call failure probability per call", "0.04 + 0.01 × D", f"0.04 + 0.01 × {D}", f"{p:.2f}"],
        ["H", "Heatwave days per month", "4 + (R mod 4)", f"4 + ({R} mod 4) = 4 + {R % 4}", f"{H}"],
        ["B", "Monthly inference budget (USD)", "1,200 + 10 × R", f"1,200 + 10 × {R} = 1,200 + {10 * R:,}", f"${B:,}"],
        ["T", "End-to-end latency target (s)", "10 − (R mod 3)", f"10 − ({R} mod 3) = 10 − {R % 3}", f"{T} s"],
    ], [1.4, 5.0, 3.0, 5.6, 2.0], size=9.5, align=[C_, None, None, None, R_])

    rep.p("These parameters lead to the following values, which the rest of the report relies on.", space_after=4)
    rep.table(["Quantity", "Value", "Section"], [
        ["Expected input tokens per query, including retries", f"{IN_Q:,.1f}", "2.1"],
        ["Expected output tokens per query, including retries", f"{OUT_Q:,.1f}", "2.1"],
        ["Expected LLM requests per query", f"{N} + {p} × {N - 1} = {CALLS_Q:.2f}", "2.1"],
        ["Queries per month (30 days, H heatwave days at 5Q)", f"{mq:,}", "2.1"],
        ["Input / output tokens per month", f"{g['monthly_in'] / 1e9:.3f} B / {g['monthly_out'] / 1e9:.3f} B", "2.1"],
        ["Peak load, heatwave day, 12:00 to 14:00", f"{g['peak_qpm']:.2f} queries/min = {g['peak_rpm']:,.0f} RPM, {peak_tpm / 1e6:.2f} M TPM", "2.3"],
    ], [8.5, 6.5, 2.0], size=9.5, align=[None, None, C_])

    rep.p(f"My answer to the technical lead is to pay for an API. At SaafHawa's traffic the cheapest model that handles tool calls well is "
          f"{best_api}, served through a hosted API for {usd(best_cost)} a month against a budget of {usd(B)}. {fallback} is the fallback at "
          f"{usd(fb_cost)} a month. Self-hosting the one open model with standard attention would take {g['gpus_needed'](Q)} GPUs and "
          f"{usd(sh(Q))} a month, and it never becomes cheaper than this API anywhere between 10 and 200,000 queries a day. The GPUs have to be "
          f"sized for a heatwave peak {ratio:.1f} times the average load, but they are billed around the clock. None of the candidates meets the "
          f"{T} s latency target with an eight-turn loop either, so the agent loop needs redesigning as well (Section 2.6).")

    # ============================================================ Section 1 (landscape)
    rep.new_section(landscape=True)
    rep.h("Section 1: Model comparison", 1)
    closed, open_ = models[models.side == "closed"], models[models.side == "open"]
    rep.p(f"I compared {len(closed)} closed-source models from {closed.provider.nunique()} providers ({', '.join(closed.provider.unique())}) and "
          f"{len(open_)} open-weight models from {open_.provider.nunique()} providers ({', '.join(open_.provider.unique())}). Each model runs in its "
          "non-reasoning or lowest reasoning setting, shown in brackets after its name. Every figure carries a source tag [S#], and I accessed all "
          "sources on 2026-11-09; the list follows the table. For the open-weight models I quote the cheapest public host for this workload's token "
          "mix, which is about 12.6 input tokens per output token, using OpenRouter's live endpoints API, and I name the host and its quantisation.",
          size=9.5)

    rows, bands = [], {}
    for i, m in enumerate(models.itertuples()):
        if i == 0:
            bands[0] = "Closed-source models"
        if m.side == "open" and "Open-weight models" not in bands.values():
            bands[i] = "Open-weight models"
        if str(m.params_total_B) == "not disclosed":
            params = "not disclosed"
        else:
            active = str(m.params_active_B)
            params = f"{m.params_total_B} ({active})" if active.startswith("dense") else f"{m.params_total_B} total\n{active} active"
        host = (f"host: {m.host}" + (f", {m.quantization}" if str(m.quantization) not in ("nan", "not stated") else "")) \
            if m.side == "open" else "first-party API"
        if fnum(m.rpm) == "not reported":
            rl = "not published"
        elif m.provider == "Anthropic":
            rl = f"{fnum(m.rpm)} RPM\n{fnum(m.itpm)} ITPM\n{fnum(m.otpm)} OTPM"
        else:
            rl = f"{fnum(m.rpm)} RPM\n{fnum(m.itpm)} TPM"
        speed = "not reported" if fnum(m.ttft_s) == "not reported" else f"{m.ttft_s} s / {m.output_tps}"
        tb4 = fnum(m.terminal_bench4_score)
        lines = [] if str(m.bench1_score) == "not reported" else [tag(f"{m.bench1_name}: {m.bench1_score}", m.bench_src)]
        if tb4 != "not reported":
            lines.append(f"Terminal-Bench 4.0: {tb4} [S33]")
        bench1 = "\n".join(lines) or tag("not reported", m.bench_src)
        bsrc = "S33" if "S33" in m.bench_src else m.bench_src
        tool_b = f"{m.bench2_tool_name}: {m.bench2_tool_score}"
        if isinstance(m.vendor_tool_note, str):
            tool_b += f"\n{m.vendor_tool_note}"
        rows.append([
            f"**{m.model}**\n({m.variant_used})",
            f"{m.provider}",
            tag(params, m.spec_src),
            tag(f"{fnum(m.context_tokens)} / {fnum(m.max_output_tokens)}", m.spec_src),
            tag(f"{m.price_in} / {m.price_out} / {m.price_cached}", m.price_src) + f"\n{host}",
            tag(rl, m.rate_src),
            tag(speed, m.speed_src),
            bench1,
            tag(tool_b, bsrc),
            tag(f"{m.tool_calling} / {m.json_output}", m.feature_src),
            tag(m.reasoning_settings, m.feature_src),
            tag(f"{m.licence} / {m.commercial_self_host}", m.feature_src),
        ])
    rep.caption("Comparison of 14 models. Every figure is sourced, and every source was accessed on 2026-11-09.")
    rep.table(["Model (setting used)", "Provider", "Params (B)\ntotal / active", "Context / max output (tokens)",
               "$ per M tokens\nin / out / cached", "Rate limit, new paid tier", "TTFT / output tok/s",
               "General benchmark(s)", "Tool-use benchmark", "Tools / JSON", "Reasoning settings offered", "Licence / commercial self-host"],
              rows, [2.6, 1.55, 1.85, 2.0, 2.6, 2.05, 1.65, 2.75, 2.75, 1.15, 2.4, 1.75], size=6.8, bands=bands)
    rep.p("AutomationBench-AA is Artificial Analysis's agentic tool-use evaluation, scored in percent, and I use it as the tool-use benchmark. "
          "The AA Intelligence Index v5.0.1 combines ten evaluations. Where Artificial Analysis did not evaluate the setting I use, the table names "
          "the setting it did evaluate.", size=8)

    rep.caption("Sources cited in Table 1, all accessed on 2026-11-09.")
    rep.table(["ID", "Source", "URL"], [[s.id, s.source, s.url] for s in sources.itertuples()], [1.2, 10.5, 15.0], size=7.5)

    rep.p("Two values in the table could not be verified, and both affect the later analysis. The first is the number of reasoning tokens that "
          "Claude Haiku 5.5, Claude Sonnet 5.5, Gemini 3.5 Flash-Lite and Gemini 3.8 Flash produce at their lowest setting. These models cannot turn "
          "thinking off, and neither the providers nor Artificial Analysis publish how many thinking tokens they average at low or minimal effort. "
          "I therefore add zero reasoning tokens for them, which makes their cost look lower than it is. If low effort produced even 300 thinking "
          "tokens per call, output per query would rise by about 80%, and the output part of the bill with it. The recommendation would stay the "
          "same, since Claude Haiku is already the more expensive of the two candidates and the Gemini and Sonnet models are over budget even "
          "without thinking tokens. Latency suffers less from this gap, because Artificial Analysis measures the time to the first answer token at "
          "that setting, and that time already includes the thinking.", size=9.5)
    rep.p("The second is the rate limit a new paid account gets on Google Gemini (Tier 1) and on the open-weight hosts. Google shows per-model RPM "
          "and TPM only inside a signed-in AI Studio console. OpenRouter says paid models have no platform-level cap but that upstream providers "
          "apply their own limits, and those hosts do not publish them. I cannot confirm that the heatwave peak in Section 2.3 fits within these "
          "limits, so I mark them as unknown. The model I recommend is one of them, which makes a written rate-limit quota, or a load test at "
          f"{g['peak_rpm']:,.0f} RPM, a condition for launch.", size=9.5)
    rep.p("A few smaller gaps appear as \"not reported\" in the table. Artificial Analysis stopped benchmarking Grok 4.20 after it was deprecated, "
          "so the table has no speed or tool-use score for it and only one benchmark. xAI's caching documentation page returns a 404 error, and "
          "DeepSeek does not state the active parameter count for V4.1 Flash.", size=9.5)

    # ============================================================ Section 2
    rep.new_section(landscape=False)
    rep.h("Section 2: Calculations", 1)
    rep.p("The notebook Assignment_Beta_Question_no_1_23K-0055.ipynb holds every calculation, and I reproduce the working here. The notebook "
          "computes the latency budget (2.6) before 2.4 and 2.5 because both of those sections use its results.")

    # ---------------- 2.1
    rep.h("2.1 Token volume", 2)
    rep.p(f"For each call k from 1 to N, with N = {N}, the input is 2,000 + 800 × (k − 1) tokens and the output is 350 tokens, except for the "
          f"final call, which produces 600. Calls 1 to N − 1 fail with probability p = {p}, and a failed call is repeated once with the same input, "
          f"so each of those calls runs 1 + p = {1 + p:.2f} times on average. The final call is never retried.")
    calls = g["calls"]
    crow = [[str(int(k)), f"2,000 + 800 × {int(k) - 1} = {int(i):,}", f"{int(o)}", f"{er:.2f}", f"{ei:,.1f}", f"{eo:,.1f}"]
            for k, i, o, er, ei, eo in calls[["call k", "input", "output", "expected runs", "expected input", "expected output"]].itertuples(index=False)]
    crow.append(["**Σ**", f"**{calls['input'].sum():,}**", f"**{calls['output'].sum():,}**", f"**{CALLS_Q:.2f}**",
                 f"**{IN_Q:,.1f}**", f"**{OUT_Q:,.1f}**"])
    rep.caption("Tokens per LLM call, including the expected retries.")
    rep.table(["Call k", "Input tokens", "Output", "Expected runs", "Expected input", "Expected output"], crow,
              [1.5, 5.0, 1.8, 2.6, 3.0, 3.1], size=9, align=[C_, None, R_, R_, R_, R_])
    first_in, retry_base_in = calls["input"].sum(), calls["input"][:-1].sum()
    first_out, retry_base_out = calls["output"].sum(), calls["output"][:-1].sum()
    nd = 30 - H
    rep.working([
        f"Input per query  = first attempt + retries = {first_in:,} + {p} × {retry_base_in:,} = {first_in:,} + {p * retry_base_in:,.1f} = {IN_Q:,.1f}",
        f"Output per query = {first_out:,} + {p} × {retry_base_out:,} = {first_out:,} + {p * retry_base_out:,.1f} = {OUT_Q:,.1f}",
        f"LLM requests per query = {N} + {p} × {N - 1} = {CALLS_Q:.2f}",
        "",
        f"Queries per month = (30 − H) × Q + H × 5Q = {nd} × {Q:,} + {H} × 5 × {Q:,} = {nd * Q:,} + {H * 5 * Q:,} = {mq:,}",
        f"Input tokens/month  = {mq:,} × {IN_Q:,.1f} = {g['monthly_in']:,.0f}  ({g['monthly_in'] / 1e9:.3f} B)",
        f"Output tokens/month = {mq:,} × {OUT_Q:,.1f} = {g['monthly_out']:,.0f}  ({g['monthly_out'] / 1e9:.3f} B)",
    ])

    # ---------------- 2.2
    rep.h("2.2 Monthly API cost, with and without prompt caching", 2)
    cpq = g["CACHED_PER_QUERY"]
    rep.p("Without caching, one query costs the expected input tokens times the input price plus the expected output tokens times the output "
          "price, divided by 10⁶. The monthly cost is that figure times the number of queries in the month. With caching, the first 1,500 input "
          f"tokens of every LLM request are the same system prompt and tool schemas, so 1,500 × {CALLS_Q:.2f} = {cpq:,.1f} input tokens per query "
          f"are billed at the cached rate and the remaining {IN_Q - cpq:,.1f} at the normal rate.")
    rep.p("Writing to the cache costs extra. Traffic averages about two queries a minute even on normal days, so the shared prefix should stay "
          "warm, but to be safe I assume the prefix is rewritten once per cache lifetime for the whole month. That is 30 × 24 × 60 / TTL writes, "
          "each costing the write price minus the cached read price. Mistral does not state a TTL, so I assume 5 minutes.")
    rep.p("I apply caching only where the provider documents a minimum cacheable prefix of 1,500 tokens or less: OpenAI (1,024), Anthropic (512) "
          "and Mistral (64). Gemini 3.x needs 4,096 tokens, xAI does not document its rules, and the third-party open-weight hosts publish no "
          "minimum, so I price those models without caching. Reasoning tokens count as zero at the chosen settings, for the reason given after Table 1.")
    r_h = models[models.model == "Claude Haiku 5.5"].iloc[0]
    writes = g["MINUTES_PER_MONTH"] / g["TTL_MINUTES"][r_h.provider]
    cq = g["cost_per_query"]
    rep.p(f"Claude Haiku 5.5 shows the arithmetic. It charges ${r_h.price_in} per million input tokens, ${r_h.price_out} for output, "
          f"${r_h.price_cached} for cached input and ${r_h.price_cache_write} for a cache write, and its cache lasts 5 minutes.")
    rep.working([
        f"No cache : ({IN_Q:,.1f} × {r_h.price_in} + {OUT_Q:,.1f} × {r_h.price_out}) / 1e6 = ${cq(r_h, False):.5f}/query",
        f"           × {mq:,} queries = ${g['monthly_cost'](r_h, False):,.2f}/month",
        f"Cache    : ({IN_Q - cpq:,.1f} × {r_h.price_in} + {cpq:,.1f} × {r_h.price_cached} + {OUT_Q:,.1f} × {r_h.price_out}) / 1e6 = ${cq(r_h, True):.5f}/query",
        f"           × {mq:,} = ${cq(r_h, True) * mq:,.2f}",
        f"Writes   : 30×24×60 / 5 = {writes:,.0f} writes × 1,500 × ({r_h.price_cache_write} − {r_h.price_cached}) / 1e6 = ${writes * 1500 * (float(r_h.price_cache_write) - r_h.price_cached) / 1e6:.2f}",
        f"Total with caching = ${g['monthly_cost'](r_h, True):,.2f}/month   (B = ${B:,})",
    ])
    order = costs.join(models[["price_cache_write"]]).sort_values("monthly $ (cache)")
    cost_rows, hi = [], set()
    for i, r in enumerate(order.itertuples()):
        wp = r.price_cache_write
        wp = f"{float(wp):g}" if re.fullmatch(r"[\d.]+", str(wp)) else str(wp)
        ttl = str(r._7).replace(" (default; 1 h optional)", " (1 h opt.)").replace(" (default)", "")
        cost_rows.append([r.model, r.side, f"{r._3} / {r._4} / {r._5}", wp,
                          str(r._6).replace("not published by host", "not published").replace(" (other Gemini 3.x: 4,096)", "; 3.x: 4,096"),
                          ttl, "yes" if r._8 else "no", usd(r._9), usd(r._10),
                          ("✓" if r._11 else "✗") + " / " + ("✓" if r._12 else "✗")])
        if r.model in list(cand.model):
            hi.add(i)
    rep.caption(f"Monthly API cost at Q = {Q:,} a day and H = {H} ({mq:,} queries). Prices are in $ per million tokens, and the two latency "
                "candidates from 2.6 are highlighted. Cache rules come from [S2] OpenAI, [S7] Anthropic, [S11] Google, [S15] xAI, [S17] Mistral "
                "and [S19] the OpenRouter hosts.")
    rep.table(["Model", "Side", "$/M in / out / cached", "Cache write $/M", "Min. prefix", "Cache TTL", "Cache applies", "No cache", "With cache", "≤ B?\nno / with"],
              cost_rows, [3.0, 1.1, 2.3, 1.7, 1.9, 1.9, 1.2, 1.5, 1.5, 1.3], size=7.5,
              align=[None, None, None, None, None, None, C_, R_, R_, C_], highlight=hi)
    inb_nc = costs[costs["within B (no cache)"]].model.tolist()
    inb_c = costs[costs["within B (cache)"]].model.tolist()
    cached = costs[costs["caching applies"]]
    saving = ", ".join(f"{r.model} {1 - r._10 / r._9:.0%}" for r in cached.itertuples())
    over_c = costs[~costs["within B (cache)"]]
    fit = (f"{len(inb_nc)} models fit the ${B:,} budget without caching, and the same {len(inb_c)} fit with it: "
           if set(inb_nc) == set(inb_c) else
           f"{len(inb_nc)} models fit the ${B:,} budget without caching and {len(inb_c)} fit with it. With caching they are ")
    rep.p(f"{fit}{', '.join(inb_c[:-1])} and {inb_c[-1]}. Caching saves {saving}, but it does not bring any model under the budget line. "
          f"{', '.join(over_c.model[:-1])} and {over_c.model.iloc[-1]} cost {over_c['monthly $ (cache)'].min() / B:.0f} to "
          f"{over_c['monthly $ (cache)'].max() / B:.0f} times the budget even with caching.")
    rep.figure(HERE / "figures" / "q1_monthly_cost.png", f"Monthly API cost with and without caching, on a log scale. The red line marks the budget B = ${B:,}.")

    # ---------------- 2.3
    rep.h("2.3 Peak load and rate limits", 2)
    hd, wq = g["heat_day_queries"], g["window_queries"]
    rep.p("On a heatwave day traffic rises to 5Q, and 35% of that day's queries arrive in the two hours between 12:00 and 14:00.")
    rep.working([
        f"Heatwave day traffic     = 5 × Q = 5 × {Q:,} = {hd:,} queries",
        f"Queries 12:00 to 14:00   = 0.35 × {hd:,} = {wq:,.2f} queries in 120 minutes",
        f"Peak queries per minute  = {wq:,.2f} / 120 = {g['peak_qpm']:.2f}",
        f"Peak requests per minute = {g['peak_qpm']:.2f} × {CALLS_Q:.2f} LLM requests/query = {g['peak_rpm']:,.1f} RPM",
        f"Peak input TPM           = {g['peak_qpm']:.2f} × {IN_Q:,.1f} = {g['peak_tpm_in']:,.0f}",
        f"Peak output TPM          = {g['peak_qpm']:.2f} × {OUT_Q:,.1f} = {g['peak_tpm_out']:,.0f}",
        f"Peak total TPM           = {peak_tpm:,.0f}",
        f"Uncached input TPM       = {g['peak_qpm']:.2f} × {IN_Q - cpq:,.1f} = {g['peak_tpm_uncached_in']:,.0f}  (Anthropic counts only these toward ITPM)",
    ])
    rate_rows = []
    for r in rate.itertuples():
        lim = "not published" if fnum(r._2) == "not reported" else (
            f"{fnum(r._2)} RPM; {fnum(r._3)} ITPM + {fnum(r._4)} OTPM" if fnum(r._4) != "not reported" else f"{fnum(r._2)} RPM; {fnum(r._3)} TPM")
        rate_rows.append([r.model, lim, str(r._5), r._6])
    rep.caption("Peak load against each provider's limits for a new paid account (sources as in Table 1).")
    rep.table(["Model", "Limit", "Tier / note", "Peak check"], rate_rows, [3.6, 4.3, 6.2, 2.9], size=7.5)
    rep.p(f"GPT-6.1 Terra's new-account limit of 1,000,000 combined tokens per minute is below the {peak_tpm / 1e6:.2f} M TPM peak, so it would "
          "throttle SaafHawa at midday during a heatwave. GPT-6 Luna, both Claude models and Grok fit. Anthropic leaves cache reads out of its "
          f"input limit, so for Claude the number to compare is {g['peak_tpm_uncached_in']:,.0f} against 2,000,000. For Gemini and the open-weight "
          "hosts I cannot say either way, for the reason given after Table 1.")

    # ---------------- 2.4
    rep.h("2.4 Self-hosting memory", 2)
    rep.p("For self-hosting I chose Ministral 3 14B Instruct, which is released under Apache 2.0. It is the only open-weight model in Table 1 with "
          "standard full grouped-query attention, with no sliding window and no linear or latent attention, so the assignment's KV-cache formula "
          "applies as written. Qwen3.8 mixes linear and full attention, Gemma 4 and gpt-oss combine sliding and global attention, and Mistral "
          "Small 4, DeepSeek and GLM use latent or sparse attention. The values below come from the text_config block of the model's config.json "
          "on Hugging Face [S24].")
    rep.table(["config.json field", "Value"], [
        ["num_hidden_layers", str(g["layers"])], ["num_key_value_heads", str(g["kv_heads"])], ["head_dim", str(g["head_dim"])],
        ["num_attention_heads", str(g["text_cfg"]["num_attention_heads"])], ["sliding_window", str(g["text_cfg"].get("sliding_window"))],
        ["parameters (safetensors total, incl. vision encoder) [S31]", f"{npar:,}"],
    ], [9.0, 4.0], size=9, align=[None, R_])
    shot = HERE / "figures" / "q1_config_json_screenshot.png"
    if shot.exists():
        rep.figure(shot, f"The text_config block of config.json for mistralai/Ministral-3-14B-Instruct-2512 on Hugging Face, showing "
                         f"num_hidden_layers = {g['layers']}, num_key_value_heads = {g['kv_heads']} and head_dim = {g['head_dim']} "
                         f"(screenshot taken 2026-11-10 from {CONFIG_URL}).", width_cm=8.5)
    else:
        ph = rep.doc.add_table(rows=1, cols=1); ph.style = "Table Grid"; ph.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = ph.rows[0].cells[0]; c.width = Cm(17); set_cell_fill(c, NOTE_FILL)
        for i, line in enumerate(["[ INSERT SCREENSHOT HERE ]",
                                  f"config.json of mistralai/Ministral-3-14B-Instruct-2512 showing num_hidden_layers = {g['layers']}, "
                                  f"num_key_value_heads = {g['kv_heads']}, head_dim = {g['head_dim']}",
                                  CONFIG_URL]):
            par = c.paragraphs[0] if i == 0 else c.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.space_before = Pt(18 if i == 0 else 2); par.paragraph_format.space_after = Pt(18 if i == 2 else 2)
            run = par.add_run(line); run.font.size = Pt(11 if i == 0 else 9); run.bold = i == 0
        rep.caption("config.json of Ministral 3 14B on Hugging Face (screenshot, accessed 2026-11-09).", kind="Figure")

    rep.p("Weight memory is the parameter count times the bytes per parameter.")
    rep.working([f"FP16  : {npar:,} × 2 bytes   = {w['FP16'] / GB:.2f} GB",
                 f"8-bit : {npar:,} × 1 byte    = {w['8-bit'] / GB:.2f} GB",
                 f"4-bit : {npar:,} × 0.5 bytes = {w['4-bit'] / GB:.2f} GB"])
    rep.p(f"For the KV cache I use the longest context in the loop. The final call (N = {N}) has {g['input_tokens'](N):,} input tokens and "
          f"generates {g['output_tokens'](N)} more, so the sequence reaches {sl:,} tokens.")
    rep.working([f"KV bytes = 2 × layers × kv_heads × head_dim × seq_len × bytes_per_value",
                 f"         = 2 × {g['layers']} × {g['kv_heads']} × {g['head_dim']} × {sl:,} × 2 = {kvb:,} bytes = {kvb / GB:.3f} GB"])
    rep.p(f"For the GPU I picked an {g['GPU_NAME']} at ${g['GPU_PRICE_HR']} an hour on demand [S32]. At 4-bit the model needs only about 7 GB, so "
          "48 GB leaves plenty of room. The card is also cheaper per concurrent sequence than an H100 80 GB, which costs $2.89 an hour and holds "
          "48 sequences against the L40S's 26 at $1.09. I reserve 10% of the memory for activations and runtime overhead.")
    rep.working([f"Usable for KV = 0.90 × {g['GPU_MEM_GB']} GB − {w['4-bit'] / GB:.2f} GB (4-bit weights) = {g['usable'] / GB:.2f} GB",
                 f"Max sequences = floor({g['usable'] / GB:.2f} / {kvb / GB:.3f}) = {g['max_seqs']}"])
    pq = g["peak_qpm"] / 60
    rep.p("To count GPUs I need the number of sessions open at the same time. Little's law gives this as the arrival rate times the time each "
          "session holds its KV cache, and a session holds it for the whole query, tool waits included. I estimate that time with the 2.6 latency "
          f"model, using Ministral 3 14B's measured speed (2.54 s TTFT, 65 tok/s [S21; S22]) as a stand-in for our own server. The result is "
          f"{g['selfhost_latency']:.1f} s per query.")
    rep.working([f"Peak arrival rate   = {g['peak_qpm']:.2f} queries/min / 60 = {pq:.3f} queries/s",
                 f"Concurrent sessions = {pq:.3f} × {g['selfhost_latency']:.1f} s = {g['conc']:.1f}",
                 f"GPUs needed         = ceil({g['conc']:.1f} / {g['max_seqs']}) = {g['gpus_needed'](Q)}"])

    # ---------------- 2.5
    rep.h("2.5 Break-even analysis", 2)
    rep.p(f"Self-hosting cost rises in steps. The number of GPUs is max(1, ceil(peak concurrency / {g['max_seqs']})), and each GPU is billed for "
          f"the full month, so the monthly cost is GPUs × ${g['GPU_PRICE_HR']} × 24 × 30 = GPUs × ${g['GPU_PRICE_HR'] * 720:,.1f}.")
    rep.p(f"The assignment asks for a comparison with the 2.6 candidate that meets T after the architectural change, or with the faster one if "
          f"neither does. Neither meets T, so the comparison is with {best_api}. Its cost grows in a straight line with volume, because queries "
          f"per month = Q(30 − H) + 5QH = Q(30 + 4H) = {30 + 4 * H}Q.")
    rep.working([
        f"At our Q = {Q:,}: self-host = {g['gpus_needed'](Q)} × ${g['GPU_PRICE_HR'] * 720:,.1f} = ${sh(Q):,.0f}/month;  API = ${g['api_monthly'](Q):,.0f}/month",
        f"Self-host $/query at large Q ≈ ${g['GPU_PRICE_HR']} × 720 / {g['max_seqs']} × concurrency per daily query / {30 + 4 * H} = ${sh_per_q:.5f}",
        f"API $/query ({best_api}) = ${api_per_q:.5f}",
        (f"Scan Q = 10 to 200,000/day: self-hosting is never cheaper, so there is no Q* against {best_api}" if not g["Q_star"]
         else f"Q* = {g['Q_star']:,} queries/day"),
        f"Average load = {mq:,} / {g['MINUTES_PER_MONTH']:,} min = {g['avg_qpm']:.2f} q/min; peak = {g['peak_qpm']:.2f}; peak/average = {ratio:.1f}",
    ])
    be = g["breakeven"].sort_values("API $/month at Q")
    be_rows = [[r._1, r.side, usd(r._3), "never (up to 200k)" if (r._4 is None or (isinstance(r._4, float) and math.isnan(r._4))) else f"{int(r._4):,}",
                "" if (isinstance(r._5, float) and math.isnan(r._5)) else f"{r._5}"] for r in be.itertuples()]
    rep.caption(f"Break-even volume Q* of self-hosting (step function, {g['max_seqs']} sessions per GPU) against every API model.")
    rep.table(["API model", "Side", "API $/month at our Q", "Q* (queries/day)", "Q* / our Q"], be_rows, [5.0, 1.6, 3.6, 3.6, 2.4], size=8,
              align=[None, None, R_, R_, R_], highlight={i for i, r in enumerate(be.itertuples()) if r._1 == best_api})
    rep.figure(HERE / "figures" / "q1_breakeven.png", "Monthly cost against daily query volume, log scales on both axes: the self-hosting step function against the API candidates.")
    prem = be[be["Q* (queries/day)"].notna()]
    prem_names = list(prem["API model"])
    rep.p(f"Against {best_api} at {usd(best_cost)} a month, the API SaafHawa would actually use, self-hosting never breaks even. It costs about "
          f"${sh_per_q:.5f} per query against the API's ${api_per_q:.5f}, and because both costs grow in proportion to volume, more traffic never "
          f"closes the gap. The GPUs have to be sized for the heatwave peak between 12:00 and 14:00, which is {ratio:.1f} times the average load, "
          f"yet they are billed 24 hours a day for 30 days. Self-hosting beats only the premium closed models ({', '.join(prem_names[:-1])} and "
          f"{prem_names[-1]}), where Q* falls between {int(prem['Q* (queries/day)'].min()):,} and {int(prem['Q* (queries/day)'].max()):,} queries "
          f"a day, far below our {Q:,}. Those models are over budget anyway and would not be chosen. GPUs that autoscale and are paid for only during "
          "peak hours, or a serving stack faster than the 65 tok/s per stream assumed here, could change the picture. For a WhatsApp assistant "
          "handling a few thousand queries a day, self-hosting is not a realistic option today.")

    # ---------------- 2.6
    rep.h("2.6 Latency budget", 2)
    rep.p("To measure the tool delay I wrote latency_logger.py. Each run calls the air-quality API (OpenAQ v3, PM2.5 and PM10 for a Karachi "
          "station) and the weather API (Open-Meteo: temperature, humidity and wind) 20 times each, and times the NOAA heat-index function and the "
          "knowledge-base lookup on my machine. Every timing goes into the raw log logs/tool_latency_log.csv, which is part of this submission.")
    if len(days) < 2:
        rep.note(f"Day-2 run pending: the log currently covers {len(days)} day ({', '.join(days)}). Re-run latency_logger.py on a second day, "
                 "re-execute the notebook and rebuild this report before submitting.")
    by_day = ok.groupby(["date", "tool"]).size().unstack(fill_value=0)
    pt_rows = []
    for tool, r in per_tool.iterrows():
        pt_rows.append([tool, str(int(r["count"])), ", ".join(f"{d}: {by_day.loc[d, tool]}" for d in by_day.index),
                        f"{r['median']:,.3f}", f"{r['mean']:,.3f}", f"{r['max']:,.3f}"])
    rep.caption(f"Tool latency measurements for successful calls ({len(log)} rows logged; {len(log) - len(ok)} failed OpenAQ connections excluded).")
    rep.table(["Tool", "n", "Calls per day", "Median (ms)", "Mean (ms)", "Max (ms)"], pt_rows, [2.6, 1.0, 5.4, 2.6, 2.6, 2.6], size=8.5,
              align=[None, R_, None, R_, R_, R_])
    TD, API_TD = g["TOOL_DELAY"], g["api_median_s"]
    rep.working([f"Median of ALL {len(ok)} measurements (as instructed) = {TD * 1000:.3f} ms; tool delay used = {TD:.6f} s",
                 f"Median of the two network APIs only (sensitivity)    = {API_TD * 1000:.1f} ms"])
    rep.p("Half of the measurements come from local functions that finish in well under a millisecond, so the pooled median is close to zero. "
          "A real agent waits about 0.5 to 0.7 s for each API tool, which means the pooled median understates the delay. I use it because the "
          f"assignment requires it and report the API-only median ({API_TD:.3f} s) as a check. That check changes no conclusion, since LLM "
          "generation takes most of the time.")
    rep.p("Calls run one after another. Each call takes t_k = TTFT + extra prefill + output tokens ÷ output speed, with TTFT and output speed taken "
          "from Artificial Analysis at the chosen setting. Artificial Analysis measures TTFT on a fixed prompt, so I add the extra input at an "
          f"assumed prefill rate of {g['PREFILL_TPS']:,} tokens per second, giving extra prefill = (input_k − 2,000) / {g['PREFILL_TPS']:,}. This "
          "adds at most 1.1 s, at call 8. The agent catches an invalid tool call before the tool runs, so a retry repeats only the LLM call, and the "
          "expected retry delay is p × Σ t_k over k < N. The total is Σ t_k + (N − 1) × tool delay + retry delay.")
    elig = g["eligible"].sort_values("bench2_tool_score", ascending=False)
    rep.p("I picked the two strongest candidates from the models that fit the budget with caching, have a single-call TTFT below T and have been "
          "benchmarked by Artificial Analysis. Among those I took the two with the highest AutomationBench-AA score, because choosing and calling "
          "tools is exactly what SaafHawa's loop does.")
    rep.caption("Models eligible as latency candidates. The two highlighted rows have the highest tool-use scores.")
    rep.table(["Model", "$/month (cache)", "TTFT (s)", "Output tok/s", "AutomationBench-AA"],
              [[r.model, usd(r.monthly_cache), f"{r.ttft_s:g}", f"{r.output_tps:g}", f"{r.bench2_tool_score:g}"] for r in elig.itertuples()],
              [5.4, 2.8, 2.2, 2.6, 3.6], size=8.5, align=[None, R_, R_, R_, R_], highlight={0, 1})
    cl = g["call_latency"]
    for r in cand.itertuples():
        times = [cl(r, k) for k in range(1, N + 1)]
        rows_k = []
        for k, tk in zip(range(1, N + 1), times):
            ex = (g["input_tokens"](k) - g["input_tokens"](1)) / g["PREFILL_TPS"]
            gen = g["output_tokens"](k) / r.output_tps
            rows_k.append([str(k), f"{g['input_tokens'](k):,}", f"{r.ttft_s:g}", f"{ex:.2f}", f"{g['output_tokens'](k)} / {r.output_tps:g} = {gen:.2f}", f"{tk:.2f}"])
        rows_k.append(["**Σ**", "", "", "", "", f"**{sum(times):.2f}**"])
        rep.caption(f"Per-call latency for {r.model} (TTFT {r.ttft_s:g} s, {r.output_tps:g} tok/s).")
        rep.table(["Call k", "Input", "TTFT (s)", "Extra prefill (s)", "Generation (s)", "t_k (s)"], rows_k, [1.6, 2.2, 2.2, 3.0, 5.0, 2.4], size=8.5,
                  align=[C_, R_, R_, R_, R_, R_])
        retry = p * sum(times[:-1]); tools = (N - 1) * TD
        rep.working([f"Σ t_k = {sum(times):.2f} s;  tools = {N - 1} × {TD:.6f} = {tools:.4f} s;  retries = {p} × {sum(times[:-1]):.2f} = {retry:.2f} s",
                     f"Total = {sum(times):.2f} + {tools:.4f} + {retry:.2f} = {g['query_latency'](r):.2f} s   (T = {T} s)   "
                     f"[with API-only tool median: {g['query_latency'](r, tool_delay=API_TD):.2f} s]"])
    rep.p(f"Both candidates miss the {T} s target by a wide margin. Most of the time goes to generating output. Seven tool-calling turns of 350 "
          "tokens plus a 600-token answer come to about 3,050 tokens per query, and at 100 to 250 tok/s that alone takes 12 to 30 s before any "
          "TTFT is counted. Claude Haiku 5.5 also pays its 6.8 s TTFT at the low setting on every turn.")
    router = g["router"].model
    rep.p("I tested the three suggested changes one at a time and then together. The first runs independent tool calls in parallel. Air quality, "
          "weather and the knowledge-base lookup do not depend on each other, so one LLM turn can request all three, which removes two LLM turns "
          f"and two tool waits and cuts N from {N} to {N - 2}. The second streams the final response, so the user starts reading at the final "
          "call's first token and that call counts only its TTFT and prefill. The third routes the tool-selection turns 1 to N − 1 to the fastest "
          f"model in the table that supports tool calling, which is {router}, while the final answer stays on the candidate. For {router} itself "
          "routing changes nothing, since it is already that model.")
    opt = options.drop(columns="meets T?")
    rep.caption("End-to-end latency in seconds after each architectural change.")
    rep.table(["Model"] + [c.replace(" (s)", "") for c in opt.columns],
              [[m] + [f"{v:.2f}" for v in opt.loc[m]] for m in opt.index],
              [3.6, 2.0, 2.8, 2.8, 2.8, 3.0], size=8.5, align=[None, R_, R_, R_, R_, R_])
    best = cand.iloc[0]
    sens = []
    for tps in [150, 200, 300, 400, 600, 800]:
        faster = best.copy(); faster["output_tps"] = tps
        sens.append(f"{g['query_latency'](faster, n=N - 2, stream_final=True, router=faster):.2f} s at {tps} tok/s")
    rep.p(f"Parallel tool calls give the largest single gain, taking {best.model} from {opt.loc[best.model].iloc[0]:.2f} s to "
          f"{opt.loc[best.model].iloc[1]:.2f} s, so that is the change I recommend. All three together bring it to "
          f"{opt.loc[best.model].iloc[-1]:.2f} s, which is still above T. Faster generation would not get there on its own either: the combined "
          f"design takes {', '.join(sens[:-1])} and {sens[-1]}. Reaching {T} s also needs shorter tool-selection turns. A compact JSON tool call "
          "is about 50 tokens, while this model assumes 350 for every turn.")

    # ============================================================ conclusions
    rep.h("Conclusions", 1)
    over = costs[~costs["within B (cache)"]]
    haiku_nc, haiku_c = c_h.loc["Claude Haiku 5.5", "monthly $ (no cache)"], c_h.loc["Claude Haiku 5.5", "monthly $ (cache)"]
    rep.p(f"SaafHawa will use {g['monthly_in'] / 1e9:.2f} billion input and {g['monthly_out'] / 1e9:.2f} billion output tokens a month. "
          f"Of the {len(costs)} models, {len(inb_c)} fit the ${B:,} budget, and prompt caching brings Claude Haiku 5.5 down from {usd(haiku_nc)} to "
          f"{usd(haiku_c)}, a saving of {1 - haiku_c / haiku_nc:.0%}. {', '.join(over.model[:-1])} and {over.model.iloc[-1]} stay over budget "
          "even with caching.")
    rep.p(f"The heatwave midday peak is {g['peak_qpm']:.1f} queries a minute, which works out to {g['peak_rpm']:,.0f} requests and "
          f"{peak_tpm / 1e6:.2f} million tokens a minute. GPT-6.1 Terra's 1 M TPM limit for a new account would throttle that. GPT-6 Luna, both "
          "Claude models and Grok fit, while Gemini and the open-weight hosts do not publish their limits.")
    rep.p(f"Neither candidate meets the {T} s target. {cand.iloc[0].model} takes {opt.iloc[0, 0]:.1f} s and {cand.iloc[1].model} takes "
          f"{opt.iloc[1, 0]:.1f} s, and nearly all of that is spent generating about 3,050 output tokens over {N} sequential turns. The measured "
          f"tool delay hardly registers. Parallel tool calls help most ({opt.iloc[0, 1]:.1f} s), all three changes together reach "
          f"{opt.iloc[0, -1]:.1f} s, and the loop also needs short JSON tool turns.")
    rep.p(f"Self-hosting Ministral 3 14B needs {w['4-bit'] / GB:.2f} GB for 4-bit weights and {kvb / GB:.2f} GB of KV cache per {sl:,}-token "
          f"session, so a 48 GB L40S holds {g['max_seqs']} sessions. The peak needs about {g['conc']:.0f} concurrent sessions, which means "
          f"{g['gpus_needed'](Q)} GPUs at {usd(sh(Q))} a month. That is over budget and {sh(Q) / g['api_monthly'](Q):.1f} times the {best_api} "
          "API bill, and it never breaks even. Ministral also scores just 0.6% on AutomationBench-AA, so it could not handle SaafHawa's tool "
          "calls in any case.")
    rep.p(f"My recommendation is {best_api} through an API, at {usd(best_cost)} a month. It is the cheapest model with strong tool use, scoring "
          "48.8% on AutomationBench-AA (close to Claude Sonnet 5.5 at low effort), and it has the fastest TTFT on the shortlist. The agent loop "
          "should be rebuilt around parallel tool calls, streamed answers and short JSON tool calls. The host does not publish its rate limits, "
          f"which is the main risk, so a written quota has to be in place before launch. {fallback} is the fallback. Its limits are documented and "
          f"cover the peak, and it supports prompt caching, but it costs {fb_cost / best_cost:.1f} times as much and is much slower at the low "
          "setting. The morning summary for municipal officers needs one more safeguard. A wrong figure there reaches people who make decisions "
          "for thousands of residents, so code should copy the numbers straight from the tool outputs instead of letting the model write them.")
    rep.p("Along with this report I am submitting the executed notebook Assignment_Beta_Question_no_1_23K-0055.ipynb, the raw timing log "
          "logs/tool_latency_log.csv together with the latency_logger.py script that produced it, the Section 1 data in data/models.csv, "
          "data/sources.csv and data/snapshots/, the two figures in the figures folder, and my declaration of the AI tools I used.")
    rep.save(REPORT)


# ------------------------------------------------------------------ declaration
def build_declaration():
    rep = Report()
    t = rep.doc.add_paragraph(style="Title"); add_runs(t, "AI Tools Declaration")
    rep.p("Agentic AI (AI4015), Assignment Beta, Question 1 (SaafHawa)", size=11.5, space_after=0)
    rep.p(f"{STUDENT}, roll no. {ROLL_NO}, {date.today():%d %B %Y}", space_after=10)
    rep.p("I used the following AI tools while preparing this submission. Text in highlighted square brackets is mine to complete or correct before submitting.")
    rows = [
        ["Claude Code (Anthropic), Claude Opus models", "Research and data collection",
         "Searched provider documentation, pricing and rate-limit pages, the OpenRouter and Hugging Face APIs and Artificial Analysis, and compiled "
         "data/models.csv and data/sources.csv with a source tag for every value.",
         "[Which figures I checked against the source pages myself; the AutomationBench-AA and Terminal-Bench values were read from my own screenshots of the AA leaderboard.]"],
        ["Claude Code", "Code",
         "Wrote latency_logger.py (OpenAQ, Open-Meteo, heat-index and knowledge-base tools plus the timing logger), the notebook generator "
         "build/build_q1.py and the report builder build/build_report.py.",
         "[I ran the logger myself on both days with my own OpenAQ key, executed the notebook and checked the outputs.]"],
        ["Claude Code", "Calculations and analysis",
         "Implemented the Section 2 formulas (token volume, caching cost, peak load, KV cache, break-even, latency) in the notebook following the assignment's assumptions.",
         "[I checked the parameter derivation from my roll number and re-did selected calculations by hand, e.g. …]"],
        ["Claude Code", "Writing",
         "Drafted the report text and tables from the notebook outputs, and this declaration.",
         "[I reviewed and edited the report; the conclusions and recommendation are ones I agree with and can defend.]"],
        ["[Any other tool, e.g. ChatGPT, Grammarly, Copilot – or delete this row]", "[Purpose]", "[What it did]", "[How I checked it]"],
    ]
    rep.table(["Tool", "Purpose", "What the tool did", "My own contribution / verification"], rows, [3.2, 2.6, 5.8, 5.4], size=9)
    for tbl in rep.doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    for run in par.runs:
                        if run.text.startswith("[") or run.text.endswith("]"):
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    rep.p("All numbers in the report are produced by the submitted notebook from the submitted data files and timing log, and every external figure "
          "is cited to its source. I understand the work and take responsibility for its correctness.")
    rep.p("Signature: ______________________        Date: ______________", space_after=0)
    rep.save(DECLARATION)


if __name__ == "__main__":
    ns = run_notebook()
    build_report(ns)
    print(f"wrote {REPORT.name}")
    if "--no-declaration" not in sys.argv:
        build_declaration()
        print(f"wrote {DECLARATION.name}")
