import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from stats_helpers import (
    bh_fdr, fmt_p, paired_test, mixed_effects_p, LEGEND_FDR_MEM,
)

# --- Data ---
df = pd.read_excel('roh1.xlsx')
df = df[df['Geschlecht_w1_m2'].isin([1, 2])]
df['Ausbild_WA1_FA2_OA3'] = df['Ausbild_WA1_FA2_OA3'].replace({3: 2})  # FAO

matching_cols = ['Ausbild_WA1_FA2_OA3', 'Laryng_diff_1Mac_2Glide_3King_4CMAC', 'tubus', 'atemweg']

matched_pairs = []
for _, group in df.groupby(matching_cols):
    w = group[group['Geschlecht_w1_m2'] == 1]
    m = group[group['Geschlecht_w1_m2'] == 2]
    n = min(len(w), len(m))
    if n > 0:
        matched_pairs.append((
            w.sample(n=n, random_state=42).reset_index(drop=True),
            m.sample(n=n, random_state=42).reset_index(drop=True)
        ))

df_w = pd.concat([w for w, _ in matched_pairs], ignore_index=True)
df_m = pd.concat([m for _, m in matched_pairs], ignore_index=True)

df_w_std = df_w[df_w['atemweg'] == 1].reset_index(drop=True)
df_m_std = df_m[df_m['atemweg'] == 1].reset_index(drop=True)
df_w_dif = df_w[df_w['atemweg'] == 2].reset_index(drop=True)
df_m_dif = df_m[df_m['atemweg'] == 2].reset_index(drop=True)

laryng = {1: 'Macintosh', 2: 'GlideScope', 3: 'King Vision', 4: 'C-MAC'}

outcome_vars = [
    ('res_max_alle_zaehne', 'Force all front teeth'),
    ('res_12_max', 'Tooth 12'),
    ('res_11_max', 'Tooth 11'),
    ('res_21_max', 'Tooth 21'),
    ('res_22_max', 'Tooth 22'),
    ('BURP', 'BURP'),
    ('Intubationsdauer', 'Intubation duration'),
    ('Sicht_Cormack', 'C&L grade'),
]


def analyse_matched(col, dw, dm):
    vw = pd.to_numeric(dw[col], errors='coerce')
    vm = pd.to_numeric(dm[col], errors='coerce')
    mask = (~vw.isna()) & (~vm.isna())
    vw_v = vw[mask]
    vm_v = vm[mask]
    if len(vw_v) < 3:
        return '-', '-', np.nan, np.nan
    diff = vw_v.values - vm_v.values
    from scipy.stats import shapiro
    try:
        p_shap = shapiro(diff)[1]
    except Exception:
        p_shap = 0
    if p_shap > 0.05:
        sw = f"{vw_v.mean():.2f} \u00b1 {vw_v.std():.2f}"
        sm = f"{vm_v.mean():.2f} \u00b1 {vm_v.std():.2f}"
    else:
        def fmt(v):
            return (f"{v.median():.2f} ({v.quantile(0.25):.2f}\u2013{v.quantile(0.75):.2f}) "
                    f"[{v.min():.2f}\u2013{v.max():.2f}]")
        sw, sm = fmt(vw_v), fmt(vm_v)
    p_raw, _ = paired_test(vw_v.values, vm_v.values)
    p_mm = mixed_effects_p(dw.loc[mask.values], dm.loc[mask.values], col)
    return sw, sm, p_raw, p_mm


def build_table_data(dwx, dmx):
    """Compute results per (variable x laryngoscope) for one airway subset.

    Returns (rows, p_list_flat) - p_list_flat in row-major order matching cells.
    """
    rows = []
    p_raw_list = []
    p_mm_list = []
    for col, label in outcome_vars:
        row = {'label': label, 'laryng': {}}
        for lv in laryng:
            ws = dwx[dwx['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv].reset_index(drop=True)
            ms = dmx[dmx['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv].reset_index(drop=True)
            sw, sm, pr, pmm = analyse_matched(col, ws, ms)
            row['laryng'][lv] = {'f': sw, 'm': sm, 'p_raw': pr, 'p_mm': pmm}
            p_raw_list.append(pr)
            p_mm_list.append(pmm)
        rows.append(row)
    qfdr = bh_fdr(p_raw_list)
    qmm = bh_fdr(p_mm_list)
    idx = 0
    for r in rows:
        for lv in laryng:
            r['laryng'][lv]['q'] = qfdr[idx]
            r['laryng'][lv]['qmm'] = qmm[idx]
            idx += 1
    return rows


table_data = build_table_data(df_w_std, df_m_std)
table_data_dif = build_table_data(df_w_dif, df_m_dif)

# --- Word Document ---
doc = Document()
section = doc.sections[-1]
section.orientation = WD_ORIENT.LANDSCAPE
w, h = section.page_height, section.page_width
section.page_width = w
section.page_height = h
section.left_margin = Cm(1)
section.right_margin = Cm(1)
section.top_margin = Cm(1.5)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="{kwargs.get("top","nil")}" w:sz="{kwargs.get("top_sz","0")}" w:space="0" w:color="{kwargs.get("top_color","000000")}"/>'
        f'<w:bottom w:val="{kwargs.get("bottom","nil")}" w:sz="{kwargs.get("bottom_sz","0")}" w:space="0" w:color="{kwargs.get("bottom_color","000000")}"/>'
        f'<w:left w:val="nil" w:sz="0" w:space="0" w:color="000000"/>'
        f'<w:right w:val="nil" w:sz="0" w:space="0" w:color="000000"/>'
        f'</w:tcBorders>')
    tcPr.append(tcBorders)


def style_cell(cell, text, bold=False, size=8, align=WD_ALIGN_PARAGRAPH.CENTER, italic=False):
    cell.text = text
    for p in cell.paragraphs:
        p.alignment = align
        for r in p.runs:
            r.bold = bold
            r.italic = italic
            r.font.size = Pt(size)
            r.font.name = 'Arial'


def write_p_cell(cell, p_raw, p_fdr, p_mm):
    cell.text = ''
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    sig = ''
    try:
        if p_fdr is not None and not np.isnan(p_fdr) and p_fdr < 0.05:
            sig += '*'
    except Exception:
        pass
    try:
        if p_mm is not None and not np.isnan(p_mm) and p_mm < 0.05:
            sig += '\u2020'
    except Exception:
        pass
    r1 = para.add_run(fmt_p(p_raw) + sig)
    r1.font.size = Pt(7)
    r1.font.name = 'Arial'
    try:
        if p_raw is not None and not np.isnan(float(p_raw)) and float(p_raw) < 0.05:
            r1.bold = True
    except Exception:
        pass
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(f"q={fmt_p(p_fdr)}")
    r2.font.size = Pt(5.5)
    r2.font.name = 'Arial'
    r2.italic = True
    p3 = cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(0)
    r3 = p3.add_run(f"pMM={fmt_p(p_mm)}")
    r3.font.size = Pt(5.5)
    r3.font.name = 'Arial'
    r3.italic = True


def build_doc_table(dwx, dmx, t_data, title_text, subtitle_text):
    title = doc.add_heading('', level=1)
    run = title.add_run(title_text)
    run.font.size = Pt(10)
    run.font.name = 'Arial'
    run.bold = True

    legend = doc.add_paragraph()
    legend.paragraph_format.space_after = Pt(4)
    run = legend.add_run(subtitle_text + ' ' + LEGEND_FDR_MEM)
    run.font.size = Pt(7.5)
    run.font.name = 'Arial'
    run.italic = True

    ncols = 13
    table = doc.add_table(rows=1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Normal Table'

    r0 = table.rows[0]
    style_cell(r0.cells[0], '', bold=True)
    for i, (lv, lname) in enumerate(laryng.items()):
        start = 1 + i * 3
        r0.cells[start].merge(r0.cells[start + 2])
        nw = len(dwx[dwx['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv])
        style_cell(r0.cells[start], f'{lname}\n(n = {nw} pairs)', bold=True, size=8)
    for cell in r0.cells:
        set_cell_border(cell, top="single", top_sz="12", top_color="000000")

    r1 = table.add_row()
    style_cell(r1.cells[0], 'Variable', bold=True, size=8, align=WD_ALIGN_PARAGRAPH.LEFT)
    for i, (lv, _) in enumerate(laryng.items()):
        start = 1 + i * 3
        nw = len(dwx[dwx['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv])
        nm = len(dmx[dmx['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv])
        style_cell(r1.cells[start], f'Female\n(n={nw})', bold=True, size=7)
        style_cell(r1.cells[start + 1], f'Male\n(n={nm})', bold=True, size=7)
        style_cell(r1.cells[start + 2], 'P / q / pMM', bold=True, size=7)
    for cell in r1.cells:
        set_cell_border(cell, bottom="single", bottom_sz="8", bottom_color="000000")

    for entry in t_data:
        row = table.add_row()
        style_cell(row.cells[0], entry['label'], size=7, align=WD_ALIGN_PARAGRAPH.LEFT)
        for i, lv in enumerate(laryng):
            start = 1 + i * 3
            vals = entry['laryng'][lv]
            style_cell(row.cells[start], vals['f'], size=6.5)
            style_cell(row.cells[start + 1], vals['m'], size=6.5)
            write_p_cell(row.cells[start + 2], vals['p_raw'], vals['q'], vals['p_mm'])

    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom="single", bottom_sz="12", bottom_color="000000")

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                pf = p.paragraph_format
                if pf.space_before is None:
                    pf.space_before = Pt(1)
                if pf.space_after is None:
                    pf.space_after = Pt(1)


build_doc_table(
    df_w_std, df_m_std, table_data,
    'Tab 3-1. Matched-Pair Comparison of Dental Force, BURP, Intubation Time, and '
    'C&L Grade by Laryngoscope Type in Standard Airway \u2014 Female vs. Male',
    'Standard airway only. Matched pairs (education level [WA/FAO], laryngoscope, '
    'tube type, airway). Data as mean \u00b1 SD or median (IQR) [range] depending on '
    'distribution (Shapiro\u2013Wilk on paired differences). '
    'C&L = Cormack\u2013Lehane; BURP = backward upward rightward pressure; '
    'WA = residents; FAO = specialists and attendings combined.'
)

doc.save('Tabelle 3_OA_laryng.docx')
print("Saved: Tabelle 3_OA_laryng.docx (Tab 3-1)")

doc.add_page_break()

build_doc_table(
    df_w_dif, df_m_dif, table_data_dif,
    'Tab 3-2. Matched-Pair Comparison of Dental Force, BURP, Intubation Time, and '
    'C&L Grade by Laryngoscope Type in Difficult Airway \u2014 Female vs. Male',
    'Difficult airway only. Matched pairs (education level [WA/FAO], laryngoscope, '
    'tube type, airway). Data as mean \u00b1 SD or median (IQR) [range] depending on '
    'distribution (Shapiro\u2013Wilk on paired differences). '
    'C&L = Cormack\u2013Lehane; BURP = backward upward rightward pressure; '
    'WA = residents; FAO = specialists and attendings combined.'
)

doc.save('Tabelle 3_OA_laryng.docx')
print("Updated: Tabelle 3_OA_laryng.docx (Tab 3-2)")
