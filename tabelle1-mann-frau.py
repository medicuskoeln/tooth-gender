import pandas as pd
import numpy as np
from scipy.stats import mannwhitneyu, chi2_contingency, fisher_exact, shapiro, ttest_ind
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# --- Daten einlesen und auf Probandenebene reduzieren ---
df = pd.read_excel('roh1.xlsx')

# Jeder Proband hat mehrere Zeilen (Versuche) - demografische Daten sind pro Proband gleich
proband_cols = ['PROBAND', 'Geschlecht_w1_m2', 'Assistent1_Facharzt2',
                'Anz_SLT_4', 'Anz_DLT_4', 'Anz_MACI_4', 'Anz_GLIDE_CMAC_4', 'Anz_KingV_4']

df_prob = df[proband_cols].drop_duplicates(subset='PROBAND')

weiblich = df_prob[df_prob['Geschlecht_w1_m2'] == 1]
maennlich = df_prob[df_prob['Geschlecht_w1_m2'] == 2]

print(f"Probanden gesamt: {len(df_prob)}, weiblich: {len(weiblich)}, männlich: {len(maennlich)}")

# --- Kategorische Variable: Assistent vs Facharzt ---
def test_categorical(col, df_w, df_m):
    """Chi-Quadrat oder Fisher-Test für kategorische Variable."""
    vals_w = df_w[col].dropna()
    vals_m = df_m[col].dropna()
    categories = sorted(set(vals_w.tolist() + vals_m.tolist()))
    # Kreuztabelle
    ct = pd.DataFrame(index=categories, columns=['weiblich', 'männlich'])
    for cat in categories:
        ct.loc[cat, 'weiblich'] = (vals_w == cat).sum()
        ct.loc[cat, 'männlich'] = (vals_m == cat).sum()
    ct = ct.astype(int)
    # Test
    if ct.shape == (2, 2) and ct.values.min() < 5:
        _, p_val = fisher_exact(ct.values)
        test_name = "Fisher"
    else:
        _, p_val, _, _ = chi2_contingency(ct.values)
        test_name = "Chi²"
    return ct, p_val, test_name

# --- Kontinuierliche Variable ---
def test_continuous(col, df_w, df_m):
    """Mann-Whitney-U oder t-Test für kontinuierliche Variable."""
    vals_w = pd.to_numeric(df_w[col], errors='coerce').dropna()
    vals_m = pd.to_numeric(df_m[col], errors='coerce').dropna()
    if len(vals_w) < 3 or len(vals_m) < 3:
        return None, None, None, None, None
    # Normalverteilung prüfen
    try:
        p_shap_w = shapiro(vals_w)[1]
        p_shap_m = shapiro(vals_m)[1]
    except:
        p_shap_w, p_shap_m = 0, 0
    if p_shap_w > 0.05 and p_shap_m > 0.05:
        _, p_val = ttest_ind(vals_w, vals_m)
        test_name = "t-Test"
        stat_w = f"{vals_w.mean():.1f} ± {vals_w.std():.1f}"
        stat_m = f"{vals_m.mean():.1f} ± {vals_m.std():.1f}"
    else:
        _, p_val = mannwhitneyu(vals_w, vals_m, alternative='two-sided')
        test_name = "Mann-Whitney"
        iqr_w = vals_w.quantile(0.75) - vals_w.quantile(0.25)
        iqr_m = vals_m.quantile(0.75) - vals_m.quantile(0.25)
        stat_w = f"{vals_w.median():.1f} ({iqr_w:.1f})"
        stat_m = f"{vals_m.median():.1f} ({iqr_m:.1f})"
    return stat_w, stat_m, p_val, test_name, len(vals_w) + len(vals_m)

# --- Ergebnisse sammeln ---
results = []

# 1) Assistent vs Facharzt (kategorisch)
ct, p_val, test_name = test_categorical('Assistent1_Facharzt2', weiblich, maennlich)
n_ass_w = (weiblich['Assistent1_Facharzt2'] == 1).sum()
n_fa_w = (weiblich['Assistent1_Facharzt2'] == 2).sum()
n_ass_m = (maennlich['Assistent1_Facharzt2'] == 1).sum()
n_fa_m = (maennlich['Assistent1_Facharzt2'] == 2).sum()
results.append({
    'Variable': 'Assistent, n (%)',
    'Weiblich': f"{n_ass_w} ({100*n_ass_w/len(weiblich):.1f}%)",
    'Männlich': f"{n_ass_m} ({100*n_ass_m/len(maennlich):.1f}%)",
    'p-Wert': f"{p_val:.4f}",
    'Test': test_name
})
results.append({
    'Variable': 'Facharzt, n (%)',
    'Weiblich': f"{n_fa_w} ({100*n_fa_w/len(weiblich):.1f}%)",
    'Männlich': f"{n_fa_m} ({100*n_fa_m/len(maennlich):.1f}%)",
    'p-Wert': '',
    'Test': ''
})

# 2) Kontinuierliche Variablen
cont_vars = [
    ('Anz_SLT_4', 'Anzahl SLT'),
    ('Anz_DLT_4', 'Anzahl DLT'),
    ('Anz_MACI_4', 'Anzahl Mac'),
    ('Anz_GLIDE_CMAC_4', 'Anzahl Glide/CMAC'),
    ('Anz_KingV_4', 'Anzahl KingVision'),
]

for col, label in cont_vars:
    stat_w, stat_m, p_val, test_name, n = test_continuous(col, weiblich, maennlich)
    if stat_w is None:
        continue
    results.append({
        'Variable': label,
        'Weiblich': stat_w,
        'Männlich': stat_m,
        'p-Wert': f"{p_val:.4f}",
        'Test': test_name
    })

# --- Anaesthesiology-Stil Hilfsfunktion ---
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>' 
        f'<w:top w:val="{kwargs.get("top", "nil")}" w:sz="{kwargs.get("top_sz", "0")}" w:space="0" w:color="{kwargs.get("top_color", "000000")}"/>' 
        f'<w:bottom w:val="{kwargs.get("bottom", "nil")}" w:sz="{kwargs.get("bottom_sz", "0")}" w:space="0" w:color="{kwargs.get("bottom_color", "000000")}"/>' 
        f'<w:left w:val="nil" w:sz="0" w:space="0" w:color="000000"/>' 
        f'<w:right w:val="nil" w:sz="0" w:space="0" w:color="000000"/>' 
        f'</w:tcBorders>')
    tcPr.append(tcBorders)

# --- Word-Dokument erstellen ---
doc = Document()

# Querformat
section = doc.sections[-1]
section.orientation = WD_ORIENT.LANDSCAPE
new_width, new_height = section.page_height, section.page_width
section.page_width = new_width
section.page_height = new_height
section.left_margin = Cm(1.5)
section.right_margin = Cm(1.5)

# Titel im Anaesthesiology-Stil
title_p = doc.add_paragraph()
title_run = title_p.add_run('Tab 1')
title_run.bold = True
title_run.font.size = Pt(11)
title_run.font.name = 'Arial'
title_p.paragraph_format.space_after = Pt(2)

subtitle_p = doc.add_paragraph()
subtitle_run = subtitle_p.add_run(
    'Demographic characteristics and intubation experience of female and male participants'
)
subtitle_run.font.size = Pt(9)
subtitle_run.font.name = 'Arial'
subtitle_p.paragraph_format.space_after = Pt(6)

# Tabelle
table = doc.add_table(rows=1, cols=5)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = 'Variable'
hdr[1].text = f'Female (n={len(weiblich)})'
hdr[2].text = f'Male (n={len(maennlich)})'
hdr[3].text = 'P value'
hdr[4].text = 'Test'

for r in results:
    row = table.add_row().cells
    row[0].text = r['Variable']
    row[1].text = r['Weiblich']
    row[2].text = r['Männlich']
    row[3].text = r['p-Wert']
    row[4].text = r['Test']

# Schriftgröße und Font
for row_idx, row in enumerate(table.rows):
    for col_idx, cell in enumerate(row.cells):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
                run.font.name = 'Arial'
                # Header fett
                if row_idx == 0:
                    run.bold = True
                # P-Wert fett wenn < 0.05
                if col_idx == 3 and row_idx > 0 and run.text not in ['', '-']:
                    try:
                        pv = float(run.text)
                        if pv < 0.05:
                            run.bold = True
                    except:
                        pass

# Anaesthesiology-Stil: nur horizontale Linien
table.style = 'Normal Table'
for cell in table.rows[0].cells:
    set_cell_border(cell, top="single", top_sz="12", top_color="000000",
                    bottom="single", bottom_sz="8", bottom_color="000000")
for cell in table.rows[-1].cells:
    set_cell_border(cell, bottom="single", bottom_sz="12", bottom_color="000000")

for row in table.rows:
    for cell in row.cells:
        for p in cell.paragraphs:
            pf = p.paragraph_format
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)

# Fußnote
fn_p = doc.add_paragraph()
fn_run = fn_p.add_run(
    f"Total participants: {len(df_prob)} (female: {len(weiblich)}, male: {len(maennlich)}). "
    "Continuous variables presented as mean \u00b1 SD (t-test) or median (IQR) (Mann-Whitney U test). "
    "Categorical variables: n (%) (Chi\u00b2 / Fisher exact test). Bold P values indicate statistical significance (P < 0.05)."
)
fn_run.font.size = Pt(7)
fn_run.font.name = 'Arial'
fn_run.italic = True
fn_p.paragraph_format.space_before = Pt(4)

doc.save('Tab1.docx')
print("Tab 1 gespeichert in Tab1.docx")

# ==============================================================
# Tabelle 2: Kategorisierte Laryngoskop-Erfahrung + Assistent/Facharzt
# ==============================================================

# Kategorisierung der Erfahrungsvariablen
def kategorisiere(series, bins, labels):
    vals = pd.to_numeric(series, errors='coerce')
    return pd.cut(vals, bins=bins, labels=labels, right=True, include_lowest=True)

df_prob['SLT_kat'] = kategorisiere(df_prob['Anz_SLT_4'],
    bins=[-1, 500, 2000, 999999], labels=['≤500', '501–2000', '>2000'])
df_prob['DLT_kat'] = kategorisiere(df_prob['Anz_DLT_4'],
    bins=[-1, 10, 200, 999999], labels=['≤10', '11–200', '>200'])
df_prob['Mac_kat'] = kategorisiere(df_prob['Anz_MACI_4'],
    bins=[-1, 500, 1500, 999999], labels=['≤500', '501–1500', '>1500'])
df_prob['GlideCMAC_kat'] = kategorisiere(df_prob['Anz_GLIDE_CMAC_4'],
    bins=[-1, 10, 100, 999999], labels=['≤10', '11–100', '>100'])
df_prob['KingV_kat'] = kategorisiere(df_prob['Anz_KingV_4'],
    bins=[-0.5, 0, 10, 999999], labels=['0', '1–10', '>10'])

weiblich2 = df_prob[df_prob['Geschlecht_w1_m2'] == 1]
maennlich2 = df_prob[df_prob['Geschlecht_w1_m2'] == 2]

def test_cat_column(col, df_w, df_m):
    """Chi² oder Fisher-Test für kategorisierte Spalte. Gibt Kreuztabelle + p zurück."""
    vals_w = df_w[col].dropna()
    vals_m = df_m[col].dropna()
    cats = list(vals_w.cat.categories) if hasattr(vals_w, 'cat') else sorted(set(vals_w.tolist() + vals_m.tolist()))
    ct_data = []
    for cat in cats:
        n_w = (vals_w == cat).sum()
        n_m = (vals_m == cat).sum()
        ct_data.append((cat, n_w, n_m))
    ct_array = np.array([[r[1], r[2]] for r in ct_data])
    # Erwartete Häufigkeiten prüfen
    if ct_array.min() < 5 and ct_array.shape[0] == 2:
        _, p_val = fisher_exact(ct_array)
        test_name = "Fisher"
    else:
        _, p_val, _, _ = chi2_contingency(ct_array)
        test_name = "Chi²"
    return ct_data, cats, p_val, test_name

results2 = []

# Assistent vs Facharzt
n_ass_w2 = (weiblich2['Assistent1_Facharzt2'] == 1).sum()
n_fa_w2 = (weiblich2['Assistent1_Facharzt2'] == 2).sum()
n_ass_m2 = (maennlich2['Assistent1_Facharzt2'] == 1).sum()
n_fa_m2 = (maennlich2['Assistent1_Facharzt2'] == 2).sum()
_, p_ass, test_ass = test_categorical('Assistent1_Facharzt2', weiblich2, maennlich2)
n_w = len(weiblich2)
n_m = len(maennlich2)
results2.append({
    'header': 'Ausbildungsstand',
    'rows': [
        ('Assistent, n (%)', f"{n_ass_w2} ({100*n_ass_w2/n_w:.1f}%)", f"{n_ass_m2} ({100*n_ass_m2/n_m:.1f}%)", f"{p_ass:.4f}", test_ass),
        ('Facharzt, n (%)', f"{n_fa_w2} ({100*n_fa_w2/n_w:.1f}%)", f"{n_fa_m2} ({100*n_fa_m2/n_m:.1f}%)", '', ''),
    ]
})

# Kategorisierte Laryngoskop-Erfahrungen
kat_vars = [
    ('SLT_kat', 'Anz. SLT (kategorisiert)'),
    ('DLT_kat', 'Anz. DLT (kategorisiert)'),
    ('Mac_kat', 'Anz. Macintosh (kategorisiert)'),
    ('GlideCMAC_kat', 'Anz. GlideScope/CMAC (kategorisiert)'),
    ('KingV_kat', 'Anz. KingVision (kategorisiert)'),
]

for col, header in kat_vars:
    ct_data, cats, p_val, test_name = test_cat_column(col, weiblich2, maennlich2)
    rows = []
    for i, (cat, n_w_cat, n_m_cat) in enumerate(ct_data):
        pct_w = 100 * n_w_cat / n_w if n_w > 0 else 0
        pct_m = 100 * n_m_cat / n_m if n_m > 0 else 0
        rows.append((
            f"  {cat}, n (%)",
            f"{n_w_cat} ({pct_w:.1f}%)",
            f"{n_m_cat} ({pct_m:.1f}%)",
            f"{p_val:.4f}" if i == 0 else '',
            test_name if i == 0 else ''
        ))
    results2.append({'header': header, 'rows': rows})

# Word-Dokument Tabelle 2
doc2 = Document()

# Querformat
section2 = doc2.sections[-1]
section2.orientation = WD_ORIENT.LANDSCAPE
new_width2, new_height2 = section2.page_height, section2.page_width
section2.page_width = new_width2
section2.page_height = new_height2
section2.left_margin = Cm(1.5)
section2.right_margin = Cm(1.5)

# Titel
title_p2 = doc2.add_paragraph()
title_run2 = title_p2.add_run('Tab 1 (continued)')
title_run2.bold = True
title_run2.font.size = Pt(11)
title_run2.font.name = 'Arial'
title_p2.paragraph_format.space_after = Pt(2)

subtitle_p2 = doc2.add_paragraph()
subtitle_run2 = subtitle_p2.add_run(
    'Categorised intubation experience and training level of female and male participants'
)
subtitle_run2.font.size = Pt(9)
subtitle_run2.font.name = 'Arial'
subtitle_p2.paragraph_format.space_after = Pt(6)

table2 = doc2.add_table(rows=1, cols=5)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr2 = table2.rows[0].cells
hdr2[0].text = 'Variable'
hdr2[1].text = f'Female (n={n_w})'
hdr2[2].text = f'Male (n={n_m})'
hdr2[3].text = 'P value'
hdr2[4].text = 'Test'

for block in results2:
    # Header-Zeile (fett)
    row = table2.add_row().cells
    row[0].text = block['header']
    for cell in row:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
    # Detail-Zeilen
    for var, val_w, val_m, p_str, t_str in block['rows']:
        row = table2.add_row().cells
        row[0].text = var
        row[1].text = val_w
        row[2].text = val_m
        row[3].text = p_str
        row[4].text = t_str

# Schriftgröße und Font
for row_idx, row in enumerate(table2.rows):
    for col_idx, cell in enumerate(row.cells):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
                run.font.name = 'Arial'
                if row_idx == 0:
                    run.bold = True
                if col_idx == 3 and row_idx > 0 and run.text not in ['', '-']:
                    try:
                        pv = float(run.text)
                        if pv < 0.05:
                            run.bold = True
                    except:
                        pass

# Anaesthesiology-Stil: nur horizontale Linien
table2.style = 'Normal Table'
for cell in table2.rows[0].cells:
    set_cell_border(cell, top="single", top_sz="12", top_color="000000",
                    bottom="single", bottom_sz="8", bottom_color="000000")
for cell in table2.rows[-1].cells:
    set_cell_border(cell, bottom="single", bottom_sz="12", bottom_color="000000")

for row in table2.rows:
    for cell in row.cells:
        for p in cell.paragraphs:
            pf = p.paragraph_format
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)

# Fußnote
fn_p2 = doc2.add_paragraph()
fn_run2 = fn_p2.add_run(
    f"Total participants: {len(df_prob)} (female: {n_w}, male: {n_m}). "
    "All variables categorical: n (%). P values from Chi\u00b2 or Fisher exact test. "
    "Bold P values indicate statistical significance (P < 0.05)."
)
fn_run2.font.size = Pt(7)
fn_run2.font.name = 'Arial'
fn_run2.italic = True
fn_p2.paragraph_format.space_before = Pt(4)

doc2.save('Tab1_continued.docx')
print("Tab 1 (continued) gespeichert in Tab1_continued.docx")
