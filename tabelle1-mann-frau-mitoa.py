import pandas as pd
import numpy as np
from scipy.stats import mannwhitneyu, chi2_contingency, fisher_exact, shapiro, ttest_ind
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- Daten einlesen und auf Probandenebene reduzieren ---
df = pd.read_excel('roh1.xlsx')

# Jeder Proband hat mehrere Zeilen (Versuche) - demografische Daten sind pro Proband gleich
proband_cols = ['PROBAND', 'Geschlecht_w1_m2', 'Assistent1_Facharzt2',
                'Anz_SLT_4', 'Anz_DLT_4', 'Anz_MACI_4', 'Anz_GLIDE_CMAC_4', 'Anz_KingV_4',
                'NL_SLT_Easy_5', 'Nl_SLT_Diff_6', 'Nl_DLT_Easy_7', 'NI_DLT_Diff_8',
                'DA_SLT_Easy_9', 'DA_SLT_Diff_10', 'DA_DLT_Easy_11', 'DA_DLT_Diff_12',
                'Dent_Risk_Low_13', 'Dent_Risk_High_14', 'Dent_Damage_SLT_15', 'Dent_Damage_DLT_16']

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
    ('NL_SLT_Easy_5', 'NL SLT Easy'),
    ('Nl_SLT_Diff_6', 'NL SLT Diff'),
    ('Nl_DLT_Easy_7', 'NL DLT Easy'),
    ('NI_DLT_Diff_8', 'NI DLT Diff'),
    ('DA_SLT_Easy_9', 'DA SLT Easy'),
    ('DA_SLT_Diff_10', 'DA SLT Diff'),
    ('DA_DLT_Easy_11', 'DA DLT Easy'),
    ('DA_DLT_Diff_12', 'DA DLT Diff'),
    ('Dent_Risk_Low_13', 'Dent Risk Low'),
    ('Dent_Risk_High_14', 'Dent Risk High'),
    ('Dent_Damage_SLT_15', 'Dent Damage SLT'),
    ('Dent_Damage_DLT_16', 'Dent Damage DLT'),
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

# --- Word-Dokument erstellen ---
doc = Document()
doc.add_heading('Tabelle 1: Vergleich Weiblich vs. Männlich', level=1)
doc.add_paragraph(
    f"Probanden gesamt: {len(df_prob)} (weiblich: {len(weiblich)}, männlich: {len(maennlich)})\n"
    "Kontinuierliche Variablen: Mittelwert ± SD (t-Test) oder Median (IQR) (Mann-Whitney-U).\n"
    "Kategorische Variablen: n (%) (Chi²/Fisher-Test)."
)

# Tabelle
table = doc.add_table(rows=1, cols=5)
table.style = 'Light Shading Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = 'Variable'
hdr[1].text = f'Weiblich (n={len(weiblich)})'
hdr[2].text = f'Männlich (n={len(maennlich)})'
hdr[3].text = 'p-Wert'
hdr[4].text = 'Test'

for r in results:
    row = table.add_row().cells
    row[0].text = r['Variable']
    row[1].text = r['Weiblich']
    row[2].text = r['Männlich']
    row[3].text = r['p-Wert']
    row[4].text = r['Test']

# Schriftgröße verkleinern
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)

# Querformat
section = doc.sections[-1]
section.orientation = WD_ORIENT.LANDSCAPE
new_width, new_height = section.page_height, section.page_width
section.page_width = new_width
section.page_height = new_height

doc.save('Tabelle1_Mann_vs_Frau.docx')
print("Tabelle 1 gespeichert in Tabelle1_Mann_vs_Frau.docx")

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
doc2.add_heading('Tabelle 2: Kategorisierte Erfahrung & Ausbildungsstand – Weiblich vs. Männlich', level=1)
doc2.add_paragraph(
    f"Probanden gesamt: {len(df_prob)} (weiblich: {n_w}, männlich: {n_m})\n"
    "Alle Variablen kategorisch: n (%). Signifikanztest: Chi² oder Fisher-Exakt-Test.\n\n"
    "Kategorien Laryngoskop-Erfahrung:\n"
    "  SLT: ≤500 / 501–2000 / >2000\n"
    "  DLT: ≤10 / 11–200 / >200\n"
    "  Macintosh: ≤500 / 501–1500 / >1500\n"
    "  GlideScope/CMAC: ≤10 / 11–100 / >100\n"
    "  KingVision: 0 / 1–10 / >10"
)

table2 = doc2.add_table(rows=1, cols=5)
table2.style = 'Light Shading Accent 1'
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr2 = table2.rows[0].cells
hdr2[0].text = 'Variable'
hdr2[1].text = f'Weiblich (n={n_w})'
hdr2[2].text = f'Männlich (n={n_m})'
hdr2[3].text = 'p-Wert'
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

# Schriftgröße
for row in table2.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)

# Querformat
section2 = doc2.sections[-1]
section2.orientation = WD_ORIENT.LANDSCAPE
new_width2, new_height2 = section2.page_height, section2.page_width
section2.page_width = new_width2
section2.page_height = new_height2

doc2.save('Tabelle2_Erfahrung_kategorisiert.docx')
print("Tabelle 2 gespeichert in Tabelle2_Erfahrung_kategorisiert.docx")
