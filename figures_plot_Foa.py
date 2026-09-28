import pandas as pd
import numpy as np
from scipy.stats import wilcoxon, shapiro, ttest_rel
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import tempfile

# =====================================================================
# DATA LOADING & MATCHING (same as Tab 2 / Tab 3)
# =====================================================================
df = pd.read_excel('roh1.xlsx')
df = df[df['Geschlecht_w1_m2'].isin([1, 2])]

# OA und FA zusammenfassen: OA (3) → FA (2) => FAO-Gruppe
df['Ausbild_WA1_FA2_OA3'] = df['Ausbild_WA1_FA2_OA3'].replace({3: 2})

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
print(f"Total matched pairs: {len(df_w)}")

# Split by airway
df_w_std = df_w[df_w['atemweg'] == 1].reset_index(drop=True)
df_m_std = df_m[df_m['atemweg'] == 1].reset_index(drop=True)
df_w_dif = df_w[df_w['atemweg'] == 2].reset_index(drop=True)
df_m_dif = df_m[df_m['atemweg'] == 2].reset_index(drop=True)

laryng = {1: 'Macintosh', 2: 'GlideScope', 3: 'King Vision', 4: 'C-MAC'}

# =====================================================================
# STYLE SETTINGS – Anaesthesiology journal
# =====================================================================
COLOR_F = '#D4738C'   # muted rose for female
COLOR_M = '#5B8FA8'   # muted steel blue for male
plt.rcParams.update({
    'font.family': 'Arial',
    'font.size': 9,
    'axes.linewidth': 0.8,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'figure.facecolor': 'white',
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'savefig.bbox': 'tight',
})

# =====================================================================
# HELPER: Significance test on matched pairs
# =====================================================================
def matched_p(col, dw, dm):
    vw = pd.to_numeric(dw[col], errors='coerce')
    vm = pd.to_numeric(dm[col], errors='coerce')
    mask = (~vw.isna()) & (~vm.isna())
    vw, vm = vw[mask], vm[mask]
    if len(vw) < 3:
        return 1.0
    diff = vw.values - vm.values
    try:
        p_shap = shapiro(diff)[1]
    except:
        p_shap = 0
    if p_shap > 0.05:
        _, p_val = ttest_rel(vw, vm)
    else:
        try:
            _, p_val = wilcoxon(vw, vm)
        except:
            return 1.0
    return p_val

def sig_stars(p):
    if p < 0.001:
        return '***'
    elif p < 0.01:
        return '**'
    elif p < 0.05:
        return '*'
    return 'n.s.'

def add_sig_bracket(ax, x1, x2, y, p, h=0.02, lw=0.8):
    """Draw a significance bracket between two box positions."""
    stars = sig_stars(p)
    ax.plot([x1, x1, x2, x2], [y, y + h, y + h, y], color='black', lw=lw)
    ax.text((x1 + x2) / 2, y + h, stars, ha='center', va='bottom', fontsize=8, fontweight='bold')

def clean_vals(col, d):
    return pd.to_numeric(d[col], errors='coerce').dropna()

# =====================================================================
# FIG 1: Intubation Duration – Standard vs Difficult Airway
# =====================================================================
fig1, (ax1a, ax1b) = plt.subplots(1, 2, figsize=(7, 4), sharey=True)

for ax, dw, dm, title_txt in [
    (ax1a, df_w_std, df_m_std, 'Standard airway'),
    (ax1b, df_w_dif, df_m_dif, 'Difficult airway')
]:
    vals_f = clean_vals('Intubationsdauer', dw)
    vals_m = clean_vals('Intubationsdauer', dm)

    bp = ax.boxplot(
        [vals_f, vals_m], positions=[1, 2], widths=0.55,
        patch_artist=True, showfliers=False,
        medianprops=dict(color='black', lw=1.5),
        whiskerprops=dict(color='black', lw=0.8),
        capprops=dict(color='black', lw=0.8),
    )
    bp['boxes'][0].set_facecolor(COLOR_F)
    bp['boxes'][0].set_alpha(0.7)
    bp['boxes'][1].set_facecolor(COLOR_M)
    bp['boxes'][1].set_alpha(0.7)

    # Individual data points (jittered)
    np.random.seed(42)
    jitter_f = np.random.normal(1, 0.06, len(vals_f))
    jitter_m = np.random.normal(2, 0.06, len(vals_m))
    ax.scatter(jitter_f, vals_f, color=COLOR_F, alpha=0.15, s=6, zorder=0, edgecolors='none')
    ax.scatter(jitter_m, vals_m, color=COLOR_M, alpha=0.15, s=6, zorder=0, edgecolors='none')

    ax.set_xticks([1, 2])
    ax.set_xticklabels([f'Female\n(n={len(vals_f)})', f'Male\n(n={len(vals_m)})'])
    ax.set_title(title_txt, fontweight='bold', fontsize=10)

    # Significance bracket
    p = matched_p('Intubationsdauer', dw, dm)
    ymax = max(vals_f.quantile(0.95), vals_m.quantile(0.95))
    add_sig_bracket(ax, 1, 2, ymax * 1.05, p)
    ax.set_ylim(0, ymax * 1.25)

ax1a.set_ylabel('Intubation duration (s)', fontsize=10)
fig1.suptitle('Fig 1.  Intubation duration by sex and airway difficulty', fontsize=11, fontweight='bold', y=1.02)

# n-pair annotation
fig1.text(0.5, -0.02,
    f'Matched pairs: standard n = {len(df_w_std)}, difficult n = {len(df_w_dif)}. '
    '* P < 0.05, ** P < 0.01, *** P < 0.001 (paired test).',
    ha='center', fontsize=7, fontstyle='italic')

fig1.tight_layout()
fig1_path = os.path.join(tempfile.gettempdir(), 'fig1_intubation.png')
fig1.savefig(fig1_path, dpi=300, bbox_inches='tight')
plt.close(fig1)
print("Fig 1 saved")

# =====================================================================
# FIG 2: Dental Force by Individual Tooth – Standard vs Difficult
# =====================================================================
teeth = [
    ('res_12_max', 'Tooth 12'),
    ('res_11_max', 'Tooth 11'),
    ('res_21_max', 'Tooth 21'),
    ('res_22_max', 'Tooth 22'),
]

fig2, axes2 = plt.subplots(2, 4, figsize=(12, 6), sharey='row')

for row_idx, (dw, dm, aw_label) in enumerate([
    (df_w_std, df_m_std, 'Standard airway'),
    (df_w_dif, df_m_dif, 'Difficult airway')
]):
    for col_idx, (col, tooth_label) in enumerate(teeth):
        ax = axes2[row_idx, col_idx]
        vals_f = clean_vals(col, dw)
        vals_m = clean_vals(col, dm)

        bp = ax.boxplot(
            [vals_f, vals_m], positions=[1, 2], widths=0.55,
            patch_artist=True, showfliers=False,
            medianprops=dict(color='black', lw=1.5),
            whiskerprops=dict(color='black', lw=0.8),
            capprops=dict(color='black', lw=0.8),
        )
        bp['boxes'][0].set_facecolor(COLOR_F)
        bp['boxes'][0].set_alpha(0.7)
        bp['boxes'][1].set_facecolor(COLOR_M)
        bp['boxes'][1].set_alpha(0.7)

        ax.set_xticks([1, 2])
        ax.set_xticklabels(['F', 'M'], fontsize=8)

        if row_idx == 0:
            ax.set_title(tooth_label, fontweight='bold', fontsize=9)

        # Significance bracket
        p = matched_p(col, dw, dm)
        ymax = max(vals_f.quantile(0.95), vals_m.quantile(0.95))
        add_sig_bracket(ax, 1, 2, ymax * 1.05, p, lw=0.6)
        ax.set_ylim(0, ymax * 1.35)

        if col_idx == 0:
            ax.set_ylabel(f'{aw_label}\nForce (N)', fontsize=8)

fig2.suptitle('Fig 2.  Dental force on individual teeth by sex and airway difficulty',
              fontsize=11, fontweight='bold', y=1.02)

# Legend
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor=COLOR_F, alpha=0.7, label='Female'),
                   Patch(facecolor=COLOR_M, alpha=0.7, label='Male')]
fig2.legend(handles=legend_elements, loc='upper right', fontsize=8, frameon=False,
            bbox_to_anchor=(0.98, 0.98))

fig2.text(0.5, -0.01,
    f'Matched pairs: standard n = {len(df_w_std)}, difficult n = {len(df_w_dif)}. '
    '* P < 0.05, ** P < 0.01, *** P < 0.001 (paired test).',
    ha='center', fontsize=7, fontstyle='italic')

fig2.tight_layout()
fig2_path = os.path.join(tempfile.gettempdir(), 'fig2_teeth.png')
fig2.savefig(fig2_path, dpi=300, bbox_inches='tight')
plt.close(fig2)
print("Fig 2 saved")

# =====================================================================
# FIG 3: Intubation Duration by Laryngoscope Type (from Tab 3)
# =====================================================================
fig3, axes3 = plt.subplots(2, 4, figsize=(13, 8), sharey='row')

for row_idx, (dw_aw, dm_aw, aw_label) in enumerate([
    (df_w_std, df_m_std, 'Standard airway'),
    (df_w_dif, df_m_dif, 'Difficult airway')
]):
    for col_idx, (lv, lname) in enumerate(laryng.items()):
        ax = axes3[row_idx, col_idx]
        w_sub = dw_aw[dw_aw['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv]
        m_sub = dm_aw[dm_aw['Laryng_diff_1Mac_2Glide_3King_4CMAC'] == lv]

        vals_f = clean_vals('Intubationsdauer', w_sub)
        vals_m = clean_vals('Intubationsdauer', m_sub)

        if len(vals_f) < 3 or len(vals_m) < 3:
            ax.text(0.5, 0.5, 'n too small', ha='center', va='center', transform=ax.transAxes)
            continue

        bp = ax.boxplot(
            [vals_f, vals_m], positions=[1, 2], widths=0.55,
            patch_artist=True, showfliers=False,
            medianprops=dict(color='black', lw=1.5),
            whiskerprops=dict(color='black', lw=0.8),
            capprops=dict(color='black', lw=0.8),
        )
        bp['boxes'][0].set_facecolor(COLOR_F)
        bp['boxes'][0].set_alpha(0.7)
        bp['boxes'][1].set_facecolor(COLOR_M)
        bp['boxes'][1].set_alpha(0.7)

        ax.set_xticks([1, 2])
        ax.set_xticklabels([f'F\n({len(vals_f)})', f'M\n({len(vals_m)})'], fontsize=7)

        if row_idx == 0:
            ax.set_title(lname, fontweight='bold', fontsize=9)

        # Significance
        p = matched_p('Intubationsdauer', w_sub.reset_index(drop=True), m_sub.reset_index(drop=True))
        ymax = max(vals_f.quantile(0.95), vals_m.quantile(0.95))
        add_sig_bracket(ax, 1, 2, ymax * 1.05, p, lw=0.6)
        ax.set_ylim(0, ymax * 1.4)

        if col_idx == 0:
            ax.set_ylabel(f'{aw_label}\nDuration (s)', fontsize=8)

fig3.suptitle('Fig 3.  Intubation duration by laryngoscope type, sex, and airway difficulty',
              fontsize=11, fontweight='bold', y=1.02)
fig3.legend(handles=legend_elements, loc='upper right', fontsize=8, frameon=False,
            bbox_to_anchor=(0.98, 0.98))
fig3.text(0.5, -0.01,
    'Matched pairs per laryngoscope. * P < 0.05, ** P < 0.01, *** P < 0.001 (paired test).',
    ha='center', fontsize=7, fontstyle='italic')

fig3.subplots_adjust(hspace=0.4, wspace=0.25)
fig3_path = os.path.join(tempfile.gettempdir(), 'fig3_laryng_duration.png')
fig3.savefig(fig3_path, dpi=300, bbox_inches='tight')
plt.close(fig3)
print("Fig 3 saved")

# =====================================================================
# FIG 4: C&L Grade and BURP – Stacked Proportions
# =====================================================================
fig4, axes4 = plt.subplots(1, 2, figsize=(8, 4))

# Panel A: C&L Grade distribution (Standard + Difficult)
for panel, (dw_aw, dm_aw, aw_label) in enumerate([
    (df_w_std, df_m_std, 'Standard'),
    (df_w_dif, df_m_dif, 'Difficult')
]):
    ax = axes4[panel]
    grades = [1, 2, 3, 4]
    grade_colors = ['#BFDFBF', '#FFD9A0', '#FFB3B3', '#D4A0A0']  # green-yellow-red scheme

    vf = pd.to_numeric(dw_aw['Sicht_Cormack'], errors='coerce').dropna()
    vm = pd.to_numeric(dm_aw['Sicht_Cormack'], errors='coerce').dropna()

    # Count proportions
    pct_f = [(vf == g).sum() / len(vf) * 100 for g in grades]
    pct_m = [(vm == g).sum() / len(vm) * 100 for g in grades]

    bar_width = 0.35
    x = np.array([0, 1])

    # Stacked bars
    bottom_f, bottom_m = 0, 0
    for g_idx, g in enumerate(grades):
        ax.bar(0, pct_f[g_idx], bottom=bottom_f, width=bar_width,
               color=grade_colors[g_idx], edgecolor='white', lw=0.5)
        ax.bar(1, pct_m[g_idx], bottom=bottom_m, width=bar_width,
               color=grade_colors[g_idx], edgecolor='white', lw=0.5)
        # Label if ≥ 5%
        if pct_f[g_idx] >= 5:
            ax.text(0, bottom_f + pct_f[g_idx]/2, f'{pct_f[g_idx]:.0f}%',
                    ha='center', va='center', fontsize=7, fontweight='bold')
        if pct_m[g_idx] >= 5:
            ax.text(1, bottom_m + pct_m[g_idx]/2, f'{pct_m[g_idx]:.0f}%',
                    ha='center', va='center', fontsize=7, fontweight='bold')
        bottom_f += pct_f[g_idx]
        bottom_m += pct_m[g_idx]

    ax.set_xticks([0, 1])
    ax.set_xticklabels([f'Female\n(n={len(vf)})', f'Male\n(n={len(vm)})'], fontsize=8)
    ax.set_ylabel('Proportion (%)', fontsize=9)
    ax.set_title(f'{aw_label} airway', fontweight='bold', fontsize=10)
    ax.set_ylim(0, 115)

    # P value annotation
    p = matched_p('Sicht_Cormack', dw_aw, dm_aw)
    stars = sig_stars(p)
    p_txt = f'P = {p:.3f}' if p >= 0.001 else 'P < 0.001'
    ax.text(0.5, 108, f'{p_txt} {stars}', ha='center', fontsize=8, fontweight='bold')

# C&L legend
from matplotlib.patches import Patch as Patch2
cl_legend = [Patch2(facecolor=grade_colors[i], label=f'C&L {g}') for i, g in enumerate(grades)]
fig4.legend(handles=cl_legend, loc='upper right', fontsize=7, frameon=False,
            ncol=1, bbox_to_anchor=(0.99, 0.95))

fig4.suptitle('Fig 4.  Cormack–Lehane grade distribution by sex and airway difficulty',
              fontsize=11, fontweight='bold', y=1.02)
fig4.text(0.5, -0.03,
    f'Matched pairs: standard n = {len(df_w_std)}, difficult n = {len(df_w_dif)}. '
    'P from paired test on ordinal grades.',
    ha='center', fontsize=7, fontstyle='italic')

fig4.tight_layout()
fig4_path = os.path.join(tempfile.gettempdir(), 'fig4_cl_grade.png')
fig4.savefig(fig4_path, dpi=300, bbox_inches='tight')
plt.close(fig4)
print("Fig 4 saved")

# =====================================================================
# ASSEMBLE figures.docx
# =====================================================================
doc = Document()
section = doc.sections[-1]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(2)
section.right_margin = Cm(2)

def add_figure(doc, img_path, fig_label, caption, width_inches=6.5):
    """Add a figure with caption in Anaesthesiology style."""
    # Figure label
    p_label = doc.add_paragraph()
    p_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_label = p_label.add_run(fig_label)
    run_label.bold = True
    run_label.font.size = Pt(10)
    run_label.font.name = 'Arial'
    p_label.paragraph_format.space_after = Pt(2)

    # Image
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_inches))
    p_img.paragraph_format.space_after = Pt(4)

    # Caption
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_cap = p_cap.add_run(caption)
    run_cap.font.size = Pt(8)
    run_cap.font.name = 'Arial'
    run_cap.italic = True
    p_cap.paragraph_format.space_after = Pt(20)

# --- Fig 1 ---
add_figure(doc, fig1_path,
    'Fig 1. Intubation duration by sex and airway difficulty',
    'Boxplots of intubation duration (seconds) for female and male anaesthesiologists '
    f'in matched pairs (standard airway: n = {len(df_w_std)} pairs; '
    f'difficult airway: n = {len(df_w_dif)} pairs). '
    'Pairs matched on education level (residents/specialists/attendings), laryngoscope type, '
    'tube type, and airway difficulty. Boxes represent interquartile range with median line; '
    'whiskers extend to 1.5 × IQR. '
    '* P < 0.05, ** P < 0.01, *** P < 0.001 (Wilcoxon signed-rank test or paired t-test '
    'depending on distribution of paired differences).')

doc.add_page_break()

# --- Fig 2 ---
add_figure(doc, fig2_path,
    'Fig 2. Dental force on individual incisors by sex and airway difficulty',
    'Boxplots of maximum force (N) exerted on each individual incisor (teeth 12, 11, 21, 22) '
    'during intubation. Upper row: standard airway; lower row: difficult airway. '
    f'Matched pairs: standard n = {len(df_w_std)}, difficult n = {len(df_w_dif)}. '
    'Matching and statistical testing as in Fig 1. F = female, M = male.',
    width_inches=6.5)

doc.add_page_break()

# --- Fig 3 ---
add_figure(doc, fig3_path,
    'Fig 3. Intubation duration by laryngoscope type, sex, and airway difficulty',
    'Boxplots of intubation duration stratified by laryngoscope (Macintosh, GlideScope, '
    'King Vision, C-MAC). Upper row: standard airway; lower row: difficult airway. '
    'Pairs matched on education level (specialists and attendings combined), laryngoscope, tube type, and airway. '
    'Sample sizes shown below each box. '
    '* P < 0.05, ** P < 0.01, *** P < 0.001 (paired test).',
    width_inches=6.5)

doc.add_page_break()

# --- Fig 4 ---
add_figure(doc, fig4_path,
    'Fig 4. Cormack–Lehane grade distribution by sex and airway difficulty',
    'Stacked bar charts showing the proportional distribution of Cormack–Lehane grades '
    '(1–4) for female and male anaesthesiologists. Left panel: standard airway; '
    f'right panel: difficult airway. Matched pairs: standard n = {len(df_w_std)}, '
    f'difficult n = {len(df_w_dif)}. '
    'P values from paired test on ordinal grade data.',
    width_inches=5.5)

doc.save('figures.docx')
print("\n>>> Saved: figures.docx")
print("4 figures: intubation duration, dental force by tooth, duration by laryngoscope, C&L grades")
