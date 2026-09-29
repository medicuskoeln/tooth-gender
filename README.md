# tooth-gender

Analysis code for the manuscript:

> **Dental force and other intubation performance metrics during tracheal intubation: sex-related differences and the effect of videolaryngoscopy in a simulated airway model**
> Katharina Hardt1*, Marc Schieren1, Jonas Weber1, Axel Schmutz2, Torsten Loop2, Manfred Staat3, Karl-Heinz Gatzweiler3, Frank Wappler1, Jerome Defosse1. *[Journal]*, [Year]. DOI: [to be added]

All statistical analyses, tables and figures of the manuscript were produced with the Python scripts in this repository.

## Requirements

Python ≥ 3.9 and the packages listed in `requirements.txt`:

```bash
pip install -r requirements.txt
```

## Input data

The scripts expect the raw dataset as an Excel file in the working directory (one row per intubation attempt). The dataset is not included in this repository; it is available from the corresponding author upon reasonable request.

Main variables:

| Column | Description | Coding |
|---|---|---|
| `PROBAND` | Participant ID (pseudonymised) | integer |
| `Geschlecht_w1_m2` | Sex of participant | 1 = female, 2 = male |
| `Ausbild_WA1_FA2_OA3` | Level of training | 1 = resident, 2 = specialist, 3 = consultant |
| `Assistent1_Facharzt2` | Training level, dichotomised | 1 = resident, 2 = specialist |
| `atemweg` | Airway scenario | 1 = standard, 2 = difficult |
| `Laryng_diff_1Mac_2Glide_3King_4CMAC` | Laryngoscope | 1 = Macintosh, 2 = GlideScope, 3 = King Vision, 4 = C-MAC |
| `hyperanguliert` | Hyperangulated blade | 0 = no, 1 = yes |
| `tubus` | Tube type | [coding] |
| `res_11_max`, `res_12_max`, `res_21_max`, `res_22_max` | Maximum value per tooth (FDI 11, 12, 21, 22) | N |
| `res_max_alle_zaehne` | Maximum value across all teeth | N |
| `Intubationsdauer` | Intubation time | s |
| `Sicht_Cormack` | Cormack–Lehane grade | 1–4 |
| `BURP` | BURP manoeuvre applied | [coding] |
| `Anz_*`, `*_kat` | Prior experience with devices (count / category) | – |
| `NL_*`, `DA_*`, `Dent_*`, `Empf_Schwierig` | Questionnaire items | [scale] |

## Scripts

Run each script from the repository root: `python <script>.py`. Output is written to the working directory as `.docx` and `.png` files; box plots and histograms go to `plots/`.

| Manuscript item | Script(s) |
|---|---|
| Table 1 – participant characteristics | `tabelle1-mann-frau.py`, `tabelle1-mann-frau-mitoa.py`, `tabelle1-3gruppen.py`, `tabelle1-3gruppen-en.py`, `tabelle1-oa-final.py` |
| Table 2 – matched-pair outcomes | `tabelle2-outcome-matched-FOA.py`, `tabelle2-outcome-matched-getrennt.py` |
| Table 3 – outcomes by laryngoscope | `tabelle3-akt-outcome-std-laryng_FOA.py`, `tabelle3-akt-outcome-std-laryng_OA.py` |
| Outcomes by airway scenario | `tabelle-outcome-airway.py`, `tabelle-outcome-matched-fao.py`, `tabelle-outcome-matched-getrennt.py` |
| Difficulty by laryngoscope (matched) | `tabelle-schwierigkeit-laryng-matched.py` |
| Figures 1–4 | `figures_plot_Foa.py`, `figures_plot_oa.py` |
| Figure 5 – tooth heat map | `figures_heatmap_v1.py`, `figures_heatmap_FOA.py`, `figures_heatmap_OA.py` |
| Matched-pair analyses (supplementary) | `neu-gen-aktuell-fa-oa.py`, `neu-gen-4-nach Laryngoskop sortiert Kopie.py`, `neu-gen-hyperanguliert2 Kopie.py` |
| Shared statistical functions | `stats_helpers.py` |

## Statistical methods (overview)

- Normality: Shapiro–Wilk test
- Group comparisons: t-test / Mann–Whitney U, χ² / Fisher's exact test, Kruskal–Wallis / one-way ANOVA
- Matched pairs (female vs. male): paired t-test / Wilcoxon signed-rank test
- Repeated measurements per participant: linear mixed-effects model with random intercept per participant (`statsmodels`)
- Multiple testing: Benjamini–Hochberg false discovery rate

Details are given in `stats_helpers.py` and in the Methods section of the manuscript.

## License

Code: MIT License (see `LICENSE`).

## Citation

If you use this code, please cite the article above and the archived version of this repository: [Zenodo DOI].
