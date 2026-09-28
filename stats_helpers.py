"""Gemeinsame Statistik-Helfer für die Matched-Pair-Analysen.

- Benjamini-Hochberg-FDR-Adjustierung der p-Werte
- Linear-Mixed-Effects-Modell mit Random-Intercept pro Proband (PROBAND)
  für wiederholte Messungen pro Anwender
"""

from __future__ import annotations

import warnings
import numpy as np
import pandas as pd
from scipy.stats import wilcoxon, shapiro, ttest_rel


# ---------------------------------------------------------------------------
# Benjamini-Hochberg
# ---------------------------------------------------------------------------

def bh_fdr(pvals):
    """Benjamini-Hochberg FDR-Adjustierung.

    Akzeptiert Liste/Array mit p-Werten (NaN/None erlaubt) und liefert ein
    Array gleicher Länge mit adjustierten q-Werten zurück. NaN-Eingaben
    bleiben NaN.
    """
    p_raw = np.array([np.nan if v is None else v for v in pvals], dtype=float)
    q = np.full_like(p_raw, np.nan, dtype=float)
    mask = ~np.isnan(p_raw)
    if mask.sum() == 0:
        return q
    pv = p_raw[mask]
    n = pv.size
    order = np.argsort(pv)
    ranked = pv[order]
    raw_q = ranked * n / (np.arange(1, n + 1))
    # Monotonie von rechts nach links erzwingen
    raw_q = np.minimum.accumulate(raw_q[::-1])[::-1]
    raw_q = np.clip(raw_q, 0.0, 1.0)
    adj = np.empty(n)
    adj[order] = raw_q
    q[mask] = adj
    return q


# ---------------------------------------------------------------------------
# Format-Helfer
# ---------------------------------------------------------------------------

def fmt_p(p):
    """Einheitliche p-Wert-Darstellung."""
    if p is None:
        return '-'
    try:
        pv = float(p)
    except (TypeError, ValueError):
        return '-'
    if np.isnan(pv):
        return '-'
    if pv < 0.001:
        return '<0.001'
    return f"{pv:.3f}"


def format_p_cell(p_raw, p_fdr=None, p_mm=None):
    """Mehrzeilige Darstellung der drei p-Werte für eine Tabellenzelle.

    Zeile 1: roher p-Wert
    Zeile 2: 'qFDR = ...' (Benjamini-Hochberg)
    Zeile 3: 'pMM = ...'  (Mixed-Effects)
    """
    lines = [fmt_p(p_raw)]
    if p_fdr is not None:
        lines.append(f"qFDR={fmt_p(p_fdr)}")
    if p_mm is not None:
        lines.append(f"pMM={fmt_p(p_mm)}")
    return "\n".join(lines)


LEGEND_FDR_MEM = (
    "Each P cell shows three values stacked vertically: (1) raw P value from "
    "the Wilcoxon signed-rank test or paired t-test, (2) qFDR \u2014 Benjamini\u2013"
    "Hochberg false-discovery-rate adjusted P value across all comparisons in "
    "this table, (3) pMM \u2014 P value for the sex fixed effect from a linear "
    "mixed-effects model with random intercept per provider (PROBAND), which "
    "accounts for clustering of repeated measurements within the same "
    "anaesthetist. Bold P indicates raw P < 0.05; an asterisk (*) marks results "
    "that remain significant after FDR adjustment (qFDR < 0.05); a dagger (\u2020) "
    "marks results that remain significant in the mixed-effects model "
    "(pMM < 0.05)."
)


def p_str_to_float(s):
    """Wandelt formatierten p-String (z.B. '0.030', '<0.001', '-') in float um."""
    if s is None:
        return np.nan
    s = str(s).strip()
    if s in ('-', '', 'nan'):
        return np.nan
    if s.startswith('<'):
        return 0.0005
    try:
        return float(s)
    except ValueError:
        return np.nan


# ---------------------------------------------------------------------------
# Paarweise Tests (Wilcoxon / paired t)
# ---------------------------------------------------------------------------

def paired_test(vals_w, vals_m):
    """Wählt Wilcoxon oder paired t-test wie im Original.

    Gibt (p_value, used_test_label) zurück.
    """
    if len(vals_w) < 3:
        return np.nan, None
    diff = np.asarray(vals_w) - np.asarray(vals_m)
    try:
        p_shap = shapiro(diff)[1]
    except Exception:
        p_shap = 0.0
    if p_shap > 0.05:
        try:
            _, p = ttest_rel(vals_w, vals_m)
            return float(p), 't'
        except Exception:
            return np.nan, None
    try:
        _, p = wilcoxon(vals_w, vals_m)
        return float(p), 'w'
    except Exception:
        return np.nan, None


# ---------------------------------------------------------------------------
# Mixed-Effects-Modell
# ---------------------------------------------------------------------------

def _build_long(df_w, df_m, col, id_col='PROBAND'):
    rows = []
    for d, sex in [(df_w, 0), (df_m, 1)]:
        if d is None or len(d) == 0 or col not in d.columns or id_col not in d.columns:
            continue
        tmp = pd.DataFrame({
            'pid': d[id_col].values,
            'y': pd.to_numeric(d[col], errors='coerce').values,
            'sex': sex,
        })
        rows.append(tmp.dropna(subset=['y', 'pid']))
    if not rows:
        return None
    data = pd.concat(rows, ignore_index=True)
    return data


def mixed_effects_p(df_w, df_m, col, id_col='PROBAND'):
    """Linear-Mixed-Effects-Modell: y ~ sex + (1 | PROBAND).

    Liefert den p-Wert des Sex-Effekts (np.nan wenn Fit fehlschlägt oder
    zu wenig Daten / Cluster).
    """
    try:
        import statsmodels.formula.api as smf
    except ImportError:
        return np.nan

    data = _build_long(df_w, df_m, col, id_col=id_col)
    if data is None:
        return np.nan
    # Mindestanforderungen
    if len(data) < 10:
        return np.nan
    if data['pid'].nunique() < 3:
        return np.nan
    if data['sex'].nunique() < 2:
        return np.nan
    if data['y'].nunique() < 2:
        return np.nan

    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        try:
            md = smf.mixedlm('y ~ sex', data, groups=data['pid'])
            res = md.fit(reml=False, method='lbfgs', disp=False)
            p = float(res.pvalues.get('sex', np.nan))
            return p
        except Exception:
            try:
                # Fallback: REML + Powell
                md = smf.mixedlm('y ~ sex', data, groups=data['pid'])
                res = md.fit(reml=True, method='powell', disp=False)
                return float(res.pvalues.get('sex', np.nan))
            except Exception:
                return np.nan
