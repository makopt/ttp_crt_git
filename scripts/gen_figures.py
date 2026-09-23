#!/usr/bin/env python3
"""Publication figures for the TTP Systematic & Critical Review.

Data figures only. The schematic figures (problem anatomy, taxonomy DAG,
PRISMA flow, recommendations roadmap) are drawn in TikZ (the .tex files in figs/).

Source data lives in data/; output goes to figs/. Paths are resolved from this
script's own location, so it can be run from anywhere:  python3 gen_figures.py
"""
import os, re, glob, math
from pathlib import Path
import numpy as np
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import gridspec

ELSEVIER_DIR = Path(__file__).resolve().parent.parent   # <repo> root
REF_TABS = ELSEVIER_DIR / "data"                  # source data

DATA_XLSX = REF_TABS / "ttp_crt_data.xlsx"
OUT  = ELSEVIER_DIR / "figs"
OUT.mkdir(parents=True, exist_ok=True)

# ----------------------------------------------------------------------
#  data access: every table the figures need is built from the files in data/
#    ttp_crt_data.xlsx (sheets Corpus, Venues, Reporting_Audit, Theme_Counts, CatABC_Objectives)
# ----------------------------------------------------------------------
import scipy.stats as ss

_WB = openpyxl.load_workbook(DATA_XLSX, data_only=True)

def _sheet(name):
    return [tuple("" if v is None else v for v in r)
            for r in _WB[name].iter_rows(values_only=True) if any(v is not None for v in r)]

def _corpus():
    return _sheet("Corpus")

def _year_variant(mc):
    iy, iv = mc[0].index("Year"), mc[0].index("Variant(s)")
    years = sorted({r[iy] for r in mc[1:]})
    vs = VORDER + ["Survey"]
    out = [("Year",) + tuple(vs)]
    for y in years:
        c = {v: 0 for v in vs}
        for r in mc[1:]:
            if r[iy] == y:
                for v in str(r[iv]).split(","):
                    if v.strip() in c: c[v.strip()] += 1
        out.append((y,) + tuple(c[v] for v in vs))
    return out

def _coverage(mc):
    ih, iv, ifam = 0, mc[0].index("Variant(s)"), mc[0].index("Family (L1)")
    body = [r for r in mc[1:] if str(r[iv]).strip() != "Survey"]
    hdr = ("Algorithm family \\ Variant",) + tuple(VORDER) + ("Row total",)
    out = [hdr]
    for f in FAM_ERA:
        sel = [r for r in body if r[ifam] == f]
        if not sel: continue
        cells = [sum(1 for r in sel if v in [x.strip() for x in str(r[iv]).split(",")]) for v in VORDER]
        out.append((f,) + tuple(cells) + (len(sel),))
    return out

def _venues():
    return _sheet("Venues")

def _catabc():
    rs = list(_WB["CatABC_Objectives"].iter_rows(values_only=True))
    algos = [str(x) for x in rs[1][9:18]]
    recs = [r for r in rs[2:] if r[1] in ("A", "B", "C") and isinstance(r[9], (int, float))]
    return algos, recs

def _rdi_tables():
    """Per-instance RDI, mean Friedman rank, best-in-pool count (Section 5.3 of the paper)."""
    algos, recs = _catabc()
    V = np.array([[r[9 + i] for i in range(len(algos))] for r in recs], dtype=float)
    rng = V.max(1) - V.min(1)
    R = (V.max(1, keepdims=True) - V) / np.where(rng == 0, 1, rng)[:, None] * 100
    ranks = np.array([ss.rankdata(-v, method="average") for v in V])
    best = {a: 0 for a in algos}
    for v, rg in zip(V, rng):
        if rg > 0: best[algos[int(np.argmax(v))]] += 1
    cats = np.array([r[1] for r in recs])
    hdr = ("Algorithm", "Cat A avg RDI", "Cat B avg RDI", "Cat C avg RDI", "Overall avg RDI",
           "Mean Friedman rank", "# best-known (of 60)")
    tab = [hdr]
    for i, a in enumerate(algos):
        tab.append((a,) + tuple(round(float(R[cats == c, i].mean()), 1) for c in "ABC")
                   + (round(float(R[:, i].mean()), 1), round(float(ranks[:, i].mean()), 2), best[a]))
    return tab, len(recs)

_MC = _corpus()
_TAB10, _NPOOL = _rdi_tables()

def rows(sheet):
    if sheet == "Master_Corpus":        return _MC
    if sheet == "Fig8_Year_Variant":    return _year_variant(_MC)
    if sheet == "Fig5_Coverage_Matrix": return _coverage(_MC)
    if sheet == "Fig9_Venues":          return _venues()
    if sheet == "Tab10_TTP1_RDI":       return _TAB10
    if sheet == "Lit_CatABC_RDI":       return [None] * (_NPOOL + 1)
    if sheet == "SW_Themes":
        return _sheet("Theme_Counts")
    if sheet == "Tab16_ReportingAudit":
        return _sheet("Reporting_Audit")
    if sheet == "Tab1_Taxonomy":
        return _sheet("Taxonomy")
    if sheet == "Tab2_Unexplored":
        return _sheet("Unexplored")
    raise KeyError(sheet)

# ----------------------------------------------------------------------
#  design system
# ----------------------------------------------------------------------
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10.5,
    "axes.titlesize": 12, "axes.titleweight": "bold", "axes.titlepad": 10,
    "axes.labelsize": 10.5, "axes.labelweight": "regular",
    "axes.edgecolor": "#8A8A8A", "axes.linewidth": 0.8,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.axisbelow": True,
    "grid.color": "#E4E4E4", "grid.linewidth": 0.7,
    "xtick.color": "#333333", "ytick.color": "#333333",
    "xtick.labelsize": 9.5, "ytick.labelsize": 9.5,
    "legend.fontsize": 9, "legend.frameon": False,
    "figure.dpi": 120, "savefig.dpi": 300, "savefig.bbox": "tight",
    "savefig.facecolor": "white",
})
INK   = "#4A6FA0"           # medium slate-blue, used for emphasis lines/text -- readable but not dark
MUTED = "#9096A0"
# ------------------------------------------------------------------
# Light, pastel-muted qualitative palette (ColorBrewer "Set2"-family
# hues): every colour is assigned to exactly one meaning and reused with
# that meaning in every figure that needs it, so colour is never
# decorative. Hues are spaced well apart within each figure's own set of
# colours (no two adjacent categories share a close hue); lightness is
# kept high and saturation moderate throughout -- no dark, neon, or
# fluorescent tones anywhere.
# ------------------------------------------------------------------
# variant colours -- fixed across every figure (and shared with the TikZ
# taxonomy DAG in the paper via the matching \definecolor block)
VORDER = ["TTP1","PWT","TTP2","MTTP","DTTP","ThOP","TTPTW","CCTTP","TTP-D"]
VCOLOR = {
    "TTP1":"#8DA0CB", "PWT":"#FFD92F", "TTP2":"#66C2A5", "MTTP":"#FC8D62",
    "DTTP":"#A6D854", "ThOP":"#B399D4", "TTPTW":"#E78AC3", "CCTTP":"#B0ECEA",
    "TTP-D":"#E5998F", "Survey":"#B3B3B3",
}
# algorithm families in chronological "era" order (old -> new): the same
# light palette, re-ordered as a cool-to-warm ramp
FAM_ERA = ["Exact methods", "Approximation scheme", "Constructive/iterative heuristic",
           "Single-solution metaheuristic", "Population-based metaheuristic",
           "Hyper-heuristic", "Learning-based", "Quantum / hybrid quantum",
           "Problem def./benchmark/analysis"]
FAM_SHORT = {
    "Exact methods":"Exact", "Approximation scheme":"Approximation",
    "Constructive/iterative heuristic":"Constructive heuristic",
    "Single-solution metaheuristic":"Single-solution metaheuristic",
    "Population-based metaheuristic":"Population-based metaheuristic",
    "Hyper-heuristic":"Hyper-heuristic", "Learning-based":"Learning-based",
    "Quantum / hybrid quantum":"Quantum", "Problem def./benchmark/analysis":"Problem def. / analysis",
}
FAM_COLOR = dict(zip(FAM_ERA, ["#8DA0CB","#C9A66B","#66C2A5","#FFD92F","#A6D854",
                               "#FC8D62","#E78AC3","#B399D4","#B3B3B3"]))
# benchmark categories A/B/C -- fixed everywhere they appear (fig_rdi)
CAT_COLOR = {"A": "#8DA0CB", "B": "#E78AC3", "C": "#66C2A5"}
# publication status -- fixed across fig_growth and fig_venues
STATUS_COLOR = {"Journal": "#E78AC3", "Conference": "#80CDB4", "Preprint": "#8DA0CB"}
# appraisal polarity -- fixed wherever strength/limitation is shown
GOOD_COLOR, BAD_COLOR = "#80CDB4", "#E78AC3"
ACCENT_COLOR = "#F0C239"   # ties/callouts that are neither good nor bad
# sequential / diverging colour maps, built from the same light family
from matplotlib.colors import LinearSegmentedColormap
CMAP_COVERAGE = LinearSegmentedColormap.from_list(
    "coverage", ["#F2F6FA", "#B7C6E3", "#8DA0CB"])       # pale -> soft blue: more coverage
CMAP_SCORE = LinearSegmentedColormap.from_list(
    "score", ["#E5998F", "#FFD98A", "#A6D854"])          # soft coral -> gold -> green: worse -> better practice

def save(fig, name):
    p = os.path.join(OUT, name)
    fig.savefig(p); plt.close(fig); print("  wrote", p)

# ======================================================================
#  Fig. growth -- publications per year by variant + cumulative + status band
# ======================================================================
def _draw_growth(fig, gtop):
    d = rows("Fig8_Year_Variant"); hdr = list(d[0])
    vcols = [c for c in hdr if c in VORDER]
    years = [r[0] for r in d[1:]]
    stacks = {v: np.array([r[hdr.index(v)] or 0 for r in d[1:]]) for v in vcols}

    # cumulative line counts UNIQUE papers (not per-variant), from Master_Corpus
    mc = rows("Master_Corpus"); mh = mc[0]
    iy, ist = mh.index("Year"), mh.index("Pub. status")
    per_year = {y: 0 for y in years}
    status = {s: np.zeros(len(years)) for s in ("Journal", "Conference", "Preprint")}
    iv = mh.index("Variant(s)")
    for r in mc[1:]:
        if str(r[iv]).strip() == "Survey":
            continue
        if r[iy] in per_year:
            per_year[r[iy]] += 1
        if r[ist] in status and r[iy] in years:
            status[r[ist]][years.index(r[iy])] += 1
    cum = np.cumsum([per_year[y] for y in years])

    g = gtop.subgridspec(2, 1, height_ratios=[4.2, 1.0], hspace=0.10)
    ax = fig.add_subplot(g[0]); axs = fig.add_subplot(g[1], sharex=ax)
    ax.set_title("(a)  Publications per year, by variant and by publication type", loc="left", fontsize=10.5)

    bottom = np.zeros(len(years))
    for v in [x for x in VORDER if x in vcols]:
        ax.bar(years, stacks[v], bottom=bottom, width=0.74, label=v,
               color=VCOLOR[v], edgecolor="white", linewidth=0.4)
        bottom += stacks[v]
    ax.set_ylabel("Publications per year")
    #ax.set_title("Growth of the TTP literature, 2013\u20132026", loc="left")
    ax.tick_params(labelbottom=False)
    ax.legend(ncol=5, loc="upper left", handlelength=1.1, columnspacing=1.0,
              borderaxespad=0.2, fontsize=8.3)
    ax.margins(x=0.01)

    ax2 = ax.twinx()
    ax2.plot(years, cum, color=INK, marker="o", ms=4.5, lw=2.0, zorder=5)
    ax2.set_ylabel("Cumulative total", color=INK)
    ax2.tick_params(axis="y", colors=INK); ax2.grid(False)
    ax2.spines["right"].set_visible(True); ax2.spines["right"].set_color(INK)
    ax2.annotate(f"{cum[-1]}", (years[-1], cum[-1]), textcoords="offset points",
                 xytext=(-4, 6), ha="right", color=INK, fontsize=9, fontweight="bold")

    b = np.zeros(len(years))
    for s in ("Journal", "Conference", "Preprint"):
        axs.bar(years, status[s], bottom=b, width=0.74, color=STATUS_COLOR[s],
                edgecolor="white", linewidth=0.4, label=s)
        b += status[s]
    axs.set_ylabel("by type", fontsize=8.5)
    axs.set_xticks(years); axs.set_xticklabels(years, rotation=45, ha="right")
    axs.legend(ncol=3, loc="upper left", fontsize=8, handlelength=1.0)
    axs.margins(x=0.01); axs.set_yticks([])
    axs.set_ylim(0, b.max() * 1.9)
    ax.set_ylim(0, bottom.max() * 1.25)

def fig_bibliometrics():
    fig = plt.figure(figsize=(9.4, 9.6))
    outer = gridspec.GridSpec(2, 1, height_ratios=[5.4, 4.4], hspace=0.30,
                              left=0.08, right=0.90, top=0.97, bottom=0.08)
    _draw_growth(fig, outer[0])
    _draw_venues(fig, outer[1])
    fig.text(0.02, 0.005, "In (a), a paper tagged with two variants appears in both colors, so a bar can exceed the number of papers that year (2013 and 2020). "
             "The line counts each paper once.", fontsize=7.6, color="#4A4A4A", ha="left", va="bottom")
    save(fig, "fig_bibliometrics.png")

# ======================================================================
#  Fig. coverage -- family x variant heatmap + families-per-variant bar
# ======================================================================
def fig_coverage():
    d = rows("Fig5_Coverage_Matrix"); hdr = list(d[0])
    vcols = hdr[1:hdr.index("Row total")]
    i_tot = hdr.index("Row total")
    body = [r for r in d[1:] if r[0] != "Column total"]
    fams = [f for f in FAM_ERA if any(r[0] == f for r in body)]
    M = np.array([[next(r[hdr.index(v)] or 0 for r in body if r[0] == f) for v in vcols]
                  for f in fams], dtype=int)
    # per-family paper count: the sheet's own "Row total" (distinct papers),
    # NOT a re-sum of the cross-tab cells -- a handful of papers are tagged
    # with two variants and would otherwise be double-counted.
    n_papers = [next(r[i_tot] for r in body if r[0] == f) for f in fams]
    # the last row is not an algorithm family (it is every study that defines
    # a benchmark, audits practice, or maps applications without proposing a
    # method), so it is excluded from the "distinct families per variant"
    # count and drawn separately below a divider, in a neutral grey rather
    # than the coverage-blue used for real families.
    algo_mask = [f != "Problem def./benchmark/analysis" for f in fams]
    n_algo = sum(algo_mask)
    fam_per_var = [(M[algo_mask, j] > 0).sum() for j in range(len(vcols))]

    # A handful of papers are tagged with two variants (e.g. a paper that
    # addresses both TTP2 and DTTP); such a paper is counted once in the
    # family's paper total but appears in two cells of its row, so a naive
    # sum of the row's cells can exceed that total by one. Mark every cell
    # that receives such a paper with a dagger so the arithmetic is visible
    # rather than merely asserted.
    mc = rows("Master_Corpus"); mh = list(mc[0])
    iF2, iV2 = mh.index("Family (L1)"), mh.index("Variant(s)")
    dual_cells = {f: set() for f in fams}
    for r in mc[1:]:
        variants = [v.strip() for v in str(r[iV2]).split(",") if v.strip()]
        if "Survey" in variants or len(variants) < 2:
            continue
        fam = r[iF2]
        if fam in dual_cells:
            dual_cells[fam].update(v for v in variants if v in vcols)

    fig = plt.figure(figsize=(10.6, 6.6))
    g = gridspec.GridSpec(2, 2, height_ratios=[1.0, 4.4], width_ratios=[len(vcols), 1.15],
                          hspace=0.06, wspace=0.05)
    axt = fig.add_subplot(g[0, 0]); ax = fig.add_subplot(g[1, 0], sharex=axt)
    axn = fig.add_subplot(g[1, 1], sharey=ax)
    fig.add_subplot(g[0, 1]).axis("off")

    x = np.arange(len(vcols))
    axt.bar(x, fam_per_var, width=0.72, color=[VCOLOR[v] for v in vcols], edgecolor="white")
    for j, v in enumerate(fam_per_var):
        axt.text(j, v + 0.15, str(v), ha="center", fontsize=8.6, fontweight="bold")
    axt.set_ylim(0, 8.4); axt.set_yticks([0, 4, 8])
    axt.set_ylabel("distinct\nfamilies", fontsize=8.5)
    axt.tick_params(labelbottom=False)
    axt.grid(axis="x", visible=False)
    for sp in ("top", "right"): axt.spines[sp].set_visible(False)

    # algorithm-family rows: coverage-blue heatmap
    im = ax.imshow(np.sqrt(M[:n_algo]), cmap=CMAP_COVERAGE, aspect="auto",
                   vmin=0, vmax=np.sqrt(M[:n_algo].max()),
                   extent=(-0.5, len(vcols) - 0.5, n_algo - 0.5, -0.5))
    # non-algorithm row: flat neutral grey, visually set apart
    if n_algo < len(fams):
        ax.imshow(np.zeros((len(fams) - n_algo, len(vcols))), cmap="Greys", vmin=0, vmax=1,
                  aspect="auto", alpha=0.16,
                  extent=(-0.5, len(vcols) - 0.5, len(fams) - 0.5, n_algo - 0.5))
    ax.set_xticks(x); ax.set_xticklabels(vcols, fontweight="bold")
    ax.set_yticks(range(len(fams)))
    ax.set_yticklabels([FAM_SHORT[f]
                        for f in fams])
    any_dual = False
    for i in range(len(fams)):
        for j in range(len(vcols)):
            v = M[i, j]
            label = str(v) if v else "\u00b7"
            if v and vcols[j] in dual_cells[fams[i]]:
                label += "$^{(+)}$"; any_dual = True
            ax.text(j, i, label, ha="center", va="center",
                    fontsize=9.5, fontweight="bold",
                    color="#2A2A2A" if algo_mask[i] else "#6A6A6A")
    ax.grid(False)
    ax.set_xticks(np.arange(-.5, len(vcols)), minor=True)
    ax.set_yticks(np.arange(-.5, len(fams)), minor=True)
    ax.grid(which="minor", color="white", lw=1.8); ax.tick_params(which="minor", length=0)
    if n_algo < len(fams):
        ax.axhline(n_algo - 0.5, color="#4A4A4A", lw=1.6, zorder=5)

    # right-hand "papers per family" panel
    bar_color = [FAM_COLOR[f] if m else "#BFBFBF" for f, m in zip(fams, algo_mask)]
    axn.barh(range(len(fams)), n_papers, color=bar_color, edgecolor="white", height=0.72)
    for i, n in enumerate(n_papers):
        axn.text(n + max(n_papers) * 0.03, i, str(n), va="center", fontsize=9.2, fontweight="bold")
    axn.set_xlim(0, max(n_papers) * 1.18)
    axn.set_title("papers", fontsize=9.5, loc="left")
    axn.tick_params(labelleft=False, labelbottom=False, left=False, bottom=False)
    axn.grid(False)
    for sp in ("top", "right", "bottom"): axn.spines[sp].set_visible(False)
    if n_algo < len(fams):
        axn.axhline(n_algo - 0.5, color="#4A4A4A", lw=1.6, zorder=5)
    note = ("Top: number of distinct algorithm families applied to each variant. "
             "Center: papers per family × variant (grey row: studies with no "
             "algorithm family). Right: papers per family.")
    if any_dual:
        note += ("\n$^{(+)}$ paper tagged with both variants; counted once in "
                  "“papers”, so its row can sum to one more than the total shown.")
    fig.text(0.01, 0.005, note, fontsize=7.6, color="#4A4A4A", ha="left")
    save(fig, "fig_coverage.png")

# ======================================================================
#  Fig. paradigm -- family share per biennium, families ordered by era
# ======================================================================
def fig_paradigm():
    """Unit chart: one square per study, stacked by algorithm family within each two-year period."""
    from matplotlib.patches import Rectangle
    mc = rows("Master_Corpus"); mh = list(mc[0])
    iy, ifam, iv = mh.index("Year"), mh.index("Family (L1)"), mh.index("Variant(s)")
    starts = list(range(2013, 2027, 2))
    labels = [f"{a}\u2013{str(a + 1)[2:]}" for a in starts]
    order = list(FAM_ERA)                      # exact ... quantum, then problem definition / analysis
    cnt = {b: {f: 0 for f in order} for b in range(len(starts))}
    tot_f = {f: 0 for f in order}
    for r in mc[1:]:
        if not r[0] or str(r[iv]).strip() == "Survey" or r[iy] is None: continue
        f = r[ifam]
        if f not in cnt[0]: continue
        b = (int(r[iy]) - 2013) // 2
        cnt[b][f] += 1; tot_f[f] += 1
    ncol = 3; cell = 1.0; gap = 0.09
    fig, ax = plt.subplots(figsize=(10.4, 5.2))
    for b in range(len(starts)):
        x0 = b * (ncol + 1.15)
        i = 0
        for f in order:
            for _ in range(cnt[b][f]):
                r_, c_ = divmod(i, ncol)
                ax.add_patch(Rectangle((x0 + c_ * cell + gap / 2, r_ * cell + gap / 2), cell - gap, cell - gap,
                                       color=FAM_COLOR[f], lw=0))
                i += 1
        n = i
        ax.text(x0 + ncol / 2, -0.85, labels[b], ha="center", va="top", fontsize=9.5, fontweight="bold")
        ax.text(x0 + ncol / 2, -2.05, f"n = {n}", ha="center", va="top", fontsize=8.4, color="#4A4A4A")
    ax.set_xlim(-0.4, len(starts) * (ncol + 1.15) - 0.5 + 0.2)
    ax.set_ylim(-3.0, 7.3)
    ax.axis("off")
    ax.set_aspect("equal")
    handles = [Rectangle((0, 0), 1, 1, color=FAM_COLOR[f]) for f in order]
    names = [f"{FAM_SHORT[f]} ({tot_f[f]})" for f in order]
    ax.legend(handles, names, loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=3, frameon=False,
              title="Family (studies in total), from the bottom to the top of each stack",
              title_fontsize=9, fontsize=8.8, handlelength=1.0, columnspacing=1.6)
    save(fig, "fig_paradigm.png")

# ======================================================================
#  Fig. venues -- top venues, full names, tagged by type
# ======================================================================
def _venue_clean(s):
    s = re.sub(r"\\textit\{([^}]*)\}", r"\1", s); s = re.sub(r"\\textbf\{([^}]*)\}", r"\1", s)
    s = s.replace("\\&", "&").replace("\\ ", " ").replace("\\", "")
    s = s.replace("{", "").replace("}", "")
    s = re.sub(r"^Proceedings of the\s*", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s
def _wrap(s, w=34):
    out, line = [], ""
    for word in s.split():
        if len(line) + len(word) + 1 > w:
            out.append(line); line = word
        else:
            line = (line + " " + word).strip()
    out.append(line)
    return "\n".join(out[:2]) + ("\u2026" if len(out) > 2 else "")

def _draw_venues(fig, gbot):
    d = rows("Fig9_Venues")
    TYPES = {"Conference", "Journal", "Preprint"}
    vs = [(_venue_clean(r[0]), r[1], r[2]) for r in d[1:]
          if isinstance(r[1], int) and r[0] and _venue_clean(r[0]) not in TYPES
          and not str(r[0]).startswith("arXiv")]
    vs.sort(key=lambda t: -t[1])
    top = vs[:11][::-1]
    def _nm(n):
        n = n.replace("Optimisation", "Optimization")
        if "Combinatorial Optimization" in n:
            return "Evolutionary Computation in Combinatorial Optimization (EvoCOP)"
        return n
    names = [_nm(t[0]) for t in top]
    vals = [t[1] for t in top]
    cols = [STATUS_COLOR.get(t[2], MUTED) for t in top]

    ax = fig.add_subplot(gbot)
    ax.set_title("(b)  Top publication venues", loc="left", fontsize=10.5)
    y = np.arange(len(top))
    ax.barh(y, vals, color=cols, edgecolor="white")
    ax.set_yticks([])
    ax.set_xlabel("Number of papers")
    # ax.set_title("Where TTP research is published (top venues)", loc="left")
    for i, (v, nm) in enumerate(zip(vals, names)):
        ax.text(v + 0.25, i, f"{v}", va="center", fontsize=9, fontweight="bold")
        ax.text(v + 0.95, i, nm, va="center", fontsize=8.8, color="#333")
    ax.set_xlim(0, max(vals) * 2.35)
    ax.set_xticks(range(0, max(vals) + 2, 2))
    ax.grid(axis="y", visible=False)
    handles = [plt.Rectangle((0, 0), 1, 1, color=STATUS_COLOR["Journal"]),
               plt.Rectangle((0, 0), 1, 1, color=STATUS_COLOR["Conference"])]
    ax.legend(handles, ["Journal", "Conference"], loc="lower right")

# ======================================================================
#  Fig. rdi -- 2 panels: RDI by category + Friedman/Nemenyi CD diagram
# ======================================================================
def _tab10():
    d = rows("Tab10_TTP1_RDI"); hdr = list(d[0])
    idx = {k: hdr.index(k) for k in ("Cat A avg RDI", "Cat B avg RDI", "Cat C avg RDI",
                                     "Overall avg RDI", "Mean Friedman rank", "# best-known (of 60)")}
    recs = [r for r in d[1:] if isinstance(r[idx["Cat A avg RDI"]], (int, float))]
    return recs, idx, hdr

def fig_rdi():
    recs, idx, hdr = _tab10()
    recs = sorted(recs, key=lambda r: r[idx["Overall avg RDI"]])
    names = [r[0] for r in recs]
    A = [r[idx["Cat A avg RDI"]] for r in recs]
    B = [r[idx["Cat B avg RDI"]] for r in recs]
    C = [r[idx["Cat C avg RDI"]] for r in recs]
    nbest = [r[idx["# best-known (of 60)"]] for r in recs]

    fig = plt.figure(figsize=(8.6, 5.8))
    g = gridspec.GridSpec(2, 1, height_ratios=[1.15, 1.0], hspace=0.42)
    ax = fig.add_subplot(g[0]); axc = fig.add_subplot(g[1])

    x = np.arange(len(names)); w = 0.26
    ax.bar(x - w, A, w, label="Category A", color=CAT_COLOR["A"])
    ax.bar(x,     B, w, label="Category B", color=CAT_COLOR["B"])
    ax.bar(x + w, C, w, label="Category C", color=CAT_COLOR["C"])
    for xi, vals in zip(x, zip(A, B, C)):
        for off, val in zip((-w, 0, w), vals):
            ax.text(xi + off, val + 1.6, f"{val:.0f}", ha="center", fontsize=7.2)
    ax.set_xticks(x); ax.set_xticklabels(names, rotation=15, ha="right", fontweight="bold")
    ax.set_ylabel("Mean RDI (%)  \u2013  lower is better")
    ax.set_ylim(0, max(A + B + C) * 1.14)
    ax.set_title("(a)  Relative deviation index by benchmark category, 60-instance pool",
                 loc="left")
    ax.legend(ncol=3, loc="upper left")
    ax.grid(axis="x", visible=False)

    # (b) Friedman + Nemenyi CD diagram -- labels split left / right
    N = max(1, len(rows("Lit_CatABC_RDI")) - 1); k = len(recs)
    q05 = {5:2.728, 6:2.850, 7:2.949, 8:3.031, 9:3.102, 10:3.164}
    CD = q05.get(k, 3.1) * math.sqrt(k * (k + 1) / (6 * N))
    pairs = sorted([(r[0], r[idx["Mean Friedman rank"]]) for r in recs], key=lambda t: t[1])
    ranks = np.array([p[1] for p in pairs])
    lo, hi = 1, math.ceil(max(ranks)) + 1
    axc.set_xlim(lo - 2.7, hi + 2.7); axc.set_ylim(-2.6, 1.9); axc.axis("off")
    axc.text(lo - 2.7, 1.75, f"(b)  Friedman + Nemenyi test (k = {k}, N = {N})",
             ha="left", fontsize=10, fontweight="bold")
    axc.plot([lo, hi], [0, 0], color="#333", lw=1.3)
    for t in range(lo, hi + 1):
        axc.plot([t, t], [0, 0.09], color="#333", lw=1.0)
        axc.text(t, 0.24, str(t), ha="center", fontsize=8.3)
    axc.text((lo + hi) / 2, 0.62, "mean Friedman rank  (1 = best)", ha="center", fontsize=8.6)

    half = (k + 1) // 2
    for i, (nm, rk) in enumerate(pairs):
        left = i < half
        lev = -1.05 - (i if left else k - 1 - i) * 0.32
        xend = lo - 0.2 if left else hi + 0.2
        axc.plot([rk, rk], [0, lev], color="#999", lw=0.9)
        axc.plot([rk, xend], [lev, lev], color="#999", lw=0.9)
        axc.text(xend + (-0.12 if left else 0.12), lev, f"{nm} ({rk:.2f})",
                 ha="right" if left else "left", va="center", fontsize=8.4, fontweight="bold")
    bn = 0; prev_j = -1
    for i in range(k):
        j = i
        while j + 1 < k and ranks[j + 1] - ranks[i] < CD:
            j += 1
        if j > i and j > prev_j:
            yy = -0.30 - bn * 0.15
            axc.plot([ranks[i] - 0.03, ranks[j] + 0.03], [yy, yy], color=ACCENT_COLOR,
                     lw=3, solid_capstyle="round")
            prev_j = j; bn += 1
    save(fig, "fig_rdi.png")

# ======================================================================
#  Fig. sw -- recurring strengths vs limitations (diverging)
# ======================================================================
def fig_sw():
    d = rows("SW_Themes")
    recs = [(r[0], r[1] or 0, r[2] or 0) for r in d[1:]
            if r[0] and not str(r[0]).startswith("Other")]
    recs.sort(key=lambda t: t[2])
    themes = [r[0] for r in recs]
    s = np.array([r[1] for r in recs]); w = np.array([r[2] for r in recs])
    y = np.arange(len(themes))
    fig, ax = plt.subplots(figsize=(9.2, 5.0))
    ax.barh(y, -w, color=BAD_COLOR, label="recurring limitation")
    ax.barh(y,  s, color=GOOD_COLOR, label="recurring strength")
    for yi, (sv, wv) in enumerate(zip(s, w)):
        if wv: ax.text(-wv - 1.4, yi, str(wv), va="center", ha="right", fontsize=8.6, fontweight="bold")
        if sv: ax.text( sv + 1.4, yi, str(sv), va="center", ha="left",  fontsize=8.6, fontweight="bold")
    ax.set_yticks(y); ax.set_yticklabels(themes)
    ax.axvline(0, color="#333", lw=0.9)
    m = max(w.max(), s.max()) * 1.18
    ax.set_xlim(-m, m)
    ax.set_xticks([t for t in ax.get_xticks()])
    ax.set_xticklabels([str(abs(int(t))) for t in ax.get_xticks()])
    ax.set_xlabel("appraisal notes            limitation  \u2190  |  \u2192  strength")
    # ax.set_title("Recurring strengths and limitations across the reviewed corpus", loc="left")
    ax.legend(loc="lower right"); ax.grid(axis="y", visible=False)
    save(fig, "fig_sw.png")

# ======================================================================
#  Fig. reporting -- merged research-practice + experimental-reporting
#  scorecard by variant, annotated with the two weakest columns' agenda items
# ======================================================================
def fig_reporting():
    d = rows("Tab16_ReportingAudit"); hdr = list(d[0])
    iv = hdr.index("Variant")
    # "Formal model" is met by all 93 studies (it follows from the eligibility
    # criteria of Section 2.2, not from the audit), so it carries no
    # information here and is left out of the scorecard
    dims = ["Uses variant's standard benchmark", "Reports dispersion (std/RSD/CI)",
            "# runs stated", "Formal statistical test", "Code available", "Compared to current SOTA"]
    dim_short = ["Standard\nbenchmark", "Dispersion\nreported",
                 "Runs\nstated", "Statistical\ntest", "Code\navailable", "Compared\nto SOTA"]
    order = [v for v in VORDER]
    num = {v: np.zeros(len(dims)) for v in order}; den = {v: 0 for v in order}
    tn = np.zeros(len(dims)); td = 0
    for r in d[1:]:
        if not r[0] or not r[iv]: continue          # skip the summary block below the data rows
        vs = [x.strip() for x in str(r[iv]).split(",") if x.strip() in order]
        if not vs: continue
        yn = [1 if r[hdr.index(c)] == "Y" else 0 for c in dims]
        for v in vs:
            num[v] += yn; den[v] += 1
        tn += yn; td += 1
    rl = [v for v in order if den[v]] + ["All"]
    M = np.array([num[v] / den[v] for v in order if den[v]] + [tn / td])
    cnt = [[(int(num[v][j]), den[v]) for j in range(len(dims))] for v in order if den[v]]
    cnt.append([(int(tn[j]), td) for j in range(len(dims))])
    all_idx = len(rl) - 1

    # extra block: the same criteria by publication period (each paper counted once)
    mc = rows("Master_Corpus"); mh = list(mc[0])
    year = {r[mh.index("CiteKey")]: r[mh.index("Year")] for r in mc[1:] if r[0]}
    periods = [("2013\u20132017", 0, 2017), ("2018\u20132021", 2018, 2021), ("2022\u20132026", 2022, 9999)]
    prow = []; pcnt = []
    for lab, lo, hi in periods:
        sel = [r for r in d[1:] if r[0] and r[iv] and lo <= year[r[0]] <= hi]
        k = np.array([sum(1 for r in sel if r[hdr.index(c)] == "Y") for c in dims])
        prow.append(k / len(sel)); pcnt.append([(int(x), len(sel)) for x in k])
    rl = rl + [p[0] for p in periods]
    M = np.vstack([M, np.array(prow)])
    cnt = cnt + pcnt

    from matplotlib.patches import Rectangle
    nA = all_idx + 1; nP = len(periods)
    fig = plt.figure(figsize=(10.4, 8.6))
    g = gridspec.GridSpec(2, 2, height_ratios=[nA, nP + 0.5], width_ratios=[1, 0.035],
                          hspace=0.26, wspace=0.04, left=0.09, right=0.93, top=0.90, bottom=0.14)
    axa = fig.add_subplot(g[0, 0]); axb = fig.add_subplot(g[1, 0])
    caxa = fig.add_subplot(g[0, 1]); caxb = fig.add_subplot(g[1, 1])
    cmap_period = LinearSegmentedColormap.from_list("period", ["#F1F8F5", "#B9E0D2", "#66C2A5"])
    ima = axa.imshow(M[:nA], cmap=CMAP_COVERAGE, vmin=0, vmax=1, aspect="auto")
    imb = axb.imshow(M[nA:], cmap=cmap_period, vmin=0, vmax=1, aspect="auto")
    for ax_, first, n_, rlab, big in ((axa, 0, nA, rl[:nA], True), (axb, nA, nP, rl[nA:], False)):
        ax_.set_yticks(range(n_)); ax_.set_yticklabels(rlab, fontweight="bold")
        ax_.set_xticks(range(len(dims)))
        for i in range(n_):
            for j in range(len(dims)):
                k, n = cnt[first + i][j]
                lab = f"{k}/{n}" if big else f"{k}/{n}\n{100 * k / n:.0f}%"
                ax_.text(j, i, lab, ha="center", va="center", fontsize=8.8 if big else 8.4, color="#1A1A1A")
        ax_.grid(False)
        ax_.set_xticks(np.arange(-.5, len(dims)), minor=True)
        ax_.set_yticks(np.arange(-.5, n_), minor=True)
        ax_.grid(which="minor", color="white", lw=1.8); ax_.tick_params(which="minor", length=0)
    axa.axhline(all_idx - 0.5, color="#333", lw=1.3)
    axa.xaxis.tick_top(); axa.set_xticklabels(dim_short, fontweight="bold")
    axb.set_xticklabels(dim_short, fontweight="bold")
    axa.set_title("(a)  By variant, with the total over all 93 studies", loc="left", fontsize=10.5, pad=34)
    axb.set_title("(b)  By publication period", loc="left", fontsize=10.5, pad=8)
    weakest = list(np.argsort(M[all_idx])[:3])
    for j in weakest:
        axa.add_patch(Rectangle((j - 0.5, all_idx - 0.5), 1, 1, fill=False, ec="#E8862A", lw=2.8, zorder=5))
    cba = fig.colorbar(ima, cax=caxa); cbb = fig.colorbar(imb, cax=caxb)
    cba.set_label("share of studies", fontsize=10); cbb.set_label("share of studies", fontsize=10)
    fig.text(0.09, 0.008,
             "- Orange frames mark the three criteria met by the fewest studies overall; the minimum reporting standard "
             "proposed in Section 6.3 targets them.\n"
             "- Two papers are tagged with two variants and counted in both rows of (a) (one TTP1 and TTP2 paper, one TTP2 and DTTP paper), so a column can add\n up to more than the total.\n- The All row and the rows of (b) count each paper once.",
             fontsize=8, color="#4A4A4A", ha="left", va="bottom")
    save(fig, "fig_reporting.png")

# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("Generating figures ->", os.path.abspath(OUT))
    # remove figures that are cut or moved to TikZ
    for stale in ["fig02_prisma_flow", "fig03_taxonomy_tree",
                  "fig04_algorithm_classification_tree", "fig05_coverage_heatmap",
                  "fig06_ttp1_rdi_comparison", "fig06b_cd_diagram", "fig06c_bestknown_counts",
                  "fig08_publications_per_year", "fig09_venue_distribution",
                  "fig10_family_paradigm_trend", "fig11_taxonomy_gap_radar",
                  "fig12_strengths_weaknesses", "fig13_practice_scorecard",
                  "fig_corpus_distributions", "fig_status_per_year", "fig_scorecard",
                  "fig_designspace"]:
        p = os.path.join(OUT, stale + ".png")
        if os.path.exists(p):
            os.remove(p); print("  removed", p)
    fig_bibliometrics()
    fig_coverage()
    fig_paradigm()
    fig_rdi()
    fig_sw()
    fig_reporting()
    print("Done.")
