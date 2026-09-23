#!/usr/bin/env python3
r"""Regenerates the 15 LaTeX tables of the paper into tables/ (byte-identical to the submitted
files).

Two tables are computed from data/ttp_crt_data.xlsx:
  tab_rdi.tex        the nine-method RDI comparison, from sheet CatABC_Objectives, merged with
                     each method's time budget, run count, and hardware/language (Section 5.4,
                     RDI_BUDGET below; a former stand-alone tab_budget.tex was folded in here,
                     since both tables shared the same nine Algorithm/Ref. rows)
  tab_reporting.tex  the per-paper audit rows and subtotals, from sheets Reporting_Audit and Corpus
The other 13 tables are hand-written text (captions, catalogs, links). Their source text is
stored in STATIC below, so that they are written out exactly as submitted.
tab_coupling.tex classifies each method-proposing study by its tour-packing coupling mechanism
(Section 5.2 of the paper); the classification is a one-off coding pass documented in the paper
text, recorded in the 'Coupling strategy' column of the Corpus sheet.
R1 (Section 6.3) states its minimum reporting items in prose, as hints rather than a formal
checklist for journals to adopt.

The paper has five appendices, all landscape material moved out of the main body:
  A  code links (tab_code_links.tex)
  B  consolidated taxonomy (tab_taxonomy.tex)
  C  unexplored combinations and variant summary (tab_unexplored.tex + tab_variant_glance.tex)
  D  variant catalogs (tab_cat_ttp1.tex + tab_cat_small.tex)
  E  reporting audit (tab_reporting.tex)
Section 3's and Section 4's prose points to all of these by \ref rather than inputting them
inline. None of the appendix `\section{...}\label{...}` headers or their intro paragraphs live
in these table files (an earlier version tried that, then tried a variant with the header in
the first table file of each appendix; both were abandoned). They live directly in the paper's
own \appendix block in ttp_slr_tax_paper.tex, one after another, with only Appendix A left in
portrait and a single \begin{landscape}...\end{landscape} wrapping B through E as one
continuous block (containing all five \section commands, \input calls, and one \newpage between
tab_cat_ttp1.tex and tab_cat_small.tex). These table files, as stored here, are therefore table
content only, with no section header of their own; they are not meant to be \input in isolation
without the paper's surrounding \appendix text.
tab_taxonomy.tex, tab_unexplored.tex, tab_variant_glance.tex, and tab_cat_small.tex use
\begin{center}...\captionof{table}{...}...\end{center} rather than a floating \begin{table},
because floating tables on these shared-landscape, multi-section appendix pages were observed
to reorder themselves ahead of the \section heading and intro paragraph that precede them in
the source (a pdflscape/float interaction, caught by reading the rendered PDF's text in order,
not just by checking that it compiled); \captionof avoids LaTeX's float placement algorithm
entirely, so content renders in strict source order. tab_cat_ttp1.tex is its own landscape
longtable (wrapped internally in \begin{landscape}) for TTP1's 57 studies, using longtable's own
\caption mechanism instead, which does not have this reordering problem. tab_cat_small.tex holds
the PWT, MTTP/DTTP, and ThOP catalogs together as three \multicolumn-headed groups (16 rows
total) as one non-floating table, so they land on one rotated page instead of three separate
ones. Both catalog tables share the same 5-column layout, Reference/Year/Algorithm/Instances/
Key result (p{2.8cm}cp{3.0cm}p{4.9cm}p{7.9cm}), with no Code column: code release is already
recorded per paper in tab_reporting.tex (Appendix E), so it was dropped here and the freed
width went to the other four columns.

Usage:  python3 scripts/gen_tables.py            (writes to tables/)
        python3 scripts/gen_tables.py --check    (compares with the files in tables/, writes nothing)
"""
import sys
from pathlib import Path
import numpy as np
import openpyxl
import scipy.stats as ss

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "ttp_crt_data.xlsx"
OUT = ROOT / "tables"
D = "\\ding{108}"

# ----------------------------------------------------------------------
#  tab_rdi.tex : computed from the raw objectives (same RDI rule as gen_figures.py)
# ----------------------------------------------------------------------
RDI_REF = {'CoCo+': '\\citet{namazi2023solving}', 'CoCo': '\\citet{namazi2019cooperative}', 'CS2SA': '\\citet{el2016population}', 'MA2B': '\\citet{el2016population}', 'S5': '\\citet{faulkner2015approximate}\\tnote{a}', 'RWS': '\\citet{zhang2021solving}', 'SAVI': '\\citet{nguyen2023simulated}', 'CS2SA*': '\\citet{el2018efficiently}', 'MATLS': '\\citet{mei2014improving}\\tnote{a}'}
# Time budget, run count, and hardware/language for each of the nine RDI-pool
# methods, as stated in the source paper for the values used below (Section 5.4).
# Merged into this table from a former stand-alone tab_budget.tex.
RDI_BUDGET = {
    'CoCo+':  ('600s\\tnote{b} / 10', 'Not stated'),
    'CoCo':   ('600s / 10', 'Intel Xeon X5650 2.66GHz'),
    'CS2SA':  ('600s / 10', 'Core i3-2370M 2.40GHz/4GB, Java'),
    'MA2B':   ('600s / 10', 'Core i7-5700HQ 2.70GHz/8GB, Java'),
    'S5':     ('600s / 10', 'Not stated'),
    'RWS':    ('600s / 10', 'Not stated'),
    'SAVI':   ('600s / 10', 'AMD Ryzen 7 5800U 1.9GHz, Java'),
    'CS2SA*': ('600s / 10', 'Java'),
    'MATLS':  ('600s / 10', 'Not stated'),
}
RDI_HEAD = r'''\begin{table}[htbp]
\centering
\scriptsize
\setlength{\tabcolsep}{3pt}
\begin{threeparttable}
\caption{Nine TTP1 algorithms on the categorized 60-instance benchmark: mean RDI (\%, lower is better) by category and overall, mean Friedman rank with its leave-one-out range, best-in-pool count, and the reported time budget, run count, and hardware or language.}
\label{tab:rdi}
\begin{tabular}{@{}llccccccclp{2.9cm}@{}}
\toprule
\textbf{Algorithm} & \textbf{Ref.} & \textbf{Cat.\,A} & \textbf{Cat.\,B} & \textbf{Cat.\,C} & \textbf{Overall} & \textbf{Rank} & \textbf{LOO range} & \textbf{\#best} & \textbf{Budget} & \textbf{Hardware / language} \\
\midrule
'''
RDI_TAIL = r'''\bottomrule
\end{tabular}
\begin{tablenotes}
\footnotesize
\item[a] Re-evaluated by \citet{zhang2021solving} under its own protocol (Section~\ref{sec:performance}), not the original authors' protocol.
\item[b] Plus an additional 1-hour robustness check beyond the 600s budget used for the reported values.
\end{tablenotes}
\end{threeparttable}
\end{table}
'''

def rdi_table(wb):
    rs = list(wb["CatABC_Objectives"].iter_rows(values_only=True))
    algos = [str(x).replace("\u2217", "*") for x in rs[1][9:18]]
    recs = [r for r in rs[2:] if r[1] in ("A", "B", "C") and isinstance(r[9], (int, float))]
    V = np.array([[r[9 + i] for i in range(len(algos))] for r in recs], dtype=float)
    rng = V.max(1) - V.min(1)
    R = (V.max(1, keepdims=True) - V) / np.where(rng == 0, 1, rng)[:, None] * 100
    ranks = np.array([ss.rankdata(-v, method="average") for v in V])
    cats = np.array([r[1] for r in recs])
    best = [0] * len(algos)
    for v, g in zip(V, rng):
        if g > 0: best[int(np.argmax(v))] += 1
    order = sorted(range(len(algos)), key=lambda i: R[:, i].mean())
    # leave-one-out sensitivity (Section 5.3): drop each method in turn (k=8)
    # and record the range of every other method's mean rank across the runs
    loo = {a: [] for a in algos}
    for drop in range(len(algos)):
        keep = [i for i in range(len(algos)) if i != drop]
        rsub = np.array([ss.rankdata(-v, method="average") for v in V[:, keep]])
        mr = rsub.mean(axis=0)
        for j, i in enumerate(keep):
            loo[algos[i]].append(mr[j])
    lines = []
    for i in order:
        a, m = algos[i], [R[cats == c, i].mean() for c in "ABC"]
        lo_r, hi_r = min(loo[a]), max(loo[a])
        budget, hw = RDI_BUDGET[a]
        lines.append("%s & %s & %.1f & %.1f & %.1f & \\textbf{%.1f} & %.2f & %.2f--%.2f & %d & %s & %s \\\\"
                     % (a, RDI_REF[a], m[0], m[1], m[2], R[:, i].mean(), ranks[:, i].mean(),
                        lo_r, hi_r, best[i], budget, hw))
    return RDI_HEAD + "\n".join(lines) + "\n" + RDI_TAIL

# ----------------------------------------------------------------------
#  tab_reporting.tex : rows and subtotals computed from the audit sheet
# ----------------------------------------------------------------------
REP_HEAD = r'''{\footnotesize
\begin{longtable}{@{}p{3.0cm}p{3.4cm}p{1.0cm}p{2.9cm}p{2.9cm}p{1.6cm}p{1.4cm}p{1.5cm}@{}}
\caption{Full research-practice and experimental-reporting audit, one row per paper.}\\\label{tab:reporting}\\
\toprule
\textbf{Reference} & \textbf{Evidence base} & \textbf{Runs} & \textbf{Reported statistic} & \textbf{Significance test} & \textbf{Std.\ bench.} & \textbf{Code} & \textbf{SOTA compar.} \\
\midrule
\endfirsthead
\multicolumn{8}{l}{\footnotesize\itshape Table~\ref{tab:reporting} continued}\\ \toprule
\textbf{Reference} & \textbf{Evidence base} & \textbf{Runs} & \textbf{Reported statistic} & \textbf{Significance test} & \textbf{Std.\ bench.} & \textbf{Code} & \textbf{SOTA compar.} \\
\midrule
\endhead
\midrule \multicolumn{8}{r}{\footnotesize continued on next page}\\
\endfoot
\bottomrule
\endlastfoot
'''
REP_TAIL = r'''\end{longtable}
}
'''
EV={"Theoretical / no computational experiment":"Theoretical/none",
 "Self-generated or synthetic instances":"Self-gen./synthetic",
 "One-off subset of the official library (used by only that paper)":"One-off subset",
 "Shared/reused benchmark within its sub-literature":"Shared subset",
 "TTP1 categorised 60-instance benchmark":"TTP1 categorized 60",
 "Full 9,720-instance library":"Full 9,720-library"}
def reporting_body(wb):
    rs=list(wb["Reporting_Audit"].iter_rows(values_only=True)); h=list(rs[0]); rs=rs[1:]
    yr={r[0]:r[1] for r in wb["Corpus"].iter_rows(min_row=2,values_only=True)}
    g=lambda r,c:r[h.index(c)]
    out=[]
    for v in ["TTP1","PWT","TTP2","MTTP","DTTP","ThOP","TTPTW","CCTTP","TTP-D"]:
        sel=sorted([r for r in rs if v in [x.strip() for x in g(r,"Variant").split(",")]],key=lambda r:(yr[r[0]],r[0]))
        n=len(sel)
        out.append("\\midrule \\multicolumn{8}{l}{\\textbf{%s} (%d papers)} \\\\ \\midrule"%(v,n))
        for r in sel:
            if g(r,"Reports dispersion (std/RSD/CI)")=="Y":
                d=g(r,"Dispersion label")
                st="Mean $+$ 95\\% CI" if d=="95% CI" else "Mean $\\pm$ "+d
            elif g(r,"Central tendency label"):
                st="Mean only" if g(r,"Central tendency label")=="Mean" else g(r,"Central tendency label")+" only"
            else: st=""
            runs=str(g(r,"# runs (value)")) if g(r,"# runs (value)") is not None else (D if g(r,"# runs stated")=="Y" else "")
            tst=(g(r,"Statistical test label") or "") if g(r,"Formal statistical test")=="Y" else ""
            f=lambda c: D if g(r,c)=="Y" else ""
            out.append("\\citet{%s} & %s & %s & %s & %s & %s & %s & %s \\\\"%(r[0],EV[g(r,"Evidence class")],runs,st,tst,f("Uses variant's standard benchmark"),f("Code available"),f("Compared to current SOTA")))
        c=lambda col:sum(1 for r in sel if g(r,col)=="Y")
        rep=sum(1 for r in sel if g(r,"Reports mean/avg")=="Y" or g(r,"Reports dispersion (std/RSD/CI)")=="Y")
        out.append("\\textit{%s subtotal} & \\textit{runs stated %d/%d; stat.\\ test %d/%d} & & \\textit{reported stat.\\ %d/%d} & & \\textbf{%d/%d} & \\textbf{%d/%d} & \\textbf{%d/%d} \\\\"%(v,c("# runs stated"),n,c("Formal statistical test"),n,rep,n,c("Uses variant's standard benchmark"),n,c("Code available"),n,c("Compared to current SOTA"),n))
        out.append("\\addlinespace[3pt]")
    N=len(rs); c=lambda col:sum(1 for r in rs if g(r,col)=="Y")
    out.append("\\midrule")
    out.append("\\textbf{All primary studies (%d)} & \\textit{runs stated %d; stat.\\ test %d} & & \\textit{mean/avg %d; dispersion %d} & & \\textbf{%d} & \\textbf{%d} & \\textbf{%d} \\\\"%(N,c("# runs stated"),c("Formal statistical test"),c("Reports mean/avg"),c("Reports dispersion (std/RSD/CI)"),c("Uses variant's standard benchmark"),c("Code available"),c("Compared to current SOTA")))
    return "\n".join(out)+"\n"

def reporting_table(wb):
    return REP_HEAD + reporting_body(wb) + REP_TAIL

# ----------------------------------------------------------------------
#  hand-written tables
# ----------------------------------------------------------------------
STATIC = {
    'tab_applications.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{What each variant adds to the original model and the motivation its authors state, with the number of studies.}
\label{tab:applications}
\setlength{\tabcolsep}{4pt}
\renewcommand{\arraystretch}{1.15}
\begin{tabular}{@{}p{1.5cm}p{5.0cm}p{8.0cm}r@{}}
\toprule
\textbf{Variant} & \textbf{Feature added} & \textbf{Stated motivation} & \textbf{Studies} \\
\midrule
TTP1, PWT & Load-dependent speed; PWT fixes the route & Benchmark for interacting components \citep{bonyadi2013travelling}; pick-up-and-delivery, waste and recyclables collection \citep{sarkar2024travelling}; last-mile cargo bikes \citep{naumov2021identifying} & 57, 4 \\
TTP2 & Second objective, profit against time \citep{blank2017solving} & No application recorded & 18 \\
MTTP & Several thieves sharing one capacity & Realism limit of a single thief \citep{chand2016fast} & 1 \\
DTTP & Items or cities change during operation & Disruptions during operation \citep{sachdeva2020dynamic,herring2020dynamic,herring2020responsive} & 3 \\
ThOP & Subset of cities within a time limit \citep{santos2018thief} & Pick-up-and-delivery logistics, per the survey \citep{sarkar2024travelling} & 8 \\
TTPTW & Per-city time windows with waiting & No application recorded \citep{angmalisang2026traveling} & 1 \\
CCTTP & Uncertain item weights & Measurement imprecision, environment, perishability \citep{pathirage2024chance,don2025weighted} & 2 \\
TTP-D & Onboard drone retrieves outlying items & Application scenarios claimed \citep{murjani2026drive} & 1 \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_benchmarks.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{Benchmark instances available for each variant (No: no public release reported).}
\label{tab:benchmarks}
\setlength{\tabcolsep}{4pt}
\begin{threeparttable}
\begin{tabular}{@{}lp{10.0cm}cp{3.5cm}@{}}
\toprule
\textbf{Variant} & \textbf{Benchmark instances} & \textbf{Online?} & \textbf{Reference} \\
\midrule
TTP1 & Standard suite: 9,720 instances from 81 TSPLIB graphs (51--85,900 cities); the categorized 60-instance subset (Category A/B/C) is the common evaluation ground. Two smaller sets support theory work: 432 instances with known optima, and 720 instances evolved to discriminate between algorithms. & Partial\tnote{a} & \citet{polyakovskiy2014comprehensive}; \citet{wu2017exact}; \citet{bossek2021generating} \\
PWT & Reuses the TTP1 standard suite, since PWT is a fixed-route restriction of TTP1; no separate instances exist. & Yes\tnote{c} & \citet{polyakovskiy2014comprehensive} \\
TTP2 & The BI-TTP competition suite: 9 instances selected from the TTP1 standard suite (EMO-2019 / GECCO-2019, cities 280--33{,}810); a 12-instance eil51-based set is separately reused by three papers, and some studies instead adapt the full TTP1 standard suite. & Yes\tnote{c} & \citet{polyakovskiy2014comprehensive}; \citet{chagas2021non} \\
MTTP & No shared suite; the sole paper reuses a 72-instance subset of the standard suite originally selected for approximation studies. & Yes\tnote{c} & \citet{chand2016fast} \\
DTTP & No shared suite; each of the three papers adds disruption scenarios to a reused subset of the standard or competition suites (9, 12, and 18 base instances respectively). & Partial\tnote{b} & \citet{sachdeva2020dynamic}; \citet{herring2020dynamic}; \citet{herring2020responsive} \\
ThOP & A dedicated 432-instance suite (cities \textasciitilde{}50--280), used by every ThOP paper to date, the only frontier variant with this level of adoption. & Partial\tnote{d} & \citet{santos2018thief} \\
TTPTW & A single derived set at three time-window tightness levels (cities 51--1000); not yet reused by a second paper. & No & \citet{angmalisang2026traveling} \\
CCTTP & A single derived set of uniform-weight scenarios at confidence levels up to 0.999 (cities 51--1000); later chance-constrained studies build on the same setting but do not yet share one standardized suite. & No & \citet{pathirage2024chance} \\
TTP-D & Two derived instance families from a single paper, one MILP-scale and one heuristic-scale; not yet reused or standardized. & No & \citet{murjani2026drive} \\
\bottomrule
\end{tabular}
\begin{tablenotes}\footnotesize
\item[a] Standard suite and evolved set are public at \url{http://cs.adelaide.edu.au/~optlog/research/ttp.php}; the availability of the exact-reference set is not stated.
\item[b] Only the disruption instances and code of \citet{sachdeva2020dynamic} are public at \url{https://github.com/ragavsachdeva/DynTTP}
\item[c] \url{http://cs.adelaide.edu.au/~optlog/research/ttp.php}
\item[d] \citet{santos2018thief} state the instances are hosted at \url{http://www.dpi.ufv.br/~andre/thop/}; this address no longer resolves (checked September 2026).
\end{tablenotes}
\end{threeparttable}
\end{table}
''',
    'tab_code_links.tex': r'''
\begin{center}
\footnotesize
\captionof{table}{Code and resource links given by the original papers, with the result of a reachability check (a single HTTP request to each address) on 21 September 2026.}
\label{tab:codelinks}
\setlength{\tabcolsep}{4pt}
\renewcommand{\arraystretch}{1.15}
\begin{tabular}{@{}p{3.6cm}lp{7.6cm}p{2.4cm}@{}}
\toprule
\textbf{Reference} & \textbf{Variant} & \textbf{Link given in the paper} & \textbf{Status} \\
\midrule
\citet{polyakovskiy2014comprehensive} & TTP1 & \url{http://cs.adelaide.edu.au/~optlog/research/ttp.php} & Reachable \\
\citet{wagner2016stealing} & TTP1 & \url{http://cs.adelaide.edu.au/~optlog/research/ttp.php} & Reachable \\
\citet{el2017local} & TTP1 & \url{https://github.com/yafrani/ttplab} & Reachable \\
\citet{martins2017hseda} & TTP1 & \url{https://bitbucket.org/marcella_engcomp/hseda-ttp} & Reachable \\
\citet{el2018efficiently} & TTP1 & \url{https://github.com/yafrani/ttplab} & Reachable \\
\citet{nieto2018guided} & TTP1 & \url{https://gitlab.com/NIFR91/TTP-solver} & Reachable \\
\citet{bossek2021generating} & TTP1 & \url{https://github.com/jakobbossek/GECCO2021-ECPERM-ttp-evolving} & Reachable \\
\citet{zhang2021solving} & TTP1 & \url{https://github.com/zhangzitong/TTP} & Reachable \\
\citet{namazi2023solving} & TTP1 & \url{https://github.com/majid75/CoCo} & Reachable \\
\citet{nguyen2023simulated} & TTP1 & \url{https://github.com/hoangnqh/TTP-SAVI} & Reachable \\
\midrule
\citet{chagas2021non} & TTP2 & \url{https://github.com/jonatasbcchagas/nds-brkga_bi-ttp} & Reachable \\
\citet{chagas2022weighted} & TTP2 & \url{https://github.com/jonatasbcchagas/wsm_bittp} & Not found (HTTP 404) \\
\citet{santiyuda2024solving} & TTP2 & \url{https://github.com/gemsanyu/phn-ttp} & Reachable \\
\midrule
\citet{chagas2020ants} & ThOP & \url{https://github.com/jonatasbcchagas/aco_thop} & Reachable \\
\citet{chagas2022efficiently} & ThOP & \url{https://github.com/jonatasbcchagas/aco_thop} & Reachable \\
\bottomrule
\end{tabular}
\end{center}
''',
    'tab_notation.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{Notation for the TTP1 model used throughout the paper.}
\label{tab:notation}
\begin{tabular}{@{}ll@{}}
\toprule
\textbf{Symbol} & \textbf{Meaning} \\
\midrule
$n$ & number of cities in the TSP component \\
$d_{ij}$ & distance between cities $i$ and $j$ \\
$\pi=(1,\pi_2,\dots,\pi_n,1)$ & the closed tour of the thief (1 denotes the starting and ending city) \\
$m_i$ & number of items offered at city $\pi_i$ \\
$p_{ik},\,w_{ik}$ & profit and weight of item $k$ at city $\pi_i$ \\
$x_{ik}\in\{0,1\}$ & packing plan: $1$ if item $k$ at city $\pi_i$ is taken, 0 otherwise \\
$C$ & knapsack capacity ($\sum w_{ik}x_{ik}\le C$) \\
$W$ & accumulated weight carried after a city \\
$\vmax,\,\vmin$ & speed of the empty and the full knapsack \\
$v(W)=\vmax-\tfrac{W}{C}(\vmax-\vmin)$ & load-dependent travel speed \\
$R$ & renting rate charged per unit of travel time \\
$t(\pi,x)$ & total travel time; objective $\sum p_{ik}x_{ik}-R\,t(\pi,x)$ \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_review_scope.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{Scope of this review compared with the only prior systematic survey of the TTP \citep{sarkar2024travelling}.}
\label{tab:reviewscope}
\setlength{\tabcolsep}{4pt}
\begin{tabular}{@{}p{5.0cm}p{5.5cm}p{5.3cm}@{}}
\toprule
\textbf{Dimension} & \textbf{\citet{sarkar2024travelling}} & \textbf{This review} \\
\midrule
Search cut-off & mid-2023  & third quarter of 2026 \\
Primary studies & 68 & 93 + the prior survey \\
Variants covered & TTP1, PWT, TTP2, MTTP, DTTP, ThOP & the same six plus TTPTW, CCTTP, TTP-D \\
Formal variant taxonomy & narrative list & seven core $+$ two refining dimensions (Table~\ref{tab:taxonomy}) \\
Solution-method classification & exact / heuristic / metaheuristic / hyper-heuristic, confined to TTP1 and TTP2 & eight-family classification covering all nine variants (Figure~\ref{fig:coverage}) \\
Reinforcement-learning and quantum methods & flagged as a future direction, none yet published & 6 learning-based papers, three of them on reinforcement learning, and 1 quantum-annealing paper classified and reviewed (Section~\ref{sec:algorithms}) \\
Cross-paper performance comparison & not attempted & RDI $+$ Friedman/Nemenyi on 60 instances (Table~\ref{tab:rdi}) \\
Application motivation & logistics motivation given for the variants covered & stated motivation of each variant, with study counts (Table~\ref{tab:applications}) \\
Critical appraisal of research practice & research-gap section only, no per-paper coding & per-paper six-field appraisal $+$ five-point practice audit (Section~\ref{sec:critical}) \\
Reproducibility / open-source audit & not reported & per-paper, one row per study (Table~\ref{tab:reporting}), and a reachability check of code links (Appendix~\ref{app:code}) \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_roadmap.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{Where each research question is answered.}
\label{tab:roadmap}
\begin{tabular}{@{}p{0.7cm}p{1.3cm}p{5.5cm}p{7.0cm}@{}}
\toprule
\textbf{RQ} & \textbf{Section} & \textbf{Deliverable} & \textbf{Headline finding} \\
\midrule
RQ1 & \ref{sec:taxonomy} & a nine-dimension variant taxonomy (Table~\ref{tab:taxonomy}) & the design space is a directed acyclic graph rooted at TTP1; many dimension combinations are still unformalized \\
RQ2 & \ref{sec:algorithms} & an eight-family algorithm classification and a TTP-specific coupling-strategy axis (Table~\ref{tab:coupling}) & population-based metaheuristics dominate with 38 of 84 method-proposing studies; only CoCo and CoCo$^+$ evaluate each move against the joint objective, which is why they lead RQ3's ranking \\
RQ3 & \ref{sec:performance} & a common-pool RDI comparison of nine TTP1 methods with Friedman and Nemenyi tests & CoCo$^+$ and CoCo lead with a statistically significant margin; earlier ``category-leading'' claims do not survive a common pool \\
RQ4 & \ref{sec:biblio}, \ref{sec:future} & an algorithm--variant coverage map and a prioritized agenda & MTTP, DTTP, TTPTW, CCTTP and TTP-D are each served by one or two families \\
RQ5 & \ref{sec:critical} & a per-paper appraisal and a five-point practice audit (Table~\ref{tab:reporting}) & about one study in three reports a statistical test, 39 of 93 use a standard benchmark, and 26 of 93 release code \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_cat_ttp1.tex': r'''{\footnotesize
\begin{longtable}{@{}p{2.8cm}cp{3.0cm}p{4.9cm}p{7.9cm}@{}}
\caption{Catalog of TTP1 contributions, grouped by the four sub-epochs of Section~\ref{sec:lit_core}.}\\\label{tab:cat_ttp1}\\
\toprule
\textbf{Reference} & \textbf{Year} & \textbf{Algorithm} & \textbf{Instances} & \textbf{Key result} \\
\midrule
\endfirsthead
\multicolumn{5}{l}{\footnotesize\itshape Table~\ref{tab:cat_ttp1} continued}\\ \toprule
\textbf{Reference} & \textbf{Year} & \textbf{Algorithm} & \textbf{Instances} & \textbf{Key result} \\
\midrule
\endhead
\midrule \multicolumn{5}{r}{\footnotesize continued on next page}\\
\endfoot
\bottomrule
\endlastfoot
\midrule \multicolumn{5}{l}{\textbf{Foundations and benchmarking, 2013--2015} (10 papers)} \\ \midrule
\citet{bonyadi2013travelling} & 2013 & RLS / EA baseline & Introduces the problem, no benchmark yet & Defines the TTP and gives the first RLS and EA baselines. \\
\citet{bonyadi2014socially} & 2014 & CoSolver, DH & 45 self-generated instances, 3--33 cities & CoSolver beats the sequential DH baseline, and the advantage grows with the rent rate. \\
\citet{mei2014improving} & 2014 & MATLS & 6 large instances, 11{,}849--33{,}810 cities & MATLS is positive where RLS and EA are negative on every large instance. \\
\citet{polyakovskiy2014comprehensive} & 2014 & SH, S1--S5, C1--C6, RLS, EA & Builds the 9{,}720-instance benchmark & Establishes the standard benchmark and the scoring and iterative heuristics used as baselines since. \\
\citet{beham2015optimization} & 2015 & HeuristicLab MA & berlin52-based TTP, KCPTP, and OP instances & A variegation operator significantly improves solution quality across all three problem types. \\
\citet{el2015cosolver2b} & 2015 & CoSolver2B & 16 uncorrelated-similar-weight instances, 52--33{,}810 cities & CoSolver2B-bestfit beats EA and RLS on most instances but times out on the largest. \\
\citet{faulkner2015approximate} & 2015 & S1--S5, C1--C6 & 72-instance representative subset & Establishes S5 as the dominant simple baseline, especially on bounded-strongly-correlated instances. \\
\citet{gupta2015greedy} & 2015 & G1 / G2 & Full 9{,}720-instance benchmark & G1 and G2 beat the SH scoring baseline on almost every instance. \\
\citet{mei2015heuristic} & 2015 & GAIN / PICKFUNC (GP) & 6 large TSPLIB bases, up to 33{,}810 cities & GP-evolved heuristics match a hand-designed heuristic and beat RLS and EA on every test instance. \\
\citet{strzezek2015divergene} & 2015 & DiverGene & berlin52 & A dissimilarity-aware selection operator improves the objective, most on the hardest (TTP) of the three problems tested. \\
\midrule \multicolumn{5}{l}{\textbf{Coordinated search, 2016--2019} (23 papers)} \\ \midrule
\citet{el2016population} & 2016 & MA2B & 20 TSPLIB bases $\times$ 3 categories & MA2B wins on small and mid-size instances; S5 and CS2SA win on Category~1 and on large instances. \\
\citet{lourencco2016evolutionary} & 2016 & EA & eil51, eil76, kroA100 & The EA beats a sequential density-heuristic baseline on 4 of 6 instances; TTP-optimal tours differ substantially from TSP-optimal ones. \\
\citet{mei2016investigation} & 2016 & Cooperative co-evolution & 39 self-generated instances, up to 100 cities & Confirms the objective cannot be cleanly separated; a plain memetic algorithm still beats every coevolution variant. \\
\citet{wagner2016stealing} & 2016 & MMAS / ACO & Over 1{,}100 instances, up to 1{,}000 cities & MMAS wins on small and medium instances; S5 dominates on the largest. \\
\citet{wu2016impact} & 2016 & Renting-rate analysis & Theoretical, no benchmark & Only a few renting-rate values yield genuinely hard instances; gives a constructed instance where a (1+1)~EA provably fails. \\
\citet{el2017local} & 2017 & JNB / J2B & 180 instances, 51--1{,}304 cities & JNB and J2B beat EA and RLS on small and medium instances; tour initialization matters far more than packing initialization. \\
\citet{karder2017solving} & 2017 & Switchover LS & 60 berlin52-based instances & Switchover and SASEGASA find a new best on 59 of 60 instances. \\
\citet{martins2017hseda} & 2017 & HSEDA & 135 instances, 59\% of the benchmark by city count & HSEDA's edge over MMAS and S5 grows with city count. \\
\citet{moeini2017hybrid} & 2017 & Hybrid heuristic & 40 instances, eil51 and eil76 classes & The hybrid EA improves several eil51 best-known values but is mostly worse on eil76. \\
\citet{vieira2017genetic} & 2017 & MCGA & 36-instance subset of \citet{faulkner2015approximate} & MCGA wins on 56\% of instances, mainly the smaller bases; S5 wins on the largest. \\
\citet{wu2017exact} & 2017 & Branch-and-bound, DP, CP & 432 exactly-solved instances, up to 20 cities & MA2B is closest to optimal overall; strongly-correlated instances are the hardest for every heuristic. \\
\citet{alharbi2018design} & 2018 & mod-ABC & eil51 and eil76, all three KP types & mod-ABC beats the density and greedy baselines on every instance but degrades with more items per city. \\
\citet{alharbi2018hybrid} & 2018 & GATS & 540 instances from the standard benchmark & GATS beats EA and RLS at every capacity level except on one TSP base. \\
\citet{araujo2018novel} & 2018 & LCRVND (GPU) & 60 instances from \citet{bonyadi2013travelling} & GPU-accelerated LCRVND reaches the literature's target faster on 54 of 60 instances. \\
\citet{el2018efficiently} & 2018 & CS2SA / CS2SA$^\ast$ & 20 TSPLIB bases $\times$ 3 categories & CS2SA$^\ast$ wins on most mid and large instances; S5 still wins Category~1. \\
\citet{el2018hyperheuristic} & 2018 & GPHS$^\ast$ & 51--783 cities, all three KP types & GPHS$^\ast$ significantly beats S5 and MA2B on several mid-size instances. \\
\citet{nieto2018guided} & 2018 & RES\_GLS & 12 instances, 52--4{,}461 cities & RES\_GLS finds 3 new best-known values, concentrated on the smaller instances. \\
\citet{przybylek2018decomposition} & 2018 & CoSolver-H, MCTS, ACO & Generated and TSPLIB-derived instances & CoSolver-H never produces a bad solution and gives the best average gain of the three decompositions. \\
\citet{wagner2018case} & 2018 & Algorithm selection & Portfolio of existing TTP1 solvers & The best selectors close 90\% of the single-best-oracle gap; CS2SA has the highest marginal contribution despite the lowest standalone performance. \\
\citet{ali2019novel} & 2019 & BMV-DE & 13 TTP instances derived from TSPLIB & BMV-DE gives the lowest average error of the compared methods. \\
\citet{namazi2019cooperative} & 2019 & CoCo (PGCH, SGCH, LGCH, MBFS) & Categorized 60-instance benchmark & Establishes CoCo as the new state of the art on the categorized benchmark. \\
\citet{namazi2019profit} & 2019 & PGCH & Categorized 60-instance benchmark & The joint reversal-and-reassignment move consistently beats plain 2-opt. \\
\citet{zouari2019new} & 2019 & CMMACS, OMMACS, DMMACS & 10 bounded-strongly-correlated a280 instances & The online variant OMMACS outperforms both the centralized and the distributed variant. \\
\midrule \multicolumn{5}{l}{\textbf{Surrogate, diversity, and learning-inspired methods, 2020--2022} (11 papers)} \\ \midrule
\citet{ali2020differential} & 2020 & DB-DE & 13 hard strongly-correlated instances & DB-DE ranks second in objective and first in runtime, saving over 90\% of the time budget. \\
\citet{ali2020hyper} & 2020 & SR-IO, KAB-IO, RD-MSA, RPD-MSA & 9 instances, eil51, eil76, a280 & The proposed hyper-heuristics win on eil51 and eil76; the prior state of the art still wins on a280. \\
\citet{maity2020efficient} & 2020 & Hybrid local search & 28 TSPLIB bases $\times$ 3 categories & Best on Category~1; MATLS and CS2SA remain better on Categories~2 and~3. \\
\citet{namazi2020surrogate} & 2020 & CoCo-SM (SVR) & 96 medium and large instances & The surrogate filters 30\% of initial tours at the cost of one missed best in ten runs, saving 30\% of solver time. \\
\citet{ali2021novel} & 2021 & ESA & Large-instance subset, 76--33{,}810 cities & Reports very large gains over S5 and CS2SA, large enough that an objective-scaling discrepancy is likely. \\
\citet{bossek2021generating} & 2021 & Evolved instances & Theoretical, instance-generation study & Generates instances with a controllable, portfolio-dependent performance gap between algorithms. \\
\citet{misir2021benchmark} & 2021 & Benchmark analysis & Subsets of the 9{,}720-instance benchmark & Most reduced subsets preserve over 90\% of the full benchmark's algorithm ranking. \\
\citet{naumov2021identifying} & 2021 & Cargo-bike CVRP & Last-mile delivery case study, not the standard benchmark & A CVRP formulation reaches the same quality as competing formulations but converges far faster. \\
\citet{zhang2021solving} & 2021 & RWS & Categorized 60-instance benchmark & RWS is strongest on Category~C, where capacity is high and items are uncorrelated. \\
\citet{nikfarjam2022coevolutionary} & 2022 & Co-EA (QD + EDO) & 18 instances, eil51, pr152, a280 & Co-EA beats standard evolutionary diversity optimization on structural entropy on 14 of 18 instances. \\
\citet{nikfarjam2022evolutionary} & 2022 & EDO (entropy) & 18 instances, eil51, pr152, a280 & A combined edge-and-item diversity measure gives the best overall diversity. \\
\midrule \multicolumn{5}{l}{\textbf{Modern paradigms and new theory, 2023--2026} (13 papers)} \\ \midrule
\citet{namazi2023solving} & 2023 & CoCoP / CoCoL (CoCo$^+$) & Categorized 60-instance benchmark & Sets new best-known values on most mid and large instances; the strongest method on the categorized benchmark to date. \\
\citet{nguyen2023simulated} & 2023 & SAVI & Categorized 60-instance benchmark & Competitive overall but the weakest of the nine pool methods on Category~A. \\
\citet{rodriguez2023sequence} & 2023 & Sequence-based hyper-heuristic & 60 self-generated 5-city instances & Trained hyper-heuristic sequences beat random sequences and generalize across correlation levels; the standard benchmark is not used. \\
\citet{ni2024leveraging} & 2024 & Symbolic-regression initialization & 450 instances from the standard benchmark & The symbolic-regression initializer ranks first on about 80\% of instances with fewer objective evaluations. \\
\citet{nikfarjam2024use} & 2024 & BMBEA (QD, journal) & 18 main plus 27 extended large-scale instances & Gives the highest mean on most instances and improves the best-known value on two large instances. \\
\citet{wu2024reinforcens} & 2024 & ReinforceNS & 60-instance benchmark plus 18 GECCO-2023 competition instances & Reports twelve new best-known results and beats CoCo on 38 of 60 instances. \\
\citet{xiang2024five} & 2024 & FECOIMO & 39 self-generated instances & Ranks first by Friedman test on all three statistics tried; the standard benchmark is not used. \\
\citet{pathirage2025evolutionary} & 2025 & Multitasking EA & 10 instances, 5 TSP bases $\times$ 2 KP types & The multitasking advantage concentrates under tight time budgets and on larger instances. \\
\citet{zhang2025two} & 2025 & ICA + GACO & Small-instance subset, 51--280 cities & GACO-GISS beats a Lin-Kernighan-based baseline despite producing longer tours, confirming TTP-aware routing helps. \\
\citet{eube2026approximation} & 2026 & Bi-criteria approximation & Theoretical, no instances & Gives the first approximation algorithm for TTP1, a polynomial-time bi-criteria Pareto set and a Weighted-TSP approximation. \\
\citet{eube2026effective} & 2026 & Weighted-TSP DP / approximation & 25 instances, eil51-based plus two larger bases & The exact DP is never worse than S5 and improves the objective by 13.6\% on average. \\
\citet{kapancioglu2026formulations} & 2026 & MILP formulations & 62 generated plus a benchmark subset & Finds new optimal values for 15 benchmark instances, including two with over 50 nodes. \\
\citet{tran2026integrating} & 2026 & ACO-RSA & TSPLIB bases plus the categorized A/B/C set & Competitive on small and medium instances; S5 and SAVI still win on the largest. \\
\end{longtable}
}
''',
    'tab_cat_small.tex': r'''\begin{center}
\footnotesize
\captionof{table}{Catalog of PWT, MTTP and DTTP, and ThOP contributions.}
\label{tab:cat_small}
\begin{tabular}{@{}p{2.8cm}cp{3.0cm}p{4.9cm}p{7.9cm}@{}}
\toprule
\textbf{Reference} & \textbf{Year} & \textbf{Algorithm} & \textbf{Instances} & \textbf{Key result} \\
\midrule
\multicolumn{5}{l}{\textbf{PWT} (4 papers)} \\ \midrule
\citet{polyakovskiy2015packing} & 2015 & PWT MIP & eil51/76/101 (small) and pla33810/85900 (large), 3 KP types & Exact MIP is optimal only on small instances; a fast heuristic is near-optimal on large ones. \\
\citet{polyakovskiy2017packing} & 2017 & PWT exact (NP-hardness) & Same small and large TSP bases as polyakovskiy\allowbreak{}2015packing & Pre-processing removes 19-32\% of items; a fast heuristic solves pla85900 in about 4 hours. \\
\citet{neumann2019fully} & 2019 & FPTAS (PWT) & 27 mid-size instances plus a large-range synthetic set & The FPTAS reaches a 100\% approximation ratio and is far faster than the exact baselines. \\
\citet{don2026greedy} & 2026 & Greedy PACK + HH & 300 PWT instances: 10 TSP bases x 30 random fixed tours & A hyper-heuristic beats the individual greedy scoring rules on most instances. \\
\midrule
\multicolumn{5}{l}{\textbf{MTTP and DTTP} (4 papers)} \\ \midrule
\citet{chand2016fast} & 2016 & MTTP fast heuristic & 72 instances, the same subset used by faulkner2015\allowbreak{}approximate & Two thieves is the sweet spot; MTTP beats the best single-thief TTP solver by 50-100\%. \\
\citet{herring2020dynamic} & 2020 & DMOTTP EA & 4 TSP bases (52-2319 cities) x 3 knapsack categories & A solver-based restart strategy gives the best hypervolume overall. \\
\citet{herring2020responsive} & 2020 & DTTP responsive EA & 6 TSP bases (52-318 cities) x 3 knapsack categories, city-location dynamics & Migrating the population after a disruption beats simple re-seeding on most generations. \\
\citet{sachdeva2020dynamic} & 2020 & DTTP EA & 9 instances from a280/fnl4461/pla33810, 72 disruption scenarios & Restarting works best after item disruptions; recovering in place works best for cities. \\
\midrule
\multicolumn{5}{l}{\textbf{ThOP} (8 papers)} \\ \midrule
\citet{santos2018thief} & 2018 & ILS / BRKGA (ThOP) & 432 instances: 4 TSP bases x item, capacity, and time-budget levels & BRKGA gives the most stable convergence overall; ILS wins only on the smallest instances. \\
\citet{chagas2020ants} & 2020 & ACO (ThOP) & The 432-instance benchmark of santos2018thief & The ACO variants beat the prior best on 95\% of instances, a 320\% average profit gain. \\
\citet{faeda2020genetic} & 2020 & thopGA & The 432-instance benchmark of santos2018thief & thopGA wins on over 70\% of instances and converges more stably than BRKGA. \\
\citet{chagas2022efficiently} & 2022 & ACO++ & The 432-instance benchmark plus 36 derived Orienteering Problem instances & ACO++ matches or beats ILS, BRKGA, GA, and ACO on almost every instance. \\
\citet{bloch2023polynomial} & 2023 & PTAS (DAG ThOP) & Theoretical (no instances) & First PTAS for ThOP on directed acyclic graphs, plus an FPTAS on restricted classes. \\
\citet{huynh2023self} & 2023 & SAAS-HC & The 432-instance benchmark of santos2018thief & SAAS-HC matches the quality of ACO++ using one parameter set instead of 48 tuned sets. \\
\citet{bloch2024thief} & 2024 & series-parallel theory & Theoretical (no instances) & A series-parallel graph transformation extends the ThOP hardness and approximation results. \\
\citet{bloch2025algorithms} & 2025 & FPTAS (ThOP) & Theoretical (no instances) & FPTAS for ThOP on constant-speed and clique-structured graph classes. \\
\bottomrule
\end{tabular}
\end{center}
''',
    'tab_coupling.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{The 84 method-proposing studies by how they handle the tour--packing coupling, against the eight algorithm families of Table~\ref{tab:reporting}.}
\label{tab:coupling}
\begin{tabular}{@{}lccccccccc@{}}
\toprule
\textbf{Coupling strategy} & \textbf{Exact} & \textbf{Approx.} & \textbf{Constr.} & \textbf{Single-sol.} & \textbf{Pop.-based} & \textbf{Hyper-h.} & \textbf{Learning} & \textbf{Quantum} & \textbf{Total} \\
\midrule
Sequential / decoupled & 2 & 5 & 5 & 0 & 0 & 0 & 1 & 0 & 13 \\
Periodic / alternating & 0 & 0 & 1 & 4 & 4 & 7 & 1 & 0 & 17 \\
Joint search & 3 & 0 & 0 & 9 & 34 & 0 & 0 & 1 & 47 \\
Move-synchronized & 0 & 0 & 0 & 3 & 0 & 0 & 0 & 0 & 3 \\
Learned coordination & 0 & 0 & 0 & 0 & 0 & 0 & 4 & 0 & 4 \\
\midrule
Total & 5 & 5 & 6 & 16 & 38 & 7 & 6 & 1 & 84 \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_taxonomy.tex': r'''\begin{center}
\footnotesize
\captionof{table}{Consolidated taxonomy of the TTP variants.}
\label{tab:taxonomy}
\begin{flushleft}
\textbf{Legend: each dimension below and the values it can take.}
\end{flushleft}
\begin{tabular}{@{}p{3.0cm}p{16.3cm}@{}}
\textbf{Objective structure} & single-objective; bi-objective (travel time vs.\ profit)\\
\textbf{Route structure} & fixed route; full tour (all cities); selective routing (subset of cities)\\
\textbf{Timing constraint} & unconstrained; per-city time windows; global time budget; epoch-based disruption; drone flight synchronization\\
\textbf{Agent multiplicity} & single thief; multiple cooperative thieves (shared capacity); vehicle $+$ drone\\
\textbf{Knapsack uncertainty} & deterministic weights; stochastic weights under a chance constraint\\
\textbf{Problem dynamics} & static instance; dynamic instance (item/city toggling, recovery)\\
\textbf{Value model} & fixed item profits; time-decaying profits with a per-item drop rate\\
\textbf{Cost / tradeoff model} & time-renting (profit $-$ $R\cdot$time); value-drop-rate; global-time-budget\\
\textbf{Speed--weight law} & linear decrease with load; linear or constant (ThOP theory; load-dependent for TTP-D)\\
\end{tabular}\\[6pt]
\centering
\setlength{\tabcolsep}{3pt}
\begin{tabular}{@{}lp{1.35cm}p{1.35cm}p{2.5cm}p{2.5cm}p{2.5cm}p{1.0cm}p{1.35cm}p{1.55cm}p{1.55cm}cp{2.0cm}@{}}
\toprule
\textbf{Variant} & \textbf{Obj.} & \textbf{Route} & \textbf{Timing} & \textbf{Agents} & \textbf{KP unc.} & \textbf{Dyn.} & \textbf{Value} & \textbf{Cost model} & \textbf{Speed law} & \textbf{Period} & \textbf{Ref. (oldest--newest)} \\
\midrule
\textbf{TTP1} & Single-objective & Full tour (all cities) & Unconstrained & Single thief & Deterministic & Static & Fixed profits & time-renting & linear & 2013--2026 & \citet{bonyadi2013travelling}; \citet{eube2026approximation} \\
\textbf{PWT} & Single-objective & Fixed route & Unconstrained & Single thief & Deterministic & Static & Fixed profits & time-renting & linear & 2015--2026 & \citet{polyakovskiy2015packing}; \citet{don2026greedy} \\
\textbf{TTP2} & Bi-objective (time vs profit) & Full tour & Unconstrained & Single thief & Deterministic & Static & Time-decaying profits (drop rate) & value-drop-rate & linear & 2013--2026 & \citet{bonyadi2013travelling}; \citet{viet2026advanced} \\
\textbf{MTTP} & Single-objective & Full tour & Unconstrained & Multiple cooperative thieves (shared capacity) & Deterministic & Static & Fixed profits & time-renting & linear & 2016 & \citet{chand2016fast} \\
\textbf{DTTP} & Single- and bi-objective & Full tour & Epoch-based disruption & Single thief & Deterministic & Dynamic (item/city toggling) & Fixed profits & time-renting & linear & 2020 & \citet{sachdeva2020dynamic} \\
\textbf{ThOP} & Single-objective & Selective routing (subset of cities) & Global time budget & Single thief & Deterministic & Static & Fixed profits & global-time-budget & linear/constant & 2018--2025 & \citet{santos2018thief}; \citet{bloch2025algorithms} \\
\textbf{TTPTW} & Single-objective & Full tour & Per-city time windows + mandatory wait & Single thief & Deterministic & Static & Fixed profits & time-renting & linear & 2026 & \citet{angmalisang2026traveling} \\
\textbf{CCTTP} & Single-objective & Full tour & Unconstrained & Single thief & Stochastic weights (chance constraint) & Static & Fixed profits & time-renting & linear & 2024--2025 & \citet{pathirage2024chance}; \citet{don2025weighted} \\
\textbf{TTP-D} & Single-objective & Full vehicle tour + drone sorties & Flight synchronization (drone launch/rendezvous) & Vehicle + drone (two cooperating agents) & Deterministic & Static & Fixed profits & time-renting & linear (load-dep.) & 2026 & \citet{murjani2026drive} \\
\bottomrule
\end{tabular}
\end{center}
''',
    'tab_ttp2_thop.tex': r'''\begin{table}[htbp]
\centering
\footnotesize
\caption{Reported evidence for TTP2 and ThOP methods, oldest first. \ding{108}\,yes, \ding{109}\,no. HV: hypervolume.}
\label{tab:ttp2thop}
\setlength{\tabcolsep}{3pt}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{@{}p{4.4cm}p{3.1cm}ccp{6.0cm}@{}}
\toprule
\textbf{Method (year)} & \textbf{Tested on} & \textbf{Rival} & \textbf{Test} & \textbf{Main reported result} \\
\midrule
\multicolumn{5}{@{}l}{\textbf{Bi-objective TTP (TTP2)}} \\
\cmidrule(r){1-5}
NTGA, specialized encoding (2019) \citep{laszczyk2019specialized} & 12 instances, eil51 (51 cities) & \ding{109} & \ding{108} & Specialized encoding raises HV over NSGA-II on all 12 \\
MO-DP, dynamic programming (2020) \citep{santana2020dynamic} & 9 competition instances & \ding{108} & \ding{109} & Close to the top three entries; HPI still leads \\
NDS-BRKGA, non-dominated-sorting biased random-key GA (2021) \citep{chagas2021non} & 9 competition instances & \ding{108} & \ding{109} & 1st at EMO-2019 (best HV on 7 of 9), 2nd at GECCO-2019 \\
NTGA2, second non-dominated tournament GA (2021) \citep{myszkowski2021diversity} & 12 instances, eil51 & \ding{108} & \ding{108} & Best HV (+0.72\% over NTGA); 3rd at GECCO-2019 \\
MOACO-PW, multi-objective ant colony (2021) \citep{yang2021moea} & Small custom set; competition instances & \ding{108} & \ding{109} & Beats its baselines on 6 of 9 competition instances; baseline values quoted \\
WSM, weighted-sum method (2022) \citep{chagas2022weighted} & About 960 instances; competition instances & \ding{108} & \ding{108} & Better than NDS-BRKGA on 82.2\% of instances, worse on 15\% \\
MOFECO-MS, five-element cycle optimization (2024) \citep{xiang2024multi} & 9 instances, up to 280 cities & \ding{109} & \ding{109} & Best HV among 8 general-purpose metaheuristics \\
MORL, multi-objective reinforcement learning (2024) \citep{santiyuda2024solving} & 10 graphs, up to 1,379 cities & \ding{108} & \ding{108} & Competitive with NDS-BRKGA only on many-item instances, but far faster \\
QA, quantum annealing (2026) \citep{viet2026advanced} & 6 instances, up to 280 cities & \ding{109} & \ding{109} & Up to 9 times faster than classical MOEAs at similar HV \\
\midrule
\multicolumn{5}{@{}l}{\textbf{Thief Orienteering Problem (ThOP), all on the same 432-instance suite}} \\
\cmidrule(r){1-5}
ILS, iterated local search, and BRKGA (2018) \citep{santos2018thief} & 432 instances & \ding{109} & \ding{109} & BRKGA is better than ILS on most instances \\
thopGA, genetic algorithm (2020) \citep{faeda2020genetic} & 432 instances & \ding{108} & \ding{108} & Best on 304 of 432; baselines quoted from another machine \\
ACO, ant colony optimization (2020) \citep{chagas2020ants} & 432 instances & \ding{108} & \ding{108} & Better than ILS and BRKGA on 410 to 419 of 432 \\
ACO++, improved ACO (2022) \citep{chagas2022efficiently} & 432 instances & \ding{108} & \ding{108} & Better on at least 96\% of instances \\
SAAS-HC, self-adaptive ant system (2023) \citep{huynh2023self} & 432 instances & \ding{108} & \ding{108} & Best known on 330 instances (ACO++: 180); vs ACO++: 176 better, 170 worse \\
PTAS, approximation scheme (2023, 2024) \citep{bloch2023polynomial,bloch2024thief} & Theory only & -- & -- & PTAS on acyclic and series-parallel graphs; in general no constant-factor approximation unless P $=$ NP \\
\bottomrule
\end{tabular}
\end{table}
''',
    'tab_unexplored.tex': r'''
\begin{center}
\footnotesize
\captionof{table}{Unexplored variant combinations in the TTP design space.}
\label{tab:unexplored}
\setlength{\tabcolsep}{3pt}
\begin{tabular}{@{}llp{5cm}cp{7.5cm}@{}}
\toprule
\textbf{Base} & \textbf{+ dimension} & \textbf{Resulting formulation} & \textbf{Studied?} & \textbf{Suggested name} \\
\midrule
CCTTP & + per-city time windows & Stochastic time-window TTP & No & Chance-constrained TTPTW (CC-TTPTW) \\
MTTP & + bi-objective (time vs profit) & Multi-thief bi-objective TTP & No & Bi-objective MTTP \\
CCTTP & + dynamic instance & Dynamic chance-constrained TTP & No & Dynamic CCTTP \\
ThOP & + multiple agents & Multi-thief orienteering & No & Multi-agent ThOP \\
PWT & + chance constraint & Stochastic packing-while-traveling & Partial & Chance-constrained PWT (\cite{don2026greedy}: greedy only) \\
TTP2 & + dynamic instance & Dynamic bi-objective TTP & Yes & DMOTTP \citep{herring2020dynamic} \\
TTPTW & + stochastic weights & Stochastic time-window TTP & No & Stochastic TTPTW \\
DTTP & + multiple thieves & Dynamic multi-thief TTP & No & Dynamic MTTP \\
TTP1 & + time-decaying profit + time windows & Time-sensitive value TTPTW & No & Value-decay TTPTW \\
TTP-D & + multiple drones / bi-objective & Multi-drone or bi-objective drone TTP & No & \citep{murjani2026drive} covers a single vehicle + single drone, single-objective \\
\bottomrule
\end{tabular}
\end{center}
''',
    'tab_variant_glance.tex': r'''\begin{center}
\footnotesize
\captionof{table}{The nine TTP variants at a glance.}
\label{tab:glance}
\setlength{\tabcolsep}{3pt}
\begin{tabular}{@{}lcp{5.5cm}p{4.5cm}p{3.0cm}p{5.0cm}@{}}
\toprule
\textbf{Variant} & \textbf{\#} & \textbf{Dominant families} & \textbf{State of the art} & \textbf{Benchmark} & \textbf{Key open gap} \\
\midrule
TTP1 & 57 & Pop.-based, single-sol., hyper-heur., learning & CoCo$^+$ \citep{namazi2023solving} & Categorized 60-inst. & Shared reporting standard; real-data studies \\
PWT & 4 & Exact, approximation, greedy & FPTAS \citep{neumann2019fully} & Derived from TTP & Chance-constrained PWT has only a greedy result \\
TTP2 & 18 & MO-evolutionary, MORL, quantum & NDS-BRKGA \citep{chagas2021non}, WSM \citep{chagas2022weighted} & BI-TTP comp.\ suite & Learning and quantum methods on the competition suite \\
MTTP & 1 & Constructive heuristic & Fast heuristics \citep{chand2016fast} & Reused TTP inst. & No EA/RL/QD solver; no benchmark \\
DTTP & 3 & Pop.-based EA (dynamic) & Restart/recover EA \citep{sachdeva2020dynamic} & 72 dynamic scenarios & No learning-based solver \\
ThOP & 8 & Swarm, approximation theory & SAAS-HC \citep{huynh2023self} & 432-inst.\ suite & No learned or hyper-heuristic solver \\
TTPTW & 1 & Pop.-based EA & DSEA \citep{angmalisang2026traveling} & Derived (3 tightness) & No exact/approx/RL \\
CCTTP & 2 & Surrogate, scenario heuristic & S5/C5 surrogate \citep{pathirage2024chance} & Derived (stochastic) & No population-based or learning-based solver \\
TTP-D & 1 & Learning-based (attention DRL) $+$ MILP & MILP $+$ DRL hybrid \citep{murjani2026drive} & Two derived families & Single paper; benchmark not standardized \\
\bottomrule
\end{tabular}
\vspace{2pt}

\footnotesize\emph{Abbreviations used above:} EA = evolutionary algorithm; RL = reinforcement learning; MORL = multi-objective reinforcement learning; QD = quality diversity; DRL = deep reinforcement learning; MILP = mixed-integer linear programming.
\end{center}
''',
}

def main():
    check = "--check" in sys.argv
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    out = dict(STATIC)
    out["tab_rdi.tex"] = rdi_table(wb)
    out["tab_reporting.tex"] = reporting_table(wb)
    bad = 0
    for name, text in sorted(out.items()):
        path = OUT / name
        if check:
            same = path.exists() and path.read_text(encoding="utf-8") == text
            print(("same     " if same else "DIFFERS  ") + name); bad += not same
        else:
            path.write_text(text, encoding="utf-8"); print("wrote", path)
    sys.exit(1 if bad else 0)

if __name__ == "__main__":
    main()
