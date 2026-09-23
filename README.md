# ttp_crt_git

Supplementary material of the review **"A Critical Review and Taxonomy of the Traveling Thief Problem and Its Variants"** (M. Khemakhem, submitted to *Computers & Operations Research*).

Corpus: 93 primary studies plus one prior survey, published between 2013 and the third quarter of 2026. Link check of code repositories: 21 September 2026.

**What this repository contains.** The per-paper fact-sheet compendium, the raw results behind the common-pool comparison, the data workbook read by the figure script, the figure script, the generated figures, the TikZ sources of the schematic figures, and the LaTeX source of the 15 tables of the paper.

**Regenerating the figures.** `pip install -r requirements.txt`, then `python3 scripts/gen_figures.py`. The six data figures are rebuilt from the files in `data/` and written to `figs/`. They were checked pixel by pixel against the figures of the paper.

**What it does not contain.** The master workbook, the audit scripts that filled its audit sheets and the appraisal-notes file are not included. The audit flags in the `Reporting_Audit` sheet are therefore provided as recorded data, not recomputed.

---

## 1. Folder map

```
ttp_crt_git/
  README.md
  requirements.txt      Python packages used by the scripts
  .gitignore
  scripts/              gen_figures.py, gen_tables.py, gen_tikz.py
  data/                 compendium and the data workbook
  figs/                 the six data figures (PNG) + the four TikZ figures, their wrapper and compiled PDF
  tables/               the 15 LaTeX tables of the paper (tab_*.tex)
```

---

## 2. `data/`

### 2.1 `ttp_biblio_meta.tex` (with `ttp_biblio_meta.pdf`)
The per-paper fact-sheet compendium. It is a landscape LaTeX document with one block per paper, grouped by variant. Each block gives the title, the abstract, the contribution, the main claims, the evidence base (instances, runs, statistical test, budget), the stated limitations, and short lists of strengths and weaknesses, all under the fixed six-field appraisal schema described in Section 2.2 of the paper. Only facts stated in a paper, or checkable by reading it, were recorded.

It is the record behind every per-paper count in the paper (Section 6.1 and Appendix A). The audit of Section 6.1 was built from its fields (Metric, Stat. test, Budget / Runs, Evidence base).

To rebuild the PDF: `pdflatex ttp_biblio_meta.tex`, `bibtex ttp_biblio_meta`, then `pdflatex` twice. It reads `ttp_references.bib` (Section 2.3).

### 2.2 `ttp_crt_data.xlsx`
The single data workbook read by `gen_figures.py`. Each sheet has a colored header row.

| Sheet | Content | Feeds |
|---|---|---|
| `Corpus` | 94 works (93 primary studies plus the prior survey) with `CiteKey`, `Year`, `Variant(s)`, `Pub. status`, `Venue`, `Family (L1)`, `Coupling strategy`. Cross-listed studies carry two variants separated by a comma. `Coupling strategy` is blank for the 10 non-algorithmic works (the survey plus the 9 studies that define a benchmark, audit practice, or map applications without proposing a method); it holds one of the five values of Table 4 for the other 84. | Figures 4, 8, 9, the period split of 7, and Table 4 |
| `Venues` | 54 journal and conference venues with paper counts and type (preprints excluded). The row order is kept because it decides which tied venues appear in the top 11. | Figure 8b |
| `Reporting_Audit` | 93 primary studies, 17 columns: evidence class, reporting items (Y/N) and the five practice flags. | Figure 7, Table 15 |
| `Theme_Counts` | Strength and weakness mentions per theme (9 rows). | Figure 6 |
| `CatABC_Objectives` | Raw input of the common-pool comparison of nine TTP1 methods (Section 5.4, Table 6, Figure 5): 60 instances (20 per category A, B, C) and one objective column per method (CoCo+, SAVI, RWS, CoCo, CS2SA*, CS2SA, MA2B, S5, MATLS). Each value comes from the results table of the paper named in the column header. S5 and MATLS values come from the re-evaluation in Zhang et al. (2021), because their own authors did not evaluate on this set. | Figure 5, Table 6 |
| `Taxonomy` | The nine variants along the seven core taxonomy dimensions (`Objective structure`, `Route structure`, `Timing constraint`, `Agent multiplicity`, `Knapsack uncertainty`, `Problem dynamics`, `Value model`), plus complexity, benchmark, year, and foundational reference. Mirrors `Tab1_Taxonomy` in the source workbook (`ref_tabs/`). | Table 10 |
| `Unexplored` | The ten unexplored-combination candidates: existing variant, the dimension it would combine with, the resulting (unexplored) formulation, whether it has been studied, and a suggested name. One row (`TTP2 + dynamic instance`) is marked `Studied? Yes` and is shown for contrast, not as a gap, since DTTP already fills it. Mirrors `Tab2_Unexplored` in the source workbook. | Table 11 |

From `CatABC_Objectives`, `gen_figures.py` computes the per-instance RDI, the mean Friedman ranks, the best-in-pool counts, the Friedman statistic (chi-square 211.2, k = 9, N = 60) and the Nemenyi critical difference (1.55).

### 2.3 `ttp_references.bib`
The bibliography of the review, all reference files merged into one. It is read by `ttp_biblio_meta.tex`.

---

## 3. `scripts/`

| Script | Output |
|---|---|
| `gen_figures.py` | the six data figures in `figs/` |
| `gen_tables.py` | the 15 tables in `tables/` (`--check` compares without writing) |
| `gen_tikz.py` | the four TikZ figures and the wrapper `tikz_figures.tex` in `figs/` (`--check` compares without writing) |

All three write files identical to the ones submitted with the paper.

**How the two new scripts work.** In `gen_tables.py`, `tab_rdi.tex` is computed from `CatABC_Objectives` and the rows and subtotals of `tab_reporting.tex` are computed from `Reporting_Audit` and `Corpus`. The other 13 tables are hand-written text (captions, catalogs, links), stored in the script and written out unchanged. `gen_tikz.py` does the same for the TikZ figures, which are drawn by hand and not computed from data. For these files the scripts restore the submitted text and do not derive it from data.

### 3.0 `gen_figures.py`

Runs from `data/ttp_crt_data.xlsx` alone and writes the six data figures to `figs/`. The original RDI script that needed the master workbook was removed, since `gen_figures.py` reproduces the same values.

### 3.1 The five practice flags
The audit of Section 6.1 was built with these rules. They are recorded here as definitions, and their results are stored in the `Reporting_Audit` sheet.

1. **Formal model**: the study states a formal model.
2. **Standard benchmark**: the study's evidence class (see below) is the TTP1 categorized 60-instance benchmark, the full 9,720-instance library, or a benchmark shared with other papers of its sub-literature. A one-off subset, self-generated instances, or no experiment do not count. 39 studies.
3. **Statistical test**: a hypothesis test is named (Wilcoxon, Friedman, Kruskal-Wallis, t-test, and similar). Confidence intervals, average ranks, descriptive statistics and decision trees do not count. Purely theoretical studies are counted as no test. 35 studies.
4. **Code available**: the study releases source code. Repositories that hold only instances or solutions do not count. 26 studies.
5. **Compared with a current baseline**: the appraisal weaknesses do not flag a dated or quoted baseline pool. 61 studies.

### 3.2 The six evidence classes
Theoretical or no computational experiment; self-generated or synthetic instances; one-off subset of the official library; benchmark shared within its sub-literature; TTP1 categorized 60-instance benchmark; full 9,720-instance library.

### 3.3 Functions

| Function | Figure |
|---|---|
| `fig_coverage()` | 4 |
| `fig_rdi()` | 5 |
| `fig_sw()` | 6 |
| `fig_reporting()` | 7 (two heatmaps: by variant with the total, and by publication period) |
| `fig_bibliometrics()` (with `_draw_growth`, `_draw_venues`) | 8 |
| `fig_paradigm()` | 9 |

Design conventions in the script: one color per variant, one per algorithm family, fixed colors for publication status, appraisal polarity and benchmark category. All figures are saved at 300 dpi.

`requirements.txt` lists the packages the script imports (numpy, openpyxl, matplotlib, scipy). Tested with Python 3.13.

---

## 4. Figures (`figs/`)

| Paper figure | File | Produced by |
|---|---|---|
| 1 Anatomy of the TTP | `fig_anatomy.tex` | TikZ |
| 2 PRISMA flow | `fig_prisma.tex` | TikZ. Every count is exact, from a logged search of Scopus and Web of Science on 5 September 2026, combined with a supplementary Google Scholar pass for preprints neither database indexes. |
| 3 Taxonomy graph | `fig_taxonomy_dag.tex` | TikZ |
| 4 Algorithm-family coverage | `fig_coverage.png` | `fig_coverage()` |
| 5 RDI and Friedman/Nemenyi | `fig_rdi.png` | `fig_rdi()` |
| 6 Strengths and limitations | `fig_sw.png` | `fig_sw()` |
| 7 Reporting scorecard | `fig_reporting.png` | `fig_reporting()` |
| 8 Growth and venues | `fig_bibliometrics.png` | `fig_bibliometrics()` |
| 9 Families by period | `fig_paradigm.png` | `fig_paradigm()` |
| 10 Recommendations roadmap | `fig_agenda.tex` | TikZ |

The four TikZ figures compile with `cd figs && pdflatex tikz_figures.tex`. `tikz_figures.tex` is a wrapper that defines the packages, macros and color palette and inputs the four TikZ files. `tikz_figures.pdf` is its compiled output. Citations in captions are replaced by plain brackets so that no bibliography is needed.

A design-space grid figure (rows = variants, columns = the seven core taxonomy dimensions, unexplored combinations hatched on the grid) was tried between Figures 3 and 4 but was cut. It restated the same edges Figure 3's DAG already labels and the same gaps Table 11 already lists, so it added a second, harder-to-read place to find information the reader already had. The `Taxonomy` and `Unexplored` sheets below remain in the workbook; they now support only Tables 10 and 11.

---

## 5. Tables (`tables/`)

The tables were written by hand. `scripts/gen_tables.py` regenerates them (Table 6 and the rows of Table 15 from the data workbook, the others from stored text). Paper numbering at submission:

| Table | File | Content |
|---|---|---|
| 1 | `tab_review_scope.tex` | Scope of this review versus the prior survey |
| 2 | `tab_roadmap.tex` | Research questions, sections and headline findings |
| 3 | `tab_notation.tex` | Notation of the TTP1 model |
| 4 | `tab_coupling.tex` | The 84 method-proposing studies by tour--packing coupling strategy, against algorithm family |
| 5 | `tab_benchmarks.tex` | Benchmark instances per variant and whether each is available online |
| 6 | `tab_rdi.tex` | RDI results of the nine TTP1 methods, merged with each one's time budget, run count, and hardware/language |
| 7 | `tab_ttp2_thop.tex` | Reported evidence for TTP2 and ThOP methods |
| 8 | `tab_applications.tex` | Feature added by each variant and the motivation its authors state |
| 9 | `tab_code_links.tex` | Code and resource links given by the original papers, with a reachability check (Appendix A) |
| 10 | `tab_taxonomy.tex` | Consolidated taxonomy of the nine variants (Appendix B, landscape) |
| 11 | `tab_unexplored.tex` | Unexplored variant combinations (Appendix C, landscape) |
| 12 | `tab_variant_glance.tex` | The nine variants at a glance (Appendix C, landscape) |
| 13 | `tab_cat_ttp1.tex` | Catalog of TTP1 studies, grouped by sub-epoch (Appendix D, landscape long table) |
| 14 | `tab_cat_small.tex` | Catalog of PWT, MTTP/DTTP, and ThOP studies, grouped by variant, one page (Appendix D, landscape) |
| 15 | `tab_reporting.tex` | Full per-paper research-practice and experimental-reporting audit (Appendix E, landscape long table) |

A stand-alone Table 6 (`tab_budget.tex`, time budget/runs/hardware for the nine RDI-pool methods) was merged into the RDI table above. Both tables keyed on the same nine `Algorithm`/`Ref.` rows, so splitting the method's performance from its budget and hardware across two adjacent tables forced the reader to cross-reference two captions for one fact; merging removed a caption/label pair with no loss of information. `rdi_table()` now appends each method's `RDI_BUDGET` entry (budget, hardware/language, hand-maintained, since it is not in the data workbook) to the row it computes from `CatABC_Objectives`.

None of the five appendix `\section{...}\label{...}` headers or their intro paragraphs live in these table files. Two earlier arrangements were tried and abandoned (dedicated `app_*_intro.tex` files, then the header in the first table file of each appendix); the current, final arrangement puts all five headers directly in the paper's own `\appendix` block in `ttp_slr_tax_paper.tex`, one after another, with only Appendix A left in portrait and a single `\begin{landscape}...\end{landscape}` wrapping B through E as one continuous block (all five `\section` commands, the `\input` calls, and one `\newpage` between `tab_cat_ttp1.tex` and `tab_cat_small.tex`, all inside that one block). The table files stored here are therefore table content only, with no section header of their own, and are not meant to be `\input` in isolation without the paper's surrounding `\appendix` text.

The paper now has five appendices, all of them landscape material moved out of the main body: A (code links, Table 9, portrait), B (consolidated taxonomy, Table 10), C (unexplored combinations and variant summary, Tables 11-12), D (variant catalogs, Tables 13-14), E (reporting audit, Table 15) — B through E share one continuous landscape block. Tables 10-12 moved out of Section 3 (the taxonomy section); Tables 13-14 moved out of Section 4, the reviewer's own suggestion that the TTP1 catalog belongs in the main body was set aside for consistency with the rest. Section 3's and Section 4's prose still introduces the underlying ideas at their original spot (the nine-dimension list and the taxonomy DAG figure stay in Section 3; the per-variant literature narrative stays in Section 4) and points forward to the appendix tables by `\ref` rather than including them inline.

`tab_taxonomy.tex`, `tab_unexplored.tex`, `tab_variant_glance.tex`, and `tab_cat_small.tex` use `\begin{center}...\captionof{table}{...}...\end{center}` rather than a floating `\begin{table}`. This is a fix, not a style choice: with a floating table, LaTeX's placement algorithm was observed to move the table ahead of the `\section` heading and intro paragraph that precede it in the source (visibly wrong reading order — caught by reading the compiled PDF's actual text order, not just by checking that it compiled). `\captionof{table}` (from the `caption` package, already used by `tab_code_links.tex`) sidesteps float placement entirely, so the page renders in exact source order. `tab_cat_ttp1.tex` keeps `longtable`'s own `\caption` mechanism instead, which does not have this problem.

Table 13 (TTP1, 57 studies) is its own landscape longtable, wrapped internally in `\begin{landscape}`. Table 14 merges the PWT, MTTP/DTTP, and ThOP catalogs (16 rows total, three `\multicolumn`-headed groups, mirroring Table 13's era-grouping style) into one non-floating table — this was originally three separate small tables (one per variant, each on its own mostly-empty landscape page), merged into one so they land on a single rotated page instead of two or three. All three former labels (`tab:cat_pwt`, `tab:cat_mttpdttp`, `tab:cat_thop`) were replaced by one, `tab:cat_small`. Both catalogs share the same 5-column layout, `p{2.8cm}cp{3.0cm}p{4.9cm}p{7.9cm}` (Reference, Year, Algorithm, Instances, Key result), with no Code column: code release is already recorded per paper in Table 15, so it was dropped from both catalogs and the freed width went to the remaining four columns (originally 6 columns including Year and Code; now 5, with the p-columns each ~0.2-1.3cm wider).

Notes on how some were produced:
- **Table 15** lists, per study, the values of the reporting audit: a filled circle means the criterion is met, a blank cell means the paper does not state it. Its subtotals and grand total were checked against the audit data after the last change.
- **Table 9**: the URLs were recorded from the papers. Each URL was requested once on 21 September 2026 with a one-off HTTP check (status 200 means reachable). That check was not saved as a script. The SAAS-HC repository is not recorded from the paper, and the appendix text says so.
- **Table 6** merges the RDI comparison with each method's reported budget and hardware. The RDI, rank, and LOO-range values come from `rdi_table()` in `gen_tables.py`, an independent computation that uses the same RDI rule as `fig_rdi()` in `gen_figures.py`; the Budget and Hardware/language columns are the hand-maintained `RDI_BUDGET` dict, also in `gen_tables.py` (a former stand-alone table, folded in here since both shared the same nine `Algorithm`/`Ref.` rows). All nine ran under the same nominal 600-second/10-run budget, so the RDI ranking is not a budget artifact, but hardware and language differ (four different machines named, three papers state none), so a machine-driven effect cannot be ruled out from what the papers report.
- **Table 4** classifies each method-proposing study by how it handles the tour-packing coupling: sequential/decoupled, periodic/alternating, joint search, move-synchronized, or learned. The classification is a one-off coding pass over each paper's stated mechanism, documented in Section 5.2 of the paper, recorded in the `Coupling strategy` column of the `Corpus` sheet. Only CoCo, CoCo$^+$, and their PGCH predecessor evaluate every move against the joint objective; the similarly-named CS2SA/CS2SA$^\ast$/SAVI/CoSolver family alternates two separately-optimized local searches instead.
- **Table 13** gives the same per-paper record for TTP1 (57 studies) that Table 14 gives for the smaller variants: algorithm, instances, and headline result, one row per paper, grouped by the four sub-epochs of Section 4.1. Neither repeats code release, already in Table 15. It complements, not replaces, the appendix audit (Table 15), which records reporting practice rather than what each paper claims.
- R1 (Section 6.3) states its minimum-reporting items in prose, as hints for a TTP author rather than a formal checklist proposed for journals to adopt. R5 was shortened to remove its long descriptive summary of the six learning-based and one quantum study, which now lives in Section 5.1 where those families are classified, rather than duplicated in the recommendations.

---

## 6. Definitions and decisions worth remembering

- **Primary studies**: 93. The prior survey (Sarkar et al.) is the 94th work and is excluded from all counts.
- **Cross-listed studies**: two studies are tagged with two variants and are counted in both variant rows of Figures 5 and 8a: `bonyadi2013travelling` (TTP1, TTP2) and `herring2020dynamic` (TTP2, DTTP). Totals and period rows count each study once.
- **Standard benchmark**: 39 studies, defined from the evidence class (see 3.1 above); this is the paper's only count for this criterion.
- **Periods** used in the paper: 2013-2017 (26 studies), 2018-2021 (35), 2022-2026 (32). Two-year bins are used for Figure 9.
- **Family classification**: exact methods, approximation scheme, constructive/iterative heuristic, single-solution metaheuristic, population-based metaheuristic, hyper-heuristic, learning-based, quantum/hybrid quantum, plus "problem definition/benchmark/analysis" for the nine studies that propose no algorithm. Sizes: population-based 38, single-solution 16, hyper-heuristic 7, constructive 6, learning-based 6, exact 5, approximation 5, quantum 1, problem definition/analysis 9.
- **Corrections made to appraisal notes**: `nikfarjam2024use` (its repository holds instances and solutions but no source code); `pathirage2024chance` (classified as constructive/iterative heuristic, not learning-based, because its "surrogate" is a deterministic Chebyshev/Hoeffding bound); `huynh2023self` (code is public at `github.com/ELO-Lab/SAAS-HC`, and whether the paper cites it was not checked).

## 7. Known limitations

- The audit flags are rule-based on free text, so a wrong annotation gives a wrong flag.
- The author coded all appraisal notes alone, following the fixed schema of Section 2.2 of the paper.
- Only Tables 11 and 15 follow the data. The other tables are curated by hand, so a data change does not propagate to them.
- The master workbook is not shared, so the audit flags cannot be recomputed from this repository. The figures can be regenerated.
