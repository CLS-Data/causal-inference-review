# Causal Inference in the CLS Cohort Studies

Online version of the guide *Causal Inference in the CLS Cohort Studies*
(Wright, Tomova, Danka, Ploubidis & Silverwood, Centre for Longitudinal Studies, UCL).

**Website:** <https://cls-data.github.io/causal-inference-review/>

## How the site is built

The site is a [Quarto](https://quarto.org) book. The chapter sources (`*.qmd`) are
generated from the Word manuscript by `tools/convert_docx.py`; the rendered site is
committed in `docs/`, which GitHub Pages serves from the `main` branch.

To update the site after editing the manuscript:

```bash
# 1. regenerate the chapter files, images and reference list from the .docx
python3 tools/convert_docx.py "/path/to/manuscript.docx" .

# 2. render the book into docs/
quarto render

# 3. commit and push
git add -A && git commit -m "Update guide" && git push
```

Requirements: [Quarto](https://quarto.org/docs/get-started/) (1.4 or later),
[pandoc](https://pandoc.org) on the PATH, and Python 3.

### Files

| File | Purpose |
|---|---|
| `_quarto.yml` | Book structure, theme and rendering options |
| `index.qmd` | Landing page (abstract, how to cite, feedback) — hand-maintained |
| `01-*.qmd` … `08-*.qmd`, `05-NN-*.qmd` | Chapters and the twenty method pages — **generated**, do not edit by hand |
| `references.qmd` | Reference list page (rendered from `references.json`) |
| `references.json` | CSL-JSON bibliography extracted from the Zotero citation fields in the manuscript — generated |
| `apa.csl` | APA citation style used for the reference list |
| `custom.scss` | Site styling |
| `images/` | Figures extracted from the manuscript, cover and favicon |
| `tools/convert_docx.py` | Word → Quarto conversion script |
| `docs/` | Rendered site (published by GitHub Pages) |

### Notes on the conversion

- Section, figure, table and equation numbers follow the manuscript (e.g. 5.10, Figure 2.8),
  so the web and PDF/Word versions can be cited interchangeably. Chapter 5 is split into
  one page per method.
- "Section 5.10", "Figure 2.8b", "Table 5.2", "Box 2.1" and "Equation 5.1" mentions are
  turned into links automatically.
- Every Zotero citation in the Word file becomes a pandoc citation, and the reference list is
  rebuilt from the citation metadata embedded in the field codes. Quarto renders both in APA
  style; each in-text citation links to its entry and shows the full reference on hover.
  Because every page cites against the full list (`nocite`), author-year labels such as
  2018a/2018b are identical on every page; `tools/prune_bibliographies.py` (run automatically
  after rendering) trims the hidden per-page copies of the list to the entries each page cites.
