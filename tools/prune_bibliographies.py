#!/usr/bin/env python3
"""
Quarto post-render step.

Every chapter is rendered with `nocite: '@*'` so that citeproc disambiguates
author-year labels (2018a/2018b, initials) against the full reference list,
which makes the in-text citations identical on every page and in the
References chapter. The side effect is that each page carries a hidden copy of
the whole bibliography. Quarto keeps that hidden `#refs` div so citation
hover pop-ups can be filled locally, but only the entries cited on the page
are needed for that. This script removes the rest.
"""
import os
import re
import sys
from pathlib import Path

out_dir = Path(os.environ.get("QUARTO_PROJECT_OUTPUT_DIR", "docs"))
ENTRY = re.compile(r'\s*<div id="(ref-[^"]+)" class="csl-entry"[^>]*>.*?</div>', re.S)

# Zotero (APA) joins several author-suppressed citations with a comma,
# "(2022, 2024)"; pandoc's citeproc uses a semicolon. Normalise to Zotero's form.
CITE_SPAN = re.compile(r'(<span class="citation" data-cites="[^"]+ [^"]+">)\((<a [^>]*>\d{4}[a-z]?</a>(?:; <a [^>]*>\d{4}[a-z]?</a>)+)\)(</span>)')
# Zotero collapses consecutive citations with the same rendered author part,
# "(Trochim, 1985, 1989)" or "(Zyphur et al., 2019, 2020)"; pandoc's citeproc
# only does so when the full author lists are identical. Merge them here.
SAME_AUTHOR = re.compile(r'(<a href="[^"]*" role="doc-biblioref">)([^<]*?), (\d{4}[a-z]?)</a>; <a href="([^"]*)" role="doc-biblioref">\2, (\d{4}[a-z]?)</a>')

# Citations where every item had its author suppressed ("Trochim (1985, 1989)"):
# pandoc's citeproc keeps the author on the second item, so strip it here.
SUPPRESSED = re.compile(r'<span class="cite-suppressed"><span class="citation"[^>]*>.*?</span></span>', re.S)


def fix_suppressed(m):
    s = re.sub(r'(<a href="[^"]*" role="doc-biblioref">)[^<]*?, (\d{4}[a-z]?)</a>', r"\1\2</a>", m.group(0))
    return s.replace("; ", ", ")


# Figure captions containing citations: Quarto copies the caption into the image's
# alt text and the lightbox title before citeproc has run, so those attributes
# still hold raw "[@ref...]" syntax. Replace them with the rendered caption text.
FIGURE = re.compile(r"<figure\b.*?</figure>", re.S)
FIGCAPTION = re.compile(r"<figcaption[^>]*>(.*?)</figcaption>", re.S)


def fix_figure_attrs(m):
    fig = m.group(0)
    if "[@ref" not in fig and "[-@ref" not in fig:
        return fig
    cap = FIGCAPTION.search(fig)
    if not cap:
        return fig
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", cap.group(1))).strip()
    text = text.replace("&", "&amp;").replace('"', "&quot;")
    fig = re.sub(r'(<a [^>]*class="lightbox"[^>]*title=")[^"]*(")', lambda a: a.group(1) + text + a.group(2), fig)
    fig = re.sub(r'(<img [^>]*alt=")[^"]*(")', lambda a: a.group(1) + text + a.group(2), fig)
    return fig


total_removed = 0
for page in sorted(out_dir.glob("*.html")):
    html = page.read_text(encoding="utf8")
    html, n_collapsed = SUPPRESSED.subn(fix_suppressed, html)
    html, n_fig = FIGURE.subn(fix_figure_attrs, html)
    n_collapsed += n_fig
    html, n = CITE_SPAN.subn(lambda m: m.group(1) + "(" + m.group(2).replace("; ", ", ") + ")" + m.group(3), html)
    n_collapsed += n
    while True:
        html, n = SAME_AUTHOR.subn(r'\1\2, \3</a>, <a href="\4" role="doc-biblioref">\5</a>', html)
        n_collapsed += n
        if not n:
            break
    if n_collapsed:
        page.write_text(html, encoding="utf8")
    if page.name == "references.html":
        continue
    m = re.search(r'<div id="refs"[^>]*>', html)
    if not m:
        continue
    # the refs div contains only csl-entry divs; find its extent by matching entries
    start = m.end()
    pos = start
    entries = []
    while True:
        em = ENTRY.match(html, pos)
        if not em:
            break
        entries.append(em)
        pos = em.end()
    end = pos
    cited = set(re.findall(r'href="(?:[^"#]*)#(ref-[^"]+)"', html))
    keep = [e.group(0) for e in entries if e.group(1) in cited]
    removed = len(entries) - len(keep)
    if removed:
        html = html[:start] + "\n" + "".join(keep) + html[end:]
        page.write_text(html, encoding="utf8")
        total_removed += removed
    print(f"{page.name}: kept {len(keep)} of {len(entries)} bibliography entries", file=sys.stderr)
print(f"pruned {total_removed} uncited hidden bibliography entries", file=sys.stderr)
