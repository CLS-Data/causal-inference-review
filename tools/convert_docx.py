#!/usr/bin/env python3
"""
Convert the Word manuscript of *Causal Inference in the CLS Cohort Studies*
into Quarto book sources.

Usage:
    python3 tools/convert_docx.py path/to/manuscript.docx [output_dir]

What it does
------------
1. Runs pandoc on the .docx (extracting images) to get pandoc markdown.
2. Drops the internal NOTES block and the Word table of contents.
3. Rebuilds figures (captions + anchors), tables (numbered captions), the
   terminology box (as a callout) and the numbered equation.
4. Splits the text into one page per chapter and one page per method
   (Chapter 5), keeping the manuscript's own section numbering
   (1, 2.1, 2.2.3, 5.10, 5.10.1 ...).
5. Turns "Section 5.10", "Figure 2.8b", "Table 5.2", "Box 2.1" and
   "Equation 5.1" mentions into links.
6. Extracts the Zotero citation metadata embedded in the Word field codes
   and writes it as CSL-JSON (references.json) so Quarto can render the
   reference list in APA style.

Files it writes (everything else in the repo is hand-maintained):
    01-*.qmd ... 08-*.qmd, 05-NN-*.qmd, references.qmd, references.json,
    images/imageN.png
"""
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# --------------------------------------------------------------------------
# configuration
# --------------------------------------------------------------------------
DOCX = Path(sys.argv[1]).expanduser()
OUT = Path(sys.argv[2]).expanduser() if len(sys.argv) > 2 else Path(__file__).resolve().parent.parent
IMG_DIR = OUT / "images"
TEXT_WIDTH_IN = 6.27           # Word text width; images are scaled relative to it
MIN_IMG_PCT, MAX_IMG_PCT = 26, 100

# Chapter files (level-1 headings in the manuscript, in order)
CHAPTER_FILES = {
    "Introduction": "01-introduction.qmd",
    "The Problem of Causal Inference with Observational Data": "02-problem.qmd",
    "Three Approaches to Causal Inference": "03-approaches.qmd",
    "Attributes of CLS's Cohort Studies That Support Causal Inference": "04-attributes.qmd",
    "The Methods of Causal Inference": "05-methods.qmd",
    "Discussion": "06-discussion.qmd",
    "Statements": "07-statements.qmd",
    "Appendix": "08-appendix.qmd",
}
# Numbered tables, in order of appearance in the manuscript
TABLE_NUMBERS = ["2.1", "5.1", "5.2", "5.3", "5.4", "5.5", "5.6"]


# --------------------------------------------------------------------------
# 1. pandoc + Zotero extraction
# --------------------------------------------------------------------------
def run_pandoc(docx: Path) -> str:
    tmp = Path(tempfile.mkdtemp())
    md = subprocess.run(
        ["pandoc", str(docx), "-t", "markdown", "--wrap=none",
         f"--extract-media={tmp}"],
        check=True, capture_output=True, text=True).stdout
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    for f in IMG_DIR.glob("image*.png"):
        f.unlink()
    for f in (tmp / "media").glob("*"):
        shutil.copy(f, IMG_DIR / f.name)
    shutil.rmtree(tmp)
    return md


def extract_zotero_items(docx: Path) -> list:
    """Collect the CSL-JSON item data that Zotero stores inside each citation field."""
    z = zipfile.ZipFile(docx)
    items = {}
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if part not in z.namelist():
            continue
        xml = z.read(part).decode("utf8")
        for m in re.finditer(r'<w:fldChar w:fldCharType="begin"/>(.*?)<w:fldChar w:fldCharType="end"/>',
                             xml, flags=re.S):
            instr = "".join(re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>", m.group(1), flags=re.S))
            if "ZOTERO_ITEM" not in instr:
                continue
            j = html.unescape(instr[instr.find("{"): instr.rfind("}") + 1])
            try:
                d = json.loads(j)
            except json.JSONDecodeError:
                continue
            for ci in d.get("citationItems", []):
                it = ci.get("itemData")
                if it:
                    items[it["id"]] = it
    out = []
    for it in items.values():
        it = dict(it)
        it["id"] = f"ref{it['id']}"
        for k in ("abstract", "note", "source", "archive", "archive_location", "call-number"):
            it.pop(k, None)
        out.append(it)
    out.sort(key=lambda d: json.dumps(d.get("author", d.get("editor", [{}]))[0]).lower())
    return out


def write_reference_page(refs_json: Path, csl: Path, out: Path):
    """Render the reference list (APA) with pandoc's citeproc into references.qmd.

    The citations in the text are plain text (Word/Zotero output), so the list is
    produced once here rather than by Quarto at render time. Quarto's book
    post-processing hides any div with id `refs`, so a different id is used.
    """
    src = "---\nnocite: '@*'\n---\n\n::: {#refs}\n:::\n"
    html_out = subprocess.run(
        ["pandoc", "-f", "markdown", "-t", "html", "--wrap=none", "--citeproc",
         "--bibliography", str(refs_json), "--csl", str(csl)],
        input=src, check=True, capture_output=True, text=True).stdout
    html_out = html_out.replace('id="refs"', 'id="bibliography"', 1)
    n = html_out.count('class="csl-entry"')
    out.write_text("# References {.unnumbered}\n\n```{=html}\n" + html_out.strip() + "\n```\n",
                   encoding="utf8")
    print(f"reference list: {n} entries")


# --------------------------------------------------------------------------
# 2. helpers
# --------------------------------------------------------------------------
def slug(text: str) -> str:
    """Pandoc's auto-identifier algorithm (approximately)."""
    t = re.sub(r"[*_`\[\]]", "", text)
    t = re.sub(r"[^\w\s.\-]", "", t, flags=re.UNICODE)
    t = re.sub(r"\s+", "-", t.strip()).lower()
    t = re.sub(r"^[^a-z]+", "", t)
    return t or "section"


def fix_math(s: str) -> str:
    s = s.replace(r"\bullet", r"\cdot")
    s = s.replace(r"for\ some\ individual\ i\ ", r"\text{ for some individual } i")
    # right-aligned bracketed remarks in the fixed-effects derivation
    s = re.sub(r"(?:\\ ){3,}\\lbrack", r"\\qquad \\lbrack", s)
    s = re.sub(r"\\lbrack(.*?)\\rbrack", r"\\left[\1\\right]", s)

    # move stray parentheses in/out of inline math so brackets balance
    def balance(m):
        body = m.group(1)
        opens, closes = body.count("("), body.count(")")
        if closes > opens and body.rstrip("\\ ").endswith(")"):
            body = body.rstrip("\\ ")[:-1]
            return f"${body}$)"
        return m.group(0)
    s = re.sub(r"\$([^$\n]+?)\$", balance, s)
    s = re.sub(r"\$([^$\n]*?\([^$\n)]*?)\$\)", r"$\1)$", s)   # '...(1, 1.3 R$)' -> inside
    return s


# --------------------------------------------------------------------------
# 3. block-level rebuilds
# --------------------------------------------------------------------------
def rebuild_html_figures(md: str) -> str:
    pat = re.compile(
        r'<figure>\s*<img src="([^"]+)" style="width:([\d.]+)in;[^"]*" alt="[^"]*" />\s*'
        r"<figcaption><p>(.*?)</p></figcaption>\s*</figure>", re.S)
    return pat.sub(lambda m: f'![{html.unescape(m.group(3))}]({m.group(1)}){{width="{m.group(2)}in"}}', md)


def merge_orphan_captions(md: str) -> str:
    pat = re.compile(r"!\[\]\(([^)]+)\)\{([^}]*)\}\n\n(Figure \d+\.\d+: [^\n]+)")
    return pat.sub(lambda m: f"![{m.group(3)}]({m.group(1)}){{{m.group(2)}}}", md)


FIG_RE = re.compile(r'!\[Figure (\d+)\.(\d+): (.*?)\]\([^)]*?(image\d+\.png)\)\{width="([\d.]+)in"[^}]*\}')


def rebuild_figures(md: str, registry: dict) -> str:
    def repl(m):
        ch, n, cap, img, w = m.groups()
        pct = int(round(float(w) / TEXT_WIDTH_IN * 100))
        pct = max(MIN_IMG_PCT, min(MAX_IMG_PCT, pct))
        cap = cap.rstrip()
        registry.setdefault("figures", {})[f"{ch}.{n}"] = f"figure-{ch}-{n}"
        # Quarto drops ids from figures that do not use its `fig-` prefix, so the
        # anchor goes on a wrapping div instead.
        return (f"::: {{#figure-{ch}-{n}}}\n\n"
                f'![**Figure {ch}.{n}:** {cap}](images/{img}){{width="{pct}%"}}\n\n:::')
    return FIG_RE.sub(repl, md)


TABLE_2_1_HTML = """::: {#table-2-1}

```{=html}
<table class="table table-bordered rct-table">
<caption><strong>Table 2.1:</strong> A hypothetical RCT of a job placement scheme that improves employment and wages, but wages are only observed among employed individuals (grey shaded cells) leading to a naïve analysis suggesting the intervention was associated with lower wages, when the causal effect was the opposite.</caption>
<thead>
<tr><th colspan="2" rowspan="2"></th><th colspan="2">Trial Arm</th><th rowspan="2"></th></tr>
<tr><th>Work Placement</th><th>Control</th></tr>
</thead>
<tbody>
<tr><th rowspan="2" class="text-end">Employability</th><th class="text-end">High</th><td class="shaded">£600 p.w.</td><td class="shaded">£550 p.w.</td><td></td></tr>
<tr><th class="text-end">Low</th><td class="shaded">£400 p.w.</td><td>£350 p.w. (Latent)</td><td></td></tr>
<tr class="summary"><th colspan="2" class="text-end">Observed Mean</th><td>£500 p.w.</td><td>£550 p.w.</td><td>− £50 p.w.</td></tr>
<tr class="summary"><th colspan="2" class="text-end">True Mean</th><td>£500 p.w.</td><td>£450 p.w.</td><td>+ £50 p.w.</td></tr>
</tbody>
</table>
```

:::"""


def is_table_line(line: str) -> bool:
    return line == "" or line[:1] in "+|" or line.startswith("  ")


def rebuild_tables_and_boxes(md: str, registry: dict) -> str:
    lines = md.split("\n")
    out_lines = list(lines)
    tables_seen = 0
    # find caption lines (": caption") scanning from the bottom so indices stay valid
    caption_idx = [i for i, l in enumerate(lines) if re.match(r"^\s*: ", l)]
    for ci in reversed(caption_idx):
        cap = lines[ci].strip()
        # locate the table block above the caption
        j = ci - 1
        while j >= 0 and is_table_line(lines[j]):
            j -= 1
        start = j + 1
        while lines[start] == "":
            start += 1
        block = lines[start:ci]
        while block and block[-1] == "":
            block.pop()
        if cap.startswith(": Box .:"):
            title = cap[len(": Box .:"):].strip().rstrip(".")
            content = []
            for l in block:
                if l.startswith("+"):
                    continue
                content.append(re.sub(r"\s*\|$", "", l[2:]) if l.startswith("| ") else l)
            while content and content[-1] == "":
                content.pop()
            registry.setdefault("boxes", {})["2.1"] = "box-2-1"
            new = ['::: {#box-2-1 .callout-note title="Box 2.1: ' + title + '" icon=false}', ""] + content + ["", ":::"]
        else:
            text = re.sub(r"^:\s*:\s*", "", cap)
            num = TABLE_NUMBERS[len(TABLE_NUMBERS) - 1 - tables_seen]
            tables_seen += 1
            anchor = "table-" + num.replace(".", "-")
            registry.setdefault("tables", {})[num] = anchor
            if text.startswith("A hypothetical RCT"):
                new = TABLE_2_1_HTML.split("\n")
            else:
                # display math inside cells -> inline math; pad so the grid columns stay aligned
                body = [re.sub(r"\$\$(.+?)\$\$", r"$\1$  ", l) for l in block]
                new = [f"::: {{#{anchor}}}", ""] + body + ["", f": **Table {num}:** {text}", "", ":::"]
        out_lines[start:ci + 1] = new
    return "\n".join(out_lines)


def rebuild_equation(md: str, registry: dict) -> str:
    pat = re.compile(r"^\s*-{10,}\n\s*(\$\$.*?)\$\$\s*\(Equation 5\.1\)\n\s*-{5,} -{5,}\n", re.M)
    registry.setdefault("equations", {})["5.1"] = "equation-5-1"
    return pat.sub(lambda m: f"\n::: {{#equation-5-1}}\n\n{m.group(1)} \\tag{{5.1}}$$\n\n:::\n", md)


def strip_subsection_toc(md: str) -> str:
    return re.sub(r"\*\*Subsections \(Links are clickable\)\*\*\n\n(?:\[[^\n]*\]\(#[^\n]*\)\n\n)+", "", md)


def note_callout(md: str) -> str:
    return re.sub(r"^\*Note: (.*?)\*$", r"::: {.callout-tip appearance=\"simple\" icon=false}\n**Note:** \1\n:::",
                  md, flags=re.M)


# --------------------------------------------------------------------------
# 4. split into files, add section numbers and ids
# --------------------------------------------------------------------------
def split_into_files(md: str, registry: dict):
    lines = md.split("\n")
    # footnotes: pull definitions out (with any indented continuation lines)
    footnotes, body = {}, []
    i = 0
    while i < len(lines):
        m = re.match(r"^\[\^(\d+)\]: (.*)$", lines[i])
        if m:
            key, text = m.group(1), [m.group(2)]
            i += 1
            while i < len(lines) and (lines[i] == "" or lines[i].startswith("    ")):
                text.append(lines[i]); i += 1
            while text and text[-1] == "":
                text.pop()
            footnotes[key] = "\n".join(text)
            continue
        body.append(lines[i]); i += 1

    files = {}          # filename -> list of lines
    order = []
    sections = registry.setdefault("sections", {})
    cur = None
    chap = 0; sub = 0; subsub = 0; sub4 = 0
    method_n = 0
    in_methods = False

    def add(fname, title_line):
        nonlocal cur
        cur = fname
        files[fname] = [title_line, ""]
        order.append(fname)

    for line in body:
        m = re.match(r"^(#{1,4}) (.*?)\s*$", line)
        if not m:
            if cur:
                files[cur].append(line)
            continue
        level, title = len(m.group(1)), m.group(2)
        if level == 1:
            if title == "References":
                cur = None
                continue
            chap += 1; sub = subsub = sub4 = 0
            fname = CHAPTER_FILES[title]
            in_methods = (title == "The Methods of Causal Inference")
            sid = slug(title)
            sections[str(chap)] = (fname, None)
            add(fname, f"# {chap} {title} {{#{sid}}}")
        elif level == 2:
            sub += 1; subsub = sub4 = 0
            num = f"{chap}.{sub}"
            if in_methods:
                method_n += 1
                fname = f"05-{method_n:02d}-{slug(title)}.qmd"
                sections[num] = (fname, None)
                add(fname, f"# {num} {title} {{#{slug(title)}}}")
            else:
                sections[num] = (cur, slug(title))
                files[cur].append(f"## {num} {title} {{#{slug(title)}}}")
        elif level == 3:
            subsub += 1; sub4 = 0
            num = f"{chap}.{sub}.{subsub}"
            sections[num] = (cur, slug(title))
            hashes = "##" if in_methods else "###"
            files[cur].append(f"{hashes} {num} {title} {{#{slug(title)}}}")
        elif level == 4:
            sub4 += 1
            num = f"{chap}.{sub}.{subsub}.{sub4}"
            sections[num] = (cur, slug(title))
            hashes = "###" if in_methods else "####"
            files[cur].append(f"{hashes} {num} {title} {{#{slug(title)}}}")

    # attach the footnotes each file uses
    for fname, flines in files.items():
        text = "\n".join(flines)
        used = sorted({k for k in re.findall(r"\[\^(\d+)\]", text) if k in footnotes}, key=int)
        if used:
            flines.append("")
            for k in used:
                flines.append(f"[^{k}]: {footnotes[k]}")
                flines.append("")
    return files, order


# --------------------------------------------------------------------------
# 5. cross-reference links
# --------------------------------------------------------------------------
NUM = r"\d+(?:\.\d+)*"
XREF = re.compile(
    rf"\b(Sections?|Figures?|Tables?|Box|Boxes|Equations?)\s+({NUM})([a-c](?:-[a-c])?)?"
    rf"((?:(?:,\s*(?:and\s+)?|\s+and\s+|\s*&\s*)(?:{NUM})(?:[a-c](?:-[a-c])?)?)*)")
ITEM = re.compile(rf"({NUM})([a-c](?:-[a-c])?)?")


def link_xrefs(files: dict, registry: dict):
    def target(kind, num, fname):
        kind = kind.lower()
        if kind.startswith("section"):
            hit = registry.get("sections", {}).get(num)
        elif kind.startswith("figure"):
            hit = registry.get("figures", {}).get(num)
            hit = (registry["fig_file"].get(num), hit) if hit else None
        elif kind.startswith("table"):
            hit = registry.get("tables", {}).get(num)
            hit = (registry["tbl_file"].get(num), hit) if hit else None
        elif kind.startswith("box"):
            hit = ("02-problem.qmd", "box-2-1") if num == "2.1" else None
        else:
            hit = registry.get("equations", {}).get(num)
            hit = (registry["eq_file"].get(num), hit) if hit else None
        if not hit or hit[0] is None:
            return None
        tfile, anchor = hit
        if tfile == fname:
            return f"#{anchor}" if anchor else None
        return f"{tfile}#{anchor}" if anchor else tfile

    unresolved = set()
    for fname, flines in files.items():
        in_simple = False
        for i, line in enumerate(flines):
            if re.match(r"^\s*-{5,}(\s+-{3,})*\s*$", line):
                in_simple = not in_simple
                continue
            if in_simple or not line.strip() or line[:1] in "#|+<:" or line.startswith("![") \
               or line.startswith("  -") or line.startswith("  :"):
                continue

            def repl(m):
                kind, first, suffix, rest = m.group(1), m.group(2), m.group(3) or "", m.group(4) or ""
                out = ""
                t = target(kind, first, fname)
                if t:
                    out += f"[{kind} {first}{suffix}]({t})"
                else:
                    unresolved.add(f"{kind} {first}"); out += f"{kind} {first}{suffix}"
                pos = 0
                for im in ITEM.finditer(rest):
                    out += rest[pos:im.start()]
                    n, sfx = im.group(1), im.group(2) or ""
                    t = target(kind, n, fname)
                    out += f"[{n}{sfx}]({t})" if t else n + sfx
                    if not t:
                        unresolved.add(f"{kind} {n}")
                    pos = im.end()
                out += rest[pos:]
                return out
            flines[i] = XREF.sub(repl, line)
    if unresolved:
        print("Unresolved cross-references:", sorted(unresolved))


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    md = run_pandoc(DOCX)
    refs = extract_zotero_items(DOCX)
    (OUT / "references.json").write_text(json.dumps(refs, indent=1, ensure_ascii=False), encoding="utf8")
    print(f"references: {len(refs)}")
    write_reference_page(OUT / "references.json", OUT / "apa.csl", OUT / "references.qmd")

    # keep everything from the Introduction heading onwards (drops NOTES + Word TOC)
    md = md[md.index("\n# Introduction\n") + 1:]
    registry = {}
    md = rebuild_html_figures(md)
    md = merge_orphan_captions(md)
    md = rebuild_figures(md, registry)
    md = rebuild_tables_and_boxes(md, registry)
    md = rebuild_equation(md, registry)
    md = strip_subsection_toc(md)
    md = note_callout(md)
    md = fix_math(md)
    md = md.replace("\u00a0", " ")

    files, order = split_into_files(md, registry)

    # where does each figure / table / equation live?
    registry["fig_file"], registry["tbl_file"], registry["eq_file"] = {}, {}, {}
    for fname, flines in files.items():
        text = "\n".join(flines)
        for num in re.findall(r"\{#figure-(\d+-\d+)", text):
            registry["fig_file"][num.replace("-", ".")] = fname
        for num in re.findall(r"\{#table-(\d+-\d+)", text):
            registry["tbl_file"][num.replace("-", ".")] = fname
        for num in re.findall(r"\{#equation-(\d+-\d+)", text):
            registry["eq_file"][num.replace("-", ".")] = fname

    # list of method pages for the Chapter 5 landing page
    methods = [f for f in order if re.match(r"05-\d\d-", f)]
    toc = ["", "**Sections**", ""]
    for f in methods:
        title = files[f][0]
        title = re.sub(r"^# (.*?) \{#.*\}$", r"\1", title)
        toc.append(f"- [{title}]({f})")
    toc.append("")
    files["05-methods.qmd"] += toc

    link_xrefs(files, registry)

    for f in OUT.glob("0[1-8]-*.qmd"):
        f.unlink()
    for fname in order:
        text = "\n".join(files[fname]).rstrip() + "\n"
        text = re.sub(r"\n{3,}", "\n\n", text)
        (OUT / fname).write_text(text, encoding="utf8")
        print(f"wrote {fname}")
    print("figures:", len(registry.get("figures", {})), "tables:", len(registry.get("tables", {})))


if __name__ == "__main__":
    main()
