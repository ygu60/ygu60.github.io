#!/usr/bin/env python3
"""Build the print-ready PDF portfolio: assets/Jimmy_Gu_Portfolio.pdf

    page 1   cover: name, target roles, contact, narrative, and a visual index of the four projects
    pages 2-5  one US Letter page per project: title, outcome, objective, KPIs, contribution bullets,
             skill tags, results against the objective, 2-4 annotated figures, a constraints line, a
             "What I learned" line, and a footer with email and links to the full web page and the code

How it works
  1. Parses index.html and the four project pages (projects/<slug>.html) with html.parser and pulls
     the live content: titles, kickers, outcome paragraphs, fact grids, KPI tiles, "What I did"
     bullets, skill tags, constraint grids, and each figure's image, size and annotation-pin positions.
  2. Combines that with the print-only choices in PRINT below (which figures to show, their layout,
     crops, and shorter captions condensed from the web captions) and writes tools/print.html: a
     dedicated one-page-per-project print layout with its own CSS (colors and fonts are read from the
     :root block of ../styles.css).
  3. Copies print.html and the images it uses to a Windows temp dir (Chrome cannot read \\\\wsl$ paths),
     swaps Google's variable web fonts for static copies of the same fonts (Chrome embeds variable fonts
     in PDFs as Type 3 outlines; static fonts embed as normal fonts), checks the layout in headless
     Chrome (nothing may spill past a page's margins, every image and font must load), prints to PDF
     with headless Chrome, and copies the PDF into assets/.
  4. Sets the PDF's title/author metadata and bookmarks, verifies it (page count, size, fonts) and can
     render PNG previews (needs PyMuPDF).

Run it with a *Windows* Python (it drives Windows Chrome). From WSL:
    /mnt/c/Users/Owner/AppData/Local/Temp/jg_venv/Scripts/python.exe "$(wslpath -w ~/portfolio/tools/build_pdf.py)"
Needs network access for the Google web fonts. Standard library only, plus Pillow (image checks,
optional downscaling) and PyMuPDF (page count, previews) when they are installed.
"""
import argparse
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from html.parser import HTMLParser

TOOLS = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.dirname(TOOLS)
OUT_HTML = os.path.join(TOOLS, "print.html")
OUT_PDF = os.path.join(SITE, "assets", "Jimmy_Gu_Portfolio.pdf")
BASE_URL = "https://ygu60.github.io"
CHROME_DEFAULT = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
FONTS_URL = ("https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700"
             "&family=Instrument+Serif:ital@0;1&family=JetBrains+Mono:wght@400;500&display=swap")
# Google serves static (one file per weight) fonts to old browsers, variable fonts to new ones.
LEGACY_UA = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:30.0) Gecko/20100101 Firefox/30.0"
EXPECTED_PAGES = 5              # cover + 4 projects
MAX_PDF_BYTES = 5 * 1024 * 1024

# Page geometry in CSS px (96 px = 1 in). Content box = page minus padding.
PAGE_W, PAGE_H = 816, 1056
PAD_X, PAD_T, PAD_B = 38, 33, 29
CONTENT_W = PAGE_W - 2 * PAD_X
ROW_GAP = 9

# Contact details for the cover and footers (no phone number, by design).
CONTACT = [
    ("Email", "g.jimmy@wustl.edu", "mailto:g.jimmy@wustl.edu"),
    ("LinkedIn", "linkedin.com/in/jimmy-yufan-gu", "https://www.linkedin.com/in/jimmy-yufan-gu"),
    ("GitHub", "github.com/ygu60", "https://github.com/ygu60"),
    ("Website", "ygu60.github.io", BASE_URL),
]

# ---------------------------------------------------------------------------------------------
# PRINT: print-only editorial choices. Everything else comes from the HTML pages.
#   one_liner  cover-page outcome line, condensed from the page's outcome paragraph
#   outcome    optional shorter outcome paragraph for print (condensed from the page's); default: verbatim
#   kpis       indexes into the page's KPI tiles (4 fill the 2 x 2 grid)
#   rows       page body, top to bottom: dict(h=image height in px, or 0 to stretch, cells=[...]).
#              "fig:<id>" = a figure, its width set by h and its aspect ratio; "side" = role, contribution
#              bullets + tags (+ scorecard when score_at == "side"); "score" = the scorecard on its own;
#              "band" = the constraints box (otherwise printed full width under the rows); a list of
#              cells = a vertical stack as wide as its figure. h == 0 stretches the row's figures to equal
#              heights across the full width.
#   figs       per figure: pins to keep (web pin numbers; renumbered 1..n in print), an optional crop
#              (px trimmed from left, top, right, bottom; pins are re-mapped), a short caption and one
#              note per kept pin, both condensed from that figure's caption on the web page.
#              kind "svg" / "flow" reuse the page's inline diagram; "w" sets their width in px. Kind "sysprint"
#              is a landscape redraw of the ODMR page's system diagram (odmr_system_svg), label-checked
#              against the page's own SVG.
#   did        optional condensed "What I did" bullets for print; each must lead with a bold verb, and
#              every bold verb must be one of the page's own bullet verbs (default: the page's bullets)
#   objective  one or two sentences condensed from the page's Motivation section (printed under the outcome)
#   score      results against the objective, condensed from the page's Results section
#   lesson     one or two sentences condensed from the page's "What I learned" box (printed under the
#              constraints)
#   constraints  which constraint items (by label, in order) to show from the page's constraint grid;
#              "Team" and "Timeline" fall back to the page's fact grid when the constraint grid lacks them
#   tags       optional subset of the page's skill tags to print (default: all)
#   band_at    "side" puts the constraints box at the foot of the contribution column instead
# Web figure references like "(Figure 4)" are stripped from bullets, since print numbering differs.
# ---------------------------------------------------------------------------------------------
PRINT = {
    "quant-alpha": dict(
        one_liner="Led an out-of-sample study of trading strategies on 10 years of prices: "
                  "one survived, but it earned less than the S&amp;P&nbsp;500.",
        outcome="I led a study of trading ideas on 10 years of real stock prices. "
                "The survivor, a risk-controlled reversal strategy, earned a <strong>Sharpe ratio "
                "of 0.44</strong> (return per unit of risk) with a <strong>worst drop of 16.8%</strong>, half the "
                "S&amp;P&nbsp;500's, but returned less than the index.",
        kpis=[0, 1, 2, 3],
        rows=[dict(h=266, cells=["fig:risk-return-map.png", "side"]),
              dict(h=0, cells=["fig:regime-robustness.png", "fig:data-hygiene.png"])],
        score_at="side",
        figs={
            "risk-return-map.png": dict(
                pins=[1, 2, 3, 4],
                cap="Every main strategy, scored the same way: out of sample, after trading costs "
                    "(plotted from the memo's results tables). Gray points were rejected.",
                notes=["The raw reversal signal (buy recent losers) held up, but its 75.6% worst drop would "
                       "wipe out most of a small account.",
                       "A risk overlay cut the worst drop to 25.5%, but halved the Sharpe to 0.23.",
                       "A filter that sits out the calmest markets restored a 0.44 Sharpe at a 16.8% worst "
                       "drop: the final strategy.",
                       "Holding the S&amp;P&nbsp;500 (SPY) scored higher (0.85), with a deeper 33.7% drop."]),
            "regime-robustness.png": dict(
                pins=[1, 2, 3],
                cap="<b>Process:</b> the robustness check behind the final strategy. The two volatility-tertile "
                    "filters were each scored at two equally reasonable lookback windows.",
                notes=["“Skip the calmest third of days” held up at both (0.44, 0.42): kept.",
                       "“Trade only the most volatile third”: 0.32, then −0.06, a fit to one setting: rejected.",
                       "A Markov-switching model scored 0.13, below the no-gate 0.23."]),
            "data-hygiene.png": dict(
                pins=[],
                cap="<b>Data cleaning:</b> of 30 candidate stocks, 7 were dropped with the reason logged, not "
                    "silently filled in (drawn from the memo's data section).",
                notes=[]),
        },
        objective="Test whether a small account can capture a well-documented market pattern too small for "
                  "large funds, the way a quant researcher would: a literature hypothesis, a realistic backtest, "
                  "true out-of-sample validation, and honest reporting of what failed.",
        score=[("met", "Met", "Literature-based idea held up out of sample"),
               ("met", "Met, with limits", "Realistic backtest (flat costs, survivorship bias)"),
               ("part", "Partly", "Out-of-sample test: probabilistic Sharpe 0.91, short of 0.95"),
               ("met", "Met", "Honest reporting: every failed idea documented"),
               ("no", "Not met", "Beating the market (SPY Sharpe 0.85)")],
        constraints=["Team", "Timeline", "Capital", "Data", "Costs"],
        lesson="<b>One good number is not evidence.</b> A filter that scored 0.32 at one lookback flipped to −0.06 at "
               "another, so I now check results across settings.",
        did=["<strong>Set</strong> the goal: a strategy a small account could run, tested the way a quant "
             "researcher would before risking money.",
             "<strong>Directed</strong> an AI coding agent (Claude Code) to build the pipeline: cleaning, a "
             "realistic backtest, walk-forward tests, a risk overlay.",
             "<strong>Directed</strong> four follow-up studies, from volatility filters to ML signal searches; a "
             "library survey caught a wrong Sortino formula.",
             "<strong>Set</strong> the bar the ML search had to clear, before it ran.",
             "<strong>Signed off</strong> on each stage through three pull requests, backed by 29 unit tests."],
        tags=["Python", "pandas", "scikit-learn", "PyTorch", "statsmodels", "Walk-forward validation",
              "Permutation testing", "Unit testing", "Data cleaning"],
    ),
    "donor-dashboard": dict(
        one_liner="Analyzed a crime victim nonprofit's donor records, then built a Streamlit dashboard that "
                  "repeats the analysis on each new GiveButter export.",
        outcome="I analyzed Crime Victim Center (CVC) donor records, found that <strong>a small group of donors gave "
                "most of the dollars</strong>, and built a <strong>five-page Streamlit dashboard</strong> that repeats "
                "that check, plus retention, map and campaign views, on each new GiveButter export. It <strong>fully "
                "covers the brief's donor-ranking goal</strong> and partly covers segmentation, tracking and outreach.",
        kpis=[0, 1, 2, 3],
        rows=[dict(h=231, cells=["fig:home-upload.png", "side"]),
              dict(h=0, cells=["fig:pareto.png", "fig:cohorts.png"])],
        score_at="side",
        figs={
            "home-upload.png": dict(
                pins=[1, 2, 3, 4, 5],
                cap="The Home page after an upload. All screenshots run the repo's code on <b>synthetic demo "
                    "data</b> (1,200 fictional donors), not CVC's, to protect privacy.",
                notes=["Staff can drop in GiveButter exports, several files at once.",
                       "Every merged file is listed, so staff can see what the numbers include.",
                       "Links to the four analysis pages.",
                       "The cleaned table follows the user to every page for the session.",
                       "Drag across the chart to pick a date range; the app labels the growth over it."]),
            "pareto.png": dict(
                pins=[2, 3, 4], crop=(0, 372, 0, 0),
                cap="Top-donor (80/20) view, below a plain-language explainer for staff (synthetic demo data).",
                notes=["Slider: the share of dollars to explain (default 80%).",
                       "Each donor's total giving in the loaded files, largest first.",
                       "Running share of dollars: 231 of 1,200 demo donors (19%) give 80%."]),
            "cohorts.png": dict(
                pins=[1, 2, 4], crop=(0, 222, 0, 40),
                cap="Which donors come back (synthetic demo data). Rows: first-gift quarter. Color: the share of "
                    "that cohort giving again in each later quarter.",
                notes=["Quarter 0 is 100% by definition.",
                       "Read a row left to right to see how a cohort holds up.",
                       "Newer cohorts have had fewer quarters to return: a staircase."]),
        },
        objective="CVC's brief asked for a donor strategy aimed at high-value and monthly donors. Five of its "
                  "goals depend on donor data; I rate each under “Against the objective.”",
        score=[("met", "Met", "Find high-value donors"),
               ("part", "Partly", "Segment donor groups"),
               ("part", "Partly", "Track donors on GiveButter"),
               ("part", "Informs only", "Recurring-gift campaign"),
               ("part", "Partly", "Guide outreach and acquisition"),
               ("no", "Not measured", "Adoption and effect on fundraising")],
        constraints=["Team", "Timeline", "Data", "Input"],
        lesson="<b>Dashboards fail quietly.</b> A 2026 re-test found four bugs that raise no error (such as a "
               "+51,062.7% growth label); I would now ship tests on known data.",
        did=["<strong>Analyzed</strong> CVC's records first: 2022 to 2024 campaign totals and a pandas Pareto "
             "(80/20) analysis of donors.",
             "<strong>Built</strong> the upload-and-clean step that merges any number of GiveButter exports into "
             "one table.",
             "<strong>Designed</strong> the retention metrics: returning vs. one-time donors, quarterly churn, "
             "cohort heatmaps.",
             "<strong>Built</strong> an adjustable top-donor (80/20) view, <strong>geocoded</strong> donor ZIP "
             "codes onto a map, and compared campaigns year over year.",
             "<strong>Wrote</strong> plain-language explainers beside the charts for CVC staff."],
        tags=["Python", "pandas", "Streamlit", "Altair", "Plotly", "Data cleaning", "Cohort analysis",
              "Pareto analysis", "Geocoding"],
    ),
    "odmr-pipeline": dict(
        one_liner="Built the Python pipeline behind a diamond magnetic-field sensor: it drives the instruments, "
                  "cleans the sweeps and fits every spectrum.",
        outcome="I built the Python pipeline behind a diamond-based "
                "magnetic-field sensor. It drives the instruments, captures <strong>10,000 sweeps</strong>, rejects "
                "glitched ones and fits every spectrum: up to <strong>8 resonance lines</strong>, <strong>R² ≥ "
                "0.97</strong> in all 10 conditions. Readings ran below the coil's predicted field; one constant "
                "0.61&nbsp;mT offset, fitted afterward and consistent with stray magnetization, brings 0.4 to "
                "0.8&nbsp;A within about 1.1σ of theory.",
        kpis=[0, 1, 2, 3],
        rows=[dict(h=277, cells=["fig:hero-fits.png", "side"]),
              dict(h=222, cells=[["fig:system", "band"], "fig:kept-vs-rejected.png"])],
        score_at="side",
        figs={
            "hero-fits.png": dict(
                pins=[1, 2, 3],
                cap="Real spectra from my pipeline (dark gray) and my adaptive fit (blue) at every other coil "
                    "current: one resonance splits into many.",
                notes=["Zero field: a clean double dip, 7.48% deep.",
                       "0.4&nbsp;A: shallow shoulders a shared-depth fit missed; per-line depths recover them.",
                       "0.8&nbsp;A: eight lines; the outermost pair (orange) gives the field."]),
            "system": dict(
                kind="sysprint",
                cap="<b>The system.</b> Gray: optics and instruments we assembled as a team. Blue: the "
                    "software I built (7 Python scripts, about 2,300 lines).",
                notes=[]),
            "kept-vs-rejected.png": dict(
                pins=[1, 2],
                cap="<b>Process:</b> automated QC on the real 0&nbsp;A capture.",
                notes=["Rejected sweeps carry single-sample spikes to ±1&nbsp;V.",
                       "Kept sweeps hint at the dip; the threshold (40.2&nbsp;mV) comes from the capture's own "
                       "jump statistics."]),
        },
        objective="Build a magnetic-field sensor from teaching-lab hardware and test it against the coil's "
                  "predicted field. Defects in diamond (NV centers) glow red under a green laser and dim by a "
                  "few percent at a microwave resonance near 2.9&nbsp;GHz; a magnetic field splits that "
                  "resonance, and the split gives the field.",
        score=[("met", "Met", "Clean without distorting: the data-derived threshold reproduced my hand-tuned split"),
               ("met", "Met", "Detect the resonance: a 7.48% dip, close to the roughly 8% in the literature"),
               ("met", "Met", "Resolve the splitting: up to 8 lines, every fit R² ≥ 0.97"),
               ("part", "Partly", "Match coil theory: readings ran low; one 0.61&nbsp;mT offset fits the gap, "
                                  "and 0.3&nbsp;A stays a 3.1σ outlier")],
        constraints=["Team", "Timeline", "Context", "Starting point", "Data volume"],
        lesson="<b>Build tools that make problems visible.</b> QC plots and a read-only instrument check found most fixes; "
               "full coil geometry cut the gap from about 30σ to 3 to 13σ.",
        did=["<strong>Automated</strong> an oscilloscope and a function generator from Python (PyVISA/SCPI): "
             "1,000 sweeps per condition.",
             "<strong>Engineered</strong> QC that rejects glitched sweeps with a data-derived threshold (85% kept).",
             "<strong>Diagnosed</strong> a trigger that never locked and fixed its root cause; <strong>fixed</strong> a "
             "jagged frequency axis and a sign bug.",
             "<strong>Designed</strong> an adaptive fit that counts dips first; per-line depths lifted R² at "
             "0.5&nbsp;A from 0.88 to 0.99.",
             "<strong>Propagated</strong> uncertainty into every field reading and the theory band.",
             "<strong>Paired</strong> with an AI assistant (Claude Code) on the code and docs; <strong>documented</strong> "
             "it all: README, uncertainty write-up, lab notebook, paper draft."],
        tags=["Python", "PyVISA / SCPI", "NumPy", "SciPy", "matplotlib", "Signal processing", "Data QC",
              "Nonlinear fitting", "Uncertainty analysis"],
    ),
    "mosaic-aesthetics": dict(
        one_liner="Wrote the MATLAB feature extractor for a WashU mosaic-art study; it reproduces a published "
                  "study's average color and brightness measures.",
        outcome="I wrote the MATLAB feature extractor for a WashU study of why people like mosaics: "
                "<strong>982 mosaics</strong> into <strong>172 features</strong> each. Benchmarked against a published "
                "study, it reproduces the original's average color and brightness measures (9 of 18 global features at "
                "|r| ≥ 0.7), not yet its hue-count, blur or segment-shape features. On 70 pilot images, the team "
                "ranked features against ratings with my table.",
        kpis=[0, 1, 2, 3],
        rows=[dict(h=274, cells=["fig:feature-anatomy.jpg", "side"]),
              dict(h=0, cells=["fig:validation.png", "fig:pilot-results.png"])],
        score_at="side",
        figs={
            "feature-anatomy.jpg": dict(
                pins=[1, 2, 3],
                cap="What my extractor computes, for one image and one channel. Values are this image's row "
                    "in my feature table.",
                notes=["Input: Pietro Cavallini, <em>Annunciation</em> (detail), 1296 to 1300, public domain "
                       "(photo: Web Gallery of Art), resized to 256 × 256.",
                       "Six of the 18 statistics, drawn on the lightness histogram.",
                       "Repeated on 9 color channels, plus frequency and color-harmony features."]),
            "validation.png": dict(
                pins=[],
                cap="<b>Benchmark:</b> I ran the published study's 826 paintings through my extractor and, for "
                    "each of its 83 features, kept the best |r| with any of mine. Hatched: region features I "
                    "later set aside. Hue-model and hatched matches re-run the original study's own code, so they "
                    "check my port rather than add new measurements.",
                notes=[]),
            "pilot-results.png": dict(
                pins=[1, 2],
                cap="<b>Results:</b> preliminary pilot (70 mosaics, 5 ratings each), as reported on our poster.",
                notes=["All correlations are modest (|r| 0.33 to 0.35); with 172 features and 70 images, this "
                       "could be chance.",
                       "LASSO ranks a different set of simple statistics."]),
        },
        objective="A 2021 study predicted people's ratings of paintings from simple and complex visual features. "
                  "Does that hold for mosaics? My piece: turn every image into numbers a regression can use, "
                  "benchmarked against the published baseline.",
        score=[("met", "Reproduced", "Average color and brightness: 9 of 18 global features at |r| ≥ 0.7"),
               ("no", "Not yet", "The original's hue-count, blur and segment-shape features"),
               ("part", "Preliminary", "Pilot (70 images) ranks candidate features; no model validated yet")],
        constraints=["Team", "Ratings cost", "Tooling"],
        lesson="<b>Validation before modeling.</b> Rebuilding legacy code, matching it feature by feature and "
               "chasing randomness is how I know which parts of the table to trust.",
        did=["<strong>Revived</strong> the original study's MATLAB code on Windows: rebuilt its C++ binaries and "
             "checked outputs against stored values.",
             "<strong>Engineered</strong> a new extractor: 18 statistics on 9 color channels plus FFT spectral "
             "centroids, reusing the original's hue-harmony models.",
             "<strong>Benchmarked</strong> it against the original 83 features: 23 matched on my first pass, 30 with "
             "its hue and region code ported.",
             "<strong>Diagnosed</strong> randomness in the hue-model choice (2 of 100 repeat runs changed) and in "
             "graph-cut segments.",
             "<strong>Evaluated</strong> graph cuts and SAM2 for tile-level features; prototyped HOG features for "
             "tile flow.",
             "<strong>Delivered</strong> the feature table for our pilot analysis; co-authored the symposium poster "
             "and abstract."],
    ),
}
ORDER = ["quant-alpha", "donor-dashboard", "odmr-pipeline", "mosaic-aesthetics"]


# ---------------------------------------------------------------------------------------------
# A small DOM on top of html.parser
# ---------------------------------------------------------------------------------------------
VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}


class Node:
    def __init__(self, tag, attrs=(), parent=None):
        self.tag, self.attrs, self.parent, self.children = tag, dict(attrs), parent, []

    @property
    def classes(self):
        return (self.attrs.get("class") or "").split()

    def walk(self):
        for c in self.children:
            if isinstance(c, Node):
                yield c
                yield from c.walk()

    def find_all(self, tag=None, cls=None, pred=None):
        return [n for n in self.walk() if (tag is None or n.tag == tag) and (cls is None or cls in n.classes)
                and (pred is None or pred(n))]

    def find(self, tag=None, cls=None, pred=None):
        r = self.find_all(tag, cls, pred)
        return r[0] if r else None

    def kids(self, tag=None, cls=None):
        return [c for c in self.children if isinstance(c, Node) and (tag is None or c.tag == tag)
                and (cls is None or cls in c.classes)]

    def closest(self, cls=None, tag=None):
        n = self.parent
        while n is not None and not ((cls is None or cls in n.classes) and (tag is None or n.tag == tag)):
            n = n.parent
        return n

    def text(self):
        out = [c if isinstance(c, str) else c.text() for c in self.children]
        return re.sub(r"\s+", " ", "".join(out)).strip()

    def inner(self):
        """Inner HTML, re-escaped, with whitespace runs collapsed (fine for inline content)."""
        out = []
        for c in self.children:
            if isinstance(c, str):
                out.append(html.escape(re.sub(r"\s+", " ", c), quote=False))
            else:
                a = "".join(f" {k}" if v is None else f' {k}="{html.escape(v)}"' for k, v in c.attrs.items())
                out.append(f"<{c.tag}{a}>" + ("" if c.tag in VOID else c.inner() + f"</{c.tag}>"))
        return "".join(out).strip()


class _Builder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = self.cur = Node("#root")

    def handle_starttag(self, tag, attrs):
        n = Node(tag, attrs, self.cur)
        self.cur.children.append(n)
        if tag not in VOID:
            self.cur = n

    def handle_startendtag(self, tag, attrs):
        self.cur.children.append(Node(tag, attrs, self.cur))

    def handle_endtag(self, tag):
        n = self.cur
        while n is not self.root and n.tag != tag:
            n = n.parent
        if n is not self.root:
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.children.append(data)


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def parse(path):
    b = _Builder()
    b.feed(read(path))
    return b.root


def die(msg):
    sys.exit(f"build_pdf: ERROR: {msg}")


# ---------------------------------------------------------------------------------------------
# Content extraction
# ---------------------------------------------------------------------------------------------
FIG_REF = [re.compile(r"\s*\((?:see )?Figures? \d+(?:\s*(?:and|to|,)\s*\d+)*\)"),  # " (Figure 4)"
           re.compile(r",\s*Figures? \d+(?:\s*(?:and|to)\s*\d+)*(?=\))")]            # ", Figure 6)"


def strip_fig_refs(s):
    for rx in FIG_REF:
        s = rx.sub("", s)
    return s


def pin_xy(style):
    m = re.search(r"left:\s*([\d.]+)%;\s*top:\s*([\d.]+)%", style or "")
    if not m:
        die(f"cannot read pin position from style {style!r}")
    return float(m.group(1)), float(m.group(2))


def extract_project(slug):
    path = os.path.join(SITE, "projects", slug + ".html")
    src = read(path)
    root = parse(path)
    hero = root.find("header", "pp-hero")
    if hero is None:
        die(f"{slug}: no header.pp-hero")
    p = dict(slug=slug)
    p["kicker"] = hero.find("p", "kicker").text()
    p["title"] = hero.find("h1").text()
    p["outcome"] = hero.find("p", "pp-outcome").inner()
    p["facts"] = [(d.find("dt").text(), d.find("dd").inner()) for d in hero.find("dl", "pp-facts").kids("div")]
    code = [a.attrs["href"] for a in hero.find_all("a") if "github.com/ygu60/" in a.attrs.get("href", "")]
    p["code"] = code[0] if code else None
    k = root.find("div", "kpis")
    p["kpis"] = [(d.find("strong").inner(), d.find("span").inner()) for d in k.kids("div")] if k else []

    # "What I did": the section whose h2 carries <small>My contribution</small>. Bullets are the first
    # non-tag list in its text column (lists inside figure captions don't count).
    sec = root.find("section", "pp-section", pred=lambda n: n.find("small") is not None
                    and "contribution" in n.find("small").text().lower())
    if sec is None:
        die(f"{slug}: no 'My contribution' section")
    txt = sec.find("div", "pp-text")
    if txt is None:
        die(f"{slug}: no text column in the contribution section")
    lists = [u for u in txt.find_all("ul") if "tags" not in u.classes and u.closest(tag="figure") is None]
    if not lists:
        die(f"{slug}: no contribution bullets")
    p["did"] = [strip_fig_refs(li.inner()) for li in lists[0].kids("li")]
    for b in p["did"]:
        if re.search(r"\bFigures?\s+\d", b):
            die(f"{slug}: a web figure reference survived in a bullet: {b[:80]}")
    tags = txt.find("ul", "tags")
    p["tags"] = [li.text() for li in tags.kids("li")] if tags else []
    cons = root.find("dl", "constraints")
    p["constraints"] = [(d.find("dt").text(), d.find("dd").inner()) for d in cons.kids("div")] if cons else []
    les = root.find("div", "lesson")
    p["lesson"] = les.text() if les else ""

    # Figures, keyed by image file name: size, alt text, annotation pins and the caption's pin notes.
    figs = {}
    for fig in root.find_all("figure"):
        for img in fig.find_all("img"):
            name = img.attrs["src"].rsplit("/", 1)[-1]
            annot = img.closest("annot")
            pins = []
            if annot is not None:
                for s in annot.kids("span", "pin"):
                    if "m" in s.classes:          # phone-only pin set (for a stacked phone image)
                        continue
                    pins.append((int(s.text()), *pin_xy(s.attrs.get("style"))))
            figs[name] = dict(src=img.attrs["src"], w=int(img.attrs["width"]), h=int(img.attrs["height"]),
                              alt=img.attrs.get("alt", ""), pins=pins)
    p["figs"] = figs

    # Inline diagrams reused as-is: the ODMR system diagram (desktop SVG) and the mosaic study-design flow.
    m = re.search(r'<svg class="sysdiag sysdiag-d".*?</svg>', src, re.S)
    p["svg"] = m.group(0) if m else None
    flow = root.find("ol", "flow")
    p["flow"] = strip_fig_refs(flow.inner()) if flow else None
    if p["flow"] and re.search(r"\bFigures?\s+\d", p["flow"]):
        die(f"{slug}: a web figure reference survived in the study-design flow")
    return p


def extract_index():
    root = parse(os.path.join(SITE, "index.html"))
    d = dict(pill=root.find("p", "pill").text(), tagline=root.find("p", "tagline").inner(),
             lead=root.find("p", "lead").inner(),
             glance=[(g.find("dt").text(), g.find("dd").inner()) for g in root.find("dl", "glance").kids("div")])
    cards = {}
    for c in root.find_all("article", "card"):
        slug = c.find("a", "card-link").attrs["href"].rsplit("/", 1)[-1].replace(".html", "")
        img = c.find("img")
        cards[slug] = dict(
            kicker=c.find("p", "kicker").text(),
            stats=[(x.find("dt").inner(), x.find("dd").inner()) for x in c.find("dl", "card-stats").kids("div")],
            thumb=img.attrs["src"], w=int(img.attrs["width"]), h=int(img.attrs["height"]),
            alt=img.attrs.get("alt", ""), thumbcap=c.find("figcaption").inner())
    d["cards"] = cards
    return d


def root_tokens():
    m = re.search(r":root\s*\{(.*?)\}", read(os.path.join(SITE, "styles.css")), re.S)
    if not m:
        die("no :root block in styles.css")
    return m.group(1).strip()


# ---------------------------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------------------------
def odmr_system_svg():
    """Landscape print version of the ODMR page's system diagram (the web one is portrait, 640 x 790, and
    prints too small to read in a half-width slot). Same boxes, labels, colors and connections, arranged in
    a 4-column grid: the hardware converges on the oscilloscope (top left), which feeds the software lane,
    read left to right and then back (a U-shaped flow). Every label is checked against the page's SVG."""
    W, BW, GAP, X0 = 772, 170, 24, 10
    X = [X0 + i * (BW + GAP) for i in range(4)]
    A, B, BH = 30, 108, 58                      # hardware rows A and B
    C, D, DH = 218, 300, 70                     # software rows C and D
    H = D + DH + 14
    out = [f'<svg class="sysprint" viewBox="0 0 {W} {H}" role="img" aria-label="System diagram: optics and '
           f'instruments set up by a team of two feed the oscilloscope; six software stages I built acquire, clean, '
           f'calibrate, average, fit and turn each spectrum into a field value with uncertainty.">',
           '<defs>']
    for mid, col in (("pah", "#8a877d"), ("pahg", "#1f9d55"), ("pahr", "#e34948"), ("pahb", "#3d5afe")):
        out.append(f'<marker id="{mid}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6.5" markerHeight="6.5" '
                   f'orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{col}"/></marker>')
    out.append('</defs>')
    # lanes
    out.append(f'<rect x="2" y="2" width="{W - 4}" height="{B + BH + 12 - 2}" rx="12" fill="#f4f3ef"/>')
    out.append(f'<text class="lane" x="{X0 + 2}" y="20">HARDWARE · SET UP AS A TEAM OF 2</text>')
    out.append(f'<rect x="2" y="{B + BH + 20}" width="{W - 4}" height="{H - (B + BH + 20) - 2}" rx="12" fill="#eef0ff"/>')
    out.append(f'<text class="lane" x="{W - X0 - 2}" y="{B + BH + 38}" text-anchor="end">SOFTWARE · BUILT BY ME</text>')

    def box(x, y, w, h, title, subs, sw=False, dashed=False):
        stroke = "#3d5afe" if sw else "#cfcabd"
        dash = ' stroke-dasharray="6 5"' if dashed else ""
        r = [f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="9" fill="#fff" stroke="{stroke}" '
             f'stroke-width="1.5"{dash}/>',
             f'<text class="t" x="{x + 12}" y="{y + 23}">{title}</text>']
        for sy, sub in subs:
            r.append(f'<text class="s" x="{x + 12}" y="{y + sy}">{sub}</text>')
        return f'<g class="{"sw" if sw else "hw"}">' + "".join(r) + "</g>"

    def arrow(x1, y1, x2, y2, col="pah", width=2):
        color = {"pah": "#8a877d", "pahg": "#1f9d55", "pahr": "#e34948", "pahb": "#3d5afe"}[col]
        return (f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{width}" '
                f'marker-end="url(#{col})"/>')

    ya, yb = A + BH // 2, B + BH // 2           # arrow heights for rows A and B
    # hardware: oscilloscope spans rows A-B in column 1; the signal chain runs right to left into it
    out.append(box(X[0], A, BW, B + BH - A, "Oscilloscope", [(42, "CH1: fluorescence"),
                                                             (yb - A + 13, "CH2: sweep voltage")]))
    out.append(box(X[1], A, BW, BH, "Photodiode", [(42, "red glow = signal")]))
    out.append(box(X[2], A, BW, BH, "Diamond", [(42, "NV centers, coil field")]))
    out.append(box(X[3], A, BW, BH, "Green laser", [(42, "532 nm excitation")]))
    out.append(box(X[1], B, BW, BH, "Function gen.", [(42, "sawtooth sweep")]))
    out.append(box(X[2], B, BW, BH, "Microwaves", [(42, "VCO, near 2.9 GHz")]))
    out.append(arrow(X[3] - 2, ya, X[2] + BW + 3, ya, "pahg", 3))           # laser -> diamond
    out.append(arrow(X[2] - 2, ya, X[1] + BW + 3, ya, "pahr", 3))           # diamond -> photodiode
    out.append(arrow(X[1] - 2, ya, X[0] + BW + 3, ya, "pahr", 2))           # photodiode -> scope CH1
    out.append(arrow(X[1] - 2, yb, X[0] + BW + 3, yb))                      # function gen -> scope CH2
    out.append(arrow(X[1] + BW + 2, yb, X[2] - 3, yb))                      # function gen -> microwaves
    mx = X[2] + BW // 2
    out.append(arrow(mx, B - 2, mx, A + BH + 3))                            # microwaves -> diamond
    # software: row C left to right, then down and back along row D
    sx = X[0] + BW // 2
    out.append(arrow(sx, B + BH + 2, sx, C - 3, "pahb", 2.5))               # scope -> acquire
    out.append(f'<text class="s" x="{sx + 10}" y="{B + BH + 38}">PyVISA + SCPI</text>')
    for i, (t, sub) in enumerate([("Acquire", "1,000 sweeps/current"), ("Clean", "drop glitches, despike"),
                                  ("Calibrate", "voltage → GHz"), ("Average", "common grid, % dip")]):
        out.append(box(X[i], C, BW, BH, t, [(42, sub)], sw=True))
        if i:
            out.append(arrow(X[i - 1] + BW + 2, C + BH // 2, X[i] - 3, C + BH // 2, "pahb"))
    out.append(arrow(X[3] + BW // 2, C + BH + 2, X[3] + BW // 2, D - 3, "pahb"))      # average -> fit
    out.append(box(X[3], D, BW, DH, "Fit", [(42, "adaptive N-Lorentzian")], sw=True))
    out.append(box(X[2], D, BW, DH, "Field ± σ", [(42, "vs. coil theory")], sw=True))
    out.append(box(X[0], D, X[1] + BW - X[0], DH, "Reports &amp; figures",
                   [(42, "grid · ridgeline · 3D terrain"), (59, "fit table · field-vs-current plot")],
                   sw=True, dashed=True))
    out.append(arrow(X[3] - 2, D + DH // 2, X[2] + BW + 3, D + DH // 2, "pahb"))      # fit -> field
    out.append(arrow(X[2] - 2, D + DH // 2, X[1] + BW + 3, D + DH // 2, "pahb"))      # field -> reports
    out.append("</svg>")
    return "\n".join(out)


def check_svg_labels(page_svg, print_svg, slug):
    """Every text label in the print diagram must appear in the web page's diagram."""
    def texts(svg):
        return [html.unescape(re.sub(r"\s+", " ", t)).strip() for t in re.findall(r"<text[^>]*>(.*?)</text>", svg, re.S)]
    web = " | ".join(texts(page_svg))
    missing = [t for t in texts(print_svg) if t not in web]
    if missing:
        die(f"{slug}: print diagram labels not in the page's diagram: {missing}")

def site_rel(src):
    """Image src on a project page (../assets/...) or index (assets/...) -> path from tools/print.html."""
    return "../" + (src[3:] if src.startswith("../") else src)


def pretty(url):
    return re.sub(r"^https?://(www\.)?", "", url).rstrip("/")


def fig_aspect(p, fid, fc):
    """Width / height of a figure's visual (after any crop)."""
    kind = fc.get("kind", "img")
    if kind in ("svg", "sysprint"):
        m = re.search(r'viewBox="0 0 ([\d.]+) ([\d.]+)"', (p["svg"] if kind == "svg" else odmr_system_svg()) or "")
        if not m:
            die(f"{p['slug']}: cannot read the system SVG's viewBox")
        return float(m.group(1)) / float(m.group(2))
    if kind == "flow":
        return None
    info = p["figs"].get(fid)
    if info is None:
        die(f"{p['slug']}: figure {fid} not found on the web page (have: {', '.join(sorted(p['figs']))})")
    l, t, r, b = fc.get("crop", (0, 0, 0, 0))
    return (info["w"] - l - r) / (info["h"] - t - b)


def render_figure(p, fid, fc, n):
    kind = fc.get("kind", "img")
    notes = fc.get("notes", [])
    if kind == "svg":
        body = f'<div class="pimg svgwrap">{p["svg"]}</div>'
    elif kind == "sysprint":
        if not p["svg"]:
            die(f"{p['slug']}: figure {fid} redraws the page's system diagram, which was not found")
        svg = odmr_system_svg()
        check_svg_labels(p["svg"], svg, p["slug"])
        body = f'<div class="pimg svgwrap">{svg}</div>'
    elif kind == "flow":
        if not p["flow"]:
            die(f"{p['slug']}: figure {fid} needs the study-design flow, not found")
        body = f'<div class="pimg flowwrap"><ol class="flow">{p["flow"]}</ol></div>'
    else:
        info = p["figs"][fid]
        W, H = info["w"], info["h"]
        l, t, r, b = fc.get("crop", (0, 0, 0, 0))
        cw, ch = W - l - r, H - t - b
        by_n = {pn: (x, y) for pn, x, y in info["pins"]}
        keep = fc.get("pins", [])
        missing = [k for k in keep if k not in by_n]
        if missing:
            die(f"{p['slug']}/{fid}: pins {missing} are not on the web page (it has {sorted(by_n)})")
        if len(notes) != len(keep):
            die(f"{p['slug']}/{fid}: {len(keep)} pins but {len(notes)} notes")
        pins = []
        for i, k in enumerate(keep, 1):
            x = (by_n[k][0] / 100 * W - l) / cw * 100      # re-map the web position into the crop
            y = (by_n[k][1] / 100 * H - t) / ch * 100
            if not (0 <= x <= 100 and 0 <= y <= 100):
                die(f"{p['slug']}/{fid}: pin {k} falls outside the crop")
            pins.append(f'<span class="pin" style="left:{x:.2f}%;top:{y:.2f}%">{i}</span>')
        img = (f'<img src="{site_rel(info["src"])}" width="{W}" height="{H}" alt="{html.escape(info["alt"])}"'
               + (f' style="position:absolute;max-width:none;width:{W / cw * 100:.3f}%;left:{-l / cw * 100:.3f}%;'
                  f'top:{-t / ch * 100:.3f}%"' if (l or t or r or b) else "") + ">")
        body = f'<div class="pimg" style="aspect-ratio:{cw}/{ch}">{img}{"".join(pins)}</div>'
    ol = "".join(f'<li><i class="pn">{i}</i>{t}</li>' for i, t in enumerate(notes, 1))
    ol = f'<ol class="notes">{ol}</ol>' if ol else ""
    return (f'<figure class="pf">{body}<figcaption><b class="fn">Fig. {n}.</b> {fc["cap"]}{ol}'
            f'</figcaption></figure>')


BADGE = {"met": "✓", "part": '<b class="half"></b>', "no": "○"}


def render_score(cfg, cls=""):
    items = "".join(f'<li><span class="v {s}"><i>{BADGE[s]}</i>{lab}</span> {txt}</li>' for s, lab, txt in cfg["score"])
    return f'<div class="scorecard {cls}"><h2>Against the objective</h2><ul class="score">{items}</ul></div>'


def fact(p, label):
    for dt, dd in p["facts"]:
        if dt.lower() == label.lower():
            return dd
    die(f"{p['slug']}: no '{label}' in the page's fact grid (it has {[d for d, _ in p['facts']]})")


def print_tags(p, cfg):
    """The page's skill tags, or the subset (in page order) named in PRINT; every name must be on the page."""
    want = cfg.get("tags")
    if not want:
        return p["tags"]
    missing = [t for t in want if t not in p["tags"]]
    if missing:
        die(f"{p['slug']}: tags {missing} are not on the page (it has {p['tags']})")
    return [t for t in p["tags"] if t in want]


def print_did(p, cfg):
    """The page's "What I did" bullets, or the condensed print versions in PRINT. A condensed bullet must
    lead with a bold verb, and every bold verb in it must be one of the page's own bullet verbs, so each
    print bullet maps back to what the page says."""
    want = cfg.get("did")
    if not want:
        return p["did"]
    verbs = {v.lower() for b in p["did"] for v in re.findall(r"<strong>(.*?)</strong>", b)}
    for b in want:
        vs = [v.lower() for v in re.findall(r"<strong>(.*?)</strong>", b)]
        if not b.startswith("<strong>") or not vs or any(v not in verbs for v in vs):
            die(f"{p['slug']}: print bullet does not map to the page's bullets (verbs {sorted(verbs)}): {b[:70]}")
    return want


def render_side(p, cfg):
    did = "".join(f"<li>{b}</li>" for b in print_did(p, cfg))
    tags = "".join(f"<li>{html.escape(t)}</li>" for t in print_tags(p, cfg))
    score = render_score(cfg) if cfg.get("score_at") == "side" and cfg["score"] else ""
    band = render_band(p, cfg, "bandcell sideband") if cfg.get("band_at") == "side" else ""
    return (f'<aside class="side"><h2>My contribution</h2><p class="role">{fact(p, "My role")}</p>'
            f'<ul class="did">{did}</ul><ul class="tags">{tags}</ul>{score}{band}</aside>')


def render_row(p, cfg, row, fign):
    """One body row. Cells: "fig:<id>", "side", "score", "band", or a list of those (a vertical stack that
    takes the width of the figure in it). With h > 0 every figure is h px tall (width from its aspect ratio)
    and text cells share what is left; with h == 0 the figures stretch to equal heights across the row."""
    flat = [c for cell in row["cells"] for c in (cell if isinstance(cell, list) else [cell])]
    has_grow = any(not isinstance(c, list) and not c.startswith("fig:") for c in row["cells"])
    fixed = row["h"] > 0
    if not fixed and any(isinstance(c, list) for c in row["cells"]):
        die(f"{p['slug']}: a stacked cell needs a row height h")
    cells, used = [], 0

    def one(cell, stacked=False):
        nonlocal fign
        if cell.startswith("fig:"):
            fid = cell[4:]
            fc = cfg["figs"].get(fid)
            if fc is None:
                die(f"{p['slug']}: no PRINT entry for figure {fid}")
            fign += 1
            fig = render_figure(p, fid, fc, fign)
            if fc.get("kind") == "flow":
                return fc["w"], f"flex:0 0 {fc['w']}px;width:{fc['w']}px", fig
            asp = fig_aspect(p, fid, fc)
            if fixed:
                w = round(row["h"] * asp)
                return w, f"flex:0 0 {w}px;width:{w}px", fig
            return 0, f"flex:{asp:.4f} 1 0;min-width:0", fig     # stretch: equal heights, full row width
        text = {"side": lambda: render_side(p, cfg), "score": lambda: render_score(cfg, "boxed"),
                "band": lambda: render_band(p, cfg, "bandcell")}.get(cell)
        if text is None:
            die(f"{p['slug']}: unknown cell {cell!r}")
        return 0, None, text()

    for cell in row["cells"]:
        if isinstance(cell, list):                  # vertical stack: as wide as its widest figure
            parts = [one(c, True) for c in cell]
            w = max(pw for pw, _, _ in parts)
            if not w:
                die(f"{p['slug']}: a stacked cell needs a figure to set its width")
            used += w
            inner = "".join(f'<div class="stack-item">{html_}</div>' for _, _, html_ in parts)
            cells.append(f'<div class="cell stack" style="flex:0 0 {w}px;width:{w}px">{inner}</div>')
            continue
        w, style, html_ = one(cell)
        used += w
        cells.append(f'<div class="cell" style="{style}">{html_}</div>' if style
                     else f'<div class="cell grow">{html_}</div>')
    span = used + ROW_GAP * (len(row["cells"]) - 1)
    if fixed and has_grow and span > CONTENT_W - 150:
        die(f"{p['slug']}: row with h={row['h']} leaves under 150 px for text; lower h")
    if fixed and not has_grow and span > CONTENT_W + 1:
        die(f"{p['slug']}: row with h={row['h']} is {span} px wide, over the {CONTENT_W} px content width; lower h")
    return f'<div class="row">{"".join(cells)}</div>', fign


def render_band(p, cfg, cls):
    # Constraint items come from the page's constraint grid; "Team" and "Timeline" fall back to the hero's
    # fact grid (on some pages they live there), since the print header has no fact grid.
    have = dict(p["constraints"])
    for label in ("Team", "Timeline"):
        if label not in have and any(dt == label for dt, _ in p["facts"]):
            have[label] = fact(p, label)
    missing = [c for c in cfg["constraints"] if c not in have]
    if missing:
        die(f"{p['slug']}: constraints {missing} not on the page (it has {list(have)})")
    cons = " · ".join(f"<b>{html.escape(c)}:</b> {have[c]}" for c in cfg["constraints"])
    return f'<div class="{cls}"><h3>Constraints</h3><p>{cons}</p></div>'


def render_lesson(p, cfg):
    """The "What I learned" strip, condensed in PRINT from the page's lesson box. Every word of its bold lead
    must appear in the page's lesson text, so the print line maps back to what the page says."""
    text = cfg.get("lesson")
    if not text:
        return ""
    if not p["lesson"]:
        die(f"{p['slug']}: PRINT has a lesson but the page has no .lesson box")
    lead = re.search(r"<b>(.*?)</b>", text)
    page = p["lesson"].lower()
    if not lead or any(w not in page for w in re.findall(r"[a-z]+", lead.group(1).lower())):
        die(f"{p['slug']}: the print lesson's bold lead is not in the page's lesson box: {text[:60]}")
    return f'<div class="lessonband"><h3>What I learned</h3><p>{text}</p></div>'


def render_objective(cfg):
    """The objective, printed in the header under the outcome (the space beside the KPI tiles)."""
    return f'<p class="objective"><b>Objective</b> {cfg["objective"]}</p>'



def render_project(p, cfg, pageno, total):
    slug = p["slug"]
    url = f"{BASE_URL}/projects/{slug}.html"
    kpis = "".join(f"<div><strong>{p['kpis'][i][0]}</strong><span>{p['kpis'][i][1]}</span></div>"
                   for i in cfg["kpis"])
    rows, fign = [], 0
    for row in cfg["rows"]:
        r, fign = render_row(p, cfg, row, fign)
        rows.append(r)
    if not 2 <= fign <= 4:
        die(f"{slug}: {fign} figures; the print layout calls for 2 to 4")
    used_cells = [c for r in cfg["rows"] for cell in r["cells"] for c in (cell if isinstance(cell, list) else [cell])]
    band = "" if "band" in used_cells or cfg.get("band_at") == "side" else render_band(p, cfg, "band")
    band += render_lesson(p, cfg)
    code = f' · Code: <a href="{p["code"]}">{pretty(p["code"])}</a>' if p["code"] else ""
    n_proj = ORDER.index(slug) + 1
    return f"""
<section class="page proj" id="p{pageno}">
  <header class="phead">
    <p class="kicker">Project {n_proj} of {len(ORDER)} · {html.escape(p["kicker"])}</p>
    <h1>{html.escape(p["title"])}</h1>
    <div class="ph-split">
      <div class="ph-text"><p class="outcome">{cfg.get("outcome") or p["outcome"]}</p>{render_objective(cfg)}</div>
      <div class="kpis">{kpis}</div>
    </div>
  </header>
  <div class="pbody">{"".join(rows)}</div>
  {band}
  <footer class="foot"><span>Jimmy Gu · <a href="mailto:{CONTACT[0][1]}">{CONTACT[0][1]}</a> · Full page: <a href="{url}">{pretty(url)}</a>{code}</span>
    <span>{pageno} / {total}</span></footer>
</section>"""


def render_cover(ix, projects, total):
    glance = "".join(f"<div><dt>{html.escape(dt)}</dt><dd>{dd}</dd></div>" for dt, dd in ix["glance"])
    contact = "".join(f'<li><span>{lab}</span><a href="{href}">{html.escape(txt)}</a></li>'
                      for lab, txt, href in CONTACT)
    cards = []
    for i, slug in enumerate(ORDER):
        c, p, cfg = ix["cards"][slug], projects[slug], PRINT[slug]
        stats = "".join(f"<div><dt>{a}</dt><dd>{b}</dd></div>" for a, b in c["stats"])
        cards.append(f"""
    <a class="ccard" href="#p{i + 2}">
      <figure class="cthumb"><div class="cimg"><img src="{site_rel(c['thumb'])}" width="{c['w']}" height="{c['h']}"
        alt="{html.escape(c['alt'])}"></div><figcaption>{c['thumbcap']}</figcaption></figure>
      <div class="cbody">
        <p class="cmeta"><span class="kicker">{html.escape(c['kicker'])}</span><span class="cpg">Page {i + 2} →</span></p>
        <h3>{html.escape(p['title'])}</h3>
        <dl class="cstats">{stats}</dl>
        <p class="cout">{cfg['one_liner']}</p>
      </div>
    </a>""")
    return f"""
<section class="page cover" id="p1">
  <header class="ctop">
    <div class="cid">
      <p class="pill"><span class="dot"></span>{html.escape(ix["pill"])}</p>
      <h1 class="cname">Jimmy Gu</h1>
      <p class="ctag">{ix["tagline"]}</p>
      <p class="clead">{ix["lead"]}</p>
      <ul class="contact">{contact}</ul>
    </div>
    <dl class="glance">{glance}</dl>
  </header>
  <div class="chead"><h2>Four projects, one page each</h2>
    <p>Each page leads with the outcome and objective, then my role, annotated figures, results and what I learned.<br>Select a project to jump to its page.</p></div>
  <div class="cgrid">{"".join(cards)}</div>
  <footer class="foot"><span>Web version, with full methods and code: <a href="{BASE_URL}">{pretty(BASE_URL)}</a></span>
    <span>1 / {total}</span></footer>
</section>"""


CSS = r"""
@page { size: 8.5in 11in; margin: 0; }
:root { __TOKENS__ }
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { background: #fff; }
body { font-family: var(--sans); color: var(--ink); font-size: 8pt; line-height: 1.36;
  -webkit-print-color-adjust: exact; print-color-adjust: exact; -webkit-font-smoothing: antialiased; }
a { color: inherit; text-decoration: none; }
img { display: block; }
ul, ol { list-style: none; }
.page { width: __PW__px; height: __PH__px; padding: __PT__px __PX__px __PB__px; display: flex; flex-direction: column;
  overflow: hidden; position: relative; background: #fff; break-after: page; page-break-after: always; }
.page:last-child { break-after: auto; page-break-after: auto; }
@media screen { html, body { background: #8a8a8a; } .page { margin: 24px auto; box-shadow: 0 4px 24px rgb(0 0 0 / .3); } }

.kicker { font: 500 6.3pt/1.25 var(--mono); letter-spacing: .07em; text-transform: uppercase; color: var(--accent); }
.tags { display: flex; flex-wrap: wrap; gap: 2.2pt; }
.tags li { font: 500 5.9pt/1 var(--mono); background: var(--accent-soft); color: #2a3cb8; padding: 2.2pt 4.6pt; border-radius: 99px; }
.foot { margin-top: auto; display: flex; justify-content: space-between; gap: 12pt; padding-top: 5pt;
  border-top: .75pt solid var(--line); font: 400 6.2pt/1.3 var(--mono); color: var(--ink-2); }
.foot a { color: #2a3cb8; }

/* ---------- project pages ---------- */
.phead h1 { font-family: var(--serif); font-weight: 400; font-size: 20.5pt; line-height: 1.03; letter-spacing: -.01em; margin: 2.5pt 0 5pt; text-wrap: balance; }
.ph-split { display: flex; gap: 12pt; align-items: flex-start; }
.ph-split .ph-text { flex: 1 1 0; min-width: 0; }
.objective { font-size: 7.1pt; line-height: 1.38; color: var(--ink-2); margin-top: 4pt; padding-top: 3.5pt;
  border-top: .75pt dashed var(--line); }
.objective b { font: 500 5.6pt/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: var(--accent);
  margin-right: 3pt; vertical-align: .6pt; }
.outcome { font-size: 8.2pt; line-height: 1.42; color: #3d424b; }
.outcome strong { color: var(--ink); font-weight: 600; }
.ph-split .kpis { flex: 0 0 172pt; display: grid; grid-template-columns: 1fr 1fr; gap: 3.5pt; align-content: start; }
.kpis div { background: var(--bg); border-radius: 5pt; padding: 3.5pt 6pt 4pt; }
.kpis strong { display: block; font-family: var(--serif); font-weight: 400; font-size: 14pt; line-height: 1; white-space: nowrap; }
.kpis span { display: block; font-size: 6pt; line-height: 1.22; color: var(--ink-2); margin-top: 1.5pt; }
.pbody { display: flex; flex-direction: column; gap: __GAP__px; margin-top: 8pt; padding-top: 8pt; border-top: .75pt solid var(--line); }
.row { display: flex; gap: __GAP__px; align-items: flex-start; }
.cell.grow { flex: 1 1 0; min-width: 0; }
.cell.stack { display: flex; flex-direction: column; gap: __GAP__px; }

.pf { border: .75pt solid var(--line); border-radius: 5pt; overflow: hidden; background: #fff; }
.pimg { position: relative; overflow: hidden; background: #fbfbfa; }
.pimg img { width: 100%; height: auto; }
.pin { position: absolute; transform: translate(-50%, -50%); width: 11pt; height: 11pt; border-radius: 50%;
  background: var(--warm); color: #fff; font: 700 6.2pt/11pt var(--sans); text-align: center;
  box-shadow: 0 0 0 1.3pt #fff, 0 .5pt 2pt rgb(0 0 0 / .35); }
.pf figcaption { padding: 3.5pt 6pt 4.5pt; font-size: 6.6pt; line-height: 1.32; color: var(--ink-2); border-top: .75pt solid var(--line); }
.pf figcaption b { color: var(--ink); font-weight: 600; }
.pf figcaption em { font-style: italic; }
.notes { display: grid; gap: 1pt; margin-top: 2pt; }
.notes li { position: relative; padding-left: 10.5pt; }
.notes .pn { position: absolute; left: 0; top: .9pt; width: 8pt; height: 8pt; border-radius: 50%; background: var(--warm);
  color: #fff; font: 700 5.3pt/8pt var(--sans); font-style: normal; text-align: center; }

.side h2, .scorecard h2 { font: 500 5.9pt/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: var(--accent); margin-bottom: 4pt; }
.side .role { font-size: 7pt; line-height: 1.3; font-weight: 600; color: var(--ink); margin: -1.5pt 0 4pt; }
.did { display: grid; gap: 2.8pt; list-style: disc; padding-left: 1.05em; }
.did li { font-size: 7.2pt; line-height: 1.35; color: #33373f; }
.did li::marker { color: var(--accent); }
.did strong { color: var(--ink); font-weight: 600; }
.side .tags { margin-top: 5.5pt; }
.side .scorecard { margin-top: 8pt; }
.scorecard.boxed { background: var(--bg); border-radius: 5pt; padding: 6pt 7pt; }
.score { display: grid; gap: 2.6pt; }
.score li { font-size: 6.7pt; line-height: 1.3; color: #33373f; }
.side .score li { display: grid; grid-template-columns: auto 1fr; column-gap: 3pt; align-items: baseline; }
.side .score li .v { margin-right: 0; }
.v { display: inline-flex; align-items: center; gap: 2pt; font: 500 5.5pt/1 var(--mono); letter-spacing: .05em; text-transform: uppercase;
  padding: 1.7pt 4pt; border-radius: 99px; margin-right: 1.5pt; vertical-align: .5pt; white-space: nowrap; }
.v i { font: 700 6pt/1 var(--sans); font-style: normal; }
.v .half { display: inline-block; width: 5pt; height: 5pt; border-radius: 50%; border: .8pt solid currentColor;
  background: linear-gradient(90deg, currentColor 50%, transparent 50%); }
.v.met { background: #e3f4ea; color: #146c3b; }
.v.part { background: #fff1e0; color: #8a4b00; }
.v.no { background: #eceae4; color: #4a4f58; }
.boxed .score { gap: 4pt; }
.boxed .score li .v { display: flex; width: max-content; margin-bottom: 1.5pt; }

.band { display: flex; gap: 8pt; align-items: baseline; margin-top: 7pt; padding: 4.5pt 7pt 5pt;
  background: var(--bg); border-radius: 5pt; }
.band h3 { flex: 0 0 auto; }
.band h3 { font: 500 5.5pt/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: var(--ink-2); margin-bottom: 2pt; }
.band p { font-size: 6.5pt; line-height: 1.33; color: #3d424b; }
.band b, .bandcell b { color: var(--ink); font-weight: 600; }
.bandcell { background: var(--bg); border-radius: 5pt; padding: 6pt 7pt; display: grid; gap: 2pt; }
.bandcell h3 { font: 500 5.5pt/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: var(--ink-2); margin-bottom: 2pt; }
.bandcell p { font-size: 6.5pt; line-height: 1.33; color: #3d424b; }
.proj .foot { margin-top: auto; }
.side .sideband { margin-top: 8pt; gap: 3pt; }
.proj .band { margin-bottom: 0; }
.lessonband { display: flex; gap: 8pt; align-items: baseline; margin-top: 4pt; margin-bottom: 5pt; padding: 4.5pt 7pt 5pt;
  background: var(--accent-soft); border-left: 2pt solid var(--accent); border-radius: 5pt; }
.lessonband h3 { flex: 0 0 auto; font: 500 5.5pt/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: #2a3cb8; }
.lessonband p { font-size: 6.6pt; line-height: 1.33; color: #2b2f45; text-wrap: pretty; }
.lessonband b { color: var(--ink); font-weight: 600; }

/* ODMR system diagram (inline SVG from the web page; sizes are SVG user units, as on the page) */
.svgwrap svg { display: block; width: 100%; height: auto; background: #fbfbfa; font-family: var(--sans); }
.svgwrap .t { font-size: 21px; font-weight: 600; fill: #16181d; }
.svgwrap .s { font-size: 15.5px; fill: #555b66; }
.svgwrap .lane { font: 500 13px var(--mono); letter-spacing: .06em; fill: #555b66; }
.svgwrap .sw .t { fill: #2a3cb8; }

.svgwrap .sysprint { display: block; width: 100%; height: auto; background: #fbfbfa; font-family: var(--sans); }
.sysprint .t { font-size: 17.5px; font-weight: 600; fill: #16181d; }
.sysprint .s { font-size: 14px; fill: #555b66; }
.sysprint .lane { font: 500 12.5px var(--mono); letter-spacing: .06em; fill: #555b66; }
.sysprint .sw .t { fill: #2a3cb8; }

/* Mosaic study-design flow (HTML from the web page) */
.flowwrap { background: #fbfbfa; }
.flow { display: grid; grid-template-columns: repeat(3, 1fr); gap: 4pt; padding: 6pt; }
.flow li { background: var(--surface); border: .75pt solid var(--line); border-radius: 4pt; padding: 3.5pt 4.5pt 4.5pt; }
.flow li.me { border: 1.2pt solid var(--accent); background: var(--accent-soft); }
.flow li.shared { border: 1.2pt solid var(--accent); }
.flow li.rd { grid-column: 1 / -1; border-style: dashed; }
.flow .step { display: flex; justify-content: space-between; gap: 3pt; font: 500 5.4pt/1.2 var(--mono); letter-spacing: .05em;
  text-transform: uppercase; color: var(--ink-2); }
.flow .step em { font-style: normal; }
.flow .me .step em, .flow .shared .step em { color: #2a3cb8; }
.flow b { display: block; font-family: var(--serif); font-weight: 400; font-size: 13pt; line-height: 1.05; margin: 2pt 0 1pt; color: var(--ink); }
.flow .d { display: block; font-size: 6pt; line-height: 1.28; color: #33373f; }
.flow em { font-style: italic; }

/* ---------- cover ---------- */
.cover { padding-top: 0.42in; }
.ctop { display: grid; grid-template-columns: minmax(0, 1.5fr) minmax(0, 1fr); gap: 20pt; align-items: start; }
.pill { display: inline-flex; align-items: center; gap: 5pt; font-size: 7.2pt; color: var(--ink-2); background: var(--surface);
  border: .75pt solid var(--line); padding: 2.5pt 8pt; border-radius: 99px; }
.dot { width: 5.5pt; height: 5.5pt; border-radius: 50%; background: #22c55e; box-shadow: 0 0 0 2.5pt #22c55e22; }
.cname { font-family: var(--serif); font-weight: 400; font-size: 42pt; line-height: .95; letter-spacing: -.02em; margin-top: 8pt; }
.ctag { font-family: var(--serif); font-size: 16pt; line-height: 1.15; margin-top: 5pt; text-wrap: balance; }
.ctag em { color: var(--accent); font-style: italic; }
.clead { font-size: 8.6pt; line-height: 1.5; color: #3d424b; margin-top: 6pt; }
.contact { display: grid; grid-template-columns: auto auto; gap: 3pt 16pt; margin-top: 9pt; justify-content: start; }
.contact li { font: 400 7.2pt/1.3 var(--mono); color: var(--ink); white-space: nowrap; }
.contact span { display: inline-block; width: 44pt; font-size: 5.7pt; letter-spacing: .07em; text-transform: uppercase; color: var(--ink-2); }
.contact a { color: #2a3cb8; }
.glance { background: var(--surface); border: .75pt solid var(--line); border-radius: 7pt; padding: 1pt 10pt; margin-top: 4pt; }
.glance div { padding: 5pt 0; border-bottom: .75pt dashed var(--line); }
.glance div:last-child { border-bottom: 0; }
.glance dt { font: 500 5.7pt/1.2 var(--mono); letter-spacing: .07em; text-transform: uppercase; color: var(--accent); }
.glance dd { font-size: 7.8pt; font-weight: 600; line-height: 1.3; margin-top: 1pt; }
.glance dd span { display: block; font-size: 6.8pt; font-weight: 400; color: var(--ink-2); margin-top: 1pt; }
.chead { display: flex; align-items: flex-end; justify-content: space-between; gap: 12pt; margin-top: 11pt; padding-top: 8pt;
  border-top: .75pt solid var(--line); }
.chead h2 { font-family: var(--serif); font-weight: 400; font-size: 17pt; line-height: 1; }
.chead p { font-size: 6.6pt; line-height: 1.35; color: var(--ink-2); text-align: right; }
.cgrid { display: grid; grid-template-columns: 1fr 1fr; gap: 9pt; margin-top: 8pt; }
.ccard { display: flex; flex-direction: column; border: .75pt solid var(--line); border-radius: 7pt; overflow: hidden; background: #fff; }
.cthumb { background: var(--media); border-bottom: .75pt solid var(--line); }
.cimg { height: 103pt; padding: 6pt 6pt 3pt; display: flex; align-items: center; justify-content: center; }
.cimg img { max-width: 100%; max-height: 100%; width: auto; height: auto; border-radius: 3pt; background: #fff;
  box-shadow: 0 .5pt 1.5pt rgb(22 24 29 / .12); }
.cthumb figcaption { font-size: 6pt; line-height: 1.3; color: var(--ink-2); text-align: center; padding: 0 6pt 3.5pt; }
.cthumb figcaption b { color: var(--ink); font-weight: 600; }
.cbody { padding: 5.5pt 9pt 7pt; }
.cmeta { display: flex; justify-content: space-between; gap: 6pt; }
.cmeta .kicker { font-size: 5.8pt; }
.cpg { font: 500 5.8pt/1.25 var(--mono); letter-spacing: .06em; text-transform: uppercase; color: #2a3cb8; white-space: nowrap; }
.ccard h3 { font-family: var(--serif); font-weight: 400; font-size: 12.5pt; line-height: 1.1; margin: 2pt 0 4pt; text-wrap: balance; }
.cstats { display: grid; grid-template-columns: 1fr 1fr; gap: 8pt; padding: 3.5pt 0; border-top: .75pt solid var(--line); border-bottom: .75pt solid var(--line); }
.cstats dt { font-family: var(--serif); font-size: 14pt; line-height: 1; white-space: nowrap; }
.cstats dd { font-size: 6.1pt; line-height: 1.28; color: var(--ink-2); margin-top: 1.5pt; }
.cout { font-size: 7.1pt; line-height: 1.4; color: #3d424b; margin-top: 4pt; }
.cover .foot { margin-top: auto; }
"""

CHECK_JS = r"""
<script>
// Layout self-check, read by build_pdf.py through Chrome's --dump-dom (runs only when the URL ends in #check). Fails if anything extends past a
// page's padding box, a text box clips its content, an image or web font fails to load. Also reports
// each page's spare height (space left above the footer) so the layout can be tuned.
window.addEventListener('load', async () => {
  if (location.hash !== '#check') return;        // only in the check pass, so the PDF keeps its real title
  await document.fonts.ready;
  const probs = [], free = [], share = [], rows = [];
  document.querySelectorAll('.page').forEach((pg, i) => {
    const r = pg.getBoundingClientRect(), cs = getComputedStyle(pg);
    const box = { l: r.left + parseFloat(cs.paddingLeft) - 1, r: r.right - parseFloat(cs.paddingRight) + 1,
                  t: r.top + parseFloat(cs.paddingTop) - 1, b: r.bottom - parseFloat(cs.paddingBottom) + 1 };
    pg.querySelectorAll('*').forEach(el => {
      if (el.closest('svg') && el.tagName.toLowerCase() !== 'svg') return;
      if (el.closest('.pimg') && el.tagName === 'IMG') return;   // cropped images overflow their frame by design
      const b = el.getBoundingClientRect();
      if (!b.width && !b.height) return;
      const over = Math.max(b.right - box.r, box.l - b.left, b.bottom - box.b, box.t - b.top);
      if (over > 0.5) probs.push(`page ${i + 1}: ${el.tagName.toLowerCase()}.${el.className} by ${Math.round(over)}px`);
    });
    pg.querySelectorAll('figcaption, .side, .band, .bandcell, .lessonband, .cbody, .phead, .scorecard').forEach(el => {
      if (el.scrollHeight > el.clientHeight + 1) probs.push(`page ${i + 1}: ${el.className || el.tagName} clipped`);
    });
    const area = (sel) => [...pg.querySelectorAll(sel)].reduce((a, e) => { const q = e.getBoundingClientRect(); return a + q.width * q.height; }, 0);
    const content = (box.r - box.l - 2) * (box.b - box.t - 2);
    rows.push(`p${i + 1} ` + [...pg.querySelectorAll('.row')].map(rw => [...rw.children].map(c => {
      const k = c.firstElementChild; return Math.round((k || c).getBoundingClientRect().height); }).join('/')).join(' + '));
    share.push(`p${i + 1} ${Math.round(100 * area('.pimg, .cimg') / content)}%/${Math.round(100 * area('.pf, .cthumb') / content)}%`);
    const foot = pg.querySelector('.foot'), prev = foot && foot.previousElementSibling;
    if (prev) { const fb = foot.getBoundingClientRect();
      free.push(`p${i + 1} ${Math.round((box.b - 1 - fb.bottom) + (fb.top - prev.getBoundingClientRect().bottom
                 - parseFloat(getComputedStyle(prev).marginBottom)))}px`); }
  });
  document.querySelectorAll('img').forEach(im => { if (!im.complete || !im.naturalWidth) probs.push('image failed: ' + im.getAttribute('src')); });
  const loaded = new Set([...document.fonts].filter(f => f.status === 'loaded').map(f => f.family.replace(/["']/g, '')));
  ['Inter', 'Instrument Serif', 'JetBrains Mono'].forEach(f => { if (!loaded.has(f)) probs.push('font not loaded: ' + f); });
  document.title = 'PRINTCHECK:' + (probs.length ? probs.slice(0, 25).join(' | ') : 'ok') + ' || spare: ' + free.join(', ') + ' || visual share (image / figure incl. caption): ' + share.join(', ') + ' || cell heights per row: ' + rows.join(', ');
});
</script>"""


def build_html():
    ix = extract_index()
    projects = {s: extract_project(s) for s in ORDER}
    for s in ORDER:
        if s not in ix["cards"]:
            die(f"index.html has no card for {s}")
    total = 1 + len(ORDER)
    pages = [render_cover(ix, projects, total)]
    pages += [render_project(projects[s], PRINT[s], i + 2, total) for i, s in enumerate(ORDER)]
    css = (CSS.replace("__TOKENS__", root_tokens()).replace("__PW__", str(PAGE_W)).replace("__PH__", str(PAGE_H))
           .replace("__PT__", str(PAD_T)).replace("__PB__", str(PAD_B)).replace("__PX__", str(PAD_X))
           .replace("__GAP__", str(ROW_GAP)))
    doc = f"""<!DOCTYPE html>
<!-- GENERATED by tools/build_pdf.py from index.html and projects/*.html. Do not edit by hand: edit the
     pages (or the PRINT settings in build_pdf.py) and rebuild. Open it in a browser to preview. -->
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="robots" content="noindex">
<title>Jimmy Gu | Portfolio</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="{html.escape(FONTS_URL)}" rel="stylesheet">
<style>{css}</style>
</head>
<body>
{"".join(pages)}
{CHECK_JS}
</body>
</html>
"""
    with open(OUT_HTML, "w", encoding="utf-8", newline="\n") as f:
        f.write(doc)
    return doc, sorted(set(re.findall(r'<img src="\.\./([^"]+)"', doc))), projects


# ---------------------------------------------------------------------------------------------
# Staging, Chrome, checks
# ---------------------------------------------------------------------------------------------
def check_images(doc, images):
    try:
        from PIL import Image
    except ImportError:
        print("  (Pillow not installed: skipping the image size check)")
        return
    for rel in images:
        with Image.open(os.path.join(SITE, *rel.split("/"))) as im:
            w, h = im.size
        for dw, dh in re.findall(r'src="\.\./' + re.escape(rel) + r'" width="(\d+)" height="(\d+)"', doc):
            if abs(int(dw) / int(dh) - w / h) > 0.01:
                die(f"{rel}: the page says {dw}x{dh} but the file is {w}x{h}; pins would be misplaced")


def fetch(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": LEGACY_UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def static_fonts_css(site):
    """Download static copies of the web fonts next to the staged print.html; return @font-face CSS."""
    css = fetch(FONTS_URL).decode("utf-8")
    fdir = os.path.join(site, "tools", "fonts")
    os.makedirs(fdir, exist_ok=True)

    def grab(m):
        url = m.group(1)
        name = re.sub(r"[^\w.-]", "_", url.split("/s/", 1)[-1])
        with open(os.path.join(fdir, name), "wb") as f:
            f.write(fetch(url))
        return f"url(fonts/{name})"

    css = re.sub(r"url\((https://fonts\.gstatic\.com/[^)]+)\)", grab, css)
    if "@font-face" not in css or "fonts/" not in css:
        raise RuntimeError("unexpected font CSS")
    return css


def stage(tmp, doc, images, max_px):
    site = os.path.join(tmp, "site")
    os.makedirs(os.path.join(site, "tools"))
    try:
        fcss = static_fonts_css(site)
        doc = re.sub(r'<link rel="preconnect"[^>]*>\n', "", doc)
        doc = doc.replace(f'<link href="{html.escape(FONTS_URL)}" rel="stylesheet">', f"<style>{fcss}</style>")
        print("  fonts: static copies staged")
    except Exception as e:  # offline or API change: fall back to the variable web fonts (Type 3 in the PDF)
        print(f"  fonts: WARNING, could not stage static fonts ({e}); using Google's web fonts")
    with open(os.path.join(site, "tools", "print.html"), "w", encoding="utf-8") as f:
        f.write(doc)
    for rel in images:
        src, dst = os.path.join(SITE, *rel.split("/")), os.path.join(site, *rel.split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if max_px:
            from PIL import Image
            with Image.open(src) as im:
                if max(im.size) > max_px:
                    im = im.copy()
                    im.thumbnail((max_px, max_px), Image.LANCZOS)
                    if rel.lower().endswith((".jpg", ".jpeg")):
                        im.convert("RGB").save(dst, quality=85, optimize=True)
                    else:
                        im.save(dst, optimize=True)
                    continue
        shutil.copy2(src, dst)
    return "file:///" + os.path.join(site, "tools", "print.html").replace("\\", "/")


def chrome(exe, tmp, prof, args, timeout=180):
    cmd = [exe, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
           "--allow-file-access-from-files", "--hide-scrollbars", "--virtual-time-budget=8000",
           f"--user-data-dir={os.path.join(tmp, prof)}"] + args
    return subprocess.run(cmd, capture_output=True, timeout=timeout)


def finalize_pdf(path, projects):
    """Set the PDF's metadata and add bookmarks (one per page). Needs PyMuPDF; skipped without it."""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            print("  (PyMuPDF not installed: skipping metadata and bookmarks)")
            return
    doc = pymupdf.open(path)
    doc.set_metadata({"title": "Jimmy Gu | Portfolio", "author": "Jimmy Gu",
                      "subject": "Project portfolio: data analysis, data pipelines and research engineering",
                      "keywords": "data analyst, data engineering, research engineering, Python, " + BASE_URL,
                      "creator": "tools/build_pdf.py (headless Chrome)", "producer": doc.metadata.get("producer", "")})
    toc = [[1, "Cover and project index", 1]]
    toc += [[1, f"{i}. {html.unescape(projects[s]['title'])}", i + 1] for i, s in enumerate(ORDER, 1)]
    if doc.page_count == len(toc):
        doc.set_toc(toc)
    tmp = path + ".tmp"
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    os.replace(tmp, path)


def verify_pdf(path, preview):
    size = os.path.getsize(path)
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            print(f"  PDF: {size / 1024:.0f} KB (install PyMuPDF to check the page count and render previews)")
            return None, []
    doc = pymupdf.open(path)
    n = doc.page_count
    sizes = {(round(p.rect.width), round(p.rect.height)) for p in doc}
    fonts = sorted({(f[3].split("+")[-1] or "?", f[2]) for p in doc for f in p.get_fonts()})
    type3 = [f for f, t in fonts if t == "Type3"]
    print(f"  PDF: {n} pages, {size / 1024:.0f} KB, page size {sizes} pt")
    print(f"  fonts: {', '.join(f'{f} ({t})' for f, t in fonts)}")
    if preview:
        os.makedirs(preview, exist_ok=True)
        for i, p in enumerate(doc, 1):
            p.get_pixmap(dpi=110).save(os.path.join(preview, f"page{i}.png"))
        print(f"  previews: {preview}")
    return n, type3


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--html-only", action="store_true", help="only (re)write tools/print.html")
    ap.add_argument("--preview", metavar="DIR", help="also render each PDF page to DIR/pageN.png (PyMuPDF)")
    ap.add_argument("--max-img-px", type=int, default=0,
                    help="downscale images in the print copy to at most N px on the long side (default: off)")
    ap.add_argument("--keep-temp", action="store_true", help="keep the Windows temp dir for debugging")
    ap.add_argument("--chrome", default=CHROME_DEFAULT, help="path to chrome.exe")
    a = ap.parse_args()

    doc, images, projects = build_html()
    print(f"wrote {OUT_HTML} ({len(images)} images)")
    check_images(doc, images)
    if a.html_only:
        return
    if os.name != "nt":
        die("run this with a Windows Python so it can drive Windows Chrome (see the docstring)")
    if not os.path.exists(a.chrome):
        die(f"Chrome not found at {a.chrome}")

    tmp = tempfile.mkdtemp(prefix="jg_pdf_")
    try:
        url = stage(tmp, doc, images, a.max_img_px)
        r = chrome(a.chrome, tmp, "prof_check", ["--window-size=1000,1400", "--dump-dom", url + "#check"])
        m = re.search(r"<title>PRINTCHECK:(.*?)</title>", r.stdout.decode("utf-8", "replace"), re.S)
        if not m:
            die("layout check did not run (no PRINTCHECK title in Chrome's DOM dump)")
        status, _, rest = html.unescape(m.group(1)).strip().partition(" || ")
        spare, _, rest = rest.partition(" || ")
        share, _, rowh = rest.partition(" || ")
        print(f"  layout check: {status}\n  {spare}\n  {share}\n  {rowh}")
        pdf_tmp = os.path.join(tmp, "out.pdf")
        r = chrome(a.chrome, tmp, "prof_pdf", ["--no-pdf-header-footer", f"--print-to-pdf={pdf_tmp}", url])
        if not os.path.exists(pdf_tmp):
            die("Chrome did not write a PDF:\n" + r.stderr.decode("utf-8", "replace")[-2000:])
        shutil.copyfile(pdf_tmp, OUT_PDF)
    finally:
        if a.keep_temp:
            print(f"  temp dir kept: {tmp}")
        else:
            shutil.rmtree(tmp, ignore_errors=True)
    finalize_pdf(OUT_PDF, projects)
    print(f"wrote {OUT_PDF}")
    n, type3 = verify_pdf(OUT_PDF, a.preview)
    problems = []
    if status != "ok":
        problems.append("layout check failed")
    if n is not None and n != EXPECTED_PAGES:
        problems.append(f"expected {EXPECTED_PAGES} pages, got {n}")
    if os.path.getsize(OUT_PDF) > MAX_PDF_BYTES:
        problems.append("the PDF is over 5 MB: try --max-img-px 1400")
    if type3:
        print(f"  note: Type 3 fonts in the PDF ({', '.join(type3)}); static font staging did not take effect")
    if problems:
        die("; ".join(problems))
    print("OK")


if __name__ == "__main__":
    main()
