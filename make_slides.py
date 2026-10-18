"""
Generate a ~15-minute presentation introducing two 2026 recommender-system papers:

  [A] LISRec: Modeling User Preferences with Learned Item Shortcuts for
      Sequential Recommendation (KDD '26)
  [B] Modeling Stage-wise Evolution of User Interests for News Recommendation (WWW '26)

Steps:
  1. Crop figures / tables out of the PDFs (PyMuPDF) into ./figs
  2. Build a 16:9 deck with python-pptx -> ./paper_presentation.pptx
     (speaker notes are loaded from 講稿.md, one '## 投影片 N' section per slide)

Usage:
    pip install pymupdf python-pptx
    python make_slides.py
"""

import re
from pathlib import Path

import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
PAPER_A = ROOT / "3770854.3780337.pdf"    # LISRec (KDD'26)
PAPER_B = ROOT / "2603.10471v1.pdf"       # stage-wise news rec (WWW'26)
FIG_DIR = ROOT / "figs"
OUT = ROOT / "paper_presentation.pptx"

# ----------------------------------------------------------------------------
# 1. Figure cropping
# ----------------------------------------------------------------------------
# (pdf, 1-based page, clip rect in PDF points (x0, y0, x1, y1)); pages are 612x792
CROPS = {
    "a_fig1_idea":      (PAPER_A, 2, (52, 80, 297, 384)),
    "a_fig2_arch":      (PAPER_A, 3, (52, 80, 563, 352)),
    "a_tab1_results":   (PAPER_A, 5, (52, 124, 563, 301)),
    "a_fig3_tau":       (PAPER_A, 7, (54, 82, 300, 192)),
    "a_fig6_effect":    (PAPER_A, 8, (52, 258, 300, 386)),
    "a_tab5_case":      (PAPER_A, 9, (52, 116, 563, 330)),
    "b_fig1_framework": (PAPER_B, 4, (58, 80, 560, 300)),
    "b_tab2_results":   (PAPER_B, 7, (78, 110, 560, 233)),
    "b_tab3_ablation":  (PAPER_B, 7, (330, 255, 560, 346)),
    "b_fig2_window":    (PAPER_B, 8, (56, 78, 298, 166)),
    "b_tab4_capability": (PAPER_B, 8, (315, 133, 562, 256)),
}


def crop_figures(dpi=250):
    FIG_DIR.mkdir(exist_ok=True)
    paths = {}
    docs = {}
    for name, (pdf, page, rect) in CROPS.items():
        doc = docs.setdefault(pdf, pymupdf.open(pdf))
        pix = doc[page - 1].get_pixmap(dpi=dpi, clip=pymupdf.Rect(*rect))
        out = FIG_DIR / f"{name}.png"
        pix.save(out)
        paths[name] = out
    for d in docs.values():
        d.close()
    return paths


# ----------------------------------------------------------------------------
# 2. Speaker notes from the Chinese script (講稿.md, "## 投影片 N｜..." sections)
# ----------------------------------------------------------------------------
SCRIPT = ROOT / "講稿.md"


def load_notes():
    notes = {}
    if not SCRIPT.exists():
        return notes
    current = None
    for line in SCRIPT.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^##\s*投影片\s*(\d+)\s*[｜|]\s*(.*)$", line)
        if m:
            current = int(m.group(1))
            notes[current] = [m.group(2).strip()]
        elif current is not None:
            notes[current].append(line)
    return {k: "\n".join(v).strip() for k, v in notes.items()}


# ----------------------------------------------------------------------------
# 3. Deck styling
#    Mostly neutral: dark navy + greys. The two papers keep one muted colour
#    each, used only for the A / B chips and small headings.
# ----------------------------------------------------------------------------
SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)

INK = RGBColor(0x1B, 0x26, 0x3B)       # dark navy (text, title/closing backgrounds)
MUTED = RGBColor(0x5B, 0x67, 0x7A)     # captions
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xA8, 0xB2, 0xC1)      # secondary text on dark slides
LIGHT = RGBColor(0xD6, 0xDC, 0xE5)
TINT = RGBColor(0xF2, 0xF3, 0xF5)      # neutral box fill (all cards)
TEAL = RGBColor(0x3D, 0x6B, 0x73)      # Paper A, muted slate-teal
CORAL = RGBColor(0x9A, 0x5B, 0x4B)     # Paper B, muted brick

HEAD_FONT = "Cambria"
BODY_FONT = "Calibri"

MARGIN = Inches(0.6)
TITLE_Y = Inches(0.45)
CONTENT_Y = Inches(1.55)


def _set_font(run, size, color=INK, bold=False, italic=False, font=BODY_FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = color


def _bullet(paragraph, level=0, char="•"):
    """Turn a paragraph into a real (non-literal) bullet with hanging indent."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.28 + 0.3 * level))))
    pPr.set("indent", str(int(-Inches(0.24))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": char}))


def add_text(slide, x, y, w, h, text, size=16, color=INK, bold=False,
             italic=False, font=BODY_FONT, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        _set_font(r, size, color, bold, italic, font)
    return tb


def add_bullets(slide, x, y, w, h, items, size=17, color=INK, space_after=8):
    """items: list of str | (str, level). 'Label: rest' renders Label in bold."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    for i, item in enumerate(items):
        text, level = (item, 0) if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _bullet(p, level, "•" if level == 0 else "–")
        p.space_after = Pt(space_after)
        sz = size if level == 0 else size - 2
        if ": " in text and text.index(": ") < 40:
            head, rest = text.split(": ", 1)
            r = p.add_run()
            r.text = head + ": "
            _set_font(r, sz, color, bold=True)
            r = p.add_run()
            r.text = rest
            _set_font(r, sz, color)
        else:
            r = p.add_run()
            r.text = text
            _set_font(r, sz, color)
    return tb


def add_box(slide, x, y, w, h, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.adjustments[0] = 0.08
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_chip(slide, x, y, label, color, d=Inches(0.62), size=20):
    """The deck's motif: a circle carrying the paper tag (A / B)."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    c.fill.solid()
    c.fill.fore_color.rgb = color
    c.line.fill.background()
    c.shadow.inherit = False
    tf = c.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    _set_font(r, size, WHITE, bold=True, font=HEAD_FONT)
    return c


def add_card(slide, x, y, w, h, head, body, head_color=INK, body_size=15):
    add_box(slide, x, y, w, h, TINT)
    add_text(slide, x + Inches(0.25), y + Inches(0.2), w - Inches(0.5),
             Inches(0.45), head, size=18, color=head_color, bold=True,
             font=HEAD_FONT)
    add_text(slide, x + Inches(0.25), y + Inches(0.7), w - Inches(0.5),
             h - Inches(0.8), body, size=body_size)


def add_stat(slide, x, y, w, big, label, big_size=44):
    add_text(slide, x, y, w, Inches(0.85), big, size=big_size, color=INK,
             bold=True, font=HEAD_FONT)
    add_text(slide, x, y + Inches(0.85), w, Inches(0.7), label, size=13,
             color=MUTED)


def add_picture_fit(slide, path, x, y, w, h, caption=None):
    """Insert image scaled to fit inside (x, y, w, h), centred, aspect kept."""
    with Image.open(path) as im:
        iw, ih = im.size
    cap_h = Inches(0.35) if caption else 0
    box_h = h - cap_h
    scale = min(w / iw, box_h / ih)
    pw, ph = int(iw * scale), int(ih * scale)
    px = x + (w - pw) // 2
    py = y + (box_h - ph) // 2
    pic = slide.shapes.add_picture(str(path), px, py, pw, ph)
    pic.line.color.rgb = RGBColor(0xD5, 0xDA, 0xE1)
    pic.line.width = Pt(0.75)
    if caption:
        add_text(slide, x, py + ph + Inches(0.08), w, Inches(0.3), caption,
                 size=11, color=MUTED, italic=True, align=PP_ALIGN.CENTER)
    return pic


def new_slide(prs, dark=False):
    s = prs.slides.add_slide(prs.slide_layouts[6])   # blank
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = INK if dark else WHITE
    return s


def slide_title(slide, text, tag=None, subtitle=None):
    x = MARGIN
    if tag:
        add_chip(slide, MARGIN, TITLE_Y + Inches(0.02), tag[0], tag[1])
        x = MARGIN + Inches(0.85)
    add_text(slide, x, TITLE_Y - Inches(0.02), SLIDE_W - x - MARGIN, Inches(0.7),
             text, size=32, bold=True, font=HEAD_FONT, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, x, TITLE_Y + Inches(0.7), SLIDE_W - x - MARGIN,
                 Inches(0.4), subtitle, size=14, color=MUTED, italic=True)


def footer(slide, n):
    add_text(slide, SLIDE_W - Inches(1.3), SLIDE_H - Inches(0.45), Inches(0.7),
             Inches(0.3), str(n), size=11, color=MUTED, align=PP_ALIGN.RIGHT)


A = ("A", TEAL)
B = ("B", CORAL)


# ----------------------------------------------------------------------------
# 4. Slides
# ----------------------------------------------------------------------------
def build(figs, notes):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    W = SLIDE_W - 2 * MARGIN
    slides = []

    def done(s, numbered=True):
        slides.append(s)
        if numbered:
            footer(s, len(slides))

    # --- 1. Title -----------------------------------------------------------
    s = new_slide(prs, dark=True)
    add_text(s, MARGIN, Inches(1.5), W, Inches(0.4),
             "BIG DATA ANALYTICS · PAPER PRESENTATION", size=14, color=SOFT, bold=True)
    add_text(s, MARGIN, Inches(2.0), W, Inches(1.0),
             "What is a user history really telling us?",
             size=44, color=WHITE, bold=True, font=HEAD_FONT)
    add_text(s, MARGIN, Inches(3.15), W, Inches(0.6),
             "Two 2026 papers on modelling user preferences from interaction logs",
             size=20, color=LIGHT, italic=True)
    for i, (tag, title, venue) in enumerate([
        (A, "LISRec: Modeling User Preferences with Learned Item Shortcuts for Sequential Recommendation",
         "Xin, Liu, Mei, Yan, Yu, Wang, Chen, Gu, Yu, Xiong · KDD '26"),
        (B, "Modeling Stage-wise Evolution of User Interests for News Recommendation",
         "Cheng, Jin, Zhang, Chen, Duan, Wang · WWW '26"),
    ]):
        y = Inches(4.3) + i * Inches(1.05)
        add_chip(s, MARGIN, y, tag[0], tag[1], d=Inches(0.55), size=18)
        add_text(s, MARGIN + Inches(0.8), y - Inches(0.04), W - Inches(0.8), Inches(0.4),
                 title, size=17, color=WHITE, bold=True)
        add_text(s, MARGIN + Inches(0.8), y + Inches(0.34), W - Inches(0.8), Inches(0.35),
                 venue, size=13, color=SOFT)
    done(s, numbered=False)

    # --- 2. Why these two papers -------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Why we picked these two papers")
    cw = (W - Inches(0.6)) / 3
    for i, (h, b) in enumerate([
        ("Same raw material",
         "Both learn user preferences only from large interaction logs "
         "(up to 35M clicks), the classic big-data setting of recommender systems."),
        ("Two blind spots",
         "Raw histories are treated as a clean, static signal. Paper A says it is "
         "not clean (noisy clicks); Paper B says it is not static (interests drift)."),
        ("Fresh & reproducible",
         "Accepted at KDD '26 and WWW '26, top venues for web / data mining; "
         "both release code, and both use graphs to restructure the history."),
    ]):
        add_card(s, MARGIN + i * (cw + Inches(0.3)), CONTENT_Y + Inches(0.1), cw,
                 Inches(2.1), h, b)
    y = Inches(4.3)
    add_text(s, MARGIN, y, W, Inches(0.4), "One question, two answers",
             size=20, bold=True, font=HEAD_FONT)
    half = (W - Inches(0.4)) / 2
    for i, (tag, head, body) in enumerate([
        (A, "Semantics: some clicks are noise",
         "Keep only the items that are semantically consistent with each other; drop the outliers."),
        (B, "Time: interests evolve",
         "Split the history into time stages and model how preferences move from stage to stage."),
    ]):
        x = MARGIN + i * (half + Inches(0.4))
        add_box(s, x, y + Inches(0.5), half, Inches(1.45), TINT)
        add_chip(s, x + Inches(0.25), y + Inches(0.75), tag[0], tag[1], d=Inches(0.55), size=18)
        add_text(s, x + Inches(1.0), y + Inches(0.65), half - Inches(1.2), Inches(0.4),
                 head, size=18, bold=True)
        add_text(s, x + Inches(1.0), y + Inches(1.05), half - Inches(1.2), Inches(0.8),
                 body, size=14)
    done(s)

    # --- 3. Paper A: problem -----------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Not every click reflects a real preference", tag=A,
                subtitle="LISRec: Modeling User Preferences with Learned Item Shortcuts · KDD '26")
    add_picture_fit(s, figs["a_fig1_idea"], MARGIN, CONTENT_Y, Inches(5.6), Inches(5.4),
                    caption="Figure 1 of Paper A: a LEGO fan clicks a Nintendo Switch")
    x0 = MARGIN + Inches(6.0)
    rw = SLIDE_W - MARGIN - x0
    add_bullets(s, x0, CONTENT_Y + Inches(0.2), rw, Inches(3.4), [
        "Histories contain noise: accidental clicks, momentary curiosity",
        "ID-based (SASRec, BERT4Rec) and text-based (TASTE, HLLM) models use every item equally → noise propagates",
        "Existing denoising re-weights items or uses implicit contrastive learning → no explicit removal, hard to interpret",
        "Idea: find the semantic \"shortcut\" - the subset of items that agree with each other - and model the user from it",
    ], size=16)
    add_box(s, x0, CONTENT_Y + Inches(3.0), rw, Inches(1.2), TINT)
    add_text(s, x0 + Inches(0.25), CONTENT_Y + Inches(3.1), rw - Inches(0.5), Inches(1.0),
             "History [LEGO Iron Man, LEGO Stitch, LEGO Bugatti, Switch Mario Kart, LEGO R2-D2] "
             "→ shortcut keeps the LEGO items and drops the Switch.",
             size=15, anchor=MSO_ANCHOR.MIDDLE)
    done(s)

    # --- 4. Paper A: architecture ------------------------------------------
    s = new_slide(prs)
    slide_title(s, "LISRec pipeline: encode → build graph → keep the core", tag=A)
    add_picture_fit(s, figs["a_fig2_arch"], MARGIN, CONTENT_Y, Inches(8.0), Inches(5.3),
                    caption="Figure 2 of Paper A: model architecture")
    x0 = MARGIN + Inches(8.3)
    sw = SLIDE_W - MARGIN - x0
    for i, (num, h, b) in enumerate([
        ("1", "Task-agnostic item encoder",
         "T5 instruction-tuned on next-item prediction across datasets, using item text only (no IDs)"),
        ("2", "Personal item graph",
         "Edge between two clicked items if cos(vi, vj) ≥ τ (τ = 0.7)"),
        ("3", "Shortcut = largest component",
         "Its items become the denoised history fed to the recommender (TASTE / BERT4Rec)"),
    ]):
        y = CONTENT_Y + Inches(0.1) + i * Inches(1.75)
        add_chip(s, x0, y, num, TEAL, d=Inches(0.5), size=16)
        add_text(s, x0 + Inches(0.7), y, sw - Inches(0.7), Inches(0.4), h, size=16, bold=True)
        add_text(s, x0 + Inches(0.7), y + Inches(0.42), sw - Inches(0.7), Inches(1.2), b, size=14)
    done(s)

    # --- 5. Paper A: case study --------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Case study: the perfume that does not belong", tag=A)
    add_picture_fit(s, figs["a_tab5_case"], MARGIN, CONTENT_Y, Inches(8.4), Inches(4.6),
                    caption="Table 5 of Paper A: Amazon Beauty case")
    x0 = MARGIN + Inches(8.7)
    add_bullets(s, x0, CONTENT_Y + Inches(0.2), SLIDE_W - MARGIN - x0, Inches(4.6), [
        "n1, n2, n4 are moisturisers; pairwise similarity ≈ 0.44–0.46",
        "n3 is a perfume: similarity to the rest only 0.20–0.25",
        "No edge reaches n3 → shortcut = {1, 2, 4}",
        "Yelp case: a vet clinic among bars & restaurants is dropped the same way",
    ], size=15)
    add_text(s, MARGIN, CONTENT_Y + Inches(4.75), W, Inches(0.5),
             "Explicit, inspectable denoising: you can see exactly which clicks were removed and why.",
             size=15, italic=True)
    done(s)

    # --- 6. Paper A: results ----------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Results: consistent gains on four datasets", tag=A)
    add_picture_fit(s, figs["a_tab1_results"], MARGIN, CONTENT_Y, Inches(8.4), Inches(3.4),
                    caption="Table 1 of Paper A: overall performance (Recall / NDCG @10, @20)")
    add_picture_fit(s, figs["a_fig3_tau"], MARGIN, CONTENT_Y + Inches(3.5), Inches(4.1),
                    Inches(2.0), caption="Figure 3: threshold τ sensitivity")
    add_picture_fit(s, figs["a_fig6_effect"], MARGIN + Inches(4.3), CONTENT_Y + Inches(3.5),
                    Inches(4.1), Inches(2.0), caption="Figure 6: closer user–target, outliers pushed out")
    x0 = MARGIN + Inches(8.8)
    sw = SLIDE_W - MARGIN - x0
    add_stat(s, x0, CONTENT_Y, sw, "+13%", "best improvement over baselines (Beauty)")
    add_stat(s, x0, CONTENT_Y + Inches(1.6), sw, "+4.7%", "over denoising methods (DSAN, HSD, DCRec)")
    add_bullets(s, x0, CONTENT_Y + Inches(3.25), sw, Inches(2.4), [
        "Big gains on shopping data, small on Yelp",
        "Even a zero-shot filter helps",
    ], size=14)
    done(s)

    # --- 7. Paper B: problem ----------------------------------------------
    s = new_slide(prs)
    slide_title(s, "News interests are time-sensitive", tag=B,
                subtitle="Modeling Stage-wise Evolution of User Interests for News Recommendation · WWW '26")
    add_bullets(s, MARGIN, CONTENT_Y + Inches(0.3), Inches(6.3), Inches(4.5), [
        "News loses value fast: users want the latest events, trending topics, seasonal contexts",
        "Two kinds of interest: stable long-term reading habits + short-term, event-driven shifts",
        ("e.g. World Cup → sports spike; pandemic → health news spike", 1),
        "Static global graph (GCN): good collaborative signal, but outdated recommendations",
        "Sequential models (LSTUR, ...): track recency, but lose stable high-order signals",
        "Goal: keep long-term stability and stay sensitive to short-term change",
    ], size=17)
    x0 = MARGIN + Inches(6.8)
    cw = SLIDE_W - MARGIN - x0
    add_card(s, x0, CONTENT_Y + Inches(0.3), cw, Inches(1.85), "Static models",
             "One graph over all time → a user's 3-month-old interest weighs as much as "
             "yesterday's. Profiles become dominated by outdated clicks.",
             head_color=CORAL, body_size=14)
    add_card(s, x0, CONTENT_Y + Inches(2.4), cw, Inches(1.85), "Pure sequential models",
             "Focus on the latest clicks → over-react to temporary trends and miss the "
             "collaborative structure of the whole graph.",
             head_color=CORAL, body_size=14)
    add_text(s, x0, CONTENT_Y + Inches(4.5), cw, Inches(0.7),
             "Key insight: clicks inside one period are coherent; transitions across periods show the evolution.",
             size=14, italic=True)
    done(s)

    # --- 8. Paper B: framework --------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Framework: global graph + stage-wise subgraphs", tag=B)
    add_picture_fit(s, figs["b_fig1_framework"], MARGIN, CONTENT_Y, Inches(8.3), Inches(5.3),
                    caption="Figure 1 of Paper B: overall framework")
    x0 = MARGIN + Inches(8.6)
    add_bullets(s, x0, CONTENT_Y + Inches(0.2), SLIDE_W - MARGIN - x0, Inches(5.2), [
        "GPM: LightGCN on the full user–news graph → long-term embeddings e_g",
        "Temporal subgraphs: split clicks into T windows (e.g. 2 weeks), one subgraph G_t each",
        "Short-term branch: GCN on each G_t (initialised by e_g) + LSTM across stages",
        "Long-range branch: self-attention over all clicks up to stage t + position embedding",
        "Fusion: user = short-term + long-range + global; score = σ(u·i)",
    ], size=15)
    done(s)

    # --- 9. Paper B: training objective ------------------------------------
    s = new_slide(prs)
    slide_title(s, "Training: one prediction loss, two regularisers", tag=B)
    cw = (W - Inches(0.6)) / 3
    for i, (h, b) in enumerate([
        ("Prediction  (λt = 0.1)",
         "Binary cross-entropy on clicks inside each time segment, with 4 sampled "
         "negatives per positive click."),
        ("Consistency  (λcl = 0.01)",
         "Contrastive (InfoNCE, τ = 0.1): a user's stage embedding should stay close to "
         "their own global embedding and far from other users'."),
        ("Smoothness  (λsl = 0.01)",
         "‖e_u^t − e_u^(t−1)‖² keeps adjacent stages close: interests drift gradually, "
         "not abruptly."),
    ]):
        add_card(s, MARGIN + i * (cw + Inches(0.3)), CONTENT_Y + Inches(0.2), cw,
                 Inches(2.05), h, b, head_color=CORAL)
    add_text(s, MARGIN, CONTENT_Y + Inches(2.7), W, Inches(0.6),
             "L  =  Σt λt·Lt  +  λcl·Lcl  +  λsl·Lsl  +  β‖Θ‖²",
             size=24, bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER)
    add_bullets(s, MARGIN, CONTENT_Y + Inches(3.7), W, Inches(1.6), [
        "Datasets: Adressa-large (3.6M users, 35M clicks, 3 months) and MIND-large (1M users, 24M clicks, 6 weeks)",
        "Metrics: AUC, MRR, nDCG@5/10 against 7 baselines (NRMS, NAML, NPA, LightGCN, CNE-SUE, TCCM, CROWN)",
    ], size=15)
    done(s)

    # --- 10. Paper B: results -----------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Results: large gains, especially on Adressa", tag=B)
    add_picture_fit(s, figs["b_tab2_results"], MARGIN, CONTENT_Y, Inches(8.2), Inches(3.3),
                    caption="Table 2 of Paper B: comparison with baselines")
    add_picture_fit(s, figs["b_tab3_ablation"], MARGIN, CONTENT_Y + Inches(3.45),
                    Inches(4.0), Inches(2.0), caption="Table 3: ablation (Adressa)")
    add_picture_fit(s, figs["b_fig2_window"], MARGIN + Inches(4.2), CONTENT_Y + Inches(3.45),
                    Inches(4.0), Inches(2.0), caption="Figure 2: window size")
    x0 = MARGIN + Inches(8.6)
    sw = SLIDE_W - MARGIN - x0
    add_stat(s, x0, CONTENT_Y, sw, "0.80", "AUC on Adressa vs 0.66 for best baseline (CROWN)")
    add_stat(s, x0, CONTENT_Y + Inches(1.6), sw, "0.58", "AUC on MIND vs 0.55 (smaller gain)")
    add_bullets(s, x0, CONTENT_Y + Inches(3.25), sw, Inches(2.4), [
        "Removing any module hurts a lot",
        "2-week window is best: shorter = too sparse, longer = blurs shifts",
    ], size=14)
    done(s)

    # --- 11. Side-by-side comparison ---------------------------------------
    s = new_slide(prs)
    slide_title(s, "Side by side")
    rows = [
        ("", "Paper A · LISRec", "Paper B · Stage-wise news rec"),
        ("Assumption challenged", "History is clean", "History is static"),
        ("Main signal", "Item text semantics", "Time stamps + collaborative graph"),
        ("Role of the graph", "Per-user item-similarity graph", "Global + per-stage user–news graphs"),
        ("Core mechanism", "Largest connected component as filter", "GCN + LSTM + self-attention"),
        ("Domain / data", "E-commerce & POI: Amazon ×3, Yelp", "News: Adressa, MIND (24–35M clicks)"),
        ("Metrics", "Recall, NDCG @10/20", "AUC, MRR, nDCG@5/10"),
        ("Plug-in?", "Yes, pre-filter for any recommender", "No, end-to-end model"),
    ]
    col_w = [Inches(2.9), Inches(4.62), Inches(4.61)]
    tbl = s.shapes.add_table(len(rows), 3, MARGIN, CONTENT_Y + Inches(0.1),
                             sum(col_w, Emu(0)), Inches(0.6) * len(rows)).table
    for j, cw_ in enumerate(col_w):
        tbl.columns[j].width = cw_
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(0.6)
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = WHITE if j == 0 else INK
            else:
                cell.fill.fore_color.rgb = TINT if j == 0 else WHITE
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.15)
            r = cell.text_frame.paragraphs[0].add_run()
            r.text = val
            _set_font(r, 15, WHITE if i == 0 else INK, bold=(i == 0 or j == 0))
    done(s)

    # --- 12. Summary & discussion ------------------------------------------
    s = new_slide(prs)
    slide_title(s, "Summary & discussion")
    half = (W - Inches(0.4)) / 2
    for x, head, items in [
        (MARGIN, "Takeaways", [
            "A: Split history by meaning - keep the semantically coherent core, drop outlier clicks",
            "B: Split history by time - a global graph for stable taste, stage subgraphs + LSTM / attention for drift",
            "Both: restructure the history before modelling instead of feeding raw logs",
            "Both report clear gains; A is plug-and-play and interpretable",
        ]),
        (MARGIN + half + Inches(0.4), "Open questions", [
            "A: Largest component keeps one dominant interest - what about users with several genuine interests?",
            "B: Several baselines score AUC ≈ 0.50 on Adressa (near random) - is the setup fair to them?",
            "B: Fixed windows; MIND has no real time stamps, so its stages are pseudo-periods",
            "Combine them? Denoise with A, then model stage-wise drift with B",
        ]),
    ]:
        add_box(s, x, CONTENT_Y, half, Inches(3.9), TINT)
        add_text(s, x + Inches(0.3), CONTENT_Y + Inches(0.2), half - Inches(0.6), Inches(0.4),
                 head, size=20, bold=True, font=HEAD_FONT)
        add_bullets(s, x + Inches(0.3), CONTENT_Y + Inches(0.8), half - Inches(0.6),
                    Inches(3.0), items, size=15)
    add_box(s, MARGIN, CONTENT_Y + Inches(4.2), W, Inches(1.0), INK)
    add_text(s, MARGIN + Inches(0.4), CONTENT_Y + Inches(4.2), W - Inches(0.8), Inches(1.0),
             "Better recommendations start with a better view of the history: "
             "which clicks really count (A) and when an interest matters (B).",
             size=18, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    done(s)

    # --- 13. Thanks ---------------------------------------------------------
    s = new_slide(prs, dark=True)
    add_text(s, MARGIN, Inches(2.3), W, Inches(1.0), "Thank you  ·  Q & A",
             size=44, color=WHITE, bold=True, font=HEAD_FONT)
    for i, (tag, ref) in enumerate([
        (A, "Xin et al. LISRec: Modeling User Preferences with Learned Item Shortcuts for Sequential "
            "Recommendation. KDD '26. doi:10.1145/3770854.3780337 · github.com/NEUIR/LISRec"),
        (B, "Cheng et al. Modeling Stage-wise Evolution of User Interests for News Recommendation. "
            "WWW '26. arXiv:2603.10471 · github.com/yikeJin/SEIN"),
    ]):
        y = Inches(4.0) + i * Inches(1.0)
        add_chip(s, MARGIN, y, tag[0], tag[1], d=Inches(0.5), size=16)
        add_text(s, MARGIN + Inches(0.75), y - Inches(0.02), W - Inches(0.75), Inches(0.8),
                 ref, size=14, color=LIGHT)
    done(s, numbered=False)

    for i, s in enumerate(slides, start=1):
        if i in notes:
            s.notes_slide.notes_text_frame.text = notes[i]

    prs.save(OUT)
    return len(slides)


if __name__ == "__main__":
    figs = crop_figures()
    print(f"Cropped {len(figs)} figures into {FIG_DIR}")
    notes = load_notes()
    print(f"Loaded speaker notes for {len(notes)} slides from {SCRIPT.name}")
    count = build(figs, notes)
    print(f"Saved {count} slides to {OUT}")
