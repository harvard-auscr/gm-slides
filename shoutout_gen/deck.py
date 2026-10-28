"""Write the PPTX: one textbox per shout-out, one click-triggered Appear each.

The XML emitted here is a faithful copy of what Google Slides exported for the
club's own shout-out slides (see output/scraped), so the generated slide can
be pasted into a GM deck and behave identically:

  textbox  ``<p:sp>`` with ``txBox=1``, ``wrap="square"``, ``<a:spAutoFit/>``,
           91425-EMU insets, Roboto set on latin/ea/cs/sym, colour ``dk1``.
  timing   main sequence of ``clickEffect`` nodes, preset Appear
           (presetID=1, presetClass=entr) that sets ``style.visibility`` to
           ``visible`` on the target shape. Click order == list order.

Only this module touches python-pptx / lxml; layout and measurement stay pure.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .metrics import BOX_INSET_EMU

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "template" / "shoutouts_template.pptx"
LAYOUT_NAME = "ONE_COLUMN_TEXT"
FONT_NAME = "Roboto"

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


@dataclass(frozen=True)
class PlacedShoutout:
    """A shout-out with its final geometry. ``slide`` is 0-based; EMU throughout."""

    text: str
    slide: int
    x_emu: int
    y_emu: int
    width_emu: int
    height_emu: int


# --------------------------------------------------------------------------- #
# Public interface
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class QotwFinale:
    """The question of the week, revealed after the last shout-out: an intro pane, then the chart."""

    question: str
    chart_png: Path


def write_deck(
    shoutouts: list[PlacedShoutout],
    out_path: Path,
    font_size_pt: float,
    template_path: Path = TEMPLATE_PATH,
    font_name: str = FONT_NAME,
    qotw: QotwFinale | None = None,
) -> Path:
    """Render ``shoutouts`` (in click order) into ``out_path`` and return it.

    Slides are created on demand from the template's ONE_COLUMN_TEXT layout;
    slide N of the output holds every shout-out whose ``slide == N``. With
    ``qotw``, the last slide gets two more clicks after its last shout-out: a
    full-slide "QOTW: <question>" pane covering the shout-outs, then the chart
    over that.
    """
    prs = Presentation(str(template_path))
    layout = _layout(prs)
    slide_count = max(max((s.slide for s in shoutouts), default=-1) + 1, 1 if qotw else 0)
    slides = [_blank_slide(prs, layout) for _ in range(slide_count)]

    # Shape ids per slide, in click order, for the animation sequence.
    click_order: dict[int, list[int]] = {i: [] for i in range(slide_count)}
    for item in shoutouts:
        shape = _add_textbox(slides[item.slide], item, font_size_pt, font_name)
        click_order[item.slide].append(shape.shape_id)
    if qotw is not None:
        last = slide_count - 1
        pane = _add_qotw_pane(slides[last], qotw.question, prs.slide_width, prs.slide_height, font_name)
        chart = slides[last].shapes.add_picture(str(qotw.chart_png), 0, 0, prs.slide_width, prs.slide_height)
        pane.name, chart.name = QOTW_PANE_NAME, QOTW_CHART_NAME
        click_order[last] += [pane.shape_id, chart.shape_id]
    for idx, slide in enumerate(slides):
        _add_click_appear_timing(slide, click_order[idx])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def write_image_slide(image_path: Path, out_path: Path, template_path: Path = TEMPLATE_PATH) -> Path:
    """One slide in the club theme holding ``image_path`` full-bleed (the image should be 16:9)."""
    prs = Presentation(str(template_path))
    slide = _blank_slide(prs, _layout(prs))
    slide.shapes.add_picture(str(image_path), 0, 0, prs.slide_width, prs.slide_height)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def write_reveal_slide(
    first_png: Path, revealed_png: Path, out_path: Path, template_path: Path = TEMPLATE_PATH,
    names: tuple[str, str] = ("first", "revealed"),
) -> Path:
    """One slide: ``first_png`` full-bleed, then ``revealed_png`` full-bleed over it on the
    first click (the same click-Appear as the QOTW chart). Both images should be 16:9."""
    prs = Presentation(str(template_path))
    slide = _blank_slide(prs, _layout(prs))
    first = slide.shapes.add_picture(str(first_png), 0, 0, prs.slide_width, prs.slide_height)
    revealed = slide.shapes.add_picture(str(revealed_png), 0, 0, prs.slide_width, prs.slide_height)
    first.name, revealed.name = names
    _add_click_appear_timing(slide, [revealed.shape_id])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def slide_size(template_path: Path = TEMPLATE_PATH) -> tuple[int, int]:
    """(width, height) in EMU of the template's page, so the packer targets the real canvas."""
    prs = Presentation(str(template_path))
    return int(prs.slide_width), int(prs.slide_height)


# --------------------------------------------------------------------------- #
# Private mechanics
# --------------------------------------------------------------------------- #
def _layout(prs):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == LAYOUT_NAME:
                return layout
    raise ValueError(f"template has no {LAYOUT_NAME} layout")


def _blank_slide(prs, layout):
    """New slide from the layout with its placeholders removed (originals have none)."""
    slide = prs.slides.add_slide(layout)
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)
    return slide


def _add_textbox(slide, item: PlacedShoutout, font_size_pt: float, font_name: str):
    shape = slide.shapes.add_textbox(
        Emu(item.x_emu), Emu(item.y_emu), Emu(item.width_emu), Emu(item.height_emu)
    )
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(BOX_INSET_EMU)

    paragraphs = item.text.split("\n")
    for i, para_text in enumerate(paragraphs):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.space_before = paragraph.space_after = Pt(0)
        run = paragraph.add_run()
        run.text = para_text
        run.font.size = Pt(font_size_pt)
        _set_run_font_and_colour(run, font_name)
    return shape


QOTW_PANE_NAME = "QOTW pane"  # shape names, so verify_render.py can tell the finale from the shout-outs
QOTW_CHART_NAME = "QOTW chart"
QOTW_PANE_PT = 40
HAUSCR_RED = "AE3646"  # the club logo's red: the "QOTW:" label here and the bars of the QOTW chart
_QOTW_PANE_MARGIN_EMU = 914_400  # 1in either side, so a long question wraps instead of touching the edge


def _add_qotw_pane(slide, question: str, width: int, height: int, font_name: str):
    """Full-slide opaque white textbox, "QOTW:" (HAUSCR red, bold) + the question, centred both ways."""
    shape = slide.shapes.add_textbox(0, 0, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(_QOTW_PANE_MARGIN_EMU)
    paragraph = tf.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    for text, accent in (("QOTW: ", True), (question, False)):
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(QOTW_PANE_PT)
        _set_run_font_and_colour(run, font_name)
        if accent:
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(HAUSCR_RED)
    return shape


def _set_run_font_and_colour(run, font_name: str) -> None:
    """The typeface on every script slot + theme text colour, matching the corpus XML byte-for-byte."""
    rPr = run._r.get_or_add_rPr()
    fill = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr")).set("val", "dk1")
    for tag in ("a:latin", "a:ea", "a:cs", "a:sym"):
        etree.SubElement(rPr, qn(tag)).set("typeface", font_name)


def _add_click_appear_timing(slide, shape_ids_in_click_order: list[int]) -> None:
    """Append a <p:timing> main sequence with one click-Appear per shape id."""
    if not shape_ids_in_click_order:
        return
    ids = _IdCounter()
    timing = etree.SubElement(slide._element, qn("p:timing"))
    tn_lst = etree.SubElement(timing, qn("p:tnLst"))
    root_par = etree.SubElement(tn_lst, qn("p:par"))
    root = _ctn(root_par, ids, dur="indefinite", nodeType="tmRoot", restart="never")
    root_children = etree.SubElement(root, qn("p:childTnLst"))
    seq = etree.SubElement(root_children, qn("p:seq"), concurrent="1", nextAc="seek")
    main = _ctn(seq, ids, dur="indefinite", nodeType="mainSeq")
    main_children = etree.SubElement(main, qn("p:childTnLst"))
    for shape_id in shape_ids_in_click_order:
        _append_click_appear(main_children, ids, shape_id)
    for list_tag, evt in (("p:prevCondLst", "onPrev"), ("p:nextCondLst", "onNext")):
        cond = etree.SubElement(etree.SubElement(seq, qn(list_tag)), qn("p:cond"), evt=evt)
        etree.SubElement(etree.SubElement(cond, qn("p:tgtEl")), qn("p:sldTgt"))


def _append_click_appear(parent, ids: "_IdCounter", shape_id: int) -> None:
    """One click step: par(indefinite) > par(delay 0) > par(clickEffect Appear) > set visibility."""
    click = _ctn(etree.SubElement(parent, qn("p:par")), ids, fill="hold")
    _start_cond(click, "indefinite")
    group = _ctn(etree.SubElement(etree.SubElement(click, qn("p:childTnLst")), qn("p:par")), ids, fill="hold")
    _start_cond(group, "0")
    effect = _ctn(
        etree.SubElement(etree.SubElement(group, qn("p:childTnLst")), qn("p:par")),
        ids, fill="hold", nodeType="clickEffect", presetClass="entr", presetID="1", presetSubtype="0",
    )
    _start_cond(effect, "0")
    set_el = etree.SubElement(etree.SubElement(effect, qn("p:childTnLst")), qn("p:set"))
    bhvr = etree.SubElement(set_el, qn("p:cBhvr"))
    _start_cond(_ctn(bhvr, ids, dur="1", fill="hold"), "0")
    etree.SubElement(etree.SubElement(bhvr, qn("p:tgtEl")), qn("p:spTgt")).set("spid", str(shape_id))
    etree.SubElement(etree.SubElement(bhvr, qn("p:attrNameLst")), qn("p:attrName")).text = "style.visibility"
    etree.SubElement(etree.SubElement(set_el, qn("p:to")), qn("p:strVal")).set("val", "visible")


class _IdCounter:
    """Unique ``id`` for every <p:cTn>; PowerPoint expects them distinct within a slide."""

    def __init__(self) -> None:
        self.next = 1

    def take(self) -> str:
        value, self.next = self.next, self.next + 1
        return str(value)


def _ctn(parent, ids: _IdCounter, **attrs: str):
    ctn = etree.SubElement(parent, qn("p:cTn"), id=ids.take())
    for key, value in attrs.items():
        ctn.set(key, value)
    return ctn


def _start_cond(ctn, delay: str) -> None:
    etree.SubElement(etree.SubElement(ctn, qn("p:stCondLst")), qn("p:cond")).set("delay", delay)
