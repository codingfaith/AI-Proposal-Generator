import math
import os
import re
from copy import deepcopy
from datetime import date
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

try:
    from PIL import ImageFont
except ImportError:  # python-pptx depends on Pillow, so this is just a safeguard
    ImageFont = None

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent

TEMPLATE_PATH = PROJECT_DIR / "assets" / "template.pptx"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PLACEHOLDER_RE = re.compile(r"\{\{[A-Za-z0-9_]+\}\}")

# ==================== OVERFLOW SETTINGS ====================

SPLIT_LONG_CONTENT = True    # spill extra text onto duplicated slides
ADD_CONT_SUFFIX = True       # "Objectives" -> "Objectives (cont.)" on extra slides
CONT_SUFFIX = " (cont.)"
MAX_EXTRA_SLIDES = 8         # safety cap per template slide (extra text beyond it is merged)
DISABLE_AUTOFIT = True       # turn off shrink/resize so layout matches our measurements
DEBUG_LAYOUT = True          # print what was detected for each content box

FILL_FACTOR = 0.92           # use at most 92% of the box height (safety margin)
DEFAULT_FONT_PT = 18         # only used if no font size is found anywhere
DEFAULT_LINE_EM = 1.2        # line height / font size, if font metrics unavailable
AVG_CHAR_WIDTH_EM = 0.5      # only used if no font file can be found

# Manual calibration per placeholder, if detection is ever off. Example:
#   {"{{OBJECTIVES}}": {"font_pt": 14, "max_lines": 12}}
# Keys: font_pt, max_lines, fill_factor
PLACEHOLDER_OVERRIDES = {}

# Template placeholder -> key in the proposal JSON
CONTENT_FIELDS = {
    "{{BACKGROUND}}": "background",
    "{{OBJECTIVES}}": "objectives",
    "{{UNIVERSE}}": "universe",
    "{{METHODOLOGY}}": "methodology",
    "{{STUDY_TYPE}}": "study_type",
    "{{INDUSTRY}}": "industry",
    "{{QUALITY_CONTROLS}}": "quality_controls",
    "{{TIMELINE}}": "timeline",
    "{{DELIVERABLES}}": "deliverables",
    "{{COSTING}}": "costing_assumptions",
    "{{RISKS}}": "risks",
    "{{ASSUMPTIONS}}": "assumptions",
    "{{VALUE_PROPOSITION}}": "value_proposition",
    "{{RECOMMENDED_APPROACH}}": "recommended_approach",
}

EMU_PER_PT = 12700
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ==================== CONTENT FORMATTING ====================

def _item_to_text(item) -> str:
    """Turn one list item into text. Handles dicts the model may return."""
    if isinstance(item, dict):
        return "; ".join(
            f"{str(k).replace('_', ' ').title()}: {v}" for k, v in item.items()
        )
    return str(item).strip()


def to_units(value) -> list:
    """Break a value into pieces (bullets or paragraphs) that can be paginated."""
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [f"• {_item_to_text(i)}" for i in value if i not in (None, "")]
    if isinstance(value, dict):
        return [f"• {str(k).replace('_', ' ').title()}: {v}" for k, v in value.items()]
    return [p.strip() for p in str(value).split("\n") if p.strip()]


def format_content(value) -> str:
    return "\n".join(to_units(value))


# ==================== FONT MEASUREMENT ====================

FONT_ALIASES = {
    "calibri": ["calibri.ttf", "carlito-regular.ttf"],
    "calibri light": ["calibril.ttf", "carlito-regular.ttf"],
    "arial": ["arial.ttf", "liberationsans-regular.ttf"],
    "helvetica": ["arial.ttf", "liberationsans-regular.ttf", "helvetica.ttc"],
    "times new roman": ["times.ttf", "liberationserif-regular.ttf"],
    "cambria": ["cambria.ttc", "caladea-regular.ttf"],
    "verdana": ["verdana.ttf", "dejavusans.ttf"],
    "tahoma": ["tahoma.ttf"],
    "segoe ui": ["segoeui.ttf"],
    "georgia": ["georgia.ttf"],
    "trebuchet ms": ["trebuc.ttf"],
    "courier new": ["cour.ttf", "liberationmono-regular.ttf"],
}
FALLBACK_FONT_FILES = [
    "calibri.ttf", "carlito-regular.ttf", "arial.ttf",
    "liberationsans-regular.ttf", "dejavusans.ttf",
]

_font_index = None
_measurers = {}
# Open fonts designed to match the width of common Microsoft fonts
METRIC_COMPATIBLE = {"carlito-regular.ttf", "liberationsans-regular.ttf",
                     "liberationserif-regular.ttf", "caladea-regular.ttf", "liberationmono-regular.ttf"}


def _build_font_index() -> dict:
    dirs = [
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
        os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"),
        "/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
        "/usr/share/fonts", "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts"),
    ]
    index = {}
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for name in files:
                if name.lower().endswith((".ttf", ".otf", ".ttc")):
                    index.setdefault(name.lower(), os.path.join(root, name))
    return index


def find_font_file(typeface):
    """Return (path, exact_match) for a typeface name, or (None, False)."""
    global _font_index
    if _font_index is None:
        _font_index = _build_font_index()

    name = (typeface or "").strip().lower()
    candidates = list(FONT_ALIASES.get(name, []))
    if name:
        candidates.insert(0, name.replace(" ", "") + ".ttf")

    for i, candidate in enumerate(candidates):
        if candidate in _font_index:
            exact = i == 0 or candidate == candidates[0]
            return _font_index[candidate], exact
    for candidate in FALLBACK_FONT_FILES:
        if candidate in _font_index:
            return _font_index[candidate], False
    return None, False


class FontMeasurer:
    SCALE = 100  # measure at a large size for precision, then scale

    def __init__(self, typeface):
        self.typeface = typeface
        self.path, self.exact = find_font_file(typeface)
        self.font = None
        self.line_em = DEFAULT_LINE_EM
        if ImageFont is not None and self.path:
            try:
                self.font = ImageFont.truetype(self.path, self.SCALE)
                ascent, descent = self.font.getmetrics()
                self.line_em = (ascent + descent) / self.SCALE
            except Exception:
                self.font = None

    def width(self, text: str, pt: float) -> float:
        if self.font is None:
            return len(text) * pt * AVG_CHAR_WIDTH_EM
        return self.font.getlength(text) * pt / self.SCALE

    def describe(self) -> str:
        if self.font is None:
            return "estimated widths (no font file found)"
        label = Path(self.path).name
        if self.exact:
            return label
        if label.lower() in METRIC_COMPATIBLE:
            return f"{label} (same metrics as '{self.typeface}')"
        return f"{label} (approximation for '{self.typeface}')"


def get_measurer(typeface) -> FontMeasurer:
    key = (typeface or "").lower()
    if key not in _measurers:
        _measurers[key] = FontMeasurer(typeface)
    return _measurers[key]


# ==================== STYLE INHERITANCE ====================
# Font size, line spacing and indents are usually NOT set on the text itself.
# They come from the shape, its layout, the slide master or the presentation
# defaults, so we walk that chain the same way PowerPoint does.

def _level_el(container, level):
    return None if container is None else container.find(qn(f"a:lvl{level + 1}pPr"))


def _style_kind(shape) -> str:
    if shape.is_placeholder:
        ptype = shape.placeholder_format.type
        if ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE,
                     PP_PLACEHOLDER.VERTICAL_TITLE):
            return "p:titleStyle"
        if ptype in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER,
                     PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.HEADER):
            return "p:otherStyle"
        return "p:bodyStyle"
    return "p:otherStyle"


def style_chain(prs, slide, shape, paragraph) -> list:
    """Style elements for this paragraph, highest priority first."""
    level = paragraph.level
    chain = [paragraph._p.pPr]

    node = shape
    for _ in range(4):  # shape -> layout placeholder -> master placeholder
        if node is None:
            break
        tx_body = node._element.find(qn("p:txBody"))
        lst = tx_body.find(qn("a:lstStyle")) if tx_body is not None else None
        chain.append(_level_el(lst, level))
        try:
            node = node._base_placeholder if node.is_placeholder else None
        except Exception:
            node = None

    master = slide.slide_layout.slide_master
    tx_styles = master._element.find(qn("p:txStyles"))
    if tx_styles is not None:
        chain.append(_level_el(tx_styles.find(qn(_style_kind(shape))), level))
    chain.append(_level_el(prs._element.find(qn("p:defaultTextStyle")), level))

    return [el for el in chain if el is not None]


def _chain_attr(chain, path, attr):
    for el in chain:
        target = el if path is None else el.find(path)
        if target is not None and target.get(attr) is not None:
            return target.get(attr)
    return None


def _theme_fonts(slide):
    try:
        theme_part = slide.slide_layout.slide_master.part.part_related_by(RT.THEME)
        root = etree.fromstring(theme_part.blob)
        ns = {"a": A_NS}
        major = root.find(".//a:majorFont/a:latin", ns).get("typeface")
        minor = root.find(".//a:minorFont/a:latin", ns).get("typeface")
        return major, minor
    except Exception:
        return None, None


class Box:
    """Everything needed to estimate how much text fits in one text box."""

    def __init__(self, placeholder, font_pt, measurer, line_pt, width_pt, height_pt, max_h):
        self.placeholder = placeholder
        self.font_pt = font_pt
        self.measurer = measurer
        self.line_pt = line_pt
        self.width_pt = width_pt
        self.height_pt = height_pt
        self.max_h = max_h

    def lines(self, text: str) -> int:
        """Word-wrap simulation using real font widths."""
        total = 0
        for paragraph in text.split("\n"):
            line, count = "", 1
            for word in paragraph.split(" "):
                candidate = word if not line else f"{line} {word}"
                if not line or self.measurer.width(candidate, self.font_pt) <= self.width_pt:
                    line = candidate
                else:
                    count += 1
                    line = word
            # a single word wider than the box wraps onto extra lines
            overflow = self.measurer.width(line, self.font_pt) / self.width_pt
            count += max(0, math.ceil(overflow) - 1)
            total += count
        return total

    def height(self, text: str) -> float:
        return self.lines(text) * self.line_pt


def analyze_box(prs, slide, shape, paragraph, placeholder):
    """Build a Box for the text box holding `placeholder`, or None if unknown."""
    overrides = PLACEHOLDER_OVERRIDES.get(placeholder, {})
    tf = shape.text_frame
    run = paragraph.runs[0] if paragraph.runs else None
    run_rpr = run._r.rPr if run is not None else None
    chain = style_chain(prs, slide, shape, paragraph)

    # Font size
    size = run_rpr.get("sz") if run_rpr is not None else None
    size = size or _chain_attr(chain, qn("a:defRPr"), "sz")
    font_pt = overrides.get("font_pt") or (int(size) / 100 if size else DEFAULT_FONT_PT)

    # Typeface (theme fonts like "+mn-lt" are resolved through the theme)
    typeface = None
    if run_rpr is not None and run_rpr.find(qn("a:latin")) is not None:
        typeface = run_rpr.find(qn("a:latin")).get("typeface")
    typeface = typeface or _chain_attr(chain, f"{qn('a:defRPr')}/{qn('a:latin')}", "typeface")
    if typeface in ("+mn-lt", "+mj-lt") or not typeface:
        major, minor = _theme_fonts(slide)
        typeface = major if typeface == "+mj-lt" else minor
    measurer = get_measurer(typeface)

    # Line height (percentage or exact spacing)
    exact = _chain_attr(chain, f"{qn('a:lnSpc')}/{qn('a:spcPts')}", "val")
    pct = _chain_attr(chain, f"{qn('a:lnSpc')}/{qn('a:spcPct')}", "val")
    if exact:
        line_pt = int(exact) / 100
    else:
        line_pt = font_pt * measurer.line_em * (int(pct) / 100000 if pct else 1.0)

    # Box size, less text insets and the paragraph's left indent
    try:
        indent = int(_chain_attr(chain, None, "marL") or 0)
        width_pt = (shape.width - tf.margin_left - tf.margin_right - indent) / EMU_PER_PT
        height_pt = (shape.height - tf.margin_top - tf.margin_bottom) / EMU_PER_PT
    except TypeError:
        return None
    if width_pt <= 0 or height_pt <= 0:
        return None

    if overrides.get("max_lines"):
        max_h = overrides["max_lines"] * line_pt
    else:
        max_h = height_pt * overrides.get("fill_factor", FILL_FACTOR)

    return Box(placeholder, font_pt, measurer, line_pt, width_pt, height_pt, max_h)


# ==================== PAGINATION ====================

def paginate(units, box: Box) -> list:
    """Pack units into chunks that each fit in the box. Returns a list of strings."""
    # A single paragraph taller than a whole slide gets split at sentence boundaries
    expanded = []
    for unit in units:
        if box.height(unit) <= box.max_h:
            expanded.append(unit)
            continue
        piece = ""
        for sentence in re.split(r"(?<=[.!?])\s+", unit):
            candidate = f"{piece} {sentence}".strip()
            if piece and box.height(candidate) > box.max_h:
                expanded.append(piece)
                piece = sentence
            else:
                piece = candidate
        if piece:
            expanded.append(piece)

    chunks, current, used = [], [], 0.0
    for unit in expanded:
        needed = box.height(unit)
        if current and used + needed > box.max_h:
            chunks.append("\n".join(current))
            current, used = [], 0.0
        current.append(unit)
        used += needed
    if current:
        chunks.append("\n".join(current))

    if len(chunks) > MAX_EXTRA_SLIDES + 1:
        print(f"Warning: content needs {len(chunks)} slides; capping at {MAX_EXTRA_SLIDES + 1}")
        chunks = chunks[:MAX_EXTRA_SLIDES] + ["\n".join(chunks[MAX_EXTRA_SLIDES:])]

    return chunks or [""]


def _paragraph_with(shape, placeholder):
    for paragraph in shape.text_frame.paragraphs:
        if placeholder in "".join(r.text for r in paragraph.runs):
            return paragraph
    return None


def plan_slide_content(prs, slide, content: dict) -> list:
    """
    Return one {placeholder: text} dict per output slide for this template slide.
    One dict means no overflow; more means the slide must be duplicated.
    Also switches off autofit on content boxes so the real layout matches the
    measurements (done here because duplicates are copied from this slide).
    """
    full = {ph: format_content(value) for ph, value in content.items()}
    chunked = {}

    if SPLIT_LONG_CONTENT:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for ph in sorted(set(PLACEHOLDER_RE.findall(shape.text_frame.text))):
                if ph not in content or ph in chunked:
                    continue
                paragraph = _paragraph_with(shape, ph)
                box = analyze_box(prs, slide, shape, paragraph, ph) if paragraph else None
                if box is None:
                    continue

                if DISABLE_AUTOFIT:
                    shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
                    shape.text_frame.word_wrap = True

                chunks = paginate(to_units(content[ph]), box)
                if DEBUG_LAYOUT:
                    print(
                        f"[layout] {ph}: {box.font_pt:g}pt, {box.measurer.describe()}, "
                        f"box {box.width_pt:.0f}x{box.height_pt:.0f}pt, "
                        f"line {box.line_pt:.1f}pt, fits ~{box.max_h / box.line_pt:.0f} lines "
                        f"-> {len(chunks)} slide(s)"
                    )
                if len(chunks) > 1:
                    chunked[ph] = chunks

    count = max((len(c) for c in chunked.values()), default=1)
    plans = []
    for k in range(count):
        plan = dict(full)
        for ph, chunks in chunked.items():
            plan[ph] = chunks[k] if k < len(chunks) else ""
        plans.append(plan)
    return plans


# ==================== SLIDE DUPLICATION ====================

def duplicate_slide(prs, source):
    """Copy a slide (shapes, images, background). python-pptx has no built-in for this."""
    dest = prs.slides.add_slide(source.slide_layout)

    # add_slide() populates layout placeholders; clear them before copying
    for shape in list(dest.shapes):
        shape._element.getparent().remove(shape._element)

    # Carry over images/media/hyperlinks and remember the new relationship ids
    rid_map = {}
    for rel in source.part.rels.values():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            rid_map[rel.rId] = dest.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rel.rId] = dest.part.rels.get_or_add(rel.reltype, rel.target_part)

    for shape in source.shapes:
        dest.shapes._spTree.insert_element_before(deepcopy(shape._element), "p:extLst")

    for element in dest.shapes._spTree.iter():
        for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
            old = element.get(attr)
            if old in rid_map:
                element.set(attr, rid_map[old])

    background = source._element.cSld.find(qn("p:bg"))
    if background is not None:
        dest._element.cSld.insert(0, deepcopy(background))

    return dest


def move_slide_after(prs, slide, after):
    """Reposition `slide` so it sits directly after `after`."""
    slide_id_list = prs.slides._sldIdLst
    moving = next(e for e in slide_id_list if int(e.id) == slide.slide_id)
    slide_id_list.remove(moving)
    anchor = next(e for e in slide_id_list if int(e.id) == after.slide_id)
    anchor.addnext(moving)


def add_cont_suffix(slide):
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        return
    for paragraph in reversed(title.text_frame.paragraphs):
        if paragraph.runs:
            paragraph.runs[-1].text += CONT_SUFFIX
            return


# ==================== PLACEHOLDER REPLACEMENT ====================

def iter_text_frames(shapes):
    """Yield every text frame, including ones inside groups and tables."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_text_frames(shape.shapes)
        elif shape.has_text_frame:
            yield shape.text_frame
        elif getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame


def replace_in_paragraph(paragraph, replacements, unresolved: set):
    """
    Replace placeholders in one paragraph.

    PowerPoint often splits a placeholder like {{CLIENT_NAME}} across several
    runs, so we work on the paragraph's combined text and write the result back
    into the first run (keeping its formatting).
    """
    runs = paragraph.runs
    if not runs:
        return

    full_text = "".join(run.text for run in runs)
    if "{{" not in full_text:
        return

    new_text = full_text
    for placeholder, value in replacements.items():
        new_text = new_text.replace(placeholder, value)

    # Anything still in {{...}} form is a typo or a field we don't fill
    unresolved.update(PLACEHOLDER_RE.findall(new_text))
    new_text = PLACEHOLDER_RE.sub("", new_text)

    if new_text == full_text:
        return

    lines = new_text.split("\n")
    first = runs[0]
    first.text = lines[0]

    for run in runs[1:]:
        run._r.getparent().remove(run._r)

    # Extra lines become line breaks, with the first run's formatting copied
    for line in lines[1:]:
        paragraph.add_line_break()
        new_run = paragraph.add_run()
        new_run.text = line
        if first._r.rPr is not None:
            new_run._r.insert(0, deepcopy(first._r.rPr))


def replace_on_slide(slide, replacements, unresolved: set):
    for text_frame in iter_text_frames(slide.shapes):
        for paragraph in text_frame.paragraphs:
            replace_in_paragraph(paragraph, replacements, unresolved)


# ==================== BUILD ====================

def build_ppt(proposal: dict, client_name: str, ppt_filename: str):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")

    prs = Presentation(str(TEMPLATE_PATH))

    base = {
        "{{CLIENT_NAME}}": client_name,
        "{{DATE}}": date.today().strftime("%d %B %Y"),
    }
    content = {ph: proposal.get(key) for ph, key in CONTENT_FIELDS.items()}
    unresolved = set()

    # Snapshot of the template's slides; duplicates are added after this list is taken
    for source in list(prs.slides):
        plans = plan_slide_content(prs, source, content)

        # Duplicate from the untouched template slide BEFORE replacing anything
        slides = [source]
        for _ in plans[1:]:
            duplicate = duplicate_slide(prs, source)
            move_slide_after(prs, duplicate, slides[-1])
            slides.append(duplicate)

        for index, (slide, plan) in enumerate(zip(slides, plans)):
            replace_on_slide(slide, {**base, **plan}, unresolved)
            if index > 0 and ADD_CONT_SUFFIX:
                add_cont_suffix(slide)

    if unresolved:
        print(f"Warning: unknown placeholders in template (blanked out): {sorted(unresolved)}")

    output_path = OUTPUT_DIR / ppt_filename
    prs.save(output_path)
    print(f"PPT saved: {output_path} ({len(prs.slides)} slides)")
    return str(output_path)