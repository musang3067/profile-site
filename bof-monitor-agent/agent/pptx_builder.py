"""일일 카드뉴스 PPTX 생성기 (4:5 세로 카드, 1080x1350 비율)."""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

W, H = Inches(7.5), Inches(9.375)
FONT = "Malgun Gothic"
NAVY = RGBColor(0x10, 0x18, 0x2B)
COPPER = RGBColor(0xB0, 0x7A, 0x4F)
INK = RGBColor(0x1F, 0x24, 0x2E)
MUTED = RGBColor(0x6B, 0x72, 0x80)
LINE = RGBColor(0xDD, 0xDF, 0xE3)
PAPER = RGBColor(0xF7, 0xF6, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLANK_GREY = RGBColor(0xA0, 0xA4, 0xAB)
ROLE_COLOR = {"대상": COPPER, "현대차그룹": RGBColor(0x00, 0x2C, 0x5F), "경쟁": RGBColor(0x4B, 0x55, 0x63)}
LABEL_COLOR = {
    "신차공개": RGBColor(0xB0, 0x7A, 0x4F), "출시/판매개시": RGBColor(0x1E, 0x6B, 0x52),
    "신기술": RGBColor(0x2F, 0x4B, 0xA0), "가격변경": RGBColor(0x8A, 0x5A, 0x00),
    "페이스리프트/연식변경": RGBColor(0x5B, 0x4B, 0x8A), "리콜/품질": RGBColor(0xA3, 0x2D, 0x2D),
    "생산/단종": RGBColor(0x4B, 0x55, 0x63), "루머/스파이샷": RGBColor(0x80, 0x80, 0x80),
}


# ---------- primitives ----------
def _text(slide, x, y, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size, r.font.bold, r.font.name = Pt(size), bold, FONT
        r.font.color.rgb = color
    return tb


def _rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def _chip(slide, x, y, text, color, size=9):
    w = Inches(0.2 + 0.018 * size * sum(1.0 if ord(c) > 0x2000 else 0.6 for c in text))
    s = _rect(slide, x, y, w, Inches(0.26), color, MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = s.text_frame
    tf.word_wrap = False
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.name = Pt(size), True, FONT
    r.font.color.rgb = WHITE
    return w


def _base(prs, date, title, kicker=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, W, H, PAPER)
    _rect(s, 0, 0, W, Inches(1.05), NAVY)
    _text(s, Inches(0.4), Inches(0.16), Inches(5), Inches(0.3),
          kicker or "BOF · GENESIS DAILY", 9, True, COPPER)
    _text(s, Inches(0.4), Inches(0.42), Inches(6.7), Inches(0.55), title, 20, True, WHITE)
    _text(s, Inches(5.1), Inches(0.16), Inches(2.0), Inches(0.3), date, 9, False,
          RGBColor(0xC9, 0xCD, 0xD4), PP_ALIGN.RIGHT)
    _text(s, Inches(0.4), H - Inches(0.38), Inches(6.7), Inches(0.25),
          "수치: 제조사 공식 홈페이지 우선 · 미확인 항목은 '–' · SI 환산(원 단위 병기) · 가격 US MSRP",
          7, False, MUTED)
    return s


def _picture(slide, path, x, y, w, h, credit=None):
    if path and Path(path).is_file():
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(str(path), x, y, w, h)
        # center-crop to box ratio
        box, img = w / h, iw / ih
        if img > box:
            c = (1 - box / img) / 2
            pic.crop_left = pic.crop_right = c
        else:
            c = (1 - img / box) / 2
            pic.crop_top = pic.crop_bottom = c
    else:
        _rect(slide, x, y, w, h, RGBColor(0xE4, 0xE5, 0xE8))
        _text(slide, x, y, w, h, "사진 미확보", 10, False, BLANK_GREY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    if credit:
        _text(slide, x, y + h - Inches(0.22), w, Inches(0.22), f"© {credit}", 6.5, False,
              WHITE if path and Path(path).is_file() else MUTED, PP_ALIGN.RIGHT)


# ---------- slides ----------
def slide_cover(prs, brief):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, W, H, NAVY)
    _rect(s, Inches(0.5), Inches(1.2), Inches(0.9), Inches(0.06), COPPER)
    _text(s, Inches(0.5), Inches(1.4), Inches(6.5), Inches(0.4), "BOF · GENESIS DAILY BRIEF", 12, True, COPPER)
    _text(s, Inches(0.5), Inches(1.85), Inches(6.5), Inches(1.6),
          ["오늘의 프리미엄", "BOF·제네시스 신차 브리핑"], 30, True, WHITE)
    _text(s, Inches(0.5), Inches(3.6), Inches(6.5), Inches(0.4), brief["date"], 14, False,
          RGBColor(0xC9, 0xCD, 0xD4))
    y = Inches(4.6)
    for i, h in enumerate((brief.get("headlines") or [])[:3], 1):
        _text(s, Inches(0.5), y, Inches(0.5), Inches(0.6), f"{i:02d}", 18, True, COPPER)
        _text(s, Inches(1.1), y + Inches(0.03), Inches(5.9), Inches(0.9), h, 14, True, WHITE)
        y += Inches(1.0)
    _text(s, Inches(0.5), H - Inches(0.8), Inches(6.5), Inches(0.5),
          "대상: 제네시스 전 차종 + BOF SUV·픽업(전장 4,800~5,800mm)", 9, False, RGBColor(0x9A, 0xA1, 0xAD))


def slide_summary(prs, brief):
    issues = brief.get("issues") or []
    s = _base(prs, brief["date"], "오늘의 주요 이슈")
    if not issues:
        _text(s, Inches(0.4), Inches(1.5), Inches(6.7), Inches(1), "중요도 기준(70점)을 넘은 신규 이슈가 없습니다.", 14, False, MUTED)
        return
    y = Inches(1.35)
    for it in issues[:6]:
        _rect(s, Inches(0.4), y, Inches(6.7), Inches(1.1), WHITE)
        _chip(s, Inches(0.55), y + Inches(0.15), it.get("label", "기타"), LABEL_COLOR.get(it.get("label"), MUTED))
        _text(s, Inches(6.0), y + Inches(0.1), Inches(1.0), Inches(0.3), f"{it.get('score', '')}점", 10, True, COPPER, PP_ALIGN.RIGHT)
        _text(s, Inches(0.55), y + Inches(0.45), Inches(6.4), Inches(0.35), it["title"], 13, True)
        one = (it.get("summary") or [""])[0]
        _text(s, Inches(0.55), y + Inches(0.75), Inches(6.4), Inches(0.3), one, 9, False, MUTED)
        y += Inches(1.22)


def slide_issue(prs, brief, it):
    s = _base(prs, brief["date"], it["title"], kicker=f"ISSUE · {it.get('label', '')}")
    _picture(s, it.get("image"), Inches(0.4), Inches(1.3), Inches(6.7), Inches(3.75), it.get("image_credit"))
    x = Inches(0.4)
    x += _chip(s, x, Inches(5.25), it.get("label", "기타"), LABEL_COLOR.get(it.get("label"), MUTED)) + Inches(0.1)
    if it.get("official") is False:
        _chip(s, x, Inches(5.25), "미확인", RGBColor(0x80, 0x80, 0x80))
    _text(s, Inches(5.6), Inches(5.22), Inches(1.5), Inches(0.3), f"중요도 {it.get('score', '')}", 10, True, COPPER, PP_ALIGN.RIGHT)
    _text(s, Inches(0.4), Inches(5.65), Inches(6.7), Inches(1.6),
          [f"• {l}" for l in (it.get("summary") or [])[:3]], 12)
    if it.get("tech"):
        _rect(s, Inches(0.4), Inches(7.3), Inches(6.7), Inches(1.15), WHITE)
        _text(s, Inches(0.55), Inches(7.35), Inches(6.4), Inches(0.3), "신기술 · 특이사항", 9, True, COPPER)
        _text(s, Inches(0.55), Inches(7.62), Inches(6.4), Inches(0.8),
              [f"– {t}" for t in it["tech"][:3]], 9.5)
    src = " | ".join(it.get("sources") or [])[:220]
    _text(s, Inches(0.4), Inches(8.55), Inches(6.7), Inches(0.4), f"출처: {src}" if src else "출처: –", 7, False, MUTED)


def _roster(comp, vehicles):
    rows = [("대상", comp["target"])]
    rows += [("현대차그룹", v) for v in comp.get("hmg", [])]
    rows += [("경쟁", v) for v in comp.get("rivals", [])[:5]]
    return [(role, vehicles[vid]) for role, vid in rows if vid in vehicles]


def slide_photos(prs, brief, comp, roster):
    s = _base(prs, brief["date"], comp.get("title", "비교 차종"), kicker="COMPARE · 비교 차종")
    cw, ch, gap = Inches(3.27), Inches(1.55), Inches(0.16)
    for i, (role, v) in enumerate(roster[:8]):
        cx = Inches(0.4) + (i % 2) * (cw + gap)
        cy = Inches(1.3) + (i // 2) * (ch + Inches(0.42))
        _picture(s, v.get("image"), cx, cy, cw, ch, v.get("image_credit"))
        _chip(s, cx + Inches(0.08), cy + Inches(0.08), role, ROLE_COLOR[role], 8)
        name = v["name"] + (f" {v['trim']}" if v.get("trim") else "")
        _text(s, cx, cy + ch + Inches(0.02), cw, Inches(0.24), name, 10, True)
        meta = " · ".join(x for x in (v.get("platform"), v.get("body"), v.get("model_year") and f"MY{v['model_year']}") if x)
        _text(s, cx, cy + ch + Inches(0.2), cw, Inches(0.2), meta, 7.5, False, MUTED)
    if comp.get("reason"):
        _text(s, Inches(0.4), Inches(8.6), Inches(6.7), Inches(0.35), f"선정 기준: {comp['reason']}", 7.5, False, MUTED)


def _table(slide, roster, keys, y=Inches(1.35)):
    ncols = len(keys) + 1
    nrows = len(roster) + 1
    row_h = Inches(0.78)
    shape = slide.shapes.add_table(nrows, ncols, Inches(0.4), y, Inches(6.7), row_h * nrows)
    tbl = shape.table
    first = Inches(1.6)
    tbl.columns[0].width = first
    other = int((Inches(6.7) - first) / len(keys))
    for c in range(1, ncols):
        tbl.columns[c].width = Emu(other)

    def put(cell, lines, size, bold=False, color=INK, fill=WHITE, align=PP_ALIGN.CENTER):
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.03)
        tf = cell.text_frame
        tf.word_wrap = True
        for i, (txt, sz, b, col) in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = txt
            r.font.size, r.font.bold, r.font.name = Pt(sz), b, FONT
            r.font.color.rgb = col

    from .specs import FIELDS
    put(tbl.cell(0, 0), [("차종", 9, True, WHITE)], 9, fill=NAVY)
    for c, k in enumerate(keys, 1):
        put(tbl.cell(0, c), [(FIELDS[k][0], 8.5, True, WHITE)], 8.5, fill=NAVY)
    tbl.rows[0].height = Inches(0.45)
    for r, (role, v) in enumerate(roster, 1):
        tbl.rows[r].height = row_h
        fill = RGBColor(0xFB, 0xF4, 0xEE) if role == "대상" else WHITE
        put(tbl.cell(r, 0), [(role, 7, True, ROLE_COLOR[role]), (v["name"], 8.5, True, INK)],
            8.5, fill=fill, align=PP_ALIGN.LEFT)
        for c, k in enumerate(keys, 1):
            f = v["fields"][k]
            if f["value"] is None:
                lines = [("–", 11, False, BLANK_GREY)]
            elif f["si_unit"] and f["original"] and f["si_unit"] not in ("USD", "s", "seats") \
                    and f["unit"] and f["unit"].lower() not in ("mm", "kg", "kw", "nm", "n·m", "km/h", "km", "l"):
                lines = [(f["display"], 9, True, INK), (f"({f['original']})", 6.5, False, MUTED)]
            else:
                lines = [(f["display"], 8.5 if k in ("engine", "transmission", "drivetrain") else 9, True, INK)]
            put(tbl.cell(r, c), lines, 9, fill=fill)
    return tbl


def slide_spec_table(prs, brief, comp, roster, title, keys):
    s = _base(prs, brief["date"], title, kicker="COMPARE · 제원")
    _table(s, roster[:7], keys)


def slide_charts(prs, brief, comp, roster):
    s = _base(prs, brief["date"], "가격 · 출력 비교", kicker="COMPARE · 가격/성능")
    names = [v["name"] for _, v in roster[:7]]

    def bar(y, key, title, scale=1.0, fmt='#,##0'):
        vals = [v["fields"][key]["si_value"] for _, v in roster[:7]]
        _text(s, Inches(0.4), y, Inches(6.7), Inches(0.3), title, 11, True)
        if all(x is None for x in vals):
            _text(s, Inches(0.4), y + Inches(0.4), Inches(6.7), Inches(0.4), "비교 가능한 데이터 없음 (–)", 10, False, BLANK_GREY)
            return
        cd = CategoryChartData()
        cd.categories = [n + (" (–)" if x is None else "") for n, x in zip(names, vals)]
        cd.add_series(title, [None if x is None else round(x * scale, 1) for x in vals])
        gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.4), y + Inches(0.3),
                                Inches(6.7), Inches(3.25), cd)
        ch = gf.chart
        ch.has_legend = False
        ch.has_title = False
        ch.font.size, ch.font.name = Pt(8), FONT
        plot = ch.plots[0]
        plot.gap_width = 60
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format, dl.number_format_is_linked = fmt, False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.font.size = Pt(8)
        ch.value_axis.has_major_gridlines = False
        ch.value_axis.visible = False
        ch.category_axis.reverse_order = True
        ch.category_axis.format.line.color.rgb = LINE
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = NAVY
        # 대상 차량 강조
        pt = ser.points[0]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = COPPER

    bar(Inches(1.3), "price_base", "시작가 (US MSRP, USD)", fmt='"$"#,##0')
    bar(Inches(5.0), "power", "최고출력 (kW)")


def slide_insights(prs, brief):
    ins = brief.get("insights") or []
    if not ins:
        return
    s = _base(prs, brief["date"], "현대차그룹 관점 시사점", kicker="INSIGHT")
    y = Inches(1.4)
    for i, t in enumerate(ins[:5], 1):
        _text(s, Inches(0.4), y, Inches(0.5), Inches(0.4), f"{i:02d}", 16, True, COPPER)
        _text(s, Inches(0.95), y + Inches(0.03), Inches(6.15), Inches(1.3), t, 11.5)
        y += Inches(1.4)


def slide_sources(prs, brief, used):
    s = _base(prs, brief["date"], "출처 · 미확인 항목", kicker="DATA")
    src, blanks = [], []
    for v in used:
        for f in v["fields"].values():
            if f["value"] is None:
                blanks.append(f"{v['name']}: {f['label']}")
            else:
                for u in f["sources"]:
                    src.append((v["name"], f["label"], f["tier"], u))
    seen, lines = set(), []
    for name, label, tier, u in src:
        key = (name, u)
        if key in seen:
            continue
        seen.add(key)
        lines.append(f"[{name}] T{tier or '?'} {u}")
    _text(s, Inches(0.4), Inches(1.3), Inches(6.7), Inches(0.3), "제원 출처 (T1 공식HP · T2 보도자료 · T3 정부DB · T4 매체 2곳 교차)", 9, True, COPPER)
    _text(s, Inches(0.4), Inches(1.6), Inches(6.7), Inches(4.4), lines[:28] or ["–"], 6.5)
    _text(s, Inches(0.4), Inches(6.1), Inches(6.7), Inches(0.3), f"미확인(blank) 항목 {len(blanks)}건", 9, True, COPPER)
    _text(s, Inches(0.4), Inches(6.4), Inches(6.7), Inches(2.4), [" · ".join(blanks[:60]) or "없음"], 6.5, False, MUTED)


def build(brief: dict, vehicles: dict, out_path: Path) -> Path:
    from .specs import FIELDS  # noqa
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    slide_cover(prs, brief)
    slide_summary(prs, brief)
    for it in (brief.get("issues") or [])[:6]:
        slide_issue(prs, brief, it)
    used = {}
    for comp in brief.get("comparisons") or []:
        roster = _roster(comp, vehicles)
        if not roster:
            continue
        for _, v in roster:
            used[v["id"]] = v
        slide_photos(prs, brief, comp, roster)
        slide_spec_table(prs, brief, comp, roster, "기본 제원 비교",
                         ["length", "width", "height", "wheelbase", "curb_weight", "seats"])
        slide_spec_table(prs, brief, comp, roster, "파워트레인 · 성능 비교",
                         ["engine", "power", "torque", "accel_0_60mph", "towing", "fuel_economy"])
        slide_charts(prs, brief, comp, roster)
    slide_insights(prs, brief)
    slide_sources(prs, brief, list(used.values()))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path
