"""/추천덱 결과를 장수 카드·전법 아이콘이 들어간 이미지(PNG)로 그린다. (Pillow, 번들 한글 폰트)"""
import base64, io, json, re
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
FONT = HERE / "fonts" / "NotoSansKR-Bold-subset.otf"
BG, PANEL, CARD, LINE = (14, 23, 22), (20, 33, 31), (26, 42, 39), (38, 58, 54)
TXT, MUTE, GOLD, JADE = (232, 227, 211), (147, 164, 158), (224, 182, 90), (94, 207, 177)
KIND = {"한방딜": (234, 90, 66), "도트딜": (240, 150, 60), "일반딜": (150, 165, 160), "회복": (94, 207, 177), "지원": (127, 166, 209)}
NATION = {"촉": (80, 170, 110), "위": (100, 140, 220), "오": (220, 100, 90), "군": (170, 170, 170)}
_IMG = {}
_FONTS = {}


def _f(sz):
    if sz not in _FONTS: _FONTS[sz] = ImageFont.truetype(str(FONT), sz)
    return _FONTS[sz]


def _imgs():
    if not _IMG:
        t = (HERE / "tier_template.html").read_text("utf-8"); dec = json.JSONDecoder()
        for k in ("HERO", "SKI"):
            i = t.index(f"const {k}=") + len(f"const {k}="); _IMG[k] = dec.raw_decode(t[i:])[0]
    return _IMG


def _pic(kind, name, size):
    b = _imgs()[kind].get(name)
    if not b: return None
    try: return Image.open(io.BytesIO(base64.b64decode(b))).convert("RGB").resize(size, Image.LANCZOS)
    except Exception: return None


def _fit(d, text, font, maxw):
    """글자가 너무 길면 …으로 줄인다."""
    if d.textlength(text, font=font) <= maxw: return text
    while text and d.textlength(text + "…", font=font) > maxw: text = text[:-1]
    return text + "…"


W, PAD, CW = 960, 20, 300          # 이미지 너비, 바깥 여백, 장수 카드 너비


def _hero_card(im, d, x, y, h, tac, hd, skd, info):
    n = h["n"]; hh = 196
    d.rounded_rectangle((x, y, x + CW - 10, y + hh), 10, fill=CARD, outline=LINE)
    if not n:
        d.text((x + 14, y + 80), "보유 장수 부족", font=_f(16), fill=MUTE); return
    p = _pic("HERO", n, (96, 150))
    if p: im.paste(p, (x + 8, y + 8))
    else: d.rounded_rectangle((x + 8, y + 8, x + 104, y + 158), 6, fill=BG, outline=LINE); d.text((x + 56, y + 83), n, font=_f(15), fill=TXT, anchor="mm")
    x2 = x + 114; x_end = x + CW - 18; v = hd.get(n) or {}
    d.text((x2, y + 8), n, font=_f(19), fill=TXT)
    nc = NATION.get(v.get("nation"), MUTE)
    if v.get("nation"):
        d.rounded_rectangle((x2 + d.textlength(n, font=_f(19)) + 8, y + 11, x2 + d.textlength(n, font=_f(19)) + 31, y + 31), 5, fill=nc)
        d.text((x2 + d.textlength(n, font=_f(19)) + 19.5, y + 21), v["nation"], font=_f(13), fill=(10, 20, 18), anchor="mm")
    d.text((x2, y + 34), _fit(d, " · ".join(z for z in (v.get("pos"), v.get("troop")) if z), _f(12), x_end - x2), font=_f(12), fill=MUTE)
    d.text((x2, y + 50), _fit(d, f"{v.get('sk', '')} ({v.get('sktype', '')})" if v.get("sk") else "", _f(12), x_end - x2), font=_f(12), fill=GOLD)
    if h.get("src"): d.text((x2, y + 66), _fit(d, "← " + h["src"], _f(12), x_end - x2), font=_f(12), fill=(240, 150, 60))
    ty = y + 88
    for t, _why in tac[:2]:
        ti = info(t); ic = _pic("SKI", t, (24, 37))
        if ic: im.paste(ic, (x2, ty))
        else: d.rounded_rectangle((x2, ty, x2 + 24, ty + 37), 4, fill=BG, outline=LINE)
        d.text((x2 + 30, ty), _fit(d, t, _f(14), x_end - x2 - 30), font=_f(14), fill=TXT)
        kc = KIND.get(ti["kind"], MUTE)
        d.text((x2 + 30, ty + 19), ti["kind"], font=_f(12), fill=kc)
        d.text((x2 + 30 + d.textlength(ti["kind"] + " ", font=_f(12)), ty + 19), f"{ti['type']}·{ti['trait']}", font=_f(12), fill=MUTE)
        ty += 44
    if not tac: d.text((x2, ty + 6), "전법 없음", font=_f(13), fill=MUTE)


def _unit_h(): return 40 + 196 + 56 + 14


def _unit(im, d, y, i, deck, unit, st, hd, skd, info, order):
    d.rounded_rectangle((PAD, y, W - PAD, y + _unit_h() - 8), 14, fill=PANEL, outline=LINE)
    title = f"{deck['tier']}  {deck['name']}" if deck else "티어덱 없음 · 자유 편성"
    d.rounded_rectangle((PAD + 12, y + 10, PAD + 66, y + 34), 8, fill=GOLD); d.text((PAD + 39, y + 22), f"{i}부대", font=_f(15), fill=(10, 20, 18), anchor="mm")
    d.text((PAD + 76, y + 12), title, font=_f(18), fill=TXT)
    if deck and deck.get("form"): d.text((W - PAD - 14, y + 14), "진형 " + deck["form"], font=_f(14), fill=MUTE, anchor="ra")
    for k, (h, tac) in enumerate(unit): _hero_card(im, d, PAD + 8 + k * CW, y + 40, h, tac, hd, skd, info)
    allt = [t for _, tac in unit for t, _w in tac]
    seq = sorted(allt, key=lambda t: order(info(t)["type"]))
    pre = [t for t in seq if order(info(t)["type"]) < 4]; post = [t for t in seq if order(info(t)["type"]) > 4]
    line = " → ".join(pre + ["평타"] + post) if seq else "전법 없음"
    d.text((PAD + 16, y + 40 + 196 + 8), "발동순서", font=_f(13), fill=GOLD)
    d.text((PAD + 84, y + 40 + 196 + 8), _fit(d, line, _f(13), W - 2 * PAD - 100), font=_f(13), fill=TXT)
    gs = ("부여 " + ", ".join(st[0]) if st[0] else "부여 없음") + (("  ·  부족 " + ", ".join(st[1])) if st[1] else "")
    d.text((PAD + 16, y + 40 + 196 + 28), "이상상태", font=_f(13), fill=GOLD)
    d.text((PAD + 84, y + 40 + 196 + 28), _fit(d, gs, _f(13), W - 2 * PAD - 100), font=_f(13), fill=MUTE)


def render(res, season, who, deal, hd, skd, info, order, per_image=3):
    """추천 결과 → PNG 바이트 목록(한 장에 부대 per_image개)."""
    out = []
    for s in range(0, len(res), per_image):
        part = res[s:s + per_image]; head = 78 if s == 0 else 14
        H = head + len(part) * _unit_h() + 30
        im = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(im)
        if s == 0:
            d.text((PAD, 14), f"{who} 추천덱 · 시즌{season}", font=_f(26), fill=TXT)
            x = PAD
            for k in ("한방딜", "도트딜", "일반딜", "회복", "지원"):
                d.text((W - PAD - 5 * 70 + ["한방딜", "도트딜", "일반딜", "회복", "지원"].index(k) * 70, 22), k, font=_f(14), fill=KIND[k])
            d.text((PAD, 54), f"딜 방식 {deal} 우선 · 한 번 쓴 장수·전법은 다시 쓰지 않음", font=_f(12), fill=MUTE)
        y = head
        for j, (deck, unit, st) in enumerate(part, s + 1):
            _unit(im, d, y, j, deck, unit, st, hd, skd, info, order); y += _unit_h()
        d.text((W - PAD, H - 22), "티어덱을 완성하지 못한 부대는 같은 역할의 보유 장수로 채웠습니다", font=_f(11), fill=MUTE, anchor="ra")
        b = io.BytesIO(); im.save(b, "PNG", optimize=True); out.append(b.getvalue())
    return out
