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


_WM = []


def _wm(b64):
    """이미지 우측하단에 '무소유' 워터마크(assets/wm.png)를 찍어 JPEG base64로 돌려준다."""
    try:
        if not _WM: _WM.append(Image.open(HERE / "assets" / "wm.png").convert("RGBA"))
        im = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGBA"); w, h = im.size; wm = _WM[0]
        tw = max(24, int(w * 0.42)); th = max(8, int(wm.size[1] * tw / wm.size[0])); m = max(2, int(w * 0.03))
        im.alpha_composite(wm.resize((tw, th), Image.LANCZOS), (w - tw - m, h - th - m))
        o = io.BytesIO(); im.convert("RGB").save(o, "JPEG", quality=88); return base64.b64encode(o.getvalue()).decode("ascii")
    except Exception: return b64


def _imgs():
    if not _IMG:
        t = (HERE / "tier_template.html").read_text("utf-8"); dec = json.JSONDecoder()
        for k in ("HERO", "SKI"):
            i = t.index(f"const {k}=") + len(f"const {k}="); _IMG[k] = {n: _wm(v) for n, v in dec.raw_decode(t[i:])[0].items()}
    return _IMG


_TPL = []


def tier_html():
    """도감 템플릿(HERO·SKI 이미지에 워터마크 적용본). 처음 한 번만 만든다."""
    if not _TPL:
        t = (HERE / "tier_template.html").read_text("utf-8"); dec = json.JSONDecoder(); im = _imgs()
        for k in ("HERO", "SKI"):
            s = t.index(f"const {k}=") + len(f"const {k}="); _, e = dec.raw_decode(t[s:])
            t = t[:s] + json.dumps(im[k], ensure_ascii=False) + t[s + e:]
        _TPL.append(t)
    return _TPL[0]


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


_STONE = {}


def _stones(im, cx, y, n, w=13, h=17, gap=2):
    """초월 돌(최대 5개): n개는 원래 색, 나머지는 어둡게. cx=가운데 x."""
    if "im" not in _STONE:
        try: _STONE["im"] = Image.open(Path(__file__).parent / "assets" / "stone.png").convert("RGBA").resize((w, h), Image.LANCZOS)
        except Exception: _STONE["im"] = None
    s = _STONE["im"]
    if s is None: return
    off = Image.merge("RGBA", (*[c.point(lambda v: int(v * 0.55)) for c in s.convert("RGB").convert("L").convert("RGB").split()], s.getchannel("A").point(lambda v: int(v * 0.85))))
    x0 = int(cx - (5 * w + 4 * gap) / 2)
    for i in range(5): im.alpha_composite(s if i < n else off, (x0 + i * (w + gap), y)) if im.mode == "RGBA" else im.paste(s if i < n else off, (x0 + i * (w + gap), y), s if i < n else off)


def _hero_card(im, d, x, y, h, tac, hd, skd, info, eff=None, rank=None):
    n = h["n"]; hh = 196
    d.rounded_rectangle((x, y, x + CW - 10, y + hh), 10, fill=CARD, outline=LINE)
    if not n:
        d.text((x + 14, y + 80), "보유 장수 부족", font=_f(16), fill=MUTE); return
    p = _pic("HERO", n, (96, 150))
    if p: im.paste(p, (x + 8, y + 8))
    else: d.rounded_rectangle((x + 8, y + 8, x + 104, y + 158), 6, fill=BG, outline=LINE); d.text((x + 56, y + 83), n, font=_f(15), fill=TXT, anchor="mm")
    x2 = x + 114; x_end = x + CW - 18; v = hd.get(n) or {}
    e0 = (eff or {}).get(n)
    if e0 and e0.get("tr"): _stones(im, x + 8 + 48, y + 8 + 150 - 46, e0["tr"])           # 초월 단계만큼 켜진 돌
    d.text((x2, y + 8), n, font=_f(19), fill=TXT)
    nc = NATION.get(v.get("nation"), MUTE)
    if v.get("nation"):
        d.rounded_rectangle((x2 + d.textlength(n, font=_f(19)) + 8, y + 11, x2 + d.textlength(n, font=_f(19)) + 31, y + 31), 5, fill=nc)
        d.text((x2 + d.textlength(n, font=_f(19)) + 19.5, y + 21), v["nation"], font=_f(13), fill=(10, 20, 18), anchor="mm")
    d.text((x2, y + 34), _fit(d, " · ".join(z for z in (v.get("pos"), v.get("troop")) if z), _f(12), x_end - x2), font=_f(12), fill=MUTE)
    d.text((x2, y + 50), _fit(d, f"{v.get('sk', '')} ({v.get('sktype', '')})" if v.get("sk") else "", _f(12), x_end - x2), font=_f(12), fill=GOLD)
    e = (eff or {}).get(n)
    if e:                                                           # 선공(초월 반영)·행동 순서
        add = "".join(f" {l}+{q}" for l, q in zip("무지통선", e.get("add", [0, 0, 0, 0])) if q)
        d.text((x2, y + 66), _fit(d, f"선공 {e['init']:g}" + (f" ·{add}" if add else (f" · 초월{e['tr']}" if e["tr"] else "")), _f(11), x_end - x2), font=_f(11), fill=JADE)
        if rank: d.rounded_rectangle((x + CW - 46, y + 10, x + CW - 20, y + 34), 8, fill=JADE); d.text((x + CW - 33, y + 22), str(rank), font=_f(15), fill=(10, 20, 18), anchor="mm")
    elif h.get("src"): d.text((x2, y + 66), _fit(d, "← " + h["src"], _f(12), x_end - x2), font=_f(12), fill=(240, 150, 60))
    if e and h.get("src"): d.text((x2, y + 78), _fit(d, "← " + h["src"], _f(11), x_end - x2), font=_f(11), fill=(240, 150, 60))
    ty = y + 94
    for t, _why in tac[:2]:
        ti = info(t); ic = _pic("SKI", t, (24, 37))
        if ic: im.paste(ic, (x2, ty))
        else: d.rounded_rectangle((x2, ty, x2 + 24, ty + 37), 4, fill=BG, outline=LINE)
        d.text((x2 + 30, ty), _fit(d, t, _f(14), x_end - x2 - 30), font=_f(14), fill=TXT)
        kc = KIND.get(ti["kind"], MUTE)
        d.text((x2 + 30, ty + 19), ti["kind"], font=_f(12), fill=kc)
        d.text((x2 + 30 + d.textlength(ti["kind"] + " ", font=_f(12)), ty + 19), f"{ti['type']}·{ti['trait']}", font=_f(12), fill=MUTE)
        ty += 44
    for _ in range(2 - len(tac[:2])):                              # 보유 전법이 모자라 비는 칸
        d.rounded_rectangle((x2, ty, x_end, ty + 37), 6, outline=LINE, width=2)
        d.text(((x2 + x_end) / 2, ty + 18), "전법 부족 (보유 전법 없음)", font=_f(12), fill=MUTE, anchor="mm"); ty += 44


def _unit_h(): return 40 + 196 + 112 + 14


def _unit(im, d, y, i, deck, unit, st, hd, skd, info, order, flow=None, eff=None):
    d.rounded_rectangle((PAD, y, W - PAD, y + _unit_h() - 8), 14, fill=PANEL, outline=LINE)
    title = f"{deck['tier']}  {deck['name']}" if deck else "티어덱 없음 · 자유 편성"
    d.rounded_rectangle((PAD + 12, y + 10, PAD + 66, y + 34), 8, fill=GOLD); d.text((PAD + 39, y + 22), f"{i}부대", font=_f(15), fill=(10, 20, 18), anchor="mm")
    d.text((PAD + 76, y + 12), title, font=_f(18), fill=TXT)
    if deck and deck.get("form"): d.text((W - PAD - 14, y + 14), "진형 " + deck["form"], font=_f(14), fill=MUTE, anchor="ra")
    rk = {n: j for j, (n, *_r) in enumerate((flow or {}).get("turn", []), 1)}
    for k, (h, tac) in enumerate(unit): _hero_card(im, d, PAD + 8 + k * CW, y + 40, h, tac, hd, skd, info, eff, rk.get(h["n"]))
    by = y + 40 + 196 + 8; lx, mw = PAD + 16, W - 2 * PAD - 100

    def row(label, text, col=TXT):
        nonlocal by
        d.text((lx, by), label, font=_f(13), fill=GOLD); d.text((lx + 74, by), _fit(d, text, _f(13), mw), font=_f(13), fill=col); by += 20
    if flow:
        st0 = flow["start"]
        row("전투 시작", "지휘 " + (" · ".join(st0["지휘"]) or "-") + "   →   패시브 " + (" · ".join(st0["패시브"]) or "-"))
        for j, (n, ini, acts, chase) in enumerate(flow["turn"], 1):
            seq = (["액티브 " + " · ".join(acts)] if acts else []) + ["평타"] + (["추격 " + " · ".join(chase)] if chase else [])
            row(f"턴 {j}번째", f"{n}({ini:g})  " + "  →  ".join(seq) if ini is not None else f"{n}  " + "  →  ".join(seq))
    else:
        allt = [t for _, tac in unit for t, _w in tac]
        seq = sorted(allt, key=lambda t: order(info(t)["type"])); pre = [t for t in seq if order(info(t)["type"]) < 4]; post = [t for t in seq if order(info(t)["type"]) > 4]
        row("발동순서", " → ".join(pre + ["평타"] + post) if seq else "전법 없음")
    gs = ("부여 " + ", ".join(st[0]) if st[0] else "부여 없음") + (("  ·  부족 " + ", ".join(st[1])) if st[1] else "")
    row("이상상태", gs, MUTE)


def render(res, season, who, deal, hd, skd, info, order, per_image=3, flows=None, eff=None):
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
            d.text((PAD, 54), "한 번 쓴 장수·전법은 다시 쓰지 않음", font=_f(12), fill=MUTE)
        y = head
        for j, (deck, unit, st) in enumerate(part, s + 1):
            _unit(im, d, y, j, deck, unit, st, hd, skd, info, order, (flows or [None] * len(res))[j - 1], eff); y += _unit_h()
        d.text((W - PAD, H - 22), "티어덱을 완성하지 못한 부대는 같은 역할의 보유 장수로 채웠습니다", font=_f(11), fill=MUTE, anchor="ra")
        b = io.BytesIO(); im.save(b, "PNG", optimize=True); out.append(b.getvalue())
    return out


_TH = {}


def thumb(kind, name, w=44, h=56):
    """내 장수·전법 탭용 작은 썸네일(JPEG base64). 없으면 빈 문자열."""
    k = (kind, name)
    if k not in _TH:
        b = _imgs()[kind].get(name); out = ""
        if b:
            try:
                im = Image.open(io.BytesIO(base64.b64decode(b))).convert("RGB").resize((w, h), Image.LANCZOS)
                o = io.BytesIO(); im.save(o, "JPEG", quality=72); out = base64.b64encode(o.getvalue()).decode("ascii")
            except Exception: out = ""
        _TH[k] = out
    return _TH[k]
