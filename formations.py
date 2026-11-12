"""진형 데이터(천하결전). g = 피격률 격자 2×3(윗줄 전열, 아랫줄 후열) — 화면에는 피격 받는 칸만 표시하고 수치는 쓰지 않는다."""
FORMS = {
    "일자진": {"s": 1, "g": [[33, 34, 33], [0, 0, 0]], "e": ["전열이 받는 피해 8% 감소"]},
    "기형진": {"s": 1, "g": [[0, 60, 0], [0, 20, 20]], "e": ["전열이 받는 피해 6% 감소", "후열이 주는 피해 12% 증가"]},
    "안형진": {"s": 1, "g": [[0, 40, 40], [20, 0, 0]], "e": ["전열 통솔 20포인트 증가", "후열이 주는 피해 15% 증가"]},
    "방원진": {"s": 1, "g": [[40, 40, 0], [0, 0, 20]], "e": ["전열이 받는 피해 5% 감소", "후열의 연타 확률 40% 증가"]},
    "추형진": {"s": 1, "g": [[0, 60, 0], [20, 0, 20]], "e": ["전열이 주는 피해 16% 증가", "후열이 받는 피해 5% 감소"]},
    "어린진": {"s": 1, "g": [[60, 0, 0], [0, 20, 20]], "e": ["전열 피신 12% 증가", "후열 회심과 묘책 확률 8% 증가"]},
    "구행진": {"s": 2, "g": [[0, 60, 0], [20, 0, 20]], "e": ["전열이 받는 피해 8% 감소", "후열의 연타 확률 25% 증가"]},
    "언월진": {"s": 2, "g": [[40, 0, 40], [0, 20, 0]], "e": ["전열이 주는 피해 14% 증가", "후열이 받는 피해 5% 감소"]},
}
ALIAS = {"안영진": "안형진", "추행진": "추형진", "기영진": "기형진", "구형진": "구행진"}      # 티어덱 데이터의 오타


def names(s):
    """'안형진/기형진' 같은 문자열 → 데이터에 있는 진형 이름 목록(오타 보정, 없는 이름은 제외)."""
    out = []
    for x in str(s or "").replace("·", "/").replace(",", "/").split("/"):
        x = ALIAS.get(x.strip(), x.strip())
        if x in FORMS and x not in out: out.append(x)
    return out


def cells(name):
    """피격 받는 칸 표시용 2×3 불리언 격자."""
    return [[1 if v else 0 for v in r] for r in FORMS[name]["g"]]


def pick(deck_form, heroes, season=1, deck_name="", chaser=False):
    """부대에 맞는 진형 하나를 고른다 → (이름, 이유).
    heroes = [{"n": 이름, "pos": 전열/후열/균형, "sktype": 지휘/패시브/액티브/추격}, ...]
    티어덱에 진형이 적혀 있으면 그 안에서, 없으면(자유 편성·개척덱) 시즌에 쓸 수 있는 전체 진형에서 고른다."""
    import re
    hs = [h for h in heroes if h and h.get("n")]
    listed = names(deck_form)
    if listed: return listed[0], "티어덱 추천 진형"               # 티어덱에 적힌 진형이 있으면 첫 번째(우선 추천)를 그대로 쓴다
    cand = [n for n, v in FORMS.items() if v["s"] <= season]
    if not cand or not hs: return "", ""
    nf = sum(h.get("pos") == "전열" for h in hs); nm = sum(h.get("pos") == "균형" for h in hs)
    front_deal = any(h.get("pos") == "전열" and h.get("sktype") in ("액티브", "추격") for h in hs)
    back = any(h.get("pos") != "전열" for h in hs); lubu = any(h["n"] == "여포" for h in hs)
    best = None
    for k, n in enumerate(cand):
        fs = sum(1 for v in FORMS[n]["g"][0] if v); why = []
        dist = 0 if nf <= fs <= nf + nm else min(abs(fs - nf), abs(fs - nf - nm))
        sc = 3 - 3 * dist
        if not dist: why.append(f"전열 {fs}·후열 {3 - fs} 배치에 맞음")
        for e in FORMS[n]["e"]:
            m = re.search(r"(\d+)", e); p = int(m.group(1)) if m else 0
            if "연타" in e and (chaser or lubu): sc += 2 + p / 20; why.append(e)
            elif "후열이 주는 피해" in e and back: sc += p / 10; why.append(e)
            elif "전열이 주는 피해" in e and front_deal: sc += p / 10; why.append(e)
            elif "주는 피해" not in e and "연타" not in e: sc += 0.5
        if lubu:                                                   # 여포: 기본 방원진(원문사극), 반격 덱이면 추형진
            want = "추형진" if "반격" in (deck_name or "") else "방원진"
            if n == want: sc += 3; why.insert(0, "여포 " + ("반격" if want == "추형진" else "연타") + " 운영")
        if best is None or sc > best[0]: best = (sc, n, why)
    return best[1], " · ".join(best[2][:3])
