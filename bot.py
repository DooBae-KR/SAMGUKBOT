"""연맹 디스코드 인증 봇
- 신규 입장자에게 [인증하기] 버튼 안내 → UID/직업/닉네임 입력 폼(모달)
- UID를 엑셀 명단과 대조해 검증 → 별명을 [직업]닉네임 으로 변경 → 역할(동맹원/간부) 부여
"""
import urllib.parse, urllib.error, hmac, hashlib, unicodedata
import base64, tempfile, os, json, re, csv, io, shutil, logging, urllib.request, datetime, asyncio, time, copy, threading, queue, contextvars
from aiohttp import web
from pathlib import Path
import discord
from discord import app_commands
from discord.ext import tasks
from dotenv import load_dotenv
import openpyxl

load_dotenv()
TOKEN = os.environ["DISCORD_TOKEN"]
GUILD_ID = int(os.environ["GUILD_ID"])                  # 제작자의 '홈' 서버(기존 데이터가 있는 서버)
_CTX = contextvars.ContextVar("gid", default=GUILD_ID)  # 지금 처리 중인 서버. 모든 저장·명단 조회는 이 값 기준으로 분리된다


def cur_gid(): return _CTX.get()
def set_gid(g): _CTX.set(int(g) if g else GUILD_ID)


LICENSED = set()                                        # 제작자가 승인한 서버(홈 서버는 항상 허가)
def licensed(gid): return gid == GUILD_ID or gid in LICENSED
EXCEL_PATH = os.environ.get("EXCEL_PATH", "261023.xlsx")
SHEET_NAME = os.environ.get("SHEET_NAME", "")            # 비우면 첫 시트
VERIFY_CHANNEL_ID = int(os.environ.get("VERIFY_CHANNEL_ID", "0"))  # 인증 안내 채널
MEMBER_ROLE = os.environ.get("MEMBER_ROLE", "동맹원")
OFFICER_ROLE = os.environ.get("OFFICER_ROLE", "간부")
DIPLOMAT_ROLE = os.environ.get("DIPLOMAT_ROLE", "외교관")     # 외교 목적으로 들어온 타 동맹원 역할
OFFICER_TITLES = {s.replace(" ", "") for s in os.environ.get("OFFICER_TITLES", "태위,군주,사공,사도,대장군,대홍려,진군장군,무군장군,익군장군,안국장군,분무장군,복파장군,파로장군,호위장군").split(",") if s.strip()}
MAX_TRIES = int(os.environ.get("MAX_TRIES", "5"))
SHEET_ID = os.environ.get("SHEET_ID", "")          # 구글 시트에서 명단을 받을 때(선택)
ROSTER_GID = os.environ.get("ROSTER_GID", "")      # 명단 탭의 gid
TALK_GID = os.environ.get("TALK_GID", "")          # 톡방참여자 탭 gid(선택)
WEB_PORT = int(os.environ.get("PORT") or os.environ.get("WEB_PORT", "8080"))   # Render는 PORT를 자동 지정  # 상황판 웹 포트(0이면 끔)
WEB_KEY = os.environ.get("WEB_KEY", "")            # 설정하면 주소 뒤에 ?key=값 이 있어야 열림
GRACE_END = os.environ.get("GRACE_END", "2026-11-06")
SYNC_MINUTES = int(os.environ.get("SYNC_MINUTES", "10"))   # 시트를 다시 읽고 역할/별명을 맞추는 주기(0이면 끔)
ADMIN_CHANNEL_ID = int(os.environ.get("ADMIN_CHANNEL_ID", "0") or 0)   # 명단 이탈 의심 알림을 받을 관리자 채널(비우면 서버장 DM)
MISSING_MAX = int(os.environ.get("MISSING_MAX", "5"))      # 한 번에 이만큼 넘게 사라지면 대량 이탈 의심으로 묶어 알림(시트 오류 방지)
LEAVE_PATH = Path(os.environ.get("LEAVE_PATH", "leave.json"))
DB_PATH = Path(os.environ.get("DB_PATH", "verified.json"))

# 엑셀 헤더 후보(실제 파일 헤더에 맞게 .env 로 바꿀 수 있음)
COL_UID = [s for s in os.environ.get("COL_UID", "UID,uid,Uid,유저ID,아이디,ID").split(",")]
COL_JOB = [s for s in os.environ.get("COL_JOB", "직업").split(",")]
COL_TITLE = [s for s in os.environ.get("COL_TITLE", "직위,직책,등급").split(",")]
COL_NAME = [s for s in os.environ.get("COL_NAME", "닉네임,이름,닉").split(",")]
DEFAULT_JOBS = [s.strip() for s in os.environ.get("JOBS", "").split(",") if s.strip()]  # 비우면 엑셀 직업 값 사용

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("bot")


# ---------- 엑셀 ----------
class Roster:
    def __init__(self, path):
        self.path = Path(path); self.mtime = 0; self.by_uid = {}; self.jobs = []; self.members = []; self.talk = set(); self.src = ""

    @staticmethod
    def _find(headers, cands):
        low = {str(h).strip().lower(): i for i, h in enumerate(headers) if h is not None}
        for c in cands:
            if c.strip().lower() in low: return low[c.strip().lower()]
        return None

    @staticmethod
    def norm_uid(v):
        if v is None: return ""
        if isinstance(v, float) and v.is_integer(): v = int(v)
        return re.sub(r"\s+", "", str(v)).lstrip("'")

    @staticmethod
    def parse_members(rows):
        rows = [list(r) for r in rows]
        hi = next((i for i, r in enumerate(rows) if "멤버" in [str(x).strip() for x in r]), None)
        if hi is None: raise RuntimeError("멤버 열을 찾지 못했습니다.")
        H = [str(x).strip() for x in rows[hi]]
        need = ["멤버", "직업", "조별", "직위", "번영", "주간 무훈", "주간 공헌", "주 공성 횟수", "주둔지"]
        miss = [n for n in need if n not in H]
        if miss: raise RuntimeError("상황판용 열이 없습니다: " + ", ".join(miss))
        ix = {n: H.index(n) for n in need}
        def num(v):
            try: return float(str(v).replace(",", "").strip() or 0)
            except ValueError: return 0
        out = []
        for r in rows[hi + 1:]:
            g = lambda k: (str(r[ix[k]]).strip() if ix[k] < len(r) and r[ix[k]] is not None else "")
            n = g("멤버")
            if not n: continue
            m = re.match(r"\s*([^(]*)\(\s*(\d+)\s*,\s*(\d+)\s*\)", g("주둔지"))
            out.append({"n": n, "j": g("직업"), "g": g("조별"), "r": g("직위"), "p": num(g("번영")), "m": num(g("주간 무훈")),
                        "c": num(g("주간 공헌")), "s": int(num(g("주 공성 횟수"))), "reg": m.group(1).strip() if m else "",
                        "x": int(m.group(2)) if m else 0, "y": int(m.group(3)) if m else 0, "t": 0})
        return out

    def _members_safe(self, rows):
        try: self.members = self.parse_members(rows)
        except Exception as e: self.members = []; log.warning("상황판 데이터 없음: %s", e)

    def _ingest(self, rows):
        rows = iter(rows); headers = None
        for r in rows:                      # 헤더 행 탐색
            if self._find(r, COL_UID) is not None: headers = r; break
        if headers is None: raise RuntimeError("명단에서 UID 열을 찾지 못했습니다. COL_UID 설정을 확인하세요.")
        iu, ij, it, inm = (self._find(headers, c) for c in (COL_UID, COL_JOB, COL_TITLE, COL_NAME))
        by, jobs = {}, []
        for r in rows:
            uid = self.norm_uid(r[iu]) if iu is not None and iu < len(r) else ""
            if not uid: continue
            g = lambda i: (str(r[i]).strip() if i is not None and i < len(r) and r[i] is not None else "")
            by[uid] = {"job": g(ij), "title": g(it), "name": g(inm)}
            if g(ij) and g(ij) not in jobs: jobs.append(g(ij))
        if not by: raise RuntimeError("명단에 CID 행이 없습니다.")
        return by, jobs

    def load(self):
        if not self.path.exists(): return      # 엑셀이 없으면 시트로 받은 명단을 그대로 사용
        m = self.path.stat().st_mtime
        if m == self.mtime: return
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws = wb[SHEET_NAME] if SHEET_NAME else wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        self.by_uid, self.jobs = self._ingest(rows); self._members_safe(rows); self.src = "엑셀 " + self.path.name; self.mtime = m
        log.info("명단 로드: %d명, 직업 %s", len(self.by_uid), self.jobs)

    def load_csv_text(self, text):
        """구글 시트 CSV로 명단을 교체(다음 엑셀 파일 변경 전까지 유지)."""
        rows = list(csv.reader(io.StringIO(text)))
        by, jobs = self._ingest(rows)
        self.by_uid, self.jobs = by, jobs; self._members_safe(rows); self.src = "구글 시트 인사 탭"
        if self.path.exists(): self.mtime = self.path.stat().st_mtime

    def get(self, uid):
        self.load(); return self.by_uid.get(self.norm_uid(uid))

    def job_list(self):
        self.load(); return DEFAULT_JOBS or self.jobs


_home_roster = Roster(EXCEL_PATH)                       # 홈 서버: 기존처럼 엑셀/구글 시트
_tenant_rosters = {}                                    # 그 외 서버: DB(sam_emp / sam_emp_stat)


def roster_for(gid):
    if gid == GUILD_ID and not SB_ON: return _home_roster          # DB를 쓰면 홈 서버도 다른 서버와 똑같이 DB 명단
    r = _tenant_rosters.get(gid)
    if r is None: r = _tenant_rosters[gid] = Roster("__none__")
    return r


class _RosterProxy:
    """roster.xxx 를 지금 처리 중인 서버의 명단으로 연결한다."""
    def __getattr__(self, k): return getattr(roster_for(cur_gid()), k)
    def __setattr__(self, k, v): setattr(roster_for(cur_gid()), k, v)


roster = _RosterProxy()


# ---------- 인증 기록 ----------
SAVE_URL = os.environ.get("SAVE_URL", "").strip()     # Apps Script 웹앱 주소(비우면 로컬 파일만 사용)
SAVE_KEY = os.environ.get("SAVE_KEY", "").strip()
_cache, _outq, _lock = {}, queue.Queue(), threading.Lock()


def _remote_get(name):
    url = f"{SAVE_URL}?{urllib.parse.urlencode({'key': SAVE_KEY, 'name': name})}"
    with urllib.request.urlopen(url, timeout=30) as r: res = json.loads(r.read().decode("utf-8"))
    if not res.get("ok"): raise RuntimeError(res.get("error", "저장소 응답 오류"))
    return res.get("data")


def _remote_put(name, data):
    body = json.dumps({"key": SAVE_KEY, "name": name, "data": data}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(SAVE_URL, data=body, headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=30) as r: res = json.loads(r.read().decode("utf-8"))
    if not res.get("ok"): raise RuntimeError(res.get("error", "저장 실패"))


# ---------- Supabase(주 저장소) ----------
_u = os.environ.get("SUPABASE_URL", "").strip()
if _u and "://" not in _u: _u = "https://" + _u                   # https:// 를 빼먹은 경우
_sp = urllib.parse.urlsplit(_u)
SB_URL = f"{_sp.scheme}://{_sp.netloc}" if _sp.netloc else ""      # /rest/v1 등 뒤에 붙은 경로는 버리고 주소만 사용
SB_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "").strip()
SB_ON = bool(SB_URL and SB_KEY)
_MISSING = object()                                        # Supabase에 행이 없음(= 아직 이전 안 됨)
_sb_tenant_ok = set()


def _sb_req(method, path, body=None, prefer=None):
    h = {"apikey": SB_KEY, "Content-Type": "application/json"}
    if SB_KEY.startswith("eyJ"): h["Authorization"] = "Bearer " + SB_KEY    # 옛 JWT 키만 Bearer 필요
    if prefer: h["Prefer"] = prefer
    data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    req = urllib.request.Request(f"{SB_URL}/rest/v1/{path}", data=data, headers=h, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as r: raw = r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Supabase {e.code}: {e.read().decode('utf-8', 'ignore')[:200]}")
    return json.loads(raw) if raw.strip() else None


def _sb_ensure_tenant(gid=GUILD_ID):
    if gid in _sb_tenant_ok: return
    _sb_req("POST", "sam_alli?on_conflict=guild_id", [{"guild_id": gid, "name": "default"}],
            prefer="resolution=ignore-duplicates,return=minimal")
    _sb_tenant_ok.add(gid)


def _sb_get(name, gid=GUILD_ID):
    rows = _sb_req("GET", f"sam_store?guild_id=eq.{gid}&name=eq.{urllib.parse.quote(name, safe='')}&select=data")
    return rows[0]["data"] if rows else _MISSING


def _sb_put(name, data, gid=GUILD_ID):
    _sb_ensure_tenant(gid)
    _sb_req("POST", "sam_store?on_conflict=guild_id,name",
            [{"guild_id": gid, "name": name, "data": data, "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat()}],
            prefer="resolution=merge-duplicates,return=minimal")


def _sb_put_tenant(d, gid=GUILD_ID):
    _sb_req("POST", "sam_alli?on_conflict=guild_id",
            [{"guild_id": gid, "name": d["name"], "game_server": str(d["server"]), "config": {"alliance": d["name"], "server": d["server"]}}],
            prefer="resolution=merge-duplicates,return=minimal")


def _writer():
    while True:
        item = _outq.get()
        try:                                               # 쌓인 요청은 (서버, 이름)별 최신 것만 보냄
            latest = {(item[0], item[1]): item[2]}
            while not _outq.empty(): g2, n2, d2 = _outq.get_nowait(); latest[(g2, n2)] = d2
            for (g, n), d in latest.items():
                if SB_ON:
                    for attempt in range(3):
                        try:
                            if n == "alliance" and d: _sb_ensure_tenant(g); _sb_put_tenant(d, g)
                            _sb_put(n, d, g); break
                        except Exception as e:
                            log.warning("Supabase 저장 실패(%s/%s) %d/3: %s", g, n, attempt + 1, e); time.sleep(3)
                if SAVE_URL and g == GUILD_ID:             # 시트 미러는 홈 서버만
                    for attempt in range(3):
                        try: _remote_put(n, d); break
                        except Exception as e:
                            log.warning("시트 저장 실패(%s) %d/3: %s", n, attempt + 1, e); time.sleep(3)
        except Exception: log.exception("저장 스레드 오류")


if SAVE_URL or SB_ON: threading.Thread(target=_writer, daemon=True).start()
log.info("저장소: %s%s", "Supabase" if SB_ON else "", (" + 시트 미러" if SB_ON and SAVE_URL else "시트" if SAVE_URL else "" if SB_ON else "로컬 파일"))


def roster_load_db(gid):
    """sam_emp(사람) + sam_emp_stat(가장 최근 날짜 수치)로 그 서버의 명단을 메모리에 올린다."""
    emp = _sb_req("GET", f"sam_emp?guild_id=eq.{gid}&select=cid,name,job,extra&limit=5000") or []
    last = _sb_req("GET", f"sam_emp_stat?guild_id=eq.{gid}&select=snap_date&order=snap_date.desc&limit=1") or []
    st = {}
    if last:
        for r in _sb_req("GET", f"sam_emp_stat?guild_id=eq.{gid}&snap_date=eq.{last[0]['snap_date']}&select=*&limit=5000") or []: st[r["cid"]] = r
    by, members, jobs = {}, [], []
    for e in emp:
        x = st.get(e["cid"], {}); job = e.get("job") or ""
        by[e["cid"]] = {"job": job, "title": x.get("title") or "", "name": e.get("name") or ""}
        if job and job not in jobs: jobs.append(job)
        members.append({"n": e.get("name") or "", "j": job, "g": (e.get("extra") or {}).get("조별", ""), "r": x.get("title") or "",
                        "p": x.get("prosperity") or 0, "m": x.get("war_merit") or 0, "c": x.get("contribution") or 0, "s": x.get("siege_count") or 0,
                        "reg": x.get("camp") or "", "x": x.get("camp_x") or 0, "y": x.get("camp_y") or 0, "t": 0})
    r = roster_for(gid); r.by_uid, r.jobs, r.members, r.src = by, jobs, members, "DB 명단"
    return len(by)


def _lpath(path, gid):
    return path if gid == GUILD_ID else path.parent / f"g{gid}" / path.name      # 홈 서버 외에는 서버별 폴더


def store_get(name, path, default):
    gid = cur_gid(); path = _lpath(path, gid); key = (gid, name)
    with _lock:
        if key in _cache: return copy.deepcopy(_cache[key])
        data, migrate = None, False
        if SB_ON:
            for attempt in range(3):
                try:
                    v = _sb_get(name, gid)
                    if v is _MISSING: migrate = True       # 처음이면 아래 시트/로컬 값을 Supabase로 이전
                    else: data = v
                    break
                except Exception as e: log.warning("Supabase 불러오기 실패(%s/%s) %d/3: %s", gid, name, attempt + 1, e); time.sleep(2)
        if data is None and SAVE_URL and gid == GUILD_ID:
            for attempt in range(3):
                try: data = _remote_get(name); break
                except Exception as e: log.warning("시트 불러오기 실패(%s) %d/3: %s", name, attempt + 1, e); time.sleep(2)
        if data is None and path.exists():
            try: data = json.loads(path.read_text("utf-8"))
            except Exception: pass
        if migrate and data is not None: _outq.put((gid, name, copy.deepcopy(data))); log.info("Supabase로 이전: %s/%s", gid, name)
        _cache[key] = data if data is not None else default
        return copy.deepcopy(_cache[key])


def store_set(name, path, data):
    gid = cur_gid(); path = _lpath(path, gid)
    with _lock: _cache[(gid, name)] = copy.deepcopy(data)
    try: path.parent.mkdir(parents=True, exist_ok=True); path.write_text(json.dumps(data, ensure_ascii=False, indent=1), "utf-8")
    except Exception: pass                                 # 읽기 전용 디스크여도 원격 저장은 계속
    if SAVE_URL or SB_ON: _outq.put((gid, name, copy.deepcopy(data)))


def db_load(): return store_get("verified", DB_PATH, {"uid": {}, "tries": {}})
def db_save(d): store_set("verified", DB_PATH, d)


def make_nick(job, nick):
    s = f"[{job}]{nick}"
    return s if len(s) <= 32 else s[:32]


async def apply_verification(member: discord.Member, uid, job, nick, info):
    guild = member.guild
    mem_role = discord.utils.get(guild.roles, name=MEMBER_ROLE)
    off_role = discord.utils.get(guild.roles, name=OFFICER_ROLE)
    is_officer = info["title"].replace(" ", "") in OFFICER_TITLES
    target = off_role if is_officer else mem_role
    if target is None: raise RuntimeError(f"서버에 '{OFFICER_ROLE if is_officer else MEMBER_ROLE}' 역할이 없습니다.")
    notes = []
    try: await member.edit(nick=make_nick(job, nick), reason="연맹 인증")
    except discord.Forbidden: notes.append("별명 변경 권한이 없습니다(서버 주인이거나 봇 역할이 낮음).")
    roles_add = [target] + ([mem_role] if is_officer and mem_role else [])   # 간부는 동맹원 역할도 함께
    await member.add_roles(*roles_add, reason="연맹 인증")
    return target.name, notes


# ---------- UI ----------
class CtxView(discord.ui.View):
    async def interaction_check(self, it): set_gid(it.guild_id); return True      # 버튼을 누른 서버 기준으로 처리


class CtxModal(discord.ui.Modal):
    async def interaction_check(self, it): set_gid(it.guild_id); return True


class VerifyModal(CtxModal, title="연맹 인증"):
    uid = discord.ui.TextInput(label="CID (캐릭터 ID)", placeholder="게임 내 캐릭터 ID", max_length=30)
    job = discord.ui.TextInput(label="직업 (비우면 명단 값 사용)", placeholder="예) 천공", max_length=10, required=False)
    nick = discord.ui.TextInput(label="닉네임 (비우면 명단 값 사용)", placeholder="게임 내 닉네임", max_length=25, required=False)

    async def on_submit(self, it: discord.Interaction):
        db = db_load(); did = str(it.user.id)
        tries = db["tries"].get(did, 0)
        if tries >= MAX_TRIES:
            return await it.response.send_message("실패 횟수를 초과했습니다. 관리자에게 문의하세요.", ephemeral=True)
        uid = Roster.norm_uid(self.uid.value); job = self.job.value.strip(); nick = self.nick.value.strip()
        try: info = roster.get(uid)
        except Exception as e:
            log.exception("명단 로드 실패"); return await it.response.send_message("명단 파일을 읽지 못했습니다. 관리자에게 문의하세요.", ephemeral=True)
        err = None
        if info is None: err = "명단에 없는 CID입니다. CID를 다시 확인해 주세요."
        elif uid in db["uid"] and db["uid"][uid] != did: err = "이미 다른 계정으로 인증된 CID입니다. 관리자에게 문의하세요."
        else:
            jobs = roster.job_list()
            job = job or info["job"]; nick = nick or info["name"]      # 비우면 인사 명단 값 사용
            if not job or not nick: err = "명단에 직업/닉네임이 없습니다. 직업과 닉네임을 직접 입력해 주세요."
            elif jobs and job not in jobs: err = f"직업은 다음 중 하나여야 합니다: {', '.join(jobs)}"
        if err:
            db["tries"][did] = tries + 1; db.setdefault("tried", {})[did] = uid; db_save(db)     # 어떤 CID로 실패했는지 기록(관리자가 CID로 초기화할 수 있게)
            return await it.response.send_message(f"❌ {err} (남은 시도 {MAX_TRIES - tries - 1}회)", ephemeral=True)
        await it.response.defer(ephemeral=True)
        try: role, notes = await apply_verification(it.user, uid, job, nick, info)
        except Exception as e:
            log.exception("역할 부여 실패"); return await it.followup.send(f"처리 중 오류: {e}", ephemeral=True)
        db["uid"][uid] = did; db.setdefault("nick", {})[uid] = nick; db["tries"].pop(did, None); db.get("tried", {}).pop(did, None); db_save(db)
        await it.followup.send(f"✅ 인증 완료! 별명: **{make_nick(job, nick)}** · 역할: **{role}**" + ("\n" + "\n".join(notes) if notes else ""), ephemeral=True)


# ---------- 외교관 인증 ----------
# ---------- 동맹 정보(서버 번호 + 동맹명) ----------
ALLY_PATH = DB_PATH.with_name("alliance.json")


def ally_load():
    d = store_get("alliance", ALLY_PATH, None)
    if not d and cur_gid() != GUILD_ID: return {"name": "", "server": ""}
    if not d:                                              # 홈 서버 최초 1회: 환경변수/기본값으로 등록
        d = {"name": os.environ.get("ALLIANCE_NAME", "공성부대").strip(), "server": os.environ.get("GAME_SERVER", "41").strip()}
        store_set("alliance", ALLY_PATH, d)
    return d


class AllyModal(CtxModal, title="동맹 등록"):
    server = discord.ui.TextInput(label="서버 번호 (숫자만 입력)", placeholder="예: 41", min_length=1, max_length=5)
    name = discord.ui.TextInput(label="동맹명", placeholder="예: 공성부대", min_length=1, max_length=12)

    def __init__(self):
        super().__init__(); cur = ally_load()
        self.server.default, self.name.default = cur["server"], cur["name"]

    async def on_submit(self, it: discord.Interaction):
        sv, nm = self.server.value.strip(), self.name.value.strip()
        if not (sv.isascii() and sv.isdigit()): return await it.response.send_message("❌ 서버 번호는 숫자만 입력할 수 있습니다. (예: 41)", ephemeral=True)
        if not nm: return await it.response.send_message("❌ 동맹명을 입력해 주세요.", ephemeral=True)
        store_set("alliance", ALLY_PATH, {"name": nm, "server": str(int(sv))})
        await it.response.send_message(f"✅ 동맹 등록 완료 — **{int(sv)}서버 · {nm}**", ephemeral=True)


class DipModal(CtxModal, title="외교관 등록"):
    ally = discord.ui.TextInput(label="동맹명", placeholder="소속 동맹 이름", max_length=12)
    uid = discord.ui.TextInput(label="CID (캐릭터 ID)", placeholder="게임 내 캐릭터 ID", max_length=30)
    nick = discord.ui.TextInput(label="닉네임", placeholder="게임 내 닉네임", max_length=20)

    async def on_submit(self, it: discord.Interaction):
        db = db_load(); did = str(it.user.id); dip = db.setdefault("dip", {})
        ally, uid, name = self.ally.value.strip(), Roster.norm_uid(self.uid.value), self.nick.value.strip()
        err = None
        if not (ally and uid and name): err = "동맹명, CID, 닉네임을 모두 입력해 주세요."
        elif roster.by_uid.get(uid): err = "우리 연맹 명단에 있는 CID입니다. 일반 인증(인증하기)을 이용해 주세요."
        elif uid in db["uid"] and db["uid"][uid] != did: err = "이미 다른 계정으로 등록된 CID입니다. 관리자에게 문의하세요."
        elif did in db["uid"].values() and db["uid"].get(uid) != did: err = "이미 인증된 계정입니다. 관리자에게 문의하세요."
        role = discord.utils.get(it.guild.roles, name=DIPLOMAT_ROLE) if it.guild else None
        if not err and role is None: err = f"서버에 '{DIPLOMAT_ROLE}' 역할이 없습니다. 관리자에게 문의하세요."
        if err: return await it.response.send_message(f"❌ {err}", ephemeral=True)
        await it.response.defer(ephemeral=True); notes = []
        nick = f"[{ally}] {name}"[:32]
        try: await it.user.edit(nick=nick, reason="외교관 등록")
        except discord.Forbidden: notes.append("별명 변경 권한이 없습니다(서버 주인이거나 봇 역할이 낮음).")
        try: await it.user.add_roles(role, reason="외교관 등록")
        except discord.Forbidden: return await it.followup.send("❌ 역할을 부여할 권한이 없습니다. 관리자에게 문의하세요.", ephemeral=True)
        stamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
        db["uid"][uid] = did; dip[uid] = {"ally": ally, "name": name, "t": stamp}; db.setdefault("nick", {})[uid] = name; db_save(db)
        try: await lv_send(it.guild, f"🤝 **외교관 등록** — {it.user.mention} · **[{ally}] {name}** · CID `{uid}`\n잘못 등록된 경우 `/외교관해제`로 정리하세요.")
        except Exception: log.exception("외교관 등록 알림 실패")
        await it.followup.send(f"✅ 외교관 등록 완료! 별명: **{nick}** · 역할: **{role.name}**" + ("\n" + "\n".join(notes) if notes else ""), ephemeral=True)


class DipView(CtxView):
    def __init__(self): super().__init__(timeout=None)

    @discord.ui.button(label="외교관 등록", style=discord.ButtonStyle.success, custom_id="dip_btn")
    async def go(self, it: discord.Interaction, _):
        await it.response.send_modal(DipModal())


class VerifyView(CtxView):
    def __init__(self): super().__init__(timeout=None)

    @discord.ui.button(label="인증하기", style=discord.ButtonStyle.primary, custom_id="verify_btn")
    async def go(self, it: discord.Interaction, _):
        await it.response.send_modal(VerifyModal())


# ---------- 명단 이탈 의심: 관리자 확인 후 처리 ----------
def leave_load(): return store_get("leave", LEAVE_PATH, {"pending": {}, "ignored": [], "log": [], "mass_at": 0})
def leave_save(d): store_set("leave", LEAVE_PATH, d)


def lv_view(uid):
    v = discord.ui.View(timeout=None); gid = cur_gid()
    for act, label, style in (("role", "역할 회수", discord.ButtonStyle.secondary), ("kick", "킥(서버 추방)", discord.ButtonStyle.danger),
                              ("keep", "유지", discord.ButtonStyle.success), ("snooze", "보류 24h", discord.ButtonStyle.secondary)):
        v.add_item(discord.ui.Button(label=label, style=style, custom_id=f"lv:{act}:{gid}:{uid}"))
    return v


async def lv_send(guild, text, view=None):
    ch = bot.get_channel(ADMIN_CHANNEL_ID) if ADMIN_CHANNEL_ID else None
    if ch is not None and getattr(ch, "guild", None) and ch.guild.id != guild.id: ch = None      # 다른 서버의 채널이면 서버장 DM
    if ch is None:
        owner = guild.owner or await guild.fetch_member(guild.owner_id)
        ch = await owner.create_dm()
    return await ch.send(text, view=view, allowed_mentions=discord.AllowedMentions.none())


async def report_missing(guild, missing):
    lv = leave_load(); now = time.time(); pend = lv["pending"]; changed = False
    have = {u for u, _, _ in missing}
    for uid in [u for u in pend if u not in have]:                         # 명단에 복구되었거나 서버를 떠난 사람은 정리
        pend.pop(uid); changed = True
    todo = [(u, d, m) for u, d, m in missing if u not in lv["ignored"] and (u not in pend or (pend[u].get("snooze") and pend[u]["snooze"] <= now))]
    if len(todo) > MISSING_MAX:
        if now - lv.get("mass_at", 0) > 86400:
            lv["mass_at"] = now; changed = True
            names = ", ".join(f"{m.display_name}" for _, _, m in todo[:20]) + (" …" if len(todo) > 20 else "")
            await lv_send(guild, f"🚨 **대량 이탈 의심** — 인사 탭에서 한 번에 **{len(todo)}명**의 CID가 사라졌습니다. 시트 연결/인사 탭 편집 오류일 수 있어 개별 처리는 보류합니다.\n{names}\n시트를 확인한 뒤 `/명단새로고침`을 실행하세요.")
        if changed: leave_save(lv)
        return
    for uid, did, m in todo:
        info = db_load().get("nick", {}).get(uid) or m.display_name
        roles = ", ".join(r.name for r in m.roles if r.name in (MEMBER_ROLE, OFFICER_ROLE)) or "-"
        msg = await lv_send(guild, f"⚠️ **명단 이탈 의심** — {m.mention} (`{m.display_name}`)\n인증 닉네임 **{info}** · CID `{uid}` · 역할 {roles}\n인사 탭에서 이 CID를 찾지 못했습니다. 확인 후 처리해 주세요.", lv_view(uid))
        pend[uid] = {"did": did, "nick": info, "ch": msg.channel.id, "msg": msg.id, "at": now}; changed = True
        await asyncio.sleep(0.5)
    if changed: leave_save(lv)


async def lv_finish(entry, text):
    try:
        ch = bot.get_channel(entry["ch"]) or await bot.fetch_channel(entry["ch"])
        msg = await ch.fetch_message(entry["msg"]); await msg.edit(content=text, view=None)
    except Exception: log.warning("이탈 알림 메시지 수정 실패")


async def lv_execute(guild, uid, entry, act, by):
    """역할 회수 / 킥. 인증 기록을 해제해 다시 인증할 수 있게 한다. → 결과 문구"""
    m = guild.get_member(int(entry["did"])); notes = []
    if m:
        drop = [r for r in (discord.utils.get(guild.roles, name=MEMBER_ROLE), discord.utils.get(guild.roles, name=OFFICER_ROLE)) if r and r in m.roles]
        try:
            if drop: await m.remove_roles(*drop, reason=f"명단 이탈 확인({by})")
            if act == "role":
                try: await m.edit(nick=None, reason=f"명단 이탈 확인({by})")
                except discord.Forbidden: notes.append("별명 초기화 권한 없음")
            else: await m.kick(reason=f"명단 이탈 확인({by})")
        except discord.Forbidden: return None, "봇 권한이 부족합니다(대상이 관리자이거나 봇 역할보다 높은 역할)."
    db = db_load(); db["uid"].pop(uid, None); db.get("nick", {}).pop(uid, None); db_save(db)
    return ("킥(서버 추방)" if act == "kick" else "역할 회수"), "; ".join(notes)


class KickConfirm(discord.ui.View):
    def __init__(self, uid, by, gid): super().__init__(timeout=60); self.uid, self.by, self.gid = uid, by, gid

    async def interaction_check(self, it): set_gid(self.gid); return True

    @discord.ui.button(label="정말 킥하기", style=discord.ButtonStyle.danger)
    async def ok(self, it: discord.Interaction, _):
        await lv_apply(it, "kick", self.uid, self.gid, confirmed=True)


async def lv_apply(it, act, uid, gid, confirmed=False):
    guild = bot.get_guild(gid); lv = leave_load(); entry = lv["pending"].get(uid)
    if not entry: return await it.response.send_message("이미 처리되었거나 목록에 없는 건입니다.", ephemeral=True)
    by = it.user.display_name; stamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%m.%d %H:%M")
    if act in ("role", "kick"):
        if roster.by_uid.get(uid):
            lv["pending"].pop(uid); leave_save(lv); await lv_finish(entry, f"✅ 명단에 다시 등록되어 처리하지 않았습니다. (CID `{uid}` · {entry['nick']})")
            return await it.response.send_message("이미 인사 탭에 있는 CID라 처리하지 않았습니다.", ephemeral=True)
        if act == "kick" and not confirmed:
            return await it.response.send_message(f"**{entry['nick']}** 님을 서버에서 추방합니다. 다시 들어오려면 초대 링크와 재인증이 필요합니다. 진행할까요?", view=KickConfirm(uid, by, gid), ephemeral=True)
        await it.response.defer(ephemeral=True)
        label, note = await lv_execute(guild, uid, entry, act, by)
        if label is None: return await it.followup.send(f"❌ {note}", ephemeral=True)
        done = f"🧹 **{label}** 처리 완료 — {entry['nick']} · CID `{uid}` · 처리자 {by} · {stamp}" + (f" ({note})" if note else "")
        lv["pending"].pop(uid); lv["log"].append({"t": stamp, "uid": uid, "nick": entry["nick"], "act": label, "by": by, "did": entry["did"]}); lv["log"] = lv["log"][-300:]; leave_save(lv)
        await lv_finish(entry, done); return await it.followup.send(done, ephemeral=True)
    if act == "keep":
        lv["ignored"].append(uid); lv["pending"].pop(uid); lv["log"].append({"t": stamp, "uid": uid, "nick": entry["nick"], "act": "유지", "by": by, "did": entry["did"]}); leave_save(lv)
        await lv_finish(entry, f"✅ **유지** — {entry['nick']} · CID `{uid}` · 처리자 {by} · {stamp}")
    else:
        entry["snooze"] = time.time() + 86400; leave_save(lv)
        await lv_finish(entry, f"⏸️ **24시간 보류** — {entry['nick']} · CID `{uid}` · 처리자 {by} · {stamp} (내일 다시 알립니다)")
    await it.response.send_message("처리했습니다.", ephemeral=True)


async def on_interaction(it: discord.Interaction):
    cid = (it.data or {}).get("custom_id", "") if it.type == discord.InteractionType.component else ""
    if cid.startswith("pm:"):
        try: await pm_decide(it, cid)
        except Exception:
            log.exception("관리자 인증 처리 오류")
            try: await it.response.send_message("처리 중 오류가 발생했습니다. 로그를 확인하세요.", ephemeral=True)
            except Exception: pass
        return
    if not cid.startswith("lv:"): return
    try:
        parts = cid.split(":"); act, uid = parts[1], parts[-1]
        gid = int(parts[2]) if len(parts) == 4 else GUILD_ID          # 예전 버튼(서버 ID 없음)은 홈 서버
        set_gid(gid); g = bot.get_guild(gid); mem = g.get_member(it.user.id) if g else None
        if not (mem and (mem.guild_permissions.administrator or (gid, it.user.id) in PERM)):
            return await it.response.send_message("관리자만 처리할 수 있습니다.", ephemeral=True)
        await lv_apply(it, act, uid, gid)
    except Exception:
        log.exception("이탈 처리 오류")
        try: await it.response.send_message("처리 중 오류가 발생했습니다. 로그를 확인하세요.", ephemeral=True)
        except Exception: pass


async def sync_members(guild: discord.Guild):
    """인증된 사람들을 명단(캐릭터 ID)과 대조해 별명/역할을 맞춘다."""
    mem_role = discord.utils.get(guild.roles, name=MEMBER_ROLE)
    off_role = discord.utils.get(guild.roles, name=OFFICER_ROLE)
    res = {"nick": 0, "role": 0, "removed": 0, "fail": 0, "total": 0}
    db = db_load(); nicks = db.setdefault("nick", {}); nchg = False; missing = []
    dips = db.get("dip", {})
    for uid, did in list(db["uid"].items()):
        if uid in dips: continue                                          # 외교관은 명단(인사 탭) 대조 대상이 아님
        m = guild.get_member(int(did))
        if m is None: continue
        res["total"] += 1
        info = roster.by_uid.get(uid)
        if info and info.get("name") and nicks.get(uid) != info["name"]: nicks[uid] = info["name"]; nchg = True     # 인증 시트에 적을 닉네임(명단 기준)
        try:
            if info is None:
                if roster.by_uid: missing.append((uid, str(did), m))        # 명단 이탈 의심 → 관리자 확인 후 처리(자동 회수 안 함)
                continue
            is_off = info["title"].replace(" ", "") in OFFICER_TITLES
            add = [r for r in ((mem_role,) + ((off_role,) if is_off else ())) if r and r not in m.roles]
            drop = [off_role] if (not is_off and off_role and off_role in m.roles) else []
            if add: await m.add_roles(*add, reason="명단 동기화"); await asyncio.sleep(0.4)
            if drop: await m.remove_roles(*drop, reason="명단 동기화"); await asyncio.sleep(0.4)
            if add or drop: res["role"] += 1
            if info["job"] and info["name"]:
                want = make_nick(info["job"], info["name"])
                if m.nick != want:
                    try: await m.edit(nick=want, reason="명단 동기화"); res["nick"] += 1; await asyncio.sleep(0.4)
                    except discord.Forbidden: pass
        except discord.Forbidden: res["fail"] += 1
        except Exception: log.exception("동기화 실패: %s", uid); res["fail"] += 1
    if nchg:
        d2 = db_load(); d2.setdefault("nick", {}).update({k: v for k, v in nicks.items() if k in d2["uid"]}); db_save(d2)
    try: await report_missing(guild, missing)
    except Exception: log.exception("명단 이탈 알림 실패")
    log.info("동기화: %s", res)
    return res


BOT_CAT = "봇"


async def bot_cat(g):
    """서버 채널 목록 맨 아래에 있는 '봇' 카테고리(없으면 만든다). 관리자가이드·명령어로그·후원 채널을 여기에 둔다. 만들 수 없으면 None."""
    cat = discord.utils.get(g.categories, name=BOT_CAT)
    try:
        if cat is None: cat = await g.create_category(BOT_CAT, reason="봇 채널 모음(최하단)")
        if cat.position != len(g.categories) - 1: await cat.edit(position=len(g.categories) - 1)
    except Exception as e: log.warning("'봇' 카테고리 준비 실패 서버=%s: %r", g.id, e); return cat if cat else None
    return cat


async def to_bottom(g, ch):
    """채널을 '봇' 카테고리(최하단)로 옮긴다. 이미 있으면 그대로, 실패해도 무시."""
    try:
        cat = await bot_cat(g)
        if cat is not None and ch is not None and getattr(ch, "category", None) != cat: await ch.edit(category=cat, sync_permissions=False)
    except Exception as e: log.warning("채널을 최하단으로 옮기지 못했습니다 서버=%s: %r", g.id, e)


LOG_CH = "명령어로그"
_log_fail = {}                                                  # 서버 ID → 채널 생성 실패 시각(1시간 동안 재시도 안 함)


async def log_channel(g, create=True):
    """#명령어로그 채널(관리자·봇만 볼 수 있음). 없으면 자동으로 만든다. 만들 수 없으면 None."""
    me = g.me
    ch = discord.utils.get(g.text_channels, name=LOG_CH)
    if ch is None and create and me is not None and time.time() - _log_fail.get(g.id, 0) > 3600:
        topic = "명령어 사용 기록 (서버 관리자만 볼 수 있습니다)"
        try:
            ow = {g.default_role: discord.PermissionOverwrite(view_channel=False), g.me: discord.PermissionOverwrite(view_channel=True, send_messages=True)}
            ch = await g.create_text_channel(LOG_CH, overwrites=ow, topic=topic, category=await bot_cat(g), reason="명령어 로그 채널 자동 생성")
        except Exception as e1:
            try:                                                   # 권한 덮어쓰기가 거부되면 일반 채널로 만든 뒤 비공개 설정을 따로 시도
                ch = await g.create_text_channel(LOG_CH, topic=topic, category=await bot_cat(g), reason="명령어 로그 채널 자동 생성")
                try: await ch.set_permissions(g.default_role, view_channel=False)
                except Exception as e3: log.warning("로그 채널 비공개 설정 실패(봇에 '역할 관리' 권한 필요) %s: %r", g.id, e3)
            except Exception as e2:
                _log_fail[g.id] = time.time(); log.warning("명령어 로그 채널 생성 실패 서버=%s 1차=%r 2차=%r (봇에 '채널 관리' 권한이 있는지 확인)", g.id, e1, e2); return None
    if ch is not None and getattr(ch, "category", None) is None: await to_bottom(g, ch)       # 예전에 위쪽에 만들어진 채널도 최하단으로
    return ch if ch and me is not None and ch.permissions_for(me).send_messages else None


async def cmd_log(it: discord.Interaction):
    """명령어 사용 기록을 #명령어로그에 남긴다. 실패해도 명령 실행에는 영향 없음."""
    try:
        ch = await log_channel(it.guild)
        if ch is None or it.command is None: return
        opts = " ".join(f"{k}={str(v)[:40]}" for k, v in (it.namespace.__dict__.items() if it.namespace else []) if v not in (None, ""))
        name = it.command.qualified_name
        await ch.send(f"🧾 <t:{int(time.time())}:T> **/{name}**" + (f" `{opts[:200]}`" if opts else "") + f" — {discord.utils.escape_markdown(it.user.display_name)} (`{it.user.id}`) · #{getattr(it.channel, 'name', '-')}",
                      allowed_mentions=discord.AllowedMentions.none())
    except Exception: log.warning("명령어 로그 기록 실패", exc_info=True)


GUIDE_CH, GUIDE_VER = "관리자가이드", 1
APP_ID = os.environ.get("APP_ID", "1563493671473520740")
GUIDE_PATH = Path(os.environ.get("GUIDE_PATH", "guide.json"))


def guide_pages():
    inv = f"https://discord.com/oauth2/authorize?client_id={APP_ID}&permissions=2550254672&scope=bot+applications.commands"
    return [
        "📘 **관리자 가이드북 ①  처음 시작하기**\n"
        "이 채널은 서버 관리자와 봇만 볼 수 있습니다.\n\n"
        "**1. 동맹 등록**: `/동맹등록`으로 서버 번호(숫자)와 동맹명을 신청합니다. 제작자가 승인하면 `동맹원`·`간부`·`외교관` 역할과 `#인증` 채널이 자동으로 만들어집니다. (등록은 한 번만, 이후 수정은 `/동맹수정`)\n"
        "**2. 명단 올리기**: 게임 명단 엑셀(.xlsx)을 `/명단업데이트`로 올립니다. 게임 엑셀은 매일 오전 8시에 갱신됩니다. 명단에서 빠진 사람은 자동으로 이탈자로 옮겨집니다.\n"
        "**3. 인증 안내**: 동맹원은 `#인증` 채널에서 `/인증`으로 CID 인증을 합니다(5회 실패하면 잠김). 인증하면 `[직책]닉네임`으로 별명이 바뀌고, 직책에 따라 `동맹원`/`간부` 역할이 부여됩니다.\n"
        "**4. 봇 권한**: 채널 관리·역할 관리·별명 관리가 필요합니다. 봇 역할은 `동맹원`·`간부`·`외교관` 역할보다 위에 있어야 합니다.",

        "📘 **관리자 가이드북 ②  관리 페이지 (`/관리`)**\n"
        "`/관리`를 입력하면 나에게만 보이는 링크가 나옵니다(본인 전용 · 6시간 유효 · 공유 금지). 간부는 쓸 수 없고 관리자만 쓸 수 있습니다. 저장하면 바로 DB에 반영되고, 화면은 4초마다 자동으로 갱신됩니다.\n\n"
        "• **톡방 참여자**: 추가·삭제·이름 수정. 명단에 없는 이름은 ⚠️로 표시됩니다.\n"
        "• **동맹원 추가**: 게임 엑셀에 아직 없는 사람을 CID·닉네임으로 직접 추가(엑셀에 반영되기 전 3일간 유지).\n"
        "• **부캐**: 본캐를 고르고 부캐 닉네임을 쉼표로 넣으면 명단과 대조해 연결합니다. 오타·없는 닉네임은 제외하고 이유를 알려 줍니다.\n"
        "• **관리자 권한**: 디스코드 사용자 ID로 공동관리자를 추가/제거합니다. 본인과 제작자의 권한은 뺄 수 없습니다.\n"
        "• **명령어 권한**: 역할별 명령어 표, 그리고 **간부 명령어를 직책별로 제한**하는 화면(`전체`를 누르면 간부 전체 허용).\n"
        "• **인증 횟수 제한**: 실패 횟수가 쌓인 계정과 잠긴 계정을 보고 바로 초기화합니다.",

        "📘 **관리자 가이드북 ③  권한 구조**\n"
        "• **동맹원**: `/인증` `/부캐` `/도감` `/추천덱` `/후원` `/명령어`\n"
        "• **간부**(`간부` 역할): 위 + `/인사` `/공성등록` `/공성` `/공성취소` `/공성예약` `/공성예약취소` `/정철설정` `/동맹수정` `/동맹원추가`\n"
        "• **관리자**(디스코드 서버 관리자, 또는 `/관리`에서 부여받은 사람): 위 전부 + `/관리` `/인증해제` `/명단업데이트` `/명단새로고침` `/외교관해제` `/부캐`의 본캐 지정\n"
        "• 간부 명령어는 `/관리` → 명령어 권한에서 **직책별로 제한**할 수 있습니다. 제한에 걸리면 \"권한이 없습니다.\"가 표시됩니다. 관리자는 제한과 관계없이 항상 사용 가능합니다.\n"
        "• 모든 명령어 사용 기록은 `#명령어로그`(관리자·봇만 열람)에 남습니다.",

        "📘 **관리자 가이드북 ④  자주 겪는 문제**\n"
        "• **채널·역할이 안 만들어져요** → 봇 역할에 `채널 관리`·`역할 관리` 권한이 있는지 확인하세요. 아래 초대 링크를 다시 열어 승인하면 봇을 내보내지 않고 권한만 갱신됩니다.\n"
        f"  {inv}\n"
        "• **인증이 잠겼어요** → `/관리` → 인증 횟수 제한에서 초기화하거나 `/인증해제`(CID)를 사용합니다.\n"
        "• **CID가 명단에 없다고 해요** → 닉네임·CID 오타를 확인하고, 신규 가입자는 `/관리` → 동맹원 추가로 먼저 넣어 주세요.\n"
        "• **링크가 만료됐어요** → `/관리`·`/인사`를 다시 입력하면 새 링크가 나옵니다.\n"
        "• **한 사람이 부계정이 여럿이에요** → `/관리` → 부캐에서 본캐에 연결합니다.\n"
        "• **봇이 응답이 없어요** → 잠시 후 다시 시도하고, 계속되면 제작자에게 문의하세요.",
    ]


async def guide_post(g):
    """#관리자가이드(관리자·봇만 열람)를 만들고 가이드북을 올린다. 버전이 바뀌면 이전 글을 지우고 새로 올린다. 만들 수 없으면 None."""
    me = g.me
    ch = discord.utils.get(g.text_channels, name=GUIDE_CH)
    if ch is None and me is not None and time.time() - _log_fail.get(("g", g.id), 0) > 3600:
        topic = "관리자 가이드북 (서버 관리자만 볼 수 있습니다)"
        try:
            ch = await g.create_text_channel(GUIDE_CH, overwrites={g.default_role: discord.PermissionOverwrite(view_channel=False), g.me: discord.PermissionOverwrite(view_channel=True, send_messages=True)}, topic=topic, category=await bot_cat(g), reason="관리자 가이드북 채널")
        except Exception as e1:
            try:
                ch = await g.create_text_channel(GUIDE_CH, topic=topic, category=await bot_cat(g), reason="관리자 가이드북 채널")
                try: await ch.set_permissions(g.default_role, view_channel=False)
                except Exception as e3: log.warning("가이드 채널 비공개 설정 실패 %s: %r", g.id, e3)
            except Exception as e2:
                _log_fail[("g", g.id)] = time.time(); log.warning("관리자 가이드 채널 생성 실패 서버=%s 1차=%r 2차=%r", g.id, e1, e2); return None
    if ch is not None and getattr(ch, "category", None) is None: await to_bottom(g, ch)
    if ch is None or me is None or not ch.permissions_for(me).send_messages: return None
    def st():
        tok = _CTX.set(g.id)
        try: return dict(store_get("guide", GUIDE_PATH, {}) or {})
        finally: _CTX.reset(tok)
    cur = await asyncio.to_thread(st)
    if cur.get("ver") == GUIDE_VER and cur.get("ch") == getattr(ch, "id", None): return ch
    for mid in cur.get("ids", []):                                     # 예전 가이드 글 정리
        try: await (await ch.fetch_message(mid)).delete()
        except Exception: pass
    ids = []
    for page in guide_pages():
        m = await ch.send(page, allowed_mentions=discord.AllowedMentions.none()); ids.append(m.id)
    def sv():
        tok = _CTX.set(g.id)
        try: store_set("guide", GUIDE_PATH, {"ver": GUIDE_VER, "ch": getattr(ch, "id", None), "ids": ids})
        finally: _CTX.reset(tok)
    await asyncio.to_thread(sv); return ch


class GuildTree(app_commands.CommandTree):
    """모든 명령 실행 전에: ① 처리할 서버를 정하고 ② 허가되지 않은 서버는 /동맹등록 외에는 막는다."""
    async def interaction_check(self, it: discord.Interaction):
        set_gid(it.guild_id)
        if it.guild_id is None:
            await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True); return False
        if licensed(it.guild_id) or (it.command and it.command.name in ("동맹등록", "명령어")):
            if it.command is not None: asyncio.create_task(cmd_log(it))
            return True
        await it.response.send_message("🔒 이 서버는 제작자의 승인이 필요합니다. `/동맹등록`으로 신청해 주세요.", ephemeral=True); return False


class Bot(discord.Client):
    def __init__(self):
        intents = discord.Intents.default(); intents.members = True
        super().__init__(intents=intents); self.tree = GuildTree(self)

    async def setup_hook(self):
        if SB_ON:                                                  # 홈 서버도 DB 명단(시트로 제어하지 않음)
            try: n = await asyncio.to_thread(roster_load_db, GUILD_ID); log.info("DB에서 명단 %d명 로드", n)
            except Exception as e: log.warning("DB 명단 로드 실패: %s", e)
        elif SHEET_ID and ROSTER_GID:
            try: roster.load_csv_text(fetch_sheet_csv()); log.info("구글 시트(인사 탭)에서 명단 %d명 로드", len(roster.by_uid))
            except Exception as e: log.warning("시트 명단 로드 실패, 엑셀 사용: %s", e)
            refresh_talk()
        await asyncio.to_thread(lambda: (db_load(), jc_load()))      # 시트에 저장된 인증·정철 상태를 미리 불러옴
        await start_web()
        self.add_view(VerifyView()); self.add_view(DipView())
        g = discord.Object(GUILD_ID)
        if OPEN_SERVERS:                                   # 다른 서버에서도 명령이 보이게 전역 등록(허가 안 된 서버는 /동맹등록만 가능)
            self.tree.clear_commands(guild=g); await self.tree.sync(guild=g); await self.tree.sync()
        else:
            self.tree.copy_global_to(guild=g)
            self.tree.remove_command("동맹등록", guild=g)          # 우리 서버는 이미 등록됨 → 메뉴에서 숨김
            await self.tree.sync(guild=g)

    async def on_ready(self):
        log.info("로그인: %s", self.user)
        if SYNC_MINUTES and not sync_loop.is_running(): sync_loop.start()
        if not jc_loop.is_running(): jc_loop.start()
        if not donate_loop.is_running(): donate_loop.start()
        hg = self.get_guild(GUILD_ID)
        if hg: await verify_channel(hg); await log_channel(hg); await guide_post(hg)                                # 우리 서버도 다른 서버와 똑같이 #인증 채널 사용(없으면 생성)
        try: await asyncio.to_thread(perm_refresh); log.info("승인된 관리자 %d명", len(PERM))
        except Exception: log.exception("승인 목록 로드 실패")

    async def on_member_join(self, m: discord.Member):
        if not licensed(m.guild.id) or m.bot: return
        set_gid(m.guild.id)
        ch = await verify_channel(m.guild)
        try:
            if ch: await ch.send(f"{m.mention} 환영합니다! 채팅창에 `/인증` 을 입력하면 나에게만 보이는 인증 창이 열립니다. CID · 직업 · 닉네임을 입력해 주세요.")
            else: await m.send("환영합니다! 아래 **인증하기** 버튼을 눌러 CID · 직업 · 닉네임을 입력해 주세요. (서버에서 `/인증` 으로도 열 수 있습니다)", view=VerifyView())
        except discord.Forbidden: log.warning("안내 메시지 전송 실패: %s", m)


VERIFY_CH = "인증"


async def verify_channel(g):
    """서버의 인증 안내 채널 = #인증 (홈 서버 포함 모든 서버 동일. 없으면 만든다). 만들 수 없으면 None."""
    me = g.me
    ok = lambda c: me is not None and c.permissions_for(me).send_messages
    c = discord.utils.get(g.text_channels, name=VERIFY_CH)
    if c is None:
        try:
            ow = {g.default_role: discord.PermissionOverwrite(view_channel=True, send_messages=False, use_application_commands=True),
                  g.me: discord.PermissionOverwrite(view_channel=True, send_messages=True)}
            c = await g.create_text_channel(VERIFY_CH, overwrites=ow, topic="처음 오셨다면 /인증 을 입력해 CID 인증을 해 주세요", reason="인증 안내 채널")
        except Exception: log.warning("인증 채널을 만들지 못했습니다(봇에 '채널 관리' 권한 필요): %s", g.id)
    return c if c and ok(c) else None


bot = Bot()
bot.event(on_interaction)


@tasks.loop(minutes=max(SYNC_MINUTES, 1))
async def sync_loop():
    set_gid(GUILD_ID)
    if SB_ON:
        try: await asyncio.to_thread(perm_refresh)         # 새로 승인/해제된 서버 반영
        except Exception: log.exception("승인 목록 갱신 실패")
        for gid in [GUILD_ID] + sorted(LICENSED - {GUILD_ID}):   # 홈 서버를 포함한 모든 서버: DB 명단
            try:
                set_gid(gid); g = bot.get_guild(gid)
                if g is None: continue
                await asyncio.to_thread(roster_load_db, gid)
                if gid == GUILD_ID: write_html()
                await sync_members(g)
            except Exception: log.exception("서버 %s 동기화 실패", gid)
        set_gid(GUILD_ID); return
    try:                                                   # (DB 미사용 환경) 엑셀/구글 시트
        if SHEET_ID and ROSTER_GID:
            text = await asyncio.to_thread(fetch_sheet_csv); roster.load_csv_text(text)
            await asyncio.to_thread(refresh_talk)
        else: roster.load()
        write_html()
        g = bot.get_guild(GUILD_ID)
        if g: await sync_members(g)
    except Exception: log.exception("주기 동기화 실패")
    set_gid(GUILD_ID)


admin = app_commands.default_permissions(manage_roles=True)


@bot.tree.command(name="외교관패널", description="외교관 등록 버튼 패널을 이 채널에 게시합니다")
@admin
async def dip_panel(it: discord.Interaction):
    await it.response.defer(ephemeral=True)
    try: await it.channel.send("**외교관 등록**\n외교 목적으로 방문하신 분은 버튼을 눌러 동맹명 · CID · 닉네임을 입력해 주세요. `[동맹명] 닉네임` 별명과 외교관 역할이 자동으로 설정됩니다.", view=DipView())
    except discord.Forbidden: return await it.followup.send("❌ 이 채널에서 봇이 메시지를 보낼 수 없습니다. 채널 권한을 확인해 주세요.", ephemeral=True)
    await it.followup.send("게시했습니다.", ephemeral=True)


@bot.tree.command(name="외교관해제", description="[관리자] 외교관 등록을 해제합니다(역할 회수, 별명 초기화, 재등록 가능)")
@admin
@app_commands.describe(cid="해제할 외교관의 CID")
async def dip_remove(it: discord.Interaction, cid: str):
    if not is_admin(it): return await it.response.send_message("관리자만 사용할 수 있습니다.", ephemeral=True)
    db = db_load(); uid = Roster.norm_uid(cid); d = db.get("dip", {}).pop(uid, None)
    if d is None: return await it.response.send_message("외교관으로 등록된 CID가 아닙니다.", ephemeral=True)
    did = db["uid"].pop(uid, None); db.get("nick", {}).pop(uid, None); db_save(db)
    m = it.guild.get_member(int(did)) if did else None; note = ""
    if m:
        role = discord.utils.get(it.guild.roles, name=DIPLOMAT_ROLE)
        try:
            if role and role in m.roles: await m.remove_roles(role, reason=f"외교관 해제({it.user.display_name})")
            await m.edit(nick=None, reason=f"외교관 해제({it.user.display_name})")
        except discord.Forbidden: note = " (역할/별명 변경 권한 부족 — 직접 정리해 주세요)"
    await it.response.send_message(f"**[{d['ally']}] {d['name']}** 외교관 등록을 해제했습니다.{note}", ephemeral=True)


@bot.tree.command(name="인증", description="CID 인증 창을 엽니다(나에게만 보임)")
async def cmd_verify(it: discord.Interaction):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    await it.response.send_message("**연맹 인증**\n버튼을 눌러 CID · 직업 · 닉네임을 입력하면 별명과 역할이 자동으로 설정됩니다.", view=VerifyView(), ephemeral=True)


@bot.tree.command(name="인증패널", description="인증 버튼 패널을 이 채널에 게시합니다")
@admin
async def panel(it: discord.Interaction):
    await it.response.defer(ephemeral=True)
    try:
        await it.channel.send("**연맹 인증**\n버튼을 눌러 CID · 직업 · 닉네임을 입력하면 별명과 역할이 자동으로 설정됩니다.", view=VerifyView())
    except discord.Forbidden:
        return await it.followup.send("❌ 이 채널에서 봇이 메시지를 보낼 수 없습니다. 채널 설정 → 권한에서 봇(또는 봇 역할)에 '채널 보기'와 '메시지 보내기'를 허용해 주세요.", ephemeral=True)
    except Exception as e:
        log.exception("패널 게시 실패"); return await it.followup.send(f"❌ 오류: {e}", ephemeral=True)
    await it.followup.send("게시했습니다.", ephemeral=True)


# ---------- 제작자 허가(관리자 인증) ----------
OPEN_SERVERS = os.environ.get("OPEN_SERVERS", "0").strip() == "1"   # 1이면 다른 서버에서도 봇 명령 표시(제작자 승인 필요)
OWNER_ID = int(os.environ.get("OWNER_ID", "0") or 0)      # 제작자 디스코드 ID(비우면 봇 소유자)
PERM = set()                                               # 승인된 (서버 ID, 사용자 ID)


def perm_refresh():
    global PERM, LICENSED
    if not SB_ON: return
    rows = _sb_req("GET", "sam_perm?status=eq.approved&select=guild_id,discord_id") or []
    PERM = {(int(r["guild_id"]), int(r["discord_id"])) for r in rows}
    LICENSED = {g for g, _ in PERM}


async def perm_owner():
    if OWNER_ID: return await bot.fetch_user(OWNER_ID)
    return (await bot.application_info()).owner


def perm_view(gid, uid):
    v = discord.ui.View(timeout=None)
    v.add_item(discord.ui.Button(label="승인", style=discord.ButtonStyle.success, custom_id=f"pm:ok:{gid}:{uid}"))
    v.add_item(discord.ui.Button(label="거절", style=discord.ButtonStyle.danger, custom_id=f"pm:no:{gid}:{uid}"))
    return v


def ally_set(gid, name, server):
    """그 서버의 동맹 정보 저장(sam_store + sam_alli). 스레드에서 호출."""
    tok = _CTX.set(int(gid))
    try: store_set("alliance", ALLY_PATH, {"name": name, "server": str(int(server))})
    finally: _CTX.reset(tok)


async def pm_setup_guild(gid):
    """승인된 서버에 기본 역할(동맹원·간부·외교관)을 만든다. → 안내 문구"""
    g = bot.get_guild(gid)
    if g is None: return "봇이 서버에 없어 역할을 만들지 못했습니다."
    made, fail = [], []
    for nm in (MEMBER_ROLE, OFFICER_ROLE, DIPLOMAT_ROLE):
        if discord.utils.get(g.roles, name=nm): continue
        try: await g.create_role(name=nm, reason="제작자 승인 — 기본 역할 생성"); made.append(nm)
        except Exception: fail.append(nm)
    vc = await verify_channel(g); await log_channel(g); await guide_post(g)
    out = (f"역할 생성: {', '.join(made)}" if made else "역할은 이미 있습니다.") + (f" / 인증 채널: #{vc.name}" if vc else " / 인증 채널을 만들지 못했습니다(봇에 '채널 관리' 권한 필요)")
    return out + (f" / 생성 실패(봇에 '역할 관리' 권한 필요): {', '.join(fail)}" if fail else "")


async def pm_decide(it: discord.Interaction, cid):
    owner = await perm_owner()
    if it.user.id != owner.id: return await it.response.send_message("제작자만 처리할 수 있습니다.", ephemeral=True)
    _, act, gid, uid = cid.split(":", 3); gid, uid = int(gid), int(uid); ok = act == "ok"
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    rows = await asyncio.to_thread(_sb_req, "GET", f"sam_perm?guild_id=eq.{gid}&discord_id=eq.{uid}&select=ally_name,game_server,guild_name")
    row = rows[0] if rows else {}
    await asyncio.to_thread(_sb_req, "PATCH", f"sam_perm?guild_id=eq.{gid}&discord_id=eq.{uid}",
                            {"status": "approved" if ok else "rejected", "decided_at": now, "decided_by": it.user.id}, "return=minimal")
    extra = ""
    if ok:
        PERM.add((gid, uid)); LICENSED.add(gid)
        if row.get("ally_name") and row.get("game_server"): await asyncio.to_thread(ally_set, gid, row["ally_name"], row["game_server"])
        extra = await pm_setup_guild(gid)
    else:
        PERM.discard((gid, uid))
        if not any(g == gid for g, _ in PERM): LICENSED.discard(gid)
    head = (it.message.content or "").split("\n")[0]
    await it.response.edit_message(content=f"{head}\n{'✅ 승인됨' if ok else '⛔ 거절됨'} — {it.user.display_name}" + (f"\n{extra}" if extra else ""), view=None)
    try:
        u = await bot.fetch_user(uid)
        await u.send((f"✅ **{row.get('guild_name') or '서버'}** 사용이 **승인**되었습니다. ({row.get('game_server')}서버 · {row.get('ally_name')})\n"
                      f"서버에서 봇 역할을 '{MEMBER_ROLE}'·'{OFFICER_ROLE}' 역할보다 위로 올린 뒤, `/명단업데이트`로 인사 명단(.xlsx)을 올려 주세요. 이후 동맹원은 `/인증패널`의 버튼으로 CID 인증을 하면 됩니다.")
                     if ok else "⛔ 동맹 등록 신청이 **거절**되었습니다.")
    except Exception: log.warning("신청자 DM 실패: %s", uid)


class AllyApplyModal(CtxModal, title="동맹 등록 신청"):
    server = discord.ui.TextInput(label="서버 번호 (숫자만 입력)", placeholder="예: 41", min_length=1, max_length=5)
    name = discord.ui.TextInput(label="동맹명", placeholder="예: 공성부대", min_length=1, max_length=12)
    reason = discord.ui.TextInput(label="사유 (선택)", placeholder="신청 사유, 직책 등", required=False, max_length=200, style=discord.TextStyle.paragraph)

    async def on_submit(self, it: discord.Interaction):
        sv, nm = self.server.value.strip(), self.name.value.strip()
        if not (sv.isascii() and sv.isdigit()): return await it.response.send_message("❌ 서버 번호는 숫자만 입력할 수 있습니다. (예: 41)", ephemeral=True)
        if not SB_ON: return await it.response.send_message("❌ DB(Supabase)가 연결되지 않아 신청을 받을 수 없습니다.", ephemeral=True)
        await it.response.defer(ephemeral=True); gid, uid = it.guild.id, it.user.id
        cur = await asyncio.to_thread(_sb_req, "GET", f"sam_perm?guild_id=eq.{gid}&discord_id=eq.{uid}&select=status")
        st = cur[0]["status"] if cur else None
        if st == "pending": return await it.followup.send("이미 신청되어 제작자의 승인을 기다리는 중입니다.", ephemeral=True)
        owner = await perm_owner(); now = datetime.datetime.now(datetime.timezone.utc).isoformat(); sv = str(int(sv))
        row = {"guild_id": gid, "discord_id": uid, "guild_name": it.guild.name, "user_name": str(it.user), "reason": self.reason.value[:200],
               "ally_name": nm, "game_server": sv, "status": "approved" if uid == owner.id else "pending",
               "requested_at": now, "decided_at": None, "decided_by": None}
        await asyncio.to_thread(_sb_req, "POST", "sam_perm?on_conflict=guild_id,discord_id", [row], "resolution=merge-duplicates,return=minimal")
        if uid == owner.id:
            PERM.add((gid, uid)); LICENSED.add(gid); await asyncio.to_thread(ally_set, gid, nm, sv); extra = await pm_setup_guild(gid)
            return await it.followup.send(f"제작자 본인이라 바로 등록되었습니다. ({sv}서버 · {nm})\n{extra}", ephemeral=True)
        try:
            await owner.send(f"🔐 **동맹 등록 신청**\n서버: **{it.guild.name}** (`{gid}`)\n동맹: **{sv}서버 · {nm}**\n신청자: {it.user.mention} ({it.user}) · `{uid}`\n사유: {self.reason.value or '(없음)'}",
                             view=perm_view(gid, uid), allowed_mentions=discord.AllowedMentions.none())
            await it.followup.send("📨 제작자에게 신청을 보냈습니다. 승인되면 DM으로 알려 드립니다.", ephemeral=True)
        except Exception:
            log.exception("제작자 알림 실패")
            await it.followup.send("신청은 접수됐지만 제작자에게 알림을 보내지 못했습니다. 제작자에게 직접 알려 주세요.", ephemeral=True)


# ---------- 후원 안내 ----------
DONATE_PATH = Path(os.environ.get("DONATE_PATH", "donate.json"))
DONATE_TEXT = "주에 한번씩 제작자의 서버비를 위해 열심히 일할 커피가 필요합니다.\n후원 한번씩 해주시면 감사하겠습니다!"
DONATE_MAX = 2 * 1024 * 1024                              # QR 이미지 최대 크기


DONATE_IMG_DEFAULT = Path(__file__).with_name("donate_qr.png")   # 저장소에 넣어 둔 기본 QR


def donate_get():
    """후원 안내(문구 + QR 이미지). 서버와 관계없이 홈 서버 저장소 한 곳에서 읽는다. 저장된 이미지가 없으면 기본 QR 파일."""
    tok = _CTX.set(GUILD_ID)
    try: d = dict(store_get("donate", DONATE_PATH, None) or {})
    finally: _CTX.reset(tok)
    if not d.get("img") and DONATE_IMG_DEFAULT.exists():
        d["img"] = base64.b64encode(DONATE_IMG_DEFAULT.read_bytes()).decode("ascii"); d["ext"] = ".png"
    return d


def donate_put(d):
    tok = _CTX.set(GUILD_ID)
    try: store_set("donate", DONATE_PATH, d)
    finally: _CTX.reset(tok)


@bot.tree.command(name="후원", description="제작자 서버비 후원 안내(카카오 QR)를 봅니다")
async def cmd_donate(it: discord.Interaction):
    await it.response.defer(ephemeral=True)
    d = await asyncio.to_thread(donate_get); text = d.get("text") or DONATE_TEXT
    if d.get("img"):
        f = discord.File(io.BytesIO(base64.b64decode(d["img"])), filename="donate" + d.get("ext", ".png"))
        return await it.followup.send(text, file=f, ephemeral=True)
    await it.followup.send(text + "\n\n(QR 이미지를 준비 중입니다.)", ephemeral=True)


@bot.tree.command(name="후원안내설정", description="[제작자] 후원 QR 이미지와 안내 문구를 설정합니다")
@app_commands.describe(이미지="카카오 QR 이미지(png/jpg)", 문구="안내 문구(비우면 기본 문구)")
async def cmd_donate_set(it: discord.Interaction, 이미지: discord.Attachment = None, 문구: str = ""):
    owner = await perm_owner()
    if it.user.id != owner.id: return await it.response.send_message("제작자만 사용할 수 있습니다.", ephemeral=True)
    await it.response.defer(ephemeral=True)
    d = await asyncio.to_thread(donate_get); d = dict(d)
    if 이미지 is not None:
        ext = Path(이미지.filename).suffix.lower()
        if ext not in (".png", ".jpg", ".jpeg", ".webp") or not (이미지.content_type or "").startswith("image/"):
            return await it.followup.send("❌ png/jpg 이미지만 올릴 수 있습니다.", ephemeral=True)
        if 이미지.size > DONATE_MAX: return await it.followup.send("❌ 이미지는 2MB 이하로 올려 주세요.", ephemeral=True)
        d["img"] = base64.b64encode(await 이미지.read()).decode("ascii"); d["ext"] = ext
    if 문구.strip(): d["text"] = 문구.strip().replace("\\n", "\n")[:800]
    await asyncio.to_thread(donate_put, d)
    f = discord.File(io.BytesIO(base64.b64decode(d["img"])), filename="donate" + d.get("ext", ".png")) if d.get("img") else None
    await it.followup.send("✅ 설정했습니다. 미리보기:\n\n" + (d.get("text") or DONATE_TEXT), file=f, ephemeral=True) if f else \
        await it.followup.send("✅ 설정했습니다. (이미지는 아직 없습니다)\n\n" + (d.get("text") or DONATE_TEXT), ephemeral=True)


KST = datetime.timezone(datetime.timedelta(hours=9))
DONATE_SENT = Path(os.environ.get("DONATE_SENT_PATH", "donate_sent.json"))


DONATE_CH = "후원"


async def donate_channel(g):
    """#후원 채널을 돌려준다. 없으면 만들고(읽기 전용), 만들 수 없으면 시스템 채널 → 쓸 수 있는 첫 텍스트 채널."""
    me = g.me
    ok = lambda c: me is not None and c.permissions_for(me).send_messages and c.permissions_for(me).attach_files
    ch = discord.utils.get(g.text_channels, name=DONATE_CH)
    if ch is None:
        try:
            ow = {g.default_role: discord.PermissionOverwrite(send_messages=False, add_reactions=True),
                  g.me: discord.PermissionOverwrite(send_messages=True, attach_files=True, view_channel=True)}
            ch = await g.create_text_channel(DONATE_CH, overwrites=ow, topic="제작자 서버비 후원 안내 (매주 토요일 19시)", category=await bot_cat(g), reason="후원 안내 채널")
        except Exception: log.warning("후원 채널을 만들지 못했습니다(봇에 '채널 관리' 권한 필요): %s", g.id)
    elif getattr(ch, "category", None) is None: await to_bottom(g, ch)
    if ch and ok(ch): return ch
    if g.system_channel and ok(g.system_channel): return g.system_channel
    return next((c for c in g.text_channels if ok(c)), None)


@tasks.loop(seconds=60)
async def donate_loop():
    """토요일 19:00(KST)에, 아직 안내받지 않은 서버에만 후원 안내를 딱 한 번 올린다(서버마다 1회)."""
    now = datetime.datetime.now(KST)
    if now.weekday() != 5 or not (now.hour == 19 and now.minute < 30): return
    tok = _CTX.set(GUILD_ID)
    try: sent = store_get("donate_sent", DONATE_SENT, None) or {}
    finally: _CTX.reset(tok)
    done = set(sent.get("guilds") or [])
    todo = [g for g in [GUILD_ID] + sorted(LICENSED - {GUILD_ID}) if g not in done]
    if not todo: return
    d = await asyncio.to_thread(donate_get); text = d.get("text") or DONATE_TEXT
    if not d.get("img"): log.warning("후원 QR 이미지가 없어 안내를 건너뜁니다"); return
    raw = base64.b64decode(d["img"])
    for gid in todo:
        g = bot.get_guild(gid); ch = await donate_channel(g) if g else None
        if not ch: log.warning("후원 안내 채널 없음: %s", gid); continue
        done.add(gid)                                                # 먼저 기록: 재시작해도 중복 전송 방지
        tok = _CTX.set(GUILD_ID)
        try: store_set("donate_sent", DONATE_SENT, {"guilds": sorted(done)})
        finally: _CTX.reset(tok)
        try: await ch.send(text, file=discord.File(io.BytesIO(raw), filename="donate" + d.get("ext", ".png")))
        except Exception: log.exception("후원 안내 전송 실패: %s", gid)


CMD_HELP = {
    "동맹원": [("/동맹등록", "서버 번호·동맹명을 등록(승인 전 서버는 제작자에게 신청)"),
             ("/인증", "CID 인증 창 열기(나에게만 보임)"),
             ("/부캐", "내 부계정을 닉네임으로 연결(닉네임1,닉네임2 …)"),
             ("/도감", "내 보유 장수·전법을 체크하는 개인 링크"),
             ("/추천덱", "내 보유 장수·전법으로 티어덱 5부대 편성(시즌 선택은 기억)"),
             ("/후원", "제작자 서버비 후원 안내(카카오 QR)"),
             ("/명령어", "내 역할로 쓸 수 있는 명령어 보기")],
    "간부": [("/인사", "인사데이터(상황판) 링크"),
           ("/공성등록", "공성 주둔지 등록"), ("/공성", "정철 개수·공성 예상 일정"), ("/공성취소", "등록한 공성 취소"),
           ("/공성예약", "공성 시각 예약(30분 전 간부 호출)"), ("/공성예약취소", "공성 예약 취소"),
           ("/정철설정", "정철 보유량 맞추기"), ("/동맹수정", "등록된 서버 번호·동맹명 수정"),
           ("/동맹원추가", "게임 엑셀(오전 8시 반영)에 아직 없는 동맹원을 직접 추가")],
    "관리자": [("/관리", "관리 페이지: 톡방 참여자·동맹원 추가·관리자 권한 부여(DB 실시간 연동)")],
}


@bot.tree.command(name="명령어", description="역할별로 사용할 수 있는 명령어를 알려 드립니다")
@app_commands.describe(역할="[제작자 전용] 어느 역할이 보는 명령어를 볼지(비우면 내 역할 기준)")
@app_commands.choices(역할=[app_commands.Choice(name="동맹원", value="동맹원"), app_commands.Choice(name="간부", value="간부")])
async def cmd_help(it: discord.Interaction, 역할: app_commands.Choice[str] = None):
    owner = await perm_owner()
    mine = "간부" if (is_staff(it) or it.user.id == owner.id) else "동맹원"
    role = 역할.value if (역할 and it.user.id == owner.id) else mine          # 역할 지정은 제작자만 가능
    lines = [f"📖 **{role}이(가) 사용할 수 있는 명령어**" + (" (제작자 조회)" if role != mine or (역할 and it.user.id == owner.id) else ""), "\n**[공통]**"]
    lines += [f"`{c}` — {d}" for c, d in CMD_HELP["동맹원"]]
    if role == "간부":
        lines.append("\n**[간부 전용]**"); lines += [f"`{c}` — {d}" for c, d in CMD_HELP["간부"] if it.user.id == owner.id or staff_ok(it, c)]
        if is_admin(it) or it.user.id == owner.id: lines.append("\n**[관리자 전용]**"); lines += [f"`{c}` — {d}" for c, d in CMD_HELP["관리자"]]
    await it.response.send_message("\n".join(lines), ephemeral=True)


def admin_only(f): return f          # 디스코드 기본 권한으로 숨기지 않고, 함수 안(is_admin)에서 판정(제작자가 승인한 관리자도 사용)


def is_admin(it: discord.Interaction):
    return isinstance(it.user, discord.Member) and (it.user.guild_permissions.administrator or (it.guild_id, it.user.id) in PERM)


def fetch_sheet_csv(gid=None):
    gid = gid or ROSTER_GID
    if not (SHEET_ID and gid): raise RuntimeError("SHEET_ID / ROSTER_GID가 .env에 설정되지 않았습니다.")
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    with urllib.request.urlopen(url, timeout=20) as r: return r.read().decode("utf-8-sig")


def refresh_talk():
    """톡방참여자 탭의 닉네임을 읽어 상황판 톡방 표시에 반영. 실패해도 기존 값 유지."""
    if not (SHEET_ID and TALK_GID): return None
    try:
        names = {c.strip() for row in csv.reader(io.StringIO(fetch_sheet_csv(TALK_GID))) for c in row if c.strip()} - {"멤버", "닉네임"}
        roster.talk = names; return len(names)
    except Exception as e:
        log.warning("톡방 탭 로드 실패: %s", e); return None


TALK_PATH = Path(os.environ.get("TALK_PATH", "talk.json"))
TALK_MAX, TALK_NAME_MAX = 2000, 40
_talk_lock = threading.Lock()


def talk_get():
    """톡방 참여자 이름 목록. 대시보드에서 한 번이라도 관리했으면 저장된 값, 아니면 시트 톡방참여자 탭(홈 서버만)."""
    d = store_get("talk", TALK_PATH, None)
    if isinstance(d, dict) and isinstance(d.get("names"), list): return list(d["names"]), d.get("date"), True
    return sorted(roster.talk), None, False


def talk_apply(body):
    """대시보드 톡방 관리 요청 처리 → (이름 목록, 날짜). 잘못된 요청은 ValueError."""
    def clean(xs):
        if not isinstance(xs, list): raise ValueError("이름 목록이 올바르지 않습니다.")
        out = []
        for x in xs:
            n = re.sub(r"\s+", " ", str(x)).strip()
            if not n: continue
            if len(n) > TALK_NAME_MAX: raise ValueError(f"이름이 너무 깁니다(최대 {TALK_NAME_MAX}자): {n[:12]}…")
            if n not in out: out.append(n)
        return out
    with _talk_lock:
        names, _, _ = talk_get(); cur = list(names); op = body.get("op")
        if op == "add": cur += [n for n in clean(body.get("names")) if n not in cur]
        elif op == "remove":
            rm = set(clean(body.get("names"))); cur = [n for n in cur if n not in rm]
        elif op == "rename":
            a, b = clean([body.get("from")]), clean([body.get("to")])
            if not a or not b: raise ValueError("이름을 입력해 주세요.")
            if a[0] not in cur: raise ValueError("목록에 없는 이름입니다.")
            if b[0] != a[0] and b[0] in cur: raise ValueError("이미 있는 이름입니다.")
            cur = [b[0] if n == a[0] else n for n in cur]
        elif op == "set": cur = clean(body.get("names"))
        else: raise ValueError("알 수 없는 요청입니다.")
        if len(cur) > TALK_MAX: raise ValueError(f"톡방 인원은 최대 {TALK_MAX}명까지입니다.")
        date = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%Y-%m-%d")
        store_set("talk", TALK_PATH, {"names": sorted(cur), "date": date})
        return sorted(cur), date


def render_html(tok=None):
    """현재 명단/톡방 데이터로 상황판 HTML 문자열을 만든다. 데이터가 없으면 None. tok이 있으면(간부 링크) 톡방 관리 탭이 켜진다."""
    if not roster.members: return None
    tpl = (Path(__file__).parent / "dashboard_template.html").read_text("utf-8")
    today = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%Y-%m-%d")
    names, tdate, managed = talk_get(); tset = set(names)
    D = [{**d, "t": 1 if d["n"] in tset else 0} for d in roster.members]
    cfg = {"rosterFile": roster.src or "명단", "rosterDate": today,
           "talkFile": ("대시보드 톡방 관리" if managed else "구글 시트 톡방참여자 탭") if names else "없음", "talkDate": tdate or today, "graceEnd": GRACE_END}
    if tok: cfg.update(edit=True, tok=tok, talkNames=sorted(names))
    body = tpl.replace("__D__", json.dumps(D, ensure_ascii=False)).replace("__CFG__", json.dumps(cfg, ensure_ascii=False))
    return '<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' + body + "</html>"


def write_html():
    """연맹_인사데이터.html 파일을 최신 데이터로 다시 만든다(파일로 공유/업로드할 때 사용)."""
    h = render_html()
    if h is None: return None
    out = Path(os.environ.get("HTML_OUT", "연맹_인사데이터.html"))
    out.write_text(h, "utf-8"); return out


async def web_index(request):
    if WEB_KEY and request.query.get("key") != WEB_KEY: return web.Response(status=403, text="접근 권한이 없습니다.")
    html = render_html()
    if html is None: return web.Response(status=503, text="명단 데이터를 아직 불러오지 못했습니다. 관리자에게 문의하세요.")
    return web.Response(text=html, content_type="text/html", headers={"Cache-Control": "no-store"})


# ---------- 도감(보유 장수·전법) 페이지 ----------
TIER_SECRET = os.environ.get("TIER_SECRET") or os.environ.get("DISCORD_TOKEN", "")
PUBLIC_URL = (os.environ.get("PUBLIC_URL") or os.environ.get("RENDER_EXTERNAL_URL") or f"http://localhost:{WEB_PORT}").rstrip("/")
HR_HOURS = int(os.environ.get("HR_HOURS", "12"))     # /인사 링크 유효시간


def hr_token(uid, exp, gid=GUILD_ID):
    return f"{uid}.{exp}.{gid}." + hmac.new(TIER_SECRET.encode(), f"hr:{uid}:{exp}:{gid}".encode(), hashlib.sha256).hexdigest()[:24]


def hr_parse(t):
    """인사 링크 검사 → 서버 ID 또는 None. (서버 ID가 없는 예전 링크는 홈 서버)"""
    p = (t or "").split(".")
    if len(p) == 3: uid, exp, gid, sig = p[0], p[1], GUILD_ID, p[2]; old = True
    elif len(p) == 4 and p[2].isdigit(): uid, exp, gid, sig = p[0], p[1], int(p[2]), p[3]; old = False
    else: return None
    if not exp.isdigit() or int(exp) < time.time(): return None
    want = hmac.new(TIER_SECRET.encode(), (f"hr:{uid}:{exp}" if old else f"hr:{uid}:{exp}:{gid}").encode(), hashlib.sha256).hexdigest()[:24]
    return gid if hmac.compare_digest(sig, want) else None


async def web_hr(request):
    gid = hr_parse(request.query.get("t")); set_gid(gid)
    if gid is None: return web.Response(status=403, text="링크가 만료되었거나 올바르지 않습니다. 디스코드에서 /인사 를 다시 입력해 주세요.")
    html = await asyncio.to_thread(lambda: (set_gid(gid), render_html(request.query.get("t")))[1])
    if html is None: return web.Response(status=503, text="명단 데이터를 아직 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.")
    return web.Response(text=html, content_type="text/html", headers={"Cache-Control": "no-store", "X-Robots-Tag": "noindex"})


async def web_talk(request):
    """대시보드 '톡방 관리' 탭의 저장 요청. 인사 링크(간부 전용 토큰)가 있어야 한다."""
    gid = hr_parse(request.query.get("t"))
    if gid is None: return web.json_response({"error": "링크가 만료되었습니다. 디스코드에서 /인사 를 다시 입력해 주세요."}, status=403)
    if request.content_length and request.content_length > 200_000: return web.json_response({"error": "요청이 너무 큽니다."}, status=413)
    try: body = await request.json()
    except Exception: return web.json_response({"error": "요청 형식이 올바르지 않습니다."}, status=400)
    if not isinstance(body, dict): return web.json_response({"error": "요청 형식이 올바르지 않습니다."}, status=400)
    try: names, date = await asyncio.to_thread(lambda: (set_gid(gid), talk_apply(body))[1])
    except ValueError as e: return web.json_response({"error": str(e)}, status=400)
    return web.json_response({"names": names, "date": date})
DEX_DIR = Path(os.environ.get("DEX_DIR", "dex"))


def _dex_sig(gid, cid): return hmac.new(TIER_SECRET.encode(), (f"dex:{cid}" if gid == GUILD_ID else f"dex:{gid}:{cid}").encode(), hashlib.sha256).hexdigest()[:20]


def tier_token(cid, gid=GUILD_ID):
    return f"{cid}.{_dex_sig(gid, cid)}" if gid == GUILD_ID else f"{gid}.{cid}.{_dex_sig(gid, cid)}"      # 홈 서버는 예전 형식 그대로


def tier_parse(t):
    """도감 링크 검사 → (서버 ID, CID) 또는 None"""
    p = (t or "").split(".")
    if len(p) == 2: gid, cid, sig = GUILD_ID, p[0], p[1]
    elif len(p) == 3 and p[0].isdigit(): gid, cid, sig = int(p[0]), p[1], p[2]
    else: return None
    return (gid, cid) if cid and hmac.compare_digest(sig, _dex_sig(gid, cid)) else None


def dex_path(cid): return DEX_DIR / f"{re.sub(r'[^0-9A-Za-z]', '', cid)}.json"


def _cell(r, i): return re.sub(r"\s*\n\s*", " ", str(r[i])).strip() if i < len(r) and r[i] is not None else ""


def parse_hd(rows):
    """시트 '무장' 탭 → {장수: {...}}  (열 순서: 시즌,나라,진영,성향,(합쳐진 칸),이름,병종,전법유형,전법특성,고유전법,확률,설명,무력,지력,통솔,선공)"""
    out = {}
    for r in rows[1:]:
        name = _cell(r, 5) or _cell(r, 4)
        if not name: continue
        out[name] = {"nation": _cell(r, 1), "pos": _cell(r, 2), "style": _cell(r, 3), "troop": _cell(r, 6), "sktype": _cell(r, 7),
                     "sktrait": _cell(r, 8), "sk": _cell(r, 9), "rate": _cell(r, 10), "desc": (str(r[11]).strip() if len(r) > 11 else ""),
                     "st": [_cell(r, i) for i in (12, 13, 14, 15)], "season": int(_cell(r, 0) or 1) if _cell(r, 0).isdigit() else 1}
    return out


def parse_skd(rows):
    """시트 '전법' 탭 → {전법: {...}}  (열 순서: 그림,이름,유형,특성,발동률,설명,등급,시즌)"""
    out = {}
    for r in rows[1:]:
        name = _cell(r, 1)
        if not name: continue
        out[name] = {"type": _cell(r, 2), "trait": _cell(r, 3), "rate": _cell(r, 4), "desc": _cell(r, 5), "grade": _cell(r, 6),
                     "season": int(_cell(r, 7)) if _cell(r, 7).isdigit() else 1}
    return out


_tab_cache = {}


def fetch_tab_rows(tab, ttl=300):
    """시트 탭을 이름으로 읽음(5분 캐시). 실패하면 이전 값, 없으면 None."""
    now = time.time(); c = _tab_cache.get(tab)
    if c and now - c[0] < ttl: return c[1]
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(tab)}"
        with urllib.request.urlopen(url, timeout=20) as r: text = r.read().decode("utf-8-sig")
        rows = list(csv.reader(io.StringIO(text)))
        if len(rows) < 2: raise RuntimeError("빈 탭")
        _tab_cache[tab] = (now, rows); return rows
    except Exception as e:
        log.warning("시트 탭 '%s' 읽기 실패: %s", tab, e); return c[1] if c else None


def tier_sheet_cfg():
    cfg = {}
    if not SHEET_ID: return cfg
    rows = fetch_tab_rows("무장")
    if rows:
        hd = parse_hd(rows)
        if hd: cfg["hd"] = hd
    rows = fetch_tab_rows("전법")
    if rows:
        skd = parse_skd(rows)
        if skd: cfg["skd"] = skd
    return cfg


async def web_tier(request):
    pr = tier_parse(request.query.get("t"))
    if not pr: return web.Response(status=403, text="링크가 올바르지 않습니다. 디스코드에서 /도감 으로 새 링크를 받으세요.")
    set_gid(pr[0]); cid = pr[1]
    info = roster.by_uid.get(cid) or {}
    dex = await asyncio.to_thread(store_get, f"dex_{cid}", dex_path(cid), {})
    cfg = await asyncio.to_thread(tier_sheet_cfg)
    cfg |= {"t": request.query["t"], "api": "/api/dex", "name": info.get("name", ""), "own": dex.get("own", []), "ownsk": dex.get("ownsk", [])}
    tpl = (Path(__file__).parent / "tier_template.html").read_text("utf-8")
    return web.Response(text=tpl.replace("__CFG__", json.dumps(cfg, ensure_ascii=False)), content_type="text/html", headers={"Cache-Control": "no-store"})


async def web_dex_save(request):
    try: b = await request.json()
    except Exception: return web.json_response({"ok": False, "error": "형식 오류"}, status=400)
    pr = tier_parse(b.get("t"))
    if not pr: return web.json_response({"ok": False, "error": "권한 없음"}, status=403)
    set_gid(pr[0]); cid = pr[1]
    clean = lambda v: sorted({str(x)[:30] for x in v if isinstance(x, str)})[:600] if isinstance(v, list) else []
    info = roster.by_uid.get(cid) or {}
    data = {"cid": cid, "name": info.get("name", ""), "own": clean(b.get("own")), "ownsk": clean(b.get("ownsk")),
            "at": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")}
    DEX_DIR.mkdir(exist_ok=True)
    store_set(f"dex_{cid}", dex_path(cid), data)
    return web.json_response({"ok": True})


# ---------- 관리 페이지(/관리: 관리자 전용) ----------
ADM_HOURS = int(os.environ.get("ADM_HOURS", "6"))


def adm_token(uid, exp, gid):
    return f"{uid}.{exp}.{gid}." + hmac.new(TIER_SECRET.encode(), f"adm:{uid}:{exp}:{gid}".encode(), hashlib.sha256).hexdigest()[:24]


def adm_parse(t):
    """관리 링크 검사 → (사용자 ID, 서버 ID) 또는 None"""
    p = (t or "").split(".")
    if len(p) != 4 or not (p[0].isdigit() and p[1].isdigit() and p[2].isdigit()) or int(p[1]) < time.time(): return None
    want = hmac.new(TIER_SECRET.encode(), f"adm:{p[0]}:{p[1]}:{p[2]}".encode(), hashlib.sha256).hexdigest()[:24]
    return (int(p[0]), int(p[2])) if hmac.compare_digest(p[3], want) else None


def adm_is_admin(uid, gid):
    """링크를 쓰는 동안에도 매번 관리자인지 다시 확인(권한을 빼면 바로 막힌다)."""
    if (gid, uid) in PERM or uid == OWNER_ID: return True
    g = bot.get_guild(gid); m = g.get_member(uid) if g else None
    return bool(m and m.guild_permissions.administrator)


def adm_state(gid):
    """관리 페이지에 보여줄 현재 DB 상태."""
    ally = store_get("alliance", ALLY_PATH, {}) or {}
    names, tdate, _ = talk_get()
    emp = _sb_req("GET", f"sam_emp?guild_id=eq.{gid}&select=cid,name,job,extra&order=name&limit=5000") or []
    mem = [{"c": e["cid"], "n": e.get("name") or "", "j": e.get("job") or "", "m": 1 if (e.get("extra") or {}).get("manual_at") else 0} for e in emp]
    perms = _sb_req("GET", f"sam_perm?guild_id=eq.{gid}&status=eq.approved&select=discord_id,user_name,decided_at,requested_at") or []
    adm = {str(p["discord_id"]): {"id": str(p["discord_id"]), "name": p.get("user_name") or "", "src": "grant"} for p in perms}
    g = bot.get_guild(gid)
    if g:
        for m in g.members:
            if not m.bot and m.guild_permissions.administrator:
                adm.setdefault(str(m.id), {"id": str(m.id), "name": str(m), "src": "discord"})
                if adm[str(m.id)]["src"] == "discord": adm[str(m.id)]["name"] = str(m)
    if OWNER_ID: adm.setdefault(str(OWNER_ID), {"id": str(OWNER_ID), "name": "", "src": "owner"})["src"] = "owner"
    db = db_load(); gg = bot.get_guild(gid); tr = []
    for d, n in (db.get("tries") or {}).items():
        if not n: continue
        m = gg.get_member(int(d)) if gg and str(d).isdigit() else None
        tr.append({"d": str(d), "name": str(m) if m else "", "cid": (db.get("tried") or {}).get(d, ""), "n": n, "lock": n >= MAX_TRIES})
    tr.sort(key=lambda x: (-x["lock"], -x["n"]))
    nm = {e["cid"]: e.get("name") or "" for e in emp}; grp = {}
    for r in _sb_req("GET", f"sam_emp_alt?guild_id=eq.{gid}&select=alt_cid,alt_name,main_cid&limit=100000") or []:
        grp.setdefault(r["main_cid"], []).append({"c": r["alt_cid"], "n": r.get("alt_name") or nm.get(r["alt_cid"], "")})
    alts = sorted(({"c": k, "n": nm.get(k, ""), "alts": sorted(v, key=lambda a: a["n"])} for k, v in grp.items()), key=lambda x: x["n"])
    return {"alts": alts, "titles": known_titles(), "staffCmds": CMD_HELP["간부"], "cmdperm": {k: v for k, v in (store_get("cmdperm", CMDPERM_PATH, {}) or {}).items() if v}, "tries": tr, "maxTries": MAX_TRIES, "cmds": CMD_HELP, "ally": ally, "talk": names, "talkDate": tdate, "members": mem, "admins": sorted(adm.values(), key=lambda a: (a["src"] != "owner", a["src"] != "discord", a["name"])), "now": int(time.time())}


async def adm_grant(gid, did, by):
    try: u = await bot.fetch_user(did)
    except Exception: raise ValueError("디스코드에서 찾을 수 없는 사용자 ID입니다.")
    if u.bot: raise ValueError("봇은 관리자로 지정할 수 없습니다.")
    g = bot.get_guild(gid)
    if g and not g.get_member(did): raise ValueError("이 서버에 참여하지 않은 사용자입니다.")
    def work():
        set_gid(gid); ally = store_get("alliance", ALLY_PATH, {}) or {}; now = datetime.datetime.now(datetime.timezone.utc).isoformat()
        sv = str(ally.get("server") or "")
        row = {"guild_id": gid, "discord_id": did, "guild_name": g.name if g else "", "user_name": str(u), "reason": "관리 페이지에서 부여",
               "ally_name": ally.get("name") or "", "game_server": sv if sv.isdigit() else None, "status": "approved",
               "requested_at": now, "decided_at": now, "decided_by": by}
        _sb_req("POST", "sam_perm?on_conflict=guild_id,discord_id", [row], "resolution=merge-duplicates,return=minimal")
    await asyncio.to_thread(work); PERM.add((gid, did)); LICENSED.add(gid); return str(u)


async def adm_revoke(gid, did, by):
    if did == OWNER_ID: raise ValueError("제작자의 권한은 뺄 수 없습니다.")
    if did == by: raise ValueError("본인의 권한은 뺄 수 없습니다. (잠금 방지)")
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    await asyncio.to_thread(_sb_req, "PATCH", f"sam_perm?guild_id=eq.{gid}&discord_id=eq.{did}", {"status": "revoked", "decided_at": now, "decided_by": by}, "return=minimal")
    PERM.discard((gid, did))
    g = bot.get_guild(gid); m = g.get_member(did) if g else None
    if m and m.guild_permissions.administrator: raise ValueError("부여 권한은 뺐지만, 이 사용자는 디스코드 서버 관리자 권한이 있어 계속 관리자입니다.")


async def web_admin(request):
    p = adm_parse(request.query.get("t"))
    if not p or not adm_is_admin(p[0], p[1]): return web.Response(status=403, text="링크가 만료되었거나 권한이 없습니다. 디스코드에서 /관리 를 다시 입력해 주세요.")
    tpl = (Path(__file__).parent / "admin_template.html").read_text("utf-8")
    return web.Response(text=tpl.replace("__T__", request.query.get("t")), content_type="text/html", headers={"Cache-Control": "no-store", "X-Robots-Tag": "noindex"})


async def web_admin_api(request):
    p = adm_parse(request.query.get("t"))
    if not p or not adm_is_admin(p[0], p[1]): return web.json_response({"error": "링크가 만료되었거나 권한이 없습니다. 디스코드에서 /관리 를 다시 입력해 주세요."}, status=403)
    uid, gid = p
    if request.content_length and request.content_length > 200_000: return web.json_response({"error": "요청이 너무 큽니다."}, status=413)
    try: body = await request.json()
    except Exception: return web.json_response({"error": "요청 형식이 올바르지 않습니다."}, status=400)
    if not isinstance(body, dict): return web.json_response({"error": "요청 형식이 올바르지 않습니다."}, status=400)
    op, msg = body.get("op"), ""
    try:
        if op == "talk":
            await asyncio.to_thread(lambda: (set_gid(gid), talk_apply(body.get("req") or {}))[1])
        elif op == "member_add":
            cid, nick, job = str(body.get("cid", "")).strip(), re.sub(r"\s+", " ", str(body.get("nick", ""))).strip(), str(body.get("job", "")).strip()[:20]
            if not (cid.isascii() and cid.isdigit() and 4 <= len(cid) <= 14): raise ValueError("CID는 숫자만 입력해 주세요.")
            if not nick or len(nick) > 20: raise ValueError("닉네임을 1~20자로 입력해 주세요.")
            msg = await asyncio.to_thread(lambda: (set_gid(gid), member_add(gid, cid, nick, job))[1])
        elif op == "admin_add":
            d = str(body.get("id", "")).strip()
            if not (d.isascii() and d.isdigit() and 15 <= len(d) <= 20): raise ValueError("디스코드 사용자 ID(숫자 17~19자리)를 입력해 주세요.")
            msg = f"✅ {await adm_grant(gid, int(d), uid)} 님에게 관리자 권한을 부여했습니다."
        elif op == "admin_remove":
            d = str(body.get("id", "")).strip()
            if not (d.isascii() and d.isdigit()): raise ValueError("사용자 ID가 올바르지 않습니다.")
            await adm_revoke(gid, int(d), uid); msg = "✅ 관리자 권한을 뺐습니다."
        elif op == "alt_add":
            key, raw = re.sub(r"\s+", " ", str(body.get("main", ""))).strip(), str(body.get("names", ""))[:2000]
            if not key or not raw.strip(): raise ValueError("본캐와 부캐 닉네임을 입력해 주세요.")
            def al():
                set_gid(gid); by = roster_for(gid).by_uid
                cs = [key] if key in by else [c for c, v in by.items() if v.get("name") == key]
                if not cs: raise ValueError(f"명단에 없는 본캐입니다: {key}")
                if len(cs) > 1: raise ValueError(f"같은 닉네임이 {len(cs)}명이라 CID로 입력해 주세요: {', '.join(cs)}")
                if _sb_req("GET", f"sam_emp_alt?guild_id=eq.{gid}&alt_cid=eq.{cs[0]}&select=alt_cid"): raise ValueError("이 본캐는 이미 다른 사람의 부캐로 연결되어 있습니다.")
                return cs[0], by[cs[0]].get("name") or cs[0], *alt_link(gid, cs[0], raw)
            mc, mn, ok, bad = await asyncio.to_thread(al)
            msg = (f"✅ {mn}에 부캐 {len(ok)}명 연결: {', '.join(n for n, _ in ok)}" if ok else "연결된 부캐가 없습니다.") + ("\n⚠️ 제외: " + " / ".join(f"{n}({w})" for n, w in bad) if bad else "")
        elif op == "alt_remove":
            c = str(body.get("cid", "")).strip()
            if not (c.isascii() and c.isdigit()): raise ValueError("CID가 올바르지 않습니다.")
            await asyncio.to_thread(_sb_req, "DELETE", f"sam_emp_alt?guild_id=eq.{gid}&alt_cid=eq.{c}", prefer="return=minimal"); msg = "✅ 부캐 연결을 해제했습니다."
        elif op == "cmdperm_set":
            cmd, ts = str(body.get("cmd", "")), body.get("titles")
            if cmd not in staff_cmds(): raise ValueError("설정할 수 없는 명령어입니다.")
            if not isinstance(ts, list): raise ValueError("직책 목록이 올바르지 않습니다.")
            def sv():
                set_gid(gid); ok = set(known_titles()); ts2 = [title_norm(t) for t in ts]
                if any(t not in ok for t in ts2): raise ValueError("알 수 없는 직책이 포함되어 있습니다.")
                d = dict(store_get("cmdperm", CMDPERM_PATH, {}) or {}); d[cmd] = sorted(set(ts2)); store_set("cmdperm", CMDPERM_PATH, d)
            await asyncio.to_thread(sv)
            msg = f"✅ {cmd} 권한을 " + (", ".join(sorted(set(title_norm(t) for t in ts))) + " 직책으로 제한했습니다." if ts else "간부 전체로 되돌렸습니다.")
        elif op == "tries_reset":
            k = str(body.get("id", "")).strip()
            if not (k.isascii() and k.isdigit()): raise ValueError("ID가 올바르지 않습니다.")
            def rs():
                set_gid(gid); db = db_load(); hit = [d for d in list(db.get("tries", {})) if d == k or db.get("tried", {}).get(d) == k]
                for d in hit: db["tries"].pop(d, None); db.get("tried", {}).pop(d, None)
                db_save(db); return len(hit)
            n = await asyncio.to_thread(rs)
            if not n: raise ValueError("해당 ID(디스코드 ID 또는 CID)의 실패 기록이 없습니다.")
            msg = f"✅ 인증 실패 횟수를 초기화했습니다 ({n}건)."
        elif op != "state": raise ValueError("알 수 없는 요청입니다.")
        st = await asyncio.to_thread(lambda: (set_gid(gid), adm_state(gid))[1])
    except ValueError as e: return web.json_response({"error": str(e)}, status=400)
    except Exception as e:
        log.exception("관리 페이지 오류"); return web.json_response({"error": "처리 중 오류가 발생했습니다: " + str(e)[:120]}, status=500)
    st["msg"] = msg; st["me"] = str(uid); return web.json_response(st)


async def start_web():
    if not WEB_PORT: return
    app = web.Application(); app.router.add_get("/", web_index)
    app.router.add_get("/tier", web_tier); app.router.add_get("/hr", web_hr); app.router.add_post("/api/dex", web_dex_save); app.router.add_post("/api/talk", web_talk)
    app.router.add_get("/admin", web_admin); app.router.add_post("/api/admin", web_admin_api)
    app.router.add_get("/health", lambda r: web.Response(text="ok"))      # 외부 핑(UptimeRobot 등)용
    runner = web.AppRunner(app); await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", WEB_PORT).start()
    log.info("상황판: http://localhost:%d/%s", WEB_PORT, f"?key={WEB_KEY}" if WEB_KEY else "")


@bot.tree.command(name="동맹등록", description="서버 번호(숫자)와 동맹명을 등록합니다. 승인 전 서버는 제작자에게 신청됩니다")
async def cmd_ally(it: discord.Interaction):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    if licensed(it.guild_id):                              # 이미 등록된 서버는 한 번만 등록 가능 → 수정은 /동맹수정
        return await it.response.send_message("이미 등록된 동맹입니다. 정보 변경은 간부가 `/동맹수정`으로 해 주세요.", ephemeral=True)
    await it.response.send_modal(AllyApplyModal())         # 승인 전 서버: 제작자에게 신청


@bot.tree.command(name="동맹수정", description="[간부] 등록된 서버 번호·동맹명을 수정합니다")
async def cmd_ally_edit(it: discord.Interaction):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    if not licensed(it.guild_id): return await it.response.send_message("먼저 `/동맹등록`으로 동맹을 등록해 주세요.", ephemeral=True)
    if not staff_ok(it, "/동맹수정"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    await it.response.send_modal(AllyModal())


# ---------- 시트 → Supabase 이관 ----------
def _sb_rows(table, rows, conflict="guild_id,cid"):
    for i in range(0, len(rows), 200):
        _sb_req("POST", f"{table}?on_conflict={conflict}", rows[i:i + 200], prefer="resolution=merge-duplicates,return=minimal")


def _sb_count(table, gid=GUILD_ID):
    return len(_sb_req("GET", f"{table}?guild_id=eq.{gid}&select=guild_id&limit=100000") or [])


def _num(v):
    try: return int(float(str(v).replace(",", "").strip() or 0))
    except ValueError: return 0


MANUAL_KEEP_DAYS = 3


def hr_import(now=None, rows=None, gid=GUILD_ID):
    """인사 탭을 읽어 중복(같은 CID)을 하나로 묶고, 사람 정보와 주간 수치를 나눠 저장한다.
    sam_emp: CID·멤버·직업·조별 / sam_emp_stat: 직위·번영·주간 무훈·주간 공헌·주둔지·주 공성 횟수(날짜별)."""
    now = now or datetime.datetime.now(datetime.timezone.utc).isoformat()
    snap = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=9)).strftime("%Y-%m-%d")   # KST 날짜
    if rows is None: raise RuntimeError("명단은 엑셀 업로드(/명단업데이트)로만 반영합니다.")
    _sb_ensure_tenant(gid)
    uidc = COL_UID + ["캐릭터 ID", "캐릭터ID", "CID"]
    hi = next((i for i, r in enumerate(rows) if roster._find(r, uidc) is not None), None)
    if hi is None: raise RuntimeError("인사 탭에서 CID 열을 찾지 못했습니다.")
    H = [h.strip() for h in rows[hi]]; ic = roster._find(H, uidc); merged, dup, warn = {}, 0, []
    for r in rows[hi + 1:]:
        cid = Roster.norm_uid(r[ic]) if ic < len(r) else ""
        if not cid: continue
        rec = {h: (r[i].strip() if i < len(r) else "") for i, h in enumerate(H) if i != ic and h}
        if cid in merged:                                  # 같은 CID가 여러 줄이면 하나로 합침(뒤 줄의 값이 우선, 빈 칸은 앞 값 유지)
            dup += 1; merged[cid].update({k: v for k, v in rec.items() if v != ""})
        else: merged[cid] = rec
    names = {}
    for cid, rec in merged.items():
        if rec.get("멤버"): names.setdefault(rec["멤버"], []).append(cid)
    same = {n: c for n, c in names.items() if len(c) > 1}
    if same: warn.append("이름은 같고 CID가 다른 멤버: " + ", ".join(f"{n}({'/'.join(c)})" for n, c in list(same.items())[:5]))
    ro, st = [], []
    for cid, g in merged.items():
        ro.append({"guild_id": gid, "cid": cid, "name": g.get("멤버", ""), "job": g.get("직업", ""), "updated_at": now,
                   "extra": {"조별": g.get("조별", "")}})
        m = re.match(r"\s*([^(]*)\(\s*(\d+)\s*,\s*(\d+)\s*\)", g.get("주둔지", ""))
        st.append({"guild_id": gid, "cid": cid, "snap_date": snap, "title": g.get("직위", ""), "prosperity": _num(g.get("번영")),
                   "war_merit": _num(g.get("주간 무훈")), "contribution": _num(g.get("주간 공헌")),
                   "camp": (m.group(1).strip() if m else g.get("주둔지", "")), "camp_x": int(m.group(2)) if m else None,
                   "camp_y": int(m.group(3)) if m else None, "siege_count": _num(g.get("주 공성 횟수"))})
    # 이번 명단에서 빠진 사람 = 접었거나 동맹을 나간 사람 → 이탈자 표(sam_emp_left)로 옮기고 sam_emp에서 제외
    old = _sb_req("GET", f"sam_emp?guild_id=eq.{gid}&select=cid,name,job,extra&limit=100000") or []
    def _manual_fresh(r):                                          # '동맹원추가'로 넣은 사람은 게임 엑셀(매일 오전 8시 일배치)에 반영될 때까지 며칠 유지
        t = (r.get("extra") or {}).get("manual_at")
        try: return bool(t) and time.time() - datetime.datetime.fromisoformat(t).timestamp() < MANUAL_KEEP_DAYS * 86400
        except Exception: return False
    kept = [r for r in old if r["cid"] not in merged and _manual_fresh(r)]
    gone = [r for r in old if r["cid"] not in merged and not _manual_fresh(r)]
    moved = 0
    if gone and len(old) >= 10 and len(gone) > len(old) * 0.5:      # 파일이 잘못 올라온 경우 전원이 이탈 처리되는 사고 방지
        warn.append(f"이번 명단에서 빠진 사람이 {len(gone)}/{len(old)}명으로 너무 많아 이탈 처리를 건너뛰었습니다. 명단 파일을 확인해 주세요.")
        gone = []
    _sb_rows("sam_emp", ro); _sb_rows("sam_emp_stat", st, conflict="guild_id,cid,snap_date")
    if gone:
        last = {}                                                  # 이탈 직전 마지막 인사 수치를 함께 보관(sam_emp 삭제 시 sam_emp_stat은 같이 지워지므로)
        for i in range(0, len(gone), 100):
            ids = ",".join(r["cid"] for r in gone[i:i + 100])
            for x in _sb_req("GET", f"sam_emp_stat?guild_id=eq.{gid}&cid=in.({ids})&order=snap_date.desc&select=cid,snap_date,title,prosperity,war_merit,contribution,camp,siege_count&limit=100000") or []:
                last.setdefault(x["cid"], x)
        al = (_sb_req("GET", f"sam_alli?guild_id=eq.{gid}&select=name,game_server") or [{}])[0]      # 서버 번호 + 동맹명으로 sam_alli에 FK
        fk = {"game_server": al["game_server"], "ally_name": al["name"]} if al.get("game_server") and al.get("name") else {}
        lf = [{"guild_id": gid, "cid": r["cid"], "name": r.get("name") or "", "job": r.get("job") or "", "left_date": snap, "left_at": now,
               "extra": r.get("extra") or {}, "last_stat": last.get(r["cid"]), **{"game_server": None, "ally_name": None, **fk}} for r in gone]
        _sb_rows("sam_emp_left", lf)
        for i in range(0, len(gone), 100):
            _sb_req("DELETE", f"sam_emp?guild_id=eq.{gid}&cid=in.({','.join(r['cid'] for r in gone[i:i + 100])})", prefer="return=minimal")
        moved = len(gone)
    back = [r["cid"] for r in (_sb_req("GET", f"sam_emp_left?guild_id=eq.{gid}&select=cid&limit=100000") or []) if r["cid"] in merged]
    for i in range(0, len(back), 100):                              # 다시 명단에 나타난 사람은 이탈자 표에서 뺀다
        _sb_req("DELETE", f"sam_emp_left?guild_id=eq.{gid}&cid=in.({','.join(back[i:i + 100])})", prefer="return=minimal")
    out = [f"명단(sam_emp): 시트 {len(rows) - hi - 1}줄 → 중복 {dup}건 병합 → {len(ro)}명 (DB 총 {_sb_count('sam_emp', gid)}명)",
           f"인사 수치(sam_emp_stat): {snap} 기준 {len(st)}건 (DB 총 {_sb_count('sam_emp_stat', gid)}건)"]
    if moved: out.append(f"이탈자(sam_emp_left): {moved}명을 이탈 처리했습니다 — " + ", ".join(f"{r.get('name') or '-'}({r['cid']})" for r in gone[:8]) + (" 외" if moved > 8 else ""))
    if kept: out.append(f"수동 추가 {len(kept)}명은 아직 엑셀에 없어 유지했습니다: " + ", ".join(r.get('name') or r['cid'] for r in kept[:8]))
    if back: out.append(f"복귀: 이탈자 표에 있던 {len(back)}명이 다시 명단에 있어 이탈 해제했습니다.")
    return out, warn




async def update_roster_tenant(it, 파일):
    gid = it.guild_id
    if 파일 is None or not 파일.filename.lower().endswith(".xlsx"):
        return await it.followup.send("❌ 인사 명단 엑셀(.xlsx)을 첨부해 주세요. 첫 시트에 `캐릭터 ID`, `멤버`, `직업`, `직위` 열이 필요합니다. (조별, 번영, 주간 무훈, 주간 공헌, 주둔지, 주 공성 횟수는 선택)", ephemeral=True)
    tmp = Path(tempfile.gettempdir()) / f"roster_{gid}.xlsx"
    try:
        await 파일.save(tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        rows = [["" if c is None else (str(int(c)) if isinstance(c, float) and c.is_integer() else str(c)) for c in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
        out, warn = await asyncio.to_thread(hr_import, None, rows, gid)
        await asyncio.to_thread(roster_load_db, gid)
    except Exception as e:
        log.exception("명단 업로드 실패"); return await it.followup.send(f"❌ 실패: {str(e)[:300]}", ephemeral=True)
    finally: tmp.unlink(missing_ok=True)
    sy = await sync_members(it.guild)
    await it.followup.send("✅ 명단을 반영했습니다.\n" + "\n".join("• " + x for x in out)
                           + f"\n• 역할 동기화: 인증자 {sy['total']}명 중 역할 변경 {sy['role']}명, 별명 변경 {sy['nick']}명"
                           + ("\n⚠️ " + "\n⚠️ ".join(warn) if warn else ""), ephemeral=True)


class MemberAddModal(CtxModal, title="동맹원 추가"):
    cid = discord.ui.TextInput(label="CID (캐릭터 ID)", placeholder="게임 내 캐릭터 ID (숫자)", max_length=30)
    nick = discord.ui.TextInput(label="닉네임", placeholder="게임 내 닉네임", max_length=25)
    job = discord.ui.TextInput(label="직업", placeholder="예) 천공", max_length=10, required=False)

    async def on_submit(self, it: discord.Interaction):
        await it.response.defer(ephemeral=True)
        if not staff_ok(it, "/동맹수정"): return await it.followup.send("권한이 없습니다.", ephemeral=True)
        gid = cur_gid(); cid = Roster.norm_uid(self.cid.value); nick = re.sub(r"\s+", " ", self.nick.value).strip(); job = self.job.value.strip()
        if not cid.isdigit(): return await it.followup.send("❌ CID는 숫자로 입력해 주세요.", ephemeral=True)
        if not nick: return await it.followup.send("❌ 닉네임을 입력해 주세요.", ephemeral=True)
        try: r = await asyncio.to_thread(member_add, gid, cid, nick, job)
        except Exception as e:
            log.exception("동맹원 추가 실패"); return await it.followup.send(f"❌ 실패: {str(e)[:200]}", ephemeral=True)
        await it.followup.send(r, ephemeral=True)


def member_add(gid, cid, nick, job):
    """게임 엑셀(일배치)에 아직 없는 사람을 명단에 직접 넣는다. 엑셀에 나타나기 전까지 MANUAL_KEEP_DAYS일 동안 이탈 처리하지 않는다."""
    _sb_ensure_tenant(gid)
    ex = _sb_req("GET", f"sam_emp?guild_id=eq.{gid}&cid=eq.{cid}&select=cid,name") or []
    if ex: return f"이미 명단에 있습니다: **{ex[0].get('name') or '-'}** ({cid})"
    dupn = _sb_req("GET", f"sam_emp?guild_id=eq.{gid}&name=eq.{urllib.parse.quote(nick, safe='')}&select=cid") or []
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    _sb_rows("sam_emp", [{"guild_id": gid, "cid": cid, "name": nick, "job": job, "updated_at": now, "extra": {"조별": "", "manual_at": now}}])
    _sb_req("DELETE", f"sam_emp_left?guild_id=eq.{gid}&cid=eq.{cid}", prefer="return=minimal")      # 이탈자였다면 복귀
    roster_load_db(gid)
    return (f"✅ 동맹원을 추가했습니다: **{nick}** ({cid})" + (f" · {job}" if job else "")
            + f"\n게임 엑셀에 반영되기 전까지 {MANUAL_KEEP_DAYS}일간 명단에 유지됩니다. 이제 인증하기에서 이 CID로 인증할 수 있습니다."
            + (f"\n⚠️ 같은 닉네임의 다른 CID가 이미 있습니다: {', '.join(r['cid'] for r in dupn)}" if dupn else ""))


@bot.tree.command(name="동맹원추가", description="[간부] 게임 엑셀(매일 오전 8시 반영)에 아직 없는 동맹원을 명단에 직접 추가합니다")
async def cmd_member_add(it: discord.Interaction):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    if not staff_ok(it, "/동맹원추가"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    if not SB_ON: return await it.response.send_message("DB가 연결되지 않아 사용할 수 없습니다.", ephemeral=True)
    await it.response.send_modal(MemberAddModal())


@bot.tree.command(name="db이관", description="[제작자] 구글 시트 인사 탭의 명단을 DB에 한 번 이관합니다(비상용)")
async def cmd_db_migrate(it: discord.Interaction):
    owner = await perm_owner()
    if it.user.id != owner.id: return await it.response.send_message("제작자만 사용할 수 있습니다.", ephemeral=True)
    if not SB_ON: return await it.response.send_message("DB가 연결되지 않았습니다.", ephemeral=True)
    if not (SHEET_ID and ROSTER_GID): return await it.response.send_message("SHEET_ID / ROSTER_GID가 설정되지 않았습니다.", ephemeral=True)
    await it.response.defer(ephemeral=True)
    try:
        rows = await asyncio.to_thread(lambda: list(csv.reader(io.StringIO(fetch_sheet_csv()))))
        out, warn = await asyncio.to_thread(hr_import, None, rows, GUILD_ID)
        n = await asyncio.to_thread(roster_load_db, GUILD_ID)
    except Exception as e:
        log.exception("DB 이관 실패"); return await it.followup.send(f"❌ 이관 실패: {str(e)[:300]}", ephemeral=True)
    await it.followup.send("✅ 시트 인사 탭을 DB로 이관했습니다.\n" + "\n".join("• " + x for x in out) + f"\n• 봇 명단 {n}명으로 갱신" + ("\n⚠️ " + "\n⚠️ ".join(warn) if warn else ""), ephemeral=True)


def alt_link(gid, main_cid, raw, replace=False):
    """부계정 닉네임 목록을 명단과 대조해 본캐에 연결한다. → (연결된 [(닉,CID)], 제외 [(닉, 사유)])"""
    import difflib
    names = []
    for x in re.split(r"[,\n，、;/]+", raw or ""):
        n = re.sub(r"\s+", " ", x).strip()
        if n and n not in names: names.append(n)
    by_name = {}
    for cid, v in roster_for(gid).by_uid.items():
        if v.get("name"): by_name.setdefault(v["name"], []).append(cid)
    taken = {r["alt_cid"]: r["main_cid"] for r in (_sb_req("GET", f"sam_emp_alt?guild_id=eq.{gid}&select=alt_cid,main_cid&limit=100000") or [])}
    ok, bad = [], []
    for n in names:
        cids = by_name.get(n)
        if not cids:
            sug = difflib.get_close_matches(n, list(by_name), n=2, cutoff=0.6)
            bad.append((n, "명단에 없는 닉네임" + (f" (혹시 {' / '.join(sug)} ?)" if sug else ""))); continue
        if len(cids) > 1: bad.append((n, f"같은 닉네임이 {len(cids)}명이라 구분할 수 없음 (CID {', '.join(cids)})")); continue
        c = cids[0]
        if c == main_cid: bad.append((n, "본캐 자신")); continue
        if taken.get(c) not in (None, main_cid): bad.append((n, "이미 다른 본캐에 연결된 계정")); continue
        ok.append((n, c))
    if replace: _sb_req("DELETE", f"sam_emp_alt?guild_id=eq.{gid}&main_cid=eq.{main_cid}", prefer="return=minimal")
    if ok: _sb_rows("sam_emp_alt", [{"guild_id": gid, "alt_cid": c, "alt_name": n, "main_cid": main_cid} for n, c in ok], conflict="guild_id,alt_cid")
    return ok, bad


@bot.tree.command(name="부캐", description="부계정을 닉네임으로 연결합니다. 예) 닉네임1,닉네임2 (비우면 목록). 관리자는 본캐를 지정해 대신 연결")
@app_commands.describe(닉네임="부계정 닉네임을 쉼표로 구분해 입력", 본캐="[관리자] 본캐 닉네임(비우면 나의 본캐)", 방식="추가(기본): 기존에 더함 / 교체: 입력한 목록으로 바꿈")
@app_commands.choices(방식=[app_commands.Choice(name="추가", value="add"), app_commands.Choice(name="교체", value="replace")])
async def cmd_alt(it: discord.Interaction, 닉네임: str = "", 본캐: str = "", 방식: app_commands.Choice[str] = None):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    if not SB_ON: return await it.response.send_message("DB가 연결되지 않아 사용할 수 없습니다.", ephemeral=True)
    gid = cur_gid(); did = str(it.user.id); target = ""
    if 본캐.strip():                                                         # 관리자: 다른 사람의 본캐를 지정
        if not is_admin(it): return await it.response.send_message("본캐를 지정하는 기능은 관리자만 사용할 수 있습니다.", ephemeral=True)
        nm = re.sub(r"\s+", " ", 본캐).strip()
        cs = [c for c, v in roster_for(gid).by_uid.items() if v.get("name") == nm]
        if not cs: return await it.response.send_message(f"❌ 명단에 없는 본캐 닉네임입니다: {nm}", ephemeral=True)
        if len(cs) > 1: return await it.response.send_message(f"❌ 같은 닉네임이 {len(cs)}명이라 구분할 수 없습니다 (CID {', '.join(cs)}).", ephemeral=True)
        main, target = cs[0], nm
    else:
        main = next((c for c, d in db_load()["uid"].items() if str(d) == did), None)
        if not main: return await it.response.send_message("먼저 `/인증`으로 본캐 CID 인증을 해 주세요.", ephemeral=True)
    await it.response.defer(ephemeral=True)
    who = f"**{target}**" if target else "내"
    try:
        if not 닉네임.strip():
            rows = await asyncio.to_thread(lambda: _sb_req("GET", f"sam_emp_alt?guild_id=eq.{gid}&main_cid=eq.{main}&select=alt_name,alt_cid&order=created_at") or [])
            return await it.followup.send(f"{who} 부계정\n" + ("\n".join(f"• {r['alt_name']} ({r['alt_cid']})" for r in rows) if rows else "연결된 부계정이 없습니다. `/부캐 닉네임1,닉네임2` 로 연결하세요."), ephemeral=True)
        ok, bad = await asyncio.to_thread(alt_link, gid, main, 닉네임, bool(방식 and 방식.value == "replace"))
    except Exception as e:
        log.exception("부캐 연결 실패"); return await it.followup.send(f"❌ 실패: {str(e)[:200]}", ephemeral=True)
    lines = [f"**{target or '내'} 부계정 연결 결과** — 연결 {len(ok)}명 · 제외 {len(bad)}명"]
    if ok: lines += ["", "✅ 연결됨"] + [f"• {n} ({c})" for n, c in ok]
    if bad: lines += ["", "❌ 제외됨 (닉네임을 확인해 다시 입력해 주세요)"] + [f"• {n} — {why}" for n, why in bad]
    await it.followup.send("\n".join(lines)[:1900], ephemeral=True)


@bot.tree.command(name="명단업데이트", description="[관리자] 게임 명단 엑셀(.xlsx)을 올려 명단을 새로 반영합니다")
@admin_only
@app_commands.describe(파일="게임에서 받은 명단 엑셀(.xlsx)")
async def update_roster(it: discord.Interaction, 파일: discord.Attachment = None):
    if not is_admin(it):
        return await it.response.send_message("관리자만 사용할 수 있습니다.", ephemeral=True)
    await it.response.defer(ephemeral=True)
    if SB_ON: return await update_roster_tenant(it, 파일)      # 홈 서버도 다른 서버와 똑같이: 엑셀을 올려 DB에 저장(시트로는 제어하지 않음)
    try: roster.load()
    except Exception: pass
    before = set(roster.by_uid)
    try:
        if 파일 is not None:
            if not 파일.filename.lower().endswith(".xlsx"): return await it.followup.send("❌ .xlsx 파일만 올릴 수 있습니다.", ephemeral=True)
            tmp = Path(EXCEL_PATH).with_suffix(".new.xlsx"); await 파일.save(tmp)
            try:
                wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
                ws = wb[SHEET_NAME] if SHEET_NAME else wb.worksheets[0]
                Roster._ingest(roster, ws.iter_rows(values_only=True))   # 검증만
            except Exception as e:
                tmp.unlink(missing_ok=True); return await it.followup.send(f"❌ 파일을 읽지 못했습니다: {e}", ephemeral=True)
            if Path(EXCEL_PATH).exists(): shutil.copy(EXCEL_PATH, Path(EXCEL_PATH).with_suffix(".bak.xlsx"))
            shutil.move(tmp, EXCEL_PATH); roster.mtime = 0; roster.load(); src = f"엑셀 {파일.filename}"
        else:
            roster.load_csv_text(fetch_sheet_csv()); src = "구글 시트"
    except Exception as e:
        log.exception("명단 업데이트 실패"); return await it.followup.send(f"❌ 실패: {e}", ephemeral=True)
    now = set(roster.by_uid)
    tk = refresh_talk()
    log.info("명단 업데이트 by %s (%s)", it.user, src)
    extra = (f"\n상황판: 멤버 {len(roster.members)}명" + (f", 톡방 {tk}명" if tk is not None else "") + " 반영") if roster.members else "\n상황판: 인사 탭의 열이 달라 반영하지 못했습니다(봇 창 로그 확인)."
    out = write_html()
    if out: extra += f"\n인사데이터 HTML 파일 갱신: {out.name}"
    sy = await sync_members(it.guild)
    extra += f"\n역할 동기화: 인증자 {sy['total']}명 중 역할 변경 {sy['role']}명, 별명 변경 {sy['nick']}명, 회수 {sy['removed']}명" + (f", 실패 {sy['fail']}명" if sy['fail'] else "")
    await it.followup.send(f"✅ {src}로 명단을 반영했습니다. 전체 {len(now)}명 · 신규 {len(now - before)}명 · 제외 {len(before - now)}명\n직업: {', '.join(roster.jobs) or '-'}{extra}", ephemeral=True)


@bot.tree.command(name="명단새로고침", description="[관리자] 엑셀 명단을 다시 읽습니다")
@admin_only
async def reload(it: discord.Interaction):
    if not is_admin(it): return await it.response.send_message("관리자만 사용할 수 있습니다.", ephemeral=True)
    if cur_gid() != GUILD_ID:
        try: n = await asyncio.to_thread(roster_load_db, cur_gid()); return await it.response.send_message(f"명단 {n}명 로드", ephemeral=True)
        except Exception as e: return await it.response.send_message(f"실패: {e}", ephemeral=True)
    roster.mtime = 0
    try: roster.load(); await it.response.send_message(f"명단 {len(roster.by_uid)}명 로드, 직업: {', '.join(roster.jobs) or '-'}", ephemeral=True)
    except Exception as e: await it.response.send_message(f"실패: {e}", ephemeral=True)


# ---------- 추천덱 ----------
_TIER = {}


def tier_data():
    """tier_template.html 안의 시즌1·2 티어덱(S1·S2)과 장수·전법 데이터(SD)를 읽는다(처음 한 번)."""
    if not _TIER:
        t = (Path(__file__).parent / "tier_template.html").read_text("utf-8"); dec = json.JSONDecoder()
        for k in ("S1", "S2", "SD"):
            i = t.index(f"const {k}=") + len(f"const {k}="); _TIER[k] = dec.raw_decode(t[i:])[0]
    return _TIER


def _nk(x): return re.sub(r"[\s·\-]", "", x or "")


STS = ["혼란", "허약", "침묵", "무장 해제", "공포", "위협", "군량 고갈", "화공", "홍수", "폭풍", "연소", "짐독", "요술", "조롱"]
_ST_CACHE = {}


def status_of(desc):
    """장수·전법 설명에서 이상상태를 (부여하는 것 g, 조건으로 쓰는 것 u)로 나눈다. (tier_template.html의 stOf와 같은 규칙)"""
    desc = re.sub(r"\s+", " ", (desc or "").replace("무장해제", "무장 해제"))
    if desc in _ST_CACHE: return _ST_CACHE[desc]
    g, u = set(), set()
    for S in STS:
        if S not in desc: continue
        r = S.replace(" ", " ?")
        if re.search(r + r"(?:을\(를\)|를|을|와\(과\)|과|와)?(?: 상태를?| 효과를?)?(?: ?\d+스택)?(?: 및 [가-힣 ]{1,8}?(?:을\(를\)|를|을| 상태를| 효과를)?)*\s*부여", desc) or re.search(r + " 상태가 된다", desc): g.add(S)
        if re.search(r"(?:이미 )?" + r + r" ?(?:상태|효과)?(?:면|를 보유|가 아닌|인 |일 |를 이미|을 보유)", desc): u.add(S)
    if re.search(r"디버프 상태 \d가지를 제거", desc): g.add("정화")
    if re.search(r"디버프 상태를 보유|디버프 상태 1개당|디버프 효과 부여 후", desc): u.add("디버프")
    _ST_CACHE[desc] = (g, u); return g, u


def unit_state(deck, heroes, hd):
    hs = [h["n"] for h in heroes if h["n"]]
    G, U = set(), set()
    for h in hs:
        g, u = status_of((hd.get(h) or {}).get("desc")); G |= g; U |= u
    names = set()
    if deck:
        for r in deck["rows"]:
            if r[0] in ("전법", "대체 전법"):
                for hero_slots in r[1]:
                    for cell in hero_slots:
                        for x in re.split(r"[/>+,]", cell): names.add(_nk(x))
    st = {"hs": hs, "G": G, "U": U, "deck": names, "picks": [], "healed": any("치유" in (hd.get(h) or {}).get("sktrait", "") or "회복" in ((hd.get(h) or {}).get("desc") or "") for h in hs),
          "styles": " ".join((hd.get(h) or {}).get("style", "") for h in hs)}
    st["trait"] = unit_trait(st, hd); return st


_TI = {}
DOT_ST = {"화공", "연소", "짐독"}
ORDER = {"지휘": 1, "패시브": 2, "액티브": 3, "추격": 5}          # 발동 순서: 지휘 → 패시브 → 액티브 → (4: 평타) → 추격


def tact_info(name, skd):
    """전법 분류: 종류·계열, 부여/활용 이상상태, 딜 방식(도트딜=누적 / 한방딜=고배율 / 일반딜), 회복류 여부."""
    t = skd.get(name) or {}
    desc = re.sub(r"\s+", " ", t.get("desc") or "")
    key = (name, desc)
    if key in _TI: return _TI[key]
    g, u = status_of(desc)
    pc = [float(x) for x in re.findall(r"→ ?(\d+(?:\.\d+)?)%의? ?(?:책략|병기|책략과 병기|병기와 책략)?(?:\s?피해)", desc)]
    peak, total = (max(pc), sum(pc)) if pc else (0, 0)
    rep = bool(re.search(r"매 턴|턴 시작 시|턴 종료 시|총 \d회|\d회 발동|행동 시", desc)) and bool(pc)
    dot = bool(g & DOT_ST) or rep
    burst = (peak >= 250 or total >= 300) and not dot
    heal = t.get("trait") == "치유" or bool(re.search(r"병력을? 회복|치유율|회유이?\(?가?\)? ?\d|정신 회복|심리", desc))
    kind = "도트딜" if dot else "한방딜" if burst else "일반딜" if pc else "회복" if heal else "지원"
    info = {"type": t.get("type", ""), "trait": t.get("trait", ""), "g": g - {"정화"}, "u": u - {"디버프"}, "peak": peak, "kind": kind, "heal": heal,
            "flood": "홍수" in g, "threat": "위협" in g, "uses_flood": "홍수" in u or "홍수" in desc and t.get("trait") == "책략"}
    _TI[key] = info; return info


def unit_trait(st, hd):
    """부대 주력 계열: 장수 기술 계열(책략/병기) 다수결."""
    c = {"책략": 0, "병기": 0}
    for h in st["hs"]:
        tr = (hd.get(h) or {}).get("sktrait", "")
        for k in c:
            if k in tr: c[k] += 1
        for k in c:
            if k in (hd.get(h) or {}).get("style", ""): c[k] += 0.5
    return "책략" if c["책략"] > c["병기"] else "병기" if c["병기"] > c["책략"] else ""


def pick_tactics(st, count, own_s, used, hd, skd, skn, deal="도트딜"):
    """한 부대에 전법을 count개 더 고른다.
    ① 딜 방식: 선호(기본 도트딜)와 같으면 +3, 반대 방식(한방딜)이면 -1.5  ② 책략 부대는 홍수 부여(+5)·홍수 활용(+3),
    병기 부대는 위협 부여(+6)  ③ 장수 기술이 필요로 하는 이상상태를 부여 +6 / 부대 이상상태 활용 +4  ④ 회복류가 없으면 회복류 +3 (회유·정신 회복 포함)
    ⑤ 티어덱 +3, 계열 적합 +2  ⑥ 지휘·패시브로 먼저 깔리는 이상상태 부여는 +1(발동 순서가 빠름)."""
    G, U = st["G"], st["U"]
    trait = st.setdefault("trait", "")
    for _ in range(count):
        best = None
        for k in sorted(own_s - used):
            n = skn.get(k, k); ti = tact_info(n, skd); g, u = ti["g"] | (status_of((skd.get(n) or {}).get("desc"))[0] & {"정화"}), ti["u"]
            why = []; sc = 0.0
            need = {x for x in U if x != "디버프"} - G
            for x in sorted((g - {"정화"}) & need): sc += 6; why.append(f"{x} 부여")
            for x in sorted(u & G): sc += 4; why.append(f"{x} 활용")
            if trait == "책략":
                if ti["flood"] and "홍수" not in G: sc += 5; why.append("홍수 부여")
                elif ti["trait"] == "책략" and "홍수" in G and ("홍수" in u or ti["uses_flood"]): sc += 3; why.append("홍수 활용")
            if trait == "병기" and ti["threat"] and "위협" not in G: sc += 6; why.append("위협 부여")
            if (u - G - g) - {"디버프"}: sc -= 3
            if ti["kind"] == deal: sc += 3; why.append(deal)
            elif ti["kind"] in ("한방딜", "도트딜"): sc -= 1.5
            if (ti["heal"] or ti["kind"] == "회복") and not st["healed"]: sc += 3; why.append("회복")
            if ti["trait"] == trait and trait: sc += 2
            if k in st["deck"]: sc += 3; why.append("티어덱")
            if ti["type"] in ("지휘", "패시브") and g: sc += 1
            sc += {"전설": 0.6, "영웅": 0.3}.get((skd.get(n) or {}).get("grade"), 0)
            if best is None or sc > best[0]: best = (sc, k, n, why, g, u, ti)
        if best is None: return
        _, k, n, why, g, u, ti = best; used.add(k); G |= g; U |= u
        if ti["heal"] or ti["kind"] == "회복": st["healed"] = True
        st["picks"].append((n, "·".join(dict.fromkeys(why))[:30] if False else "·".join(list(dict.fromkeys(why))[:2]) or "보유 전법"))


def finish_tactics(st, hd, skd):
    """고른 전법을 장수당 2개까지 나눠 준다(전법이 적은 장수 우선, 같으면 역할이 맞는 장수). → ({장수: [(전법, 이유)]}, (부여되는 이상상태, 부족한 이상상태))"""
    out = {h: [] for h in st["hs"]}
    for n, why in st["picks"]:
        tr = (skd.get(n) or {}).get("trait", "")
        order = sorted((h for h in st["hs"] if len(out[h]) < 2), key=lambda h: (len(out[h]), tr not in (hd.get(h) or {}).get("style", "")))
        if order: out[order[0]].append((n, why))
    for h in out: out[h].sort(key=lambda x: ORDER.get(tact_info(x[0], skd)["type"], 3))
    return out, (sorted(st["G"] - {"정화"}), sorted(x for x in st["U"] - st["G"] if x != "디버프"))


def recommend_decks(own_h, own_s, season, hd, skd, units=5, deal="도트딜"):
    """티어덱 순서대로 5부대 편성. 한 번 쓴 장수·전법은 다시 쓰지 않는다.
    덱을 통째로 못 만들면 2명 이상, 그다음 1명 이상 보유한 덱을 고르고 빈자리는 같은 역할의 미사용 보유 장수로 채운다."""
    decks = tier_data()["S1" if season == 1 else "S2"]
    own_h = {h for h in own_h if (hd.get(h) or {}).get("season", 1) <= season}
    skn = {_nk(k): k for k in skd}
    own_s = {_nk(x) for x in own_s if (skd.get(x) or {}).get("season", 1) <= season}
    freq = {}
    for d in decks:
        for h in d["heroes"]:
            for a in h.split("/"): freq[a] = freq.get(a, 0) + 1
    usedH, usedS, usedD, plan = set(), set(), set(), []

    def pick(minc):
        for i, d in enumerate(decks):
            if i in usedD: continue
            got = []
            for slot in d["heroes"]:
                alts = [a for a in slot.split("/") if a in own_h and a not in usedH and a not in got]
                got.append(alts[0] if alts else None)
            if sum(1 for g in got if g) >= minc: return i, got
        return None

    def sim(a, b):
        A, B = hd.get(a), hd.get(b)
        if not A or not B: return 0
        return 2 * len(set(A["style"].split()) & set(B["style"].split())) + (A["pos"] == B["pos"]) + (A["troop"] == B["troop"]) + (A["sktype"] == B["sktype"])
    for _ in range(units):
        sel = next((r for m in (2, 1) for r in [pick(m)] if r), None)
        deck, got = (decks[sel[0]], sel[1]) if sel else (None, [None, None, None])
        if sel: usedD.add(sel[0])
        hs = []
        for g in got:
            if g: hs.append({"n": g, "src": ""}); usedH.add(g)
        for k, g in enumerate(got):
            if g: continue
            want = deck["heroes"][k].split("/")[0] if deck else None
            pool = [h for h in own_h if h not in usedH]
            if not pool: hs.append({"n": None, "src": "보유 장수 부족"}); continue
            best = max(pool, key=lambda h: ((sim(want, h) if want else 0), -freq.get(h, 0), -ord(h[0])))
            hs.append({"n": best, "src": f"{want} 대신" if want else "자유 편성"}); usedH.add(best)
        plan.append((deck, hs))
    sts = [unit_state(deck, hs, hd) for deck, hs in plan]
    for rnd in (1, 2):                                       # 1라운드: 모든 장수가 전법 1개씩(티어덱 순서), 2라운드: 남은 전법으로 2번째 칸
        for st in sts:
            pick_tactics(st, len(st["hs"]), own_s, usedS, hd, skd, skn, deal)
    res = []
    for (deck, hs), st in zip(plan, sts):
        tac, ss = finish_tactics(st, hd, skd)
        res.append((deck, [(h, tac.get(h["n"], [])) for h in hs], ss))
    return res


def reco_chunks(text, lim=1900):
    out, cur = [], ""
    for blk in text.split("\n\n"):
        if cur and len(cur) + len(blk) + 2 > lim: out.append(cur); cur = ""
        cur += ("\n\n" if cur else "") + blk
    return out + [cur] if cur else out


ICON = {"한방딜": "💥", "도트딜": "🔥", "일반딜": "⚔️", "회복": "💚", "지원": "🛡️"}


def reco_text(season, res, who, skd=None, deal="도트딜"):
    skd = skd or {}
    lines = [f"🃏 **{who} 추천덱 — 시즌{season}** (티어덱 순서 · 한 번 쓴 장수·전법은 다시 쓰지 않음 · 딜 방식 **{deal}** 우선)",
             "💥한방딜(고배율) 🔥도트딜(누적) ⚔️일반딜 💚회복 🛡️지원 · 책략↔홍수 / 병기↔위협 궁합 반영"]
    for i, (deck, unit, st) in enumerate(res, 1):
        lines.append(f"\n**{i}부대** " + (f"{deck['tier']} {deck['name']} ({'·'.join(deck['heroes'])}) · {deck.get('form') or '진형 미정'}" if deck else "티어덱 없음 · 자유 편성"))
        allt = []
        for h, tac in unit:
            if not h["n"]: lines.append("• (보유 장수 부족)"); continue
            note = f" ← {h['src']}" if h["src"] else ""
            def fmt(t, w):
                ti = tact_info(t, skd); return f"{ICON[ti['kind']]}{t}[{ti['type']}·{ti['trait']}]" + (f"({w})" if w else "")
            lines.append(f"• {h['n']}{note} — " + (", ".join(fmt(t, w) for t, w in tac) if tac else "전법 없음(보유 전법 부족)"))
            allt += [t for t, _ in tac]
        if allt:
            seq = sorted(allt, key=lambda t: ORDER.get(tact_info(t, skd)["type"], 3))
            lines.append("  ↳ 발동순서: " + " → ".join(f"{t}" for t in seq if ORDER.get(tact_info(t, skd)["type"], 3) < 4) + (" → 평타 → " if any(ORDER.get(tact_info(t, skd)["type"], 3) < 4 for t in seq) else "평타 → ") + (", ".join(t for t in seq if ORDER.get(tact_info(t, skd)["type"], 3) > 4) or "(추격 없음)"))
        if st[0] or st[1]:
            lines.append("  ↳ 이상상태: " + (f"부여 {', '.join(st[0])}" if st[0] else "부여 없음") + (f" · 부족 {', '.join(st[1])}" if st[1] else ""))
    lines.append("\n티어덱을 완성하지 못한 부대는 같은 역할의 보유 장수로 채웠습니다. 변경: `/추천덱 시즌:2` · `딜:한방딜`")
    return "\n".join(lines)


def pref_get(did): return (store_get("pref", PREF_PATH, {}) or {}).get(str(did), {})


def pref_set(did, **kw):
    d = dict(store_get("pref", PREF_PATH, {}) or {}); d[str(did)] = {**d.get(str(did), {}), **kw}; store_set("pref", PREF_PATH, d)


PREF_PATH = Path(os.environ.get("PREF_PATH", "pref.json"))


def reco_build(did, season):
    """내 도감(보유 장수·전법) 기준 추천덱 문장. 오류면 (None, 안내문)."""
    cid = next((c for c, d in db_load()["uid"].items() if str(d) == str(did)), None)
    if not cid: return None, "먼저 `/인증`으로 CID 인증을 해 주세요."
    dex = store_get(f"dex_{cid}", dex_path(cid), {}) or {}
    if not dex.get("own"): return None, "보유 장수가 없습니다. `/도감`에서 보유 장수와 전법을 먼저 체크해 주세요."
    sd = tier_data()["SD"]; cfg = tier_sheet_cfg()
    hd, skd = {**sd["HD"], **cfg.get("hd", {})}, {**sd["SKD"], **cfg.get("skd", {})}
    deal = pref_get(did).get("deal", "도트딜")
    res = recommend_decks(dex.get("own", []), dex.get("ownsk", []), season, hd, skd, deal=deal)
    who = (roster.by_uid.get(cid) or {}).get("name") or "내"
    return reco_text(season, res, who, skd, deal), None


class SeasonPick(CtxView):
    def __init__(self, did): super().__init__(timeout=300); self.did = did

    async def _pick(self, it, season):
        await it.response.defer(ephemeral=True)
        await asyncio.to_thread(pref_set, self.did, season=season)          # 선택값 기억
        text, err = await asyncio.to_thread(reco_build, self.did, season)
        ch = reco_chunks(err or text); await it.edit_original_response(content=ch[0], view=None)
        for c in ch[1:]: await it.followup.send(c, ephemeral=True)

    @discord.ui.button(label="시즌1", style=discord.ButtonStyle.primary)
    async def s1(self, it, _): await self._pick(it, 1)

    @discord.ui.button(label="시즌2", style=discord.ButtonStyle.primary)
    async def s2(self, it, _): await self._pick(it, 2)


@bot.tree.command(name="추천덱", description="내 보유 장수·전법으로 티어덱 순서대로 5부대를 편성해 줍니다")
@app_commands.describe(시즌="시즌1 / 시즌2 (처음 한 번 고르면 기억하고, 바꿀 때만 지정)", 딜="전법 우선 방식: 도트딜(누적, 기본) / 한방딜(고배율) — 고르면 기억")
@app_commands.choices(시즌=[app_commands.Choice(name="시즌1", value=1), app_commands.Choice(name="시즌2", value=2)],
                      딜=[app_commands.Choice(name="도트딜", value="도트딜"), app_commands.Choice(name="한방딜", value="한방딜")])
async def cmd_reco(it: discord.Interaction, 시즌: app_commands.Choice[int] = None, 딜: app_commands.Choice[str] = None):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    did = it.user.id
    if 딜 is not None: await asyncio.to_thread(pref_set, did, deal=딜.value)
    if 시즌 is None:
        saved = await asyncio.to_thread(lambda: pref_get(did).get("season"))
        if saved not in (1, 2):                                          # 처음이면 시즌을 묻는다
            return await it.response.send_message("어느 시즌 기준으로 편성할까요? (한 번 고르면 기억합니다)", view=SeasonPick(did), ephemeral=True)
        season = saved
    else:
        season = 시즌.value; await asyncio.to_thread(pref_set, did, season=season)
    await it.response.defer(ephemeral=True)
    text, err = await asyncio.to_thread(reco_build, did, season)
    for chunk in reco_chunks(err or text): await it.followup.send(chunk, ephemeral=True)


@bot.tree.command(name="도감", description="내 보유 장수·전법을 체크하는 개인 링크를 받습니다(시트에 자동 저장)")
async def dex_link(it: discord.Interaction):
    cid = next((c for c, d in db_load()["uid"].items() if str(d) == str(it.user.id)), None)
    if not cid: return await it.response.send_message("먼저 인증(CID 등록)을 완료해 주세요.", ephemeral=True)
    await it.response.send_message(f"🔗 내 도감 링크 (다른 사람에게 공유하지 마세요)\n{PUBLIC_URL}/tier?t={tier_token(cid, cur_gid())}\n장수·전법을 체크하면 자동으로 시트에 저장됩니다.", ephemeral=True)


@bot.tree.command(name="관리", description="[관리자] 톡방 참여자·동맹원 추가·관리자 권한을 DB와 실시간으로 관리하는 페이지 링크")
async def cmd_admin_page(it: discord.Interaction):
    if not it.guild: return await it.response.send_message("서버 안에서 사용해 주세요.", ephemeral=True)
    if not is_admin(it): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    if not SB_ON: return await it.response.send_message("DB가 연결되지 않아 사용할 수 없습니다.", ephemeral=True)
    exp = int(time.time()) + ADM_HOURS * 3600
    asyncio.create_task(guide_post(it.guild))                                   # 가이드북 채널이 없으면 만든다
    await it.response.send_message(f"🛠️ 관리 페이지 (본인 전용 · {ADM_HOURS}시간 유효 · 외부 공유 금지)\n{PUBLIC_URL}/admin?t={adm_token(it.user.id, exp, it.guild_id)}", ephemeral=True)


@bot.tree.command(name="인사", description="[간부] 인사데이터(상황판) 링크를 받습니다")
async def hr_link(it: discord.Interaction):
    if not staff_ok(it, "/인사"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    if not roster.members: return await it.response.send_message("명단 데이터를 아직 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.", ephemeral=True)
    exp = int(time.time()) + HR_HOURS * 3600
    await it.response.send_message(f"📋 인사데이터 링크 (본인 전용 · {HR_HOURS}시간 유효 · 외부 공유 금지)\n{PUBLIC_URL}/hr?t={hr_token(it.user.id, exp, cur_gid())}\n멤버 {len(roster.members)}명 기준", ephemeral=True)


@bot.tree.command(name="인증해제", description="[관리자] CID의 인증 기록을 지우고 실패 횟수를 초기화합니다(재인증 허용)")
@admin
async def unlink(it: discord.Interaction, uid: str):
    db = db_load(); uid = Roster.norm_uid(uid); r = db["uid"].pop(uid, None)
    fails = [d for d, c in db.get("tried", {}).items() if c == uid]              # 그 CID로 인증하다 실패한 계정의 실패 횟수도 초기화
    for d in fails: db["tries"].pop(d, None); db["tried"].pop(d, None)
    db_save(db)
    msg = ("인증 기록을 해제했습니다." if r else "인증 기록은 없습니다.") + (f" 실패 횟수를 초기화했습니다({len(fails)}명)." if fails else " 이 CID로 실패한 기록은 없습니다.")
    await it.response.send_message(msg, ephemeral=True)


# ---------- 정철 타이머 ----------
JC_PATH = Path(os.environ.get("JC_PATH", "jungchul.json"))
JC_RATE = int(os.environ.get("JC_RATE", "160"))            # 시간당 생산량
JC_STOCK = int(os.environ.get("JC_STOCK", "1923"))         # 처음 한 번만 쓰는 현재 보유량
JC_STEP = 300                                              # 5분마다 지급
JC_PAT = [13, 13, 14]                                      # 지급량 반복(15분에 40개 = 시간당 160개)
JC_PHASE = int(os.environ.get("JC_PHASE", "0"))            # 앵커 직후 첫 지급이 패턴의 몇 번째인지(0=13, 2=14)
JC_ANCHOR = datetime.datetime(2026, 10, 25, 21, 23, 10, tzinfo=datetime.timezone(datetime.timedelta(hours=9))).timestamp()  # 실제 지급 시각 기준점
ALERT_MIN = int(os.environ.get("ALERT_MIN", "30"))         # 필요량 도달 N분 전에 '곧 도달' 알림
# 성지(관문) 레벨별 기본 소모량. 12레벨은 관문 기준
COST = {4: 500, 6: 750, 8: 1000, 9: 1000, 10: 1500, 11: 1500, 12: 500, 13: 2000, 14: 2000, 15: 2500, 16: 2500, 18: 3000, 20: 6000}


def extra_cost(level):             # 아군 성지/관문과 관로로 연결되지 않았을 때 추가 소모
    return 500 if level <= 8 else 1500 if level <= 12 else 3000


def jc_load(): return store_get("jungchul", JC_PATH, {"stock": JC_STOCK if cur_gid() == GUILD_ID else 0, "at": time.time(), "rate": JC_RATE, "timers": []})


def jc_save(st): store_set("jungchul", JC_PATH, st)


def _tick_n(t): return int((t - JC_ANCHOR) // JC_STEP)       # t 시점까지 지급된 횟수 번호


def _tick_amt(n): return JC_PAT[(n - 1 + JC_PHASE) % len(JC_PAT)]   # n번째 지급량 (13,13,14 반복)


def jc_stock_now(st, now=None):
    now = now or time.time()
    n0, n1 = _tick_n(st["at"]), _tick_n(now)
    return st["stock"] + sum(_tick_amt(n) for n in range(n0 + 1, n1 + 1))


def jc_eta(st, need):               # need 개 이상이 되는 정확한 시각(timestamp)
    cur, n = st["stock"], _tick_n(st["at"])
    if cur >= need: return st["at"]
    for _ in range(100000):
        n += 1; cur += _tick_amt(n)
        if cur >= need: return JC_ANCHOR + n * JC_STEP
    return float("inf")


def jc_need(level, unlinked=False, enemy=False):
    base = COST[level] + (extra_cost(level) if unlinked else 0)
    return int(base * 1.5) if enemy else base          # 타 진영 시작 지역은 총량의 50% 추가


def fmt_dur(hours):
    m = max(0, int(round(hours * 60)))
    h, m = divmod(m, 60)
    return (f"{h}시간 " if h else "") + f"{m}분"


def jc_label(level): return f"{level}레벨" + (" 관문" if level == 12 else "")


def is_staff(it):
    return isinstance(it.user, discord.Member) and (it.user.guild_permissions.administrator or (it.guild_id, it.user.id) in PERM or any(r.name == OFFICER_ROLE for r in it.user.roles))


CMDPERM_PATH = Path(os.environ.get("CMDPERM_PATH", "cmdperm.json"))


def staff_cmds(): return [c for c, _ in CMD_HELP["간부"]]


def title_norm(t): return (t or "").replace(" ", "")


def known_titles():
    """이 서버에서 고를 수 있는 직책: 기본 간부 직책 + 명단에 있는 직책."""
    ts = {t for t in OFFICER_TITLES}
    ts |= {title_norm(v.get("title")) for v in roster.by_uid.values() if v.get("title")}
    return sorted(t for t in ts if t)


def my_title(did):
    cid = next((c for c, d in db_load()["uid"].items() if str(d) == str(did)), None)
    return title_norm((roster.by_uid.get(cid) or {}).get("title")) if cid else ""


def cmd_titles(cmd):
    """그 명령어를 쓸 수 있게 제한한 직책 목록(비어 있으면 간부 전체)."""
    return list(((store_get("cmdperm", CMDPERM_PATH, {}) or {}).get(cmd)) or [])


def staff_ok(it, cmd):
    """간부 명령어 사용 가능 여부. 관리자(디스코드 관리자·승인 관리자)는 항상 가능, 간부는 관리자가 정한 직책 제한을 따른다."""
    if not is_staff(it): return False
    if is_admin(it): return True
    lim = cmd_titles(cmd)
    return not lim or my_title(it.user.id) in lim


def jc_plan(st):
    """등록 순서대로 정철이 모이는 시각을 시뮬레이션. 앞선 공성에 쓴 정철은 차감하고 계산 → [(공성, 도달시각, 차감 후 잔여)]"""
    cur = st["stock"]; n = _tick_n(st["at"]); t = st["at"]; out = []
    for e in st.get("sieges", []):
        while cur < e["cost"] and len(out) < 1000:
            n += 1; cur += _tick_amt(n); t = JC_ANCHOR + n * JC_STEP
        cur -= e["cost"]; out.append((e, t, cur))
    return out


def siege_name(e): return f"{e['tags']}{jc_label(e['level'])}"

KST = datetime.timezone(datetime.timedelta(hours=9))
BUILD_SEC = int(os.environ.get("BUILD_MIN", "60")) * 60            # 주둔지 건축 소요 시간(분)
NIGHT_START = int(os.environ.get("NIGHT_START", "1"))              # 이 시각(시)부터
NIGHT_END = int(os.environ.get("NIGHT_END", "8"))                  # 이 시각(시)까지는 건축이 멈춤


def build_end(start, dur=BUILD_SEC):
    """start에 건축을 시작해 dur초 작업하면 끝나는 시각. 야간(01~08시)에는 멈췄다가 08시에 재개."""
    t, rem = float(start), float(dur)
    for _ in range(100):
        d = datetime.datetime.fromtimestamp(t, KST)
        ns = d.replace(hour=NIGHT_START, minute=0, second=0, microsecond=0)
        ne = d.replace(hour=NIGHT_END, minute=0, second=0, microsecond=0)
        if ns <= d < ne: t = ne.timestamp(); continue                 # 야간이면 08시까지 대기
        nxt = ns if d < ns else ns + datetime.timedelta(days=1)       # 다음 야간 시작
        avail = nxt.timestamp() - t
        if rem <= avail: return t + rem
        rem -= avail; t = nxt.timestamp()
    return t + rem


def fmt_t(ts): return datetime.datetime.fromtimestamp(ts, KST).strftime("%m.%d.%H:%M")


def _w(s): return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)
def _pad(s, w): return s + " " * max(0, w - _w(s))


def siege_rows(st, now=None):
    """/공성 표용 [(등록자, 레벨표시, 완공예상, 상태)] — 건축 중 + 정철 대기, 완공 빠른 순"""
    now = now or time.time(); rows = []
    for b in st.get("building", []): rows.append((b.get("name", ""), b["level"], b["t_end"], "건축"))
    for e, eta, _ in jc_plan(st): rows.append((e.get("name", ""), e["level"], build_end(max(eta, e.get("reg", 0))), "대기"))      # 실제 시작 시각(jc_settle의 fund)과 동일하게 계산: 정철이 이미 모였으면 모인 시각 기준
    return sorted(rows, key=lambda r: r[2])


@bot.tree.command(name="공성등록", description="[간부] 공성 주둔지를 등록합니다. 정철이 모이면 자동으로 차감하고 알려 줍니다")
@app_commands.describe(레벨="건설할 성지(관문) 레벨", 연결안됨="아군 성지/관문과 관로로 연결되지 않음(추가 소모)", 타진영="타 진영 시작 지역 성지(총량 +50%)")
@app_commands.choices(레벨=[app_commands.Choice(name=jc_label(k), value=k) for k in sorted(COST)])
async def siege_add(it: discord.Interaction, 레벨: app_commands.Choice[int], 연결안됨: bool = False, 타진영: bool = False):
    if not staff_ok(it, "/공성등록"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    st = jc_load(); lv = 레벨.value; cost = jc_need(lv, 연결안됨, 타진영)
    tags = ("연결 안 됨 " if 연결안됨 else "") + ("타 진영 " if 타진영 else "")
    st.setdefault("sieges", []).append({"user": it.user.id, "name": it.user.display_name, "channel": it.channel_id, "level": lv, "cost": cost, "tags": tags, "soon": False, "reg": time.time()})
    st.pop("timers", None); mine = st["sieges"][-1]
    started = jc_settle(st, time.time()); jc_save(st)                     # 정철이 충분하면 즉시 차감하고 건축 시작
    for e, end, left in started:
        if e is mine or (e["reg"], e["user"]) == (mine["reg"], mine["user"]):
            return await it.response.send_message(
                f"🏗️ {siege_name(e)} 등록 — 정철이 충분해 **{cost:,}**개를 바로 차감하고 건축을 시작했습니다. (남은 정철 약 {int(left):,}개)\n"
                f"주둔지 건설 완료 예상 **{fmt_t(end)}** (<t:{int(end)}:R>)" + (" · 01~08시 건축 중지 반영" if end - time.time() > BUILD_SEC + 60 else "") + "\n전체 일정은 `/공성`으로 확인하세요.")
    n = len(st["sieges"]); e, eta, left = jc_plan(st)[-1]; fund = max(eta, time.time()); end = build_end(fund)
    night = " (야간 01~08시 건축 중지 반영)" if build_end(fund, BUILD_SEC) - fund > BUILD_SEC + 1 else ""
    await it.response.send_message(
        f"📌 **{n}번** {siege_name(e)} 등록 — 필요 정철 **{cost:,}**개\n"
        f"정철 도달 예상 <t:{int(fund)}:f> (<t:{int(fund)}:R>) · 도달하면 보유 정철에서 {cost:,}개를 자동 차감하고 건축을 시작합니다.\n"
        f"(봇 계산 보유 정철 약 {int(jc_stock_now(st)):,}개. 게임 내 실제 수량과 다르면 `/정철설정`으로 맞춰야 예상 시각이 정확합니다.)\n"
        f"주둔지 건설 완료 예상 **{fmt_t(end)}**{night}\n"
        f"정철 도달 {ALERT_MIN}분 전, 도달·건축 시작, 주둔지 건설 완료 때 알려 드립니다. 전체 일정은 `/공성`으로 확인하세요.")


@bot.tree.command(name="공성", description="[간부] 현재 정철 개수와 진행 중인 공성 예상 일정을 봅니다(대외비)")
async def siege_view(it: discord.Interaction):
    if not staff_ok(it, "/공성"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    st = jc_load(); now = time.time(); cur = jc_stock_now(st, now)
    nxt = JC_ANCHOR + (_tick_n(now) + 1) * JC_STEP
    head = f"🪙 현재 정철 약 **{int(cur):,}**개 · 5분마다 13·13·14개 지급(시간당 160개) · 다음 지급 <t:{int(nxt)}:R>"
    rows = siege_rows(st, now); rs = sorted(st.get("reserve", []), key=lambda r: r["t"])
    rsv = ("\n**📅 공성 예약**\n```\n" + "\n".join(f"{i}. {fmt_t(r['t'])} | {siege_name(r)} | {r.get('name', '-')}" for i, r in enumerate(rs, 1)) + "\n```") if rs else ""
    if not rows: return await it.response.send_message(head + rsv + "\n등록된 공성이 없습니다. `/공성등록`으로 주둔지를 등록하세요.", ephemeral=True)
    names = [r[0] or "-" for r in rows]; w = max(_w("등록자"), *(_w(n) for n in names))
    lines = [f"{_pad('등록자', w)} | 성지레벨 | 주둔지 건설 완료시간"]
    for n, (_, lv, end, kind) in zip(names, rows):
        lines.append(f"{_pad(n, w)} | {_pad(str(lv) + ('관문' if lv == 12 else ''), 8)} | {fmt_t(end)}{' *' if kind == '대기' else ''}")
    await it.response.send_message(head + "\n```\n" + "\n".join(lines) + "\n```\n`*` 정철이 아직 모이지 않은 대기 건(정철 도달 후 건축 시작). 01~08시에는 건축이 멈추고 08시에 재개됩니다." + rsv, ephemeral=True)


@bot.tree.command(name="공성취소", description="[간부] 등록한 공성을 취소합니다(번호 생략 시 마지막 등록)")
@app_commands.describe(번호="/공성 목록의 번호", 전체="모두 취소")
async def siege_cancel(it: discord.Interaction, 번호: int = 0, 전체: bool = False):
    if not staff_ok(it, "/공성취소"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    st = jc_load(); sg = st.setdefault("sieges", [])
    if 전체: n = len(sg); sg.clear(); st["reserve"] = []; msg = f"공성 {n}건을 모두 취소했습니다."
    elif not sg: return await it.response.send_message("취소할 공성이 없습니다.", ephemeral=True)
    else:
        i = (번호 or len(sg)) - 1
        if not 0 <= i < len(sg): return await it.response.send_message(f"1~{len(sg)} 사이의 번호를 입력하세요.", ephemeral=True)
        e = sg.pop(i); st["reserve"] = [r for r in st.get("reserve", []) if r.get("reg") != e.get("reg")]; msg = f"{i + 1}번 {siege_name(e)} 공성을 취소했습니다."
    jc_save(st); await it.response.send_message(msg)


# ---------- 공성 예약(공성 시각 지정) ----------
DECLARE_SEC = int(os.environ.get("DECLARE_MIN", "60")) * 60        # 선전포고 후 공성까지 최소 소요(분)
SIEGE_ALERT_MIN = int(os.environ.get("SIEGE_ALERT_MIN", "30"))     # 공성 N분 전에 간부 호출


def parse_when(s, now=None):
    """'21:00' '내일 21시 30분' '10.29 21:00' → timestamp(KST). 시각만 쓰면 오늘, 이미 지났으면 내일로 본다."""
    now = now or time.time(); n = datetime.datetime.fromtimestamp(now, KST); s = (s or "").strip()
    day = n.date(); explicit = False
    if "모레" in s: day += datetime.timedelta(days=2); explicit = True
    elif "내일" in s: day += datetime.timedelta(days=1); explicit = True
    elif "오늘" in s: explicit = True
    m = re.search(r"(\d{1,2})\s*[./월]\s*(\d{1,2})\s*일?", s)
    if m:
        try: day = datetime.date(n.year, int(m.group(1)), int(m.group(2))); explicit = True
        except ValueError: return None
        s = s[:m.start()] + " " + s[m.end():]
    tm = re.search(r"(\d{1,2})\s*(?::|시)\s*(\d{1,2})?", s)
    if not tm: return None
    h, mi = int(tm.group(1)), int(tm.group(2) or 0)
    if h > 23 or mi > 59: return None
    d = datetime.datetime(day.year, day.month, day.day, h, mi, tzinfo=KST)
    if d.timestamp() <= now and not explicit: d += datetime.timedelta(days=1)
    return d.timestamp()


def siege_check(st, cost, now=None):
    """지금 이 건을 큐 맨 뒤에 넣었을 때 → (정철 도달, 건축 시작, 건축 완료, 가장 빠른 공성)"""
    now = now or time.time(); tmp = copy.deepcopy(st)
    tmp.setdefault("sieges", []).append({"cost": cost, "level": 0, "tags": "", "reg": now})
    _, eta, _ = jc_plan(tmp)[-1]; fund = max(eta, now); end = build_end(fund)
    return eta, fund, end, end + DECLARE_SEC


@bot.tree.command(name="공성예약", description="[간부] 공성 시각을 정하면 그 시각에 주둔지 건설이 가능한지 알려 주고, 가능하면 예약합니다(30분 전 간부 호출)")
@app_commands.describe(레벨="건설할 성지(관문) 레벨", 시각="공성하고 싶은 시각 예) 21:00 / 내일 21시 30분 / 10.29 21:00", 연결안됨="아군 성지/관문과 관로로 연결되지 않음(추가 소모)", 타진영="타 진영 시작 지역 성지(총량 +50%)", 확인만="예약하지 않고 가능 여부만 확인")
@app_commands.choices(레벨=[app_commands.Choice(name=jc_label(k), value=k) for k in sorted(COST)])
async def siege_reserve(it: discord.Interaction, 레벨: app_commands.Choice[int], 시각: str, 연결안됨: bool = False, 타진영: bool = False, 확인만: bool = False):
    if not staff_ok(it, "/공성예약"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    now = time.time(); T = parse_when(시각, now)
    if T is None: return await it.response.send_message("시각을 읽지 못했습니다. 예) `21:00`, `내일 21시 30분`, `10.29 21:00`", ephemeral=True)
    if T <= now: return await it.response.send_message("이미 지난 시각입니다.", ephemeral=True)
    st = jc_load(); lv = 레벨.value; cost = jc_need(lv, 연결안됨, 타진영)
    tags = ("연결 안 됨 " if 연결안됨 else "") + ("타 진영 " if 타진영 else "")
    eta, fund, end, earliest = siege_check(st, cost, now)
    tl = (f"정철 {cost:,}개 도달 <t:{int(fund)}:f> → 건축 완료 **{fmt_t(end)}** (건축 {BUILD_SEC // 60}분, 01~08시 중지 반영) → "
          f"선전포고 후 최소 {DECLARE_SEC // 60}분 → **{fmt_t(earliest)}** 공성 가능")
    if T < earliest:
        wait = earliest - now; late = earliest - T
        return await it.response.send_message(
            f"❌ **{fmt_t(T)}** 공성은 어렵습니다. {siege_name({'tags': tags, 'level': lv})} 기준 요청 시각보다 **{fmt_dur(late / 3600)}** 부족합니다.\n"
            f"가장 빠른 공성 가능 시각은 **{fmt_t(earliest)}** (지금부터 **{fmt_dur(wait / 3600)}** 뒤) 입니다.\n{tl}", ephemeral=True)
    slack = T - earliest
    if 확인만:
        return await it.response.send_message(f"✅ **{fmt_t(T)}** 공성 가능합니다 (여유 {fmt_dur(slack / 3600)}).\n{tl}\n예약하려면 `확인만` 없이 다시 입력하세요.", ephemeral=True)
    ent = {"user": it.user.id, "name": it.user.display_name, "channel": it.channel_id, "level": lv, "cost": cost, "tags": tags, "soon": False, "reg": now}
    st.setdefault("sieges", []).append(ent); st.pop("timers", None)
    st.setdefault("reserve", []).append({"user": it.user.id, "name": it.user.display_name, "channel": it.channel_id, "level": lv, "tags": tags, "t": T, "alerted": False, "reg": now})
    started = jc_settle(st, now); jc_save(st)
    extra = "정철이 충분해 바로 차감하고 건축을 시작했습니다." if started else "정철이 모이면 자동으로 차감하고 건축을 시작합니다."
    await it.response.send_message(
        f"✅ **{fmt_t(T)}** 공성 예약 완료 — {siege_name(ent)} (여유 {fmt_dur(slack / 3600)})\n{tl}\n{extra}\n"
        f"공성 **{SIEGE_ALERT_MIN}분 전({fmt_t(T - SIEGE_ALERT_MIN * 60)})**에 `{OFFICER_ROLE}` 역할을 이 채널에서 호출합니다. 전체 일정은 `/공성`으로 확인하세요.", ephemeral=True)


@bot.tree.command(name="공성예약취소", description="[간부] 공성 예약을 취소합니다(번호 생략 시 마지막 예약)")
@app_commands.describe(번호="/공성 의 예약 번호")
async def siege_reserve_cancel(it: discord.Interaction, 번호: int = 0):
    if not staff_ok(it, "/공성예약취소"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    st = jc_load(); rs = sorted(st.setdefault("reserve", []), key=lambda r: r["t"])
    if not rs: return await it.response.send_message("취소할 예약이 없습니다.", ephemeral=True)
    i = (번호 or len(rs)) - 1
    if not 0 <= i < len(rs): return await it.response.send_message(f"1~{len(rs)} 사이의 번호를 입력하세요.", ephemeral=True)
    r = rs[i]; st["reserve"] = [x for x in st["reserve"] if x is not r and x.get("reg") != r.get("reg")]
    st["sieges"] = [e for e in st.get("sieges", []) if e.get("reg") != r.get("reg")]       # 같이 등록된 정철 대기 건도 취소(이미 건축 시작한 건은 유지)
    jc_save(st); await it.response.send_message(f"{i + 1}번 예약({fmt_t(r['t'])} · {siege_name(r)})을 취소했습니다.", ephemeral=True)


@bot.tree.command(name="정철설정", description="[간부] 현재 정철 보유량을 맞춥니다")
@app_commands.describe(보유량="지금 동맹 정철 보유량")
async def jc_set(it: discord.Interaction, 보유량: int):
    if not staff_ok(it, "/정철설정"): return await it.response.send_message("권한이 없습니다.", ephemeral=True)
    if 보유량 < 0: return await it.response.send_message("0 이상의 숫자를 입력하세요.", ephemeral=True)
    st = jc_load(); st["stock"] = 보유량; st["at"] = time.time()
    jc_save(st)
    await it.response.send_message(f"정철 보유량을 **{보유량:,}**개로 설정했습니다.")


async def _say(ch_id, text):
    ch = bot.get_channel(ch_id)
    if ch:
        try: await ch.send(text)
        except Exception: log.exception("정철 알림 전송 실패")


async def officer_call(ch_id, text):
    ch = bot.get_channel(ch_id)
    if not ch: log.warning("공성 알림 채널을 찾지 못했습니다: %s", ch_id); return
    role = discord.utils.get(ch.guild.roles, name=OFFICER_ROLE)
    ids = sorted({m.id for m in role.members if not m.bot}) if role else []
    tags = [f"<@{i}>" for i in ids]
    if not tags: tags = [role.mention] if role else []
    head, body = ("", "\n".join([text]))
    chunks, cur = [], ""
    for tg in tags:
        if len(cur) + len(tg) + 1 > 1500: chunks.append(cur); cur = ""
        cur += tg + " "
    if cur: chunks.append(cur)
    am = discord.AllowedMentions(users=True, roles=True)
    try:
        await ch.send((chunks[0] + "\n" if chunks else "") + text, allowed_mentions=am)
        for c in chunks[1:]: await ch.send(c, allowed_mentions=am)
    except Exception: log.exception("공성 알림 전송 실패")


def jc_settle(st, now, quiet_id=None):
    """정철이 충분한 대기 건을 순서대로 차감하고 건축을 시작시킨다. → 시작된 건 목록 [(공성, 완공시각, 남은정철)]"""
    sg = st.setdefault("sieges", []); bl = st.setdefault("building", []); started = []
    while sg and jc_stock_now(st, now) >= sg[0]["cost"]:
        e = sg.pop(0); cur = jc_stock_now(st, now)
        fund = min(now, max(jc_eta(st, e["cost"]), e.get("reg", 0))); end = build_end(fund)
        st["stock"] = cur - e["cost"]; st["at"] = now
        bl.append({"user": e["user"], "name": e.get("name", ""), "channel": e["channel"], "level": e["level"], "tags": e["tags"], "t_start": fund, "t_end": end})
        started.append((e, end, st["stock"]))
    return started


async def jc_tick():
    try:
        st = jc_load(); sg = st.setdefault("sieges", []); bl = st.setdefault("building", [])
        rv = st.setdefault("reserve", [])
        if not sg and not bl and not rv: return
        now = time.time(); changed = False
        for b in [b for b in bl if b["t_end"] <= now]:                    # 축성 완료 알림
            bl.remove(b); changed = True
            await _say(b["channel"], f"🏯 <@{b['user']}> **{b['tags']}{jc_label(b['level'])}** 주둔지 건설이 완료되었습니다!")
        for e, end, left in jc_settle(st, now):                         # 도달 → 자동 차감, 건축 시작
            changed = True
            nxt = f"\n다음: **{siege_name(sg[0])}** {sg[0]['cost']:,}개 대기 중" if sg else ""
            await _say(e["channel"], f"🔔 <@{e['user']}> **{siege_name(e)}** 정철이 모였습니다! 필요 {e['cost']:,}개를 차감하고 건축을 시작합니다. (남은 정철 약 {int(left):,}개)\n🏗️ 주둔지 건설 완료 예상 **{fmt_t(end)}** (<t:{int(end)}:R>){nxt}")
        for e, eta, _ in jc_plan(st):                                   # 도달 직전 알림
            if not e.get("soon") and eta - now <= ALERT_MIN * 60:
                e["soon"] = True; changed = True
                await _say(e["channel"], f"⏰ <@{e['user']}> **{siege_name(e)}** 정철이 곧 모입니다. 약 {fmt_dur((eta - now) / 3600)} 남았습니다. (<t:{int(eta)}:t>)")
        for r in list(rv):                                             # 공성 N분 전 간부 호출
            if r["t"] + 600 < now: rv.remove(r); changed = True; continue
            if not r.get("alerted") and r["t"] - now <= SIEGE_ALERT_MIN * 60:
                r["alerted"] = True; changed = True
                await officer_call(r["channel"], f"⚔️ **공성 {fmt_dur(max(0, r['t'] - now) / 3600)} 전!** {siege_name(r)} 공성이 **{fmt_t(r['t'])}** (<t:{int(r['t'])}:R>)에 예정되어 있습니다. 모두 준비해 주세요. (예약: {r.get('name', '-')})")
        if changed: jc_save(st)
    except Exception: log.exception("정철 타이머 오류")


@tasks.loop(seconds=30)
async def jc_loop():
    for gid in [GUILD_ID] + sorted(LICENSED - {GUILD_ID}):
        set_gid(gid); await jc_tick()
    set_gid(GUILD_ID)


if __name__ == "__main__":
    bot.run(TOKEN)
