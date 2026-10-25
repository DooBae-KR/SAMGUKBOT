"""연맹 디스코드 인증 봇
- 신규 입장자에게 [인증하기] 버튼 안내 → UID/직업/닉네임 입력 폼(모달)
- UID를 엑셀 명단과 대조해 검증 → 별명을 [직업]닉네임 으로 변경 → 역할(동맹원/간부) 부여
"""
import urllib.parse, hmac, hashlib
import os, json, re, csv, io, shutil, logging, urllib.request, datetime, asyncio, time, copy, threading, queue
from aiohttp import web
from pathlib import Path
import discord
from discord import app_commands
from discord.ext import tasks
from dotenv import load_dotenv
import openpyxl

load_dotenv()
TOKEN = os.environ["DISCORD_TOKEN"]
GUILD_ID = int(os.environ["GUILD_ID"])
EXCEL_PATH = os.environ.get("EXCEL_PATH", "261023.xlsx")
SHEET_NAME = os.environ.get("SHEET_NAME", "")            # 비우면 첫 시트
VERIFY_CHANNEL_ID = int(os.environ.get("VERIFY_CHANNEL_ID", "0"))  # 인증 안내 채널
MEMBER_ROLE = os.environ.get("MEMBER_ROLE", "동맹원")
OFFICER_ROLE = os.environ.get("OFFICER_ROLE", "간부")
OFFICER_TITLES = {s.replace(" ", "") for s in os.environ.get("OFFICER_TITLES", "태위,군주,사공,사도,대장군,대홍려,진군장군,무군장군,익군장군,안국장군,분무장군,복파장군,파로장군,호위장군").split(",") if s.strip()}
MAX_TRIES = int(os.environ.get("MAX_TRIES", "5"))
SHEET_ID = os.environ.get("SHEET_ID", "")          # 구글 시트에서 명단을 받을 때(선택)
ROSTER_GID = os.environ.get("ROSTER_GID", "")      # 명단 탭의 gid
TALK_GID = os.environ.get("TALK_GID", "")          # 톡방참여자 탭 gid(선택)
WEB_PORT = int(os.environ.get("PORT") or os.environ.get("WEB_PORT", "8080"))   # Render는 PORT를 자동 지정  # 상황판 웹 포트(0이면 끔)
WEB_KEY = os.environ.get("WEB_KEY", "")            # 설정하면 주소 뒤에 ?key=값 이 있어야 열림
GRACE_END = os.environ.get("GRACE_END", "2026-11-06")
SYNC_MINUTES = int(os.environ.get("SYNC_MINUTES", "10"))   # 시트를 다시 읽고 역할/별명을 맞추는 주기(0이면 끔)
SYNC_REMOVE = os.environ.get("SYNC_REMOVE", "1") == "1"    # 명단에서 사라진 사람의 역할 회수
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


roster = Roster(EXCEL_PATH)


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


def _writer():
    while True:
        name, data = _outq.get()
        try:                                               # 쌓인 요청은 이름별 최신 것만 보냄
            latest = {name: data}
            while not _outq.empty(): n, d = _outq.get_nowait(); latest[n] = d
            for n, d in latest.items():
                for attempt in range(3):
                    try: _remote_put(n, d); break
                    except Exception as e:
                        log.warning("시트 저장 실패(%s) %d/3: %s", n, attempt + 1, e); time.sleep(3)
        except Exception: log.exception("저장 스레드 오류")


if SAVE_URL: threading.Thread(target=_writer, daemon=True).start()


def store_get(name, path, default):
    with _lock:
        if name in _cache: return copy.deepcopy(_cache[name])
        data = None
        if SAVE_URL:
            for attempt in range(3):
                try: data = _remote_get(name); break
                except Exception as e: log.warning("시트 불러오기 실패(%s) %d/3: %s", name, attempt + 1, e); time.sleep(2)
        if data is None and path.exists():
            try: data = json.loads(path.read_text("utf-8"))
            except Exception: pass
        _cache[name] = data if data is not None else default
        return copy.deepcopy(_cache[name])


def store_set(name, path, data):
    with _lock: _cache[name] = copy.deepcopy(data)
    try: path.write_text(json.dumps(data, ensure_ascii=False, indent=1), "utf-8")
    except Exception: pass                                 # 읽기 전용 디스크여도 시트 저장은 계속
    if SAVE_URL: _outq.put((name, copy.deepcopy(data)))


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
class VerifyModal(discord.ui.Modal, title="연맹 인증"):
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
            db["tries"][did] = tries + 1; db_save(db)
            return await it.response.send_message(f"❌ {err} (남은 시도 {MAX_TRIES - tries - 1}회)", ephemeral=True)
        await it.response.defer(ephemeral=True)
        try: role, notes = await apply_verification(it.user, uid, job, nick, info)
        except Exception as e:
            log.exception("역할 부여 실패"); return await it.followup.send(f"처리 중 오류: {e}", ephemeral=True)
        db["uid"][uid] = did; db["tries"].pop(did, None); db_save(db)
        await it.followup.send(f"✅ 인증 완료! 별명: **{make_nick(job, nick)}** · 역할: **{role}**" + ("\n" + "\n".join(notes) if notes else ""), ephemeral=True)


class VerifyView(discord.ui.View):
    def __init__(self): super().__init__(timeout=None)

    @discord.ui.button(label="인증하기", style=discord.ButtonStyle.primary, custom_id="verify_btn")
    async def go(self, it: discord.Interaction, _):
        await it.response.send_modal(VerifyModal())


async def sync_members(guild: discord.Guild):
    """인증된 사람들을 명단(캐릭터 ID)과 대조해 별명/역할을 맞춘다."""
    mem_role = discord.utils.get(guild.roles, name=MEMBER_ROLE)
    off_role = discord.utils.get(guild.roles, name=OFFICER_ROLE)
    res = {"nick": 0, "role": 0, "removed": 0, "fail": 0, "total": 0}
    for uid, did in list(db_load()["uid"].items()):
        m = guild.get_member(int(did))
        if m is None: continue
        res["total"] += 1
        info = roster.by_uid.get(uid)
        try:
            if info is None:
                if SYNC_REMOVE:
                    drop = [r for r in (mem_role, off_role) if r and r in m.roles]
                    if drop: await m.remove_roles(*drop, reason="명단에서 제외"); res["removed"] += 1; await asyncio.sleep(0.4)
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
    log.info("동기화: %s", res)
    return res


class Bot(discord.Client):
    def __init__(self):
        intents = discord.Intents.default(); intents.members = True
        super().__init__(intents=intents); self.tree = app_commands.CommandTree(self)

    async def setup_hook(self):
        if SHEET_ID and ROSTER_GID:
            try: roster.load_csv_text(fetch_sheet_csv()); log.info("구글 시트(인사 탭)에서 명단 %d명 로드", len(roster.by_uid))
            except Exception as e: log.warning("시트 명단 로드 실패, 엑셀 사용: %s", e)
        refresh_talk()
        await asyncio.to_thread(lambda: (db_load(), jc_load()))      # 시트에 저장된 인증·정철 상태를 미리 불러옴
        await start_web()
        self.add_view(VerifyView())
        g = discord.Object(GUILD_ID); self.tree.copy_global_to(guild=g); await self.tree.sync(guild=g)

    async def on_ready(self):
        log.info("로그인: %s", self.user)
        if SYNC_MINUTES and not sync_loop.is_running(): sync_loop.start()
        if not jc_loop.is_running(): jc_loop.start()

    async def on_member_join(self, m: discord.Member):
        if m.guild.id != GUILD_ID or m.bot: return
        text = f"{m.mention} 환영합니다! 아래 **인증하기** 버튼을 눌러 CID · 직업 · 닉네임을 입력해 주세요."
        ch = m.guild.get_channel(VERIFY_CHANNEL_ID) if VERIFY_CHANNEL_ID else None
        try:
            if ch: await ch.send(text, view=VerifyView())
            else: await m.send(text, view=VerifyView())
        except discord.Forbidden: log.warning("안내 메시지 전송 실패: %s", m)


bot = Bot()


@tasks.loop(minutes=max(SYNC_MINUTES, 1))
async def sync_loop():
    try:
        if SHEET_ID and ROSTER_GID:
            text = await asyncio.to_thread(fetch_sheet_csv); roster.load_csv_text(text)
            await asyncio.to_thread(refresh_talk)
        else: roster.load()
        write_html()
        g = bot.get_guild(GUILD_ID)
        if g: await sync_members(g)
    except Exception: log.exception("주기 동기화 실패")

admin = app_commands.default_permissions(manage_roles=True)


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


admin_only = app_commands.default_permissions(administrator=True)


def is_admin(it: discord.Interaction):
    return isinstance(it.user, discord.Member) and it.user.guild_permissions.administrator


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


def render_html():
    """현재 명단/톡방 데이터로 상황판 HTML 문자열을 만든다. 데이터가 없으면 None."""
    if not roster.members: return None
    tpl = (Path(__file__).parent / "dashboard_template.html").read_text("utf-8")
    today = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%Y-%m-%d")
    D = [{**d, "t": 1 if d["n"] in roster.talk else 0} for d in roster.members]
    cfg = {"rosterFile": roster.src or "명단", "rosterDate": today, "talkFile": "구글 시트 톡방참여자 탭" if roster.talk else "없음", "talkDate": today, "graceEnd": GRACE_END}
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
DEX_DIR = Path(os.environ.get("DEX_DIR", "dex"))


def tier_token(cid): return f"{cid}.{hmac.new(TIER_SECRET.encode(), f'dex:{cid}'.encode(), hashlib.sha256).hexdigest()[:20]}"


def tier_cid(t):
    cid, _, sig = (t or "").partition(".")
    return cid if cid and hmac.compare_digest(sig, tier_token(cid).partition(".")[2]) else None


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
    cid = tier_cid(request.query.get("t"))
    if not cid: return web.Response(status=403, text="링크가 올바르지 않습니다. 디스코드에서 /도감 으로 새 링크를 받으세요.")
    info = roster.by_uid.get(cid) or {}
    dex = await asyncio.to_thread(store_get, f"dex_{cid}", dex_path(cid), {})
    cfg = await asyncio.to_thread(tier_sheet_cfg)
    cfg |= {"t": request.query["t"], "api": "/api/dex", "name": info.get("name", ""), "own": dex.get("own", []), "ownsk": dex.get("ownsk", [])}
    tpl = (Path(__file__).parent / "tier_template.html").read_text("utf-8")
    return web.Response(text=tpl.replace("__CFG__", json.dumps(cfg, ensure_ascii=False)), content_type="text/html", headers={"Cache-Control": "no-store"})


async def web_dex_save(request):
    try: b = await request.json()
    except Exception: return web.json_response({"ok": False, "error": "형식 오류"}, status=400)
    cid = tier_cid(b.get("t"))
    if not cid: return web.json_response({"ok": False, "error": "권한 없음"}, status=403)
    clean = lambda v: sorted({str(x)[:30] for x in v if isinstance(x, str)})[:600] if isinstance(v, list) else []
    info = roster.by_uid.get(cid) or {}
    data = {"cid": cid, "name": info.get("name", ""), "own": clean(b.get("own")), "ownsk": clean(b.get("ownsk")),
            "at": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")}
    DEX_DIR.mkdir(exist_ok=True)
    store_set(f"dex_{cid}", dex_path(cid), data)
    return web.json_response({"ok": True})


async def start_web():
    if not WEB_PORT: return
    app = web.Application(); app.router.add_get("/", web_index)
    app.router.add_get("/tier", web_tier); app.router.add_post("/api/dex", web_dex_save)
    app.router.add_get("/health", lambda r: web.Response(text="ok"))      # 외부 핑(UptimeRobot 등)용
    runner = web.AppRunner(app); await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", WEB_PORT).start()
    log.info("상황판: http://localhost:%d/%s", WEB_PORT, f"?key={WEB_KEY}" if WEB_KEY else "")


@bot.tree.command(name="명단업데이트", description="[관리자] 명단을 새로 반영합니다 (엑셀 첨부 또는 구글 시트)")
@admin_only
@app_commands.describe(파일="새 명단 엑셀(.xlsx). 비우면 구글 시트에서 받아옵니다")
async def update_roster(it: discord.Interaction, 파일: discord.Attachment = None):
    if not is_admin(it):
        return await it.response.send_message("관리자만 사용할 수 있습니다.", ephemeral=True)
    await it.response.defer(ephemeral=True)
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
    roster.mtime = 0
    try: roster.load(); await it.response.send_message(f"명단 {len(roster.by_uid)}명 로드, 직업: {', '.join(roster.jobs) or '-'}", ephemeral=True)
    except Exception as e: await it.response.send_message(f"실패: {e}", ephemeral=True)


@bot.tree.command(name="도감", description="내 보유 장수·전법을 체크하는 개인 링크를 받습니다(시트에 자동 저장)")
async def dex_link(it: discord.Interaction):
    cid = next((c for c, d in db_load()["uid"].items() if str(d) == str(it.user.id)), None)
    if not cid: return await it.response.send_message("먼저 인증(CID 등록)을 완료해 주세요.", ephemeral=True)
    await it.response.send_message(f"🔗 내 도감 링크 (다른 사람에게 공유하지 마세요)\n{PUBLIC_URL}/tier?t={tier_token(cid)}\n장수·전법을 체크하면 자동으로 시트에 저장됩니다.", ephemeral=True)


@bot.tree.command(name="인증해제", description="CID의 인증 기록을 지웁니다(재인증 허용)")
@admin
async def unlink(it: discord.Interaction, uid: str):
    db = db_load(); uid = Roster.norm_uid(uid); r = db["uid"].pop(uid, None); db_save(db)
    await it.response.send_message("해제했습니다." if r else "기록이 없습니다.", ephemeral=True)


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


def jc_load(): return store_get("jungchul", JC_PATH, {"stock": JC_STOCK, "at": time.time(), "rate": JC_RATE, "timers": []})


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


@bot.tree.command(name="공성", description="성지 레벨의 정철이 모이면 알려 줍니다")
@app_commands.describe(레벨="목표 성지(관문) 레벨", 연결안됨="아군 성지/관문과 관로로 연결되지 않음(추가 소모)", 타진영="타 진영 시작 지역 성지(총량 +50%)")
@app_commands.choices(레벨=[app_commands.Choice(name=jc_label(k), value=k) for k in sorted(COST)])
async def siege(it: discord.Interaction, 레벨: app_commands.Choice[int], 연결안됨: bool = False, 타진영: bool = False):
    st = jc_load(); lv = 레벨.value; need = jc_need(lv, 연결안됨, 타진영)
    cur = jc_stock_now(st); rem = need - cur; eta = jc_eta(st, need)
    tags = ("연결 안 됨 " if 연결안됨 else "") + ("타 진영 " if 타진영 else "")
    if rem <= 0:
        return await it.response.send_message(f"✅ {tags}{jc_label(lv)} 필요 정철 **{need:,}**개, 지금 보유 **{int(cur):,}**개로 이미 충분합니다.")
    st["timers"] = [t for t in st["timers"] if not (t["user"] == it.user.id and t["level"] == lv)]
    st["timers"].append({"user": it.user.id, "channel": it.channel_id, "level": lv, "need": need, "tags": tags, "soon": False})
    jc_save(st)
    await it.response.send_message(
        f"⏳ {tags}**{jc_label(lv)}** 필요 정철 **{need:,}**개 · 보유 {int(cur):,}개 · 부족 **{int(rem + 0.999):,}**개\n"
        f"5분마다 13·13·14개 지급(시간당 160개) 기준 **{fmt_dur((eta - time.time()) / 3600)}** 뒤 도달 → <t:{int(eta)}:t> (<t:{int(eta)}:R>)\n"
        f"도달 {ALERT_MIN}분 전과 도달했을 때 이 채널에서 {it.user.mention}님을 불러 드립니다.")


@bot.tree.command(name="정철설정", description="[간부] 현재 정철 보유량(과 시간당 생산량)을 맞춥니다")
@app_commands.describe(보유량="지금 동맹 정철 보유량")
async def jc_set(it: discord.Interaction, 보유량: int):
    staff = isinstance(it.user, discord.Member) and (it.user.guild_permissions.administrator or any(r.name == OFFICER_ROLE for r in it.user.roles))
    if not staff: return await it.response.send_message("간부 이상만 설정할 수 있습니다.", ephemeral=True)
    if 보유량 < 0: return await it.response.send_message("0 이상의 숫자를 입력하세요.", ephemeral=True)
    st = jc_load(); st["stock"] = 보유량; st["at"] = time.time()
    jc_save(st)
    await it.response.send_message(f"정철 보유량 **{보유량:,}**개로 설정했습니다.")


@bot.tree.command(name="정철현황", description="현재 정철 보유량과 대기 중인 알림을 봅니다")
async def jc_status(it: discord.Interaction):
    st = jc_load(); cur = jc_stock_now(st)
    lines = [f"현재 정철 약 **{int(cur):,}**개 (시간당 {st['rate']}개)"]
    for t in st["timers"]:
        rem = t["need"] - cur; eta = jc_eta(st, t["need"])
        lines.append(f"· <@{t['user']}> {t['tags']}{jc_label(t['level'])} 필요 {t['need']:,}개 — " + ("도달" if rem <= 0 else f"{fmt_dur((eta - time.time()) / 3600)} 남음 (<t:{int(eta)}:t>)"))
    if not st["timers"]: lines.append("대기 중인 알림이 없습니다.")
    await it.response.send_message("\n".join(lines))


@bot.tree.command(name="공성취소", description="내가 걸어 둔 정철 알림을 모두 취소합니다")
async def siege_cancel(it: discord.Interaction):
    st = jc_load(); n = len(st["timers"]); st["timers"] = [t for t in st["timers"] if t["user"] != it.user.id]; jc_save(st)
    await it.response.send_message(f"알림 {n - len(st['timers'])}개를 취소했습니다.", ephemeral=True)


@tasks.loop(seconds=30)
async def jc_loop():
    try:
        st = jc_load()
        if not st["timers"]: return
        cur = jc_stock_now(st); keep = []; changed = False
        for t in st["timers"]:
            rem = t["need"] - cur; left = jc_eta(st, t["need"]) - time.time()
            ch = bot.get_channel(t["channel"])
            try:
                if rem <= 0:
                    if ch: await ch.send(f"🔔 <@{t['user']}> {t['tags']}**{jc_label(t['level'])}** 정철이 모였습니다! (필요 {t['need']:,}개 · 현재 약 {int(cur):,}개)")
                    changed = True; continue                      # 알림 후 제거
                if not t["soon"] and left <= ALERT_MIN * 60:
                    if ch: await ch.send(f"⏰ <@{t['user']}> {t['tags']}**{jc_label(t['level'])}** 정철이 곧 모입니다. 약 {fmt_dur(left / 3600)} 남았습니다. (부족 {int(rem + 0.999):,}개)")
                    t["soon"] = True; changed = True
            except Exception: log.exception("정철 알림 전송 실패")
            keep.append(t)
        if changed: st["timers"] = keep; jc_save(st)
    except Exception: log.exception("정철 타이머 오류")


if __name__ == "__main__":
    bot.run(TOKEN)
