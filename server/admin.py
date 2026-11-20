#!/usr/bin/env python3
"""녹음 PC 서버 - 관리자 모드 (접속 기록)

  * 관리자 비밀번호: 녹음 PC에서  bash ~/edit/server/admin-password.sh  (웹에서는 만들 수 없음)
  * 웹: http://녹음PC/admin/  -> 비밀번호로 들어가면
      - 접속 기록: 누가(IP·이름) 언제 무엇을 봤는지 (첫 화면 · 운행표 · TV · 녹음 듣기 · 구간 저장 · PDF)
      - 관리자 기록: 관리자 로그인 성공/실패, IP 이름 바꾸기
  * 접속 기록은 nginx 가 /var/log/nginx/edit-access.log 에 한 줄씩 JSON 으로 남긴다 (install.sh 가 설정)
    지난 기록(.1, .2.gz …)도 함께 읽는다 (nginx 기본 정리: 14일).
"""
import datetime
import glob
import gzip
import hashlib
import hmac
import json
import os
import re
import secrets
import sys
import threading
import time
import urllib.parse

HOME = os.environ.get("EDIT_MIRROR_HOME") or os.path.join(os.path.expanduser("~"), "edit-mirror")
CONFIG_FILE = os.path.join(HOME, "config.json")
ADMIN_FILE = os.path.join(HOME, "admin.json")              # 비밀번호(소금+해시)만. 600 권한
EVENTS_FILE = os.path.join(HOME, "admin-events.log")
ACCESS_LOG = os.environ.get("EDIT_ACCESS_LOG") or "/var/log/nginx/edit-access.log"
SESSION_SEC = 12 * 3600
FAIL_MAX, FAIL_WINDOW = 5, 600            # 10분에 5번 틀리면 10분 잠금

_lock = threading.Lock()
_sessions = {}                            # token -> (만료 시각, ip)
_fails = {}                               # ip -> [틀린 시각…]


# ---------------------------------------------------------------- 비밀번호
def _load(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, type(default)) else default
    except Exception:
        return default


def _save(path, data, mode=None):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    if mode:
        os.chmod(tmp, mode)
    os.replace(tmp, path)


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), 200_000).hex()


def has_password():
    return bool(_load(ADMIN_FILE, {}).get("hash"))


def set_password(pw):
    salt = secrets.token_hex(16)
    _save(ADMIN_FILE, {"salt": salt, "hash": _hash(pw, salt),
                       "set": datetime.datetime.now().isoformat(timespec="seconds")}, 0o600)


def check_password(pw):
    d = _load(ADMIN_FILE, {})
    if not d.get("hash") or not d.get("salt"):
        return False
    return hmac.compare_digest(_hash(pw, d["salt"]), d["hash"])


def event(ip, what):
    try:
        os.makedirs(HOME, exist_ok=True)
        with open(EVENTS_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps({"t": datetime.datetime.now().isoformat(timespec="seconds"), "ip": ip,
                                "what": what}, ensure_ascii=False) + "\n")
    except Exception:
        pass


def login(ip, pw):
    """반환: (토큰 또는 None, 알림 글)"""
    now = time.time()
    with _lock:
        f = [t for t in _fails.get(ip, []) if now - t < FAIL_WINDOW]
        _fails[ip] = f
        if len(f) >= FAIL_MAX:
            return None, f"비밀번호를 여러 번 틀려 잠겼습니다. {int((FAIL_WINDOW - (now - f[0])) // 60) + 1}분 뒤 다시 하세요."
    if not has_password():
        return None, "관리자 비밀번호가 아직 없습니다. 녹음 PC에서 bash ~/edit/server/admin-password.sh 를 실행하세요."
    if not check_password(pw or ""):
        with _lock:
            _fails.setdefault(ip, []).append(now)
        event(ip, "로그인 실패")
        return None, "비밀번호가 맞지 않습니다."
    tok = secrets.token_urlsafe(32)
    with _lock:
        _fails.pop(ip, None)
        for k in [k for k, (exp, _) in _sessions.items() if exp < now]:
            _sessions.pop(k, None)
        _sessions[tok] = (now + SESSION_SEC, ip)
    event(ip, "로그인")
    return tok, ""


def session_ok(cookie_header):
    m = re.search(r"(?:^|;\s*)edit_admin=([A-Za-z0-9_\-]+)", cookie_header or "")
    if not m:
        return False
    with _lock:
        s = _sessions.get(m.group(1))
        return bool(s and s[0] > time.time())


def logout(cookie_header):
    m = re.search(r"(?:^|;\s*)edit_admin=([A-Za-z0-9_\-]+)", cookie_header or "")
    if m:
        with _lock:
            _sessions.pop(m.group(1), None)


# ---------------------------------------------------------------- 접속 기록 읽기
def _log_files():
    """새 것부터: edit-access.log, .1, .2.gz …"""
    fs = [ACCESS_LOG] if os.path.exists(ACCESS_LOG) else []
    olds = glob.glob(ACCESS_LOG + ".*")
    num = lambda p: int(re.search(r"\.(\d+)(?:\.gz)?$", p).group(1)) if re.search(r"\.(\d+)(?:\.gz)?$", p) else 999
    return fs + sorted(olds, key=num)


def _read_lines(path):
    op = gzip.open if path.endswith(".gz") else open
    try:
        with op(path, "rt", encoding="utf-8", errors="replace") as f:
            return f.readlines()
    except PermissionError:
        raise
    except Exception:
        return []


AGENTS = (("Edg/", "Edge"), ("Whale/", "웨일"), ("SamsungBrowser", "삼성 인터넷"), ("Chrome/", "크롬"),
          ("Firefox/", "파이어폭스"), ("Safari/", "사파리"), ("curl/", "curl"), ("python", "프로그램"))


def short_agent(ua):
    ua = ua or ""
    b = next((n for k, n in AGENTS if k in ua), "기타")
    dev = "휴대폰" if ("Mobile" in ua or "Android" in ua or "iPhone" in ua) else ("윈도우" if "Windows" in ua else
                                                                                ("맥" if "Macintosh" in ua else ""))
    return f"{b}{' · ' + dev if dev else ''}"


def classify(m, uri):
    """요청 -> (종류, 내용). 화면이 1분마다 묻는 것(상태 등)은 None (기본 화면에서는 숨김)"""
    p = urllib.parse.urlparse(uri)
    path = urllib.parse.unquote(p.path)
    q = urllib.parse.parse_qs(p.query)
    if path in ("/", "/index.html"):
        return "첫 화면", ""
    if path in ("/radio/", "/radio/index.html"):
        return "라디오 운행표", "달력 · 정산 · 검색"
    mm = re.match(r"^/radio/runsheet/(.+)\.pdf$", path)
    if mm:
        return "운행표 PDF", mm.group(1)
    if path in ("/tv/", "/tv/index.html"):
        return "TV 편성표", ""
    mm = re.match(r"^/tv/(?:tv/)?(.+)\.pdf$", path)
    if mm:
        return "TV 편성표 PDF", mm.group(1)
    if path in ("/rec/player.html",):
        return "AM 녹음 플레이어", ""
    if path in ("/rec-fm/player.html",):
        return "FM 녹음 플레이어", ""
    mm = re.match(r"^/rec(-fm)?/(.+\.mp3)$", path)
    if mm:
        return ("FM 녹음 듣기" if mm.group(1) else "AM 녹음 듣기"), mm.group(2)
    if path == "/api/clip" and not q.get("info"):
        src = (q.get("src") or [""])[0]
        s, e = (q.get("s") or ["?"])[0], (q.get("e") or ["?"])[0]
        try:
            span = f"{int(float(s)) // 60}:{int(float(s)) % 60:02d}~{int(float(e)) // 60}:{int(float(e)) % 60:02d}"
        except ValueError:
            span = ""
        return "녹음 구간 저장", f"{os.path.basename(urllib.parse.unquote(src))} {span}".strip()
    if path.startswith("/admin"):
        return "관리자 화면", ""
    if path == "/api/admin/login" and m == "POST":
        return "관리자 로그인", ""
    return None


def access_rows(days=1, show_all=False, ip_filter="", text="", limit=3000):
    """최근 days 일 접속 기록 (새 것부터). 녹음 mp3 는 같은 사람·같은 파일을 30분 안에 다시 받으면 한 줄로"""
    cutoff = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - \
        datetime.timedelta(days=days - 1)      # 1=오늘, 2=어제부터, 7=7일
    rows, seen = [], {}
    names = _load(CONFIG_FILE, {}).get("ip_names") or {}
    stop = False
    for path in _log_files():
        lines = _read_lines(path)
        for ln in reversed(lines):
            try:
                d = json.loads(ln)
            except Exception:
                continue
            try:
                t = datetime.datetime.fromisoformat(d.get("t", "")).replace(tzinfo=None)
            except ValueError:
                continue
            if t < cutoff:
                stop = True
                break
            ip = d.get("ip", "")
            if ip_filter and ip_filter not in ip and ip_filter not in names.get(ip, ""):
                continue
            c = classify(d.get("m", ""), d.get("u", ""))
            if c is None and not show_all:
                continue
            kind, what = c if c else ("기타 요청", urllib.parse.unquote(d.get("u", ""))[:120])
            if text and text not in kind and text not in what:
                continue
            if kind.endswith("녹음 듣기") and not show_all:
                k = (ip, what)
                if k in seen and abs((seen[k]["_t"] - t).total_seconds()) < 1800:
                    seen[k]["n"] += 1
                    continue
            row = {"t": t.strftime("%Y-%m-%d %H:%M:%S"), "_t": t, "ip": ip, "name": names.get(ip, ""),
                   "kind": kind, "what": what, "s": d.get("s", 0), "ua": short_agent(d.get("ua", "")),
                   "b": d.get("b", 0), "n": 1}
            if kind.endswith("녹음 듣기"):
                seen[(ip, what)] = row
            rows.append(row)
            if len(rows) >= limit:
                stop = True
                break
        if stop:
            break
    for r in rows:
        r.pop("_t", None)
    return rows


def summary(rows):
    today = datetime.date.today().isoformat()
    td = [r for r in rows if r["t"].startswith(today)]
    kinds = {}
    for r in rows:
        kinds[r["kind"]] = kinds.get(r["kind"], 0) + 1
    people = {}
    for r in rows:
        p = people.setdefault(r["ip"], {"ip": r["ip"], "name": r["name"], "n": 0, "last": r["t"], "kinds": {}})
        p["n"] += 1
        p["kinds"][r["kind"]] = p["kinds"].get(r["kind"], 0) + 1
    return {"today_ips": len({r["ip"] for r in td}), "today_n": len(td), "kinds": kinds,
            "people": sorted(people.values(), key=lambda p: p["last"], reverse=True)[:50]}


def events(limit=200):
    try:
        with open(EVENTS_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()[-limit:]
        return [json.loads(x) for x in reversed(lines) if x.strip()]
    except Exception:
        return []


def set_ip_name(ip, name, who):
    if not re.fullmatch(r"[0-9a-fA-F:.]{3,45}", ip or ""):
        raise ValueError("IP 주소가 올바르지 않습니다.")
    c = _load(CONFIG_FILE, {})
    names = c.get("ip_names") or {}
    name = (name or "").strip()[:40]
    if name:
        names[ip] = name
    else:
        names.pop(ip, None)
    c["ip_names"] = names
    _save(CONFIG_FILE, c)
    event(who, f"IP 이름: {ip} -> {name or '(지움)'}")


# ---------------------------------------------------------------- 백업
BK = os.path.join(HOME, "backup")
JOBS_STATUS = os.path.join(BK, "status-jobs.json")
EDITORS_STATUS = os.path.join(BK, "status-editors.json")
FM_CONF = "/etc/nginx/edit-extra/fm.conf"
BACKUP_DIRS = {"server": "server", "peer": "peer", "editors_radio": "editors/radio", "editors_tv": "editors/tv",
               "editors_tv_save": "editors/tv-save"}


def peer_addr():
    """상대 PC: 예비 서버면 config 의 peer(본 서버), 본 서버면 fm.conf 의 FM PC"""
    c = _load(CONFIG_FILE, {})
    if c.get("peer"):
        return str(c["peer"])
    try:
        with open(FM_CONF, encoding="utf-8") as f:
            m = re.search(r"proxy_pass\s+http://([0-9.]+)", f.read())
        return m.group(1) if m else ""
    except Exception:
        return ""


def _files(sub):
    d = os.path.join(BK, sub)
    out = []
    for f in sorted(glob.glob(os.path.join(d, "*.zip")), key=os.path.getmtime, reverse=True):
        out.append({"name": os.path.basename(f), "size": os.path.getsize(f),
                    "t": datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M")})
    return out


def backup_status_public():
    """상대 PC 화면에 보여 줄 짧은 상태 (비밀 없음): /api/backup-status"""
    import socket
    jobs = _load(JOBS_STATUS, {})
    ed = _load(EDITORS_STATUS, {})
    c = _load(CONFIG_FILE, {})
    pick = lambda d, ks: {k: d.get(k) for k in ks if k in d}
    return {"host": socket.gethostname(), "role": c.get("role", "main"),
            "server": pick(jobs.get("server", {}), ("at", "ok", "msg", "err", "file", "size")),
            "sync": pick(jobs.get("sync", {}), ("at", "ok", "msg", "err", "last_ok", "dates", "days", "newest_file")),
            "editors": {k: pick(ed.get(k, {}), ("at", "ok", "err", "msg", "from", "file")) for k in ("radio", "tv", "tv_save")}}


def backup_status():
    import urllib.request
    c = _load(CONFIG_FILE, {})
    peer = peer_addr()
    pst = None
    if peer:
        try:
            with urllib.request.urlopen(f"http://{peer}/api/backup-status", timeout=3) as r:
                pst = json.loads(r.read().decode("utf-8"))
        except Exception as exc:
            pst = {"down": True, "err": str(exc)[:120]}
    jobs = _load(JOBS_STATUS, {})
    return {"me": backup_status_public(), "role": c.get("role", "main"), "peer": peer, "peer_status": pst,
            "history": jobs.get("history", [])[:100],
            "files": {k: _files(v) for k, v in BACKUP_DIRS.items()},
            "keep": {"am_keep_days": c.get("am_keep_days", 2), "backup_keep": c.get("backup_keep", 7),
                     "editors_keep": c.get("editors_keep", 14)}}


def run_backup(job, who):
    import subprocess
    if job == "editors":
        os.makedirs(BK, exist_ok=True)
        open(os.path.join(BK, ".editors-now"), "w").close()     # 자료 모으기 서비스가 1분 안에 가져감
        event(who, "편집기 설정 지금 가져오기")
        return "1분 안에 가져옵니다."
    if job not in ("server", "sync"):
        raise ValueError("없는 작업입니다.")
    if job == "sync" and _load(CONFIG_FILE, {}).get("role") != "backup":
        raise ValueError("AM 녹음 복사는 예비 서버(FM PC)에서 합니다. 예비 서버 관리 화면에서 누르세요.")
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup_job.py")
    subprocess.Popen(["nice", "-n", "19", sys.executable, script, job], start_new_session=True,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    event(who, {"server": "설정 백업 지금 하기", "sync": "AM 녹음 지금 복사"}[job])
    return "시작했습니다. 잠시 뒤 [새로 고침]으로 확인하세요."


def set_keep(b, who):
    c = _load(CONFIG_FILE, {})
    lim = {"am_keep_days": (1, 30), "backup_keep": (1, 60), "editors_keep": (1, 90)}
    msg = []
    for k, (lo, hi) in lim.items():
        if k in b:
            v = int(b[k])
            if not lo <= v <= hi:
                raise ValueError(f"{k}: {lo}~{hi}")
            c[k] = v
            msg.append(f"{k}={v}")
    _save(CONFIG_FILE, c)
    event(who, "백업 설정: " + ", ".join(msg))


# ---------------------------------------------------------------- 웹 (edit_mirror 의 8081 처리기에서 부름)
def handle(h, method):
    """처리했으면 True. h: BaseHTTPRequestHandler (h._send, h._json 이 있음)"""
    u = urllib.parse.urlparse(h.path)
    path = urllib.parse.unquote(u.path)
    q = urllib.parse.parse_qs(u.query)
    ip = h.headers.get("X-Real-IP") or h.client_address[0]
    cookie = h.headers.get("Cookie", "")
    if path in ("/admin", "/admin/", "/admin/index.html") and method in ("GET", "HEAD"):
        h._send(200, ADMIN_HTML, "text/html; charset=utf-8")
        return True
    if path == "/api/backup-status":            # 상대 PC 관리자 화면이 읽는 짧은 상태 (비밀 없음)
        h._json(backup_status_public())
        return True
    if not path.startswith("/api/admin/"):
        return False
    if path == "/api/admin/me":
        h._json({"login": session_ok(cookie), "has_password": has_password()})
        return True
    if path == "/api/admin/login" and method == "POST":
        body = _body(h)
        tok, msg = login(ip, body.get("pw", ""))
        if not tok:
            h._json({"error": msg}, 403)
            return True
        h._send(200, json.dumps({"ok": True}), "application/json; charset=utf-8",
                {"Set-Cookie": f"edit_admin={tok}; Path=/; HttpOnly; SameSite=Strict; Max-Age={SESSION_SEC}"})
        return True
    if path == "/api/admin/logout" and method == "POST":
        logout(cookie)
        h._send(200, json.dumps({"ok": True}), "application/json; charset=utf-8",
                {"Set-Cookie": "edit_admin=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0"})
        return True
    if not session_ok(cookie):
        h._json({"error": "로그인이 필요합니다.", "login": False}, 401)
        return True
    g = lambda k, d="": (q.get(k) or [d])[0]
    if path == "/api/admin/access":
        try:
            days = max(1, min(31, int(g("days", "1"))))
        except ValueError:
            days = 1
        try:
            rows = access_rows(days, g("all") == "1", g("ip").strip(), g("q").strip())
        except PermissionError:
            h._json({"error": f"접속 기록 파일({ACCESS_LOG})을 읽을 권한이 없습니다. "
                              "녹음 PC에서 cd ~/edit && git pull && bash server/install.sh 를 한 번 실행하세요."}, 500)
            return True
        h._json({"rows": rows[:2000], "summary": summary(rows), "log": ACCESS_LOG,
                 "log_ok": os.path.exists(ACCESS_LOG)})
        return True
    if path == "/api/admin/backup":
        h._json(backup_status())
        return True
    if path == "/api/admin/backup/run" and method == "POST":
        try:
            h._json({"ok": True, "msg": run_backup(_body(h).get("job", ""), ip)})
        except ValueError as exc:
            h._json({"error": str(exc)}, 400)
        return True
    if path == "/api/admin/backup/keep" and method == "POST":
        try:
            set_keep(_body(h), ip)
            h._json({"ok": True})
        except (ValueError, TypeError) as exc:
            h._json({"error": f"숫자를 확인하세요 ({exc})"}, 400)
        return True
    if path == "/api/admin/backup/file":
        sub = BACKUP_DIRS.get(g("dir"))
        name = g("name")
        f = os.path.join(BK, sub or "x", name)
        if not sub or not re.fullmatch(r"[0-9A-Za-z가-힣_.\-]{1,120}\.zip", name) or not os.path.isfile(f):
            h._json({"error": "파일이 없습니다."}, 404)
            return True
        event(ip, f"백업 파일 받기: {g('dir')}/{name}")
        with open(f, "rb") as fh:
            data = fh.read()
        h._send(200, data, "application/zip",
                {"Content-Disposition": f"attachment; filename*=UTF-8''{urllib.parse.quote(name)}"})
        return True
    if path == "/api/admin/events":
        h._json({"rows": events()})
        return True
    if path == "/api/admin/ipname" and method == "POST":
        b = _body(h)
        try:
            set_ip_name(b.get("ip", ""), b.get("name", ""), ip)
        except ValueError as exc:
            h._json({"error": str(exc)}, 400)
            return True
        h._json({"ok": True})
        return True
    h._json({"error": "없는 주소입니다."}, 404)
    return True


def _body(h):
    try:
        n = min(4096, int(h.headers.get("Content-Length") or 0))
        return json.loads(h.rfile.read(n).decode("utf-8") or "{}") if n else {}
    except Exception:
        return {}


ADMIN_HTML = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>관리자 · 접속 기록</title><script>(function(){var t='light';try{t=localStorage.getItem('edit_theme')||'light'}catch(e){}
var d=t==='dark'||(t==='auto'&&window.matchMedia&&matchMedia('(prefers-color-scheme: dark)').matches);document.documentElement.classList.toggle('dark',d);
window.applyTheme=function(){var t='light';try{t=localStorage.getItem('edit_theme')||'light'}catch(e){}document.documentElement.classList.toggle('dark',t==='dark'||(t==='auto'&&matchMedia('(prefers-color-scheme: dark)').matches))}})()</script>
<style>
*{box-sizing:border-box}body{margin:0;font-family:"Malgun Gothic","맑은 고딕","Noto Sans KR",sans-serif;font-size:14px;background:#f1f5f9;color:#0f172a}
header{background:#0f172a;color:#fff;padding:10px 16px;display:flex;gap:12px;align-items:center}header b{font-size:16px}
header .sp{flex:1}button,select,input{font:inherit;padding:5px 9px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;color:#0f172a}
button{cursor:pointer}button.p{background:#2563eb;color:#fff;border-color:#2563eb}
main{padding:14px 16px;max-width:1300px;margin:0 auto}
.box{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin-bottom:12px}
#login{max-width:360px;margin:60px auto;text-align:center}#login input{width:100%;margin:10px 0;padding:9px}
.kpis{display:flex;gap:10px;flex-wrap:wrap}.kpi{flex:1;min-width:150px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:10px 12px}
.kpi .l{font-size:12px;color:#64748b}.kpi .v{font-size:22px;font-weight:bold;margin-top:2px}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
table{border-collapse:collapse;width:100%}th,td{padding:5px 8px;border-bottom:1px solid #eef2f7;text-align:left;font-size:13px;vertical-align:top}
th{background:#f8fafc;position:sticky;top:0;font-size:12px;color:#475569}td.n{text-align:right}
.k{display:inline-block;padding:1px 7px;border-radius:9px;font-size:12px;background:#e0e7ff;color:#3730a3;white-space:nowrap}
.k.rec{background:#dcfce7;color:#166534}.k.fm{background:#ffedd5;color:#9a3412}.k.pdf{background:#fef9c3;color:#854d0e}.k.adm{background:#fee2e2;color:#991b1b}.k.oth{background:#f1f5f9;color:#64748b}
.ipn{color:#2563eb;cursor:pointer;text-decoration:underline dotted}.muted{color:#64748b;font-size:12px}
.tabs button{border-radius:6px 6px 0 0;border-bottom:none}.tabs button.on{background:#2563eb;color:#fff}
.err{color:#b91c1c;font-weight:bold}.warnbar{background:#fee2e2;color:#991b1b;border:1px solid #fecaca;border-radius:8px;padding:8px 12px;margin-bottom:10px;font-weight:bold}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:10px;margin-bottom:12px}.card{border:1px solid #e5e7eb;border-radius:10px;padding:10px 12px}
.card h4{margin:0 0 6px;font-size:14px}.ok{color:#15803d;font-weight:bold}.bad{color:#b91c1c;font-weight:bold}.grey{color:#94a3b8}
h3{font-size:14px;margin:14px 0 6px}html.dark .card{border-color:#1f2937}html.dark .warnbar{background:#450a0a;color:#fecaca;border-color:#7f1d1d}.scroll{max-height:62vh;overflow:auto}
html.dark body{background:#0b1220;color:#e2e8f0}html.dark .box,html.dark .kpi{background:#111827;border-color:#1f2937}
html.dark th{background:#1f2937;color:#cbd5e1}html.dark td{border-color:#1f2937}html.dark button,html.dark select,html.dark input{background:#1f2937;color:#e2e8f0;border-color:#334155}
html.dark button.p{background:#2563eb}
</style></head><body>
<header><b>🔐 관리자 · 접속 기록</b><span class="muted" style="color:#94a3b8">녹음 PC 웹을 누가 언제 무엇을 봤는지</span><span class="sp"></span>
<button id="out" style="display:none">로그아웃</button></header>
<main>
<div id="login" class="box" style="display:none"><b>관리자 비밀번호</b>
 <input type="password" id="pw" autocomplete="current-password" placeholder="비밀번호"><button class="p" id="go" style="width:100%">들어가기</button>
 <div id="lmsg" class="err" style="margin-top:10px"></div></div>
<div id="app" style="display:none">
 <div class="kpis" id="kpis"></div>
 <div class="tabs" style="margin-top:12px"><button data-t="acc" class="on">접속 기록</button> <button data-t="ppl">사람별</button> <button data-t="ev">관리자 기록</button> <button data-t="bk">💾 백업</button></div>
 <div class="box" style="border-radius:0 10px 10px 10px">
  <div class="bar" id="fbar">기간 <select id="days"><option value="1">오늘</option><option value="2">어제부터</option><option value="7">7일</option><option value="14">14일</option></select>
   IP·이름 <input id="ip" size="14" placeholder="예: 218.237"> 찾기 <input id="q" size="14" placeholder="예: 녹음, 운행표">
   <label><input type="checkbox" id="all"> 모든 요청 (1분마다 묻는 것 포함)</label><button class="p" id="run">보기</button>
   <span class="muted" id="cnt"></span></div>
  <div id="msg" class="err"></div>
  <div class="scroll"><table id="tbl"></table></div>
 </div>
 <div class="muted">· 녹음 듣기는 같은 사람·같은 파일을 30분 안에 다시 받으면 한 줄로 묶어 '횟수'로 보여 줍니다. · IP를 누르면 이름(예: 라디오 편집 PC)을 붙일 수 있습니다. · 기록은 14일 보관합니다.</div>
</div>
</main>
<script>
const $=id=>document.getElementById(id);const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let tab='acc',last=null;
async function J(u,o){const r=await fetch(u,Object.assign({cache:'no-store',credentials:'same-origin'},o||{}));let d={};try{d=await r.json()}catch(e){}
 if(r.status===401){show(false);throw new Error('login')}if(!r.ok)throw new Error(d.error||('HTTP '+r.status));return d}
function show(on){$('login').style.display=on?'none':'';$('app').style.display=on?'':'none';$('out').style.display=on?'':'none';if(!on)setTimeout(()=>$('pw').focus(),50)}
async function me(){try{const d=await J('../api/admin/me');if(d.login){show(true);load()}else{show(false);
 if(!d.has_password)$('lmsg').textContent='관리자 비밀번호가 아직 없습니다. 녹음 PC에서 bash ~/edit/server/admin-password.sh 를 실행하세요.'}}catch(e){}}
async function doLogin(){$('lmsg').textContent='';try{await J('../api/admin/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:$('pw').value})});
 $('pw').value='';show(true);load()}catch(e){if(e.message!=='login')$('lmsg').textContent=e.message}}
$('go').onclick=doLogin;$('pw').addEventListener('keydown',e=>{if(e.key==='Enter')doLogin()});
$('out').onclick=async()=>{try{await J('../api/admin/logout',{method:'POST'})}catch(e){}show(false)};
document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>{tab=b.dataset.t;document.querySelectorAll('.tabs button').forEach(x=>x.classList.toggle('on',x===b));render()});
$('run').onclick=load;['ip','q'].forEach(i=>$(i).addEventListener('keydown',e=>{if(e.key==='Enter')load()}));$('days').onchange=load;$('all').onchange=load;
function kcls(k){return /FM/.test(k)?'fm':/녹음/.test(k)?'rec':/PDF/.test(k)?'pdf':/관리자/.test(k)?'adm':/기타/.test(k)?'oth':''}
function ipCell(ip,name){return '<span class="ipn" data-ip="'+esc(ip)+'" data-n="'+esc(name)+'">'+esc(name||ip)+'</span>'+(name?'<div class="muted">'+esc(ip)+'</div>':'')}
async function load(){$('msg').textContent='';$('cnt').textContent='불러오는 중…';
 try{const u='../api/admin/access?days='+$('days').value+'&all='+($('all').checked?1:0)+'&ip='+encodeURIComponent($('ip').value)+'&q='+encodeURIComponent($('q').value);
  last=await J(u);if(!last.log_ok)$('msg').textContent='접속 기록 파일이 아직 없습니다 ('+last.log+'). 녹음 PC에서 cd ~/edit && git pull && bash server/install.sh 를 실행하세요.';render()}
 catch(e){if(e.message!=='login'){$('msg').textContent=e.message;$('cnt').textContent=''}}}
function render(){if(tab==='ev')return evs();if(tab==='bk')return bk();if(!last)return;const s=last.summary,k=s.kinds;
 const n=x=>Number(k[x]||0);
 $('kpis').innerHTML=[['오늘 접속한 PC',s.today_ips+'대'],['오늘 기록',s.today_n+'건'],['녹음 듣기',n('AM 녹음 듣기')+n('FM 녹음 듣기')+'건'],['구간 저장',n('녹음 구간 저장')+'건'],['운행표·TV',n('라디오 운행표')+n('운행표 PDF')+n('TV 편성표')+n('TV 편성표 PDF')+'건']]
  .map(([l,v])=>'<div class="kpi"><div class="l">'+l+'</div><div class="v">'+v+'</div></div>').join('');
 $('fbar').style.display='';$('cnt').textContent=last.rows.length+'줄';
 if(tab==='ppl'){$('tbl').innerHTML='<tr><th>PC (IP)</th><th>마지막</th><th class="n">기록</th><th>본 것</th></tr>'+s.people.map(p=>'<tr><td>'+ipCell(p.ip,p.name)+'</td><td>'+esc(p.last.slice(5))+'</td><td class="n">'+p.n+'</td><td>'
   +Object.entries(p.kinds).sort((a,b)=>b[1]-a[1]).map(([kk,v])=>'<span class="k '+kcls(kk)+'">'+esc(kk)+' '+v+'</span>').join(' ')+'</td></tr>').join('');}
 else{$('tbl').innerHTML='<tr><th>시각</th><th>PC (IP)</th><th>무엇</th><th>내용</th><th class="n">횟수</th><th>상태</th><th>브라우저</th></tr>'+last.rows.map(r=>'<tr><td>'+esc(r.t.slice(5))+'</td><td>'+ipCell(r.ip,r.name)+'</td><td><span class="k '+kcls(r.kind)+'">'+esc(r.kind)+'</span></td><td>'+esc(r.what)+'</td><td class="n">'+(r.n>1?r.n:'')+'</td><td>'+(r.s>=400?'<b style="color:#b91c1c">'+r.s+'</b>':r.s)+'</td><td class="muted">'+esc(r.ua)+'</td></tr>').join('');}
 bindIp()}
const KB=n=>n>=1073741824?(n/1073741824).toFixed(1)+' GB':n>=1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round((n||0)/1024))+' KB';
const ago=t=>{if(!t)return '';const m=Math.round((Date.now()-new Date(t.replace(' ','T')))/60000);return m<1?'방금':m<60?m+'분 전':m<1440?Math.floor(m/60)+'시간 전':Math.floor(m/1440)+'일 전'};
const okc=(x,lbl)=>!x||!x.at?'<span class="grey">아직 없음</span>':(x.ok?'<span class="ok">✔ '+(lbl||'정상')+'</span>':'<span class="bad">✖ 실패</span>')+' <span class="muted">'+esc(x.at.slice(5,16))+' ('+ago(x.at)+')</span>'+(x.err?'<div class="bad" style="font-weight:normal;font-size:12px">'+esc(x.err)+'</div>':'')+(x.msg?'<div class="muted">'+esc(x.msg)+'</div>':'');
async function bk(){$('fbar').style.display='none';$('cnt').textContent='';$('msg').textContent='';
 let d;try{d=await J('../api/admin/backup')}catch(e){if(e.message!=='login')$('msg').textContent=e.message;return}
 const me=d.me,ps=d.peer_status||{},iamB=d.role==='backup',bkS=iamB?me:(ps.down?null:ps);
 const warn=[];if(d.peer&&ps.down)warn.push((iamB?'본 서버(AM PC)':'예비 서버(FM PC)')+'에 연결되지 않습니다 ('+d.peer+')');
 const sy=bkS&&bkS.sync;if(d.peer&&sy&&sy.last_ok&&(Date.now()-new Date(sy.last_ok.replace(' ','T')))>3*3600e3)warn.push('AM 녹음 복사가 '+ago(sy.last_ok)+' 이후 되지 않았습니다');
 if(sy&&sy.at&&!sy.ok)warn.push('AM 녹음 복사 실패: '+(sy.err||''));
 for(const k of ['radio','tv']){const e=(me.editors||{})[k];if(e&&e.at&&!e.ok)warn.push((k==='radio'?'라디오':'TV')+' 편집기 설정 백업 실패: '+(e.err||''))}
 if(me.server&&me.server.at&&!me.server.ok)warn.push('설정 백업 실패: '+(me.server.err||''));
 const dates=sy&&sy.dates?Object.entries(sy.dates).map(([dd,v])=>esc(dd.slice(5))+' <b>'+v[0]+'</b>개 '+KB(v[1])).join('<br>'):'';
 let h=warn.map(w=>'<div class="warnbar">⚠ '+esc(w)+'</div>').join('');
 h+='<div class="cards">'
  +'<div class="card"><h4>🖥 이 서버: '+esc(me.host)+' · '+(iamB?'예비 서버':'본 서버')+'</h4>'+(d.peer?'상대 '+(iamB?'본 서버':'예비 서버')+': '+esc(d.peer)+' '+(ps.down?'<span class="bad">✖ 연결 안 됨</span>':'<span class="ok">✔ '+esc(ps.host||'')+'</span>'):'<span class="grey">예비 서버 없음</span>')
  +(d.peer&&!iamB?'<div style="margin-top:6px"><a href="http://'+esc(d.peer)+'/admin/" target="_blank">예비 서버 관리 화면 열기 ↗</a> <span class="muted">(사내망 안쪽에서만)</span></div>':'')+'</div>'
  +'<div class="card"><h4>🎧 AM 녹음 복사 (예비 서버, 최근 '+esc((sy&&sy.days)||d.keep.am_keep_days)+'일)</h4>'+(sy?okc(sy)+'<div style="margin-top:4px">'+dates+'</div>'+(sy.newest_file?'<div class="muted">가장 새 파일 '+esc(sy.newest_file)+'</div>':''):'<span class="grey">예비 서버 정보 없음</span>')
  +(iamB?'<div style="margin-top:6px"><button data-run="sync">지금 복사</button></div>':'')+'</div>'
  +'<div class="card"><h4>🗂 설정 백업 (이 서버)</h4>'+okc(me.server)+'<div style="margin-top:6px"><button data-run="server">지금 백업</button></div>'
  +(d.peer&&!ps.down&&ps.server?'<h4 style="margin-top:10px">상대 서버 설정 백업</h4>'+okc(ps.server):'')+'</div>'
  +'<div class="card"><h4>⭐ 편집기 설정 (즐겨찾기·설정)</h4>'+['radio','tv'].map(k=>'<div><b>'+(k==='radio'?'라디오 운행표':'TV 편성표·DMB·엑셀')+'</b> '+okc((me.editors||{})[k])+(((me.editors||{})[k]||{}).from?'<div class="muted">'+esc(me.editors[k].from)+'</div>':'')+'</div>').join('')
  +'<div style="margin-top:6px"><button data-run="editors">지금 가져오기</button> <span class="muted">하루 한 번 자동</span></div></div>'
  +'</div>';
 h+='<h3>보관 설정</h3><div class="bar">AM 녹음 복사 <input id="k1" size="3" value="'+d.keep.am_keep_days+'">일 · 설정 백업 <input id="k2" size="3" value="'+d.keep.backup_keep+'">개 · 편집기 설정 <input id="k3" size="3" value="'+d.keep.editors_keep+'">개 <button id="ksave">저장</button> <span class="muted">AM 녹음 복사 일수는 예비 서버의 설정입니다</span></div>';
 const F=(key,title)=>{const L=d.files[key]||[];return '<h3>'+title+' <span class="muted">('+L.length+'개)</span></h3>'+(L.length?'<table><tr><th>파일</th><th>시각</th><th class="n">크기</th><th></th></tr>'+L.map(f=>'<tr><td>'+esc(f.name)+'</td><td>'+esc(f.t)+'</td><td class="n">'+KB(f.size)+'</td><td><a href="../api/admin/backup/file?dir='+key+'&name='+encodeURIComponent(f.name)+'">⬇ 받기</a></td></tr>').join('')+'</table>':'<div class="muted">없음</div>')};
 h+=F('server','이 서버 설정 백업')+F('peer','상대 서버 설정 백업 (서로 백업)')+F('editors_radio','라디오 편집기 설정')+F('editors_tv','TV 편집기 설정')+F('editors_tv_save','엑셀 편성표 저장할 때마다 (프로젝트·즐겨찾기·설정)');
 h+='<h3>백업 기록 (30일)</h3>'+((d.history||[]).length?'<table><tr><th>시각</th><th>작업</th><th>결과</th><th>내용</th><th class="n">걸린 시간</th></tr>'+d.history.map(x=>'<tr><td>'+esc(x.t.slice(5,16))+'</td><td>'+({server:'설정 백업',sync:'AM 녹음 복사'}[x.job]||esc(x.job))+'</td><td>'+(x.ok?'<span class="ok">성공</span>':'<span class="bad">실패</span>')+'</td><td>'+esc(x.msg)+'</td><td class="n">'+(x.sec||0)+'초</td></tr>').join('')+'</table>':'<div class="muted">아직 없음</div>');
 h+='<h3>복원 방법</h3><div class="muted">설정 백업 zip 안의 <b>복원방법.txt</b>에 새 SSD·새 PC에 다시 꾸미는 순서가 있습니다. 편집기 설정 zip은 편집기 PC의 C:\\EDIT\\radio\\데이터 / C:\\EDIT\\TV\\데이터 에 풀면 됩니다 (편집기를 끄고).</div>';
 $('tbl').innerHTML='<tr><td style="border:none;padding:0">'+h+'</td></tr>';
 document.querySelectorAll('[data-run]').forEach(b=>b.onclick=async()=>{b.disabled=true;try{const r=await J('../api/admin/backup/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({job:b.dataset.run})});alert(r.msg);setTimeout(bk,5000)}catch(e){if(e.message!=='login')alert(e.message)}b.disabled=false});
 $('ksave').onclick=async()=>{try{await J('../api/admin/backup/keep',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({am_keep_days:+$('k1').value,backup_keep:+$('k2').value,editors_keep:+$('k3').value})});alert('저장했습니다.');bk()}catch(e){if(e.message!=='login')alert(e.message)}};
}
async function evs(){$('fbar').style.display='none';try{const d=await J('../api/admin/events');$('cnt').textContent='';
 $('tbl').innerHTML='<tr><th>시각</th><th>IP</th><th>한 일</th></tr>'+d.rows.map(r=>'<tr><td>'+esc(r.t.replace('T',' ').slice(5))+'</td><td>'+esc(r.ip)+'</td><td>'+esc(r.what)+'</td></tr>').join('')}catch(e){}}
function bindIp(){document.querySelectorAll('.ipn').forEach(a=>a.onclick=async()=>{const nm=prompt(a.dataset.ip+' 의 이름 (비우면 지움)\n예: 라디오 편집 PC, 3층 주조',a.dataset.n||'');if(nm===null)return;
 try{await J('../api/admin/ipname',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ip:a.dataset.ip,name:nm})});load()}catch(e){if(e.message!=='login')alert(e.message)}})}
setInterval(()=>{if(!document.hidden&&$('app').style.display===''&&(tab==='acc'||tab==='ppl'))load()},60000);
me();
</script></body></html>"""


if __name__ == "__main__":
    if "--set-password" in sys.argv:
        import getpass
        a = getpass.getpass("새 관리자 비밀번호: ")
        if len(a) < 4:
            sys.exit("4자 이상으로 정해 주세요.")
        if getpass.getpass("한 번 더: ") != a:
            sys.exit("두 번 입력한 것이 다릅니다.")
        set_password(a)
        event("녹음 PC", "관리자 비밀번호 바꿈")
        print(f"관리자 비밀번호를 저장했습니다 ({ADMIN_FILE}). 웹: http://<녹음PC>/admin/")
