#!/usr/bin/env python3
"""녹음 PC 서버 - 녹음 프로그램(rr_rec.py) 기록 보기 (관리자 화면의 '📻 녹음 기록' 탭)

  * 녹음 프로그램은 고치지 않는다. 남기는 기록 파일을 읽기만 한다.
      ~/.local/share/radio_recorder/logs/log_YYYY-MM-DD.txt   (커지면 log_YYYY-MM-DD_old.txt)
      한 줄: [2026-11-23 06:00:00] [INFO] 메시지
  * 녹음 상태: 프로그램이 켜져 있는지(/proc), 지금 쓰는 파일이 커지고 있는지, 디스크 남은 용량
  * 시간표: ~/Recordings/날짜/Radio_주파수_HHMM.mp3 를 시간마다 묶어 길이(크기로 어림)를 본다
  * 상대 PC(AM↔FM)의 기록은 상대 서버의 8081 /api/rec-log-peer 에서 받는다.
    상대 PC 주소(관리자 화면의 본/예비 서버 주소)에서 온 요청만 대답한다.
"""
import datetime
import glob
import json
import os
import re
import shutil
import time
import urllib.error
import urllib.parse
import urllib.request

HOME_DIR = os.path.expanduser("~")
DEFAULT_LOG_DIRS = [os.path.join(HOME_DIR, ".local", "share", "radio_recorder", "logs"),
                    os.path.join(HOME_DIR, "Recordings", "recording_logs")]
DEFAULT_REC_DIR = os.path.join(HOME_DIR, "Recordings")      # 이 PC가 녹음하는 곳 (예비 서버여도 FM 녹음은 여기)
PEER_PORT = 8081
LINE_RE = re.compile(r"^\[(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})\]\s*\[([^\]]{1,20})\]\s?(.*)$")
FILE_RE = re.compile(r"^(?:.*?_)?([0-9]+(?:\.[0-9]+)?)_(\d{2})(\d{2})(?:_\d+)?\.(mp3|m4a|wav|ogg)$", re.I)
AUDIO_RE = re.compile(r"\.(mp3|m4a|wav|ogg)$", re.I)
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ERR_LV = {"ERROR", "ERR", "CRITICAL", "FATAL", "오류", "에러"}
WARN_LV = {"WARN", "WARNING", "경고", "주의"}
ERR_WORDS = re.compile(r"실패|오류|에러|끊김|error|fail|exception|traceback", re.I)
MAX_READ = 3 * 1024 * 1024          # 한 기록 파일에서 읽는 양 (끝에서부터)


def _cfg(c, k, d):
    v = c.get(k)
    return v if v not in (None, "", []) else d


def settings(c):
    """config.json 으로 바꿀 수 있는 값"""
    role = c.get("role", "main")
    return {"label": str(_cfg(c, "rec_label", "FM" if role == "backup" else "AM")),
            "log_dirs": _cfg(c, "rec_log_dirs", DEFAULT_LOG_DIRS),
            "rec_dir": _cfg(c, "rec_own_dir", DEFAULT_REC_DIR),
            "kbps": float(_cfg(c, "rec_kbps", 128)),
            "short_min": int(_cfg(c, "rec_short_min", 50)),
            "stall_min": int(_cfg(c, "rec_stall_min", 5))}


# ---------------------------------------------------------------- 기록 파일
def log_dir(s):
    """기록 파일이 있는 첫 폴더"""
    for d in s["log_dirs"]:
        if glob.glob(os.path.join(d, "log_*.txt")):
            return d
    return ""


def log_dates(s):
    d = log_dir(s)
    out = set()
    for f in glob.glob(os.path.join(d, "log_*.txt")) if d else []:
        m = re.search(r"log_(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        if m:
            out.add(m.group(1))
    return sorted(out, reverse=True)


def _files_for(s, date):
    d = log_dir(s)
    if not d or not DATE_RE.match(date or ""):
        return []
    return [f for f in (os.path.join(d, f"log_{date}_old.txt"), os.path.join(d, f"log_{date}.txt")) if os.path.isfile(f)]


def _tail_text(path, nbytes=MAX_READ):
    with open(path, "rb") as f:
        f.seek(0, 2)
        size = f.tell()
        f.seek(max(0, size - nbytes))
        data = f.read()
    if size > nbytes:
        data = data[data.find(b"\n") + 1:]
    return data.decode("utf-8", errors="replace")


def raw_text(s, date):
    return "".join(_tail_text(f) for f in _files_for(s, date))


def level_of(lv, msg):
    u = (lv or "").upper()
    if u in ERR_LV or lv in ERR_LV:
        return "err"
    if u in WARN_LV or lv in WARN_LV:
        return "warn"
    if ERR_WORDS.search(msg or ""):
        return "warn"
    return "info"


def parse(text):
    rows = []
    for ln in text.splitlines():
        m = LINE_RE.match(ln)
        if m:
            rows.append({"t": m.group(2), "d": m.group(1), "lv": m.group(3).strip(), "m": m.group(4),
                         "c": level_of(m.group(3).strip(), m.group(4))})
        elif ln.strip() and rows:                  # 여러 줄짜리 (오류 내용 등) 는 앞 줄에 붙임
            rows[-1]["m"] += "\n" + ln.rstrip()
            if rows[-1]["c"] == "info" and ERR_WORDS.search(ln):
                rows[-1]["c"] = "warn"
    return rows


def lines(s, date, lv="all", q="", hour="", limit=1500):
    rows = parse(raw_text(s, date))
    counts = {"err": sum(r["c"] == "err" for r in rows), "warn": sum(r["c"] == "warn" for r in rows), "all": len(rows)}
    if lv == "err":
        rows = [r for r in rows if r["c"] == "err"]
    elif lv == "warn":
        rows = [r for r in rows if r["c"] != "info"]
    if hour and re.fullmatch(r"\d{2}", hour):
        rows = [r for r in rows if r["t"].startswith(hour + ":")]
    if q:
        ql = q.lower()
        rows = [r for r in rows if ql in r["m"].lower() or ql in r["lv"].lower()]
    return rows[-limit:], counts


# ---------------------------------------------------------------- 녹음 상태
def recorder_proc():
    """rr_rec 프로그램이 켜져 있으면 (pid, 켜진 시각)"""
    try:
        boot = None
        with open("/proc/stat") as f:
            for ln in f:
                if ln.startswith("btime"):
                    boot = int(ln.split()[1])
        hz = os.sysconf("SC_CLK_TCK")
    except Exception:
        boot, hz = None, 100
    for p in glob.glob("/proc/[0-9]*/cmdline"):
        try:
            with open(p, "rb") as f:
                cmd = f.read().replace(b"\0", b" ").decode("utf-8", "replace")
        except OSError:
            continue
        if "rr_rec" in cmd and "python" in cmd:
            pid = int(p.split("/")[2])
            started = None
            try:
                with open(f"/proc/{pid}/stat") as f:
                    st = f.read().rsplit(")", 1)[1].split()
                if boot:
                    started = boot + int(st[19]) / hz
            except Exception:
                pass
            return pid, started
    return None, None


def _audio(d):
    try:
        with os.scandir(d) as it:
            return [e for e in it if e.is_file() and AUDIO_RE.search(e.name)]
    except OSError:
        return []


def _rec_dates(s, n):
    try:
        ds = sorted((x for x in os.listdir(s["rec_dir"]) if DATE_RE.match(x)), reverse=True)
    except OSError:
        ds = []
    return ds[:n]


def grid(s, days=7):
    """날짜마다 24칸: ok / short / none / now / future (+ 분, 파일)"""
    now = datetime.datetime.now()
    today = now.strftime("%Y-%m-%d")
    bps = s["kbps"] * 1000 / 8
    dates = set(_rec_dates(s, days))
    dates.add(today)
    out = []
    for date in sorted(dates, reverse=True)[:days]:
        hours = {}
        for e in _audio(os.path.join(s["rec_dir"], date)):
            m = FILE_RE.match(e.name)
            if m:
                h = int(m.group(2))
            else:
                h = datetime.datetime.fromtimestamp(e.stat().st_mtime).hour
            x = hours.setdefault(h, {"sec": 0, "files": [], "freq": m.group(1) if m else ""})
            x["sec"] += e.stat().st_size / bps
            x["files"].append(e.name)
        row = []
        for h in range(24):
            x = hours.get(h)
            cur = date == today and h == now.hour
            fut = date > today or (date == today and h > now.hour)
            if fut:
                st = "future"
            elif cur:
                st = "now" if x else "none"
            elif not x:
                st = "none"
            else:
                st = "ok" if x["sec"] >= s["short_min"] * 60 else "short"
            row.append({"h": h, "st": st, "min": round(x["sec"] / 60) if x else 0,
                        "files": sorted(x["files"]) if x else [], "freq": x["freq"] if x else ""})
        out.append({"date": date, "hours": row})
    return out


def status(s):
    now = time.time()
    pid, started = recorder_proc()
    today = datetime.date.today().isoformat()
    newest = None
    for date in _rec_dates(s, 2) + [today]:
        for e in _audio(os.path.join(s["rec_dir"], date)):
            st = e.stat()
            if not newest or st.st_mtime > newest[1]:
                newest = (f"{date}/{e.name}", st.st_mtime, st.st_size)
    try:
        du = shutil.disk_usage(s["rec_dir"])
        disk = {"free": du.free, "total": du.total}
    except OSError:
        disk = None
    d = log_dir(s)
    last_line = None
    for date in log_dates(s)[:1]:
        rows = parse(raw_text(s, date)[-20000:])
        if rows:
            last_line = rows[-1]["d"] + " " + rows[-1]["t"]
    _, counts = lines(s, today) if _files_for(s, today) else ([], {"err": 0, "warn": 0, "all": 0})
    g = grid(s, 1)[0]["hours"]
    nowh = datetime.datetime.now().hour
    missing = [x["h"] for x in g if x["st"] == "none" and x["h"] < nowh]
    short = [x["h"] for x in g if x["st"] == "short"]
    files_today = sum(1 for x in g if x["files"])
    age = (now - newest[1]) if newest else None
    recording = bool(newest and age is not None and age < s["stall_min"] * 60)
    freq = ""
    if newest:
        m = FILE_RE.match(os.path.basename(newest[0]))
        freq = m.group(1) if m else ""
    warn = []
    if not pid:
        warn.append("녹음 프로그램이 꺼져 있습니다")
    if not recording:
        warn.append("녹음 파일이 " + (f"{int(age // 60)}분째 커지지 않습니다" if age is not None else "없습니다"))
    if missing:
        warn.append("오늘 " + ", ".join(f"{h}시" for h in missing[-6:]) + (" 등" if len(missing) > 6 else "") + " 녹음 파일 없음")
    if disk and disk["free"] < 5 * 2 ** 30:
        warn.append(f"녹음 디스크 남은 용량 {disk['free'] / 2 ** 30:.1f} GB")
    if not d:
        warn.append("녹음 프로그램 기록 파일을 찾지 못했습니다")
    return {"label": s["label"], "host": os.uname().nodename, "pid": pid,
            "started": datetime.datetime.fromtimestamp(started).strftime("%Y-%m-%d %H:%M") if started else "",
            "recording": recording, "newest": newest[0] if newest else "", "newest_age": int(age) if age is not None else None,
            "newest_size": newest[2] if newest else 0, "freq": freq, "disk": disk, "log_dir": d, "last_line": last_line or "",
            "today": {"files": files_today, "missing": missing, "short": short, "err": counts["err"], "warn": counts["warn"]},
            "warn": warn, "at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def local(c, q):
    """q: date, lv, q, h, part(status|log|grid 콤마), days"""
    s = settings(c)
    g = lambda k, d="": (q.get(k) or [d])[0]
    parts = set((g("part") or "status,log,grid").split(","))
    date = g("date") or datetime.date.today().isoformat()
    if not DATE_RE.match(date):
        date = datetime.date.today().isoformat()
    out = {"ok": True, "label": s["label"], "date": date}
    if "status" in parts:
        out["status"] = status(s)
    if "log" in parts:
        rows, counts = lines(s, date, g("lv") or "all", g("q").strip()[:60], g("h"))
        out.update({"lines": rows, "counts": counts, "dates": log_dates(s)[:62]})
    if "grid" in parts:
        try:
            days = max(1, min(31, int(g("days") or 7)))
        except ValueError:
            days = 7
        out["grid"] = grid(s, days)
    return out


# ---------------------------------------------------------------- 상대 PC
def peer_get(peer, query, raw=False, timeout=5):
    url = f"http://{peer}:{PEER_PORT}/api/rec-log-peer?" + urllib.parse.urlencode(query)
    with urllib.request.urlopen(url, timeout=timeout) as r:
        data = r.read()
    return data if raw else json.loads(data.decode("utf-8"))


def remote(peer, qd, raw=False):
    if not peer:
        return {"ok": False, "err": "상대 녹음 PC가 등록되어 있지 않습니다."}
    try:
        return peer_get(peer, qd, raw)
    except urllib.error.HTTPError as exc:
        if exc.code == 403:
            return {"ok": False, "err": f"상대 PC({peer})가 이 PC를 상대 서버로 알지 못합니다 (주소가 바뀌었는지 확인)."}
        if exc.code == 404:
            return {"ok": False, "err": f"상대 PC({peer}) 서버가 예전 버전입니다. 1시간 안에 저절로 새 버전을 받습니다."}
        return {"ok": False, "err": f"상대 PC({peer}) 오류 {exc.code}"}
    except Exception as exc:
        return {"ok": False, "err": f"상대 PC({peer})에 연결되지 않습니다 ({str(exc)[:80]})"}
