#!/usr/bin/env python3
"""녹음 PC 서버 (사본 모음)

라디오 운행표 편집기·TV 편성표 편집기가 저장한 데이터를 사내망에서 자동으로 모아 두고,
편집기 PC가 꺼져 있어도 웹으로 항상 볼 수 있게 한다.

  * 사내망 찾기: 편집기/런처가 UDP 50808 로 보내는 '조회 있음' 알림을 듣고 PC 주소를 안다 (IP가 바뀌어도 됨)
  * 모으기: 1분마다 편집기의 조회 웹(읽기 전용)에서
      - 라디오: 데이터 사본(/api/snapshot), 운행표 미리보기 PDF
      - TV   : 주간 편성표 PDF
    를 가져온다. 바뀐 것이 없으면 받지 않는다.
  * 보여 주기:
      - 127.0.0.1:8080  운행표 정산/검색/달력 — 편집기에 들어 있는 조회 화면을 그대로 실행 (사본 데이터로)
      - 0.0.0.0:8081    첫 화면(/), 상태(/api/status), TV 편성표(/tv/)
    nginx 가 http://녹음PC/ 아래에 /radio/ /tv/ /rec/ 로 묶어 준다.

편집 기능은 없다 (모두 읽기 전용). 원본은 언제나 편집기 PC에 있다.
"""
import json
import os
import re
import socket
import sqlite3
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
HOME = os.environ.get("EDIT_MIRROR_HOME") or os.path.join(os.path.expanduser("~"), "edit-mirror")
CONFIG_FILE = os.path.join(HOME, "config.json")

DISCOVERY_PORT = 50808
RADIO_PORT = 8080           # 운행표 조회 화면 (편집기 코드)
HUB_PORT = 8081             # 첫 화면 + TV + 상태
SYNC_EVERY = 60             # 초
PEER_TIMEOUT = 90           # 이 시간 동안 알림이 없으면 '꺼짐'
GIT_PULL_EVERY = 6 * 3600   # 저장소를 새로 받아 조회 화면을 편집기와 같은 버전으로 유지
TV_KEY_RE = re.compile(r"^(schedule|dmb)-[0-9A-Za-z가-힣_\-]{1,60}$")
RUNSHEET_KEY_RE = re.compile(r"^(AM|FM)\d{8}$")
KEEP_DAYS = 40              # 이보다 오래된 미리보기 PDF는 지운다

RADIO_HOME = os.path.join(HOME, "radio")                  # 편집기 코드가 보는 폴더 (RADIO_HOME)
RADIO_DATA = os.path.join(RADIO_HOME, "데이터")
RADIO_DB = os.path.join(RADIO_DATA, "pdf_data_editor.db")
RADIO_PREVIEW = os.path.join(RADIO_DATA, "미리보기")
TV_DIR = os.path.join(HOME, "tv")
STATE_FILE = os.path.join(HOME, "state.json")


def log(*a):
    print(time.strftime("%Y-%m-%d %H:%M:%S"), *a, flush=True)


def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, type(default)) else default
    except Exception:
        return default


def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def write_atomic(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("10.255.255.255", 1))
            return s.getsockname()[0]
        finally:
            s.close()
    except Exception:
        return "127.0.0.1"


def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- 사내망 찾기
class Discovery:
    def __init__(self):
        self.peers = {}          # (kind, ip, port) -> {ip, port, kind, name, seen}
        self.lock = threading.Lock()

    def start(self):
        threading.Thread(target=self._listen, daemon=True).start()

    def _listen(self):
        while True:
            try:
                sk = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                sk.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                sk.bind(("", DISCOVERY_PORT))
                break
            except Exception as exc:
                log("사내망 찾기 포트를 열 수 없음, 30초 뒤 다시:", exc)
                time.sleep(30)
        log(f"사내망 찾기 시작 (UDP {DISCOVERY_PORT})")
        while True:
            try:
                data, addr = sk.recvfrom(2048)
                d = json.loads(data.decode("utf-8"))
                if d.get("app") != "edit-view" or not isinstance(d.get("port"), int):
                    continue
                kind = str(d.get("kind"))
                if kind not in ("radio", "tv"):
                    continue
                with self.lock:
                    key = (kind, addr[0], d["port"])
                    if key not in self.peers:
                        log(f"찾음: {kind} {d.get('name', '')} {addr[0]}:{d['port']}")
                    self.peers[key] = {"ip": addr[0], "port": d["port"], "kind": kind,
                                       "name": str(d.get("name", ""))[:40], "seen": time.time()}
            except Exception:
                time.sleep(0.5)

    def list(self, kind):
        now = time.time()
        with self.lock:
            rows = [dict(v) for v in self.peers.values() if v["kind"] == kind and now - v["seen"] < PEER_TIMEOUT]
        rows.sort(key=lambda r: -r["seen"])
        return rows


def sources(discovery, kind, config):
    """가져올 곳 목록: 사내망에서 찾은 것 + config.json 에 직접 적은 주소 (방송 신호가 막힌 네트워크용)."""
    out, seen = [], set()
    for r in discovery.list(kind):
        url = f"http://{r['ip']}:{r['port']}"
        if url not in seen:
            seen.add(url)
            out.append({"url": url, "name": r["name"]})
    for u in config.get(kind, []) or []:
        u = str(u).strip().rstrip("/")
        if u and not u.startswith("http"):
            u = "http://" + u
        if u and u not in seen:
            seen.add(u)
            out.append({"url": u, "name": "직접 넣은 주소"})
    return out


# ---------------------------------------------------------------- 가져오기
def http_get(url, timeout=15, headers=None):
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, dict(r.headers), r.read()


def http_json(url, timeout=10):
    _, _, body = http_get(url, timeout)
    return json.loads(body.decode("utf-8"))


class Mirror:
    def __init__(self, discovery):
        self.discovery = discovery
        self.state = load_json(STATE_FILE, {})
        self.state.setdefault("radio", {})
        self.state.setdefault("tv", {})
        self.lock = threading.Lock()

    def _set(self, kind, **kw):
        with self.lock:
            self.state[kind].update(kw)
            save_json(STATE_FILE, self.state)

    def status(self):
        with self.lock:
            st = json.loads(json.dumps(self.state))
        # 연결됨 = 최근 3분 안에 실제로 자료를 받아 왔음 (사내망 알림이 막혀 직접 넣은 주소로 받는 경우도 포함)
        for k in ("radio", "tv"):
            x = st[k]
            fresh = bool(x.get("synced")) and not x.get("error") and \
                time.time() - time.mktime(time.strptime(x["synced"], "%Y-%m-%d %H:%M:%S")) < 180
            names = x.get("sources") or [x.get("source_name") or x.get("source") or "편집기"]
            x["online"] = names if fresh else []
        st["server"] = {"name": socket.gethostname(), "ip": lan_ip(), "now": now_str()}
        return st

    # ---- 라디오
    def sync_radio(self, config):
        srcs = sources(self.discovery, "radio", config)
        if not srcs:
            self._set("radio", checked=now_str(), error="편집기 PC를 찾지 못함 (꺼져 있거나 '다른 PC 조회 허용'이 꺼짐)")
            return
        last_err = ""
        for s in srcs:
            try:
                self._sync_radio_from(s)
                return
            except Exception as exc:
                last_err = f"{s['name'] or s['url']}: {exc}"
        self._set("radio", checked=now_str(), error=last_err)

    def _sync_radio_from(self, s):
        base = s["url"]
        since = self.state["radio"].get("tag", "") if self.state["radio"].get("source") == base else ""
        q = "?since=" + urllib.parse.quote(since) if since and os.path.exists(RADIO_DB) else ""
        try:
            code, hdr, body = http_get(base + "/api/snapshot" + q, timeout=60)
        except urllib.error.HTTPError as e:
            if e.code == 304:
                code, hdr, body = 304, dict(e.headers), b""
            elif e.code == 404:
                raise RuntimeError("편집기가 예전 버전입니다 (운행표 편집기 2.12.0 이상으로 업데이트 필요)")
            else:
                raise
        changed = False
        if code == 200 and body:
            tmp = RADIO_DB + ".new"
            write_atomic(tmp, body)
            c = sqlite3.connect(tmp)
            try:
                ok = c.execute("PRAGMA quick_check").fetchone()[0]
                n = c.execute("SELECT COUNT(*) FROM records").fetchone()[0]
            finally:
                c.close()
            if ok != "ok":
                os.remove(tmp)
                raise RuntimeError("받은 데이터 파일 점검 실패")
            os.replace(tmp, RADIO_DB)
            changed = True
            name = urllib.parse.unquote(hdr.get("X-Source-Name", "")) or s["name"]
            self._set("radio", tag=hdr.get("X-Data-Tag", ""), source=base, source_name=name,
                      editor_version=hdr.get("X-App-Version", ""), records=n, data_at=now_str())
            log(f"라디오 데이터 갱신: {name} ({n}건)")
        # 운행표 미리보기 PDF
        rows = http_json(base + "/api/runsheets").get("rows", [])
        local = load_json(os.path.join(RADIO_PREVIEW, "status.json"), {})
        got = 0
        for r in rows:
            key = r.get("key", "")
            if not RUNSHEET_KEY_RE.match(key):
                continue
            ent = {k: r.get(k, "") for k in ("air_date", "media", "state", "saved_at", "updated")}
            ent["dirty"] = bool(r.get("dirty"))
            pdf = os.path.join(RADIO_PREVIEW, key + ".pdf")
            old = local.get(key) or {}
            if r.get("has_pdf") and (not os.path.exists(pdf) or old.get("saved_at") != ent["saved_at"] or old.get("updated") != ent["updated"]):
                _, _, data = http_get(f"{base}/runsheet/{key}.pdf", timeout=60)
                if data[:4] == b"%PDF":
                    write_atomic(pdf, data)
                    got += 1
            local[key] = ent
        self._cleanup(local, RADIO_PREVIEW, "air_date")
        save_json(os.path.join(RADIO_PREVIEW, "status.json"), local)
        if got:
            log(f"운행표 미리보기 {got}개 받음")
        self._set("radio", checked=now_str(), synced=now_str(), source=base,
                  source_name=self.state["radio"].get("source_name") or s["name"], error="",
                  previews=len(rows), changed=changed or bool(got))

    # ---- TV
    def sync_tv(self, config):
        srcs = sources(self.discovery, "tv", config)
        if not srcs:
            self._set("tv", checked=now_str(), error="TV 편성표 PC를 찾지 못함 (런처가 꺼져 있거나 '다른 PC에서 보기'가 꺼짐)")
            return
        local = load_json(os.path.join(TV_DIR, "status.json"), {})
        ok_any, errs, got = False, [], 0
        for s in srcs:
            try:
                d = http_json(s["url"] + "/api/tv")
                for r in d.get("rows", []):
                    key = r.get("key", "")
                    if not TV_KEY_RE.match(key):
                        continue
                    ent = {k: r.get(k, "") for k in ("kind", "title", "src", "state", "saved_at", "updated")}
                    ent["dirty"] = bool(r.get("dirty"))
                    ent["from"] = d.get("name") or s["name"]
                    pdf = os.path.join(TV_DIR, key + ".pdf")
                    old = local.get(key) or {}
                    if r.get("has_pdf") and (not os.path.exists(pdf) or old.get("saved_at") != ent["saved_at"] or old.get("updated") != ent["updated"]):
                        _, _, data = http_get(f"{s['url']}/tv/{urllib.parse.quote(key)}.pdf", timeout=60)
                        if data[:4] == b"%PDF":
                            write_atomic(pdf, data)
                            got += 1
                    local[key] = ent
                ok_any = True
            except Exception as exc:
                errs.append(f"{s['name'] or s['url']}: {exc}")
        self._cleanup(local, TV_DIR, "saved_at")
        save_json(os.path.join(TV_DIR, "status.json"), local)
        if got:
            log(f"TV 편성표 {got}개 받음")
        if ok_any:
            self._set("tv", checked=now_str(), synced=now_str(), error="; ".join(errs), count=len(local),
                      sources=[s["name"] or s["url"] for s in srcs])
        else:
            self._set("tv", checked=now_str(), error="; ".join(errs))

    @staticmethod
    def _cleanup(status, folder, date_field):
        cutoff = time.strftime("%Y-%m-%d", time.localtime(time.time() - KEEP_DAYS * 86400))
        for k in [k for k, v in status.items() if (v.get(date_field) or "9999")[:10] < cutoff]:
            status.pop(k, None)
            try:
                os.remove(os.path.join(folder, k + ".pdf"))
            except OSError:
                pass

    def loop(self):
        while True:
            config = load_json(CONFIG_FILE, {})
            for fn in (self.sync_radio, self.sync_tv):
                try:
                    fn(config)
                except Exception as exc:
                    log("동기화 오류:", fn.__name__, exc)
            time.sleep(int(config.get("interval", SYNC_EVERY)))


# ---------------------------------------------------------------- 운행표 조회 화면 (편집기 코드 그대로)
class RadioView:
    def __init__(self, mirror):
        self.mirror = mirror
        self.mod = None
        self.store = None
        self.srv = None

    def start(self):
        os.environ["RADIO_HOME"] = RADIO_HOME
        os.makedirs(RADIO_PREVIEW, exist_ok=True)
        import importlib.util
        path = os.path.join(REPO, "apps", "runsheet", "pdf_data_editor.py")
        spec = importlib.util.spec_from_file_location("runsheet_mirror_view", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.PREVIEW_DIR = RADIO_PREVIEW
        mod.computer_name = lambda: f"녹음 PC 서버 (사본 · 원본 {self.mirror.state['radio'].get('source_name') or '-'}, {self.mirror.state['radio'].get('synced', '')[5:16]} 기준)"
        self.mod = mod

        def get_store():
            if not os.path.exists(RADIO_DB):
                return None
            if self.store is None:
                self.store = mod.Store(RADIO_DB)
            return self.store

        def peers():
            tv = load_json(os.path.join(TV_DIR, "status.json"), {})
            return [{"ip": lan_ip(), "port": HUB_PORT, "name": "TV 편성표 (녹음 PC 서버 사본)"}] if tv else []

        try:
            self.srv = mod.ViewServer(get_store, "녹음 PC 서버(사본)", peers_fn=peers, announce=False)
        except TypeError:
            self.srv = mod.ViewServer(get_store, "녹음 PC 서버(사본)", peers_fn=peers)
        if not self.srv.start(RADIO_PORT):
            raise RuntimeError(self.srv.error)
        log(f"운행표 조회 화면: http://127.0.0.1:{RADIO_PORT}  (편집기 v{mod.APP_VERSION} 코드)")


# ---------------------------------------------------------------- 첫 화면 · TV · 상태
HUB_HTML = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>방송 자료 조회</title><style>
:root{--side:#0f172a;--side2:#1e293b;--line:#e5e7eb;--muted:#64748b;--acc:#2563eb;--ok:#16a34a;--warn:#d97706;--bad:#dc2626}
*{box-sizing:border-box}
html,body{height:100%;margin:0}
body{font-family:"Malgun Gothic","맑은 고딕","Noto Sans KR",sans-serif;background:#f1f5f9;color:#0f172a;font-size:14px;display:flex}
a{color:inherit}
/* ---- 왼쪽 패널 */
#side{width:260px;flex-shrink:0;background:var(--side);color:#cbd5e1;display:flex;flex-direction:column;height:100vh;overflow:hidden}
.brand{padding:18px 18px 14px;border-bottom:1px solid #1e293b}
.brand b{display:block;color:#fff;font-size:17px;letter-spacing:-.3px}.brand span{font-size:11px;color:#64748b}
.scroll{flex:1;overflow-y:auto;padding:8px 0 12px}
.grp{padding:14px 18px 6px;font-size:11px;color:#64748b;letter-spacing:.5px;font-weight:bold}
.nav{display:flex;align-items:center;gap:10px;padding:9px 18px;margin:1px 8px;border-radius:8px;cursor:pointer;text-decoration:none;color:#cbd5e1}
.nav:hover{background:var(--side2);color:#fff}.nav.on{background:var(--acc);color:#fff}
.nav .i{width:22px;text-align:center;font-size:16px}.nav .n{flex:1}.nav .c{font-size:11px;background:#334155;color:#e2e8f0;border-radius:10px;padding:1px 7px}
.nav.on .c{background:rgba(255,255,255,.25)}
.sub{display:flex;align-items:center;gap:8px;padding:6px 18px 6px 26px;margin:0 8px;border-radius:6px;cursor:pointer;font-size:13px;color:#94a3b8}
.sub:hover{background:var(--side2);color:#fff}.sub.on{background:#1e3a8a;color:#fff}
.sub .d{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;background:#475569}.dot.ok{background:#22c55e}.dot.warn{background:#f59e0b}
.empty{padding:4px 18px 4px 26px;font-size:12px;color:#475569}
.status{border-top:1px solid #1e293b;padding:12px 18px;font-size:12px;line-height:1.75}
.status .row{display:flex;align-items:center;gap:8px}.status .row span{flex:1}.status small{color:#64748b}
#clock{color:#fff;font-size:20px;font-variant-numeric:tabular-nums;margin-bottom:4px}
/* ---- 오른쪽 */
#main{flex:1;display:flex;flex-direction:column;min-width:0;height:100vh}
#top{height:52px;flex-shrink:0;background:#fff;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:12px;padding:0 18px}
#top h1{font-size:16px;margin:0;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb{border:1px solid var(--line);background:#fff;border-radius:8px;padding:6px 12px;font:inherit;font-size:13px;cursor:pointer;text-decoration:none;color:#334155}
.tb:hover{background:#f8fafc;border-color:#cbd5e1}
#menu{display:none}
#frame{flex:1;border:0;width:100%;background:#fff;display:none}
#home{flex:1;overflow-y:auto;padding:22px 26px}
/* ---- 오늘 한눈에 */
.hd{display:flex;align-items:baseline;gap:12px;margin-bottom:16px}.hd b{font-size:22px}.hd span{color:var(--muted)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:18px}
.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:16px 18px;cursor:pointer}
.kpi:hover{border-color:#93c5fd}
.kpi .l{font-size:12px;color:var(--muted)}.kpi .v{font-size:28px;font-weight:bold;margin:4px 0 2px;letter-spacing:-.5px}.kpi .s{font-size:12px;color:var(--muted)}
.badge{display:inline-block;font-size:11px;border-radius:10px;padding:1px 8px;margin-right:4px;font-weight:bold}
.b-ok{background:#dcfce7;color:#166534}.b-warn{background:#fef3c7;color:#92400e}.b-none{background:#f1f5f9;color:#64748b}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;display:flex;flex-direction:column;max-height:520px}
.card h3{margin:0;font-size:14px;padding:12px 16px;border-bottom:1px solid var(--line);display:flex;gap:8px;align-items:center}
.card h3 span{flex:1}.card h3 a{font-size:12px;font-weight:normal;color:var(--acc);text-decoration:none;cursor:pointer}
.card .body{overflow-y:auto}
table{border-collapse:collapse;width:100%}th,td{padding:7px 14px;text-align:left;border-bottom:1px solid #f1f5f9;white-space:nowrap}
th{background:#f8fafc;font-size:12px;color:var(--muted);position:sticky;top:0}td.n,th.n{text-align:right}
.bar{height:6px;background:#dbeafe;border-radius:3px;overflow:hidden;min-width:60px}.bar i{display:block;height:100%;background:#3b82f6}
.none{padding:22px;text-align:center;color:#94a3b8}
.am{color:#1d4ed8;font-weight:bold}.fm{color:#c2410c;font-weight:bold}
@media(max-width:1000px){.cols{grid-template-columns:1fr}}
@media(max-width:820px){#side{position:fixed;left:-270px;top:0;z-index:20;transition:left .2s;box-shadow:4px 0 20px rgba(0,0,0,.3)}
 #side.open{left:0}#menu{display:inline-block}}
</style></head><body>
<aside id="side">
 <div class="brand"><b>📡 방송 자료 조회</b><span id="srv">녹음 PC 서버</span></div>
 <div class="scroll">
  <div class="grp">보기</div>
  <a class="nav" data-p="home"><span class="i">🏠</span><span class="n">오늘 한눈에</span></a>
  <a class="nav" data-p="radio"><span class="i">💰</span><span class="n">라디오 운행표</span></a>
  <a class="nav" data-p="tv"><span class="i">📺</span><span class="n">TV 편성표</span><span class="c" id="ctv"></span></a>
  <a class="nav" data-p="rec"><span class="i">🎧</span><span class="n">녹음 듣기</span></a>
  <div class="grp">운행표 PDF</div><div id="lrs"><div class="empty">불러오는 중…</div></div>
  <div class="grp">TV 편성표 PDF</div><div id="ltv"><div class="empty">불러오는 중…</div></div>
 </div>
 <div class="status">
  <div id="clock">--:--:--</div>
  <div class="row"><i class="dot" id="dr"></i><span>라디오 편집기</span><small id="sr">-</small></div>
  <div class="row"><i class="dot" id="dt"></i><span>TV 편집기</span><small id="st">-</small></div>
  <small id="sm">편집기에서 저장한 내용을 1분마다 모읍니다</small>
 </div>
</aside>
<div id="main">
 <div id="top"><button class="tb" id="menu">☰</button><h1 id="ttl">오늘 한눈에</h1>
  <button class="tb" id="reload">⟳ 새로고침</button><a class="tb" id="newwin" target="_blank" style="display:none">↗ 새 창</a></div>
 <iframe id="frame" title="내용"></iframe>
 <div id="home"></div>
</div>
<script>
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const pad=n=>String(n).padStart(2,'0');const num=v=>Number(v||0).toLocaleString('ko-KR');
const WK='일월화수목금토';
function today(){const d=new Date();return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
function secs(s){s=+s||0;const m=Math.floor(s/60);return m?m+'분 '+(s%60?s%60+'초':''):s+'초'}
const PAGES={home:['🏠 오늘 한눈에',''],radio:['💰 라디오 운행표 — 달력 · 정산 · 검색','radio/'],tv:['📺 TV 주간 편성표','tv/'],rec:['🎧 녹음 듣기','rec/player.html']};
let cur='';
function go(p,src,title){
 if(!PAGES[p]&&!src)p='home';cur=p;
 const [t,u]=PAGES[p]||[title,src];const url=src||u;
 document.querySelectorAll('.nav').forEach(a=>a.classList.toggle('on',a.dataset.p===p));
 document.querySelectorAll('.sub').forEach(a=>a.classList.toggle('on',!!src&&a.dataset.u===src));
 $('ttl').textContent=title||t;
 if(p==='home'){$('frame').style.display='none';$('home').style.display='';$('newwin').style.display='none';home()}
 else{$('home').style.display='none';const f=$('frame');f.style.display='block';if(f.getAttribute('src')!==url)f.src=url;
  $('newwin').href=url;$('newwin').style.display=''}
 if(!src){try{localStorage.setItem('hub_page',p)}catch(e){}}
 if(location.hash||location.search)history.replaceState(null,'',location.pathname);   // 주소창에는 기본 주소만
 $('side').classList.remove('open');
}
document.querySelectorAll('.nav').forEach(a=>a.onclick=()=>go(a.dataset.p));
$('menu').onclick=()=>$('side').classList.toggle('open');
$('reload').onclick=()=>{if(cur==='home')home();else{const f=$('frame');try{f.contentWindow.location.reload()}catch(e){f.src=f.src}}side()};
async function J(u){const r=await fetch(u,{cache:'no-store'});if(!r.ok)throw new Error(r.status);return r.json()}
function sub(box,items,emptyMsg){
 $(box).innerHTML=items.length?items.map(x=>'<div class="sub" data-u="'+esc(x.u)+'" data-t="'+esc(x.t)+'"><i class="dot '+x.dot+'"></i><span class="d">'+x.h+'</span></div>').join(''):'<div class="empty">'+emptyMsg+'</div>';
 $(box).querySelectorAll('.sub').forEach(a=>a.onclick=()=>go('pdf',a.dataset.u,a.dataset.t));
}
function stTxt(x){return x.dirty||x.state!=='수정완료'?['warn','편집중']:['ok','수정완료']}
async function side(){
 try{const d=await J('api/status');$('srv').textContent=d.server.name+' · '+d.server.ip;
  const f=(x,dot,lab)=>{const on=(x.online||[]).length;$(dot).className='dot '+(on?'ok':(x.synced?'warn':''));
   $(lab).textContent=on?'연결됨':(x.synced?'꺼짐':'없음');$(lab).title=x.synced?'마지막으로 모은 시각 '+x.synced:''};
  f(d.radio,'dr','sr');f(d.tv,'dt','st');
  const last=[d.radio.synced,d.tv.synced].filter(Boolean).sort().pop();$('sm').textContent=last?'마지막으로 모은 시각 '+last.slice(5,16):'아직 모은 자료가 없습니다';
 }catch(e){$('sm').textContent='⚠ 서버 상태를 읽지 못했습니다'}
 try{const d=await J('radio/api/runsheets');
  sub('lrs',d.rows.filter(x=>x.has_pdf).slice(0,10).map(x=>{const [c,l]=stTxt(x);const dt=x.air_date.slice(5).replace('-','/');
   return {u:'radio/runsheet/'+x.key+'.pdf',t:'📄 '+x.media+' 운행표 '+x.air_date+' ('+l+')',dot:c,
    h:'<span class="'+(x.media==='FM'?'fm':'am')+'">'+esc(x.media)+'</span> '+esc(dt)+(x.rel?' <small>'+esc(x.rel)+'</small>':'')}}),'운행표가 없습니다')}
 catch(e){$('lrs').innerHTML='<div class="empty">운행표를 읽지 못했습니다</div>'}
 try{const d=await J('tv/api/tv');$('ctv').textContent=d.rows.length||'';
  sub('ltv',d.rows.filter(x=>x.has_pdf).slice(0,8).map(x=>{const [c,l]=stTxt(x);
   return {u:'tv/tv/'+encodeURIComponent(x.key)+'.pdf',t:'📺 '+(x.title||x.key)+' ('+l+')',dot:c,h:esc(x.title||x.src||x.key)}}),'편성표가 없습니다')}
 catch(e){$('ltv').innerHTML='<div class="empty">편성표를 읽지 못했습니다</div>'}
}
async function recToday(){
 try{const r=await fetch('rec/'+today()+'/',{cache:'no-store'});if(!r.ok)return [];
  const doc=new DOMParser().parseFromString(await r.text(),'text/html');
  return [...doc.querySelectorAll('a[href]')].map(a=>decodeURIComponent(a.getAttribute('href'))).filter(h=>/\.(mp3|m4a|wav|ogg)$/i.test(h)).sort()}catch(e){return []}
}
async function home(){
 const t=today(),dd=new Date();
 let h='<div class="hd"><b>'+(dd.getMonth()+1)+'월 '+dd.getDate()+'일 ('+WK[dd.getDay()]+')</b><span>오늘 한눈에 보기</span></div><div class="kpis" id="kp"></div><div class="cols"><div class="card"><h3><span>📋 오늘 광고주별</span><a onclick="go(\'radio\')">달력·정산에서 보기 →</a></h3><div class="body" id="cn"><div class="none">불러오는 중…</div></div></div><div class="card"><h3><span>🕘 오늘 운행 내역 (시각순)</span></h3><div class="body" id="cr"><div class="none">불러오는 중…</div></div></div></div>';
 if(!$('kp'))$('home').innerHTML=h;
 let cal={},day={names:[],rows:[],previews:[]},tvn=0,recs=[];
 try{cal=(await J('radio/api/calendar?ym='+t.slice(0,7))).days||{}}catch(e){}
 try{day=await J('radio/api/day?date='+t)}catch(e){}
 try{tvn=(await J('tv/api/tv')).rows.length}catch(e){}
 recs=await recToday();
 const td=cal[t]||[0,0,0,0];let mc=0,ms=0;for(const k in cal){mc+=cal[k][0];ms+=cal[k][1]}
 const pv=m=>{const x=(day.previews||[]).find(v=>v.media===m);return x?'<span class="badge '+(x.state==='수정완료'?'b-ok':'b-warn')+'">'+m+' '+esc(x.state)+'</span>':'<span class="badge b-none">'+m+' 없음</span>'};
 $('kp').innerHTML=
  '<div class="kpi" onclick="go(\'radio\')"><div class="l">오늘 추가한 광고</div><div class="v">'+num(td[0])+'<small style="font-size:14px"> 건</small></div><div class="s">총 '+secs(td[1])+' · <span class="am">AM '+td[2]+'</span> · <span class="fm">FM '+td[3]+'</span></div></div>'
  +'<div class="kpi" onclick="go(\'radio\')"><div class="l">이번 달 누적</div><div class="v">'+num(mc)+'<small style="font-size:14px"> 건</small></div><div class="s">총 '+secs(ms)+' · '+Object.keys(cal).length+'일</div></div>'
  +'<div class="kpi" onclick="go(\'radio\')"><div class="l">오늘 운행표</div><div class="v" style="font-size:16px;margin:12px 0 8px">'+pv('AM')+pv('FM')+'</div><div class="s">편집기에서 저장하면 바뀝니다</div></div>'
  +'<div class="kpi" onclick="go(\'tv\')"><div class="l">TV 편성표</div><div class="v">'+tvn+'<small style="font-size:14px"> 개</small></div><div class="s">주·판별 PDF</div></div>'
  +'<div class="kpi" onclick="go(\'rec\')"><div class="l">오늘 녹음</div><div class="v">'+recs.length+'<small style="font-size:14px"> 개</small></div><div class="s">'+(recs.length?'최근 '+esc(recs[recs.length-1]).replace(/^.*_(\d\d)(\d\d)\.\w+$/,'$1:$2')+' 시작':'아직 없음')+'</div></div>';
 const names=day.names||[],mx=Math.max(1,...names.map(n=>n[2]));
 $('cn').innerHTML=names.length?'<table><tr><th>광고주</th><th>그룹</th><th class="n">횟수</th><th class="n">총 초수</th><th></th></tr>'+names.map(n=>'<tr><td><b>'+esc(n[0])+'</b></td><td>'+esc(n[1])+'</td><td class="n">'+n[2]+'</td><td class="n">'+num(n[3])+'</td><td style="width:30%"><div class="bar"><i style="width:'+(n[2]/mx*100)+'%"></i></div></td></tr>').join('')+'</table>':'<div class="none">오늘 추가한 광고가 없습니다</div>';
 const rows=day.rows||[];
 $('cr').innerHTML=rows.length?'<table><tr><th>시각</th><th>매체</th><th>프로그램</th><th>광고주</th><th class="n">초</th></tr>'+rows.map(r=>'<tr><td>'+esc(r[0])+'</td><td class="'+(r[1]==='FM'?'fm':'am')+'">'+esc(r[1]||'AM')+'</td><td>'+esc(r[2])+'</td><td>'+esc(r[3])+'</td><td class="n">'+r[4]+'</td></tr>').join('')+'</table>':'<div class="none">운행 내역이 없습니다</div>';
}
function tick(){const d=new Date();$('clock').textContent=pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds())}
setInterval(tick,1000);tick();
let first=(location.hash||'').slice(1);if(!PAGES[first]){try{first=localStorage.getItem('hub_page')||'home'}catch(e){first='home'}}
side();go(PAGES[first]?first:'home');
setInterval(()=>{if(!document.hidden){side();if(cur==='home')home()}},60000);
</script></body></html>"""

TV_HTML = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TV 주간 편성표 (조회 전용)</title><style>
body{font-family:"Malgun Gothic","맑은 고딕","Noto Sans KR",sans-serif;margin:0;background:#f5f6f8;color:#1f2937;font-size:14px}
header{background:#7c2d12;color:#fff;padding:10px 18px}header b{font-size:17px}header a{color:#fed7aa;font-size:13px;margin-right:12px;text-decoration:none}
main{padding:14px 18px;max-width:1100px}table{border-collapse:collapse;width:100%;background:#fff}
th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left}th{background:#fff7ed}.msg{color:#6b7280;font-size:12px;margin:6px 0}
</style></head><body><header><a href="../" id="back">← 처음</a><b>📺 TV 주간 편성표</b> <span id="who" style="font-size:12px;opacity:.85"></span></header>
<main><div class="msg" id="msg">불러오는 중…</div><table id="tbl"></table></main><script>
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
if(window.top!==window.self)document.getElementById('back').style.display='none';
async function load(){try{const d=await (await fetch('api/tv')).json();document.getElementById('who').textContent=' — '+d.name;
 let h='<tr><th>편성표</th><th>종류</th><th>상태</th><th>마지막 저장</th><th>편집기 PC</th><th>미리보기 · 인쇄</th></tr>';
 for(const x of d.rows){const st=x.dirty?'<b style="color:#c2410c">✏ 편집중</b> <span style="font-size:12px;color:#9a3412">(마지막 저장 상태)</span>':(x.state==='수정완료'?'<b style="color:#15803d">✔ 수정완료</b>':'<b style="color:#c2410c">✏ 편집중</b>');
  h+='<tr><td>'+esc(x.title||x.src)+'</td><td>'+esc(x.kind)+'</td><td>'+st+'</td><td>'+esc((x.saved_at||'').slice(5,16))+'</td><td>'+esc(x.from)+'</td><td>'+(x.has_pdf?'<a target="_blank" href="tv/'+encodeURIComponent(x.key)+'.pdf">📄 보기 · 인쇄</a>':'<span style="color:#9ca3af">아직 저장 전</span>')+'</td></tr>'}
 document.getElementById('tbl').innerHTML=h;document.getElementById('msg').textContent=d.rows.length?('TV 편집기에서 저장한 편성표를 1분마다 모읍니다. 마지막으로 모은 시각: '+(d.synced||'-').slice(5,16)):'아직 모은 편성표가 없습니다.'}
 catch(e){document.getElementById('msg').textContent='⚠ 연결하지 못했습니다.'}}
load();setInterval(()=>{if(!document.hidden)load()},30000);
</script></body></html>"""


def tv_rows():
    st = load_json(os.path.join(TV_DIR, "status.json"), {})
    out = []
    for k, v in st.items():
        if not TV_KEY_RE.match(k):
            continue
        r = dict(v)
        r["key"] = k
        r["has_pdf"] = os.path.exists(os.path.join(TV_DIR, k + ".pdf"))
        out.append(r)
    out.sort(key=lambda r: r.get("saved_at") or r.get("updated") or "", reverse=True)
    return out


def make_hub_handler(mirror):
    class H(BaseHTTPRequestHandler):
        server_version = "EditMirror/1"

        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype, extra=None):
            data = body.encode("utf-8") if isinstance(body, str) else body
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Access-Control-Allow-Origin", "*")       # 운행표 조회 화면의 'TV' 탭이 다른 포트에서 읽는다
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            if self.command != "HEAD":
                self.wfile.write(data)

        def _json(self, obj, code=200):
            self._send(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

        def do_HEAD(self):
            self.do_GET()

        def do_GET(self):
            try:
                path = urllib.parse.unquote(urllib.parse.urlparse(self.path).path)
                if path in ("/", "/index.html"):
                    return self._send(200, HUB_HTML, "text/html; charset=utf-8")
                if path == "/api/status":
                    return self._json(mirror.status())
                if path in ("/tv", "/tv/", "/tv/index.html"):
                    return self._send(200, TV_HTML, "text/html; charset=utf-8")
                if path in ("/api/tv", "/tv/api/tv"):
                    return self._json({"app": "edit-tv-view", "name": "녹음 PC 서버 (사본)",
                                       "synced": mirror.state["tv"].get("synced", ""), "rows": tv_rows()})
                m = re.fullmatch(r"/tv/(?:tv/)?(.+)\.pdf", path)
                if m and TV_KEY_RE.match(m.group(1)):
                    f = os.path.join(TV_DIR, m.group(1) + ".pdf")
                    if os.path.exists(f):
                        with open(f, "rb") as fh:
                            data = fh.read()
                        return self._send(200, data, "application/pdf",
                                          {"Content-Disposition": f"inline; filename*=UTF-8''{urllib.parse.quote(m.group(1))}.pdf"})
                return self._json({"error": "없는 주소입니다."}, 404)
            except (BrokenPipeError, ConnectionResetError):
                pass
            except Exception as exc:
                try:
                    self._json({"error": str(exc)}, 500)
                except Exception:
                    pass

        def _deny(self):
            self._json({"error": "조회 전용입니다."}, 405)

        do_POST = do_PUT = do_DELETE = do_PATCH = _deny
    return H


# ---------------------------------------------------------------- 저장소 새로 받기 (조회 화면을 편집기와 같은 버전으로)
def git_pull_loop():
    while True:
        time.sleep(GIT_PULL_EVERY)
        try:
            before = subprocess.run(["git", "-C", REPO, "rev-parse", "HEAD"], capture_output=True, text=True, timeout=30).stdout.strip()
            r = subprocess.run(["git", "-C", REPO, "pull", "--ff-only", "-q"], capture_output=True, text=True, timeout=300)
            after = subprocess.run(["git", "-C", REPO, "rev-parse", "HEAD"], capture_output=True, text=True, timeout=30).stdout.strip()
            if r.returncode != 0:
                log("저장소 받기 실패:", r.stderr.strip()[:200])
            elif before != after:
                log("새 버전을 받았습니다. 다시 시작합니다.")
                os._exit(3)               # systemd 가 다시 켠다
        except Exception as exc:
            log("저장소 받기 오류:", exc)


def main():
    os.makedirs(RADIO_PREVIEW, exist_ok=True)
    os.makedirs(TV_DIR, exist_ok=True)
    discovery = Discovery()
    discovery.start()
    mirror = Mirror(discovery)
    try:
        RadioView(mirror).start()
    except Exception as exc:
        log("운행표 조회 화면을 열지 못함:", exc)
    hub = ThreadingHTTPServer(("0.0.0.0", HUB_PORT), make_hub_handler(mirror))
    hub.daemon_threads = True
    threading.Thread(target=hub.serve_forever, daemon=True).start()
    log(f"첫 화면 · TV: http://{lan_ip()}:{HUB_PORT}")
    if "--no-git" not in sys.argv:
        threading.Thread(target=git_pull_loop, daemon=True).start()
    time.sleep(6)                     # 알림을 한두 번 들은 뒤 첫 동기화
    mirror.loop()


if __name__ == "__main__":
    main()
