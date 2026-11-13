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
        st["radio"]["online"] = [s["name"] or s["url"] for s in sources(self.discovery, "radio", {})]
        st["tv"]["online"] = [s["name"] or s["url"] for s in sources(self.discovery, "tv", {})]
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
body{font-family:"Malgun Gothic","맑은 고딕","Noto Sans KR",sans-serif;margin:0;background:#f5f6f8;color:#1f2937}
header{background:#111827;color:#fff;padding:14px 22px}header b{font-size:19px}header span{opacity:.7;font-size:12px;margin-left:10px}
main{max-width:1000px;margin:0 auto;padding:22px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px}
a.card{display:block;background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:20px;text-decoration:none;color:inherit;box-shadow:0 1px 2px rgba(0,0,0,.04)}
a.card:hover{border-color:#3b82f6;box-shadow:0 4px 14px rgba(59,130,246,.15)}
.ic{font-size:30px}.t{font-size:18px;font-weight:bold;margin:8px 0 4px}.d{font-size:13px;color:#6b7280}
.st{margin-top:12px;font-size:12px;line-height:1.6}.ok{color:#15803d}.warn{color:#b45309}.bad{color:#b91c1c}
.foot{margin-top:22px;font-size:12px;color:#9ca3af}
</style></head><body><header><b>📡 방송 자료 조회</b><span id="srv"></span></header><main>
<div class="cards">
<a class="card" href="radio/"><div class="ic">💰</div><div class="t">라디오 운행표</div><div class="d">달력 · 일/월/년 정산 · 검색 · 운행표 미리보기</div><div class="st" id="sr">확인 중…</div></a>
<a class="card" href="tv/"><div class="ic">📺</div><div class="t">TV 주간 편성표</div><div class="d">주·판별 편성표 PDF 보기 · 인쇄</div><div class="st" id="st">확인 중…</div></a>
<a class="card" href="rec/player.html"><div class="ic">🎧</div><div class="t">녹음 듣기</div><div class="d">정각 단위 녹음 파일 재생 · 다운로드</div><div class="st ok">항상 이용 가능</div></a>
</div>
<div class="foot">편집은 각 편집기 PC에서 합니다. 이 페이지는 편집기에서 저장한 내용을 1분마다 모아 둔 사본이라, 편집기 PC가 꺼져 있어도 볼 수 있습니다.</div>
</main><script>
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function line(x,on){let h='';
 if(x.synced)h+='<span class="ok">✔ 마지막으로 모은 시각 '+esc(x.synced.slice(5,16))+'</span><br>';else h+='<span class="warn">아직 모은 자료가 없습니다</span><br>';
 h+=on.length?'<span class="ok">● 편집기 연결됨: '+esc(on.join(', '))+'</span>':'<span class="warn">○ 편집기 PC 꺼짐 — 마지막 사본을 보여 줍니다</span>';
 if(x.error&&on.length)h+='<br><span class="bad">⚠ '+esc(x.error)+'</span>';return h}
async function load(){try{const d=await (await fetch('api/status')).json();
 document.getElementById('srv').textContent=d.server.name+' · '+d.server.ip;
 let r=line(d.radio,d.radio.online);if(d.radio.records!=null)r+='<br>입력 기록 '+Number(d.radio.records).toLocaleString('ko-KR')+'건';
 document.getElementById('sr').innerHTML=r;
 let t=line(d.tv,d.tv.online);if(d.tv.count!=null)t+='<br>편성표 '+d.tv.count+'개';document.getElementById('st').innerHTML=t;
}catch(e){document.getElementById('sr').textContent=document.getElementById('st').textContent='⚠ 서버 상태를 읽지 못했습니다'}}
load();setInterval(()=>{if(!document.hidden)load()},30000);
</script></body></html>"""

TV_HTML = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TV 주간 편성표 (조회 전용)</title><style>
body{font-family:"Malgun Gothic","맑은 고딕","Noto Sans KR",sans-serif;margin:0;background:#f5f6f8;color:#1f2937;font-size:14px}
header{background:#7c2d12;color:#fff;padding:10px 18px}header b{font-size:17px}header a{color:#fed7aa;font-size:13px;margin-right:12px;text-decoration:none}
main{padding:14px 18px;max-width:1100px}table{border-collapse:collapse;width:100%;background:#fff}
th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left}th{background:#fff7ed}.msg{color:#6b7280;font-size:12px;margin:6px 0}
</style></head><body><header><a href="../">← 처음</a><b>📺 TV 주간 편성표</b> <span id="who" style="font-size:12px;opacity:.85"></span></header>
<main><div class="msg" id="msg">불러오는 중…</div><table id="tbl"></table></main><script>
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
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
