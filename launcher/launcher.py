"""
편집 프로그램 런처 — RADIO / TV 프로그램 중 필요한 것만 골라 설치하고, 업데이트하고, 실행한다.

폴더 (C:\\EDIT)
  런처\\EditLauncher.exe    런처 + 공용 실행 환경(파이썬·PyMuPDF·Pillow·openpyxl)  ← RADIO·TV 가 같이 씀
  업데이트\\               기본 업데이트 폴더 (catalog.json + 꾸러미 zip). 설정에서 공유 폴더·GitHub 주소로 바꿀 수 있다
  launcher.json             런처 설정
  radio\\앱\\<id>\\        RADIO 프로그램 (.py + app.json)
  radio\\데이터\\  radio\\백업\\
  TV\\앱\\  TV\\데이터\\  TV\\백업\\

실행 모드
  EditLauncher.exe                     런처 화면
  EditLauncher.exe --run radio:runsheet  설치한 프로그램 실행 (런처가 이 방법으로 별도 창을 띄운다)
  EditLauncher.exe --selftest 결과.txt   빌드 점검
"""
import datetime
import glob
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import zipfile

LAUNCHER_VERSION = "1.0.0"
RUNTIME = 1                         # 공용 실행 환경 세대. 필요한 부품이 바뀌면 올린다 (앱의 runtime 값과 비교)
EXE_NAME = "EditLauncher.exe"
MUTEX = "EditLauncher_Running"      # 설치 프로그램(setup)이 '실행 중'을 알아보는 이름

GROUPS = {"radio": "RADIO", "tv": "TV"}              # 그룹 id: 탭 이름
GROUP_FOLDER = {"radio": "radio", "tv": "TV"}        # C:\\EDIT 아래 폴더 이름

APP_INFO = {   # 기본 이름·설명·그룹 (catalog.json 에 있으면 그쪽을 쓴다)
    "runsheet": {"name": "운행표 데이터 편집기", "desc": "운행표 PDF에 광고주 데이터 입력 · 정산 · 일정표", "group": "radio"},
    "schedule": {"name": "편성표 PDF 편집기", "desc": "편성표 PDF 편집 (칸 색 · 글자 · 분할선)", "group": "radio"},
}


def _root():
    if os.environ.get("EDIT_HOME"):
        return os.environ["EDIT_HOME"]
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(base) if os.path.basename(base) in ("런처", "launcher") else base
    return r"C:\EDIT" if os.name == "nt" else os.path.join(os.path.expanduser("~"), "EDIT")


ROOT = _root()
LAUNCHER_DIR = os.path.join(ROOT, "런처")
UPDATE_DIR_DEFAULT = os.path.join(ROOT, "업데이트")
SETTINGS_FILE = os.path.join(ROOT, "launcher.json")
OLD_RADIO_ROOT = r"C:\RADIO"                      # 예전(오늘 이전) 설치 위치 -> 기록 옮기기
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def group_root(group):
    return os.path.join(ROOT, GROUP_FOLDER.get(group, group))


def apps_dir(group):
    return os.path.join(group_root(group), "앱")


def data_dir(group):
    return os.path.join(group_root(group), "데이터")


def backup_dir(group):
    return os.path.join(group_root(group), "백업")


def group_of(app_id, cat=None):
    ent = ((cat or {}).get("apps") or {}).get(app_id) or {}
    return ent.get("group") or APP_INFO.get(app_id, {}).get("group") or "radio"


def _bundle_hint():
    """호출하지 않는다. PyInstaller 가 이 import 들을 보고 앱에 필요한 부품을 런처에 함께 넣는다."""
    import base64, calendar, colorsys, copy, csv, ctypes, io, threading, time  # noqa
    import pymupdf  # noqa
    import fitz  # noqa
    import openpyxl  # noqa
    import openpyxl.styles  # noqa
    from PIL import Image, ImageDraw, ImageFont, ImageTk  # noqa
    from tkinter import colorchooser, simpledialog  # noqa
    import winreg  # noqa


def version_tuple(v):
    out = []
    for part in re.split(r"[.\-]", str(v or "0")):
        m = re.match(r"\d+", part)
        out.append(int(m.group()) if m else 0)
    return tuple(out + [0] * (4 - len(out)))[:4]


def hold_mutex():
    try:
        import ctypes
        return ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX)
    except Exception:
        return None


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            d = json.load(f) or {}
    except Exception:
        d = {}
    d.setdefault("update_dir", UPDATE_DIR_DEFAULT)
    return d


def save_settings(d):
    try:
        os.makedirs(ROOT, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def installed_info(app_id, group=None):
    for g in ([group] if group else list(GROUPS)):
        folder = os.path.join(apps_dir(g), app_id)
        try:
            with open(os.path.join(folder, "app.json"), "r", encoding="utf-8") as f:
                info = json.load(f)
            if os.path.exists(os.path.join(folder, info["entry"])):
                info["group"] = g
                info["_dir"] = folder
                return info
        except Exception:
            pass
    return None


GITHUB_RE = re.compile(r"github\.com[/:]+([^/\s]+)/([^/\s#?]+?)(?:\.git)?/?$", re.I)


def is_github(src):
    return bool(src and GITHUB_RE.search(src.strip()))


def _http_get(url, accept="application/json", timeout=15):
    import urllib.request
    req = urllib.request.Request(url, headers={"Accept": accept, "User-Agent": f"EditLauncher/{LAUNCHER_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def read_catalog_github(src):
    """공개 GitHub 저장소의 최신 Release 에서 catalog.json 만 받아 온다 (작은 파일).
    꾸러미 zip 은 설치·업데이트할 때 필요한 것만 받는다 (ensure_local)."""
    m = GITHUB_RE.search(src.strip())
    owner, repo = m.group(1), m.group(2)
    rel = json.loads(_http_get(f"https://api.github.com/repos/{owner}/{repo}/releases/latest").decode("utf-8"))
    assets = {a["name"]: a["browser_download_url"] for a in rel.get("assets", [])}
    if "catalog.json" not in assets:
        return None
    cache = os.path.join(UPDATE_DIR_DEFAULT, "github")
    os.makedirs(cache, exist_ok=True)
    data = _http_get(assets["catalog.json"], accept="application/octet-stream")
    with open(os.path.join(cache, "catalog.json"), "wb") as f:
        f.write(data)
    cat = json.loads(data.decode("utf-8-sig"))
    cat["_dir"] = cache
    cat["_remote"] = assets
    cat.setdefault("apps", {})
    return cat


def ensure_local(cat, fname):
    """꾸러미 파일 경로. GitHub 에서 읽은 목록이면 필요할 때 받아 온다."""
    path = os.path.join(cat["_dir"], os.path.basename(fname))
    remote = cat.get("_remote") or {}
    if not os.path.exists(path) and fname in remote:
        tmp = path + ".part"
        import urllib.request
        req = urllib.request.Request(remote[fname], headers={"Accept": "application/octet-stream", "User-Agent": f"EditLauncher/{LAUNCHER_VERSION}"})
        with urllib.request.urlopen(req, timeout=60) as r, open(tmp, "wb") as f:
            shutil.copyfileobj(r, f, 1024 * 256)
        os.replace(tmp, path)
    return path


def read_catalog(folder):
    """업데이트 폴더(또는 공개 GitHub 저장소 주소)의 catalog.json. 없거나 못 읽으면 None."""
    if not folder:
        return None
    if is_github(folder):
        try:
            return read_catalog_github(folder)
        except Exception:
            return None
    try:
        with open(os.path.join(folder, "catalog.json"), "r", encoding="utf-8-sig") as f:
            cat = json.load(f)
    except Exception:
        return None
    if not isinstance(cat, dict):
        return None
    cat["_dir"] = folder
    cat.setdefault("apps", {})
    return cat


def count_running():
    """EditLauncher.exe 프로세스 수 (런처 화면 + 실행 중인 프로그램들)."""
    if os.name != "nt":
        return 1
    try:
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {EXE_NAME}", "/NH"], capture_output=True,
                             creationflags=NO_WINDOW).stdout.decode("mbcs", errors="ignore")
        return sum(1 for ln in out.splitlines() if EXE_NAME.lower() in ln.lower())
    except Exception:
        return 1


def backup_data(group="radio", tag="업데이트전", keep=10):
    """그 그룹의 데이터 폴더를 통째로 zip 으로 백업한다. DB 는 SQLite 백업 기능으로 안전하게 복사해서 넣는다."""
    DATA_DIR, BACKUP_DIR = data_dir(group), backup_dir(group)
    if not os.path.isdir(DATA_DIR) or not os.listdir(DATA_DIR):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(BACKUP_DIR, f"{tag}_{stamp}.zip")
    tmpdir = tempfile.mkdtemp(prefix="radio_bk_")
    try:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for base, _dirs, files in os.walk(DATA_DIR):
                for fn in files:
                    full = os.path.join(base, fn)
                    arc = os.path.relpath(full, DATA_DIR)
                    if fn.lower().endswith(".db"):
                        tmp = os.path.join(tmpdir, "x.db")
                        try:
                            s, d = sqlite3.connect(full), sqlite3.connect(tmp)
                            try:
                                s.backup(d)
                            finally:
                                d.close()
                                s.close()
                            z.write(tmp, arc)
                            os.remove(tmp)
                            continue
                        except Exception:
                            pass
                    try:
                        z.write(full, arc)
                    except OSError:
                        pass
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    olds = sorted(glob.glob(os.path.join(BACKUP_DIR, f"{tag}_*.zip")))
    for old in olds[:-keep]:
        try:
            os.remove(old)
        except OSError:
            pass
    return path


def install_app(app_id, cat):
    """catalog 의 꾸러미로 앱을 설치/업데이트한다. 데이터는 건드리지 않는다. 성공하면 새 app.json 내용."""
    ent = cat["apps"][app_id]
    zpath = ensure_local(cat, ent["file"])
    if not os.path.exists(zpath):
        raise FileNotFoundError(f"꾸러미 파일이 없습니다:\n{zpath}")
    work = tempfile.mkdtemp(prefix=f"radio_{app_id}_")
    try:
        with zipfile.ZipFile(zpath) as z:
            z.extractall(work)
        if not os.path.exists(os.path.join(work, ent["entry"])):
            raise RuntimeError(f"꾸러미 안에 {ent['entry']} 가 없습니다.")
        group = group_of(app_id, cat)
        info = {"id": app_id, "name": ent.get("name", app_id), "version": ent["version"], "entry": ent["entry"], "group": group,
                "runtime": int(ent.get("runtime", 1)), "installed_at": datetime.datetime.now().isoformat(timespec="seconds")}
        with open(os.path.join(work, "app.json"), "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False, indent=1)
        os.makedirs(apps_dir(group), exist_ok=True)
        for d in (data_dir(group), backup_dir(group)):
            os.makedirs(d, exist_ok=True)
        dst = os.path.join(apps_dir(group), app_id)
        old = dst + ".old"
        shutil.rmtree(old, ignore_errors=True)
        if os.path.exists(dst):
            os.replace(dst, old)            # .py 는 실행 중에도 잠기지 않아 바로 바꿀 수 있다
        shutil.copytree(work, dst)
        shutil.rmtree(old, ignore_errors=True)
        return info
    finally:
        shutil.rmtree(work, ignore_errors=True)


def run_app_mode(spec):
    """--run 모드 (spec = '그룹:앱' 또는 '앱'): 설치한 앱을 이 프로세스에서 실행한다 (공용 실행 환경 사용)."""
    hold_mutex()
    group, _, app_id = spec.rpartition(":")
    info = installed_info(app_id, group or None)
    if info is None:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("편집 프로그램", f"'{APP_INFO.get(app_id, {}).get('name', app_id)}' 이(가) 설치되어 있지 않습니다.\n편집 프로그램(런처)에서 설치해 주세요.")
        return 1
    if int(info.get("runtime", 1)) > RUNTIME:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("편집 프로그램", "이 프로그램은 새 런처가 필요합니다.\n편집 프로그램(런처)에서 [런처 업데이트]를 먼저 해 주세요.")
        return 1
    folder = info["_dir"]
    script = os.path.join(folder, info["entry"])
    os.environ["RADIO_HOME"] = group_root(info["group"])     # 앱은 이 폴더 아래 데이터·백업을 쓴다 (C:\\EDIT\\radio 등)
    os.environ["EDIT_HOME"] = ROOT
    os.environ["RADIO_LAUNCHER"] = "1"
    os.environ["RADIO_APP_ID"] = app_id
    sys.argv = [script]
    sys.path.insert(0, folder)
    os.chdir(folder)
    import runpy
    runpy.run_path(script, run_name="__main__")
    return 0


def start_app_process(app_id):
    info = installed_info(app_id)
    spec = f"{info['group']}:{app_id}"
    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--run", spec]
    else:
        cmd = [sys.executable, os.path.abspath(__file__), "--run", spec]
    return subprocess.Popen(cmd, cwd=info["_dir"], close_fds=True)


def migrate_old_radio(ask):
    """예전 C:\\RADIO 의 기록·설정을 C:\\EDIT\\radio 로 옮긴다 (복사, 원본은 그대로). 옮겼으면 안내 문장."""
    old_data = os.path.join(OLD_RADIO_ROOT, "데이터")
    new_data = data_dir("radio")
    if os.path.abspath(old_data) == os.path.abspath(new_data) or not os.path.isdir(old_data):
        return ""
    if os.path.exists(os.path.join(new_data, "pdf_data_editor.db")) or os.path.exists(os.path.join(ROOT, ".radio_migrated")):
        return ""
    if not any(os.scandir(old_data)):
        return ""
    if not ask("예전 기록 옮기기", f"예전 위치에 RADIO 기록이 있습니다.\n{old_data}\n\n새 위치로 복사해서 이어서 쓸까요?\n{new_data}\n\n(원본은 지우지 않습니다)"):
        open(os.path.join(ROOT, ".radio_migrated"), "w").close()      # 다시 묻지 않음
        return ""
    os.makedirs(new_data, exist_ok=True)
    for base, _d, files in os.walk(old_data):
        for fn in files:
            src = os.path.join(base, fn)
            rel = os.path.relpath(src, old_data)
            dst = os.path.join(new_data, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            if fn.lower().endswith(".db"):
                s_, d_ = sqlite3.connect(src), sqlite3.connect(dst)
                try:
                    s_.backup(d_)
                finally:
                    d_.close()
                    s_.close()
            else:
                shutil.copy2(src, dst)
    # settings.json 안의 예전 경로(C:\\RADIO\\데이터\\...)는 새 경로로
    sp = os.path.join(new_data, "settings.json")
    if os.path.exists(sp):
        try:
            with open(sp, "r", encoding="utf-8") as f:
                st = json.load(f)
            if str(st.get("db_path", "")).lower().startswith(OLD_RADIO_ROOT.lower()):
                st["db_path"] = os.path.join(new_data, "pdf_data_editor.db")
            if str(st.get("backup_dir", "")).lower().startswith(OLD_RADIO_ROOT.lower()):
                st["backup_dir"] = ""
            if str(st.get("update_dir", "")).lower().startswith(OLD_RADIO_ROOT.lower()):
                st["update_dir"] = ""
            with open(sp, "w", encoding="utf-8") as f:
                json.dump(st, f, ensure_ascii=False, indent=1)
        except Exception:
            pass
    open(os.path.join(ROOT, ".radio_migrated"), "w").close()
    return f"예전 기록을 옮겼습니다:\n{new_data}"


# ============================================================================ 화면
class LauncherApp:
    def __init__(self, root):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk = tk, ttk
        self.root = root
        root.title(f"편집 프로그램  (런처 v{LAUNCHER_VERSION})")
        root.geometry("780x520")
        root.minsize(640, 400)
        self.settings = load_settings()
        self.cat = None
        self.font = "맑은 고딕" if os.name == "nt" else "TkDefaultFont"
        for d in [UPDATE_DIR_DEFAULT] + [f(g) for g in GROUPS for f in (apps_dir, data_dir, backup_dir)]:
            try:
                os.makedirs(d, exist_ok=True)
            except OSError:
                pass
        self._build()
        self.refresh()
        root.after(300, self._first_run_hint)

    # ------------------------------------------------------------------ 화면 뼈대
    def _build(self):
        tk, ttk = self.tk, self.ttk
        top = tk.Frame(self.root, background="#1e3a8a")
        top.pack(fill="x")
        tk.Label(top, text="  편집 프로그램", font=(self.font, 15, "bold"), fg="white", bg="#1e3a8a", pady=10).pack(side="left")
        self.banner = tk.Label(top, text="", font=(self.font, 9, "bold"), fg="#fde68a", bg="#1e3a8a", cursor="hand2")
        self.banner.pack(side="right", padx=12)
        self.banner.bind("<Button-1>", lambda e: self.update_launcher())

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(8, 0))
        self.tabs = {}
        for g, title in GROUPS.items():
            fr = ttk.Frame(self.nb, padding=(10, 8))
            self.nb.add(fr, text=f"   {title}   ")
            self.tabs[g] = fr
        try:
            last = self.settings.get("last_tab", "radio")
            self.nb.select(list(GROUPS).index(last))
        except Exception:
            pass
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._remember_tab())

        bot = ttk.Frame(self.root, padding=(14, 6))
        bot.pack(fill="x")
        ttk.Button(bot, text="🔄 다시 확인", command=self.refresh).pack(side="left")
        ttk.Button(bot, text="⬆ 모두 업데이트", command=self.update_all).pack(side="left", padx=6)
        ttk.Button(bot, text="📦 업데이트 파일 가져오기 (USB 등)", command=self.import_package).pack(side="left")
        ttk.Button(bot, text="📁 폴더 열기", command=lambda: self._open_folder(group_root(self._cur_group()))).pack(side="right")

        src = ttk.Frame(self.root, padding=(14, 0, 14, 10))
        src.pack(fill="x")
        ttk.Label(src, text="업데이트 위치 (폴더 또는 GitHub 주소):").pack(side="left")
        self.src_var = tk.StringVar(value=self.settings["update_dir"])
        ttk.Entry(src, textvariable=self.src_var).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(src, text="폴더 선택", command=self._pick_src).pack(side="left")
        ttk.Button(src, text="저장", command=self._save_src).pack(side="left", padx=(4, 0))
        self.status = ttk.Label(self.root, text="", foreground="#555", padding=(14, 0, 14, 8))
        self.status.pack(fill="x")

    def _cur_group(self):
        try:
            return list(GROUPS)[self.nb.index(self.nb.select())]
        except Exception:
            return "radio"

    def _remember_tab(self):
        self.settings["last_tab"] = self._cur_group()
        save_settings(self.settings)

    def _pick_src(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(title="업데이트 폴더 (catalog.json 이 있는 곳, 예: 공유 폴더)", parent=self.root)
        if d:
            self.src_var.set(os.path.normpath(d))
            self._save_src()

    def _save_src(self):
        v = self.src_var.get().strip().strip('"')
        self.settings["update_dir"] = v if is_github(v) else (os.path.normpath(v) if v else UPDATE_DIR_DEFAULT)
        save_settings(self.settings)
        self.refresh()

    def _open_folder(self, path):
        try:
            os.makedirs(path, exist_ok=True)
            os.startfile(path)
        except Exception:
            pass

    # ------------------------------------------------------------------ 목록
    def refresh(self):
        tk, ttk = self.tk, self.ttk
        self.cat = read_catalog(self.settings["update_dir"])
        if self.cat is None and (is_github(self.settings["update_dir"]) or os.path.abspath(self.settings["update_dir"]) != os.path.abspath(UPDATE_DIR_DEFAULT)):
            self.cat = read_catalog(UPDATE_DIR_DEFAULT)      # 공유 폴더에 못 닿으면 PC 안의 업데이트 폴더라도
            src_note = f"업데이트 위치에 닿지 않아 이 PC의 업데이트 폴더를 봅니다: {UPDATE_DIR_DEFAULT}"
        else:
            src_note = ""
        for fr in self.tabs.values():
            for w in fr.winfo_children():
                w.destroy()
        ids = list(APP_INFO) + [k for k in (self.cat or {}).get("apps", {}) if k not in APP_INFO]
        shown = {g: 0 for g in GROUPS}
        for app_id in ids:
            g = group_of(app_id, self.cat)
            if g not in self.tabs:
                continue
            self._row(app_id, self.tabs[g])
            shown[g] += 1
        for g, n in shown.items():
            if not n:
                ttk.Label(self.tabs[g], text=f"아직 {GROUPS[g]} 프로그램이 없습니다.\n새 프로그램이 업데이트 꾸러미에 들어오면 여기에 보입니다.",
                          foreground="#64748b", justify="left").pack(anchor="w", pady=20)
        # 런처 자체 업데이트 안내
        lv = ((self.cat or {}).get("launcher") or {}).get("version")
        if lv and version_tuple(lv) > version_tuple(LAUNCHER_VERSION):
            self.banner.config(text=f"⬆ 새 런처 v{lv} — 눌러서 업데이트  ")
        else:
            self.banner.config(text="")
        if self.cat is None:
            self.status.config(text=f"업데이트 정보(catalog.json)를 찾지 못했습니다: {self.settings['update_dir']}   — 설치·업데이트를 하려면 위치를 확인하거나 [업데이트 파일 가져오기]를 쓰세요.",
                               foreground="#b45309", wraplength=720)
        else:
            built = self.cat.get("built", "")
            self.status.config(text=(src_note or f"업데이트 위치 확인됨" + (f"  (꾸러미 만든 날 {built})" if built else "")), foreground="#555", wraplength=720)

    def _row(self, app_id, parent):
        tk, ttk = self.tk, self.ttk
        meta = dict(APP_INFO.get(app_id, {}))
        ent = (self.cat or {}).get("apps", {}).get(app_id)
        if ent:
            meta.update({k: ent[k] for k in ("name", "desc") if ent.get(k)})
        inst = installed_info(app_id)
        f = tk.Frame(parent, background="#ffffff", highlightthickness=1, highlightbackground="#cbd5e1", padx=12, pady=10)
        f.pack(fill="x", pady=5)
        left = tk.Frame(f, background="#ffffff")
        left.pack(side="left", fill="x", expand=True)
        tk.Label(left, text=meta.get("name", app_id), font=(self.font, 12, "bold"), bg="#ffffff").pack(anchor="w")
        tk.Label(left, text=meta.get("desc", ""), font=(self.font, 9), fg="#64748b", bg="#ffffff").pack(anchor="w")
        newer = bool(ent and inst and version_tuple(ent["version"]) > version_tuple(inst["version"]))
        if inst:
            st = f"설치됨 v{inst['version']}" + (f"   →   새 버전 v{ent['version']}" if newer else ("   (최신)" if ent else ""))
            col = "#c2410c" if newer else "#15803d"
        else:
            st = "설치 안 함" + (f"   (설치 가능 v{ent['version']})" if ent else "")
            col = "#64748b"
        tk.Label(left, text=st, font=(self.font, 9, "bold"), fg=col, bg="#ffffff").pack(anchor="w", pady=(4, 0))
        if newer and ent.get("notes"):
            tk.Label(left, text="바뀐 내용: " + ent["notes"], font=(self.font, 8), fg="#92400e", bg="#ffffff", wraplength=420, justify="left").pack(anchor="w")
        btns = tk.Frame(f, background="#ffffff")
        btns.pack(side="right")
        if inst:
            ttk.Button(btns, text="▶ 실행", width=10, command=lambda: self.run(app_id)).pack(side="left", padx=2)
            if newer:
                ttk.Button(btns, text="⬆ 업데이트", width=10, command=lambda: self.install(app_id, update=True)).pack(side="left", padx=2)
            ttk.Button(btns, text="제거", width=6, command=lambda: self.remove(app_id)).pack(side="left", padx=2)
        elif ent:
            ttk.Button(btns, text="⬇ 설치", width=10, command=lambda: self.install(app_id)).pack(side="left", padx=2)

    def _first_run_hint(self):
        """운행표 편집기를 쓰던 PC(데이터가 있음)인데 아직 설치하지 않았으면 설치를 권한다."""
        from tkinter import messagebox
        try:
            note = migrate_old_radio(lambda t, m: messagebox.askyesno(t, m, parent=self.root))
        except Exception as exc:
            note = f"예전 기록을 옮기지 못했습니다: {exc}"
        if note:
            messagebox.showinfo("편집 프로그램", note, parent=self.root)
        had_runsheet = os.path.exists(os.path.join(data_dir("radio"), "pdf_data_editor.db"))
        if had_runsheet and not installed_info("runsheet") and self.cat and "runsheet" in self.cat["apps"]:
            if messagebox.askyesno("편집 프로그램", "이 PC에는 운행표 데이터 편집기의 기록이 있습니다.\n운행표 데이터 편집기를 지금 설치할까요?\n(기록은 그대로 이어서 씁니다)", parent=self.root):
                self.install("runsheet")

    # ------------------------------------------------------------------ 동작
    def run(self, app_id):
        from tkinter import messagebox
        try:
            start_app_process(app_id)
            self.status.config(text=f"{APP_INFO.get(app_id, {}).get('name', app_id)} 을(를) 실행했습니다.", foreground="#1d4ed8")
        except Exception as exc:
            messagebox.showerror("실행 실패", str(exc), parent=self.root)

    def install(self, app_id, update=False, quiet=False):
        from tkinter import messagebox
        if not self.cat or app_id not in self.cat["apps"]:
            messagebox.showwarning("편집 프로그램", "업데이트 정보에 이 프로그램이 없습니다.", parent=self.root)
            return False
        ent = self.cat["apps"][app_id]
        if int(ent.get("runtime", 1)) > RUNTIME:
            messagebox.showinfo("편집 프로그램", "이 버전은 새 런처가 필요합니다.\n먼저 런처를 업데이트합니다.", parent=self.root)
            self.update_launcher()
            return False
        name = ent.get("name") or APP_INFO.get(app_id, {}).get("name", app_id)
        if update:
            try:
                backup_data(group_of(app_id, self.cat))
            except Exception as exc:
                if not messagebox.askyesno("백업 실패", f"데이터를 백업하지 못했습니다:\n{exc}\n\n그래도 업데이트할까요?", parent=self.root):
                    return False
        try:
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            info = install_app(app_id, self.cat)
        except Exception as exc:
            messagebox.showerror("설치 실패", f"{name}\n\n{exc}", parent=self.root)
            return False
        finally:
            self.root.config(cursor="")
        if not quiet:
            msg = f"{name} v{info['version']} 을(를) {'업데이트' if update else '설치'}했습니다."
            if update and count_running() > 1:
                msg += "\n\n지금 켜져 있는 프로그램은 닫았다가 다시 켜면 새 버전이 됩니다."
            messagebox.showinfo("편집 프로그램", msg, parent=self.root)
        self.refresh()
        return True

    def remove(self, app_id):
        from tkinter import messagebox
        name = APP_INFO.get(app_id, {}).get("name", app_id)
        if not messagebox.askyesno("제거", f"{name} 을(를) 이 PC에서 제거할까요?\n(입력 기록·설정 데이터는 지우지 않습니다)", parent=self.root):
            return
        info = installed_info(app_id)
        if info:
            shutil.rmtree(info["_dir"], ignore_errors=True)
        self.refresh()

    def update_all(self):
        from tkinter import messagebox
        self.refresh()
        if not self.cat:
            messagebox.showinfo("모두 업데이트", "업데이트 정보를 찾지 못했습니다.", parent=self.root)
            return
        todo = [a for a, ent in self.cat["apps"].items()
                if installed_info(a) and version_tuple(ent["version"]) > version_tuple(installed_info(a)["version"])]
        lv = (self.cat.get("launcher") or {}).get("version")
        need_launcher = bool(lv and version_tuple(lv) > version_tuple(LAUNCHER_VERSION))
        if not todo and not need_launcher:
            messagebox.showinfo("모두 업데이트", "모두 최신입니다.", parent=self.root)
            return
        for g in GROUPS:
            try:
                backup_data(g)
            except Exception:
                pass
        done = []
        for a in todo:
            if int(self.cat["apps"][a].get("runtime", 1)) > RUNTIME:
                continue                     # 런처를 먼저 바꿔야 하는 앱은 런처 업데이트 뒤 다시 켜서
            if self.install(a, quiet=True):
                done.append(self.cat["apps"][a].get("name", a))
        if done:
            messagebox.showinfo("모두 업데이트", "업데이트했습니다:\n" + "\n".join(done), parent=self.root)
        if need_launcher:
            self.update_launcher()

    def import_package(self):
        """USB 등으로 가져온 업데이트 꾸러미(zip 또는 catalog.json 이 있는 폴더)를 이 PC의 업데이트 폴더로 복사한다."""
        from tkinter import filedialog, messagebox
        path = filedialog.askopenfilename(title="업데이트 꾸러미 선택 (EDIT_update_*.zip 또는 catalog.json)",
                                          filetypes=[("업데이트 꾸러미", "*.zip;catalog.json"), ("모든 파일", "*.*")], parent=self.root)
        if not path:
            return
        dst = UPDATE_DIR_DEFAULT
        try:
            os.makedirs(dst, exist_ok=True)
            if path.lower().endswith(".zip"):
                with zipfile.ZipFile(path) as z:
                    names = z.namelist()
                    cat_name = next((n for n in names if n.split("/")[-1] == "catalog.json"), None)
                    if cat_name is None:
                        raise RuntimeError("이 zip 에는 catalog.json 이 없습니다. EDIT_update 꾸러미가 맞는지 확인해 주세요.")
                    prefix = cat_name[: -len("catalog.json")]
                    for n in names:
                        if n.startswith(prefix) and not n.endswith("/"):
                            target = os.path.join(dst, n[len(prefix):])
                            os.makedirs(os.path.dirname(target), exist_ok=True)
                            with z.open(n) as src_f, open(target, "wb") as out:
                                shutil.copyfileobj(src_f, out)
            else:
                folder = os.path.dirname(path)
                if os.path.abspath(folder) != os.path.abspath(dst):
                    for fn in os.listdir(folder):
                        if fn == "catalog.json" or fn.lower().endswith(".zip"):
                            shutil.copy2(os.path.join(folder, fn), os.path.join(dst, fn))
        except Exception as exc:
            messagebox.showerror("가져오기 실패", str(exc), parent=self.root)
            return
        if os.path.abspath(self.settings["update_dir"]) != os.path.abspath(dst):
            if messagebox.askyesno("가져오기", f"가져왔습니다: {dst}\n\n업데이트 위치를 이 PC의 업데이트 폴더로 바꿀까요?\n(아니오: 이번에만 이 폴더를 봅니다)", parent=self.root):
                self.settings["update_dir"] = dst
                self.src_var.set(dst)
                save_settings(self.settings)
                self.refresh()
            else:
                self.cat = read_catalog(dst)
                saved = self.settings["update_dir"]
                self.settings["update_dir"] = dst
                self.refresh()
                self.settings["update_dir"] = saved
        else:
            self.refresh()
            messagebox.showinfo("가져오기", "가져왔습니다. 목록에서 설치·업데이트하세요.", parent=self.root)

    def update_launcher(self):
        """런처(공용 실행 환경 포함) 업데이트. 켜진 프로그램을 모두 닫아야 한다."""
        from tkinter import messagebox
        lent = (self.cat or {}).get("launcher") or {}
        if not lent.get("file"):
            messagebox.showinfo("런처 업데이트", "새 런처 정보가 없습니다.", parent=self.root)
            return
        try:
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            zpath = ensure_local(self.cat, lent["file"])
        except Exception as exc:
            messagebox.showerror("런처 업데이트", f"런처 꾸러미를 받지 못했습니다:\n{exc}", parent=self.root)
            return
        finally:
            self.root.config(cursor="")
        if not os.path.exists(zpath):
            messagebox.showerror("런처 업데이트", f"런처 꾸러미가 없습니다:\n{zpath}", parent=self.root)
            return
        if not getattr(sys, "frozen", False):
            messagebox.showinfo("런처 업데이트", f"개발 모드(.py)에서는 런처를 바꾸지 않습니다.\n{zpath}", parent=self.root)
            return
        while count_running() > 1:
            if not messagebox.askretrycancel("런처 업데이트", "켜져 있는 RADIO·TV 프로그램을 모두 저장하고 닫은 뒤 [다시 시도]를 눌러 주세요.", parent=self.root):
                return
        if not messagebox.askyesno("런처 업데이트", f"런처를 v{lent['version']} 로 업데이트합니다.\n런처가 잠시 닫혔다가 다시 켜집니다.\n(데이터는 그대로입니다)\n\n계속할까요?", parent=self.root):
            return
        try:
            for g in GROUPS:
                backup_data(g)
            work = tempfile.mkdtemp(prefix="edit_launcher_")
            with zipfile.ZipFile(zpath) as z:
                z.extractall(work)
            new_dir = os.path.join(work, "EditLauncher")
            if not os.path.exists(os.path.join(new_dir, EXE_NAME)):
                raise RuntimeError("꾸러미 안에 EditLauncher 폴더가 없습니다.")
            cur_dir = os.path.dirname(os.path.abspath(sys.executable))
            ps1 = os.path.join(work, "update.ps1")
            qn, qc = new_dir.replace("'", "''"), cur_dir.replace("'", "''")
            with open(ps1, "w", encoding="utf-8-sig") as f:
                f.write(f"""$new = '{qn}'
$cur = '{qc}'
for ($i = 0; $i -lt 60; $i++) {{
    if (-not (Get-Process EditLauncher -ErrorAction SilentlyContinue)) {{ break }}
    Start-Sleep -Seconds 1
}}
robocopy $new $cur /MIR /R:20 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
Start-Process (Join-Path $cur '{EXE_NAME}') -WorkingDirectory $cur
""")
            subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", ps1],
                             creationflags=NO_WINDOW, close_fds=True)
        except Exception as exc:
            messagebox.showerror("런처 업데이트 실패", str(exc), parent=self.root)
            return
        self.root.destroy()


def selftest(out_path):
    """빌드 점검용: 앱들이 쓰는 부품이 실행 환경에 다 들어 있는지 실제로 불러 본다."""
    mods = ["tkinter", "tkinter.ttk", "tkinter.colorchooser", "tkinter.simpledialog", "tkinter.filedialog", "sqlite3", "csv",
            "calendar", "colorsys", "ctypes", "winreg", "pymupdf", "fitz", "openpyxl", "openpyxl.styles",
            "PIL.Image", "PIL.ImageDraw", "PIL.ImageFont", "PIL.ImageTk"]
    import importlib
    bad = []
    for m in mods:
        try:
            importlib.import_module(m)
        except Exception as exc:
            bad.append(f"{m}: {exc}")
    try:
        import pymupdf
        d = pymupdf.open()
        d.new_page().insert_text((50, 50), "test")
        d.tobytes()
    except Exception as exc:
        bad.append(f"pymupdf 동작: {exc}")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("OK" if not bad else "\n".join(bad))
    return 0 if not bad else 1


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2]))
    if len(sys.argv) >= 3 and sys.argv[1] == "--run":
        sys.exit(run_app_mode(sys.argv[2]))
    import tkinter as tk
    hold_mutex()
    root = tk.Tk()
    LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
