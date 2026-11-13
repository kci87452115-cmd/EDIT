import os
import re
import csv
import json
import time
import datetime
import calendar
import copy
import base64
import sqlite3
import glob
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    import winreg
except ImportError:
    winreg = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

APP_VERSION = "2.12.0"
APP_TITLE = f"운행표 데이터 편집기 v{APP_VERSION}"
APP_MUTEX = "RunSheetEditor_Running"      # 설치 프로그램이 '실행 중'을 알아보는 이름 (setup.iss 의 AppMutex 와 같아야 함)
RUN_BY_LAUNCHER = bool(os.environ.get("RADIO_LAUNCHER"))   # 편집 프로그램 런처로 실행됨 -> 업데이트는 런처가 한다

# ---- 폴더 구조 ---------------------------------------------------------------
#   C:\RADIO\프로그램\RunSheetEditor.exe   ← 업데이트 때 이 폴더만 바뀐다
#   C:\RADIO\데이터\pdf_data_editor.db, settings.json
#   C:\RADIO\백업\                        ← 날짜별 자동 백업
#   C:\RADIO\업데이트\                     ← (기본) 새 버전 설치 파일을 두는 곳. 설정에서 공유 폴더로 바꿀 수 있다
def _app_root():
    import sys
    if os.environ.get("RADIO_HOME"):
        return os.environ["RADIO_HOME"]
    if getattr(sys, "frozen", False):                       # exe 로 실행: ...\RADIO\프로그램\RunSheetEditor.exe
        base = os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(base) if os.path.basename(base) in ("프로그램", "program", "런처") else base
    return r"C:\RADIO" if os.name == "nt" else os.path.join(os.path.expanduser("~"), "RADIO")


APP_ROOT = _app_root()
DATA_DIR = os.path.join(APP_ROOT, "데이터")
BACKUP_DIR = os.path.join(APP_ROOT, "백업")
UPDATE_DIR_DEFAULT = os.path.join(APP_ROOT, "업데이트")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
DB_FILE = os.path.join(DATA_DIR, "pdf_data_editor.db")
SEED_DB_FILE = os.path.join(DATA_DIR, "기본데이터.db")    # 설치 파일이 넣어 두는 기본 데이터 (처음 한 번만 쓰인다)
OLD_SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".pdf_data_editor_settings.json")   # v2.0 까지의 위치
OLD_DB_FILE = os.path.join(os.path.expanduser("~"), "pdf_data_editor.db")


def version_tuple(v):
    out = []
    for part in re.split(r"[.\-]", str(v or "0")):
        m = re.match(r"\d+", part)
        out.append(int(m.group()) if m else 0)
    return tuple(out + [0] * (4 - len(out)))[:4]


def hold_app_mutex():
    """실행 중 표시. 설치 프로그램이 이 이름을 보고 '프로그램을 닫아 주세요'라고 안내한다. (윈도우 전용, 실패해도 무시)"""
    try:
        import ctypes
        return ctypes.windll.kernel32.CreateMutexW(None, False, APP_MUTEX)
    except Exception:
        return None


def copy_db(src, dst):
    """SQLite 백업 기능으로 안전하게 복사한다 (쓰는 중이어도 깨지지 않게)."""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    tmp = dst + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)
    s = sqlite3.connect(src)
    d = sqlite3.connect(tmp)
    try:
        s.backup(d)
    finally:
        d.close()
        s.close()
    os.replace(tmp, dst)


def prepare_data_folder(ask):
    """처음 실행 준비. 폴더를 만들고, 설정·기록을 정한다.
    ask(title, msg) -> bool : 사용자에게 묻는 함수.
    순서: ① 이미 C:\\RADIO\\데이터 에 DB가 있으면 그대로
          ② 예전 위치(사용자 폴더)에 기록이 있으면 옮길지 묻는다
          ③ 설치 파일이 넣은 기본 데이터(기본데이터.db)가 있으면 그것으로 시작
          ④ 아무것도 없으면 빈 DB 로 시작 (Store 가 만든다)
    반환: 안내할 문장 (없으면 '')"""
    for d in (DATA_DIR, BACKUP_DIR):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            pass
    note = ""
    if os.path.exists(SETTINGS_FILE):              # 이미 새 위치에서 쓰고 있다. 다른 DB 위치를 지정해 두었으면 손대지 않는다
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                cur_db = (json.load(f) or {}).get("db_path") or DB_FILE
        except Exception:
            cur_db = DB_FILE
        if os.path.abspath(cur_db) != os.path.abspath(DB_FILE):
            return ""
    old_settings = {}
    if os.path.exists(OLD_SETTINGS_FILE):
        try:
            with open(OLD_SETTINGS_FILE, "r", encoding="utf-8") as f:
                old_settings = json.load(f) or {}
        except Exception:
            old_settings = {}
    if not os.path.exists(SETTINGS_FILE) and old_settings:
        # 예전 설정(글꼴·정렬·열 설정 등)을 이어받는다. DB·백업 위치는 아래에서 다시 정한다
        carry = dict(old_settings)
        old_db = carry.get("db_path") or OLD_DB_FILE
        custom_db = os.path.abspath(old_db) != os.path.abspath(OLD_DB_FILE)   # 공유 폴더 등 직접 지정했던 위치
        if not custom_db:
            carry["db_path"] = DB_FILE
            carry["backup_dir"] = ""
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(carry, f, ensure_ascii=False)
        except Exception:
            pass
        if custom_db:
            return f"예전 설정의 데이터 파일 위치를 그대로 씁니다:\n{old_db}"
    if os.path.exists(DB_FILE):
        return note
    if os.path.exists(OLD_DB_FILE):
        if ask("기존 기록 옮기기", f"예전 위치에 입력 기록과 미리 만든 데이터가 있습니다.\n{OLD_DB_FILE}\n\n"
                                  f"새 위치로 복사해서 계속 쓸까요?\n{DB_FILE}\n\n(원본 파일은 지우지 않고 그대로 둡니다)"):
            try:
                copy_db(OLD_DB_FILE, DB_FILE)
                return f"기존 기록을 옮겼습니다:\n{DB_FILE}"
            except Exception as exc:
                return f"기존 기록을 옮기지 못했습니다: {exc}\n빈 데이터로 시작합니다."
    if os.path.exists(SEED_DB_FILE):
        try:
            copy_db(SEED_DB_FILE, DB_FILE)
            return "기본 데이터(미리 만든 데이터·재원 설정·그룹)로 시작합니다."
        except Exception as exc:
            return f"기본 데이터를 넣지 못했습니다: {exc}"
    return note


def make_seed_db(src, dst, keep_records=False):
    """설치 파일에 넣을 기본 데이터 파일을 만든다. 기본은 입력 기록을 빼고 미리 만든 데이터·재원 설정·그룹만 남긴다."""
    copy_db(src, dst)
    c = sqlite3.connect(dst)
    try:
        if not keep_records:
            c.execute("DELETE FROM records")
        c.commit()
        c.execute("VACUUM")
    finally:
        c.close()
    c = sqlite3.connect(dst)
    try:
        n = {t: c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in ("presets", "cap_settings", "groups", "records")}
    finally:
        c.close()
    return n


def import_base_data(dst_db, src_db):
    """다른 PC에서 만든 기본 데이터 파일(또는 백업 파일)에서 미리 만든 데이터·재원 설정·그룹을 가져와 합친다.
    이미 있는 것(같은 이름+초수, 같은 시각+프로그램)은 그대로 두고 없는 것만 더한다. 입력 기록은 가져오지 않는다.
    반환: {'presets': n, 'cap_settings': n, 'groups': n} 새로 더한 개수"""
    if os.path.abspath(dst_db) == os.path.abspath(src_db):
        raise ValueError("지금 쓰는 데이터 파일과 같은 파일입니다.")
    c = sqlite3.connect(dst_db)
    added = {}
    try:
        c.execute("ATTACH DATABASE ? AS src", (src_db,))
        src_tables = {r[0] for r in c.execute("SELECT name FROM src.sqlite_master WHERE type='table'")}
        for t in ("presets", "cap_settings", "groups"):
            if t not in src_tables:
                added[t] = 0
                continue
            dst_cols = [r[1] for r in c.execute(f"PRAGMA main.table_info({t})") if r[1] != "id"]
            src_cols = {r[1] for r in c.execute(f"PRAGMA src.table_info({t})")}
            cols = [x for x in dst_cols if x in src_cols]
            if t == "presets" and "base" in dst_cols and "base" not in src_cols:     # 아주 예전 파일: base 가 없으면 이름으로
                sel = ", ".join(cols + ["name"])
                cols_sql = ", ".join(cols + ["base"])
            else:
                sel = cols_sql = ", ".join(cols)
            before = c.total_changes
            c.execute(f"INSERT OR IGNORE INTO main.{t}({cols_sql}) SELECT {sel} FROM src.{t}")
            added[t] = c.total_changes - before
        c.commit()
        c.execute("DETACH DATABASE src")
    finally:
        c.close()
    return added


def read_update_info(folder):
    """업데이트 폴더의 version.json 을 읽는다. 없거나 못 읽으면 None.
    version.json 예: {"version": "2.1.0", "package": "RunSheetEditor_2.1.0.zip", "setup": "RunSheetEditor_setup_2.1.0.exe", "notes": "바뀐 내용"}
    package(zip) 와 setup(exe) 중 있는 것을 쓴다. 둘 다 있으면 zip 을 쓴다 (설치 도구 없이 만든 것도 되도록)."""
    if not folder:
        return None
    path = os.path.join(folder, "version.json")
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            info = json.load(f)
    except Exception:
        return None
    if not isinstance(info, dict) or not info.get("version") or not (info.get("setup") or info.get("package")):
        return None
    info["package_path"] = os.path.join(folder, os.path.basename(info["package"])) if info.get("package") else ""
    info["setup_path"] = os.path.join(folder, os.path.basename(info["setup"])) if info.get("setup") else ""
    if info["package_path"] and os.path.exists(info["package_path"]):
        info["file_path"] = info["package_path"]
    elif info["setup_path"] and os.path.exists(info["setup_path"]):
        info["file_path"] = info["setup_path"]
    else:
        info["file_path"] = ""
    return info


M = 18                       # 캔버스 안쪽 여백(px)
ZOOM_LEVELS = (0.50, 0.60, 0.75, 0.90, 1.00, 1.25, 1.50, 1.75, 2.00, 2.50, 3.00)

# ---- 운행표 칸 구조 (A4 세로, 단위 pt) -------------------------------------
# 칸 위치가 화면에서 조금 어긋나 보이면 아래 값만 고치면 된다.
AD_BORDERS = (206.0, 286.0, 366.0, 446.0, 536.0)   # 광고주 4칸의 좌우 경계 (마지막 값은 판매/재원 칸 시작)
AD_TEXT_START = (216.0, 296.0, 376.0, 456.0)       # 각 칸 안에서 글자가 시작하는 x (원본의 '20' 위치)
COMPRESS_MIN = 5            # 한 줄의 광고(원본 + 추가)가 이 수 이상이면 '줄임 표시'(초수 + 앞 두 글자)
SPLIT_INSET = 3.0           # 반으로 나눈 칸의 왼쪽 글자 간격
MIN_XSCALE = 0.7            # 글자 폭을 줄이는(간격을 좁히는) 최소 비율. 그래도 넘치면 끝을 자른다
OVERFLOW_MAX = 4            # 칸이 모자랄 때 '+N' 으로만 기록할 수 있는 최대 수
MARK_X = 582.0              # '+N' 표시 위치(행 오른쪽 여백)
# '+N' 으로만 기록된 광고(원본 칸이 다 찬 줄)의 이름표: 그 줄 바로 아래(또는 위) 경계선에 걸쳐 작게 쓴다.
# 위치가 어색하면 아래 값만 고치면 된다.
OVER_LABEL_POS = "below"    # 'below' = 줄 아래 경계선, 'above' = 줄 위 경계선
OVER_LABEL_FS = 5.5         # 이름표 글자 크기(pt)
OVER_LABEL_X0 = 206.0       # 이름표가 쓸 수 있는 왼쪽 끝 (광고주1 칸 시작)
OVER_LABEL_X1 = 580.0       # 이름표 오른쪽 끝 (오른쪽 맞춤, '+N' 바로 옆)
OVER_LABEL_DY = 0.0         # 위아래로 조금 옮기고 싶을 때 (+ 는 아래로)
OVER_LABEL_PAD = 0.6        # 이름표 바탕(흰색) 위아래 여백(pt)
HL_COLOR = (0.11, 0.31, 0.85)  # 출력물 강조 테두리 색 (진한 파랑, 흑백 출력에서는 진한 회색)
HL_WIDTH = 1.0              # 출력물 강조 테두리 두께(pt)
# ※ 위 이름표·강조 값은 프로그램 [🏷 표시 설정]에서 바꿀 수 있고, 여기 값은 처음(기본값)으로 쓰인다.
AMOUNT_X1 = 580.0                                  # 판매/재원 칸의 오른쪽 끝 (재원 표시 테두리용)
CELL_HALF_H = 7.3                                  # 칸 높이의 절반
DATA_Y_MIN, DATA_Y_MAX = 100.0, 800.0              # 표 본문이 있는 y 범위 (머리글/바닥글 제외)
FONT_SIZE = 8.0                                    # 추가하는 글자 크기 (원본과 같은 8pt)
TEXT_INSET = 10.0                                  # 칸 왼쪽 끝에서 글자까지 간격
TEXT_MIN_SIZE = 5.0                                # 칸에 안 들어가면 줄여 쓰는 최소 크기

COLS_KO = ("시분초", "프로그램명", "형태", "녹/생", "광고주1", "광고주2", "광고주3", "광고주4", "판매/재원")
BY_OPTIONS = {"광고주별": "name", "그룹별": "grp", "프로그램별": "program", "날짜별": "air_date", "월별": "substr(air_date,1,7)", "매체별": "CASE WHEN media='' THEN 'AM' ELSE media END"}   # 매체가 비어 있는 예전 기록은 AM (일정표와 같은 기준)
SEARCH_TARGETS = {"광고주만": "name", "프로그램만": "program", "광고주+프로그램": "both"}


# ----------------------------------------------------------------------------
# 글꼴 도우미
# ----------------------------------------------------------------------------
def get_system_korean_font_path():
    try:
        reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                val_name, val_data, _ = winreg.EnumValue(key, i)
                if any(k in val_name.lower() for k in ["malgun", "맑은 고딕", "nanum", "나눔고딕", "gulim", "굴림", "batang", "바탕"]):
                    if val_data.lower().endswith(('.ttf', '.ttc', '.otf')):
                        full = val_data if os.path.isabs(val_data) else os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", val_data)
                        if os.path.exists(full):
                            return full
    except Exception:
        pass
    for p in ("C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/malgunbd.ttf", "C:/Windows/Fonts/gulim.ttc",
              "C:/Windows/Fonts/batang.ttc", "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
              "/System/Library/Fonts/Supplemental/AppleGothic.ttf"):
        if os.path.exists(p):
            return p
    return None


def get_tk_font_family(font_path):
    if not font_path or not os.path.exists(font_path):
        return "맑은 고딕"
    name = os.path.splitext(os.path.basename(font_path))[0].lower()
    for key, fam in (("malgun", "맑은 고딕"), ("gulim", "굴림"), ("batang", "바탕"), ("nanum", "나눔고딕")):
        if key in name:
            return fam
    return "맑은 고딕"


def est_text_width(text, fs):
    """글자 폭 추정(영문/숫자는 좁게, 한글은 넓게)."""
    return sum(fs * (0.55 if ord(ch) < 128 else 1.0) for ch in text)


# ----------------------------------------------------------------------------
# 운행표 읽기 (단어 좌표 -> 행/칸)
# ----------------------------------------------------------------------------
BASE_PAGE_W = 595.0     # 칸 위치 상수들은 A4 폭(595pt)으로 인쇄된 운행표를 기준으로 잡은 값이다
REF_TIME_X0 = 27.36     # 기준 운행표에서 머리글 '시분초' 의 왼쪽 x
REF_AMT_X0 = 538.08     # 기준 운행표에서 머리글 '판매/재원' 의 왼쪽 x
REF_HEAD_Y0 = 92.49     # 기준 운행표에서 머리글 줄의 y


def page_transform(page):
    """이 페이지의 인쇄 배율·위치를 기준 운행표(A4)와 비교해 (s, tx, ty) 를 돌려준다.
    실제 좌표 = s * 기준좌표 + (tx, ty).  A3 등으로 인쇄했거나 여백이 달라도 머리글('시분초', '판매/재원')을 기준으로 맞춘다.
    머리글을 못 찾으면 페이지 폭 비율만 쓴다. 기준과 같은 PDF 면 (1, 0, 0)."""
    try:
        words = page.get_text("words")
        head = [w for w in words if w[1] < 0.2 * float(page.rect.height)]
        t = next((w for w in head if w[4] == "시분초"), None)
        a = next((w for w in head if w[4] == "판매/재원"), None)
        if t and a and a[0] - t[0] > 100:
            sc = (a[0] - t[0]) / (REF_AMT_X0 - REF_TIME_X0)
            if 0.5 < sc < 4:
                return (sc, t[0] - sc * REF_TIME_X0, t[1] - sc * REF_HEAD_Y0)
    except Exception:
        pass
    try:
        k = float(page.rect.width) / BASE_PAGE_W
    except Exception:
        return (1.0, 0.0, 0.0)
    return (1.0, 0.0, 0.0) if abs(k - 1.0) < 0.02 or k <= 0 else (k, 0.0, 0.0)


def page_scale(page):
    return page_transform(page)[0]


def words_from_page(page):
    """페이지의 단어를 기준 운행표 좌표로 바꿔서 돌려준다. (모든 칸 계산은 이 좌표를 쓴다)"""
    sc, tx, ty = page_transform(page)
    return [{"x0": (w[0] - tx) / sc, "y0": (w[1] - ty) / sc, "x1": (w[2] - tx) / sc, "y1": (w[3] - ty) / sc, "t": w[4]} for w in page.get_text("words")]


def parse_air_date(words):
    """머리글의 '2026년 10월 30일' 에서 방송일을 읽는다. 못 찾으면 None."""
    head = sorted([w for w in words if w["y0"] < DATA_Y_MIN], key=lambda w: (round(w["y0"]), w["x0"]))
    text = " ".join(w["t"] for w in head)
    m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", text)
    if not m:
        return None
    return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def parse_rows(words, page_no, hour=None):
    """한 페이지의 단어들을 프로그램 행으로 묶는다.
    반환: (rows, 마지막 시각의 '시'). 다음 페이지 첫 행이 '시'를 생략할 때를 위해 넘겨준다.
    row = {page, time, program, form, kind, lines:[{ads[4], amount, cy}]}"""
    ws = [w for w in words if DATA_Y_MIN < w["y0"] < DATA_Y_MAX]
    ws.sort(key=lambda w: (w["y0"], w["x0"]))
    lines = []
    for w in ws:
        if lines and abs(w["y0"] - lines[-1]["y"]) <= 3.5:
            lines[-1]["w"].append(w)
        else:
            lines.append({"y": w["y0"], "w": [w]})

    rows = []
    for ln in lines:
        lw = sorted(ln["w"], key=lambda w: w["x0"])
        tcol = [w for w in lw if w["x0"] < 60]
        pcol = [w for w in lw if 60 <= w["x0"] < 166]
        fcol = [w for w in lw if 166 <= w["x0"] < 192]
        kcol = [w for w in lw if 192 <= w["x0"] < AD_BORDERS[0]]
        acol = [w for w in lw if AD_BORDERS[0] <= w["x0"] < AD_BORDERS[4]]
        mcol = [w for w in lw if w["x0"] >= AD_BORDERS[4]]

        # 광고주: 칸 시작 위치의 숫자(초수)를 만나면 새 칸, 아니면 앞 칸에 이어 붙인다 (긴 이름이 옆으로 넘친 경우)
        ads = ["", "", "", ""]
        cur = 0
        for w in acol:
            if re.fullmatch(r"\d+", w["t"]):
                near = min(range(4), key=lambda k: abs(w["x0"] - AD_TEXT_START[k]))
                if abs(w["x0"] - AD_TEXT_START[near]) <= 8:
                    cur = near
            ads[cur] = (ads[cur] + " " + w["t"]).strip()

        # 판매/재원 칸에는 숫자와 '/' 만 있다. 긴 광고주 이름이 칸을 넘쳐 들어온 글자(예: '주식회사')는 마지막 광고주 칸의 이어진 글자로 본다.
        mtxt = [w for w in mcol if not re.fullmatch(r"[\d,/.]+", w["t"])]
        mcol = [w for w in mcol if re.fullmatch(r"[\d,/.]+", w["t"])]
        for w in mtxt:
            ads[cur] = (ads[cur] + " " + w["t"]).strip()

        ys0 = [w["y0"] for w in lw]
        ys1 = [w["y1"] for w in lw]
        cy = (min(ys0) + max(ys1)) / 2.0
        amount = "".join(w["t"] for w in mcol)
        line = {"ads": ads, "amount": amount, "cy": cy}

        if tcol or pcol or fcol or kcol:
            raw = "".join(w["t"] for w in tcol).replace(" ", "")
            if re.fullmatch(r"\d{2}:\d{2}:\d{2}", raw):
                hour = raw[:2]
                full = raw
            elif re.fullmatch(r"\d{2}:\d{2}", raw) and hour:
                full = f"{hour}:{raw}"
            else:
                full = raw
            rows.append({"page": page_no, "time": full, "program": " ".join(w["t"] for w in pcol),
                         "form": " ".join(w["t"] for w in fcol), "kind": " ".join(w["t"] for w in kcol),
                         "lines": [line]})
        elif rows:
            rows[-1]["lines"].append(line)
    return rows, hour


def parse_amount(row):
    """행의 '판매/재원'(예: '100/330')을 (판매, 재원)으로. 읽지 못하면 None."""
    for line in row["lines"]:
        m = re.fullmatch(r"(\d+)/(\d+)", (line.get("amount") or "").replace(" ", ""))
        if m:
            return int(m.group(1)), int(m.group(2))
    return None


def cell_rect(cy, k):
    return (AD_BORDERS[k], cy - CELL_HALF_H, AD_BORDERS[k + 1], cy + CELL_HALF_H)


# ----------------------------------------------------------------------------
# 날짜 도우미 / 달력 팝업
# ----------------------------------------------------------------------------
def parse_date_input(text):
    """'26/10/31', '2026-10-31', '2026.10.31' 등을 'YYYY-MM-DD'로. 빈 값은 ''. 형식 오류는 ValueError."""
    t = (text or "").strip()
    if not t:
        return ""
    m = re.fullmatch(r"(\d{2}|\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})", t)
    if not m:
        raise ValueError("형식 오류")
    y = int(m.group(1))
    if y < 100:
        y += 2000
    return datetime.date(y, int(m.group(2)), int(m.group(3))).isoformat()   # 없는 날짜면 ValueError


def short_date(iso):
    return f"{iso[2:4]}/{iso[5:7]}/{iso[8:10]}" if iso else "-"


def today_iso():
    return datetime.date.today().isoformat()


def is_expired(end_date, ref=None):
    """종료일 다음 날부터 만료. 종료일이 비어 있으면 만료 없음."""
    return bool(end_date) and end_date < (ref or today_iso())


def pick_date(parent, initial_iso, on_pick):
    """작은 달력 팝업. 날짜를 누르면 on_pick('YYYY-MM-DD') 호출."""
    try:
        d0 = datetime.date.fromisoformat(initial_iso) if initial_iso else datetime.date.today()
    except ValueError:
        d0 = datetime.date.today()
    win = tk.Toplevel(parent)
    win.title("날짜 선택")
    win.transient(parent)
    win.grab_set()
    win.resizable(False, False)
    state = {"y": d0.year, "m": d0.month}
    head = ttk.Frame(win)
    head.pack(fill="x", padx=8, pady=(8, 2))
    title = ttk.Label(head, text="", anchor="center", font=("맑은 고딕", 10, "bold"))
    body = ttk.Frame(win)
    body.pack(padx=8, pady=4)

    def shift(dm):
        m = state["m"] + dm
        y = state["y"] + (m - 1) // 12
        state["y"], state["m"] = y, (m - 1) % 12 + 1
        draw()

    ttk.Button(head, text="◀", width=3, command=lambda: shift(-1)).pack(side="left")
    title.pack(side="left", fill="x", expand=True)
    ttk.Button(head, text="▶", width=3, command=lambda: shift(1)).pack(side="right")

    def choose(d):
        win.destroy()
        on_pick(d.isoformat())

    def draw():
        for w in body.winfo_children():
            w.destroy()
        title.config(text=f"{state['y']}년 {state['m']}월")
        for c_, nm in enumerate("일월화수목금토"):
            ttk.Label(body, text=nm, width=4, anchor="center", foreground="#dc2626" if c_ == 0 else "#2563eb" if c_ == 6 else "#333").grid(row=0, column=c_)
        for r_, week in enumerate(calendar.Calendar(firstweekday=6).monthdayscalendar(state["y"], state["m"]), start=1):
            for c_, day in enumerate(week):
                if day:
                    d = datetime.date(state["y"], state["m"], day)
                    txt = f"[{day}]" if d == datetime.date.today() else str(day)
                    tk.Button(body, text=txt, width=4, relief="flat", command=lambda d=d: choose(d)).grid(row=r_, column=c_, padx=1, pady=1)

    draw()
    bf = ttk.Frame(win)
    bf.pack(pady=(2, 8))
    ttk.Button(bf, text="오늘", command=lambda: choose(datetime.date.today())).pack(side="left", padx=4)
    ttk.Button(bf, text="닫기", command=win.destroy).pack(side="left", padx=4)


# ----------------------------------------------------------------------------
# 데이터베이스 (미리 만든 데이터 + 입력 기록)
# ----------------------------------------------------------------------------
class Store:
    def backup(self, dest_dir, keep=30):
        """DB를 날짜별 파일로 복사한다. 같은 날 다시 하면 그 날 파일을 최신으로 덮어쓰고,
        오래된 백업은 keep 개만 남기고 지운다. SQLite 백업 기능을 써서 저장 중이어도 안전하다."""
        os.makedirs(dest_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(self.path))[0]
        final = os.path.join(dest_dir, f"{stem}_{datetime.date.today().strftime('%Y%m%d')}.db")
        tmp = final + ".tmp"
        if os.path.exists(tmp):
            os.remove(tmp)
        src = sqlite3.connect(self.path)
        dst = sqlite3.connect(tmp)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        os.replace(tmp, final)
        files = sorted(glob.glob(os.path.join(dest_dir, f"{glob.escape(stem)}_[0-9]" + "[0-9]" * 7 + ".db")))
        for old in files[:max(0, len(files) - max(1, int(keep)))]:
            try:
                os.remove(old)
            except OSError:
                pass
        return final

    PRESET_COLS = ("start_date", "min_am", "min_fm")

    def __init__(self, path):
        self.path = path
        with self._conn() as c:
            # name = 화면·기록에 쓰는 이름(구분이 있으면 '이름(구분)'), base = 입력한 이름, variant = 구분. 같은 (name, seconds) 만 중복 금지.
            preset_ddl = """CREATE TABLE presets(
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, base TEXT NOT NULL DEFAULT '', variant TEXT NOT NULL DEFAULT '',
                seconds INTEGER NOT NULL DEFAULT 20, price REAL NOT NULL DEFAULT 0, grp TEXT NOT NULL DEFAULT '',
                end_date TEXT NOT NULL DEFAULT '', start_date TEXT NOT NULL DEFAULT '',
                min_am INTEGER NOT NULL DEFAULT 0, min_fm INTEGER NOT NULL DEFAULT 0, short_name TEXT NOT NULL DEFAULT '',
                UNIQUE(name, seconds))"""
            if not c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='presets'").fetchone():
                c.execute(preset_ddl)
            have = [r[1] for r in c.execute("PRAGMA table_info(presets)")]
            if "short_name" not in have:
                c.execute("ALTER TABLE presets ADD COLUMN short_name TEXT NOT NULL DEFAULT ''")
            if "end_date" not in have:
                c.execute("ALTER TABLE presets ADD COLUMN end_date TEXT NOT NULL DEFAULT ''")
            if "start_date" not in have:
                c.execute("ALTER TABLE presets ADD COLUMN start_date TEXT NOT NULL DEFAULT ''")
            if "min_am" not in have:
                c.execute("ALTER TABLE presets ADD COLUMN min_am INTEGER NOT NULL DEFAULT 0")
            if "min_fm" not in have:
                c.execute("ALTER TABLE presets ADD COLUMN min_fm INTEGER NOT NULL DEFAULT 0")
            if "variant" not in have:
                # 예전 DB: 이름만으로 중복을 막던 표 -> (이름, 초수)로 막는 표로 바꾼다. 바꾸기 전에 파일을 통째로 복사해 둔다.
                try:
                    import shutil
                    shutil.copyfile(path, path + ".v1.3-전환전.bak")
                except Exception:
                    pass
                c.execute("ALTER TABLE presets RENAME TO presets_old")
                c.execute(preset_ddl)
                c.execute("""INSERT INTO presets(id,name,base,variant,seconds,price,grp,end_date,start_date,min_am,min_fm,short_name)
                             SELECT id,name,name,'',seconds,price,grp,end_date,start_date,min_am,min_fm,short_name FROM presets_old""")
                c.execute("DROP TABLE presets_old")

            rec_ddl = """CREATE TABLE records(
                id INTEGER PRIMARY KEY AUTOINCREMENT, air_date TEXT NOT NULL, media TEXT NOT NULL DEFAULT '',
                page INTEGER, time TEXT, program TEXT,
                line INTEGER NOT NULL DEFAULT 0, cell INTEGER NOT NULL, name TEXT NOT NULL, seconds INTEGER NOT NULL DEFAULT 0,
                price REAL NOT NULL DEFAULT 0, grp TEXT NOT NULL DEFAULT '', src_file TEXT, saved_at TEXT,
                UNIQUE(air_date, media, time, program, line, cell))"""
            exists = c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='records'").fetchone()
            if not exists:
                c.execute(rec_ddl)
            elif "media" not in [r[1] for r in c.execute("PRAGMA table_info(records)")]:
                # 예전 DB: 매체(AM/FM) 열이 없다 -> 표를 새로 만들어 기존 기록을 옮긴다 (매체는 비워 둠)
                c.execute("ALTER TABLE records RENAME TO records_old")
                c.execute(rec_ddl)
                c.execute("""INSERT INTO records(id,air_date,media,page,time,program,line,cell,name,seconds,price,grp,src_file,saved_at)
                             SELECT id,air_date,'',page,time,program,line,cell,name,seconds,price,grp,src_file,saved_at FROM records_old""")
                c.execute("DROP TABLE records_old")
            c.execute("CREATE INDEX IF NOT EXISTS ix_rec_date ON records(air_date)")
            c.execute("CREATE INDEX IF NOT EXISTS ix_rec_name ON records(name)")
            c.execute("CREATE TABLE IF NOT EXISTS groups(name TEXT PRIMARY KEY)")
            c.execute("""CREATE TABLE IF NOT EXISTS cap_settings(
                time TEXT NOT NULL, program TEXT NOT NULL, cap INTEGER NOT NULL,
                base_sold INTEGER NOT NULL DEFAULT 0, base_cap INTEGER NOT NULL DEFAULT 0, saved_at TEXT,
                PRIMARY KEY(time, program))""")
            # 지운 기록 (다시 불러올 때 '오늘 지운 것'을 보여 주는 데 쓴다. 30일이 지나면 정리)
            c.execute("""CREATE TABLE IF NOT EXISTS removed_records(
                id INTEGER PRIMARY KEY AUTOINCREMENT, air_date TEXT NOT NULL, media TEXT NOT NULL DEFAULT '',
                page INTEGER, time TEXT, program TEXT, line INTEGER NOT NULL DEFAULT 0, cell INTEGER NOT NULL,
                name TEXT NOT NULL, seconds INTEGER NOT NULL DEFAULT 0, removed_at TEXT)""")
            c.execute("CREATE INDEX IF NOT EXISTS ix_rm_date ON removed_records(air_date)")
            # 예전에 입력한 그룹 이름도 목록에 합쳐 둔다
            c.execute("INSERT OR IGNORE INTO groups(name) SELECT DISTINCT grp FROM presets WHERE grp<>''")

    def _conn(self):
        c = sqlite3.connect(self.path)
        c.row_factory = sqlite3.Row
        return c

    # --- 미리 만든 데이터
    def list_presets(self, keyword=""):
        with self._conn() as c:
            if keyword:
                like = f"%{keyword}%"
                cur = c.execute("SELECT * FROM presets WHERE name LIKE ? OR grp LIKE ? ORDER BY name, seconds", (like, like))
            else:
                cur = c.execute("SELECT * FROM presets ORDER BY name, seconds")
            return [dict(r) for r in cur.fetchall()]

    def get_preset_by_name(self, name, seconds=None):
        with self._conn() as c:
            if seconds is None:
                r = c.execute("SELECT * FROM presets WHERE name=? ORDER BY seconds", (name,)).fetchone()
            else:
                r = c.execute("SELECT * FROM presets WHERE name=? AND seconds=?", (name, int(seconds))).fetchone()
            return dict(r) if r else None

    @staticmethod
    def compose_name(base, variant=""):
        base, variant = (base or "").strip(), (variant or "").strip()
        return f"{base}({variant})" if variant else base

    @staticmethod
    def _clean_preset(name, seconds, grp, min_am, min_fm):
        name = (name or "").strip()
        if not name:
            raise ValueError("이름을 입력해 주세요.")
        return name, int(seconds), (grp or "").strip(), max(0, int(min_am or 0)), max(0, int(min_fm or 0))

    def add_preset(self, name, seconds, price=0, grp="", end_date="", start_date="", min_am=0, min_fm=0, short_name="", variant=""):
        base, seconds, grp, min_am, min_fm = self._clean_preset(name, seconds, grp, min_am, min_fm)
        variant = (variant or "").strip()
        name = self.compose_name(base, variant)
        if start_date and end_date and end_date < start_date:
            raise ValueError("종료일이 시작일보다 앞섭니다.")
        try:
            with self._conn() as c:
                cur = c.execute("INSERT INTO presets(name,base,variant,seconds,price,grp,end_date,start_date,min_am,min_fm,short_name) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                (name, base, variant, seconds, float(price), grp, end_date or "", start_date or "", min_am, min_fm, (short_name or "").strip()))
                return cur.lastrowid
        except sqlite3.IntegrityError:
            raise ValueError(f"이름과 초수가 같은 데이터가 이미 있습니다: {seconds}초 {name}\n(다른 광고라면 '구분'란에 구분을 적어 주세요.)")

    def update_preset(self, pid, name, seconds, price=0, grp="", end_date="", start_date="", min_am=0, min_fm=0, short_name="", variant=""):
        base, seconds, grp, min_am, min_fm = self._clean_preset(name, seconds, grp, min_am, min_fm)
        variant = (variant or "").strip()
        name = self.compose_name(base, variant)
        if start_date and end_date and end_date < start_date:
            raise ValueError("종료일이 시작일보다 앞섭니다.")
        try:
            with self._conn() as c:
                c.execute("UPDATE presets SET name=?,base=?,variant=?,seconds=?,price=?,grp=?,end_date=?,start_date=?,min_am=?,min_fm=?,short_name=? WHERE id=?",
                          (name, base, variant, seconds, float(price), grp, end_date or "", start_date or "", min_am, min_fm, (short_name or "").strip(), pid))
        except sqlite3.IntegrityError:
            raise ValueError(f"이름과 초수가 같은 데이터가 이미 있습니다: {seconds}초 {name}\n(다른 광고라면 '구분'란에 구분을 적어 주세요.)")

    def delete_preset(self, pid):
        with self._conn() as c:
            c.execute("DELETE FROM presets WHERE id=?", (pid,))

    # --- 행별 재원 설정 (시각 + 프로그램명 기준으로 매일 계속 적용)
    def list_caps(self):
        with self._conn() as c:
            return {(r["time"], r["program"]): dict(r) for r in c.execute("SELECT * FROM cap_settings")}

    def set_cap(self, time_, program, cap, base_sold, base_cap):
        with self._conn() as c:
            c.execute("INSERT OR REPLACE INTO cap_settings(time,program,cap,base_sold,base_cap,saved_at) VALUES(?,?,?,?,?,?)",
                      (time_, program, int(cap), int(base_sold), int(base_cap), time.strftime("%Y-%m-%d %H:%M:%S")))

    def delete_cap(self, time_, program):
        with self._conn() as c:
            c.execute("DELETE FROM cap_settings WHERE time=? AND program=?", (time_, program))

    # --- 그룹 목록 (데이터 추가 창에서 선택용)
    def list_groups(self):
        with self._conn() as c:
            return [r[0] for r in c.execute("SELECT name FROM groups ORDER BY name")]

    def add_group(self, name):
        name = (name or "").strip()
        if not name:
            raise ValueError("그룹 이름을 입력해 주세요.")
        with self._conn() as c:
            if c.execute("SELECT 1 FROM groups WHERE name=?", (name,)).fetchone():
                raise ValueError(f"이미 있는 그룹입니다: {name}")
            c.execute("INSERT INTO groups(name) VALUES(?)", (name,))

    def delete_group(self, name):
        """그룹을 목록에서 지운다. 이 그룹을 쓰던 미리 만든 데이터는 '그룹 없음'이 된다(이미 저장된 기록은 그대로)."""
        with self._conn() as c:
            c.execute("DELETE FROM groups WHERE name=?", (name,))
            c.execute("UPDATE presets SET grp='' WHERE grp=?", (name,))

    def distinct_groups(self):
        with self._conn() as c:
            g = {r[0] for r in c.execute("SELECT name FROM groups")}
            g |= {r[0] for r in c.execute("SELECT DISTINCT grp FROM presets WHERE grp<>''")}
            g |= {r[0] for r in c.execute("SELECT DISTINCT grp FROM records WHERE grp<>''")}
        return sorted(g)

    # --- 입력 기록
    def load_records(self, air_date, media=""):
        """해당 방송일(과 매체)의 기록. 매체가 비어 있는 예전 기록도 함께 불러와 이번 저장 때 매체를 채운다."""
        with self._conn() as c:
            if media:
                cur = c.execute("SELECT * FROM records WHERE air_date=? AND (media=? OR media='') ORDER BY time,line,cell", (air_date, media))
            else:
                cur = c.execute("SELECT * FROM records WHERE air_date=? ORDER BY time,line,cell", (air_date,))
            return [dict(r) for r in cur.fetchall()]

    @staticmethod
    def _rec_ident(r):
        return (r["time"], r["program"], int(r.get("line") or 0), int(r["cell"]), r["name"], int(r.get("seconds") or 0))

    def replace_records(self, air_date, recs, src_file, media=""):
        """해당 방송일(과 매체)의 기록을 현재 입력 내용으로 통째로 바꾼다.
        - 그대로 둔 기록은 처음 저장한 시각(saved_at)을 그대로 지킨다 -> 다시 불러오면 '오늘 고친 것'만 따로 표시할 수 있다.
        - 없어진 기록은 removed_records 에 남긴다 -> '오늘 지운 것'을 표시할 수 있다."""
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        with self._conn() as c:
            if media:
                where, par = "air_date=? AND (media=? OR media='')", (air_date, media)
            else:
                where, par = "air_date=?", (air_date,)
            old = [dict(r) for r in c.execute(f"SELECT * FROM records WHERE {where}", par)]
            new_ids = {self._rec_ident(r) for r in recs}
            for o in old:
                if self._rec_ident(o) not in new_ids:
                    c.execute("""INSERT INTO removed_records(air_date,media,page,time,program,line,cell,name,seconds,removed_at)
                                 VALUES(?,?,?,?,?,?,?,?,?,?)""",
                              (air_date, o.get("media") or media, o.get("page"), o["time"], o["program"], o.get("line") or 0,
                               o["cell"], o["name"], o["seconds"], now))
            for r in recs:                     # 지웠다가 다시 넣은 것은 '지운 것'에서 뺀다
                c.execute(f"""DELETE FROM removed_records WHERE {where} AND time=? AND program=? AND line=? AND cell=? AND name=? AND seconds=?""",
                          par + self._rec_ident(r))
            c.execute("DELETE FROM removed_records WHERE removed_at < ?", (time.strftime("%Y-%m-%d", time.localtime(time.time() - 30 * 86400)),))
            c.execute(f"DELETE FROM records WHERE {where}", par)
            for r in recs:
                c.execute("""INSERT OR REPLACE INTO records(air_date,media,page,time,program,line,cell,name,seconds,price,grp,src_file,saved_at)
                             VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                          (air_date, r.get("media") or media, r.get("page"), r["time"], r["program"], r.get("line", 0), r["cell"], r["name"],
                           int(r.get("seconds", 0)), float(r.get("price", 0)), r.get("grp", ""), src_file, r.get("saved_at") or now))
        return now

    def load_removed(self, air_date, media="", day=None):
        """해당 방송일(과 매체)에서 day(기본: 오늘) 에 지운 기록."""
        day = day or time.strftime("%Y-%m-%d")
        with self._conn() as c:
            if media:
                cur = c.execute("SELECT * FROM removed_records WHERE air_date=? AND (media=? OR media='') AND substr(removed_at,1,10)=? ORDER BY time,line,cell",
                                (air_date, media, day))
            else:
                cur = c.execute("SELECT * FROM removed_records WHERE air_date=? AND substr(removed_at,1,10)=? ORDER BY time,line,cell", (air_date, day))
            return [dict(r) for r in cur.fetchall()]

    @staticmethod
    def _where(date_from, date_to, keyword, grp, media="", target="both"):
        cond, par = [], []
        if date_from:
            cond.append("air_date>=?"); par.append(date_from)
        if date_to:
            cond.append("air_date<=?"); par.append(date_to)
        if keyword:
            if target == "name":
                cond.append("name LIKE ?"); par.append(f"%{keyword}%")
            elif target == "program":
                cond.append("program LIKE ?"); par.append(f"%{keyword}%")
            else:
                cond.append("(name LIKE ? OR program LIKE ?)"); par += [f"%{keyword}%", f"%{keyword}%"]
        if grp:
            cond.append("grp=?"); par.append(grp)
        if media == "AM":                       # 매체가 비어 있는 예전 기록은 AM 으로 본다 (일정표와 같은 기준)
            cond.append("(media=? OR media='')"); par.append(media)
        elif media:
            cond.append("media=?"); par.append(media)
        return (" WHERE " + " AND ".join(cond)) if cond else "", par

    def aggregate(self, date_from, date_to, keyword="", grp="", by="name", media="", target="both"):
        expr = by if by in BY_OPTIONS.values() else "name"
        where, par = self._where(date_from, date_to, keyword, grp, media, target)
        with self._conn() as c:
            cur = c.execute(f"SELECT {expr} AS k, COUNT(*) AS cnt, SUM(seconds) AS sec, SUM(price) AS amt FROM records{where} GROUP BY k ORDER BY k", par)
            return [(r["k"], r["cnt"], r["sec"] or 0, r["amt"] or 0) for r in cur.fetchall()]

    def detail(self, date_from, date_to, keyword="", grp="", media="", target="both"):
        where, par = self._where(date_from, date_to, keyword, grp, media, target)
        with self._conn() as c:
            cur = c.execute(f"SELECT air_date,CASE WHEN media='' THEN 'AM' ELSE media END AS media,time,program,cell,name,seconds,price,grp "
                            f"FROM records{where} ORDER BY air_date,media,time,line,cell", par)
            return [tuple(r) for r in cur.fetchall()]

    # --- 웹 달력용
    def calendar_counts(self, date_from, date_to, keyword="", grp="", media="", target="name"):
        """방송일별 [건수, 총 초수, AM 건수, FM 건수]."""
        where, par = self._where(date_from, date_to, keyword, grp, media, target)
        with self._conn() as c:
            cur = c.execute("SELECT air_date, COUNT(*), SUM(seconds), SUM(CASE WHEN media IN ('AM','') THEN 1 ELSE 0 END), "
                            f"SUM(CASE WHEN media='FM' THEN 1 ELSE 0 END) FROM records{where} GROUP BY air_date", par)
            return {r[0]: [r[1], r[2] or 0, r[3] or 0, r[4] or 0] for r in cur.fetchall()}

    def day_by_name(self, air_date, keyword="", grp="", media="", target="name"):
        """그날 광고주별 (광고주, 그룹, 횟수, 총 초수). 횟수가 많은 순."""
        where, par = self._where(air_date, air_date, keyword, grp, media, target)
        with self._conn() as c:
            cur = c.execute(f"SELECT name, GROUP_CONCAT(DISTINCT NULLIF(grp,'')), COUNT(*), SUM(seconds) FROM records{where} "
                            "GROUP BY name ORDER BY COUNT(*) DESC, name", par)
            return [(r[0], r[1] or "", r[2], r[3] or 0) for r in cur.fetchall()]

    # --- 일정표용
    def schedule_records(self, name, date_from, date_to, seconds=None):
        sec_sql, par = ("", [name, date_from, date_to]) if seconds is None else (" AND seconds=?", [name, date_from, date_to, int(seconds)])
        with self._conn() as c:
            cur = c.execute("SELECT air_date,media,time,program,seconds FROM records WHERE name=? AND air_date>=? AND air_date<=?" + sec_sql + " ORDER BY air_date,time,line,cell", par)
            return [dict(r) for r in cur.fetchall()]

    def record_date_span(self, name, seconds=None):
        sec_sql, par = ("", [name]) if seconds is None else (" AND seconds=?", [name, int(seconds)])
        with self._conn() as c:
            r = c.execute("SELECT MIN(air_date), MAX(air_date) FROM records WHERE name=?" + sec_sql, par).fetchone()
            return (r[0] or "", r[1] or "")

    def processed_dates(self, date_from, date_to):
        """기록이 하나라도 저장된 방송일(= 입력 작업을 한 날)."""
        with self._conn() as c:
            return {r[0] for r in c.execute("SELECT DISTINCT air_date FROM records WHERE air_date>=? AND air_date<=?", (date_from, date_to))}

    def last_record_date(self, before, media=""):
        """before 보다 앞선 날 중에서 기록이 저장된 가장 최근 방송일. 없으면 ''."""
        with self._conn() as c:
            if media:
                r = c.execute("SELECT MAX(air_date) FROM records WHERE air_date<? AND (media=? OR media='')", (before, media)).fetchone()
            else:
                r = c.execute("SELECT MAX(air_date) FROM records WHERE air_date<?", (before,)).fetchone()
            return (r[0] if r else "") or ""

    def record_keys(self):
        """기록에 있는 (이름, 초수) 목록."""
        with self._conn() as c:
            return [(r[0], r[1]) for r in c.execute("SELECT DISTINCT name, seconds FROM records ORDER BY name, seconds")]

    def keys_in_period(self, date_from, date_to):
        """기간 안에 기록이 하나라도 있는 (이름, 초수) 목록."""
        with self._conn() as c:
            cur = c.execute("SELECT DISTINCT name, seconds FROM records WHERE air_date>=? AND air_date<=? ORDER BY name, seconds",
                            (date_from or "0000-01-01", date_to or "9999-12-31"))
            return [(r[0], r[1]) for r in cur.fetchall()]

    def record_names(self):
        with self._conn() as c:
            return [r[0] for r in c.execute("SELECT DISTINCT name FROM records ORDER BY name")]


def period_range(kind, key):
    """일/월/년 정산의 기간. key 예: 2026-10-30 / 2026-10 / 2026."""
    key = key.strip()
    if kind == "일":
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", key):
            raise ValueError("날짜를 2026-10-30 형식으로 입력해 주세요.")
        return key, key
    if kind == "월":
        if not re.fullmatch(r"\d{4}-\d{2}", key):
            raise ValueError("월을 2026-10 형식으로 입력해 주세요.")
        return key + "-01", key + "-31"
    if not re.fullmatch(r"\d{4}", key):
        raise ValueError("연도를 2026 형식으로 입력해 주세요.")
    return key + "-01-01", key + "-12-31"


def fmt_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{int(f):,}" if f == int(f) else f"{f:,.2f}"


_CORP_PREFIX = re.compile(r"^(?:\(주\)|\(유\)|\(사\)|\(재\)|㈜|㈲|주식회사|유한회사|사단법인|재단법인|학교법인|사회복지법인|의료법인|농업회사법인|\s)+")


def abbreviate(name, n=2):
    """줄임 표시용 약칭: '주식회사 설총' -> '설총', '(주)와우프레스' -> '와우'."""
    base = _CORP_PREFIX.sub("", name or "").strip()
    base = re.sub(r"[\s()·\-_/\[\]]+", "", base)
    return (base or (name or "").strip())[:n]


def cell_label(c):
    """기록의 칸 번호(1-based)를 사람이 읽는 말로. 1~4 칸, 5~8 칸의 오른쪽 반, 9~ '+N'."""
    c = int(c)
    if c <= 4:
        return str(c)
    if c <= 8:
        return f"{c - 4}(오른쪽)"
    return f"+{c - 8}"


# ----------------------------------------------------------------------------
# 광고주별 일정표 (날짜 x AM/FM, Spot 시간 코드)
# ----------------------------------------------------------------------------
MEDIA_LIST = ("AM", "FM")


def detect_media(path):
    """파일 이름에서 AM / FM 을 찾는다. 없으면 ''."""
    name = os.path.basename(path or "")
    m = re.search(r"(?<![A-Za-z])(AM|FM)(?![A-Za-z])", name, flags=re.I)
    return m.group(1).upper() if m else ""


def detect_media_from_words(words):
    """운행표 제목('AM 방송운행표' / 'FM 방송운행표')에서 매체를 읽는다. 못 찾거나 둘 다 있으면 ''."""
    found = {w["t"].strip().upper() for w in words
             if w["y0"] < DATA_Y_MIN and w["t"].strip().upper() in MEDIA_LIST}
    return found.pop() if len(found) == 1 else ""


def spot_code(program, time_):
    """'Spot 0759' -> '0759'. Spot 이 아니면 시각(07:58)을 대신 쓴다."""
    m = re.search(r"[Ss]pot\s*(\d{4})", program or "")
    if m:
        return m.group(1)
    return (time_ or "")[:5]


def date_range_iso(d0, d1, limit=800):
    a = datetime.date.fromisoformat(d0)
    b = datetime.date.fromisoformat(d1)
    out = []
    while a <= b and len(out) < limit:
        out.append(a.isoformat())
        a += datetime.timedelta(days=1)
    return out


def month_end_iso(iso):
    d = datetime.date.fromisoformat(iso)
    return datetime.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1]).isoformat()


def _md(iso):
    return f"{iso[5:7]}.{iso[8:10]}" if iso else "?"


def schedule_title(p):
    """예: '60초 무등산환경대학 63기 모집 (07.29-08.26) 3회'"""
    t = f"{p['name']} (전체 초수)" if p.get("all_seconds") else f"{p.get('seconds', 0)}초 {p['name']}"
    if p.get("start_date") or p.get("end_date"):
        t += f" ({_md(p.get('start_date'))}-{_md(p.get('end_date'))})"
    a, f = int(p.get("min_am") or 0), int(p.get("min_fm") or 0)
    if a or f:
        t += f" {a}회" if a == f else f" AM {a}회 FM {f}회"
    return t


def combine_presets(name, plist):
    """같은 이름(광고주)의 여러 초수 데이터를 하나로 합친 '전체 초수' 설정.
    기간은 가장 이른 시작일 ~ 가장 늦은 종료일, 하루 최소 횟수는 합쳐서 센다고 보고 가장 큰 값을 쓴다."""
    starts = [x["start_date"] for x in plist if x.get("start_date")]
    ends = [x["end_date"] for x in plist if x.get("end_date")]
    return {"name": name, "seconds": 0, "all_seconds": True, "start_date": min(starts) if starts else "", "end_date": max(ends) if ends else "",
            "min_am": max([int(x.get("min_am") or 0) for x in plist] or [0]), "min_fm": max([int(x.get("min_fm") or 0) for x in plist] or [0])}


def build_schedule(p, recs, d0, d1, processed):
    """한 광고의 날짜별 일정. recs: 그 광고의 기록(air_date, media, time, program).
    매체가 비어 있는 예전 기록은 AM 으로 본다. 최소 횟수 점검은 '입력 작업을 한 날'(processed)에만 한다."""
    by = {}
    for r in recs:
        med = r.get("media") or "AM"
        code = spot_code(r.get("program"), r.get("time"))
        if p.get("all_seconds"):
            code = f"{code}({r.get('seconds', 0)})"      # 전체 초수 보기: 칸마다 초수를 붙여 구분
        by.setdefault((r["air_date"], med), []).append(code)
    start, end = p.get("start_date") or "", p.get("end_date") or ""
    mins = {"AM": int(p.get("min_am") or 0), "FM": int(p.get("min_fm") or 0)}
    rows = []
    for d in date_range_iso(d0, d1):
        outside = bool((start and d < start) or (end and d > end))
        row = {"date": d, "outside": outside}
        shorts = []
        for med in MEDIA_LIST:
            codes = by.get((d, med), [])
            row[med] = " ".join(sorted(set(codes)))
            row[med + "_n"] = len(codes)
            short = (not outside) and d in processed and mins[med] > 0 and len(codes) < mins[med]
            row[med + "_short"] = short
            if short:
                shorts.append(f"{med} {len(codes)}/{mins[med]}")
        row["check"] = " · ".join(shorts)
        rows.append(row)
    return rows


def sheet_name_for(p, used):
    base = re.sub(r"[\[\]:*?/\\]", "", (f"{p['name']} 전체" if p.get("all_seconds") else f"{p.get('seconds', 0)}초 {p['name']}")).strip()[:31] or "광고"
    name, i = base, 2
    while name in used:
        suf = f" ({i})"
        name = base[:31 - len(suf)] + suf
        i += 1
    used.add(name)
    return name


def export_schedule_xlsx(path, sheets):
    """sheets: [{'title','sheet','rows'}]. 사진과 같은 모양: 제목 줄, 날짜/AM/FM, 기간 밖은 회색, 횟수 부족은 노랑."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    gray = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
    yellow = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    center = Alignment(horizontal="center", vertical="center")
    for sh in sheets:
        ws = wb.create_sheet(sh["sheet"])
        ws.merge_cells("A1:C1")
        ws["A1"] = sh["title"]
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = center
        ws.row_dimensions[1].height = 26
        for col in "ABC":
            ws[f"{col}1"].border = border
        for i, h in enumerate(("날짜", "AM", "FM")):
            c = ws.cell(row=2, column=i + 1, value=h)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = border
        for r_i, row in enumerate(sh["rows"], start=3):
            vals = (f"{row['date'][2:4]}.{row['date'][5:7]}.{row['date'][8:10]}", row["AM"], row["FM"])
            for c_i, v in enumerate(vals, start=1):
                c = ws.cell(row=r_i, column=c_i, value=v)
                c.alignment = center
                c.border = border
                if c_i > 1:
                    if row["outside"]:
                        c.fill = gray
                    elif row["AM_short" if c_i == 2 else "FM_short"]:
                        c.fill = yellow
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 46
        ws.column_dimensions["C"].width = 46
        ws.freeze_panes = "A3"
        ws.page_setup.orientation = "landscape"          # 인쇄할 때 AM/FM 열이 잘리지 않게 가로 + 폭 맞춤
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    if not wb.sheetnames:
        wb.create_sheet("빈 시트")
    wb.save(path)


# ----------------------------------------------------------------------------
# 메인 앱
# ----------------------------------------------------------------------------
# ============================================================================
# 다른 PC에서 정산/검색 조회 (읽기 전용 웹 페이지)
#   설정에서 '다른 PC 조회 허용'을 켠 편집기만 사내망에 작은 웹 서버를 연다.
#   다른 PC는 브라우저로 http://이PC:포트 에 들어가 정산/검색만 본다. 기록을 바꾸는 기능은 없다.
# ============================================================================
VIEW_PORT_DEFAULT = 8080
VIEW_DETAIL_LIMIT = 5000
PREVIEW_DIR = os.path.join(DATA_DIR, "미리보기")      # 다른 PC 조회용 운행표 미리보기 PDF + 상태(status.json)
PREVIEW_KEEP_DAYS = 7
_PREVIEW_KEY_RE = re.compile(r"^(AM|FM)\d{8}$")


def read_preview_status():
    try:
        with open(os.path.join(PREVIEW_DIR, "status.json"), "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def update_preview_status(key, **fields):
    """운행표 하나의 상태를 고친다. (상태: '수정완료' / '편집중', dirty: 저장 안 된 수정이 있음)"""
    try:
        os.makedirs(PREVIEW_DIR, exist_ok=True)
        d = read_preview_status()
        cur = d.get(key) or {}
        cur.update(fields)
        cur["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
        d[key] = cur
        cutoff = (datetime.date.today() - datetime.timedelta(days=PREVIEW_KEEP_DAYS)).isoformat()
        for k in [k for k, v in d.items() if (v.get("air_date") or "9999") < cutoff]:   # 오래된 것 정리
            d.pop(k, None)
            try:
                os.remove(os.path.join(PREVIEW_DIR, k + ".pdf"))
            except OSError:
                pass
        tmp = os.path.join(PREVIEW_DIR, "status.json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        os.replace(tmp, os.path.join(PREVIEW_DIR, "status.json"))
    except Exception:
        pass


def list_previews():
    """조회 페이지용 목록: 최근 7일 + 앞으로의 운행표. 최신 방송일이 위."""
    d = read_preview_status()
    today = datetime.date.today()
    cutoff = (today - datetime.timedelta(days=PREVIEW_KEEP_DAYS)).isoformat()
    out = []
    for k, v in d.items():
        ad = v.get("air_date") or ""
        if not _PREVIEW_KEY_RE.match(k) or ad < cutoff:
            continue
        rel = {0: "오늘", 1: "내일", -1: "어제"}.get((datetime.date.fromisoformat(ad) - today).days, "") if re.fullmatch(r"\d{4}-\d{2}-\d{2}", ad) else ""
        out.append({"key": k, "media": v.get("media", ""), "air_date": ad, "rel": rel, "state": v.get("state", ""),
                    "dirty": bool(v.get("dirty")), "saved_at": v.get("saved_at", ""), "updated": v.get("updated", ""),
                    "has_pdf": os.path.exists(os.path.join(PREVIEW_DIR, k + ".pdf"))})
    out.sort(key=lambda x: (x["air_date"], x["media"]), reverse=True)
    return out


def local_addresses():
    """이 PC의 사내망 주소들 (127.x 제외). 가장 그럴듯한 것이 앞."""
    import socket
    ips = []
    try:   # 바깥으로 실제로 보내지는 않는다. 어느 랜카드를 쓰는지만 알아본다
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("10.255.255.255", 1))
            ips.append(s.getsockname()[0])
        finally:
            s.close()
    except Exception:
        pass
    try:
        for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
            if ip not in ips:
                ips.append(ip)
    except Exception:
        pass
    return [ip for ip in ips if not ip.startswith("127.")]


def computer_name():
    import socket
    return os.environ.get("COMPUTERNAME") or socket.gethostname() or "편집기"


def quote_hdr(s):
    """HTTP 머리글에 한글을 넣을 수 있게 %XX 로 바꾼다."""
    from urllib.parse import quote
    return quote(str(s or ""), safe="")


# 사내망에서 서로 찾기 (런처의 Discovery 와 같은 약속: UDP 50808, {"app": "edit-view", ...})
#   조회 페이지를 연 편집기/런처가 '여기 운행표 조회 있음'을 알리면 녹음 PC 서버(사본)가 찾아와 데이터를 모아 간다.
DISCOVERY_PORT = 50808


def announce_view(kind, port):
    import socket
    msg = json.dumps({"app": "edit-view", "kind": kind, "port": int(port), "name": computer_name()}).encode("utf-8")
    try:
        sk = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sk.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            sk.sendto(msg, ("255.255.255.255", DISCOVERY_PORT))
        finally:
            sk.close()
    except Exception:
        pass


_VIEW_HTML = r"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>💰 정산 / 검색 (조회 전용)</title>
<style>
body{font-family:"Malgun Gothic","맑은 고딕",sans-serif;margin:0;background:#f5f6f8;color:#1f2937;font-size:14px}
header{background:#1e3a8a;color:#fff;padding:10px 18px;display:flex;gap:16px;align-items:baseline;flex-wrap:wrap}
header b{font-size:17px} header span{opacity:.85;font-size:12px}
main{padding:12px 18px;max-width:1300px}
.tabs{display:flex;gap:2px;flex-wrap:wrap;border-bottom:1px solid #cbd5e1}
.tabs button{background:#e5e7eb;color:#1f2937;border:1px solid #cbd5e1;border-bottom:none;border-radius:6px 6px 0 0;padding:6px 14px;cursor:pointer}
.tabs button.on{background:#fff;color:#1e3a8a;font-weight:bold;position:relative;top:1px}
.pane{background:#fff;border:1px solid #cbd5e1;border-top:none;padding:10px 12px}
.row{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin:3px 0 6px}
label{font-weight:bold;font-size:13px;color:#1f2937;margin-left:8px} label:first-child{margin-left:0}
.hint{color:#666;font-size:12px}
input,select,button{font:inherit;padding:3px 6px;border:1px solid #cbd5e1;border-radius:4px;background:#fff}
.btn{background:#2563eb;color:#fff;border-color:#2563eb;cursor:pointer;padding:4px 14px}
.btn.lite{background:#fff;color:#1e3a8a}
.right{margin-left:auto}
table{border-collapse:collapse;width:100%;background:#fff}
th,td{border:1px solid #e5e7eb;padding:3px 7px;text-align:left;white-space:nowrap}
th{background:#eef2ff;position:sticky;top:0}
td.n,th.n{text-align:right}
.scroll{height:62vh;overflow:auto;border:1px solid #e5e7eb}
.total{text-align:right;font-weight:bold;margin-top:5px}
.msg{color:#6b7280;font-size:12px;margin:4px 0}
.cal{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin:6px 0}
.cal .h{text-align:center;font-weight:bold;color:#374151;padding:3px 0}
.cal .d{border:1px solid #e5e7eb;border-radius:6px;min-height:68px;padding:4px 6px;cursor:pointer;background:#fff;position:relative}
.cal .d:hover{border-color:#93c5fd;background:#f8fbff}
.cal .d.out{background:#f9fafb;color:#c0c4cc;cursor:default}
.cal .d.today{border:2px solid #1e3a8a}
.cal .d.sel{background:#dbeafe;border-color:#2563eb}
.cal .dn{font-weight:bold;font-size:15px} .cal .sun{color:#dc2626} .cal .sat{color:#2563eb}
.cal .cnt{font-size:20px;font-weight:bold;color:#1e3a8a;text-align:right;margin-top:2px}
.cal .sub{font-size:11px;color:#6b7280;text-align:right}
.cal .st{position:absolute;top:5px;right:6px;font-size:11px}
.dayh{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:10px 0 4px;font-weight:bold;font-size:15px}
.two{display:grid;grid-template-columns:1fr 1fr;gap:0 10px}
@media(max-width:900px){.two{grid-template-columns:1fr}}
</style></head><body>
<header><b>💰 정산 / 검색</b><span id="who"></span><span>조회 전용 — 이 화면에서는 기록을 바꿀 수 없습니다</span></header>
<main>
<div class="tabs" id="tabs"></div>
<div class="pane">
 <div id="fS">
  <div class="row"><label id="kLbl">기준일:</label><input id="kref" size="11"><span class="hint" id="kHint"></span>
   <label>묶음 기준:</label><select id="sby"></select>
   <label>매체:</label><select id="smedia"><option value="">전체</option><option>AM</option><option>FM</option></select>
   <button class="btn" onclick="run()">조회</button>
   <a class="right" id="scsv" href="#">⬇ CSV 저장</a></div>
 </div>
 <div id="fC" style="display:none">
  <div class="row"><button class="btn lite" onclick="mv(-1)">◀</button><select id="cy"></select><select id="cm"></select><button class="btn lite" onclick="mv(1)">▶</button>
   <button class="btn lite" onclick="thisMonth()">이번 달</button>
   <label>매체:</label><select id="cmedia" onchange="cal()"><option value="">전체</option><option>AM</option><option>FM</option></select>
   <label>광고주:</label><input id="ckw" size="16" placeholder="비우면 전체">
   <label>그룹:</label><input id="cgrp" list="grps" size="12">
   <button class="btn" onclick="cal()">조회</button></div>
  <div class="msg" id="cmsg"></div>
  <div id="calv"></div>
  <div id="dayv"></div>
 </div>
 <div id="fQ" style="display:none">
  <div class="row"><label>검색어:</label><input id="kw" size="22">
   <label>찾는 곳:</label><select id="target"></select>
   <label>매체:</label><select id="qmedia"><option value="">전체</option><option>AM</option><option>FM</option></select>
   <label>그룹:</label><input id="grp" list="grps" size="14"><datalist id="grps"></datalist></div>
  <div class="row"><label>기간:</label><input type="date" id="from"> ~ <input type="date" id="to">
   <span class="hint">(비우면 전체 기간, 기간 안의 기록만 나옵니다)</span>
   <label>묶어서 보기:</label><select id="qby"></select>
   <button class="btn" onclick="run()">검색</button>
   <a class="right" id="qcsv" href="#">⬇ CSV 저장</a></div>
 </div>
 <div class="msg" id="msg"></div>
 <div class="scroll"><table id="tbl"></table></div>
 <div class="total" id="total"></div>
</div>
</main>
<script>
const $=id=>document.getElementById(id);let cur='C',info=null;const WK='일월화수목금토';
const TABS=[['C','📅 달력'],['일',' 일 정산 '],['월',' 월 정산 '],['년',' 년 정산 '],['Q',' 검색 / 그룹 검색 '],['P','📄 운행표 미리보기 · 인쇄'],['T','📺 TV 주간 편성표']];
const pad=n=>String(n).padStart(2,'0');const t0=new Date();const today=t0.getFullYear()+'-'+pad(t0.getMonth()+1)+'-'+pad(t0.getDate());
const refs={'일':today,'월':today.slice(0,7),'년':today.slice(0,4)},bys={'일':'광고주별','월':'광고주별','년':'광고주별'},meds={'일':'','월':'','년':''};
const num=v=>Number(v||0).toLocaleString('ko-KR');
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function drawTabs(){$('tabs').innerHTML=TABS.map(([k,t])=>'<button class="'+(k===cur?'on':'')+'" onclick="tab(\''+k+'\')">'+esc(t)+'</button>').join('')}
function keep(){if('일월년'.includes(cur)){refs[cur]=$('kref').value;bys[cur]=$('sby').value;meds[cur]=$('smedia').value}}
function tab(k){keep();cur=k;drawTabs();const s='일월년'.includes(k);$('fS').style.display=s?'':'none';$('fQ').style.display=k==='Q'?'':'none';
 $('fC').style.display=k==='C'?'':'none';$('msg').textContent='';document.querySelector('.scroll').style.display=k==='C'?'none':'';
 $('total').textContent='';
 if(s){$('kLbl').textContent={'일':'기준일:','월':'기준월:','년':'기준연도:'}[k];$('kHint').textContent={'일':'(2026-10-30)','월':'(2026-10)','년':'(2026)'}[k];
  $('kref').value=refs[k];$('sby').value=bys[k];$('smedia').value=meds[k]}
 run()}
function sq(){return new URLSearchParams({kind:cur,ref:$('kref').value.trim(),by:$('sby').value,media:$('smedia').value}).toString()}
function qq(){return new URLSearchParams({kw:$('kw').value.trim(),target:$('target').value,media:$('qmedia').value,grp:$('grp').value.trim(),from:$('from').value,to:$('to').value,by:$('qby').value}).toString()}
function aggTable(by,rows){let h='<tr><th>'+esc(by.replace('별',''))+'</th><th class=n>건수</th><th class=n>총 초수</th></tr>';
 for(const x of rows)h+='<tr><td>'+esc(x[0]||'(없음)')+'</td><td class=n>'+num(x[1])+'</td><td class=n>'+num(x[2])+'</td></tr>';return h}
async function get(u){const r=await fetch(u);const d=await r.json();if(d.error)throw new Error(d.error);return d}
const cq=()=>({media:$('cmedia').value,kw:$('ckw').value.trim(),grp:$('cgrp').value.trim(),target:'광고주만'});
let selDay=today;
function thisMonth(){$('cy').value=today.slice(0,4);$('cm').value=today.slice(5,7);cal()}
function mv(n){let y=+$('cy').value,m=+$('cm').value+n;if(m<1){m=12;y--}if(m>12){m=1;y++}
 if(![...$('cy').options].some(o=>o.value==y)){const o=document.createElement('option');o.value=o.textContent=y;$('cy').appendChild(o)}
 $('cy').value=y;$('cm').value=pad(m);cal()}
async function cal(){const y=+$('cy').value,m=+$('cm').value,ym=y+'-'+pad(m);$('cmsg').textContent='불러오는 중…';
 try{const d=await get('api/calendar?'+new URLSearchParams(Object.assign({ym},cq())));
  const first=new Date(y,m-1,1),last=new Date(y,m,0).getDate();let h='';
  for(const [i,w] of [...WK].entries())h+='<div class="h '+(i===0?'sun':i===6?'sat':'')+'">'+w+'</div>';
  for(let i=0;i<first.getDay();i++)h+='<div class="d out"></div>';
  let mc=0,ms=0;
  for(let dd=1;dd<=last;dd++){const iso=ym+'-'+pad(dd),x=d.days[iso],wd=(first.getDay()+dd-1)%7;
   if(x){mc+=x[0];ms+=x[1]}
   const st=(d.status[iso]||[]).map(s=>'<span title="'+esc(s.media+' '+s.state)+'" style="color:'+(s.state==='수정완료'?'#15803d':'#c2410c')+'">'+(s.state==='수정완료'?'✔':'✏')+esc(s.media)+'</span>').join(' ');
   h+='<div class="d'+(iso===today?' today':'')+(iso===selDay?' sel':'')+'" onclick="day(\''+iso+'\')"><div class="dn '+(wd===0?'sun':wd===6?'sat':'')+'">'+dd+'</div>'
    +(st?'<div class="st">'+st+'</div>':'')
    +(x?'<div class="cnt">'+num(x[0])+'</div><div class="sub">AM '+num(x[2])+' · FM '+num(x[3])+' · '+num(x[1])+'초</div>':'')+'</div>'}
  $('calv').innerHTML='<div class="cal">'+h+'</div>';
  $('cmsg').textContent=y+'년 '+m+'월'+($('ckw').value.trim()?' · 「'+$('ckw').value.trim()+'」':'')+' — 건수 '+num(mc)+' · 총 초수 '+num(ms)+'  (날짜를 누르면 그날 내역이 아래에 나옵니다)';
  if(selDay&&selDay.slice(0,7)===ym)day(selDay,true);else $('dayv').innerHTML=''}
 catch(e){$('cmsg').textContent='⚠ '+(e instanceof TypeError?'편집기 PC에 연결하지 못했습니다.':e.message)}}
function shift(iso,n){const d=new Date(iso+'T00:00:00');d.setDate(d.getDate()+n);return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
async function day(iso,keepCal){selDay=iso;
 if(iso.slice(0,7)!==$('cy').value+'-'+$('cm').value){$('cy').value=iso.slice(0,4);$('cm').value=iso.slice(5,7);if(![...$('cy').options].some(o=>o.value==iso.slice(0,4))){const o=document.createElement('option');o.value=o.textContent=iso.slice(0,4);$('cy').appendChild(o);$('cy').value=iso.slice(0,4)}return cal()}
 if(!keepCal)document.querySelectorAll('.cal .d').forEach(e=>e.classList.toggle('sel',(e.getAttribute('onclick')||'').includes(iso)));
 try{const qs=new URLSearchParams(Object.assign({date:iso},cq())).toString();const d=await get('api/day?'+qs);
  const pv=d.previews.map(x=>x.has_pdf?'<a target="_blank" href="runsheet/'+esc(x.key)+'.pdf">📄 '+esc(x.media)+' 운행표 보기·인쇄</a> <span style="font-size:12px;color:'+(x.state==='수정완료'?'#15803d':'#c2410c')+'">('+esc(x.state)+')</span>':'').join(' &nbsp; ');
  let h='<div class="dayh"><button class="btn lite" onclick="day(shift(selDay,-1))">◀</button>'+esc(iso)+' ('+WK[new Date(iso+'T00:00:00').getDay()]+') 운행 내역<button class="btn lite" onclick="day(shift(selDay,1))">▶</button>'
   +'<span style="font-weight:normal;font-size:13px">'+(pv||'<span class="hint">(이 날 운행표 미리보기 없음)</span>')+'</span></div>';
  h+='<div class="row"><b>광고주별 횟수</b><span class="hint">총 '+d.names.length+'곳 · '+num(d.rows.length)+'건</span><a class="right" href="api/day.csv?part=names&'+qs+'">⬇ CSV 저장</a></div>';
  h+='<div class="scroll" style="height:auto;max-height:300px"><table><tr><th>광고주</th><th>그룹</th><th class=n>횟수</th><th class=n>총 초수</th></tr>'
   +d.names.map(x=>'<tr><td>'+esc(x[0])+'</td><td>'+esc(x[1])+'</td><td class=n>'+num(x[2])+'회</td><td class=n>'+num(x[3])+'</td></tr>').join('')+'</table></div>';
  const half=Math.ceil(d.rows.length/2),tb=rs=>'<table><tr><th>시분초</th><th>매체</th><th>프로그램</th><th>광고주</th><th class=n>초수</th></tr>'
   +rs.map(x=>'<tr><td>'+esc(x[0])+'</td><td>'+esc(x[1])+'</td><td>'+esc(x[2])+'</td><td>'+esc(x[3])+'</td><td class=n>'+num(x[4])+'</td></tr>').join('')+'</table>';
  h+='<div class="row" style="margin-top:10px"><b>운행 내역</b><span class="hint">시각 순</span><a class="right" href="api/day.csv?part=rows&'+qs+'">⬇ CSV 저장</a></div>';
  h+=d.rows.length?'<div class="two"><div>'+tb(d.rows.slice(0,half))+'</div><div>'+tb(d.rows.slice(half))+'</div></div>':'<div class="hint">이 날 기록이 없습니다.</div>';
  $('dayv').innerHTML=h;if(!keepCal)$('dayv').scrollIntoView({behavior:'smooth',block:'start'})}
 catch(e){$('dayv').innerHTML='<div class="msg">⚠ '+esc(e instanceof TypeError?'편집기 PC에 연결하지 못했습니다.':e.message)+'</div>'}}
async function run(){if(cur==='C')return cal();if(cur==='P')return previews();if(cur==='T')return tvsheets();
 $('msg').textContent='불러오는 중…';
 try{if(cur==='Q'){$('qcsv').href='api/search.csv?'+qq();const d=await get('api/search?'+qq());
   if(d.mode==='detail'){let h='<tr><th>방송일</th><th>매체</th><th>시분초</th><th>프로그램</th><th>칸</th><th>광고주</th><th class=n>초수</th><th>그룹</th></tr>';
    for(const x of d.rows)h+='<tr><td>'+esc(x[0])+'</td><td>'+esc(x[1])+'</td><td>'+esc(x[2])+'</td><td>'+esc(x[3])+'</td><td>'+esc(x[4])+'</td><td>'+esc(x[5])+'</td><td class=n>'+num(x[6])+'</td><td>'+esc(x[7])+'</td></tr>';
    $('tbl').innerHTML=h;$('total').textContent=num(d.total)+'건   총 초수 '+num(d.sec);
    $('msg').textContent=d.total>d.rows.length?'화면에는 앞의 '+num(d.rows.length)+'건만 보입니다 (전체는 CSV 저장으로 받으세요)':(d.total?'':'조건에 맞는 기록이 없습니다.')}
   else{$('tbl').innerHTML=aggTable($('qby').value,d.rows);$('total').textContent=d.rows.length+'개 묶음   건수 '+num(d.rows.reduce((a,x)=>a+x[1],0));
    $('msg').textContent=d.rows.length?'':'조건에 맞는 기록이 없습니다.'}}
  else{$('scsv').href='api/settle.csv?'+sq();const d=await get('api/settle?'+sq());
   $('tbl').innerHTML=aggTable($('sby').value,d.rows);
   $('total').textContent='합계  건수 '+num(d.rows.reduce((a,x)=>a+x[1],0))+'   총 초수 '+num(d.rows.reduce((a,x)=>a+x[2],0));
   $('msg').textContent=(d.from===d.to?d.from:d.from+' ~ '+d.to)+(d.rows.length?'':' — 기록이 없습니다.')}}
 catch(e){$('tbl').innerHTML='';$('total').textContent='';$('msg').textContent='⚠ '+(e instanceof TypeError?'편집기 PC에 연결하지 못했습니다. 편집기가 켜져 있는지 확인해 주세요.':e.message)}}
function tvManual(){const v=prompt('TV 편성표 PC 주소 (예: http://192.168.0.20:8081)\n자동으로 못 찾을 때만 넣으세요. 비우면 지웁니다.',localStorage.getItem('tvaddr')||'');
 if(v===null)return;try{v.trim()?localStorage.setItem('tvaddr',v.trim().replace(/\/+$/,'')):localStorage.removeItem('tvaddr')}catch(e){}tvsheets()}
async function tvsheets(){$('msg').textContent='TV 편성표 PC를 찾는 중…';
 let bases=[];try{const p=await (await fetch('api/peers')).json();for(const x of p.rows)bases.push({url:'http://'+x.ip+':'+x.port,name:x.name})}catch(e){}
 let man='';try{man=localStorage.getItem('tvaddr')||''}catch(e){}
 if(man&&!bases.some(b=>b.url===man))bases.push({url:man,name:'직접 넣은 주소'});
 let h='<tr><th>편성표</th><th>종류</th><th>상태</th><th>마지막 저장</th><th>미리보기 · 인쇄</th><th>TV PC</th></tr>',n=0,bad=[];
 for(const b of bases){try{const d=await (await fetch(b.url+'/api/tv')).json();
   for(const x of d.rows){n++;const st=x.dirty?'<b style="color:#c2410c">✏ 현재 편집중</b>':(x.state==='수정완료'?'<b style="color:#15803d">✔ 수정완료</b>':'<b style="color:#c2410c">✏ 편집중</b>');
    h+='<tr><td>'+esc(x.title||x.src)+'</td><td>'+esc(x.kind)+'</td><td>'+st+'</td><td>'+esc((x.saved_at||'').slice(5,16))+'</td><td>'
     +(x.has_pdf?'<a target="_blank" href="'+esc(b.url)+'/tv/'+encodeURIComponent(x.key)+'.pdf">📄 보기 · 인쇄</a>':'<span style="color:#9ca3af">아직 저장 전</span>')+'</td><td>'+esc(d.name||b.name)+'</td></tr>'}}
  catch(e){bad.push(b.url)}}
 $('tbl').innerHTML=h;
 $('msg').innerHTML=(bases.length?(n?'TV 편집기에서 PDF를 저장할 때마다 바뀝니다. (30초마다 새로고침)':'TV 편성표 PC는 찾았지만 저장된 편성표가 아직 없습니다.')
   :'TV 편성표 PC를 찾지 못했습니다. TV 편집기 ⚙ 설정에서 「다른 PC에서 보기」를 켜 주세요.')
  +(bad.length?' ⚠ 연결 안 됨: '+esc(bad.join(', ')):'')+' <a href="#" onclick="tvManual();return false">주소 직접 넣기</a>'}
setInterval(()=>{if(cur==='T'&&!document.hidden)tvsheets()},30000);
function dlabel(s){const d=new Date(s+'T00:00:00');return isNaN(d)?s:(s.slice(5).replace('-','/')+'('+WK[d.getDay()]+')')}
async function previews(){$('msg').textContent='불러오는 중…';
 try{const d=await (await fetch('api/runsheets')).json();
  let h='<tr><th>방송일</th><th>매체</th><th>상태</th><th>마지막 저장</th><th>미리보기 · 인쇄</th></tr>';
  for(const x of d.rows){
   const st=x.dirty?'<b style="color:#c2410c">✏ 현재 편집중</b> <span style="color:#9a3412;font-size:12px">(저장 안 된 수정 있음 · 아래 미리보기는 마지막 저장 상태)</span>'
     :(x.state==='수정완료'?'<b style="color:#15803d">✔ 수정완료</b>':'<b style="color:#c2410c">✏ 편집중</b> <span style="color:#9a3412;font-size:12px">(기록만 저장됨)</span>');
   const rel=x.rel?'<b style="background:'+(x.rel==='오늘'?'#1e3a8a':'#64748b')+';color:#fff;border-radius:4px;padding:1px 6px;margin-right:6px">'+x.rel+'</b>':'';
   h+='<tr'+(x.rel==='오늘'||x.rel==='내일'?' style="background:#f8fafc"':'')+'><td>'+rel+esc(dlabel(x.air_date))+'</td><td>'+esc(x.media)+'</td><td>'+st+'</td><td>'+esc((x.saved_at||'').slice(5,16))+'</td><td>'
     +(x.has_pdf?'<a href="runsheet/'+esc(x.key)+'.pdf" target="_blank">📄 보기 · 인쇄</a>':'<span style="color:#9ca3af">미리보기 없음</span>')+'</td></tr>'}
  $('tbl').innerHTML=h;$('msg').textContent=d.rows.length?'편집기에서 저장할 때마다 바뀝니다. 「보기 · 인쇄」를 누르면 새 창에서 열리고, 브라우저의 인쇄 버튼으로 인쇄합니다. (30초마다 새로고침)':'아직 미리보기가 없습니다. 편집기에서 운행표를 저장하면 여기에 나타납니다.'}
 catch(e){$('msg').textContent='⚠ 편집기 PC에 연결하지 못했습니다.'}}
setInterval(()=>{if(cur==='P'&&!document.hidden)previews()},30000);
const opt=(sel,list)=>{for(const t of list){const o=document.createElement('option');o.textContent=t;sel.appendChild(o)}};
(async()=>{drawTabs();
 try{info=await (await fetch('api/info')).json();
  $('who').textContent='편집기: '+info.name+' · v'+info.version;document.title+=' — '+info.name;
  for(const g of info.groups){const o=document.createElement('option');o.value=g;$('grps').appendChild(o)}
  opt($('target'),info.targets);opt($('sby'),info.by);opt($('qby'),['묶지 않음'].concat(info.by));
  $('target').value='광고주만';$('sby').value='광고주별';}catch(e){}
 $('kw').addEventListener('keydown',e=>{if(e.key==='Enter')run()});
 $('kref').addEventListener('keydown',e=>{if(e.key==='Enter')run()});
 $('ckw').addEventListener('keydown',e=>{if(e.key==='Enter')cal()});$('cgrp').addEventListener('keydown',e=>{if(e.key==='Enter')cal()});
 const Y=+today.slice(0,4);for(let y=Y-3;y<=Y+1;y++){const o=document.createElement('option');o.value=o.textContent=y;$('cy').appendChild(o)}
 for(let m=1;m<=12;m++){const o=document.createElement('option');o.value=pad(m);o.textContent=m+'월';$('cm').appendChild(o)}
 $('cy').onchange=$('cm').onchange=()=>cal();$('cy').value=Y;$('cm').value=today.slice(5,7);
 tab('C')})();
</script></body></html>
"""


def probe_view(port, timeout=1.0):
    """이 PC의 port 에서 이미 조회 페이지가 돌고 있으면 그 정보(dict), 아니면 None. (런처·편집기가 서로 확인)"""
    try:
        import urllib.request
        with urllib.request.urlopen(f"http://127.0.0.1:{int(port)}/api/info", timeout=timeout) as r:
            d = json.loads(r.read().decode("utf-8"))
        return d if isinstance(d, dict) and d.get("app") == "runsheet-view" else None
    except Exception:
        return None


def load_view_settings(data_dir):
    """편집기 설정에서 조회 관련 값 (켜짐, 포트, DB 위치). 런처도 이걸 읽어 조회 페이지를 연다."""
    d = {}
    try:
        with open(os.path.join(data_dir, "settings.json"), "r", encoding="utf-8") as f:
            d = json.load(f) or {}
    except Exception:
        pass
    try:
        port = min(65535, max(1024, int(d.get("view_port", VIEW_PORT_DEFAULT))))
    except (TypeError, ValueError):
        port = VIEW_PORT_DEFAULT
    return bool(d.get("view_on", False)), port, d.get("db_path") or os.path.join(data_dir, "pdf_data_editor.db")


class ViewServer:
    """읽기 전용 조회 서버. start() 로 켜고 stop() 으로 끈다.
    같은 포트를 런처(또는 편집기)가 이미 열어 두었으면 그것을 쓴다 (external)."""

    def __init__(self, get_store, host_kind="편집기", peers_fn=None, announce=True):
        self.get_store = get_store
        self.host_kind = host_kind
        self.peers_fn = peers_fn          # 런처가 사내망에서 찾은 TV 편성표 조회 PC 목록
        self.announce = announce          # 사내망에 '운행표 조회 있음' 알리기 (녹음 PC 서버의 사본은 알리지 않는다)
        self.httpd = None
        self.port = None
        self.error = ""
        self.external = None          # 다른 쪽(런처/편집기)이 열어 둔 조회 페이지 정보

    @property
    def running(self):
        return self.httpd is not None or self.external is not None

    def start(self, port):
        self.stop()
        self.error = ""
        try:
            import threading
            from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
            from urllib.parse import urlparse, parse_qs
        except Exception as exc:
            self.error = f"이 실행 환경에는 웹 서버 부품이 없습니다. 편집 프로그램 런처를 업데이트해 주세요.\n({exc})"
            return False
        owner = self

        class Handler(BaseHTTPRequestHandler):
            server_version = "RunSheetView/1"

            def log_message(self, *a):
                pass

            def _send(self, code, body, ctype, extra=None):
                data = body.encode("utf-8") if isinstance(body, str) else body
                self.send_response(code)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                for k, v in (extra or {}).items():
                    self.send_header(k, v)
                self.end_headers()
                if self.command != "HEAD":
                    self.wfile.write(data)

            def _json(self, obj, code=200):
                self._send(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

            def _csv(self, header, rows, fname):
                import io
                from urllib.parse import quote
                buf = io.StringIO()
                w = csv.writer(buf)
                w.writerow(header)
                w.writerows(rows)
                return self._send(200, "\ufeff" + buf.getvalue(), "text/csv; charset=utf-8",
                                  {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"})

            def do_HEAD(self):
                self.do_GET()

            def do_GET(self):
                try:
                    u = urlparse(self.path)
                    p = {k: v[0] for k, v in parse_qs(u.query).items()}
                    route = u.path.rstrip("/") or "/"
                    if route == "/":
                        return self._send(200, _VIEW_HTML, "text/html; charset=utf-8")
                    if route in ("/api/peers", "/api/runsheets") or route.startswith("/runsheet/"):
                        store = None
                    else:
                        store = owner.get_store()
                    if store is None and not (route in ("/api/peers", "/api/runsheets") or route.startswith("/runsheet/")):
                        return self._json({"error": "편집기의 데이터 파일을 열 수 없습니다."}, 503)
                    if route == "/api/peers":
                        try:
                            peers = owner.peers_fn() if owner.peers_fn else []
                        except Exception:
                            peers = []
                        return self._json({"rows": peers, "auto": owner.peers_fn is not None})
                    if route == "/api/runsheets":
                        return self._json({"rows": list_previews()})
                    if route.startswith("/runsheet/") and route.endswith(".pdf"):
                        key = route[len("/runsheet/"):-4]
                        path = os.path.join(PREVIEW_DIR, key + ".pdf")
                        if not _PREVIEW_KEY_RE.match(key) or not os.path.exists(path):
                            return self._json({"error": "미리보기가 없습니다."}, 404)
                        with open(path, "rb") as f:
                            data = f.read()
                        return self._send(200, data, "application/pdf", {"Content-Disposition": f'inline; filename="{key}.pdf"'})
                    if route == "/api/snapshot":                       # 녹음 PC 서버(사본)가 가져가는 데이터 사본 (읽기 전용)
                        return self._snapshot(store, p.get("since", ""))
                    if route == "/api/info":
                        return self._json({"app": "runsheet-view", "host": owner.host_kind,
                                           "name": computer_name(), "version": APP_VERSION, "groups": store.distinct_groups(),
                                           "by": list(BY_OPTIONS), "targets": list(SEARCH_TARGETS)})
                    if route == "/api/calendar":                       # 📅 달력: 한 달의 날짜별 건수
                        ym = p.get("ym", "")
                        if not re.fullmatch(r"\d{4}-\d{2}", ym):
                            return self._json({"error": "연-월을 확인해 주세요."}, 400)
                        media = p.get("media", "") if p.get("media", "") in ("", "AM", "FM") else ""
                        kw, grp = p.get("kw", "").strip()[:100], p.get("grp", "").strip()[:100]
                        target = SEARCH_TARGETS.get(p.get("target", ""), "name")
                        days = store.calendar_counts(ym + "-01", ym + "-31", kw, grp, media, target)
                        status = {}
                        for x in list_previews():
                            if x["air_date"][:7] == ym:
                                status.setdefault(x["air_date"], []).append({"media": x["media"], "key": x["key"], "has_pdf": x["has_pdf"],
                                                                             "state": "편집중" if x["dirty"] or x["state"] != "수정완료" else "수정완료"})
                        return self._json({"days": days, "status": status})
                    if route in ("/api/day", "/api/day.csv"):         # 📅 달력: 날짜를 누르면 나오는 그날 내역
                        day = p.get("date", "")
                        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day):
                            return self._json({"error": "날짜를 확인해 주세요."}, 400)
                        media = p.get("media", "") if p.get("media", "") in ("", "AM", "FM") else ""
                        kw, grp = p.get("kw", "").strip()[:100], p.get("grp", "").strip()[:100]
                        target = SEARCH_TARGETS.get(p.get("target", ""), "name")
                        agg = store.day_by_name(day, kw, grp, media, target)
                        rows = sorted(((t, m, pg, n, sec) for a, m, t, pg, c_, n, sec, pr, g in store.detail(day, day, kw, grp, media=media, target=target)),
                                      key=lambda r: (r[0], r[1]))
                        if route == "/api/day.csv":
                            if p.get("part") == "names":
                                return self._csv(["광고주", "그룹", "횟수", "총 초수"], agg, f"{day}_광고주별횟수.csv")
                            return self._csv(["시분초", "매체", "프로그램", "광고주", "초수"], rows, f"{day}_운행내역.csv")
                        pv = [{"media": x["media"], "key": x["key"], "has_pdf": x["has_pdf"],
                               "state": "편집중" if x["dirty"] or x["state"] != "수정완료" else "수정완료"}
                              for x in list_previews() if x["air_date"] == day]
                        return self._json({"names": agg, "rows": rows, "previews": pv})
                    if route in ("/api/settle", "/api/settle.csv"):      # 편집기 「일/월/년 정산」 탭과 같은 조회
                        kind = p.get("kind", "")
                        if kind not in ("일", "월", "년"):
                            return self._json({"error": "정산 종류를 확인해 주세요."}, 400)
                        try:
                            d0, d1 = period_range(kind, p.get("ref", ""))
                        except ValueError as exc:
                            return self._json({"error": str(exc)}, 400)
                        by_lbl = p.get("by", "") if p.get("by", "") in BY_OPTIONS else "광고주별"
                        media = p.get("media", "") if p.get("media", "") in ("", "AM", "FM") else ""
                        data = store.aggregate(d0, d1, by=BY_OPTIONS[by_lbl], media=media)
                        rows = [(k or "(없음)", cnt, sec) for k, cnt, sec, amt in data]
                        if route == "/api/settle":
                            return self._json({"from": d0, "to": d1, "rows": rows})
                        return self._csv([by_lbl.replace("별", ""), "건수", "총 초수"], rows, f"{kind}정산_{p.get('ref', '').strip()}_{by_lbl}.csv")
                    if route in ("/api/search", "/api/search.csv"):      # 편집기 「검색 / 그룹 검색」 탭과 같은 조회
                        d0, d1 = p.get("from", ""), p.get("to", "")
                        for d in (d0, d1):
                            if d and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
                                return self._json({"error": "기간은 2026-10-30 형식으로 입력해 주세요."}, 400)
                        media = p.get("media", "") if p.get("media", "") in ("", "AM", "FM") else ""
                        kw, grp = p.get("kw", "").strip()[:100], p.get("grp", "").strip()[:100]
                        target = SEARCH_TARGETS.get(p.get("target", ""), "name")
                        by_lbl = p.get("by", "")
                        fname = f"검색_{kw or '전체'}_{d0 or '처음'}_{d1 or '끝'}.csv"
                        if by_lbl in BY_OPTIONS:
                            data = store.aggregate(d0, d1, kw, grp, BY_OPTIONS[by_lbl], media=media, target=target)
                            rows = [(k or "(없음)", cnt, sec) for k, cnt, sec, amt in data]
                            if route == "/api/search":
                                return self._json({"mode": "agg", "rows": rows})
                            return self._csv([by_lbl.replace("별", ""), "건수", "총 초수"], rows, fname)
                        data = store.detail(d0, d1, kw, grp, media=media, target=target)
                        rows = [(a, m, t, pg, cell_label(c_), n, s, g) for a, m, t, pg, c_, n, s, pr, g in data]
                        if route == "/api/search":
                            return self._json({"mode": "detail", "total": len(rows), "sec": sum(r[6] or 0 for r in rows),
                                               "rows": rows[:VIEW_DETAIL_LIMIT]})
                        return self._csv(["방송일", "매체", "시분초", "프로그램", "칸", "광고주", "초수", "그룹"], rows, fname)
                    if route in ("/api/aggregate", "/api/detail", "/api/detail.csv"):
                        d0, d1 = p.get("from", ""), p.get("to", "")
                        for d in (d0, d1):
                            if d and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
                                return self._json({"error": "기간을 확인해 주세요."}, 400)
                        media = p.get("media", "") if p.get("media", "") in ("", "AM", "FM") else ""
                        kw, grp = p.get("kw", "").strip()[:100], p.get("grp", "")[:100]
                        target = SEARCH_TARGETS.get(p.get("target", ""), "both")
                        if route == "/api/aggregate":
                            by = BY_OPTIONS.get(p.get("by", ""), "name")
                            return self._json({"rows": store.aggregate(d0, d1, kw, grp, by=by, media=media, target=target)})
                        rows = store.detail(d0, d1, kw, grp, media=media, target=target)
                        if route == "/api/detail":
                            return self._json({"total": len(rows), "rows": rows[:VIEW_DETAIL_LIMIT]})
                        import io
                        buf = io.StringIO()
                        w = csv.writer(buf)
                        w.writerow(["방송일", "매체", "시각", "프로그램", "칸", "광고주", "초", "금액", "그룹"])
                        w.writerows(rows)
                        fname = f"정산상세_{d0 or '처음'}_{d1 or '끝'}.csv"
                        from urllib.parse import quote
                        return self._send(200, "﻿" + buf.getvalue(), "text/csv; charset=utf-8",
                                          {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"})
                    return self._json({"error": "없는 주소입니다."}, 404)
                except Exception as exc:
                    try:
                        self._json({"error": f"조회 중 오류: {exc}"}, 500)
                    except Exception:
                        pass

            def _snapshot(self, store, since):
                """데이터 파일(.db)을 SQLite 백업 기능으로 안전하게 복사해 보낸다. 바뀐 것이 없으면 304."""
                import tempfile
                db = getattr(store, "path", "") or ""
                if not db or not os.path.exists(db):
                    return self._json({"error": "데이터 파일이 없습니다."}, 404)
                st = os.stat(db)
                tag = f"{st.st_mtime_ns}-{st.st_size}"
                if since and since == tag:
                    self.send_response(304)
                    self.send_header("X-Data-Tag", tag)
                    self.end_headers()
                    return
                fd, tmp = tempfile.mkstemp(suffix=".db")
                os.close(fd)
                try:
                    src = sqlite3.connect(db)
                    dst = sqlite3.connect(tmp)
                    try:
                        src.backup(dst)
                    finally:
                        dst.close()
                        src.close()
                    with open(tmp, "rb") as f:
                        data = f.read()
                finally:
                    try:
                        os.remove(tmp)
                    except OSError:
                        pass
                return self._send(200, data, "application/octet-stream",
                                  {"X-Data-Tag": tag, "X-App-Version": APP_VERSION, "X-Source-Name": quote_hdr(computer_name())})

            def _deny(self):
                self._json({"error": "조회 전용입니다."}, 405)

            do_POST = do_PUT = do_DELETE = do_PATCH = _deny

        info = probe_view(port, timeout=0.5)
        if info:                                   # 이미 이 PC에서 조회 페이지가 열려 있음 (런처 또는 편집기)
            self.external, self.port = info, int(port)
            return True

        class ExclusiveServer(ThreadingHTTPServer):
            # 윈도우의 SO_REUSEADDR 은 같은 포트를 두 프로그램이 함께 열게 해 버린다 -> 윈도우에서는 끈다
            # (그러면 런처와 편집기 중 나중에 여는 쪽은 실패하고, 먼저 연 쪽 것을 그대로 쓴다)
            allow_reuse_address = os.name != "nt"

        try:
            httpd = ExclusiveServer(("0.0.0.0", int(port)), Handler)
        except OSError as exc:
            info = probe_view(port)
            if info:                               # 런처(또는 다른 편집기 창)가 이미 이 PC에서 조회 페이지를 열어 둠 -> 그대로 쓴다
                self.external, self.port = info, int(port)
                return True
            self.error = f"{port}번 포트를 열 수 없습니다. 다른 프로그램이 쓰고 있을 수 있습니다. 다른 포트 번호로 바꿔 보세요.\n({exc})"
            return False
        httpd.daemon_threads = True
        self.httpd, self.port = httpd, int(port)
        threading.Thread(target=httpd.serve_forever, kwargs={"poll_interval": 0.5}, daemon=True).start()
        if self.announce:
            def _announce_loop(h=httpd, p=int(port)):
                while self.httpd is h:                       # 이 서버가 켜져 있는 동안 5초마다 알린다
                    announce_view("radio", p)
                    time.sleep(5)
            threading.Thread(target=_announce_loop, daemon=True).start()
        return True

    def stop(self):
        self.external = None
        h, self.httpd = self.httpd, None
        if h is not None:
            try:
                h.shutdown()
                h.server_close()
            except Exception:
                pass

    def urls(self):
        if not self.running:
            return []
        return [f"http://{computer_name()}:{self.port}"] + [f"http://{ip}:{self.port}" for ip in local_addresses()]



def flow_layout(frame, pad_x=8, pad_y=4):
    """툴바의 버튼들을 창 너비에 맞게 자동으로 줄 바꿈 (넓으면 1줄, 좁으면 2줄 이상).
    버튼 크기·글씨는 그대로, 순서도 그대로 (오른쪽에 붙여 둔 것은 맨 뒤)."""
    try:
        items = [(w, w.pack_info()) for w in frame.pack_slaves()]
    except Exception:
        return
    order = [it for it in items if it[1].get("side") != "right"] + [it for it in items if it[1].get("side") == "right"]

    def _pad(v):
        try:
            if isinstance(v, (tuple, list)):
                return int(float(v[0])), int(float(v[-1]))
            p = [int(float(x)) for x in str(v).split()]
            return (p[0], p[-1]) if p else (0, 0)
        except Exception:
            return 0, 0
    lay = []
    for w, info in order:
        w.pack_forget()
        lay.append((w, _pad(info.get("padx", 0)), info.get("fill") in ("y", "both")))
    frame._flow_last_w = None

    def reflow(_e=None):
        W = frame.winfo_width()
        if W <= 1:
            W = frame.winfo_toplevel().winfo_width()
        if W <= 1 or W == frame._flow_last_w:
            return
        frame._flow_last_w = W
        x, y, rowh, rows = pad_x, pad_y, 0, [[]]
        for w, (pl, pr), fill_y in lay:
            if not w.winfo_exists():
                continue
            ww = w.winfo_reqwidth() + pl + pr
            hh = w.winfo_reqheight() + 4
            if x > pad_x and x + ww > W - pad_x:            # 자리가 모자라면 다음 줄로
                x, y, rowh = pad_x, y + rowh, 0
                rows.append([])
            rows[-1].append((w, x + pl, y, fill_y))
            x += ww
            rowh = max(rowh, hh)
        top = pad_y
        for row in rows:
            if not row:
                continue
            rh = max(w_.winfo_reqheight() for w_, *_ in row) + 4
            for w_, xx, _yy, fill_y in row:
                if fill_y:
                    w_.place(x=xx, y=top + 2, height=rh - 4)
                else:
                    w_.place(x=xx, y=top + (rh - w_.winfo_reqheight()) // 2)
            top += rh
        frame.configure(height=top + pad_y)

    frame.bind("<Configure>", reflow)
    frame.after_idle(reflow)
    frame.after(150, lambda: (setattr(frame, "_flow_last_w", None), reflow()))


class DataEditor:
    # ---- 탭(운행표 여러 장) --------------------------------------------------
    # 지금 보고 있는 탭의 상태는 self 에 그대로 두고(기존 함수들이 그대로 동작),
    # 다른 탭으로 넘어갈 때만 아래 항목들을 탭 묶음(dict)에 넣었다 꺼낸다.
    @staticmethod
    def _blank_state():
        return {"doc": None, "src_bytes": None, "pdf_path": None, "page_count": 0, "cur": 0, "zoom": 1.0,
                "page_origin": (0.0, 0.0), "page_size": (595.0, 842.0), "rows": {}, "cells": {}, "cell_by_key": {},
                "entries": {}, "orphans": [], "removed": [], "history": [], "selected_key": None, "air_date": "", "media": "",
                "_dirty": False, "cap_flag": set(), "_bg_key": None, "_bg_photo": None, "_coal": None}

    @property
    def dirty(self):
        return getattr(self, "_dirty", False)

    @dirty.setter
    def dirty(self, v):
        was = getattr(self, "_dirty", False)
        self._dirty = bool(v)
        if bool(v) and not was and getattr(self, "doc", None) is not None:
            try:
                self._set_preview_dirty(True)          # 다른 PC 조회에 '현재 편집중' 알림
            except Exception:
                pass
        if getattr(self, "tabs", None):
            self._refresh_tab_titles()

    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1480x1020")
        root.minsize(1000, 700)
        self.tabs = []              # 열린 운행표 탭들 (dict). 지금 탭의 상태는 self 에 있다
        self.active_tab = None
        self._pdf_ui = False
        self.page_total_var = tk.StringVar(value="/ 0")
        self.file_var = tk.StringVar(value="")

        self.doc = None
        self.src_bytes = None
        self.pdf_path = None
        self.page_count = 0
        self.cur = 0
        self.zoom = 1.0
        self.zoom_label = tk.StringVar(value="100%")
        self.page_origin = (0.0, 0.0)
        self.page_size = (595.0, 842.0)

        self.rows = {}              # 페이지 -> 행 목록
        self.cells = {}             # 페이지 -> 칸 목록
        self.cell_by_key = {}
        self.caps = {}              # (시각, 프로그램명) -> 사용자가 설정한 재원
        self.entries = {}           # 칸 key -> 추가한 데이터
        self.orphans = []           # 기록은 있는데 현재 PDF에서 자리를 못 찾은 것 (저장 때 그대로 보존)
        self.history = []
        self.selected_key = None
        self.air_date = ""
        self.media = ""             # 'AM' / 'FM'
        self.media_var = tk.StringVar(value="")
        self.count_var = tk.StringVar(value="")
        self.dirty = False
        self._coal = None

        self._bg_key = None
        self._bg_photo = None
        self.show_cells = tk.BooleanVar(value=True)
        self.page_var = tk.StringVar(value="1")
        self.date_var = tk.StringVar(value="")
        self.filter_var = tk.StringVar(value="")
        self.cell_info_var = tk.StringVar(value="PDF에서 칸을 클릭하세요.")

        self.sort_key = "name"
        self.sort_desc = False
        self.show_caps = tk.BooleanVar(value=True)
        self.preview_var = tk.BooleanVar(value=True)
        self.highlight_var = tk.BooleanVar(value=False)   # 저장·출력 PDF에도 추가한 칸을 형광펜처럼 강조
        self.show_today = tk.BooleanVar(value=True)       # 화면에만: 오늘 넣거나 고친 칸(초록)·오늘 지운 칸(회색 취소선)
        self.preset_cols = self._default_preset_cols()   # 미리 만든 데이터 목록의 열 순서·숨김·너비
        self.mark_opts = self._default_mark_opts()        # 이름표·강조 표시 설정
        self.cap_flag = set()       # 원본 재원이 설정 때와 달라진 행 (시각, 프로그램명)
        self._mutex = hold_app_mutex()
        self.settings_path = SETTINGS_FILE
        self.font_path = get_system_korean_font_path()
        self.db_path = DB_FILE
        self.backup_dir = ""         # 비우면 C:\RADIO\백업
        self.backup_keep = 30
        self.backup_on = True
        self.update_dir = UPDATE_DIR_DEFAULT
        self.update_check = True     # 켤 때 새 버전 확인
        self.page_check_on = True     # 다음 장으로 넘어갈 때 시간 초과·부족 행 알림
        self.preview_view = None      # 미리보기 마지막 보기: 'fit' / 'width' / 배율(숫자), None 이면 편집 화면 배율
        self.preview_geom = ""
        self.preset_detached = False  # '미리 만든 데이터'를 따로 창으로 띄워 둠
        self.preset_win_geom = ""
        self.preset_win_top = False
        self.preset_win = None
        self.view_on = False         # 다른 PC에서 정산/검색 조회 허용 (이 편집기에서만)
        self.view_port = VIEW_PORT_DEFAULT
        self.view_server = ViewServer(lambda: self.store, "편집기")
        startup_note = prepare_data_folder(lambda t, m: messagebox.askyesno(t, m))
        self._load_settings()
        self.ui_font_family = get_tk_font_family(self.font_path)
        self.store = None
        self._open_store()

        self._build_ui()
        self.draw_blank_canvas()
        self.root.after(400, self.alert_expired)
        if startup_note:
            self.root.after(300, lambda: messagebox.showinfo("데이터 준비", startup_note))
        self.startup_backup()
        if self.update_check and not RUN_BY_LAUNCHER:
            self.root.after(1500, lambda: self.check_update(quiet=True))
        if self.view_on:
            self.root.after(800, lambda: self.apply_view_server(quiet=True))
        self.root.after(15000, self._view_watch)

    def _view_watch(self):
        """런처가 조회 페이지를 맡고 있다가 꺼지면 편집기가 이어받는다 (15초마다 확인)."""
        try:
            vs = self.view_server
            if self.view_on and vs.external is not None and not probe_view(vs.port):
                vs.external = None
                self.apply_view_server(quiet=True)
        except Exception:
            pass
        self.root.after(15000, self._view_watch)

    # ------------------------------------------------------------------ 다른 PC 조회
    def _app_title(self):
        return APP_TITLE + ("   · 🌐 다른 PC 조회 허용 중" if self.view_server.running else "")

    def apply_view_server(self, quiet=False, parent=None):
        """설정대로 조회 서버를 켜거나 끈다."""
        was = self.view_server.running
        self._apply_view_server(quiet, parent)
        if self.view_server.running != was:
            self._rebuild_toolbar()
        if self.view_server.running:
            self._write_view_address_file()
            if not was and not quiet:                 # 방금 켰으면 접속 주소를 바로 알려 준다
                self.root.after(150, lambda: self.show_view_addresses(parent if parent is not None and parent.winfo_exists() else None))

    def _apply_view_server(self, quiet=False, parent=None):
        if self.view_on:
            if not (self.view_server.running and self.view_server.port == self.view_port):
                if not self.view_server.start(self.view_port) and not quiet:
                    messagebox.showwarning("다른 PC 조회", self.view_server.error, parent=parent or self.root)
                elif not self.view_server.running and quiet:
                    self.root.after(0, lambda: messagebox.showwarning("다른 PC 조회", "다른 PC 조회를 켜지 못했습니다.\n" + self.view_server.error))
        else:
            self.view_server.stop()
        try:
            if getattr(self, "active_tab", None) is not None and self.doc:
                self.root.title(f"{self._app_title()}  —  {self._tab_title(self.media, self.air_date)}{' *' if self.dirty else ''}")
            else:
                self.root.title(self._app_title())
        except Exception:
            pass

    # ------------------------------------------------------------------ 설정 / DB
    def _load_settings(self):
        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("font_path") and os.path.exists(data["font_path"]):
                self.font_path = data["font_path"]
            if data.get("db_path"):
                self.db_path = data["db_path"]
            self.backup_dir = data.get("backup_dir", "") or ""
            try:
                self.backup_keep = max(1, int(data.get("backup_keep", 30)))
            except (TypeError, ValueError):
                self.backup_keep = 30
            self.backup_on = bool(data.get("backup_on", True))
            if data.get("sort_key") in ("name", "sec", "end"):
                self.sort_key = data["sort_key"]
            self.sort_desc = bool(data.get("sort_desc", False))
            self.preview_var.set(bool(data.get("preview_before_save", True)))
            self.highlight_var.set(bool(data.get("highlight_output", False)))
            self.show_today.set(bool(data.get("show_today", True)))
            self.preset_cols = self._clean_preset_cols(data.get("preset_cols"))
            self.mark_opts = self._clean_mark_opts(data.get("mark_opts"))
            self.update_dir = data.get("update_dir") or UPDATE_DIR_DEFAULT
            self.update_check = bool(data.get("update_check", True))
            self.view_on = bool(data.get("view_on", False))
            self.preset_detached = bool(data.get("preset_detached", False))
            self.page_check_on = bool(data.get("page_check_on", True))
            pv_ = data.get("preview_view")
            self.preview_view = pv_ if pv_ in ("fit", "width") or isinstance(pv_, (int, float)) else None
            self.preview_geom = str(data.get("preview_geom", "") or "")
            self.preset_win_geom = str(data.get("preset_win_geom", "") or "")
            self.preset_win_top = bool(data.get("preset_win_top", False))
            try:
                self.view_port = min(65535, max(1024, int(data.get("view_port", VIEW_PORT_DEFAULT))))
            except (TypeError, ValueError):
                self.view_port = VIEW_PORT_DEFAULT
        except Exception:
            pass

    def _save_settings(self):
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump({"font_path": self.font_path, "db_path": self.db_path, "sort_key": self.sort_key, "sort_desc": self.sort_desc, "preview_before_save": self.preview_var.get(), "highlight_output": self.highlight_var.get(), "show_today": self.show_today.get(), "preset_cols": self.preset_cols, "mark_opts": self.mark_opts,
                           "backup_dir": self.backup_dir, "backup_keep": self.backup_keep, "backup_on": self.backup_on,
                           "update_dir": self.update_dir, "update_check": self.update_check, "view_on": self.view_on, "view_port": self.view_port,
                           "preset_detached": self.preset_detached, "page_check_on": self.page_check_on, "preview_view": self.preview_view, "preview_geom": self.preview_geom, "preset_win_geom": self.preset_win_geom, "preset_win_top": self.preset_win_top,
                           "app_version": APP_VERSION}, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _open_store(self):
        try:
            self.store = Store(self.db_path)
        except Exception as exc:
            self.store = None
            messagebox.showerror("데이터베이스 열기 실패", f"{self.db_path}\n\n{exc}")

    def effective_backup_dir(self):
        return self.backup_dir or BACKUP_DIR

    def run_backup(self, quiet=True):
        """백업 실행. 성공하면 파일 경로, 실패하면 None (quiet=False 이면 안내창)."""
        if not self.store:
            return None
        try:
            path = self.store.backup(self.effective_backup_dir(), self.backup_keep)
            return path
        except Exception as exc:
            if not quiet:
                messagebox.showerror("백업 실패", f"{self.effective_backup_dir()}\n\n{exc}")
            else:
                self.backup_error = str(exc)
            return None

    def startup_backup(self):
        """프로그램을 켤 때, 오늘 백업이 아직 없으면 한 번 만든다."""
        if not self.backup_on or not self.store or not os.path.exists(self.db_path):
            return
        stem = os.path.splitext(os.path.basename(self.db_path))[0]
        today_file = os.path.join(self.effective_backup_dir(), f"{stem}_{datetime.date.today().strftime('%Y%m%d')}.db")
        if not os.path.exists(today_file):
            self.run_backup()

    def open_settings_window(self):
        win = tk.Toplevel(self.root)
        win.title("⚙️ 설정")
        win.geometry("660x900")
        win.transient(self.root)
        win.grab_set()
        pad = {"padx": 15, "pady": 4}

        ttk.Label(win, text="추가하는 글자에 쓸 한글 글꼴 (.ttf / .ttc / .otf)", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(14, 2))
        r1 = ttk.Frame(win)
        r1.pack(fill="x", **pad)
        fvar = tk.StringVar(value=self.font_path or "")
        ttk.Entry(r1, textvariable=fvar).pack(side="left", fill="x", expand=True, padx=(0, 5))

        def browse_font():
            p = filedialog.askopenfilename(title="글꼴 파일 선택", filetypes=[("Font Files", "*.ttf;*.ttc;*.otf"), ("모든 파일", "*.*")], parent=win)
            if p:
                fvar.set(p)

        ttk.Button(r1, text="찾아보기", command=browse_font).pack(side="right")

        ttk.Label(win, text="입력 기록·미리 만든 데이터 저장 파일 (.db)", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(12, 2))
        r2 = ttk.Frame(win)
        r2.pack(fill="x", **pad)
        dvar = tk.StringVar(value=self.db_path)
        ttk.Entry(r2, textvariable=dvar).pack(side="left", fill="x", expand=True, padx=(0, 5))

        def browse_db():
            p = filedialog.asksaveasfilename(title="데이터 파일 위치", defaultextension=".db", initialfile=os.path.basename(dvar.get() or "pdf_data_editor.db"),
                                             filetypes=[("데이터 파일", "*.db"), ("모든 파일", "*.*")], parent=win, confirmoverwrite=False)
            if p:
                dvar.set(p)

        ttk.Button(r2, text="위치 변경", command=browse_db).pack(side="right")
        ttk.Label(win, text="※ 네트워크 드라이브나 공유 폴더에 두면 여러 PC에서 같은 기록을 쓸 수 있지만, 동시에 저장하면 충돌할 수 있습니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=560).pack(anchor="w", padx=15)

        ttk.Label(win, text="자동 백업", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(14, 2))
        bon = tk.BooleanVar(value=self.backup_on)
        ttk.Checkbutton(win, text="프로그램을 켤 때와 끌 때 자동으로 백업 (하루에 파일 1개, 날짜별)", variable=bon).pack(anchor="w", padx=15)
        r3 = ttk.Frame(win)
        r3.pack(fill="x", **pad)
        bvar = tk.StringVar(value=self.backup_dir)
        ttk.Entry(r3, textvariable=bvar).pack(side="left", fill="x", expand=True, padx=(0, 5))

        def browse_bk():
            d = filedialog.askdirectory(title="백업 폴더 선택", parent=win)
            if d:
                bvar.set(d)

        ttk.Button(r3, text="폴더 선택", command=browse_bk).pack(side="right")
        r4 = ttk.Frame(win)
        r4.pack(fill="x", padx=15)
        ttk.Label(r4, text="보관 개수(최근 N일치):").pack(side="left")
        kvar = tk.StringVar(value=str(self.backup_keep))
        ttk.Spinbox(r4, from_=1, to=3650, width=6, textvariable=kvar).pack(side="left", padx=6)

        def backup_now():
            old = (self.backup_dir, self.backup_keep)
            self.backup_dir = bvar.get().strip()
            try:
                self.backup_keep = max(1, int(kvar.get()))
            except ValueError:
                pass
            path = self.run_backup(quiet=False)
            if path:
                messagebox.showinfo("백업 완료", f"백업했습니다:\n{path}", parent=win)
            else:
                self.backup_dir, self.backup_keep = old

        ttk.Button(r4, text="지금 백업", command=backup_now).pack(side="left", padx=12)
        ttk.Label(win, text=f"※ 백업 폴더를 비워 두면 {BACKUP_DIR} 을 씁니다. OneDrive·구글드라이브 폴더로 지정하면 다른 곳에도 보관됩니다.\n   복원: 백업 파일을 복사해 데이터 파일(.db) 이름으로 바꾸면 됩니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=560, justify="left").pack(anchor="w", padx=15, pady=(4, 0))

        ttk.Label(win, text="업데이트" + ("  (편집 프로그램 런처에서 관리합니다)" if RUN_BY_LAUNCHER else ""), font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(14, 2))
        r5 = ttk.Frame(win)
        r5.pack(fill="x", **pad)
        uvar = tk.StringVar(value=self.update_dir)
        ttk.Entry(r5, textvariable=uvar).pack(side="left", fill="x", expand=True, padx=(0, 5))

        def browse_up():
            d = filedialog.askdirectory(title="업데이트 폴더 선택 (version.json 과 설치 파일이 있는 곳)", parent=win)
            if d:
                uvar.set(d)

        ttk.Button(r5, text="폴더 선택", command=browse_up).pack(side="right")
        r6 = ttk.Frame(win)
        r6.pack(fill="x", padx=15)
        ucheck = tk.BooleanVar(value=self.update_check)
        ttk.Checkbutton(r6, text="프로그램을 켤 때 새 버전 확인", variable=ucheck).pack(side="left")

        def check_now():
            self.update_dir = uvar.get().strip() or UPDATE_DIR_DEFAULT
            self._save_settings()
            self.check_update(quiet=False, parent=win)

        ttk.Button(r6, text="지금 확인", command=check_now).pack(side="left", padx=12)
        ttk.Label(win, text="※ 여러 PC가 같이 쓰려면 공유 폴더(예: \\\\서버\\RADIO\\업데이트)로 지정하세요. 그 폴더에 새 버전을 올리면 각 PC가 켤 때 알려 줍니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=600, justify="left").pack(anchor="w", padx=15, pady=(2, 0))

        pc_var = tk.BooleanVar(value=self.page_check_on)

        def _pc():
            self.page_check_on = bool(pc_var.get())
            self._save_settings()
        ttk.Checkbutton(win, text="다음 장으로 넘어갈 때 Spot·캠페인 시각 행의 시간 초과·부족 알려 주기 (예: 1329(+20초), 0859(-10초))",
                        variable=pc_var, command=_pc).pack(anchor="w", padx=15, pady=(10, 0))

        ttk.Label(win, text="다른 PC에서 정산 / 검색 조회", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(14, 2))
        rv = ttk.Frame(win)
        rv.pack(fill="x", padx=15)
        von = tk.BooleanVar(value=self.view_on)
        ttk.Checkbutton(rv, text="이 편집기의 정산/검색을 사내 다른 PC에서 볼 수 있게 허용 (조회 전용)", variable=von).pack(side="left")
        rv2 = ttk.Frame(win)
        rv2.pack(fill="x", padx=15, pady=(2, 0))
        ttk.Label(rv2, text="포트 번호:").pack(side="left")
        vport = tk.StringVar(value=str(self.view_port))
        ttk.Spinbox(rv2, from_=1024, to=65535, width=7, textvariable=vport).pack(side="left", padx=6)
        view_lbl = ttk.Label(win, text="", font=(self.ui_font_family, 9), foreground="#1e3a8a", wraplength=600, justify="left", cursor="hand2")
        view_lbl.pack(anchor="w", padx=15, pady=(2, 0))

        def show_view_status():
            urls = self.view_server.urls()
            if urls:
                view_lbl.config(text="✅ 켜짐 — 다른 PC의 브라우저에서 아래 주소로 들어가면 됩니다 (클릭하면 복사):\n   " + "\n   ".join(urls))
            elif self.view_on and self.view_server.error:
                view_lbl.config(text="⚠ " + self.view_server.error, foreground="#b45309")
            else:
                view_lbl.config(text="꺼짐", foreground="#6b7280")

        def copy_url(_e=None):
            urls = self.view_server.urls()
            if urls:
                win.clipboard_clear()
                win.clipboard_append(urls[-1] if len(urls) > 1 else urls[0])

        def apply_view():
            self.view_on = bool(von.get())
            try:
                self.view_port = min(65535, max(1024, int(vport.get())))
            except ValueError:
                self.view_port = VIEW_PORT_DEFAULT
            self._save_settings()
            self.apply_view_server(parent=win)
            view_lbl.config(foreground="#1e3a8a")
            show_view_status()

        ttk.Button(rv2, text="적용", command=apply_view).pack(side="left", padx=8)
        view_lbl.bind("<Button-1>", copy_url)
        show_view_status()
        ttk.Label(win, text="※ 켠 편집기만 조회 페이지를 엽니다. 편집기가 2대라도 조회를 보여 줄 편집기에서만 켜세요.\n"
                            "   처음 켤 때 Windows 방화벽 창이 뜨면 '허용'(개인/도메인 네트워크)을 눌러 주세요. 편집기를 꺼도 편집 프로그램 런처가 켜져 있으면 런처가 이어서 조회를 엽니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=600, justify="left").pack(anchor="w", padx=15, pady=(2, 0))

        ttk.Label(win, text="기본 데이터 / 프로그램 정보", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=15, pady=(14, 2))
        r7 = ttk.Frame(win)
        r7.pack(fill="x", padx=15)
        ttk.Button(r7, text="기본 데이터 파일 만들기 (다른 PC용)", command=lambda: self.export_seed_db(parent=win)).pack(side="left")
        ttk.Button(r7, text="기본 데이터 가져오기", command=lambda: self.import_seed_db(parent=win)).pack(side="left", padx=6)
        ttk.Label(win, text=f"버전 {APP_VERSION}   ·   프로그램 폴더 {APP_ROOT}", foreground="#555").pack(anchor="w", padx=15, pady=(6, 0))

        def save_close():
            self.update_dir = uvar.get().strip() or UPDATE_DIR_DEFAULT
            self.update_check = bool(ucheck.get())
            self.backup_on = bool(bon.get())
            self.view_on = bool(von.get())
            try:
                self.view_port = min(65535, max(1024, int(vport.get())))
            except ValueError:
                self.view_port = VIEW_PORT_DEFAULT
            self.backup_dir = bvar.get().strip()
            try:
                self.backup_keep = max(1, int(kvar.get()))
            except ValueError:
                self.backup_keep = 30
            fp = fvar.get().strip()
            if fp and os.path.exists(fp):
                self.font_path = fp
                self.ui_font_family = get_tk_font_family(fp)
            new_db = dvar.get().strip()
            if new_db and new_db != self.db_path:
                self.db_path = new_db
                self._open_store()
                self.refresh_presets()
            self._save_settings()
            self.apply_view_server(parent=win)
            self.draw_page()
            win.destroy()

        ttk.Button(win, text="설정 저장", command=save_close).pack(pady=14)

    # ------------------------------------------------------------------ 업데이트 / 기본 데이터
    def check_update(self, quiet=True, parent=None):
        """업데이트 폴더의 version.json 을 보고 새 버전이 있으면 업데이트할지 묻는다.
        quiet=True(켤 때): 폴더가 없거나 최신이면 아무 말도 하지 않는다."""
        owner = parent or self.root
        if RUN_BY_LAUNCHER:
            if not quiet:
                messagebox.showinfo("업데이트 확인", "업데이트는 편집 프로그램(런처)에서 합니다.\n런처 화면에서 [업데이트]를 눌러 주세요.", parent=owner)
            return
        info = read_update_info(self.update_dir)
        if info is None:
            if not quiet:
                messagebox.showinfo("업데이트 확인", f"업데이트 정보를 찾지 못했습니다.\n폴더: {self.update_dir}\n\n"
                                    "(그 폴더에 version.json 과 설치 파일이 있어야 합니다)", parent=owner)
            return
        if version_tuple(info["version"]) <= version_tuple(APP_VERSION):
            if not quiet:
                messagebox.showinfo("업데이트 확인", f"최신 버전입니다.  (지금 {APP_VERSION})", parent=owner)
            return
        if not info["file_path"]:
            if not quiet:
                messagebox.showwarning("업데이트 확인", f"새 버전 {info['version']} 정보는 있는데 설치 파일(zip/exe)이 없습니다:\n{self.update_dir}", parent=owner)
            return
        notes = (info.get("notes") or "").strip()
        msg = (f"새 버전 {info['version']} 이 있습니다.  (지금 {APP_VERSION})\n\n" + (f"바뀐 내용:\n{notes}\n\n" if notes else "") +
               "지금 업데이트할까요?\n"
               "· 업데이트 전에 데이터를 자동으로 백업합니다.\n"
               "· 프로그램이 닫히고, 설치가 끝나면 다시 켜집니다.\n"
               "· 입력 기록·미리 만든 데이터·설정은 그대로 유지됩니다.")
        if messagebox.askyesno("🔄 새 버전", msg, parent=owner):
            self.run_update(info, owner)

    def run_update(self, info, owner=None):
        import sys
        import shutil
        import tempfile
        import subprocess
        owner = owner or self.root
        if not getattr(sys, "frozen", False):
            messagebox.showinfo("업데이트", "지금은 파이썬 파일(.py)로 실행 중이라 자동 설치를 하지 않습니다.\n"
                                f"새 버전 파일:\n{info['file_path']}", parent=owner)
            return
        # 1) 저장하지 않은 탭은 기록을 저장한다 (PDF 파일은 그대로)
        self._commit_date()
        pending = [t for t in self._sorted_tabs() if self._tab_info(t)["dirty"]]
        if pending:
            names = ", ".join(self._tab_title(self._tab_info(t)["media"], self._tab_info(t)["air_date"]) for t in pending)
            if not messagebox.askyesno("업데이트", f"저장하지 않은 탭이 있습니다: {names}\n\n입력 기록을 저장하고 업데이트할까요?\n(아니오: 업데이트 취소)", parent=owner):
                return
            for t in pending:
                self._activate(t)
                if not self.save_records():
                    messagebox.showerror("업데이트", "기록을 저장하지 못해 업데이트를 멈춥니다.", parent=owner)
                    return
                self.dirty = False
        # 2) 백업
        if self.store is not None and self.run_backup(quiet=False) is None:
            if not messagebox.askyesno("업데이트", "백업하지 못했습니다. 그래도 업데이트할까요?", parent=owner):
                return
        # 3) 새 버전을 임시 폴더에 준비한 뒤, 이 프로그램이 닫히면 교체하고 다시 실행한다
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            src = info["file_path"]
            if src.lower().endswith(".zip"):
                # zip: 압축을 풀어 프로그램 폴더를 통째로 바꾼다 (데이터·백업 폴더는 건드리지 않음)
                import zipfile
                work = tempfile.mkdtemp(prefix="RunSheetEditor_update_")
                with zipfile.ZipFile(src) as z:
                    z.extractall(work)
                new_dir = os.path.join(work, "RunSheetEditor")
                if not os.path.exists(os.path.join(new_dir, "RunSheetEditor.exe")):
                    raise RuntimeError("zip 안에 RunSheetEditor\\RunSheetEditor.exe 가 없습니다.")
                prog_dir = os.path.dirname(os.path.abspath(sys.executable))
                ps1 = os.path.join(work, "update.ps1")
                q_new, q_prog = new_dir.replace("'", "''"), prog_dir.replace("'", "''")
                with open(ps1, "w", encoding="utf-8-sig") as f:      # 한글 경로 때문에 BOM 있는 UTF-8 로
                    f.write(f"""Start-Sleep -Seconds 2
$new = '{q_new}'
$prog = '{q_prog}'
robocopy $new $prog /MIR /R:20 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
Start-Process (Join-Path $prog 'RunSheetEditor.exe') -WorkingDirectory $prog
""")
                subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", ps1],
                                 creationflags=flags, close_fds=True)
            else:
                # exe(설치 프로그램): 조용히 설치 -> 끝나면 다시 실행
                dst = os.path.join(tempfile.gettempdir(), os.path.basename(src))
                shutil.copy2(src, dst)
                subprocess.Popen(f'cmd /c timeout /t 2 /nobreak >nul & "{dst}" /SILENT /NORESTART', shell=False, creationflags=flags, close_fds=True)
        except Exception as exc:
            messagebox.showerror("업데이트 실패", f"새 버전을 준비하지 못했습니다:\n{exc}", parent=owner)
            return
        self.root.destroy()

    def import_seed_db(self, parent=None):
        """다른 PC에서 만든 기본데이터.db(또는 백업 파일)의 미리 만든 데이터·재원 설정·그룹을 이 PC에 합친다."""
        owner = parent or self.root
        if self.store is None:
            messagebox.showwarning("기본 데이터 가져오기", "데이터 파일을 열 수 없습니다.", parent=owner)
            return
        path = filedialog.askopenfilename(title="가져올 기본 데이터 파일 (기본데이터.db 또는 백업 .db)", filetypes=[("데이터 파일", "*.db")], parent=owner)
        if not path:
            return
        if not messagebox.askyesno("기본 데이터 가져오기", f"{path}\n\n이 파일의 미리 만든 데이터·재원 설정·그룹 중 이 PC에 없는 것만 더합니다.\n"
                                   "(이미 있는 것은 그대로, 입력 기록은 가져오지 않습니다. 가져오기 전에 자동 백업합니다)\n\n계속할까요?", parent=owner):
            return
        self.run_backup(quiet=True)
        try:
            n = import_base_data(self.db_path, path)
        except Exception as exc:
            messagebox.showerror("기본 데이터 가져오기", f"가져오지 못했습니다:\n{exc}", parent=owner)
            return
        try:
            self.caps = self.store.list_caps()
        except Exception:
            pass
        self.refresh_presets()
        if self.doc:
            self.draw_page()
        messagebox.showinfo("기본 데이터 가져오기", f"새로 더했습니다:\n미리 만든 데이터 {n.get('presets', 0)}개 · 재원 설정 {n.get('cap_settings', 0)}개 · 그룹 {n.get('groups', 0)}개", parent=owner)

    def export_seed_db(self, parent=None):
        """설치 파일에 넣을 '기본데이터.db' 를 만든다. 기본은 입력 기록을 빼고 미리 만든 데이터·재원 설정·그룹만."""
        owner = parent or self.root
        if self.store is None or not os.path.exists(self.db_path):
            messagebox.showwarning("기본 데이터", "데이터 파일을 열 수 없습니다.", parent=owner)
            return
        keep = messagebox.askyesnocancel("기본 데이터 파일 만들기",
                                         "설치용 기본 데이터에는 미리 만든 데이터·재원 설정·그룹이 들어갑니다.\n\n"
                                         "지금까지의 입력 기록(정산 자료)도 같이 넣을까요?\n"
                                         "(예: 기록 포함  /  아니오: 기록 빼고  /  취소)", parent=owner)
        if keep is None:
            return
        path = filedialog.asksaveasfilename(title="기본 데이터 파일 저장 (설치 파일 만드는 폴더에 두세요)", initialfile="기본데이터.db",
                                            defaultextension=".db", filetypes=[("데이터 파일", "*.db")], parent=owner)
        if not path:
            return
        if os.path.abspath(path) == os.path.abspath(self.db_path):
            messagebox.showwarning("기본 데이터", "지금 쓰는 데이터 파일과 다른 이름으로 저장해 주세요.", parent=owner)
            return
        try:
            n = make_seed_db(self.db_path, path, keep_records=bool(keep))
        except Exception as exc:
            messagebox.showerror("기본 데이터", str(exc), parent=owner)
            return
        messagebox.showinfo("기본 데이터", f"만들었습니다:\n{path}\n\n미리 만든 데이터 {n['presets']}개 · 재원 설정 {n['cap_settings']}개 · "
                                         f"그룹 {n['groups']}개 · 입력 기록 {n['records']}건", parent=owner)

    # ------------------------------------------------------------------ UI 뼈대
    def _build_ui(self):
        self.toolbar = ttk.Frame(self.root, padding=(8, 6))
        self.toolbar.pack(side="top", fill="x")
        self._build_default_toolbar()
        self.tabbar = tk.Frame(self.root, background="#cbd5e1")
        self.tabbar.pack(side="top", fill="x")
        self._rebuild_tabbar()

        body = ttk.Panedwindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True)
        board = ttk.Frame(body)
        self.side = ttk.Frame(body, width=390)
        body.add(board, weight=1)
        body.add(self.side, weight=0)
        self._build_sidebar_scroll_wrapper()

        self.canvas = tk.Canvas(board, background="#e5e7eb", highlightthickness=0)
        vbar = ttk.Scrollbar(board, orient="vertical", command=self.canvas.yview)
        hbar = ttk.Scrollbar(board, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        board.rowconfigure(0, weight=1)
        board.columnconfigure(0, weight=1)

        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)

        r = self.root
        r.bind("<Control-z>", lambda e: None if self._typing() else self.undo_action())
        r.bind("<Delete>", lambda e: None if self._typing() else self.clear_selected_cell())
        r.bind("<Prior>", lambda e: None if self._typing() else self.goto_page(self.cur - 1))
        r.bind("<Next>", lambda e: None if self._typing() else self.goto_page(self.cur + 1))
        r.bind("<Left>", lambda e: None if self._typing() else self.move_selection(0, -1))
        r.bind("<Right>", lambda e: None if self._typing() else self.move_selection(0, 1))
        r.bind("<Up>", lambda e: None if self._typing() else self.move_selection(-1, 0))
        r.bind("<Down>", lambda e: None if self._typing() else self.move_selection(1, 0))
        r.bind_all("<Control-Tab>", lambda e: (self.next_tab(1), "break")[1])
        r.bind_all("<Control-Shift-Tab>", lambda e: (self.next_tab(-1), "break")[1])
        try:
            r.bind_all("<Control-ISO_Left_Tab>", lambda e: (self.next_tab(-1), "break")[1])   # 리눅스에서 Ctrl+Shift+Tab
        except tk.TclError:
            pass
        r.bind("<Control-Key-1>", lambda e: self.switch_media("AM") if self.tabs else None)
        r.bind("<Control-Key-2>", lambda e: self.switch_media("FM") if self.tabs else None)
        r.bind("<Control-w>", lambda e: self.close_tab(self.active_tab) if self.active_tab else None)
        r.protocol("WM_DELETE_WINDOW", self.on_closing_app)

        self._build_default_sidebar()

    def _build_sidebar_scroll_wrapper(self):
        self.side_canvas = tk.Canvas(self.side, highlightthickness=0, background="#f1f5f9")
        sb = ttk.Scrollbar(self.side, orient="vertical", command=self.side_canvas.yview)
        self.sidebar_content_frame = ttk.Frame(self.side_canvas, padding=10)
        self.sidebar_content_frame.bind("<Configure>", lambda e: self.side_canvas.configure(scrollregion=self.side_canvas.bbox("all")))
        wid = self.side_canvas.create_window((0, 0), window=self.sidebar_content_frame, anchor="nw")
        self.side_canvas.configure(yscrollcommand=sb.set)
        self.side_canvas.bind("<Configure>", lambda e: self.side_canvas.itemconfig(wid, width=e.width))
        self.side_canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def _wheel(event):
            # 목록(Treeview)·글상자 위에서는 각자 스크롤하게 두고, 그 밖에서만 패널을 스크롤
            if isinstance(event.widget, (ttk.Treeview, tk.Text)):
                return
            if event.delta:
                self.side_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self.side_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.side_canvas.yview_scroll(1, "units")

        self.side_canvas.bind_all("<MouseWheel>", _wheel, add="+")
        self.side_canvas.bind_all("<Button-4>", _wheel, add="+")
        self.side_canvas.bind_all("<Button-5>", _wheel, add="+")

    def _typing(self):
        w = self.root.focus_get()
        return isinstance(w, (tk.Entry, ttk.Entry, tk.Text, ttk.Combobox, ttk.Spinbox, ttk.Treeview))

    def clear_toolbar(self):
        for w in self.toolbar.winfo_children():
            w.destroy()

    def clear_sidebar(self):
        for w in self.sidebar_content_frame.winfo_children():
            w.destroy()

    def _add_view_badge(self, t):
        """다른 PC 조회가 켜져 있으면 툴바 오른쪽 끝에 접속 주소 버튼."""
        if self.view_server.running:
            urls = self.view_server.urls()
            ttk.Button(t, text=f"🌐 조회 주소: {urls[-1] if len(urls) > 1 else urls[0]}",
                       command=self.show_view_addresses).pack(side="right", padx=4)

    def _rebuild_toolbar(self):
        try:
            if getattr(self, "_pdf_ui", False):
                self._build_pdf_toolbar()
            else:
                self._build_default_toolbar()
        except Exception:
            pass

    def show_view_addresses(self, parent=None):
        """다른 PC에서 들어올 주소를 보여 주고 복사·열기. 주소는 데이터 폴더의 '다른PC_조회주소.txt' 에도 남긴다."""
        owner = parent or self.root
        urls = self.view_server.urls()
        if not urls:
            messagebox.showinfo("다른 PC 조회", "다른 PC 조회가 꺼져 있습니다.\n⚙️ 설정에서 켜 주세요.", parent=owner)
            return
        win = tk.Toplevel(owner)
        win.title("🌐 다른 PC 조회 주소")
        win.transient(owner)
        try:
            win.grab_set()                                # 설정 창이 잡고 있어도 이 창을 쓸 수 있게
        except Exception:
            pass
        f = ttk.Frame(win, padding=14)
        f.pack(fill="both", expand=True)
        by = "편집 프로그램 런처" if (self.view_server.external or {}).get("host") == "런처" else "이 편집기"
        ttk.Label(f, text=f"다른 PC의 브라우저(크롬·엣지) 주소창에 아래 주소 중 하나를 넣으면 정산/검색을 볼 수 있습니다.\n(지금은 {by}가 조회 페이지를 열고 있습니다. 런처만 켜 두어도 조회됩니다.)",
                  wraplength=520, justify="left").pack(anchor="w", pady=(0, 8))
        order = urls[1:] + urls[:1]                       # IP 주소를 먼저, PC 이름 주소는 마지막
        for i, u in enumerate(order):
            r = ttk.Frame(f)
            r.pack(fill="x", pady=2)
            e = ttk.Entry(r, width=36, font=(self.ui_font_family, 11))
            e.insert(0, u)
            e.configure(state="readonly")
            e.pack(side="left", padx=(0, 6))
            ttk.Button(r, text="복사", width=6, command=lambda u_=u: (win.clipboard_clear(), win.clipboard_append(u_), done_lbl.config(text=f"복사했습니다: {u_}"))).pack(side="left")
            ttk.Button(r, text="열어 보기", width=8, command=lambda u_=u: self._open_url(u_)).pack(side="left", padx=4)
            ttk.Label(r, text=("← 권장 (IP)" if i == 0 and len(order) > 1 else "PC 이름 (IP가 바뀌어도 됨)" if u == urls[0] else ""),
                      foreground="#64748b", font=(self.ui_font_family, 8)).pack(side="left")
        done_lbl = ttk.Label(f, text="", foreground="#15803d")
        done_lbl.pack(anchor="w", pady=(6, 0))
        path = self._write_view_address_file()
        ttk.Label(f, text=("주소를 '" + path + "' 에도 적어 두었습니다." if path else "") +
                          "\n안 열리면: 이 PC가 켜져 있고 편집기가 실행 중인지, 방화벽에서 '허용'했는지 확인해 주세요.",
                  foreground="#64748b", font=(self.ui_font_family, 8), wraplength=520, justify="left").pack(anchor="w", pady=(4, 0))
        ttk.Button(f, text="닫기", command=win.destroy).pack(pady=(10, 0))

    @staticmethod
    def _open_url(url):
        try:
            os.startfile(url)                             # 윈도우: 기본 브라우저로
        except Exception:
            try:
                import webbrowser
                webbrowser.open(url)
            except Exception:
                pass

    def _write_view_address_file(self):
        urls = self.view_server.urls()
        if not urls:
            return ""
        path = os.path.join(DATA_DIR, "다른PC_조회주소.txt")
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write("운행표 정산/검색 - 다른 PC에서 조회하는 주소\n")
                f.write(f"편집기 PC: {computer_name()}   (기록: {time.strftime('%Y-%m-%d %H:%M')})\n\n")
                for u in urls[1:] + urls[:1]:
                    f.write(u + "\n")
                f.write("\n다른 PC의 브라우저 주소창에 위 주소 중 하나를 넣으세요. 이 PC에서 편집 프로그램 런처나 편집기가 켜져 있을 때 열립니다.\n")
            return path
        except Exception:
            return ""

    def _build_default_toolbar(self):
        self._build_default_toolbar_items()
        flow_layout(self.toolbar)                     # 창이 좁으면 버튼이 둘째 줄로

    def _build_default_toolbar_items(self):
        self.clear_toolbar()
        t = self.toolbar
        self._add_view_badge(t)
        ttk.Button(t, text="PDF 열기", command=self.open_pdf).pack(side="left", padx=3)
        ttk.Button(t, text="⚙️ 설정", command=self.open_settings_window).pack(side="left", padx=3)
        ttk.Button(t, text="💰 정산 / 검색", command=self.show_report).pack(side="left", padx=3)
        ttk.Label(t, text="  상단의 'PDF 열기'를 눌러 시작하세요. (여러 파일을 한 번에 골라 탭으로 열 수 있습니다)").pack(side="left", padx=12)

    def _build_pdf_toolbar(self):
        self._build_pdf_toolbar_items()
        flow_layout(self.toolbar)                     # 창이 좁으면 버튼이 둘째 줄로

    def _build_pdf_toolbar_items(self):
        self.clear_toolbar()
        t = self.toolbar
        self._add_view_badge(t)
        ttk.Button(t, text="PDF 열기", command=self.open_pdf).pack(side="left", padx=3)
        ttk.Button(t, text="⚙️ 설정", command=self.open_settings_window).pack(side="left", padx=3)
        ttk.Separator(t, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(t, text="◀", width=3, command=lambda: self.goto_page(self.cur - 1)).pack(side="left", padx=1)
        e = ttk.Entry(t, textvariable=self.page_var, width=4, justify="center")
        e.pack(side="left", padx=1)
        e.bind("<Return>", lambda ev: self._goto_from_entry())
        ttk.Label(t, textvariable=self.page_total_var).pack(side="left", padx=2)
        ttk.Button(t, text="▶", width=3, command=lambda: self.goto_page(self.cur + 1)).pack(side="left", padx=1)
        ttk.Separator(t, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Checkbutton(t, text="칸 표시", variable=self.show_cells, command=self.draw_page).pack(side="left", padx=3)
        ttk.Checkbutton(t, text="재원 표시", variable=self.show_caps, command=self.draw_page).pack(side="left", padx=3)
        ttk.Checkbutton(t, text="오늘 수정 표시", variable=self.show_today, command=lambda: (self._save_settings(), self.draw_page())).pack(side="left", padx=3)
        ttk.Button(t, text="📝 오늘 수정 목록", command=self.show_today_list).pack(side="left", padx=3)
        ttk.Button(t, text="↩ 되돌리기 (Ctrl+Z)", command=self.undo_action).pack(side="left", padx=3)
        ttk.Button(t, text="📥 다른 날 가져오기", command=self.import_dialog).pack(side="left", padx=3)
        ttk.Separator(t, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(t, text="📋 읽은 데이터 표", command=self.show_table).pack(side="left", padx=3)
        ttk.Button(t, text="📌 재원 설정 목록", command=self.show_cap_list).pack(side="left", padx=3)
        ttk.Button(t, text="💰 정산 / 검색", command=self.show_report).pack(side="left", padx=3)
        ttk.Separator(t, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(t, text="💾 PDF 저장 + 기록", command=self.save_with_preview).pack(side="left", padx=3)
        ttk.Checkbutton(t, text="저장 전 미리보기", variable=self.preview_var, command=self._save_settings).pack(side="left", padx=2)
        ttk.Checkbutton(t, text="출력물 강조", variable=self.highlight_var, command=lambda: (self._save_settings(), self.draw_page())).pack(side="left", padx=2)
        ttk.Button(t, text="🏷 표시 설정", command=self.mark_opts_dialog).pack(side="left", padx=2)
        ttk.Button(t, text="기록만 저장", command=self.save_records_only).pack(side="left", padx=3)
        ttk.Separator(t, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(t, text="🔍+", command=lambda: self.change_zoom(1)).pack(side="left", padx=2)
        ttk.Label(t, textvariable=self.zoom_label, width=5, anchor="center").pack(side="left")
        ttk.Button(t, text="🔍-", command=lambda: self.change_zoom(-1)).pack(side="left", padx=2)

    def _build_default_sidebar(self):
        self.clear_sidebar()
        ttk.Label(self.sidebar_content_frame, text="데이터 입력", font=(self.ui_font_family, 13, "bold")).pack(anchor="w")
        ttk.Label(self.sidebar_content_frame, text="PDF 파일을 열면 칸에 넣을 데이터 목록과 입력 패널이 활성화됩니다.\n정산 / 검색은 PDF를 열지 않아도 쓸 수 있습니다.",
                  foreground="#555", wraplength=340).pack(anchor="w", pady=(2, 8))

    def _build_pdf_sidebar(self):
        self.clear_sidebar()
        f = self.sidebar_content_frame
        fnt = (self.ui_font_family, 8, "bold")
        ttk.Label(f, text="📦 데이터 입력", font=(self.ui_font_family, 13, "bold")).pack(anchor="w")
        ttk.Label(f, text="① 아래 목록에서 데이터를 고르고 ② PDF의 빈 칸을 더블클릭(또는 [선택 칸에 넣기])하세요.",
                  foreground="#555", wraplength=350).pack(anchor="w", pady=(2, 6))
        ttk.Separator(f).pack(fill="x", pady=4)

        info = ttk.LabelFrame(f, text=" 📄 파일 / 방송일 ", padding=8)
        info.pack(fill="x", pady=4)
        ttk.Label(info, textvariable=self.file_var, font=(self.ui_font_family, 8), wraplength=330).pack(anchor="w")
        drow = ttk.Frame(info)
        drow.pack(fill="x", pady=(4, 0))
        ttk.Label(drow, text="방송일:", font=fnt).pack(side="left")
        de = ttk.Entry(drow, textvariable=self.date_var, width=12)
        de.pack(side="left", padx=4)
        de.bind("<Return>", lambda e: self._commit_date())
        de.bind("<FocusOut>", lambda e: self._commit_date())
        ttk.Label(drow, text="(2026-10-30 형식)", font=(self.ui_font_family, 7), foreground="#666").pack(side="left")
        mrow = ttk.Frame(info)
        mrow.pack(fill="x", pady=(4, 0))
        ttk.Label(mrow, text="매체:", font=fnt).pack(side="left")
        mc = ttk.Combobox(mrow, textvariable=self.media_var, values=list(MEDIA_LIST), state="readonly", width=6)
        mc.pack(side="left", padx=4)
        mc.bind("<<ComboboxSelected>>", lambda e: self._commit_media())
        ttk.Label(mrow, text="(저장할 때 AM/FM 구분에 쓰입니다)", font=(self.ui_font_family, 7), foreground="#666").pack(side="left")

        cell = ttk.LabelFrame(f, text=" 🎯 선택된 칸 ", padding=8)
        cell.pack(fill="x", pady=4)
        ttk.Label(cell, textvariable=self.cell_info_var, font=(self.ui_font_family, 9), wraplength=330, justify="left").pack(anchor="w")
        crow = ttk.Frame(cell)
        crow.pack(fill="x", pady=(6, 0))
        ttk.Button(crow, text="선택 칸에 넣기", command=self.insert_selected_preset).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(crow, text="추가한 내용 지우기", command=self.clear_selected_cell).pack(side="left", fill="x", expand=True, padx=(2, 0))
        ttk.Button(cell, text="이 행 재원 설정", command=self.set_row_cap_dialog).pack(fill="x", pady=(4, 0))

        if self.preset_detached:                      # 미리 만든 데이터를 따로 창으로 띄워 둔 때
            self._open_preset_window()
            ph = ttk.LabelFrame(f, text=" ⭐ 미리 만든 데이터 ", padding=8)
            ph.pack(fill="x", pady=4)
            ttk.Label(ph, text="따로 창으로 띄워 두었습니다. (자유롭게 옮기고 크기를 바꿀 수 있습니다)", foreground="#555", wraplength=330).pack(anchor="w")
            r_ = ttk.Frame(ph)
            r_.pack(fill="x", pady=(4, 0))
            ttk.Button(r_, text="창 보이기", command=self._open_preset_window).pack(side="left", fill="x", expand=True, padx=(0, 2))
            ttk.Button(r_, text="⇤ 여기로 다시 붙이기", command=self.attach_presets).pack(side="left", fill="x", expand=True, padx=(2, 0))
        else:
            self._build_preset_panel(f)
        self.refresh_presets()
        self.update_cell_info()

    def _build_preset_panel(self, parent, in_window=False):
        """'미리 만든 데이터' 목록 (사이드바 안 또는 따로 띄운 창 안)."""
        fnt = (self.ui_font_family, 8, "bold")
        pre = ttk.LabelFrame(parent, text=" ⭐ 미리 만든 데이터 ", padding=8)
        pre.pack(fill="both", expand=True, pady=4, padx=(6 if in_window else 0))
        top_ = ttk.Frame(pre)
        top_.pack(fill="x", pady=(0, 4))
        if in_window:
            tv = tk.BooleanVar(value=self.preset_win_top)

            def _top():
                self.preset_win_top = bool(tv.get())
                try:
                    self.preset_win.attributes("-topmost", self.preset_win_top)
                except Exception:
                    pass
                self._save_settings()
            ttk.Checkbutton(top_, text="항상 위", variable=tv, command=_top).pack(side="left")
            ttk.Button(top_, text="⇤ 옆에 다시 붙이기", command=self.attach_presets).pack(side="right")
        else:
            ttk.Button(top_, text="⧉ 따로 창으로 띄우기", command=self.detach_presets).pack(side="right")
        frow = ttk.Frame(pre)
        frow.pack(fill="x", pady=(0, 4))
        ttk.Label(frow, text="찾기:", font=fnt).pack(side="left")
        fe = ttk.Entry(frow, textvariable=self.filter_var)
        fe.pack(side="left", fill="x", expand=True, padx=4)
        self.filter_var.trace_add("write", lambda *a: self.refresh_presets())

        srow = ttk.Frame(pre)
        srow.pack(fill="x", pady=(0, 4))
        ttk.Label(srow, text="정렬:", font=fnt).pack(side="left")
        self.sort_label_var = tk.StringVar(value=self.SORT_NAMES[self.sort_key])
        sc = ttk.Combobox(srow, textvariable=self.sort_label_var, values=list(self.SORT_KEYS), state="readonly", width=8)
        sc.pack(side="left", padx=4)
        sc.bind("<<ComboboxSelected>>", lambda e: self.set_sort(self.SORT_KEYS[self.sort_label_var.get()]))
        self.sort_dir_btn = ttk.Button(srow, text="", width=9, command=lambda: self.set_sort(self.sort_key, toggle=True))
        self.sort_dir_btn.pack(side="left")
        self._update_sort_btn()
        ttk.Button(srow, text="⚙ 열", width=5, command=self.preset_col_dialog).pack(side="right")

        tf = ttk.Frame(pre)
        tf.pack(fill="both", expand=True)
        self.preset_tree = ttk.Treeview(tf, columns=tuple(self.PRESET_COL_DEFS), show="headings", height=12, selectmode="browse")
        for c_, (title, _w, anchor, _st) in self.PRESET_COL_DEFS.items():
            sort_key = {"name": "name", "sec": "sec", "end": "end"}.get(c_)
            if sort_key:
                self.preset_tree.heading(c_, text=title, command=lambda k=sort_key: self.set_sort(k))
            else:
                self.preset_tree.heading(c_, text=title)
            self.preset_tree.column(c_, anchor=anchor)
        self._apply_preset_cols()
        self.preset_tree.bind("<Button-3>", self._preset_header_menu)          # 머리글 우클릭: 열 설정
        self.preset_tree.bind("<Button-2>", self._preset_header_menu)          # (맥 트랙패드)
        self.preset_tree.bind("<ButtonRelease-1>", lambda e: self._remember_preset_widths(), add="+")
        self.preset_tree.tag_configure("expired", foreground="#dc2626")
        ps = ttk.Scrollbar(tf, orient="vertical", command=self.preset_tree.yview)
        self.preset_tree.configure(yscrollcommand=ps.set)
        self.preset_tree.pack(side="left", fill="both", expand=True)
        ps.pack(side="right", fill="y")
        self.preset_tree.bind("<Double-Button-1>", lambda e: self.insert_selected_preset())
        self.preset_tree.bind("<Return>", lambda e: self.insert_selected_preset())

        ttk.Label(pre, textvariable=self.count_var, wraplength=330, foreground="#374151", justify="left").pack(anchor="w", pady=(4, 0))
        self.preset_tree.bind("<<TreeviewSelect>>", lambda e: self.update_count_label())
        brow = ttk.Frame(pre)
        brow.pack(fill="x", pady=(6, 0))
        ttk.Button(brow, text="추가", command=self.add_preset_dialog).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(brow, text="수정", command=self.edit_preset_dialog).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(brow, text="삭제", command=self.delete_preset_action).pack(side="left", fill="x", expand=True, padx=(2, 0))
        self.refresh_presets()

    def _open_preset_window(self):
        """미리 만든 데이터를 따로 창으로 (자유롭게 옮기기·크기 바꾸기, 다른 모니터로도)."""
        w = getattr(self, "preset_win", None)
        if w is not None and w.winfo_exists():
            w.deiconify()
            w.lift()
            return
        w = self.preset_win = tk.Toplevel(self.root)
        w.title("⭐ 미리 만든 데이터")
        try:
            w.geometry(self.preset_win_geom or "400x620")
        except Exception:
            w.geometry("400x620")
        if self.preset_win_top:
            w.attributes("-topmost", True)
        w.protocol("WM_DELETE_WINDOW", self.attach_presets)          # 창을 닫으면 사이드바로 다시 붙인다
        w.bind("<Configure>", lambda e: e.widget is w and self._remember_preset_geom())
        self._build_preset_panel(w, in_window=True)

    def _remember_preset_geom(self):
        w = getattr(self, "preset_win", None)
        if w is not None and w.winfo_exists():
            self.preset_win_geom = w.geometry()

    def detach_presets(self):
        self.preset_detached = True
        self._save_settings()
        self._open_preset_window()
        self._rebuild_side_after_preset()

    def attach_presets(self):
        self._remember_preset_geom()
        self.preset_detached = False
        self._save_settings()
        w = getattr(self, "preset_win", None)
        self.preset_win = None
        if w is not None:
            try:
                w.destroy()
            except Exception:
                pass
        self._rebuild_side_after_preset()

    def _rebuild_side_after_preset(self):
        if getattr(self, "_pdf_ui", False):
            self._build_pdf_sidebar()

    def draw_blank_canvas(self):
        c = self.canvas
        c.delete("all")
        c.configure(scrollregion=(0, 0, 1000, 800))
        c.create_text(500, 350, text="PDF 파일이 열리지 않았습니다.", font=(self.ui_font_family, 16, "bold"), fill="#888888")
        c.create_text(500, 390, text="상단의 [PDF 열기]를 눌러 시작하세요.", font=(self.ui_font_family, 11), fill="#aaaaaa")

    # ------------------------------------------------------------------ PDF 열기 / 읽기
    def open_pdf(self):
        if not fitz:
            messagebox.showerror("PDF 기능 필요", "PyMuPDF가 설치되어 있지 않습니다.\npip install pymupdf 명령어로 설치해주세요.")
            return
        if self.store is None:
            messagebox.showerror("열기 불가", "데이터베이스를 열 수 없어 PDF를 열 수 없습니다.\n⚙️ 설정에서 데이터 파일 위치를 확인해 주세요.")
            return
        paths = filedialog.askopenfilenames(title="운행표 PDF 열기 (Ctrl·Shift로 여러 개 선택 가능)", filetypes=[("PDF 파일", "*.pdf")])
        if not paths:
            return
        self._commit_date()
        lines, go_to = [], None
        for path in paths:
            res = self._read_pdf(path)
            if res is None:
                continue
            same = self._find_tab(res["media"], res["air_date"])
            if same is not None:
                res["doc"].close()
                lines.append(f"· {os.path.basename(path)}: {self._tab_title(res['media'], res['air_date'])} 은(는) 이미 열려 있어 그 탭을 씁니다.")
                go_to = same
                continue
            loaded, lost = self._add_tab(res)
            go_to = self.active_tab
            t = f"· {self._tab_title(res['media'], res['air_date'])}  ({res['page_count']}쪽, {res['total']}행)"
            if loaded:
                t += f" — 저장된 입력 {loaded}건 불러옴"
            n_t, n_d = self.today_counts()
            if n_t or n_d:
                t += f"\n     📝 오늘 수정: 넣거나 고친 것 {n_t}건" + (f", 지운 것 {n_d}건" if n_d else "") + " (화면에 초록/회색으로 표시)"
            if lost:
                t += f" (자리 못 찾은 {lost}건은 저장 때 보존)"
            lines.append(t)
        if go_to is not None:
            self._activate(go_to)
        if lines:
            messagebox.showinfo("PDF 읽기 완료", "\n".join(lines))
            self.alert_expired()

    def _read_pdf(self, path):
        """PDF 한 개를 읽어 탭을 만들 재료를 돌려준다. 실패하거나 취소하면 None (안내는 여기서)."""
        fname = os.path.basename(path)
        try:
            with open(path, "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            if getattr(doc, "needs_pass", False):
                doc.close()
                messagebox.showerror("열기 실패", f"{fname}\n\n암호가 걸린 PDF는 지원하지 않습니다.")
                return None
            if not doc.page_count:
                doc.close()
                raise ValueError("PDF에 페이지가 없습니다.")
            rows, date, first_words = {}, None, None
            hour = None
            for i in range(doc.page_count):
                words = words_from_page(doc.load_page(i))
                if i == 0:
                    first_words = words
                    date = parse_air_date(words)
                rows[i], hour = parse_rows(words, i, hour)
        except Exception as exc:
            messagebox.showerror("PDF 읽기 실패", f"{fname}\n\n{exc}")
            return None

        total = sum(len(v) for v in rows.values())
        if total == 0:
            doc.close()
            messagebox.showerror("읽기 실패", f"{fname}\n\n표에서 프로그램 행을 찾지 못했습니다.\n운행표 PDF가 맞는지 확인해 주세요.")
            return None
        if not date:
            ans = simpledialog.askstring("방송일 입력", f"{fname}\n\nPDF에서 방송일을 찾지 못했습니다.\n방송일을 2026-10-30 형식으로 입력해 주세요.", parent=self.root)
            if not ans or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", ans.strip()):
                doc.close()
                return None
            date = ans.strip()

        media_in = detect_media_from_words(first_words or [])     # PDF 안의 제목에서 읽은 매체 (가장 정확)
        media_nm = detect_media(path)                              # 파일 이름에서 읽은 매체
        media = media_in or media_nm
        if media_in and media_nm and media_in != media_nm:
            messagebox.showwarning("매체 확인", f"{fname}\n\n파일 이름은 {media_nm} 인데 PDF 안의 제목은 {media_in} 입니다.\n제목({media_in})을 기준으로 엽니다.\n맞지 않으면 오른쪽 '매체'에서 바꿔 주세요.")
        if not media:
            media = self.ask_media(fname)
            if not media:
                doc.close()
                return None
        return {"doc": doc, "data": data, "path": path, "rows": rows, "air_date": date, "media": media,
                "page_count": doc.page_count, "total": total}

    def _add_tab(self, res):
        """읽은 PDF로 새 탭을 만들고 그 탭을 지금 탭으로 한다. (loaded, lost) 를 돌려준다."""
        self._stash()
        self.active_tab = None
        for f, v in self._blank_state().items():
            setattr(self, f, v)
        doc = res["doc"]
        self.doc, self.src_bytes, self.pdf_path = doc, res["data"], res["path"]
        self.page_count = res["page_count"]
        self.rows = res["rows"]
        self.air_date = res["air_date"]
        self.media = res["media"]
        self._build_cells()
        try:
            self.caps = self.store.list_caps()
        except Exception:
            self.caps = {}
        self._reload_preset_map()
        loaded, lost = self._restore_records()
        pw = doc.load_page(0).rect.width / page_transform(doc.load_page(0))[0]
        self.zoom = round(min(max(900.0 / pw, 0.5), 2.0), 2) if pw > 0 else 1.0
        tab = {}
        self.tabs.append(tab)
        self.active_tab = tab
        self._show_active(view=(0.0, 0.0))
        self.alert_cap_changes()
        return loaded, lost

    # ------------------------------------------------------------------ 탭
    @staticmethod
    def _tab_title(media, air_date):
        """저장할 파일 이름과 같은 규칙: 매체 + 방송일 (예: AM20261030)."""
        return f"{media or '??'}{(air_date or '').replace('-', '')}"

    def _tab_info(self, tab):
        if tab is self.active_tab:
            return {"media": self.media, "air_date": self.air_date, "dirty": self.dirty, "path": self.pdf_path, "entries": self.entries}
        return {"media": tab["media"], "air_date": tab["air_date"], "dirty": tab["_dirty"], "path": tab["pdf_path"], "entries": tab["entries"]}

    def _sorted_tabs(self):
        def k(t):
            i = self._tab_info(t)
            return (i["air_date"], MEDIA_LIST.index(i["media"]) if i["media"] in MEDIA_LIST else 9)
        return sorted(self.tabs, key=k)

    def _find_tab(self, media, air_date, exclude=None):
        for t in self.tabs:
            if t is exclude:
                continue
            i = self._tab_info(t)
            if i["media"] == media and i["air_date"] == air_date:
                return t
        return None

    def _stash(self):
        """지금 탭의 상태를 탭 묶음에 넣어 둔다."""
        tab = self.active_tab
        if tab is None:
            return
        for f in self._blank_state():
            tab[f] = getattr(self, f)
        tab["_bg_key"] = tab["_bg_photo"] = None      # 화면 그림은 다시 만들면 되므로 메모리를 비운다
        try:
            tab["_view"] = (self.canvas.xview()[0], self.canvas.yview()[0])
        except Exception:
            tab["_view"] = (0.0, 0.0)

    def _activate(self, tab):
        """다른 탭으로 넘어간다."""
        if tab is None or tab not in self.tabs:
            return
        if tab is self.active_tab:
            self._rebuild_tabbar()
            return
        self._commit_date()
        self._stash()
        for f in self._blank_state():
            setattr(self, f, tab[f])
        self.active_tab = tab
        try:
            self.caps = self.store.list_caps() if self.store else {}   # 다른 탭에서 바꾼 재원 설정도 반영
        except Exception:
            pass
        self._reload_preset_map()
        self._show_active(view=tab.get("_view", (0.0, 0.0)))

    def _show_active(self, view=(0.0, 0.0)):
        """지금 탭의 내용으로 화면·입력칸·탭 줄을 맞춘다."""
        if not self._pdf_ui:
            self._build_pdf_toolbar()
            self._build_pdf_sidebar()
            self._pdf_ui = True
        self.date_var.set(self.air_date)
        self.media_var.set(self.media)
        self.page_var.set(str(self.cur + 1))
        self.page_total_var.set(f"/ {self.page_count}")
        self.zoom_label.set(f"{round(self.zoom * 100)}%")
        self.file_var.set(os.path.basename(self.pdf_path or ""))
        self.draw_page()
        self.update_idletasks_safe()
        try:
            self.canvas.xview_moveto(view[0])
            self.canvas.yview_moveto(view[1])
        except Exception:
            pass
        self.update_cell_info()
        self._rebuild_tabbar()

    def update_idletasks_safe(self):
        try:
            self.root.update_idletasks()
        except tk.TclError:
            pass

    def next_tab(self, step):
        if self.active_tab is None:
            return
        order = self._media_tabs(self.media)          # 같은 매체 안에서 날짜만 넘긴다
        if len(order) < 2 or self.active_tab not in order:
            return
        i = order.index(self.active_tab)
        self._activate(order[(i + step) % len(order)])

    def close_tab(self, tab):
        """탭을 닫는다. 저장하지 않은 입력이 있으면 묻는다. 닫았으면 True."""
        if tab is None or tab not in self.tabs:
            return False
        if tab is not self.active_tab:
            self._activate(tab)
        self._commit_date()
        if self.dirty:
            name = self._tab_title(self.media, self.air_date)
            ans = messagebox.askyesnocancel("탭 닫기", f"{name} 에 저장하지 않은 입력이 있습니다.\nPDF와 기록을 저장하고 닫을까요?\n(아니오: 저장하지 않고 닫기)")
            if ans is None:
                return False
            if ans and not self.export_pdf():
                return False
            if not ans:
                self._set_preview_dirty(False)         # 저장하지 않고 닫음 -> 편집중 표시만 지운다
        same = self._media_tabs(self.media)            # 닫은 뒤에는 같은 매체의 옆 날짜로, 없으면 다른 매체로
        order = same if len(same) > 1 else self._sorted_tabs()
        idx = order.index(tab)
        try:
            self.doc.close()
        except Exception:
            pass
        self.tabs.remove(tab)
        self.active_tab = None
        for f, v in self._blank_state().items():
            setattr(self, f, v)
        if self.tabs:
            order.remove(tab)
            self._activate(order[min(idx, len(order) - 1)])
        else:
            self._pdf_ui = False
            self._build_default_toolbar()
            self._build_default_sidebar()
            self.draw_blank_canvas()
            self.root.title(self._app_title())
            self._rebuild_tabbar()
        return True

    def save_all_records(self):
        """저장하지 않은 모든 탭의 입력 기록을 저장한다. (PDF 파일은 만들지 않는다)"""
        self._commit_date()
        done, fail = [], []
        back = self.active_tab
        for t in self._sorted_tabs():
            if not self._tab_info(t)["dirty"]:
                continue
            self._activate(t)
            name = self._tab_title(self.media, self.air_date)
            if self.save_records():
                self.dirty = False
                done.append(name)
            else:
                fail.append(name)
        self._activate(back)
        if not done and not fail:
            messagebox.showinfo("기록 모두 저장", "저장할 변경 내용이 없습니다.")
        else:
            msg = f"입력 기록을 저장했습니다: {', '.join(done)}" if done else ""
            if fail:
                msg += f"\n저장하지 못한 탭: {', '.join(fail)}"
            messagebox.showinfo("기록 모두 저장", msg.strip())

    MEDIA_COLORS = {"AM": "#2563eb", "FM": "#16a34a"}
    WEEK_KO = "월화수목금토일"

    def _date_label(self, air_date):
        try:
            d = datetime.date.fromisoformat(air_date)
            return f"{air_date.replace('-', '')} ({self.WEEK_KO[d.weekday()]})"
        except (ValueError, TypeError):
            return (air_date or "날짜 없음").replace("-", "")

    def _media_tabs(self, media):
        return [t for t in self._sorted_tabs() if self._tab_info(t)["media"] == media]

    def switch_media(self, media):
        """AM ↔ FM 전환. 같은 날짜가 열려 있으면 그 날짜, 없으면 그 매체에서 마지막으로 보던 탭."""
        if self.active_tab is not None and self.media == media:
            return
        group = self._media_tabs(media)
        if not group:
            if messagebox.askyesno(f"{media} 운행표", f"열려 있는 {media} 운행표가 없습니다.\n지금 PDF를 열까요?"):
                self.open_pdf()
            return
        target = self._find_tab(media, self.air_date) if self.active_tab is not None else None
        if target is None:
            last = getattr(self, "_last_by_media", {}).get(media)
            target = last if last in group else group[0]
        self._activate(target)

    def _rebuild_tabbar(self):
        bar = getattr(self, "tabbar", None)
        if bar is None:
            return
        for w in bar.winfo_children():
            w.destroy()
        if self.active_tab is not None:
            if not hasattr(self, "_last_by_media"):
                self._last_by_media = {}
            self._last_by_media[self.media] = self.active_tab
        top = tk.Frame(bar, background="#cbd5e1")
        top.pack(side="top", fill="x")
        cur_media = self.media if self.active_tab is not None else ""
        self._media_btns = {}
        medias = list(MEDIA_LIST) + sorted({self._tab_info(t)["media"] for t in self.tabs} - set(MEDIA_LIST))
        for m in medias:
            col = self.MEDIA_COLORS.get(m, "#475569")
            on = m == cur_media
            has = bool(self._media_tabs(m))
            b = tk.Label(top, text="", cursor="hand2", padx=14, pady=4, font=(self.ui_font_family, 11, "bold"),
                         background=col if on else "#e2e8f0", foreground="#ffffff" if on else (col if has else "#94a3b8"),
                         highlightthickness=1, highlightbackground=col)
            b.pack(side="left", padx=(6 if m == medias[0] else 3, 0), pady=(5, 0))
            b.bind("<Button-1>", lambda e, m=m: self.switch_media(m))
            self._media_btns[m] = b
        tk.Button(top, text="💾 기록 모두 저장", relief="flat", background="#cbd5e1", activebackground="#e2e8f0",
                  font=(self.ui_font_family, 9), command=self.save_all_records).pack(side="right", padx=6, pady=(4, 0))
        tk.Label(top, text="Ctrl+1 / Ctrl+2: AM / FM   Ctrl+Tab: 다음 날짜   Ctrl+W: 탭 닫기", background="#cbd5e1",
                 foreground="#475569", font=(self.ui_font_family, 8)).pack(side="right", padx=6)

        col = self.MEDIA_COLORS.get(cur_media, "#475569")
        sub = tk.Frame(bar, background="#f1f5f9", highlightthickness=0)
        sub.pack(side="top", fill="x")
        tk.Frame(sub, background=col, height=3).pack(side="top", fill="x")     # 매체 색 줄
        row = tk.Frame(sub, background="#f1f5f9")
        row.pack(side="top", fill="x")
        if not self.tabs:
            tk.Label(row, text="  열린 운행표가 없습니다.  [PDF 열기]에서 여러 파일을 고르면 AM / FM 별로 날짜 탭이 생깁니다.",
                     background="#f1f5f9", foreground="#475569", font=(self.ui_font_family, 9)).pack(side="left", pady=4)
            self._refresh_tab_titles()
            return
        for t in self._media_tabs(cur_media):
            active = t is self.active_tab
            bg = "#ffffff" if active else "#e2e8f0"
            f = tk.Frame(row, background=bg, highlightthickness=1, highlightbackground=col if active else "#94a3b8")
            f.pack(side="left", padx=(4, 0), pady=4)
            lbl = tk.Label(f, text="", background=bg, cursor="hand2", padx=8, pady=2,
                           font=(self.ui_font_family, 10, "bold" if active else "normal"))
            lbl.pack(side="left")
            x = tk.Label(f, text="✕", background=bg, foreground="#64748b", cursor="hand2", padx=4)
            x.pack(side="left")
            for w in (f, lbl):
                w.bind("<Button-1>", lambda e, t=t: self._activate(t))
            lbl.bind("<Button-2>", lambda e, t=t: self.close_tab(t))     # 가운데 버튼으로 닫기
            x.bind("<Button-1>", lambda e, t=t: self.close_tab(t))
            x.bind("<Enter>", lambda e, x=x: x.config(foreground="#dc2626"))
            x.bind("<Leave>", lambda e, x=x: x.config(foreground="#64748b"))
            t["_lbl"] = lbl
        self._refresh_tab_titles()

    def _refresh_tab_titles(self):
        for t in self.tabs:
            lbl = t.get("_lbl")
            if lbl is None:
                continue
            try:
                if not lbl.winfo_exists():
                    continue
                i = self._tab_info(t)
                lbl.config(text=self._date_label(i["air_date"]) + (" *" if i["dirty"] else ""),
                           foreground="#c2410c" if i["dirty"] else "#111827")
            except tk.TclError:
                pass
        for m, b in getattr(self, "_media_btns", {}).items():
            try:
                if not b.winfo_exists():
                    continue
                group = self._media_tabs(m)
                star = " *" if any(self._tab_info(t)["dirty"] for t in group) else ""
                b.config(text=f"{m}  ({len(group)}){star}")
            except tk.TclError:
                pass
        if self.active_tab is not None:
            self.root.title(f"{self._app_title()}  —  {self._tab_title(self.media, self.air_date)}{' *' if self.dirty else ''}")

    def ask_media(self, fname=""):
        """파일 이름에서 AM/FM 을 알 수 없을 때 물어본다. 취소하면 None."""
        win = tk.Toplevel(self.root)
        win.title("매체 선택")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        res = {"v": None}
        ttk.Label(win, text=(f"{fname}\n\n" if fname else "") + "파일 이름에서 AM / FM 을 알 수 없습니다.\n이 운행표는 어느 쪽인가요?", justify="left").pack(padx=18, pady=(16, 8))
        bf = ttk.Frame(win)
        bf.pack(pady=(0, 14))

        def pick(v):
            res["v"] = v
            win.destroy()

        ttk.Button(bf, text="AM", width=8, command=lambda: pick("AM")).pack(side="left", padx=6)
        ttk.Button(bf, text="FM", width=8, command=lambda: pick("FM")).pack(side="left", padx=6)
        ttk.Button(bf, text="취소", width=8, command=win.destroy).pack(side="left", padx=6)
        win.wait_window(win)
        return res["v"]

    def _commit_media(self):
        v = self.media_var.get().strip()
        if v in MEDIA_LIST and v != self.media:
            if self.media in MEDIA_LIST:
                n = len(self.entries)
                msg = (f"매체를 {self.media} → {v} 로 바꾸려고 합니다.\n\n"
                       f"· 저장하는 기록이 {v} 운행표로 들어가고, 일정표·정산도 {v} 로 집계됩니다.\n"
                       f"· 파일이 {self.media} 운행표가 맞다면 바꾸지 마세요.")
                if n:
                    msg += f"\n· 지금 입력해 둔 {n}건도 {v} 기록으로 저장됩니다."
                if not messagebox.askyesno("⚠ 매체 변경 확인", msg + "\n\n정말 바꿀까요?", icon="warning"):
                    self.media_var.set(self.media)
                    return
            if not self._check_new_key(v, self.air_date):
                self.media_var.set(self.media)
                return
            self.media = v
            self.dirty = True
            self.update_cell_info()
            self._rebuild_tabbar()

    def _check_new_key(self, media, air_date):
        """이 탭의 매체·방송일을 바꿔도 되는지. 다른 탭과 겹치면 막고, 이미 저장된 기록이 있으면 묻는다."""
        name = self._tab_title(media, air_date)
        if self._find_tab(media, air_date, exclude=self.active_tab) is not None:
            messagebox.showwarning("이미 열린 운행표", f"{name} 은(는) 다른 탭에 이미 열려 있습니다.\n같은 운행표를 두 탭에서 고치면 나중에 저장한 쪽이 앞의 입력을 지우므로 바꿀 수 없습니다.")
            return False
        try:
            n = len(self.store.load_records(air_date, media)) if self.store else 0
        except Exception:
            n = 0
        if n and not messagebox.askyesno("저장된 기록 있음", f"{name} 에는 이미 저장된 입력 기록 {n}건이 있습니다.\n이 탭을 저장하면 그 기록이 이 탭의 내용으로 바뀝니다.\n\n계속할까요?", icon="warning"):
            return False
        return True

    def _build_cells(self):
        self.cells = {}
        self.cell_by_key = {}
        for p, rlist in self.rows.items():
            lst = []
            for row in rlist:
                for li, line in enumerate(row["lines"]):
                    for k in range(4):
                        key = (p, row["time"], row["program"], li, k)
                        cd = {"key": key, "rect": cell_rect(line["cy"], k), "text": line["ads"][k], "cy": line["cy"], "row": row}
                        lst.append(cd)
                        self.cell_by_key[key] = cd
                    # 가상 칸: 4~7 = 각 칸의 오른쪽 반(칸 나누기), 8~11 = 칸이 모자랄 때 '+N' 으로만 기록하는 자리
                    for k in range(4, 12):
                        host = self.cell_by_key[(p, row["time"], row["program"], li, k % 4 if k < 8 else 3)]
                        key = (p, row["time"], row["program"], li, k)
                        self.cell_by_key[key] = {"key": key, "rect": host["rect"], "text": host["text"] if k < 8 else "",
                                                 "cy": line["cy"], "row": row, "virtual": True}
            self.cells[p] = lst

    def _restore_records(self):
        self.entries = {}
        self.orphans = []
        loaded = 0
        for r in self.store.load_records(self.air_date, self.media):
            key = (r["page"], r["time"], r["program"], r["line"], r["cell"] - 1)
            cd = self.cell_by_key.get(key)
            entry = {"name": r["name"], "seconds": r["seconds"], "price": r["price"], "grp": r["grp"], "saved_at": r.get("saved_at") or ""}
            if cd is not None and not cd["text"]:
                self.entries[key] = entry
                loaded += 1
            else:
                self.orphans.append(dict(r))
        self._load_removed()
        return loaded, len(self.orphans)

    # ------------------------------------------------------------------ 오늘 고친 것 표시
    def _load_removed(self):
        try:
            self.removed = self.store.load_removed(self.air_date, self.media) if self.store else []
        except Exception:
            self.removed = []

    @staticmethod
    def is_today_entry(ent):
        """오늘 넣거나 고친 칸인가. 아직 저장하지 않은 것(saved_at 없음)도 오늘 것으로 본다."""
        sa = (ent or {}).get("saved_at") or ""
        return not sa or sa[:10] == today_iso()

    def removed_ghosts(self, page):
        """오늘 지웠고 지금은 비어 있는 칸. [(key, 지운 기록)]"""
        out = []
        for r in getattr(self, "removed", []) or []:
            if r.get("page") != page:
                continue
            key = (r["page"], r["time"], r["program"], r.get("line") or 0, int(r["cell"]) - 1)
            if key[4] >= 4 or key in self.entries or (key[:4] + (key[4] + 4,)) in self.entries:
                continue
            cd = self.cell_by_key.get(key)
            if cd is None or cd["text"]:
                continue
            out.append((key, r))
        return out

    def today_counts(self):
        n = sum(1 for e in self.entries.values() if self.is_today_entry(e))
        return n, len(getattr(self, "removed", []) or [])

    def _commit_date(self):
        v = self.date_var.get().strip()
        if v == self.air_date:
            return
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
            messagebox.showwarning("방송일", "방송일은 2026-10-30 형식으로 입력해 주세요.")
            self.date_var.set(self.air_date)
            return
        if self.active_tab is not None and not self._check_new_key(self.media, v):
            self.date_var.set(self.air_date)
            return
        self.air_date = v
        self.dirty = True
        self._rebuild_tabbar()

    # ------------------------------------------------------------------ 좌표 / 그리기
    def c2d(self, x, y):
        return ((x - M) / self.zoom + self.page_origin[0], (y - M) / self.zoom + self.page_origin[1])

    def d2c(self, x, y):
        return (M + (x - self.page_origin[0]) * self.zoom, M + (y - self.page_origin[1]) * self.zoom)

    def _render_bg(self):
        key = (self.cur, self.zoom)
        if self._bg_key == key and self._bg_photo is not None:
            return
        page = self.doc.load_page(self.cur)
        k, tx, ty = page_transform(page)
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom / k, self.zoom / k), alpha=False)
        self._bg_photo = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"))
        r = page.rect
        self.page_origin = ((r.x0 - tx) / k, (r.y0 - ty) / k)
        self.page_size = (r.width / k, r.height / k)
        self._bg_key = key

    def _reload_preset_map(self):
        try:
            self._preset_map = {(p["name"], p["seconds"]): p for p in self.store.list_presets()} if self.store else {}
        except Exception:
            self._preset_map = {}

    def short_of(self, name, seconds=None):
        pm = getattr(self, "_preset_map", {})
        p = pm.get((name, seconds)) or next((v for (n, _), v in pm.items() if n == name), None)
        return ((p or {}).get("short_name") or "").strip() or abbreviate(name)

    def entry_layout(self, page, width_fn=None):
        """한 페이지에 그릴 추가 데이터의 자리·글자·크기. 화면과 저장 PDF가 같은 계산을 쓴다.
        - 한 줄의 광고(원본 + 추가)가 5개 이상이면 줄임 표시('20 설총')로 바꾼다. 다른 줄은 그대로.
        - 칸을 반으로 나눈 경우 왼쪽/오른쪽 반에 각각 넣는다.
        - 그래도 넘치면 글자 폭을 70%까지 줄이고(간격 좁히기), 그래도 넘치면 끝을 자른다. 기록은 그대로다."""
        wfn = width_fn or est_text_width
        lines = {}
        for key, ent in self.entries.items():
            if key[0] == page and key in self.cell_by_key:
                lines.setdefault(key[:4], []).append((key, ent))
        items, marks = [], []
        for lk, lst in lines.items():
            cd0 = self.cell_by_key.get(lk + (0,))
            if cd0 is None:
                continue
            ads = cd0["row"]["lines"][lk[3]]["ads"]
            compressed = sum(1 for a in ads if a) + len(lst) >= COMPRESS_MIN
            cy = cd0["cy"]
            over = []
            for key, ent in sorted(lst, key=lambda t: t[0][4]):
                c = key[4]
                if c >= 8:
                    over.append((key, ent))
                    continue
                host, slot = c % 4, c // 4
                x0, y0, x1, y1 = cell_rect(cy, host)
                if (lk + (host,)) in self.entries and (lk + (host + 4,)) in self.entries:
                    mid = (x0 + x1) / 2.0
                    rx0, rx1, inset = (x0, mid, SPLIT_INSET) if slot == 0 else (mid, x1, SPLIT_INSET)
                else:
                    rx0, rx1, inset = x0, x1, TEXT_INSET
                full = self.entry_text(ent)
                text = f"{ent['seconds']} {self.short_of(ent['name'], ent['seconds'])}" if compressed else full
                avail = rx1 - rx0 - inset - 1.5
                fs, xs = FONT_SIZE, 1.0
                if compressed or slot == 1 or (lk + (host + 4,)) in self.entries:
                    w = wfn(text, fs)
                    if w > avail:
                        xs = max(MIN_XSCALE, avail / w)
                        while len(text) > 1 and wfn(text, fs) * xs > avail + 0.01:
                            text = text[:-1]
                else:
                    while fs > TEXT_MIN_SIZE and wfn(text, fs) > avail:
                        fs -= 0.5
                items.append({"key": key, "rect": (rx0, y0, rx1, y1), "text": text, "fs": fs, "xs": xs, "inset": inset,
                              "full": full, "compressed": compressed})
            if over:
                marks.append({"lk": lk, "cy": cy, "n": len(over), "keys": [k for k, _ in over],
                              "names": [self.entry_text(e) for _, e in over]})
        return items, marks

    def _fit_size(self, text, max_w):
        fs = FONT_SIZE
        while fs > TEXT_MIN_SIZE and est_text_width(text, fs) > max_w:
            fs -= 0.5
        return fs

    @staticmethod
    def entry_text(ent):
        return f"{ent['seconds']} {ent['name']}"

    def over_labels(self, marks, width_fn=None):
        """'+N' 으로만 기록된 광고의 이름표 자리. 그 줄의 아래(위) 경계선에 걸쳐 오른쪽 맞춤으로 쓴다.
        화면과 저장 PDF가 같은 계산을 쓴다. 너무 길면 글자 폭을 줄이고, 그래도 넘치면 끝을 자른다."""
        wfn = width_fn or est_text_width
        out = []
        o = getattr(self, "mark_opts", None) or self._default_mark_opts()
        avail = OVER_LABEL_X1 - OVER_LABEL_X0
        for m in marks:
            fs = o["fs"]
            text = "+ " + ", ".join(m["names"])
            xs = 1.0
            w = wfn(text, fs)
            if w > avail:
                xs = max(0.6, avail / w)
                while len(text) > 3 and wfn(text, fs) * xs > avail + 0.01:
                    text = text[:-1]
                w = wfn(text, fs) * xs
            yc = m["cy"] + (CELL_HALF_H if o["pos"] == "below" else -CELL_HALF_H) + o["dy"]
            x0 = OVER_LABEL_X1 - w
            half = fs * 0.5 + o["pad"]
            out.append({"text": text, "fs": fs, "xs": xs, "x0": x0, "yc": yc, "full": ", ".join(m["names"]),
                        "rect": (x0 - 1.2, yc - half, OVER_LABEL_X1 + 1.2, yc + half)})
        return out

    def draw_page(self):
        if not self.doc:
            return
        c = self.canvas
        c.delete("all")
        self._render_bg()
        pw, ph = self.page_size
        c.create_image(M, M, image=self._bg_photo, anchor="nw")
        c.configure(scrollregion=(0, 0, pw * self.zoom + 2 * M, ph * self.zoom + 2 * M))
        z = self.zoom
        items, marks = self.entry_layout(self.cur)
        self._layout_cache = (self.cur, id(self.entries), items, marks, self.over_labels(marks))   # 마우스 움직일 때 다시 계산하지 않게
        by_key = {it["key"]: it for it in items}
        busy = {it["key"][:4] + (it["key"][4] % 4,) for it in items if it["key"][4] < 8}
        for cd in self.cells.get(self.cur, []):
            if cd["key"] in busy:
                continue
            if self.show_cells.get() and not cd["text"]:
                x0, y0, x1, y1 = cd["rect"]
                a, b = self.d2c(x0, y0)
                cc, d = self.d2c(x1, y1)
                c.create_rectangle(a, b, cc, d, outline="#cbd5e1", dash=(2, 3))
        show_today = self.show_today.get()
        if show_today:                                    # 오늘 지운 칸: 회색 점선 + 취소선 (화면에만)
            for key, r in self.removed_ghosts(self.cur):
                x0, y0, x1, y1 = self.cell_by_key[key]["rect"]
                a, b = self.d2c(x0, y0)
                cc, d = self.d2c(x1, y1)
                c.create_rectangle(a, b, cc, d, outline="#9ca3af", dash=(3, 2))
                px = max(4, int(round(FONT_SIZE * 0.85 * z)))
                c.create_text(a + TEXT_INSET * z, (b + d) / 2, text=f"{r['seconds']} {r['name']}", anchor="w",
                              font=(self.ui_font_family, -px, "overstrike"), fill="#9ca3af")
        for it in items:
            x0, y0, x1, y1 = it["rect"]
            a, b = self.d2c(x0, y0)
            cc, d = self.d2c(x1, y1)
            over = self.is_over(self.row_check(self.cell_by_key[it["key"]]["row"]))
            today = show_today and self.is_today_entry(self.entries.get(it["key"]))
            if today:                                     # 오늘 넣거나 고친 칸: 연초록 바탕 (화면에만)
                c.create_rectangle(a, b, cc, d, fill="#dcfce7", outline="#fca5a5" if over else "#16a34a", width=max(1, round(1.2 * z)))
            else:
                c.create_rectangle(a, b, cc, d, outline="#fca5a5" if over else "#93c5fd")
            px = max(4, int(round(it["fs"] * it["xs"] * z)))
            c.create_text(a + it["inset"] * z, (b + d) / 2, text=it["text"], anchor="w", font=(self.ui_font_family, -px),
                          fill="#dc2626" if over else "#15803d" if today else "#1d4ed8")
        for m in marks:
            mx, my = self.d2c(MARK_X, m["cy"])
            today = show_today and any(self.is_today_entry(self.entries.get(k_)) for k_ in m["keys"])
            c.create_text(mx, my, text=f"+{m['n']}", anchor="w", font=(self.ui_font_family, -max(5, int(round(7 * z))), "bold"),
                          fill="#15803d" if today else "#2563eb")
        for m, lb in zip(marks, self.over_labels(marks)):   # 칸이 모자란 줄의 추가 광고 이름표 (저장 PDF와 같은 자리)
            x0, y0, x1, y1 = lb["rect"]
            a, b = self.d2c(x0, y0)
            cc, d = self.d2c(x1, y1)
            hl_on = self.highlight_var.get()
            today = show_today and any(self.is_today_entry(self.entries.get(k_)) for k_ in m["keys"])
            c.create_rectangle(a, b, cc, d, fill="#dcfce7" if today else "#ffffff" if self.mark_opts["fill"] else "",
                               outline="#1d4ed8" if hl_on else "#93c5fd", width=max(1, round(self.mark_opts["hl_width"] * z)) if hl_on else 1)
            lx, ly = self.d2c(lb["x0"], lb["yc"])
            c.create_text(lx, ly, text=lb["text"], anchor="w", font=(self.ui_font_family, -max(4, int(round(lb["fs"] * lb["xs"] * z * 1.05)))), fill="#1d4ed8")
        if self.show_caps.get():
            seen_link = set()
            for row in self.rows.get(self.cur, []):        # Spot·캠페인 묶음: 오른쪽에 합산 '합 판매/재원' (화면에만)
                lk = self._link_key(row)
                if not lk or lk in seen_link:
                    continue
                seen_link.add(lk)
                grp = self.linked_rows(row)
                chk = self.row_check(row) if len(grp) > 1 else None
                if not chk:
                    continue
                col = self.CAP_COLORS[self.cap_state(row)]
                ys = [r["lines"][0]["cy"] for r in grp]
                a, b = self.d2c(AMOUNT_X1 - 1, min(ys) - CELL_HALF_H + 1)
                cc, d = self.d2c(AMOUNT_X1 + 1, max(ys) + CELL_HALF_H - 1)
                c.create_line(a, b, a, d, fill=col, width=max(2, round(2 * z)))
                tx, ty = self.d2c(AMOUNT_X1 + 3, (min(ys) + max(ys)) / 2)
                c.create_text(tx, ty, text=f"합 {chk[3]}/{chk[1]}", anchor="w", fill=col,
                              font=(self.ui_font_family, -max(6, int(round(7 * z))), "bold"))
            for row in self.rows.get(self.cur, []):
                if not self.row_is_custom(row):
                    continue
                cy = row["lines"][0]["cy"]
                col = self.CAP_COLORS[self.cap_state(row)]
                a, b = self.d2c(AD_BORDERS[4], cy - CELL_HALF_H)
                cc, d = self.d2c(AMOUNT_X1, cy + CELL_HALF_H)
                c.create_rectangle(a, b, cc, d, outline=col, width=2)
                dx, dy = self.d2c(12, cy)
                rad = 3 * z + 1
                c.create_oval(dx - rad, dy - rad, dx + rad, dy + rad, fill=col, outline=col)
        cd = self.cell_by_key.get(self.selected_key) if self.selected_key else None
        if cd and self.selected_key[0] == self.cur:
            x0, y0, x1, y1 = by_key[self.selected_key]["rect"] if self.selected_key in by_key else cd["rect"]
            a, b = self.d2c(x0, y0)
            cc, d = self.d2c(x1, y1)
            c.create_rectangle(a, b, cc, d, outline="#dc2626" if cd["text"] else "#2563eb", width=2)

    # ------------------------------------------------------------------ 페이지 / 확대
    def _goto_from_entry(self):
        try:
            n = int(self.page_var.get())
        except ValueError:
            self.page_var.set(str(self.cur + 1))
            return
        self.goto_page(n - 1)
        self.canvas.focus_set()

    def goto_page(self, idx):
        if not self.doc:
            return
        idx = max(0, min(self.page_count - 1, idx))
        self.page_var.set(str(idx + 1))
        if idx == self.cur:
            return
        if idx > self.cur and self.page_check_on and not self._confirm_leave_page(self.cur):
            self.page_var.set(str(self.cur + 1))
            return
        self.cur = idx
        if self.selected_key and self.selected_key[0] != idx:
            self.selected_key = None
            self.update_cell_info()
        self.draw_page()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def _page_problems(self, page):
        """이 장에서 재원을 넘친 행 / 모자란 행: ([(시각, 초)], [(시각, 초)]). Spot·캠페인 59는 한 번만."""
        over, short, seen = [], [], set()
        for row in self.rows.get(page, []):
            lk = self._link_key(row)
            if not lk:
                continue                              # Spot0859·캠페인1329 처럼 Spot/캠페인 + 시각 행만 본다
            k = ("L", lk)
            if k in seen:
                continue
            seen.add(k)
            chk = self.row_check(row)
            if not chk:
                continue
            sold, cap, added, total = chk
            if cap <= 0:
                continue
            if total > cap:
                over.append((lk, total - cap))
            elif total < cap:
                short.append((lk, cap - total))
        return over, short

    def _confirm_leave_page(self, page):
        """다음 장으로 넘어가기 전에 재원을 넘친 행·모자란 행을 알려 준다. 넘어가면 True."""
        over, short = self._page_problems(page)
        if not over and not short:
            return True

        def fmt(lst, sign):
            txt = ", ".join(f"{t}({sign}{n}초)" for t, n in lst[:40])
            return txt + (f" 외 {len(lst) - 40}개" if len(lst) > 40 else "")
        msg = f"{page + 1}장 확인\n\n"
        if over:
            msg += f"⚠ 시간 초과 {len(over)}개: {fmt(over, '+')}\n\n"
        if short:
            msg += f"△ 부족 {len(short)}개: {fmt(short, '-')}\n\n"
        msg += "다음 장으로 넘어갈까요?\n(아니요: 이 장에 남아서 고치기 · ⚙ 설정에서 이 확인을 끌 수 있습니다)"
        return messagebox.askyesno("재원 확인", msg, icon="warning" if over else "info")

    def change_zoom(self, direction):
        if not self.doc:
            return
        cur = min(range(len(ZOOM_LEVELS)), key=lambda i: abs(ZOOM_LEVELS[i] - self.zoom))
        tgt = max(0, min(len(ZOOM_LEVELS) - 1, cur + direction))
        if tgt == cur:
            return
        view = (self.canvas.xview()[0], self.canvas.yview()[0])
        self.zoom = ZOOM_LEVELS[tgt]
        self.zoom_label.set(f"{round(self.zoom * 100)}%")
        self.draw_page()
        try:
            self.canvas.xview_moveto(view[0])
            self.canvas.yview_moveto(view[1])
        except Exception:
            pass

    def on_mouse_wheel(self, event):
        if event.state & 0x0004:
            self.change_zoom(1 if (event.delta > 0 or event.num == 4) else -1)
            return "break"
        move = -1 if (event.delta > 0 or event.num == 4) else 1 if (event.delta < 0 or event.num == 5) else 0
        self.canvas.yview_scroll(move, "units")
        return "break"

    # ------------------------------------------------------------------ 칸 선택 / 입력
    def hit_cell(self, x, y):
        for cd in self.cells.get(self.cur, []):
            x0, y0, x1, y1 = cd["rect"]
            if x0 <= x < x1 and y0 <= y < y1:
                return cd
        return None

    def select_cell(self, key):
        self.selected_key = key
        self.update_cell_info()
        self.draw_page()

    def _pick_key(self, dx, dy):
        """클릭한 자리의 칸 key. '+N' 표시를 누르면 그 줄의 마지막 추가 항목, 반으로 나눈 칸이면 누른 쪽 반."""
        items, marks = self.entry_layout(self.cur)
        for m in marks:
            if MARK_X - 2 <= dx <= MARK_X + 14 and abs(dy - m["cy"]) <= CELL_HALF_H:
                return max(m["keys"], key=lambda k: k[4])
        cd = self.hit_cell(dx, dy)
        if not cd:
            return None
        key = cd["key"]
        for it in items:
            k2 = it["key"]
            if k2[:4] == key[:4] and k2[4] < 8 and k2[4] % 4 == key[4] and it["rect"][0] <= dx < it["rect"][2]:
                return k2
        return key

    def on_press(self, event):
        if not self.doc:
            return
        self.canvas.focus_set()
        dx, dy = self.c2d(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.select_cell(self._pick_key(dx, dy))

    def on_double_click(self, event):
        if not self.doc:
            return
        dx, dy = self.c2d(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        key = self._pick_key(dx, dy)
        if key:
            self.selected_key = key
            self.insert_selected_preset()

    def move_selection(self, drow, dcol):
        """방향키로 선택 칸 이동: 좌우는 같은 줄의 칸, 상하는 위/아래 줄의 같은 칸 번호."""
        if not self.doc or not self.selected_key:
            return
        p = self.selected_key[0]
        lst = self.cells.get(p, [])
        keys = [cd["key"] for cd in lst]
        if self.selected_key not in keys:
            return
        i = keys.index(self.selected_key)
        if dcol:
            j = i + dcol
        else:
            j = i + drow * 4
            if drow:
                # 줄이 4칸씩 이어져 있으므로 4칸씩 건너뛰면 같은 칸 번호로 위/아래 이동
                pass
        if 0 <= j < len(lst):
            self.cur = p
            self.select_cell(lst[j]["key"])
            self._scroll_to_cell(lst[j])

    def _scroll_to_cell(self, cd):
        try:
            x0, y0, x1, y1 = cd["rect"]
            _, cy = self.d2c(x0, (y0 + y1) / 2)
            total = self.page_size[1] * self.zoom + 2 * M
            view = self.canvas.yview()
            top, bot = view[0] * total, view[1] * total
            if cy < top + 40 or cy > bot - 40:
                self.canvas.yview_moveto(max(0.0, (cy - (bot - top) / 2) / total))
        except Exception:
            pass

    def media_min(self, preset):
        if not preset:
            return 0
        return int(preset.get("min_fm" if self.media == "FM" else "min_am") or 0)

    def count_text(self, name, seconds=None):
        """이 PDF(오늘, 현재 매체)에서 그 광고를 넣은 횟수와 최소 횟수 비교."""
        p = None
        try:
            p = self.store.get_preset_by_name(name, seconds) if self.store else None
        except Exception:
            pass
        n = sum(1 for e in self.entries.values() if e["name"] == name and (seconds is None or e["seconds"] == seconds))
        mn = self.media_min(p)
        t = f"오늘 {self.media or ''} {n}회".replace("  ", " ")
        if mn:
            t += f" / 최소 {mn}회 " + ("✓ 충족" if n >= mn else f"(아직 {mn - n}회 부족)")
        return f"{(str(seconds) + '초 ') if seconds else ''}{name}: {t}"

    def update_count_label(self):
        try:
            preset = self._selected_preset()
        except Exception:
            preset = None
        self.count_var.set(self.count_text(preset["name"], preset["seconds"]) if (preset and self.doc) else "")

    @staticmethod
    def remain_text(cap, total, custom):
        if total > cap:
            return f"{total - cap}초 초과"
        if total == cap:
            return "딱 맞음"
        return f"{cap - total}초 부족" if custom else f"남은 {cap - total}초"

    def update_cell_info(self):
        self.update_count_label()
        key = self.selected_key
        if not key or key not in self.cell_by_key:
            self.cell_info_var.set("PDF에서 칸을 클릭하세요.")
            return
        cd = self.cell_by_key[key]
        row = cd["row"]
        if key[4] < 4:
            where = f"광고주 {key[4] + 1}칸"
        elif key[4] < 8:
            where = f"광고주 {key[4] - 3}칸 (반으로 나눈 오른쪽)"
        else:
            where = f"행 오른쪽 '+N' 표시 ({key[4] - 7}번째)"
        head = f"{row['time']}  {row['program']}\n{where}" + (" (둘째 줄)" if key[3] else "")
        ent = self.entries.get(key)
        if cd["text"]:
            body = f"원본에 이미 내용이 있는 칸입니다:\n{cd['text']}"
        elif ent:
            sa = ent.get("saved_at") or ""
            when = ("아직 저장 안 함" if not sa else f"오늘 {sa[11:16]} 저장" if sa[:10] == today_iso() else f"{short_date(sa[:10])} {sa[11:16]} 저장")
            body = f"추가한 내용: {self.entry_text(ent)}  (그룹 {ent['grp'] or '-'})\n🕒 {when}\n" + self.count_text(ent["name"], ent["seconds"])
        else:
            body = "비어 있음 — 넣을 데이터를 골라 주세요."
            gone = [r for k_, r in self.removed_ghosts(key[0]) if k_ == key]
            if gone:
                body += f"\n(오늘 {gone[0]['removed_at'][11:16]}에 '{gone[0]['seconds']} {gone[0]['name']}' 지움)"
        chk = self.row_check(row)
        extra = ""
        tag = " (설정)" if any(self.row_is_custom(r) for r in self.linked_rows(row)) else ""
        if len(self.linked_rows(row)) > 1:
            parts = []
            for r in self.linked_rows(row):
                am = self.row_amount(r)
                parts.append(f"{r['program']} {am[0]}/{am[1]}" if am else r["program"])
            tag += "  [" + " + ".join(parts) + " 합산]"
        if chk:
            sold, cap, added, total = chk
            rt = self.remain_text(cap, total, bool(tag))
            if added > 0:
                extra = ("\n⚠ " if total > cap else "\n") + f"판매 {sold} + 추가 {added} = {total} / 재원 {cap}{tag}  → {rt}"
            else:
                extra = f"\n판매/재원 {sold} / {cap}{tag}  → {rt}"
        elif (parse_amount(row) or (0, 1))[1] == 0:
            extra = "\n원본 재원이 0입니다. [이 행 재원 설정]으로 재원을 정하면 비교해 드립니다."
        self.cell_info_var.set(head + "\n" + body + extra)

    def row_amount(self, row):
        """비교에 쓸 (판매, 재원). 사용자가 설정한 재원이 있으면 그것을, 없으면 원본을 쓴다.
        원본 재원이 0 이고 설정도 없으면 비교하지 않는다(None)."""
        orig = parse_amount(row)
        c = self.caps.get((row["time"], row["program"]))
        if c:
            return (orig[0] if orig else 0), c["cap"]
        if orig is None or orig[1] == 0:
            return None
        return orig

    def row_is_custom(self, row):
        return (row["time"], row["program"]) in self.caps

    # Spot0859 · Spot 1429 · 캠페인1329 · 0659캠페인 처럼 'Spot/캠페인 + 네 자리 시각(HHMM)' 인 행
    _SLOT_RE = re.compile(r"(?i)^\s*(?:spot|캠페인)\s*(\d{4})\s*$|^\s*(\d{4})\s*(?:spot|캠페인)\s*$")

    def _link_key(self, row):
        """Spot·캠페인 시각 행이면 그 시각('0659'), 아니면 None.
        같은 시각의 Spot 과 캠페인(예: Spot 0659 · 0659캠페인, Spot1329 · 캠페인1329)은 같은 자리로 본다."""
        m = self._SLOT_RE.match(row.get("program") or "")
        return (m.group(1) or m.group(2)) if m else None

    def linked_rows(self, row):
        """같은 자리로 묶이는 행들 (이 행 포함). 묶이지 않으면 [row]."""
        k = self._link_key(row)
        if not k:
            return [row]
        out = [r for r in self.rows.get(row.get("page"), []) if self._link_key(r) == k]
        return out if row in out else [row] + out

    def row_check(self, row, extra_key=None, extra_sec=0):
        """(원본 판매, 재원, 추가한 초, 합계). 재원을 못 읽은 행은 None.
        extra_key 칸의 기존 입력은 빼고 extra_sec 를 더해 '넣으면 어떻게 되는지'를 미리 계산할 수 있다.
        같은 시각의 Spot 과 캠페인(예: Spot 1659 · 1659캠페인)은 한 자리로 보고 판매·재원·추가 초를 모두 합친다."""
        group = self.linked_rows(row)
        amts = [a for a in (self.row_amount(r) for r in group) if a]
        if not amts:
            return None
        sold = sum(a[0] for a in amts)
        cap = sum(a[1] for a in amts)          # 두 행의 재원 총량을 합산 (예: 80 + 70 = 150)
        rks = {(r["page"], r["time"], r["program"]) for r in group}
        added = sum(e["seconds"] for k, e in self.entries.items() if k[:3] in rks and k != extra_key) + extra_sec
        return sold, cap, added, sold + added

    @staticmethod
    def is_over(chk):
        return bool(chk) and chk[2] > 0 and chk[3] > chk[1]

    CAP_COLORS = {"over": "#dc2626", "changed": "#f97316", "short": "#eab308", "full": "#16a34a", "set": "#2563eb"}

    def cap_state(self, row):
        if self.is_over(self.row_check(row)):
            return "over"
        if (row["time"], row["program"]) in self.cap_flag:
            return "changed"
        chk = self.row_check(row)
        if chk is None:
            return "set"
        return "short" if chk[3] < chk[1] else "full"

    def _show_tip(self, cx, cy_, text, right):
        c = self.canvas
        t = c.create_text(cx + (-14 if right else 14), cy_ + 16, text=text, anchor="ne" if right else "nw", tags="tip",
                          font=(self.ui_font_family, 9), fill="#111")
        bb = c.bbox(t)
        c.create_rectangle(bb[0] - 4, bb[1] - 3, bb[2] + 4, bb[3] + 3, fill="#fffbe6", outline="#999", tags="tip")
        c.tag_raise(t)

    def on_motion(self, event):
        c = self.canvas
        c.delete("tip")
        if not self.doc:
            return
        cx, cy_ = c.canvasx(event.x), c.canvasy(event.y)
        x, y = self.c2d(cx, cy_)
        lc = getattr(self, "_layout_cache", None)
        if lc and lc[0] == self.cur and lc[1] == id(self.entries):
            items, marks, labels = lc[2], lc[3], lc[4]
        else:
            items, marks = self.entry_layout(self.cur)
            labels = self.over_labels(marks)
        # 줄어들거나 잘려 보이는 글자는 마우스를 올리면 전체 이름을 보여 준다
        for it in items:
            x0, y0, x1, y1 = it["rect"]
            if x0 <= x < x1 and y0 <= y < y1 and (it["compressed"] or it["text"] != it["full"] or it["xs"] < 0.999):
                self._show_tip(cx, cy_, it["full"], x > 300)
                return
        if self.show_today.get():
            for key, r in self.removed_ghosts(self.cur):
                x0, y0, x1, y1 = self.cell_by_key[key]["rect"]
                if x0 <= x < x1 and y0 <= y < y1:
                    self._show_tip(cx, cy_, f"오늘 {r['removed_at'][11:16]}에 지움: {r['seconds']} {r['name']}", x > 300)
                    return
        for m in marks:
            if MARK_X - 2 <= x <= MARK_X + 14 and abs(y - m["cy"]) <= CELL_HALF_H:
                self._show_tip(cx, cy_, "행 오른쪽 '+N' 표시 (칸이 모자라 줄 아래에 이름표로 표시)\n" + "\n".join(m["names"]), True)
                return
        for lb in labels:
            x0, y0, x1, y1 = lb["rect"]
            if x0 <= x <= x1 and y0 <= y <= y1:
                self._show_tip(cx, cy_, lb["full"].replace(", ", "\n"), True)
                return
        if not self.show_caps.get() or not self.caps:
            return
        for row in self.rows.get(self.cur, []):
            if not self.row_is_custom(row):
                continue
            ry = row["lines"][0]["cy"]
            if abs(y - ry) <= CELL_HALF_H and (x >= AD_BORDERS[4] or x <= 24):
                cap = self.caps[(row["time"], row["program"])]["cap"]
                orig = parse_amount(row)
                text = f"재원 {cap}초 설정됨" + (f" (원본 {orig[0]}/{orig[1]})" if orig else "")
                chk = self.row_check(row)
                if chk:
                    rt = self.remain_text(chk[1], chk[3], True)
                    text += (f"\n판매 {chk[0]} + 추가 {chk[2]} = {chk[3]}" if chk[2] > 0 else f"\n판매 {chk[0]}") + f"  → {rt}"
                if (row["time"], row["program"]) in self.cap_flag:
                    text += "\n⚠ 원본 재원이 설정할 때와 달라졌습니다"
                self._show_tip(cx, cy_, text, x >= AD_BORDERS[4])
                return

    def show_cap_list(self):
        if self.store is None:
            return
        self.caps = self.store.list_caps()
        win = tk.Toplevel(self.root)
        win.title("📌 재원 설정 목록")
        win.geometry("820x480")
        cols = ("time", "program", "orig", "cap", "now", "gap", "state")
        only_short = tk.BooleanVar(value=False)
        ttk.Checkbutton(win, text="부족한 행만 보기", variable=only_short, command=lambda: refresh()).pack(anchor="w", padx=10, pady=(8, 0))
        tf = ttk.Frame(win)
        tf.pack(fill="both", expand=True, padx=8, pady=8)
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c_, h, w in zip(cols, ("시각", "프로그램명", "설정 당시 원본", "설정 재원", "현재(판매+추가)", "부족/초과", "상태"), (80, 200, 100, 70, 100, 90, 170)):
            tree.heading(c_, text=h)
            tree.column(c_, width=w, anchor="w" if c_ == "program" else "center")
        vs = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        tree.tag_configure("warn", foreground="#d97706")
        tree.tag_configure("short", foreground="#a16207")
        tree.tag_configure("over", foreground="#dc2626")
        info = ttk.Label(win, text="더블클릭하면 PDF의 해당 행으로 이동합니다. (PDF를 연 경우)", foreground="#555")
        info.pack(anchor="w", padx=10)

        def refresh():
            changed_keys, missing_keys = set(), set()
            rowmap = {}
            if self.doc:
                rowmap = {(r["time"], r["program"]): r for pg in self.rows for r in self.rows[pg]}
                ch, mi = self.cap_changes()
                changed_keys = {(t, pr) for t, pr, _, _ in ch} | set(self.cap_flag)
                missing_keys = {(t, pr) for t, pr, _ in mi}
            for i in tree.get_children():
                tree.delete(i)
            for (t, pr), c in sorted(self.caps.items()):
                if not self.doc:
                    st = "-"
                elif (t, pr) in missing_keys:
                    st = "이 PDF에서 행을 찾을 수 없음"
                elif (t, pr) in changed_keys:
                    st = "원본 재원 변경됨 (확인 필요)"
                else:
                    st = "정상"
                row = rowmap.get((t, pr))
                chk = self.row_check(row) if row else None
                now_s, gap_s, short, over = "-", "-", False, False
                if chk:
                    now_s = str(chk[3])
                    gap = chk[1] - chk[3]
                    short, over = gap > 0, gap < 0
                    gap_s = f"{gap}초 부족" if gap > 0 else ("딱 맞음" if gap == 0 else f"{-gap}초 초과")
                if only_short.get() and not short:
                    continue
                tag = ("over",) if over else ("short",) if short else ("warn",) if st not in ("정상", "-") else ()
                tree.insert("", "end", iid=f"{t}|{pr}", values=(t, pr, f"{c['base_sold']} / {c['base_cap']}", c["cap"], now_s, gap_s, st), tags=tag)

        def sel_key():
            sel = tree.selection()
            if not sel:
                return None
            t, pr = sel[0].split("|", 1)
            return t, pr

        def goto(event=None):
            k = sel_key()
            if not k or not self.doc:
                return
            row = next((r for p in self.rows for r in self.rows[p] if (r["time"], r["program"]) == k), None)
            if row is None:
                messagebox.showinfo("이동", "이 PDF에서 해당 행을 찾을 수 없습니다.", parent=win)
                return
            key0 = (row["page"], row["time"], row["program"], 0, 0)
            self.cur = row["page"]
            self.page_var.set(str(row["page"] + 1))
            self.selected_key = key0
            self.update_cell_info()
            self.draw_page()
            self._scroll_to_cell(self.cell_by_key[key0])

        def remove():
            k = sel_key()
            if not k:
                messagebox.showinfo("설정 해제", "해제할 행을 먼저 선택해 주세요.", parent=win)
                return
            if messagebox.askyesno("설정 해제", f"{k[0]} {k[1]}\n이 행의 재원 설정을 해제할까요?", parent=win):
                self.store.delete_cap(*k)
                self.caps = self.store.list_caps()
                refresh()
                if self.doc:
                    self.update_cell_info()
                    self.draw_page()

        tree.bind("<Double-Button-1>", goto)
        bf = ttk.Frame(win)
        bf.pack(pady=8)
        ttk.Button(bf, text="해당 행으로 이동", command=goto).pack(side="left", padx=4)
        ttk.Button(bf, text="설정 해제", command=remove).pack(side="left", padx=4)
        ttk.Button(bf, text="닫기", command=win.destroy).pack(side="left", padx=4)
        refresh()

    def set_row_cap_dialog(self):
        key = self.selected_key
        if not self.doc or not key or key not in self.cell_by_key:
            messagebox.showinfo("재원 설정", "재원을 설정할 행의 칸을 PDF에서 먼저 클릭해 주세요.")
            return
        row = self.cell_by_key[key]["row"]
        rk = (row["time"], row["program"])
        orig = parse_amount(row)
        cur = self.caps.get(rk)
        win = tk.Toplevel(self.root)
        win.title("이 행 재원 설정")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        fb = (self.ui_font_family, 9, "bold")
        ttk.Label(win, text=f"{row['time']}  {row['program']}", font=(self.ui_font_family, 10, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        ttk.Label(win, text=f"원본 판매/재원: {orig[0]} / {orig[1]}" if orig else "원본 판매/재원을 읽지 못했습니다.", foreground="#555").pack(anchor="w", padx=14)
        r1 = ttk.Frame(win)
        r1.pack(fill="x", padx=14, pady=8)
        ttk.Label(r1, text="재원(초):", font=fb).pack(side="left")
        cap_v = tk.StringVar(value=str(cur["cap"]) if cur else "")
        ttk.Entry(r1, textvariable=cap_v, width=10).pack(side="left", padx=6)
        ttk.Label(win, text="설정한 재원은 같은 시각·같은 프로그램명의 행에 매일 계속 적용됩니다.\nPDF의 원본 숫자는 바뀌지 않습니다.\n나중에 이 행의 원본 재원이 달라지면 PDF를 열 때 알려 드립니다.",
                  font=(self.ui_font_family, 8), foreground="#666", justify="left").pack(anchor="w", padx=14)

        def done():
            self.caps = self.store.list_caps()
            win.destroy()
            self.update_cell_info()
            self.draw_page()

        def save():
            try:
                cap = int(cap_v.get().strip())
                if cap <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("입력 확인", "재원은 1 이상의 숫자(초)로 입력해 주세요.", parent=win)
                return
            self.store.set_cap(rk[0], rk[1], cap, orig[0] if orig else 0, orig[1] if orig else 0)
            done()

        def clear():
            self.store.delete_cap(rk[0], rk[1])
            done()

        bf = ttk.Frame(win)
        bf.pack(pady=12)
        ttk.Button(bf, text="저장", command=save).pack(side="left", padx=4)
        if cur:
            ttk.Button(bf, text="설정 해제", command=clear).pack(side="left", padx=4)
        ttk.Button(bf, text="취소", command=win.destroy).pack(side="left", padx=4)

    def cap_changes(self):
        """저장된 재원 설정과 지금 읽은 PDF를 비교한다.
        changed: 원본 재원이 설정할 때와 달라진 행, missing: 이 PDF의 시간 범위 안인데 행이 없어진 설정."""
        allrows = {(r["time"], r["program"]): r for p in self.rows for r in self.rows[p]}
        times = sorted(t for t, _ in allrows)
        lo, hi = (times[0], times[-1]) if times else ("", "")
        changed, missing = [], []
        for (t, pr), c in sorted(self.caps.items()):
            row = allrows.get((t, pr))
            if row is None:
                if lo <= t <= hi:
                    missing.append((t, pr, c))
                continue
            orig = parse_amount(row)
            if orig and orig[1] != c["base_cap"]:
                changed.append((t, pr, c, orig))
        return changed, missing

    def alert_cap_changes(self):
        changed, missing = self.cap_changes()
        self.cap_flag = {(t, pr) for t, pr, _, _ in changed}
        if not changed and not missing:
            return
        self.draw_page()
        lines = []
        for t, pr, c, orig in changed[:10]:
            lines.append(f"· {t} {pr}: 원본 재원 {c['base_cap']} → {orig[1]}  (설정값 {c['cap']} 적용 중)")
        for t, pr, c in missing[:10]:
            lines.append(f"· {t} {pr}: 이 PDF에서 행을 찾을 수 없음  (설정값 {c['cap']})")
        msg = "재원 설정과 다른 내용이 있습니다.\n\n" + "\n".join(lines)
        if len(changed) > 10 or len(missing) > 10:
            msg += "\n… (일부 생략)"
        msg += "\n\n확인 후 필요하면 [이 행 재원 설정]에서 수정하거나 해제해 주세요.\n(원본 재원 변경은 이번 한 번만 알려 드립니다.)"
        messagebox.showwarning("재원 변경 알림", msg)
        for t, pr, c, orig in changed:     # 같은 변경을 매일 반복해서 알리지 않도록 기준값을 갱신
            try:
                self.store.set_cap(t, pr, c["cap"], orig[0], orig[1])
            except Exception:
                pass
        self.caps = self.store.list_caps()

    def _snapshot(self):
        return copy.deepcopy(self.entries)

    def _push(self):
        self.history.append(self._snapshot())
        if len(self.history) > 50:
            self.history.pop(0)
        self.dirty = True

    def undo_action(self):
        if not self.history:
            messagebox.showinfo("되돌리기", "더 이상 되돌릴 작업이 없습니다.")
            return
        self.entries = self.history.pop()
        self.dirty = True
        self.update_cell_info()
        self.draw_page()

    SORT_KEYS = {"이름순": "name", "초수순": "sec", "종료일순": "end"}
    SORT_NAMES = {"name": "이름순", "sec": "초수순", "end": "종료일순"}

    def _sort_presets(self):
        k, desc = self.sort_key, self.sort_desc
        if k == "sec":
            self.presets.sort(key=lambda p: (p["seconds"], p["name"].casefold()), reverse=desc)
        elif k == "end":
            have = sorted([p for p in self.presets if p.get("end_date")], key=lambda p: (p["end_date"], p["name"].casefold()), reverse=desc)
            none = sorted([p for p in self.presets if not p.get("end_date")], key=lambda p: p["name"].casefold())
            self.presets = have + none      # 종료일 없는 광고주는 항상 맨 뒤
        else:
            self.presets.sort(key=lambda p: p["name"].casefold(), reverse=desc)

    def _update_sort_btn(self):
        btn = getattr(self, "sort_dir_btn", None)
        if btn is not None:
            btn.config(text="▼ 내림차순" if self.sort_desc else "▲ 오름차순")

    def set_sort(self, key, toggle=False):
        if toggle or key == self.sort_key:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_key, self.sort_desc = key, False
        if hasattr(self, "sort_label_var"):
            self.sort_label_var.set(self.SORT_NAMES[self.sort_key])
        self._update_sort_btn()
        self._save_settings()
        self.refresh_presets()

    def alert_expired(self):
        """종료일이 오늘보다 앞선 광고주를 알려 준다."""
        if self.store is None:
            return
        try:
            exp = [p for p in self.store.list_presets() if is_expired(p.get("end_date", ""))]
        except Exception:
            return
        if not exp:
            return
        exp.sort(key=lambda p: (p["end_date"], p["name"]))
        lines = [f"· {p['name']}  (종료일 {short_date(p['end_date'])})" for p in exp[:15]]
        more = f"\n… 외 {len(exp) - 15}건" if len(exp) > 15 else ""
        messagebox.showwarning("종료된 광고주 알림", f"종료일이 지난 광고주가 {len(exp)}건 있습니다.\n\n" + "\n".join(lines) + more)

    def _selected_preset(self):
        tree = getattr(self, "preset_tree", None)
        if tree is None:
            return None
        sel = tree.selection()
        if not sel:
            return None
        pid = int(sel[0])
        return next((p for p in self.presets if p["id"] == pid), None)

    def _line_state(self, lk):
        """한 줄의 칸 상태. empty: 아직 쓸 수 있는 칸(빈 칸·왼쪽 반이 빈 칸), over: 이미 '+N' 으로 기록한 키들."""
        empty = []
        for k in range(4):
            cd = self.cell_by_key.get(lk + (k,))
            if cd is None or cd["text"]:
                continue
            if (lk + (k,)) not in self.entries:
                empty.append(k)
        over = [k for k in range(8, 12) if (lk + (k,)) in self.entries]
        return empty, over

    def _resolve_target(self, key):
        """선택한 칸에 넣을 때 실제로 기록할 칸 key. 넣을 수 없으면 안내하고 None."""
        cd = self.cell_by_key[key]
        lk, k = key[:4], key[4]
        empty, over = self._line_state(lk)
        if k >= 4:
            messagebox.showinfo("데이터 넣기", "이미 데이터가 있는 자리입니다.\n[추가한 내용 지우기]로 비우거나 다른 칸을 선택해 주세요.")
            return None
        # 차 있는 칸(원본 또는 이미 넣은 칸)을 골랐고 오른쪽에 빈칸이 있으면 그 빈칸으로 자동 이동
        if cd["text"] or key in self.entries:
            right = [x for x in empty if x > k]
            if right:
                self.selected_key = lk + (right[0],)
                self.update_cell_info()
                self.draw_page()
                return self.selected_key
        if cd["text"]:
            splittable = [x for x in range(4) if not self.cell_by_key[lk + (x,)]["text"] and (lk + (x + 4,)) not in self.entries]
            if empty or splittable:
                messagebox.showinfo("데이터 넣기", "원본에 이미 내용이 있는 칸입니다.\n비어 있는 칸이나 내가 넣은 칸을 선택해 주세요.")
                return None
            if len(over) >= OVERFLOW_MAX:
                messagebox.showinfo("데이터 넣기", f"이 행에는 더 넣을 자리가 없습니다. ('+N' 표시는 최대 {OVERFLOW_MAX}개)")
                return None
            if not messagebox.askyesno("칸이 부족합니다", "이 행에는 더 넣을 칸이 없습니다.\n기록만 저장하고 행 오른쪽에 '+N'으로 표시할까요?"):
                return None
            return lk + (next(i for i in range(8, 12) if (lk + (i,)) not in self.entries),)
        slot0, slot1 = key in self.entries, (lk + (k + 4,)) in self.entries
        if not slot0:
            return key
        if slot1:
            messagebox.showinfo("데이터 넣기", "이 칸은 이미 두 개가 들어 있습니다.\n하나를 지우거나 다른 칸을 선택해 주세요.")
            return None
        if empty:
            messagebox.showinfo("데이터 넣기", "비어 있는 칸이 남아 있습니다.\n먼저 빈 칸에 넣어 주세요. (칸이 모두 찼을 때만 한 칸을 반으로 나눕니다)")
            return None
        if not messagebox.askyesno("칸 나누기", "이 칸에는 이미 데이터가 있습니다.\n칸을 반으로 나눠 오른쪽에 하나 더 넣을까요?\n(이 행은 이름이 '초수 + 앞 두 글자'로 줄어 보입니다. 기록에는 전체 이름이 남습니다.)"):
            return None
        return lk + (k + 4,)

    def insert_selected_preset(self):
        if not self.doc:
            return
        key = self.selected_key
        if not key or key not in self.cell_by_key:
            messagebox.showinfo("데이터 넣기", "넣을 칸을 PDF에서 먼저 클릭해 주세요.")
            return
        preset = self._selected_preset()
        if not preset:
            messagebox.showinfo("데이터 넣기", "넣을 데이터를 오른쪽 목록에서 먼저 선택해 주세요.")
            return
        target = self._resolve_target(key)
        if target is None:
            return
        cd = self.cell_by_key[target]
        end = preset.get("end_date", "")
        if end and self.air_date > end:
            if not messagebox.askyesno("종료된 광고주", f"'{preset['name']}' 은(는) 종료일({short_date(end)})이 방송일({short_date(self.air_date)})보다 앞섭니다.\n그래도 넣을까요?"):
                return
        st = preset.get("start_date", "")
        if st and self.air_date < st:
            if not messagebox.askyesno("시작 전 광고주", f"'{preset['name']}' 은(는) 시작일({short_date(st)})이 방송일({short_date(self.air_date)})보다 늦습니다.\n그래도 넣을까요?"):
                return
        chk = self.row_check(cd["row"], extra_key=target, extra_sec=preset["seconds"])
        if chk and chk[3] > chk[1]:
            sold, cap, added, total = chk
            if not messagebox.askyesno("재원 초과", f"이 행의 재원은 {cap}초인데,\n판매 {sold}초 + 추가 {added}초 = {total}초로 {total - cap}초 초과합니다.\n그래도 넣을까요?"):
                return
        self._push()
        self.entries[target] = {"name": preset["name"], "seconds": preset["seconds"], "price": preset["price"], "grp": preset["grp"]}
        self.update_cell_info()
        self.draw_page()

    def clear_selected_cell(self):
        key = self.selected_key
        if not key or key not in self.entries:
            return
        self._push()
        del self.entries[key]
        self.update_cell_info()
        self.draw_page()

    # ------------------------------------------------------------------ 이름표·강조 표시 설정
    MARK_OPT_RANGE = {   # 값 이름: (최소, 최대, 한 칸)
        "fs": (3.5, 9.0, 0.5),
        "dy": (-8.0, 8.0, 0.5),
        "pad": (0.0, 4.0, 0.2),
        "hl_width": (0.3, 3.0, 0.1),
    }

    @staticmethod
    def _default_mark_opts():
        return {"fs": OVER_LABEL_FS, "pos": OVER_LABEL_POS if OVER_LABEL_POS in ("below", "above") else "below",
                "dy": OVER_LABEL_DY, "pad": OVER_LABEL_PAD, "fill": True, "hl_width": HL_WIDTH}

    @classmethod
    def _clean_mark_opts(cls, v):
        d = cls._default_mark_opts()
        if not isinstance(v, dict):
            return d
        for k, (lo, hi, _st) in cls.MARK_OPT_RANGE.items():
            try:
                d[k] = round(min(hi, max(lo, float(v.get(k, d[k])))), 2)
            except (TypeError, ValueError):
                pass
        if v.get("pos") in ("below", "above"):
            d["pos"] = v["pos"]
        if "fill" in v:
            d["fill"] = bool(v["fill"])
        return d

    def mark_opts_dialog(self, parent=None, on_change=None):
        """칸이 모자란 줄의 이름표(글자 크기·위치·바탕)와 강조 테두리 두께를 바꾼다. 바꾸는 즉시 화면에 반영된다.
        parent/on_change: 미리보기 창에서 열면 그 창을 다시 그리도록 넘겨받는다."""
        owner = parent or self.root
        win = tk.Toplevel(owner)
        win.title("🏷 표시 설정")
        win.transient(owner)
        win.resizable(False, False)
        win.grab_set()
        o = self.mark_opts
        fb = (self.ui_font_family, 9, "bold")
        vars_ = {k: tk.StringVar(value=str(o[k])) for k in self.MARK_OPT_RANGE}
        pos_v = tk.StringVar(value=o["pos"])
        fill_v = tk.BooleanVar(value=o["fill"])
        hl_v = tk.BooleanVar(value=self.highlight_var.get())
        pending = {"id": None}

        def apply(*_a):
            new = {"pos": pos_v.get(), "fill": bool(fill_v.get())}
            for k in self.MARK_OPT_RANGE:
                try:
                    new[k] = float(vars_[k].get())
                except ValueError:
                    return                         # 입력 중인 값(예: '-')은 다 쓸 때까지 기다린다
            self.mark_opts = self._clean_mark_opts(new)
            self.highlight_var.set(bool(hl_v.get()))
            self._save_settings()
            if self.doc:
                self.draw_page()
            if on_change:                          # 미리보기는 PDF를 다시 만들어야 하므로 잠깐 모았다가 한 번에
                if pending["id"]:
                    win.after_cancel(pending["id"])
                pending["id"] = win.after(350, on_change)

        def spin(parent_, key, label, unit, note=""):
            lo, hi, st = self.MARK_OPT_RANGE[key]
            r = ttk.Frame(parent_)
            r.pack(fill="x", pady=3)
            ttk.Label(r, text=label, width=16).pack(side="left")
            sp = ttk.Spinbox(r, from_=lo, to=hi, increment=st, width=6, textvariable=vars_[key], command=apply)
            sp.pack(side="left")
            sp.bind("<Return>", apply)
            sp.bind("<FocusOut>", apply)
            ttk.Label(r, text=unit, foreground="#555").pack(side="left", padx=(4, 8))
            if note:
                ttk.Label(r, text=note, foreground="#888", font=(self.ui_font_family, 8)).pack(side="left")

        lf = ttk.LabelFrame(win, text=" 칸이 모자란 줄의 이름표 ( + 20 설총, … ) ", padding=10)
        lf.pack(fill="x", padx=12, pady=(12, 6))
        spin(lf, "fs", "글자 크기", "pt", "원본 글자는 8pt")
        pr = ttk.Frame(lf)
        pr.pack(fill="x", pady=3)
        ttk.Label(pr, text="위치", width=16).pack(side="left")
        ttk.Radiobutton(pr, text="줄 아래 경계선", variable=pos_v, value="below", command=apply).pack(side="left")
        ttk.Radiobutton(pr, text="줄 위 경계선", variable=pos_v, value="above", command=apply).pack(side="left", padx=6)
        spin(lf, "dy", "위아래 미세 이동", "pt", "+ 는 아래로, - 는 위로")
        fr = ttk.Frame(lf)
        fr.pack(fill="x", pady=3)
        ttk.Label(fr, text="흰 바탕 채우기", width=16).pack(side="left")
        ttk.Checkbutton(fr, text="선·다른 글자 위에서도 읽히게 흰 바탕을 깐다", variable=fill_v, command=apply).pack(side="left")
        spin(lf, "pad", "바탕 크기(위아래 여백)", "pt", "줄이면 위아래 행을 덜 가린다")

        hf = ttk.LabelFrame(win, text=" 출력물 강조 ", padding=10)
        hf.pack(fill="x", padx=12, pady=6)
        ttk.Checkbutton(hf, text="저장·출력 PDF에서 추가한 칸·'+N'·이름표에 테두리 강조", variable=hl_v, command=apply).pack(anchor="w")
        spin(hf, "hl_width", "테두리 두께", "pt")

        ttk.Label(win, text="※ 바꾸면 화면에 바로 보입니다. 미리보기에서 열었다면 미리보기도 다시 그려집니다.",
                  foreground="#666", font=(self.ui_font_family, 8)).pack(anchor="w", padx=14)

        def reset():
            d = self._default_mark_opts()
            for k in self.MARK_OPT_RANGE:
                vars_[k].set(str(d[k]))
            pos_v.set(d["pos"])
            fill_v.set(d["fill"])
            apply()

        def close():
            apply()
            if pending["id"]:                       # 남은 미리보기 다시 그리기는 지금 한 번에
                win.after_cancel(pending["id"])
                pending["id"] = None
                if on_change:
                    on_change()
            win.destroy()
            if parent is not None:
                try:
                    parent.grab_set()              # 미리보기 창으로 다시 돌려준다
                except tk.TclError:
                    pass

        bf = ttk.Frame(win)
        bf.pack(pady=12)
        ttk.Button(bf, text="기본값으로", command=reset).pack(side="left", padx=4)
        ttk.Button(bf, text="닫기", command=close).pack(side="left", padx=4)
        win.protocol("WM_DELETE_WINDOW", close)

    # ------------------------------------------------------------------ 미리 만든 데이터 목록: 열 설정
    PRESET_COL_DEFS = {          # 열 id: (머리글, 기본 너비, 정렬, 남는 폭을 차지하는지)
        "name": ("이름", 120, "w", True),
        "sec": ("초", 34, "center", False),
        "end": ("종료일", 62, "center", False),
        "grp": ("그룹", 70, "w", False),
    }

    @classmethod
    def _default_preset_cols(cls):
        return {"order": list(cls.PRESET_COL_DEFS), "hidden": [], "widths": {}}

    @classmethod
    def _clean_preset_cols(cls, v):
        """설정 파일의 열 설정을 검사한다. 이상하면 기본값."""
        d = cls._default_preset_cols()
        if not isinstance(v, dict):
            return d
        order = [c for c in (v.get("order") or []) if c in cls.PRESET_COL_DEFS]
        order = list(dict.fromkeys(order)) + [c for c in cls.PRESET_COL_DEFS if c not in order]
        hidden = [c for c in (v.get("hidden") or []) if c in cls.PRESET_COL_DEFS and c != "name"]
        widths = {}
        for c, w in (v.get("widths") or {}).items():
            try:
                if c in cls.PRESET_COL_DEFS and 20 <= int(w) <= 600:
                    widths[c] = int(w)
            except (TypeError, ValueError):
                pass
        return {"order": order, "hidden": hidden, "widths": widths}

    @staticmethod
    def _preset_values(p, order):
        val = {"name": p["name"], "sec": p["seconds"], "end": short_date(p.get("end_date", "")), "grp": p["grp"]}
        return tuple(val[c] for c in order)

    def _apply_preset_cols(self):
        tree = getattr(self, "preset_tree", None)
        if tree is None:
            return
        try:
            if not tree.winfo_exists():
                return
        except tk.TclError:
            return
        pc = self.preset_cols
        tree["displaycolumns"] = [c for c in pc["order"] if c not in pc["hidden"]] or ["name"]
        for c, (_t, w, _a, stretch) in self.PRESET_COL_DEFS.items():
            tree.column(c, width=pc["widths"].get(c, w), stretch=stretch, minwidth=20)

    def _remember_preset_widths(self):
        """사용자가 머리글 경계선을 끌어 바꾼 열 너비를 기억한다."""
        tree = getattr(self, "preset_tree", None)
        if tree is None:
            return
        try:
            now = {c: int(tree.column(c, "width")) for c in self.PRESET_COL_DEFS if c not in self.preset_cols["hidden"]}
        except tk.TclError:
            return
        old = {c: self.preset_cols["widths"].get(c, self.PRESET_COL_DEFS[c][1]) for c in now}
        if now != old:
            self.preset_cols["widths"].update(now)
            self._save_settings()

    def _preset_header_menu(self, event):
        tree = self.preset_tree
        if tree.identify_region(event.x, event.y) not in ("heading", "separator"):
            return
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="⚙ 열 순서 / 보이기 설정…", command=self.preset_col_dialog)
        menu.add_separator()
        for c, (title, *_r) in self.PRESET_COL_DEFS.items():
            if c == "name":
                continue
            var = tk.BooleanVar(value=c not in self.preset_cols["hidden"])
            menu.add_checkbutton(label=f"{title} 보이기", variable=var, command=lambda c=c: self._toggle_preset_col(c))
        menu.add_separator()
        menu.add_command(label="기본값으로", command=self._reset_preset_cols)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def _toggle_preset_col(self, c):
        h = self.preset_cols["hidden"]
        if c in h:
            h.remove(c)
        elif c != "name":
            h.append(c)
        self._apply_preset_cols()
        self._save_settings()

    def _reset_preset_cols(self):
        self.preset_cols = self._default_preset_cols()
        self._apply_preset_cols()
        self._save_settings()

    def preset_col_dialog(self):
        """열 순서(▲▼)와 보이기를 바꾸는 작은 창. 바꾸는 즉시 목록에 반영된다."""
        win = tk.Toplevel(self.root)
        win.title("⚙ 목록 열 설정")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        ttk.Label(win, text="미리 만든 데이터 목록의 열 순서와 보이기를 정합니다.\n열 너비는 목록 머리글의 경계선을 끌면 저장됩니다.",
                  foreground="#555", justify="left").pack(anchor="w", padx=14, pady=(12, 6))
        box = ttk.Frame(win)
        box.pack(fill="x", padx=14)

        def move(c, d):
            o = self.preset_cols["order"]
            i = o.index(c)
            j = i + d
            if 0 <= j < len(o):
                o[i], o[j] = o[j], o[i]
                changed()

        def changed():
            self._apply_preset_cols()
            self._save_settings()
            draw()

        def draw():
            for w in box.winfo_children():
                w.destroy()
            order = self.preset_cols["order"]
            for i, c in enumerate(order):
                title = self.PRESET_COL_DEFS[c][0]
                r = ttk.Frame(box)
                r.pack(fill="x", pady=2)
                ttk.Label(r, text=f"{i + 1}.", width=3).pack(side="left")
                if c == "name":
                    ttk.Label(r, text=f"{title}  (항상 보임)", width=18).pack(side="left")
                else:
                    v = tk.BooleanVar(value=c not in self.preset_cols["hidden"])
                    ttk.Checkbutton(r, text=title, variable=v, width=16, command=lambda c=c: (self._toggle_preset_col(c), draw())).pack(side="left")
                b1 = ttk.Button(r, text="▲", width=3, command=lambda c=c: move(c, -1))
                b2 = ttk.Button(r, text="▼", width=3, command=lambda c=c: move(c, 1))
                b1.pack(side="left", padx=2)
                b2.pack(side="left")
                if i == 0:
                    b1.state(["disabled"])
                if i == len(order) - 1:
                    b2.state(["disabled"])

        draw()
        bf = ttk.Frame(win)
        bf.pack(pady=12)
        ttk.Button(bf, text="기본값으로", command=lambda: (self._reset_preset_cols(), draw())).pack(side="left", padx=4)
        ttk.Button(bf, text="닫기", command=win.destroy).pack(side="left", padx=4)

    # ------------------------------------------------------------------ 미리 만든 데이터
    def refresh_presets(self):
        self._reload_preset_map()
        tree = getattr(self, "preset_tree", None)
        if self.store is None:
            self.presets = []
            return
        try:
            self.presets = self.store.list_presets(self.filter_var.get().strip())
        except Exception as exc:
            self.presets = []
            messagebox.showerror("목록 불러오기 실패", str(exc))
        if tree is None:
            return
        try:
            if not tree.winfo_exists():
                return
        except tk.TclError:
            return
        self._sort_presets()
        keep = tree.selection()
        for i in tree.get_children():
            tree.delete(i)
        for p in self.presets:
            tree.insert("", "end", iid=str(p["id"]), values=self._preset_values(p, tuple(self.PRESET_COL_DEFS)),
                        tags=("expired",) if is_expired(p.get("end_date", "")) else ())
        if keep and tree.exists(keep[0]):
            tree.selection_set(keep[0])

    SECONDS_CHOICES = ("10", "20", "40", "50", "60", "70")

    def _preset_dialog(self, preset=None):
        win = tk.Toplevel(self.root)
        win.title("데이터 수정" if preset else "데이터 추가")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        fb = (self.ui_font_family, 9, "bold")
        name_v = tk.StringVar(value=(preset.get("base") or preset["name"]) if preset else "")
        var_v = tk.StringVar(value=preset.get("variant", "") if preset else "")
        sec_v = tk.StringVar(value=str(preset["seconds"]) if preset else "20")
        grp_v = tk.StringVar(value=preset["grp"] if preset else "")
        start_v = tk.StringVar(value=short_date(preset["start_date"]) if preset and preset.get("start_date") else ("" if preset else short_date(today_iso())))
        end_v = tk.StringVar(value=short_date(preset["end_date"]) if preset and preset.get("end_date") else "")
        am_v = tk.StringVar(value=str(preset.get("min_am", 0)) if preset else "0")
        fm_v = tk.StringVar(value=str(preset.get("min_fm", 0)) if preset else "0")

        ttk.Label(win, text="이름(광고주):", font=fb).grid(row=0, column=0, sticky="e", padx=(14, 4), pady=6)
        ttk.Entry(win, textvariable=name_v, width=30).grid(row=0, column=1, columnspan=2, sticky="w", padx=(0, 14), pady=6)

        ttk.Label(win, text="구분(선택):", font=fb).grid(row=1, column=0, sticky="e", padx=(14, 4), pady=6)
        ttk.Entry(win, textvariable=var_v, width=14).grid(row=1, column=1, sticky="w", pady=6)
        ttk.Label(win, text="이름이 같은 다른 광고를 구분할 때만 적습니다. 예: 호남대학교 + 입학 → '호남대학교(입학)'",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=260, justify="left").grid(row=1, column=2, sticky="w", padx=(6, 14))

        short_v = tk.StringVar(value=preset.get("short_name", "") if preset else "")
        ttk.Label(win, text="약칭(선택):", font=fb).grid(row=2, column=0, sticky="e", padx=(14, 4), pady=6)
        ttk.Entry(win, textvariable=short_v, width=10).grid(row=2, column=1, sticky="w", pady=6)
        ttk.Label(win, text="한 행에 광고가 5개 이상일 때 PDF에 보이는 줄임말. 비우면 앞 두 글자(예: 주식회사 설총 → 설총)", font=(self.ui_font_family, 8), foreground="#666",
                  wraplength=260, justify="left").grid(row=2, column=2, sticky="w", padx=(6, 14))

        ttk.Label(win, text="초수:", font=fb).grid(row=3, column=0, sticky="e", padx=(14, 4), pady=6)
        ttk.Combobox(win, textvariable=sec_v, values=self.SECONDS_CHOICES, width=8).grid(row=3, column=1, sticky="w", pady=6)
        ttk.Label(win, text="목록에서 고르거나 직접 입력", font=(self.ui_font_family, 8), foreground="#666").grid(row=3, column=2, sticky="w", padx=(6, 14))

        ttk.Label(win, text="그룹:", font=fb).grid(row=4, column=0, sticky="e", padx=(14, 4), pady=6)
        gcombo = ttk.Combobox(win, textvariable=grp_v, state="readonly", width=20)
        gcombo.grid(row=4, column=1, sticky="w", pady=6)
        gb = ttk.Frame(win)
        gb.grid(row=4, column=2, sticky="w", padx=(6, 14))

        def reload_groups(select=None):
            groups = self.store.list_groups()
            gcombo.configure(values=[""] + groups)
            if select is not None:
                grp_v.set(select)
            elif grp_v.get() not in groups:
                grp_v.set("")

        def add_group():
            nm = simpledialog.askstring("그룹 추가", "새 그룹 이름:", parent=win)
            if nm is None:
                return
            try:
                self.store.add_group(nm)
            except ValueError as exc:
                messagebox.showwarning("그룹 추가", str(exc), parent=win)
                return
            reload_groups(nm.strip())

        def del_group():
            nm = grp_v.get()
            if not nm:
                messagebox.showinfo("그룹 삭제", "삭제할 그룹을 먼저 선택해 주세요.", parent=win)
                return
            if messagebox.askyesno("그룹 삭제", f"'{nm}' 그룹을 삭제할까요?\n이 그룹을 쓰던 미리 만든 데이터는 '그룹 없음'이 됩니다.\n(이미 저장된 기록은 그대로 남습니다.)", parent=win):
                self.store.delete_group(nm)
                reload_groups("")
                self.refresh_presets()

        ttk.Button(gb, text="그룹 추가", width=9, command=add_group).pack(side="left", padx=(0, 3))
        ttk.Button(gb, text="그룹 삭제", width=9, command=del_group).pack(side="left")
        reload_groups(preset["grp"] if preset else "")

        def date_row(r, label, var):
            ttk.Label(win, text=label, font=fb).grid(row=r, column=0, sticky="e", padx=(14, 4), pady=6)
            fr = ttk.Frame(win)
            fr.grid(row=r, column=1, columnspan=2, sticky="w", pady=6)
            ttk.Entry(fr, textvariable=var, width=12).pack(side="left")

            def open_cal():
                try:
                    cur = parse_date_input(var.get())
                except ValueError:
                    cur = ""
                pick_date(win, cur, lambda iso: var.set(short_date(iso)))

            ttk.Button(fr, text="달력", width=6, command=open_cal).pack(side="left", padx=4)
            ttk.Button(fr, text="지우기", width=6, command=lambda: var.set("")).pack(side="left")

        date_row(5, "시작일:", start_v)
        date_row(6, "종료일:", end_v)
        ttk.Label(win, text="계약 기간입니다. 시작일은 기본 오늘, 종료일 다음 날부터 목록에 빨간색으로 표시되고 프로그램을 켤 때와 PDF를 열 때 알려 드립니다. 비워 두면 기간 제한 없이 계속 사용합니다.   예) 26/10/31",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=420, justify="left").grid(row=7, column=0, columnspan=3, sticky="w", padx=14)
        status = tk.Label(win, text="", font=(self.ui_font_family, 9, "bold"), anchor="w")
        status.grid(row=8, column=0, columnspan=3, sticky="w", padx=14, pady=(2, 0))

        def update_status(*_):
            try:
                iso = parse_date_input(end_v.get())
                parse_date_input(start_v.get())
            except ValueError:
                status.config(text="날짜 형식이 맞지 않습니다. 달력에서 선택해 주세요", fg="#dc2626")
                return
            if not iso:
                status.config(text="종료일 없음 – 계속 사용", fg="#6b7280")
                return
            diff = (datetime.date.fromisoformat(iso) - datetime.date.today()).days
            if diff > 0:
                status.config(text=f"{short_date(iso)} 까지 사용 (D-{diff})", fg="#15803d")
            elif diff == 0:
                status.config(text="오늘이 마지막 날입니다", fg="#d97706")
            else:
                status.config(text="이미 종료된 날짜입니다. 저장하면 목록에 빨간색으로 표시됩니다", fg="#dc2626")

        end_v.trace_add("write", update_status)
        start_v.trace_add("write", update_status)
        update_status()

        ttk.Label(win, text="하루 최소 횟수:", font=fb).grid(row=9, column=0, sticky="e", padx=(14, 4), pady=6)
        mr = ttk.Frame(win)
        mr.grid(row=9, column=1, columnspan=2, sticky="w", pady=6)
        ttk.Label(mr, text="AM").pack(side="left")
        ttk.Spinbox(mr, from_=0, to=99, width=4, textvariable=am_v).pack(side="left", padx=(3, 12))
        ttk.Label(mr, text="FM").pack(side="left")
        ttk.Spinbox(mr, from_=0, to=99, width=4, textvariable=fm_v).pack(side="left", padx=3)
        ttk.Label(win, text="계약에 정해진 하루 최소 방송 횟수입니다 (예: AM 3, FM 3). 0이면 점검하지 않습니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=420, justify="left").grid(row=10, column=0, columnspan=3, sticky="w", padx=14)

        def ok():
            try:
                sec = int(sec_v.get().strip())
                if sec <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("입력 확인", "초수는 1 이상의 숫자로 입력해 주세요.", parent=win)
                return
            try:
                start_iso = parse_date_input(start_v.get())
                end_iso = parse_date_input(end_v.get())
            except ValueError:
                messagebox.showwarning("입력 확인", "날짜 형식이 맞지 않습니다.\n달력에서 고르거나 26/10/31 형식으로 입력해 주세요.", parent=win)
                return
            if start_iso and end_iso and end_iso < start_iso:
                messagebox.showwarning("입력 확인", "종료일이 시작일보다 앞섭니다.", parent=win)
                return
            try:
                mam, mfm = int(am_v.get().strip() or "0"), int(fm_v.get().strip() or "0")
                if mam < 0 or mfm < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("입력 확인", "최소 횟수는 0 이상의 숫자로 입력해 주세요.", parent=win)
                return
            try:
                if preset:
                    self.store.update_preset(preset["id"], name_v.get(), sec, preset.get("price", 0), grp_v.get(), end_iso, start_iso, mam, mfm, short_v.get(), var_v.get())
                else:
                    self.store.add_preset(name_v.get(), sec, 0, grp_v.get(), end_iso, start_iso, mam, mfm, short_v.get(), var_v.get())
            except ValueError as exc:
                messagebox.showwarning("입력 확인", str(exc), parent=win)
                return
            win.destroy()
            self.refresh_presets()
            self.update_count_label()

        bf = ttk.Frame(win)
        bf.grid(row=11, column=0, columnspan=3, pady=12)
        ttk.Button(bf, text="저장", command=ok).pack(side="left", padx=4)
        ttk.Button(bf, text="취소", command=win.destroy).pack(side="left", padx=4)

    def add_preset_dialog(self):
        if self.store:
            self._preset_dialog(None)

    def edit_preset_dialog(self):
        p = self._selected_preset()
        if not p:
            messagebox.showinfo("수정", "수정할 데이터를 목록에서 먼저 선택해 주세요.")
            return
        self._preset_dialog(p)

    def delete_preset_action(self):
        p = self._selected_preset()
        if not p:
            messagebox.showinfo("삭제", "삭제할 데이터를 목록에서 먼저 선택해 주세요.")
            return
        if messagebox.askyesno("삭제 확인", f"'{p['name']}' 을(를) 목록에서 삭제할까요?\n(이미 PDF에 넣었거나 기록된 내용은 그대로 남습니다.)"):
            self.store.delete_preset(p["id"])
            self.refresh_presets()

    # ------------------------------------------------------------------ 읽은 데이터 표
    def show_today_list(self):
        """지금 탭에서 오늘 넣거나 고친 것 / 오늘 지운 것 목록. 더블클릭하면 그 칸으로 간다."""
        if not self.doc:
            return
        win = tk.Toplevel(self.root)
        win.title(f"📝 오늘 수정 목록 — {self._tab_title(self.media, self.air_date)}  ({today_iso()})")
        win.geometry("760x520")
        cols = ("kind", "when", "page", "time", "program", "where", "data")
        heads = ("구분", "시각", "쪽", "시분초", "프로그램명", "자리", "데이터")
        widths = (70, 70, 40, 80, 200, 90, 200)
        tf = ttk.Frame(win)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c_, h, w in zip(cols, heads, widths):
            tree.heading(c_, text=h)
            tree.column(c_, width=w, anchor="w" if c_ in ("program", "data") else "center")
        vs = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        tree.tag_configure("add", foreground="#15803d")
        tree.tag_configure("new", foreground="#b45309")
        tree.tag_configure("del", foreground="#6b7280")

        def where(k):
            if k < 4:
                return f"{k + 1}칸"
            if k < 8:
                return f"{k - 3}칸 오른쪽"
            return f"+N ({k - 7})"

        rows, keys = [], {}
        for key, ent in self.entries.items():
            if not self.is_today_entry(ent):
                continue
            sa = ent.get("saved_at") or ""
            rows.append((("new" if not sa else "add"), ("저장 전" if not sa else sa[11:16]), key, self.entry_text(ent)))
        for r in getattr(self, "removed", []) or []:
            key = (r["page"], r["time"], r["program"], r.get("line") or 0, int(r["cell"]) - 1)
            rows.append(("del", r["removed_at"][11:16], key, f"{r['seconds']} {r['name']}"))
        rows.sort(key=lambda t: (t[2][0] if t[2][0] is not None else 999, t[2][1], t[2][3], t[2][4]))
        label = {"add": "넣음/고침", "new": "저장 전", "del": "지움"}
        for kind, when, key, data in rows:
            p, tm, prog, li, k = key
            iid = tree.insert("", "end", values=(label[kind], when, (p + 1) if p is not None else "", tm, prog,
                                                  where(k) + (" (둘째 줄)" if li else ""), data), tags=(kind,))
            keys[iid] = key

        def goto(event=None):
            sel = tree.selection()
            if not sel:
                return
            key = keys[sel[0]]
            cd = self.cell_by_key.get(key)
            if cd is None:
                messagebox.showinfo("이동", "이 운행표에서 그 칸을 찾지 못했습니다.", parent=win)
                return
            self.goto_page(key[0])
            self.select_cell(key)
            self._scroll_to_cell(cd)

        tree.bind("<Double-Button-1>", goto)
        n_add = sum(1 for r in rows if r[0] != "del")
        n_del = sum(1 for r in rows if r[0] == "del")
        ttk.Label(win, text=f"오늘 넣거나 고친 것 {n_add}건 · 오늘 지운 것 {n_del}건   — 더블클릭하면 그 칸으로 이동합니다."
                            "\n화면에서 초록 바탕은 오늘 넣거나 고친 칸, 회색 취소선은 오늘 지운 칸입니다. (저장 PDF에는 나오지 않습니다)",
                  font=(self.ui_font_family, 8), foreground="#555", justify="left").pack(anchor="w", padx=8, pady=4)
        if not rows:
            tree.insert("", "end", values=("", "", "", "", "오늘 수정한 내용이 없습니다.", "", ""))

    def show_table(self):
        if not self.doc:
            return
        win = tk.Toplevel(self.root)
        win.title(f"📋 읽은 데이터 표 — {self.air_date}")
        win.geometry("1250x720")
        cols = ("page", "time", "program", "form", "kind", "ad1", "ad2", "ad3", "ad4", "amount")
        heads = ("쪽", "시분초", "프로그램명", "형태", "녹/생", "광고주1", "광고주2", "광고주3", "광고주4", "판매/재원")
        widths = (40, 80, 230, 60, 50, 150, 150, 150, 150, 90)
        tf = ttk.Frame(win)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c_, h, w in zip(cols, heads, widths):
            tree.heading(c_, text=h)
            tree.column(c_, width=w, anchor="center" if c_ in ("page", "time", "form", "kind") else "w")
        vs = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(tf, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        tree.tag_configure("added", foreground="#1d4ed8")
        tree.tag_configure("over", foreground="#dc2626")

        first_of = {}
        for p in sorted(self.rows):
            for row in self.rows[p]:
                for li, line in enumerate(row["lines"]):
                    shown, added = [], False
                    lk = (p, row["time"], row["program"], li)
                    for k in range(4):
                        parts = [("＋ " + self.entry_text(e)) for e in (self.entries.get(lk + (k,)), self.entries.get(lk + (k + 4,))) if e]
                        if parts:
                            shown.append("  /  ".join(parts))
                            added = True
                        else:
                            shown.append(line["ads"][k])
                    ov = [self.entry_text(self.entries[lk + (i,)]) for i in range(8, 12) if (lk + (i,)) in self.entries]
                    if ov:
                        shown[3] = (shown[3] + "   " if shown[3] else "") + f"(+{len(ov)}: " + ", ".join(ov) + ")"
                        added = True
                    amt = re.sub(r"\s*/\s*", " / ", line["amount"]) if line["amount"] else ""
                    chk = self.row_check(row)
                    over = self.is_over(chk)
                    custom = self.row_is_custom(row)
                    if li == 0 and chk and (chk[2] > 0 or custom):
                        amt = f"{chk[3]} / {chk[1]}" + (f"  (+{chk[2]})" if chk[2] > 0 else "") + ("  [설정]" if custom else "")
                    iid = tree.insert("", "end", values=(p + 1, row["time"] if li == 0 else "", row["program"] if li == 0 else "",
                                                          row["form"] if li == 0 else "", row["kind"] if li == 0 else "", *shown, amt),
                                      tags=("over",) if over else ("added",) if added else ())
                    first_of[iid] = (p, row, li)

        def goto(event=None):
            sel = tree.selection()
            if not sel:
                return
            p, row, li = first_of[sel[0]]
            k = next((k for k in range(4) if not row["lines"][li]["ads"][k] and (p, row["time"], row["program"], li, k) not in self.entries), 0)
            key = (p, row["time"], row["program"], li, k)
            self.cur = p
            self.page_var.set(str(p + 1))
            self.selected_key = key
            self.update_cell_info()
            self.draw_page()
            self._scroll_to_cell(self.cell_by_key[key])

        tree.bind("<Double-Button-1>", goto)
        ttk.Label(win, text="행을 더블클릭하면 PDF 화면의 해당 칸으로 이동합니다.   파란 글씨(＋)는 새로 추가한 데이터, 빨간 글씨는 추가한 초가 재원을 넘은 행입니다.",
                  font=(self.ui_font_family, 8), foreground="#555").pack(anchor="w", padx=8, pady=4)

    # ------------------------------------------------------------------ 정산 / 검색
    def show_report(self):
        if self.store is None:
            messagebox.showerror("정산 / 검색", "데이터베이스를 열 수 없습니다.\n⚙️ 설정에서 데이터 파일 위치를 확인해 주세요.")
            return
        win = tk.Toplevel(self.root)
        win.title("💰 정산 / 검색")
        win.geometry("980x640")
        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        fnt = (self.ui_font_family, 9, "bold")
        today = self.air_date or time.strftime("%Y-%m-%d")
        defaults = {"일": today, "월": today[:7], "년": today[:4]}

        def make_tree(parent):
            tf = ttk.Frame(parent)
            tf.pack(fill="both", expand=True, pady=4)
            tree = ttk.Treeview(tf, show="headings")
            vs = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vs.set)
            tree.pack(side="left", fill="both", expand=True)
            vs.pack(side="right", fill="y")
            return tree

        def fill(tree, cols, rows, numeric_from=1):
            tree.configure(columns=[c for c, _ in cols])
            for c_, w in cols:
                tree.heading(c_, text=c_)
                tree.column(c_, width=w, anchor="e" if c_ in ("건수", "총 초수", "금액", "초수") else "w")
            for i in tree.get_children():
                tree.delete(i)
            for r in rows:
                tree.insert("", "end", values=r)

        def save_csv(cols, rows):
            if not rows:
                messagebox.showinfo("CSV 저장", "저장할 내용이 없습니다.", parent=win)
                return
            path = filedialog.asksaveasfilename(title="CSV 저장", defaultextension=".csv", filetypes=[("CSV 파일", "*.csv")], parent=win)
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8-sig", newline="") as f:
                    w = csv.writer(f)
                    w.writerow([c for c, _ in cols])
                    w.writerows(rows)
                messagebox.showinfo("CSV 저장", f"저장했습니다:\n{path}", parent=win)
            except Exception as exc:
                messagebox.showerror("CSV 저장 실패", str(exc), parent=win)

        def cal_btn(parent, var, keep=10):
            """입력칸 옆 📅 버튼: 달력에서 날짜를 고르면 입력칸에 넣는다. keep=10 날짜, 7 연-월(월 정산), 4 연도(년 정산)."""
            def open_cal():
                cur = (var.get() or "").strip()
                if re.fullmatch(r"\d{4}-\d{2}", cur):
                    cur += "-01"
                elif re.fullmatch(r"\d{4}", cur):
                    cur += "-01-01"
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", cur):
                    cur = ""
                pick_date(win, cur, lambda iso: var.set(iso[:keep]))
            ttk.Button(parent, text="📅", width=3, command=open_cal).pack(side="left")

        # --- 일/월/년 정산 탭
        for kind in ("일", "월", "년"):
            tab = ttk.Frame(nb, padding=8)
            nb.add(tab, text=f" {kind} 정산 ")
            top = ttk.Frame(tab)
            top.pack(fill="x")
            ttk.Label(top, text={"일": "기준일:", "월": "기준월:", "년": "기준연도:"}[kind], font=fnt).pack(side="left")
            kvar = tk.StringVar(value=defaults[kind])
            ttk.Entry(top, textvariable=kvar, width=12).pack(side="left", padx=(4, 0))
            cal_btn(top, kvar, {"일": 10, "월": 7, "년": 4}[kind])
            ttk.Label(top, text={"일": "(2026-10-30)", "월": "(2026-10)", "년": "(2026)"}[kind], foreground="#666").pack(side="left")
            ttk.Label(top, text="   묶음 기준:", font=fnt).pack(side="left")
            bvar = tk.StringVar(value="광고주별")
            ttk.Combobox(top, textvariable=bvar, values=list(BY_OPTIONS), state="readonly", width=9).pack(side="left", padx=4)
            ttk.Label(top, text="매체:", font=fnt).pack(side="left")
            mvar = tk.StringVar(value="전체")
            ttk.Combobox(top, textvariable=mvar, values=["전체", "AM", "FM"], state="readonly", width=5).pack(side="left", padx=4)
            tree = make_tree(tab)
            total_lbl = ttk.Label(tab, text="", font=fnt)
            total_lbl.pack(anchor="e")
            state = {"cols": [], "rows": []}

            def run(kind=kind, kvar=kvar, bvar=bvar, mvar=mvar, tree=tree, total_lbl=total_lbl, state=state):
                try:
                    d0, d1 = period_range(kind, kvar.get())
                except ValueError as exc:
                    messagebox.showwarning("정산", str(exc), parent=win)
                    return
                by = BY_OPTIONS[bvar.get()]
                data = self.store.aggregate(d0, d1, by=by, media="" if mvar.get() == "전체" else mvar.get())
                cols = [(bvar.get().replace("별", ""), 300), ("건수", 90), ("총 초수", 110)]
                rows = [(k or "(없음)", cnt, fmt_num(sec)) for k, cnt, sec, amt in data]
                fill(tree, cols, rows)
                state["cols"], state["rows"] = cols, rows
                total_lbl.config(text=f"합계  건수 {fmt_num(sum(r[1] for r in data))}   총 초수 {fmt_num(sum(r[2] for r in data))}")

            ttk.Button(top, text="조회", command=run).pack(side="left", padx=6)
            ttk.Button(top, text="CSV 저장", command=lambda s=state: save_csv(s["cols"], s["rows"])).pack(side="right")
            run()

        # --- 검색 탭
        tab = ttk.Frame(nb, padding=8)
        nb.add(tab, text=" 검색 / 그룹 검색 ")
        r1 = ttk.Frame(tab)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="검색어:", font=fnt).pack(side="left")
        kw = tk.StringVar()
        kw_e = ttk.Entry(r1, textvariable=kw, width=22)
        kw_e.pack(side="left", padx=4)
        kw_e.bind("<Return>", lambda e: do_search())
        ttk.Label(r1, text="찾는 곳:", font=fnt).pack(side="left", padx=(6, 0))
        tgt = tk.StringVar(value="광고주만")
        ttk.Combobox(r1, textvariable=tgt, values=list(SEARCH_TARGETS), state="readonly", width=13).pack(side="left", padx=4)
        ttk.Label(r1, text="매체:", font=fnt).pack(side="left", padx=(10, 0))
        sm = tk.StringVar(value="전체")
        ttk.Combobox(r1, textvariable=sm, values=["전체", "AM", "FM"], state="readonly", width=5).pack(side="left", padx=4)
        ttk.Label(r1, text="그룹:", font=fnt).pack(side="left", padx=(10, 0))
        gv = tk.StringVar()
        gcombo = ttk.Combobox(r1, textvariable=gv, values=[""] + self.store.distinct_groups(), width=14)
        gcombo.pack(side="left", padx=4)
        r2 = ttk.Frame(tab)
        r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="기간:", font=fnt).pack(side="left")
        df = tk.StringVar()
        dt = tk.StringVar()
        ttk.Entry(r2, textvariable=df, width=12).pack(side="left", padx=(4, 0))
        cal_btn(r2, df)
        ttk.Label(r2, text=" ~").pack(side="left")
        ttk.Entry(r2, textvariable=dt, width=12).pack(side="left", padx=(4, 0))
        cal_btn(r2, dt)
        ttk.Label(r2, text="(비우면 전체 기간, 기간 안의 기록만 나옵니다)", foreground="#666").pack(side="left")
        ttk.Label(r2, text="   묶어서 보기:", font=fnt).pack(side="left")
        sby = tk.StringVar(value="묶지 않음")
        ttk.Combobox(r2, textvariable=sby, values=["묶지 않음"] + list(BY_OPTIONS), state="readonly", width=10).pack(side="left", padx=4)
        stree = make_tree(tab)
        slbl = ttk.Label(tab, text="", font=fnt)
        slbl.pack(anchor="e")
        sstate = {"cols": [], "rows": []}

        def do_search():
            d0, d1 = df.get().strip(), dt.get().strip()
            for v in (d0, d1):
                if v and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    messagebox.showwarning("검색", "기간은 2026-10-30 형식으로 입력해 주세요.", parent=win)
                    return
            gcombo.configure(values=[""] + self.store.distinct_groups())
            if sby.get() == "묶지 않음":
                data = self.store.detail(d0, d1, kw.get().strip(), gv.get().strip(), media="" if sm.get() == "전체" else sm.get(),
                                         target=SEARCH_TARGETS[tgt.get()])
                cols = [("방송일", 95), ("매체", 50), ("시분초", 80), ("프로그램", 220), ("칸", 40), ("광고주", 200), ("초수", 60), ("그룹", 100)]
                rows = [(a, m, t, p, cell_label(c_), n, s, g) for a, m, t, p, c_, n, s, pr, g in data]
                fill(stree, cols, rows)
                slbl.config(text=f"{len(rows)}건   총 초수 {fmt_num(sum(r[6] for r in data))}")
            else:
                data = self.store.aggregate(d0, d1, kw.get().strip(), gv.get().strip(), BY_OPTIONS[sby.get()], media="" if sm.get() == "전체" else sm.get(),
                                            target=SEARCH_TARGETS[tgt.get()])
                cols = [(sby.get().replace("별", ""), 300), ("건수", 90), ("총 초수", 110)]
                rows = [(k or "(없음)", cnt, fmt_num(sec)) for k, cnt, sec, amt in data]
                fill(stree, cols, rows)
                slbl.config(text=f"{len(rows)}개 묶음   건수 {fmt_num(sum(r[1] for r in data))}")
            sstate["cols"], sstate["rows"] = cols, rows

        ttk.Button(r2, text="검색", command=do_search).pack(side="left", padx=6)
        ttk.Button(r2, text="CSV 저장", command=lambda: save_csv(sstate["cols"], sstate["rows"])).pack(side="right")
        do_search()

        # --- 광고주별 일정표 탭
        tab = ttk.Frame(nb, padding=8)
        nb.add(tab, text=" 광고주별 일정표 ")
        # 일정표는 (이름, 초수) 단위. 목록에는 '60초 이름' 형태로 보인다
        keys = sorted({(p["name"], p["seconds"]) for p in self.store.list_presets()} | set(self.store.record_keys()), key=lambda t: (t[0], t[1]))
        key_of_label = {f"{sec}초 {nm}": (nm, sec) for nm, sec in keys}
        secs_of = {}
        for nm, sec in keys:
            secs_of.setdefault(nm, set()).add(sec)
        multi = sorted(nm for nm, v in secs_of.items() if len(v) >= 2)
        for nm in multi:                       # 초수가 둘 이상인 광고주는 '전체 초수'(합산) 보기도 고를 수 있다
            key_of_label[f"{nm} (전체 초수)"] = (nm, None)
        names = list(key_of_label)
        q1 = ttk.Frame(tab)
        q1.pack(fill="x", pady=2)
        ttk.Label(q1, text="광고:", font=fnt).pack(side="left")
        adv = tk.StringVar(value=names[0] if names else "")
        adv_cb = ttk.Combobox(q1, textvariable=adv, values=names, width=36)
        adv_cb.pack(side="left", padx=4)
        ttk.Label(q1, text="  기간:", font=fnt).pack(side="left")
        sf, ef = tk.StringVar(), tk.StringVar()
        ttk.Entry(q1, textvariable=sf, width=12).pack(side="left", padx=(2, 0))
        cal_btn(q1, sf)
        ttk.Label(q1, text=" ~").pack(side="left")
        ttk.Entry(q1, textvariable=ef, width=12).pack(side="left", padx=(2, 0))
        cal_btn(q1, ef)
        ttk.Label(q1, text="(비우면 광고의 계약 기간으로 자동, 형식 2026-08-01)", foreground="#666").pack(side="left")
        q3 = ttk.Frame(tab)
        q3.pack(fill="x", pady=2)
        with_contract = tk.BooleanVar(value=False)
        list_info = ttk.Label(q3, text="", foreground="#1d4ed8")

        def period_or_none():
            """입력한 기간. 다 쓰지 않은(형식이 틀린) 값이 있으면 None."""
            a, b = sf.get().strip(), ef.get().strip()
            for v in (a, b):
                if v and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    return None
            return a, b

        def labels_for(a, b):
            """목록에 보일 광고. 기간이 없으면 전체, 있으면 그 기간에 기록이 있는 광고만
            ([기록 없는 계약 광고도 보기]를 켜면 계약 기간이 겹치는 광고도)."""
            if not a and not b:
                return names
            ks = set(self.store.keys_in_period(a, b))
            if with_contract.get():
                lo, hi = a or "0000-01-01", b or "9999-12-31"
                for pr_ in self.store.list_presets():
                    st_, en_ = pr_.get("start_date") or "", pr_.get("end_date") or ""
                    if (st_ or en_) and (not st_ or st_ <= hi) and (not en_ or en_ >= lo):
                        ks.add((pr_["name"], pr_["seconds"]))
            secs = {}
            for nm_, sec_ in ks:
                secs.setdefault(nm_, set()).add(sec_)
            out = [f"{sec_}초 {nm_}" for nm_, sec_ in sorted(ks, key=lambda t: (t[0], t[1]))]
            out += [f"{nm_} (전체 초수)" for nm_ in sorted(secs) if len(secs[nm_]) >= 2]
            return [x for x in out if x in key_of_label]

        def refresh_list(*_a):
            per = period_or_none()
            if per is None:
                return
            lst = labels_for(*per)
            adv_cb["values"] = lst
            if per == ("", ""):
                list_info.config(text=f"전체 광고 {len(lst)}개", foreground="#555")
            else:
                list_info.config(text=f"이 기간에 조회되는 광고 {len(lst)}개" + ("" if lst else " — 기간 안의 기록이 없습니다"),
                                 foreground="#1d4ed8" if lst else "#dc2626")
            if adv.get() not in lst:
                adv.set(lst[0] if lst else "")

        ttk.Checkbutton(q3, text="기록 없는 계약 광고도 보기 (계약 기간이 겹치는 광고)", variable=with_contract, command=refresh_list).pack(side="left")
        list_info.pack(side="left", padx=12)
        sf.trace_add("write", refresh_list)
        ef.trace_add("write", refresh_list)
        q2 = ttk.Frame(tab)
        q2.pack(fill="x", pady=2)
        title_lbl = ttk.Label(tab, text="", font=(self.ui_font_family, 11, "bold"))
        stree2 = make_tree(tab)
        stree2.tag_configure("out", background="#d9d9d9", foreground="#6b7280")
        stree2.tag_configure("short", background="#fff2cc")
        sum_lbl = ttk.Label(tab, text="", font=fnt)
        sum_lbl.pack(anchor="e")
        ttk.Label(tab, text="초수가 여러 개인 광고주는 '(전체 초수)'를 고르면 합산해서 봅니다 (칸에 0759(20) 처럼 초수 표시, 하루 최소 횟수는 합쳐서 점검).",
                  font=(self.ui_font_family, 8), foreground="#666").pack(anchor="w")
        ttk.Label(tab, text="회색: 계약 기간 밖  /  노랑: 입력 작업을 한 날인데 하루 최소 횟수에 모자람  (입력하지 않은 날은 점검하지 않습니다)",
                  font=(self.ui_font_family, 8), foreground="#666").pack(anchor="w")
        sched = {"p": None, "rows": [], "cols": []}

        def build_for(name, d0s, d1s):
            name = (name or "").strip()
            if not name:
                raise ValueError("광고를 선택해 주세요.")
            for v in (d0s, d1s):
                if v and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    raise ValueError("기간은 2026-08-01 형식으로 입력해 주세요.")
            if name in key_of_label:
                name, sec = key_of_label[name]
            else:
                m = re.match(r"^(\d+)초\s+(.+)$", name)
                name, sec = (m.group(2).strip(), int(m.group(1))) if m else (name, None)
            if sec is None and name in multi:
                p = combine_presets(name, [x for x in self.store.list_presets() if x["name"] == name])
            else:
                p = self.store.get_preset_by_name(name, sec)
            span = self.store.record_date_span(name, sec)
            if p is None:
                recs0 = self.store.schedule_records(name, "0000-01-01", "9999-12-31", sec)
                p = {"name": name, "seconds": sec if sec is not None else (recs0[0]["seconds"] if recs0 else 0), "start_date": "", "end_date": "", "min_am": 0, "min_fm": 0}
            d0 = d0s or p.get("start_date") or span[0]
            d1 = d1s or (month_end_iso(p["end_date"]) if p.get("end_date") else span[1])
            if not d0 or not d1:
                raise ValueError("이 광고의 기록이 아직 없어 기간을 알 수 없습니다.\n기간을 직접 입력해 주세요.")
            if d1 < d0:
                raise ValueError("기간이 올바르지 않습니다. 시작이 끝보다 늦습니다.")
            if (datetime.date.fromisoformat(d1) - datetime.date.fromisoformat(d0)).days + 1 > 800:
                raise ValueError("기간이 너무 깁니다. 800일 이내로 입력해 주세요.")
            recs = self.store.schedule_records(name, d0, d1, sec)
            rows = build_schedule(p, recs, d0, d1, self.store.processed_dates(d0, d1))
            return p, rows, d0, d1

        def show_sched():
            try:
                p, rows, d0, d1 = build_for(adv.get(), sf.get().strip(), ef.get().strip())
            except ValueError as exc:
                messagebox.showwarning("일정표", str(exc), parent=win)
                return
            sched["p"], sched["rows"] = p, rows
            title_lbl.config(text=schedule_title(p))
            cols = [("날짜", 90), ("AM", 330), ("FM", 330), ("점검", 140)]
            sched["cols"] = cols
            tbl = [(f"{r['date'][2:4]}.{r['date'][5:7]}.{r['date'][8:10]}", r["AM"], r["FM"], r["check"]) for r in rows]
            stree2.configure(columns=[c for c, _ in cols])
            for c_, w in cols:
                stree2.heading(c_, text=c_)
                stree2.column(c_, width=w, anchor="center")
            for i in stree2.get_children():
                stree2.delete(i)
            for r, v in zip(rows, tbl):
                stree2.insert("", "end", values=v, tags=("out",) if r["outside"] else ("short",) if r["check"] else ())
            am_t, fm_t = sum(r["AM_n"] for r in rows), sum(r["FM_n"] for r in rows)
            sum_lbl.config(text=f"{d0} ~ {d1}   AM 총 {am_t}회   FM 총 {fm_t}회   횟수 부족한 날 {sum(1 for r in rows if r['check'])}일")

        def sched_csv():
            if not sched["rows"]:
                messagebox.showinfo("CSV 저장", "먼저 [보기]를 눌러 일정표를 만들어 주세요.", parent=win)
                return
            rows = [(f"{r['date'][2:4]}.{r['date'][5:7]}.{r['date'][8:10]}", r["AM"], r["FM"], r["check"]) for r in sched["rows"]]
            save_csv(sched["cols"], rows)

        def save_xlsx(sheets):
            if not sheets:
                messagebox.showinfo("엑셀 저장", "저장할 내용이 없습니다.", parent=win)
                return
            path = filedialog.asksaveasfilename(title="엑셀 저장", defaultextension=".xlsx", filetypes=[("엑셀 파일", "*.xlsx")], parent=win)
            if not path:
                return
            try:
                export_schedule_xlsx(path, sheets)
            except ImportError:
                messagebox.showerror("엑셀 저장", "엑셀 파일을 만들려면 openpyxl 이 필요합니다.\n명령 프롬프트에서  pip install openpyxl  을 실행한 뒤 다시 시도해 주세요.\n(CSV 저장은 그대로 쓸 수 있습니다.)", parent=win)
                return
            except Exception as exc:
                messagebox.showerror("엑셀 저장 실패", str(exc), parent=win)
                return
            messagebox.showinfo("엑셀 저장", f"저장했습니다 (시트 {len(sheets)}개):\n{path}", parent=win)

        def sched_xlsx():
            if not sched["rows"]:
                messagebox.showinfo("엑셀 저장", "먼저 [보기]를 눌러 일정표를 만들어 주세요.", parent=win)
                return
            p = sched["p"]
            save_xlsx([{"title": schedule_title(p), "sheet": sheet_name_for(p, set()), "rows": sched["rows"]}])

        def all_xlsx():
            used, sheets, skipped = set(), [], 0
            per = period_or_none()
            if per is None:
                messagebox.showwarning("엑셀 저장", "기간은 2026-08-01 형식으로 입력해 주세요.", parent=win)
                return
            if per == ("", ""):
                labels_all = [f"{sec_}초 {nm_}" for nm_, sec_ in self.store.record_keys()] + [f"{nm_} (전체 초수)" for nm_ in multi]
            else:
                labels_all = labels_for(*per)             # 기간을 넣었으면 그 기간에 조회되는 광고만 시트로 만든다
            for lb in labels_all:
                try:
                    p, rows, _, _ = build_for(lb, sf.get().strip(), ef.get().strip())
                except ValueError:
                    skipped += 1
                    continue
                sheets.append({"title": schedule_title(p), "sheet": sheet_name_for(p, used), "rows": rows})
            if skipped and sheets:
                messagebox.showinfo("엑셀 저장", f"기간을 정할 수 없는 광고 {skipped}개는 제외하고 저장합니다.", parent=win)
            save_xlsx(sheets)

        ttk.Button(q2, text="보기", command=show_sched).pack(side="left", padx=2)
        ttk.Button(q2, text="CSV 저장", command=sched_csv).pack(side="left", padx=2)
        ttk.Button(q2, text="엑셀 저장 (이 광고)", command=sched_xlsx).pack(side="left", padx=2)
        ttk.Button(q2, text="엑셀 저장 (목록의 모든 광고, 시트별)", command=all_xlsx).pack(side="left", padx=2)
        title_lbl.pack(anchor="w", pady=(6, 0), before=stree2.master)
        if names:
            show_sched()

    # ------------------------------------------------------------------ 다른 날 입력 가져오기
    IMPORT_STATUS = {"ok": "그대로", "moved": "자리 바뀜", "check": "확인 필요", "skip": "이미 있음", "fail": "못 넣음"}

    @staticmethod
    def _src_items_from_state(entries, orphans):
        items = []
        for (p, tm, prog, li, k), e in entries.items():
            items.append({"time": tm, "program": prog, "line": li, "k": k, "name": e["name"], "seconds": e["seconds"],
                          "price": e.get("price", 0), "grp": e.get("grp", "")})
        for o in orphans:
            items.append({"time": o["time"], "program": o["program"], "line": o.get("line") or 0, "k": int(o["cell"]) - 1,
                          "name": o["name"], "seconds": o["seconds"], "price": o.get("price", 0), "grp": o.get("grp", "")})
        return items

    def _src_items_from_db(self, air_date, media):
        recs = self.store.load_records(air_date, media) if self.store else []
        return [{"time": r["time"], "program": r["program"], "line": r["line"] or 0, "k": int(r["cell"]) - 1, "name": r["name"],
                 "seconds": r["seconds"], "price": r["price"], "grp": r["grp"]} for r in recs]

    def _find_slot(self, row, line, k, occupied):
        """가져온 광고를 이 운행표의 행에 넣을 자리. (칸 key, 바뀐 내용 설명) / 자리가 없으면 (None, '')."""
        p, t, pr = row["page"], row["time"], row["program"]
        nl = len(row["lines"])
        li0 = line if line < nl else 0
        order = [li0] + [x for x in range(nl) if x != li0]

        def key(li, c):
            return (p, t, pr, li, c)

        def free(li, c):
            cd = self.cell_by_key.get(key(li, c))
            return cd is not None and not cd["text"] and key(li, c) not in occupied

        same_line = line < nl
        if same_line and k < 4 and free(li0, k):
            return key(li0, k), ""
        if same_line and 4 <= k < 8:                       # 반으로 나눈 오른쪽: 왼쪽 반에 이미 넣었으면 같은 자리
            host = self.cell_by_key.get(key(li0, k - 4))
            if host and not host["text"] and key(li0, k - 4) in occupied and key(li0, k) not in occupied:
                return key(li0, k), ""
        for li in order:                                   # 같은 줄(또는 다른 줄)의 빈 칸
            for c in range(4):
                if free(li, c):
                    if k >= 8:
                        return key(li, c), "빈 칸으로 ('+N' 대신)"
                    return key(li, c), ("다른 칸으로" if li == li0 and same_line else "다른 줄로")
        for li in order:                                   # 빈 칸이 없으면 내가 넣은 칸을 반으로 나눈다
            for c in range(4):
                cd = self.cell_by_key.get(key(li, c))
                if cd and not cd["text"] and key(li, c) in occupied and key(li, c + 4) not in occupied:
                    return key(li, c + 4), ("" if k == c + 4 and li == li0 else "칸 나눔")
        for li in order:                                   # 그래도 없으면 '+N'
            for c in range(8, 8 + OVERFLOW_MAX):
                if key(li, c) not in occupied:
                    return key(li, c), ("" if k >= 8 else "'+N' 표시로")
        return None, ""

    def _plan_import(self, items, replace):
        """가져올 광고를 이 운행표와 하나씩 비교해 넣을 자리와 결과를 정한다."""
        rowmap, by_time, by_prog = {}, {}, {}
        for p in sorted(self.rows):
            for r in self.rows[p]:
                rowmap.setdefault((r["time"], r["program"]), r)
                by_time.setdefault(r["time"], []).append(r)
                by_prog.setdefault(r["program"], []).append(r)
        have = {}                                          # 이미 넣어 둔 광고 (같은 행·같은 광고) 개수
        if not replace:
            for k_, e in self.entries.items():
                hk = (k_[1], k_[2], e["name"], e["seconds"])
                have[hk] = have.get(hk, 0) + 1
        pm = getattr(self, "_preset_map", {})
        plan = []
        for it in sorted(items, key=lambda i: (i["time"], i["program"], i["line"], i["k"])):
            preset = pm.get((it["name"], it["seconds"]))
            ent = {"name": it["name"], "seconds": it["seconds"],
                   "price": preset["price"] if preset else it["price"], "grp": preset["grp"] if preset else it["grp"]}
            pi = {"src": it, "ent": ent, "row": None, "key": None, "status": "ok", "check": True, "notes": []}
            plan.append(pi)
            if preset is None:
                pi["notes"].append("미리 만든 데이터에 없음")
            else:
                end, st = preset.get("end_date", ""), preset.get("start_date", "")
                if end and self.air_date > end:
                    pi["notes"].append(f"종료됨({short_date(end)})")
                    pi["status"], pi["check"] = "check", False
                if st and self.air_date < st:
                    pi["notes"].append(f"시작 전({short_date(st)})")
                    pi["status"], pi["check"] = "check", False
            row = rowmap.get((it["time"], it["program"]))
            if row is None:
                same_t = by_time.get(it["time"], [])
                same_p = by_prog.get(it["program"], [])
                if len(same_t) == 1:
                    row = same_t[0]
                    pi["notes"].append(f"프로그램명 바뀜 → {row['program']}")
                elif len(same_p) == 1:
                    row = same_p[0]
                    pi["notes"].append(f"시각 바뀜 → {row['time']}")
                if row is None:
                    pi["status"], pi["check"] = "fail", False
                    pi["notes"].append("이 운행표에 같은 행이 없음")
                    continue
                pi["status"], pi["check"] = "check", False
            pi["row"] = row
            hk = (row["time"], row["program"], it["name"], it["seconds"])
            if have.get(hk, 0) > 0:
                have[hk] -= 1
                pi["status"], pi["check"] = "skip", False
                pi["notes"].append("같은 행에 이미 들어 있음")
        occupied = set() if replace else set(self.entries)
        # 체크된 것부터, 그리고 같은 자리에 그대로 들어가는 것부터 자리를 준다 (옮겨야 하는 것이 남의 제자리를 뺏지 않게)
        for first in (True, False):
            for exact in (True, False):
                for pi in plan:
                    if pi["row"] is None or pi["key"] or pi["status"] in ("skip", "fail") or pi["check"] != first:
                        continue
                    key, how = self._find_slot(pi["row"], pi["src"]["line"], pi["src"]["k"], occupied)
                    if exact and (key is None or how):
                        continue
                    if key is None:
                        pi["status"], pi["check"] = "fail", False
                        pi["notes"].append("넣을 자리가 없음")
                        continue
                    occupied.add(key)
                    pi["key"] = key
                    if how:
                        pi["notes"].append(how)
                        if pi["status"] == "ok":
                            pi["status"] = "moved"
        return plan

    def _plan_caps(self, plan, replace):
        """체크된 항목을 넣었을 때 재원을 넘는 행. {(page, time, program): (합계, 재원)}"""
        add = {}
        if not replace:
            for k_, e in self.entries.items():
                add[k_[:3]] = add.get(k_[:3], 0) + e["seconds"]
        for pi in plan:
            if pi["check"] and pi["key"]:
                add[pi["key"][:3]] = add.get(pi["key"][:3], 0) + pi["ent"]["seconds"]
        over = {}
        for pi in plan:
            if not pi["key"]:
                continue
            rk = pi["key"][:3]
            amt = self.row_amount(pi["row"])
            if amt and rk in add and amt[0] + add[rk] > amt[1]:
                over[rk] = (amt[0] + add[rk], amt[1])
        return over

    def import_dialog(self):
        """다른 날(열린 탭 또는 저장된 기록)의 입력을 지금 탭에 넣는다. 행을 비교해 다른 경우는 표시하고 고르게 한다."""
        if not self.doc:
            return
        self._commit_date()
        me = self.active_tab
        others = [t for t in self._sorted_tabs() if t is not me]
        names = [self._tab_title(self._tab_info(t)["media"], self._tab_info(t)["air_date"]) for t in others]
        tab_by_name = dict(zip(names, others))
        my_name = self._tab_title(self.media, self.air_date)

        win = tk.Toplevel(self.root)
        win.title(f"📥 다른 날 입력 가져오기 → {my_name}")
        win.geometry("1060x700")
        win.transient(self.root)
        win.grab_set()
        fb = (self.ui_font_family, 9, "bold")
        ttk.Label(win, text=f"받는 운행표: {my_name}  (지금 탭)", font=(self.ui_font_family, 11, "bold")).pack(anchor="w", padx=12, pady=(10, 2))

        src = ttk.LabelFrame(win, text=" 가져올 곳 ", padding=8)
        src.pack(fill="x", padx=12, pady=4)
        kind = tk.StringVar()
        tab_v = tk.StringVar()
        db_date = tk.StringVar()
        db_media = tk.StringVar(value=self.media)
        # 기본값: 같은 매체의 바로 앞 날짜 탭 → 없으면 같은 매체의 다른 탭 → 없으면 저장된 기록 중 가장 최근 날
        prev = [n for n, t in tab_by_name.items() if self._tab_info(t)["media"] == self.media and self._tab_info(t)["air_date"] < self.air_date]
        same_m = [n for n, t in tab_by_name.items() if self._tab_info(t)["media"] == self.media]
        pick = (prev[-1] if prev else same_m[0] if same_m else names[0] if names else "")
        tab_v.set(pick)
        try:
            last = self.store.last_record_date(self.air_date, self.media) if self.store else ""
        except Exception:
            last = ""
        auto = bool(pick or last)          # 가져올 곳이 분명하면 창을 열자마자 비교해 보여 준다
        if not last:
            try:
                last = (datetime.date.fromisoformat(self.air_date) - datetime.timedelta(days=1)).isoformat()
            except ValueError:
                last = ""
        db_date.set(last)
        kind.set("tab" if pick else "db")

        r1 = ttk.Frame(src)
        r1.pack(fill="x", pady=2)
        rb1 = ttk.Radiobutton(r1, text="열린 탭에서:", variable=kind, value="tab")
        rb1.pack(side="left")
        cb = ttk.Combobox(r1, textvariable=tab_v, values=names, state="readonly", width=14)
        cb.pack(side="left", padx=6)
        ttk.Label(r1, text="(저장하지 않은 입력도 그대로 가져옵니다)", foreground="#666", font=(self.ui_font_family, 8)).pack(side="left")
        if not names:
            rb1.state(["disabled"])
            cb.state(["disabled"])
        r2 = ttk.Frame(src)
        r2.pack(fill="x", pady=2)
        ttk.Radiobutton(r2, text="저장된 기록에서:", variable=kind, value="db").pack(side="left")
        ttk.Entry(r2, textvariable=db_date, width=12).pack(side="left", padx=(6, 2))
        ttk.Button(r2, text="📅", width=3, command=lambda: pick_date(win, db_date.get(), lambda d: (db_date.set(d), kind.set("db")))).pack(side="left")
        ttk.Combobox(r2, textvariable=db_media, values=list(MEDIA_LIST), state="readonly", width=5).pack(side="left", padx=6)
        ttk.Label(r2, text="(PDF를 열지 않아도 DB에 저장된 날의 입력을 가져옵니다)", foreground="#666", font=(self.ui_font_family, 8)).pack(side="left")

        r3 = ttk.Frame(src)
        r3.pack(fill="x", pady=(6, 0))
        mode = tk.StringVar(value="merge")
        ttk.Label(r3, text="넣는 방식:", font=fb).pack(side="left")
        ttk.Radiobutton(r3, text="합치기 (지금 넣어 둔 것은 그대로, 빈 자리에만)", variable=mode, value="merge").pack(side="left", padx=4)
        ttk.Radiobutton(r3, text="바꾸기 (지금 넣어 둔 것을 지우고 가져온 것으로)", variable=mode, value="replace").pack(side="left", padx=4)
        ttk.Button(r3, text="🔍 비교하기", command=lambda: compare()).pack(side="right")

        ttk.Label(win, text="그대로: 같은 자리  ·  자리 바뀜: 원본 광고가 생겼거나 줄이 달라 다른 빈 칸으로  ·  확인 필요(기본 체크 해제): 종료/시작 전 광고, 시각·프로그램명이 바뀐 행\n"
                            "이미 있음: 같은 행에 같은 광고가 이미 있음  ·  못 넣음: 같은 행이 없거나 자리가 없음  ·  빨간 글자: 넣으면 재원 초과   — 첫 칸(☑)을 누르면 체크를 바꿉니다.",
                  foreground="#555", font=(self.ui_font_family, 8), justify="left").pack(anchor="w", padx=12, pady=(4, 0))
        flt = tk.BooleanVar(value=False)
        top2 = ttk.Frame(win)
        top2.pack(fill="x", padx=12, pady=(4, 0))
        summary = ttk.Label(top2, text="[비교하기]를 누르세요.", font=fb)
        summary.pack(side="left")
        ttk.Checkbutton(top2, text="그대로인 것은 숨기기", variable=flt, command=lambda: fill()).pack(side="right")

        tf = ttk.Frame(win)
        tf.pack(fill="both", expand=True, padx=12, pady=6)
        cols = ("chk", "time", "program", "ad", "src", "dst", "status", "note")
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        for c_, h, w in zip(cols, ("넣기", "시각", "프로그램명", "광고", "가져온 자리", "넣을 자리", "결과", "내용"), (44, 70, 180, 190, 80, 80, 74, 300)):
            tree.heading(c_, text=h)
            tree.column(c_, width=w, anchor="w" if c_ in ("program", "ad", "note") else "center", stretch=c_ in ("program", "ad", "note"))
        vs = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        for tg, col in (("ok", "#111827"), ("moved", "#1d4ed8"), ("check", "#b45309"), ("skip", "#9ca3af"), ("fail", "#9ca3af"), ("over", "#dc2626")):
            tree.tag_configure(tg, foreground=col)

        st = {"plan": [], "replace": False, "src_name": ""}

        def where(line, k):
            return (f"{line + 1}줄 " if line else "") + cell_label(k + 1)

        def compare():
            if kind.get() == "tab":
                t = tab_by_name.get(tab_v.get())
                if t is None:
                    messagebox.showinfo("가져오기", "가져올 탭을 골라 주세요.", parent=win)
                    return
                info = self._tab_info(t)
                items = self._src_items_from_state(info["entries"], t["orphans"])
                st["src_name"], src_media = tab_v.get(), info["media"]
            else:
                try:
                    d = parse_date_input(db_date.get())
                except ValueError:
                    d = ""
                if not d:
                    messagebox.showwarning("가져오기", "날짜를 2026-10-30 형식으로 입력해 주세요.", parent=win)
                    return
                db_date.set(d)
                if d == self.air_date and db_media.get() == self.media:
                    messagebox.showwarning("가져오기", "받는 운행표와 같은 날짜·매체입니다. 다른 날을 골라 주세요.", parent=win)
                    return
                try:
                    items = self._src_items_from_db(d, db_media.get())
                except Exception as exc:
                    messagebox.showerror("가져오기", str(exc), parent=win)
                    return
                st["src_name"], src_media = f"{self._tab_title(db_media.get(), d)} (저장된 기록)", db_media.get()
            if not items:
                messagebox.showinfo("가져오기", f"{st['src_name']} 에는 가져올 입력이 없습니다.", parent=win)
                return
            st["replace"] = mode.get() == "replace"
            st["plan"] = self._plan_import(items, st["replace"])
            st["warn_media"] = src_media != self.media
            fill()

        def fill():
            for i in tree.get_children():
                tree.delete(i)
            plan = st["plan"]
            over = self._plan_caps(plan, st["replace"])
            cnt = {k: 0 for k in self.IMPORT_STATUS}
            for idx, pi in enumerate(plan):
                cnt[pi["status"]] += 1
                if flt.get() and pi["status"] == "ok" and not (pi["key"] and pi["key"][:3] in over):
                    continue
                s_ = pi["src"]
                notes = list(pi["notes"])
                tags = [pi["status"]]
                if pi["key"] and pi["key"][:3] in over:
                    tot, cap = over[pi["key"][:3]]
                    notes.append(f"재원 초과 {tot}/{cap}")
                    if pi["check"]:
                        tags = ["over"]
                dst = where(pi["key"][3], pi["key"][4]) if pi["key"] else "-"
                ad = f"{s_['seconds']} {s_['name']}"
                can = pi["key"] is not None
                tree.insert("", "end", iid=str(idx), tags=tags,
                            values=("☑" if pi["check"] else ("☐" if can else "—"), s_["time"], s_["program"], ad,
                                    where(s_["line"], s_["k"]), dst, self.IMPORT_STATUS[pi["status"]], ", ".join(notes)))
            n_chk = sum(1 for pi in plan if pi["check"])
            txt = (f"{st['src_name']}: {len(plan)}건  →  넣을 것 {n_chk}건   "
                   f"(그대로 {cnt['ok']} · 자리 바뀜 {cnt['moved']} · 확인 필요 {cnt['check']} · 이미 있음 {cnt['skip']} · 못 넣음 {cnt['fail']})")
            if over:
                txt += f"   ⚠ 재원 초과 행 {len(over)}개"
            if st.get("warn_media"):
                txt += "   ⚠ 매체가 다릅니다"
            summary.config(text=txt, foreground="#dc2626" if over else "#111827")

        def toggle(ids):
            changed = False
            for iid in ids:
                pi = st["plan"][int(iid)]
                if pi["key"] is None:
                    continue
                pi["check"] = not pi["check"]
                changed = True
            if changed:
                keep = tree.yview()[0]
                fill()
                tree.yview_moveto(keep)

        def on_click(e):
            if tree.identify_region(e.x, e.y) == "cell" and tree.identify_column(e.x) == "#1":
                iid = tree.identify_row(e.y)
                if iid:
                    toggle([iid])
                    return "break"

        tree.bind("<Button-1>", on_click)
        tree.bind("<space>", lambda e: toggle(tree.selection()))

        def set_all(v):
            for pi in st["plan"]:
                if pi["key"] is not None and pi["status"] not in ("skip",):
                    pi["check"] = v
            fill()

        def apply():
            sel = [pi for pi in st["plan"] if pi["check"] and pi["key"]]
            if not sel:
                messagebox.showinfo("가져오기", "넣을 항목이 없습니다. (☑ 체크된 항목만 넣습니다)", parent=win)
                return
            over = self._plan_caps(st["plan"], st["replace"])
            msg = f"{st['src_name']} 에서 {len(sel)}건을 {my_name} 에 넣습니다."
            if st["replace"] and self.entries:
                msg += f"\n지금 넣어 둔 {len(self.entries)}건은 지워집니다."
            if over:
                msg += f"\n\n⚠ 재원을 넘는 행이 {len(over)}개 생깁니다."
            if not messagebox.askyesno("가져오기", msg + "\n\n계속할까요? (Ctrl+Z로 한 번에 되돌릴 수 있습니다)", parent=win):
                return
            self._push()
            if st["replace"]:
                self.entries = {}
            for pi in sel:
                self.entries[pi["key"]] = dict(pi["ent"])
            win.destroy()
            self.update_cell_info()
            self.draw_page()
            messagebox.showinfo("가져오기 완료", f"{len(sel)}건을 넣었습니다.\n확인한 뒤 저장해 주세요.")

        bf = ttk.Frame(win, padding=8)
        bf.pack()
        ttk.Button(bf, text="모두 체크", command=lambda: set_all(True)).pack(side="left", padx=4)
        ttk.Button(bf, text="모두 해제", command=lambda: set_all(False)).pack(side="left", padx=4)
        ttk.Button(bf, text="✅ 체크한 항목 넣기", command=apply).pack(side="left", padx=12)
        ttk.Button(bf, text="닫기", command=win.destroy).pack(side="left", padx=4)
        if auto:
            compare()

    # ------------------------------------------------------------------ 저장
    def _collect_records(self):
        recs = []
        for key, ent in self.entries.items():
            p, tm, prog, line, cell = key
            recs.append({"page": p, "time": tm, "program": prog, "line": line, "cell": cell + 1,
                         "name": ent["name"], "seconds": ent["seconds"], "price": ent["price"], "grp": ent["grp"], "media": self.media,
                         "saved_at": ent.get("saved_at") or ""})
        recs.sort(key=lambda r: (r["page"], r["time"], r["line"], r["cell"]))
        for o in self.orphans:
            recs.append({"page": o.get("page"), "time": o["time"], "program": o["program"], "line": o.get("line", 0), "cell": o["cell"],
                         "name": o["name"], "seconds": o["seconds"], "price": o["price"], "grp": o["grp"], "saved_at": o.get("saved_at") or ""})
        return recs

    def _preview_key(self):
        return f"{self.media}{(self.air_date or '').replace('-', '')}" if self.media and self.air_date else ""

    def _publish_preview(self, state):
        """다른 PC 조회용: 지금 입력이 들어간 운행표 PDF를 미리보기 폴더에 두고 상태를 적는다."""
        key = self._preview_key()
        if not key or not self.doc or not _PREVIEW_KEY_RE.match(key):
            return
        try:
            data = self._build_output_bytes()
            os.makedirs(PREVIEW_DIR, exist_ok=True)
            tmp = os.path.join(PREVIEW_DIR, key + ".pdf.tmp")
            with open(tmp, "wb") as f:
                f.write(data)
            os.replace(tmp, os.path.join(PREVIEW_DIR, key + ".pdf"))
        except Exception:
            pass                                    # 미리보기는 실패해도 저장은 계속
        update_preview_status(key, media=self.media, air_date=self.air_date, state=state, dirty=False,
                              saved_at=time.strftime("%Y-%m-%d %H:%M:%S"), pc=computer_name())

    def _set_preview_dirty(self, v):
        """편집 중 표시: 저장 안 된 수정이 생기면 '현재 편집중' (상태 파일에만, 가볍게)."""
        key = self._preview_key()
        if not key or not _PREVIEW_KEY_RE.match(key):
            return
        cur = read_preview_status().get(key) or {}
        if v:
            before = cur.get("state_before") if cur.get("dirty") else cur.get("state", "")
            update_preview_status(key, media=self.media, air_date=self.air_date, dirty=True, state="편집중",
                                  state_before=before or "", pc=computer_name())
        elif cur.get("dirty"):                      # 저장하지 않고 버림 -> 마지막 저장 때 상태로 되돌린다
            update_preview_status(key, dirty=False, state=cur.get("state_before") or cur.get("state", ""))

    def save_records(self, state="편집중"):
        if not self.air_date or self.store is None:
            return False
        try:
            now = self.store.replace_records(self.air_date, self._collect_records(), os.path.basename(self.pdf_path or ""), media=self.media)
        except Exception as exc:
            messagebox.showerror("기록 저장 실패", str(exc))
            return False
        for ent in self.entries.values():          # 방금 저장한 것에 저장 시각을 붙여 둔다
            if not ent.get("saved_at"):
                ent["saved_at"] = now
        self._load_removed()
        if self.doc:
            self.draw_page()
        self._publish_preview(state)
        return True

    def save_records_only(self):
        if not self.doc:
            return
        self._commit_date()
        if self.save_records():
            self.dirty = False
            messagebox.showinfo("기록 저장", f"{self.air_date} 입력 기록 {len(self._collect_records())}건을 저장했습니다.")

    def _build_output_bytes(self, highlight=None):
        """저장·출력할 PDF. highlight 가 None 이면 [출력물 강조] 설정을 따른다."""
        if highlight is None:
            highlight = bool(self.highlight_var.get())
        doc = fitz.open(stream=self.src_bytes, filetype="pdf")
        try:
            font_ok = bool(self.font_path and os.path.exists(self.font_path))
            measure = None
            if font_ok:
                try:
                    measure = fitz.Font(fontfile=self.font_path)
                except Exception:
                    measure = None

            def wfn(text, fs):
                try:
                    return measure.text_length(text, fontsize=fs) if measure else est_text_width(text, fs)
                except Exception:
                    return est_text_width(text, fs)

            for p in sorted({k[0] for k in self.entries}):
                items, marks = self.entry_layout(p, width_fn=wfn)
                if not items and not marks:
                    continue
                page = doc.load_page(p)
                font_name = "helv"
                if font_ok:
                    try:
                        page.insert_font(fontname="kfont", fontfile=self.font_path)
                        font_name = "kfont"
                    except Exception:
                        font_name = "helv"
                labels = self.over_labels(marks, width_fn=wfn)
                mo = self.mark_opts
                hlw = mo["hl_width"]
                if font_name == "helv" and any(not it["text"].isascii() for it in list(items) + labels):
                    raise RuntimeError("한글 폰트를 불러오지 못했습니다.\n⚙️ 설정에서 .ttf/.ttc 한글 폰트 파일을 지정해 주세요.")
                k, tx, ty = page_transform(page)   # A3 등으로 인쇄된 PDF 면 위치·글자 크기에 배율과 위치를 반영한다
                for it in items:
                    x0, y0, x1, y1 = it["rect"]
                    fs = it["fs"]
                    if highlight:              # 강조: 추가한 칸 안쪽에 테두리만 (채우기 없음)
                        page.draw_rect(fitz.Rect((x0 + 0.9) * k + tx, (y0 + 0.9) * k + ty, (x1 - 0.9) * k + tx, (y1 - 0.9) * k + ty),
                                       color=HL_COLOR, fill=None, width=hlw * k)
                    pt = fitz.Point((x0 + it["inset"]) * k + tx, ((y0 + y1) / 2.0 + fs * 0.35) * k + ty)
                    kw = {}
                    if it["xs"] < 0.999:       # 글자 폭(간격)을 줄여서 칸에 맞춘다
                        kw["morph"] = (pt, fitz.Matrix(it["xs"], 1))
                    page.insert_text(pt, it["text"], fontsize=fs * k, fontname=font_name, color=(0, 0, 0), **kw)
                for m in marks:
                    if highlight:
                        page.draw_rect(fitz.Rect((MARK_X - 1) * k + tx, (m["cy"] - 4) * k + ty, (MARK_X + 11) * k + tx, (m["cy"] + 4) * k + ty),
                                       color=HL_COLOR, fill=None, width=hlw * k)
                    page.insert_text(fitz.Point(MARK_X * k + tx, (m["cy"] + 2.3) * k + ty), f"+{m['n']}", fontsize=6.5 * k, fontname="helv", color=(0, 0, 0))
                for lb in labels:              # 칸이 모자란 줄: 줄 경계선에 걸쳐 추가 광고 이름을 작게 (바탕을 칠해 선 위에서도 읽히게)
                    x0, y0, x1, y1 = lb["rect"]
                    if mo["fill"] or highlight:
                        page.draw_rect(fitz.Rect(x0 * k + tx, y0 * k + ty, x1 * k + tx, y1 * k + ty),
                                       color=HL_COLOR if highlight else None, fill=(1, 1, 1) if mo["fill"] else None,
                                       width=(hlw * k) if highlight else 0)
                    pt = fitz.Point(lb["x0"] * k + tx, (lb["yc"] + lb["fs"] * 0.35) * k + ty)
                    kw = {"morph": (pt, fitz.Matrix(lb["xs"], 1))} if lb["xs"] < 0.999 else {}
                    page.insert_text(pt, lb["text"], fontsize=lb["fs"] * k, fontname=font_name, color=(0, 0, 0), **kw)
            return doc.tobytes(deflate=True)
        finally:
            doc.close()

    def save_with_preview(self):
        if not self.doc:
            return
        self._commit_date()
        if not self.preview_var.get():
            self.export_pdf()
            return
        try:
            data = self._build_output_bytes()
        except Exception as exc:
            messagebox.showerror("미리보기 실패", str(exc))
            return
        data = self.show_preview(data)
        if data is not None:
            self.export_pdf(data)

    def show_preview(self, data):
        """저장될 PDF를 그대로 보여 준다. [이 내용으로 저장]을 누르면 그 PDF 데이터, 아니면 None.
        [출력물 강조]를 바꾸면 PDF를 다시 만들어 보여 주므로, 보이는 그대로 저장·출력된다."""
        try:
            pdoc = fitz.open(stream=data, filetype="pdf")
        except Exception as exc:
            messagebox.showerror("미리보기 실패", str(exc))
            return None
        win = tk.Toplevel(self.root)
        win.title("👁 저장 전 미리보기")
        try:
            win.geometry(self.preview_geom or "980x960")       # 마지막 창 크기·위치
        except Exception:
            win.geometry("980x960")
        win.bind("<Configure>", lambda e: e.widget is win and setattr(self, "preview_geom", win.geometry()), add="+")
        win.transient(self.root)
        win.grab_set()
        st = {"page": min(self.cur, pdoc.page_count - 1), "zoom": self.zoom, "photo": None, "ok": False, "data": data, "doc": pdoc}
        hl = tk.BooleanVar(value=self.highlight_var.get())
        pv = tk.StringVar()

        over_rows = {k[:3] for k in self.entries if k in self.cell_by_key and self.is_over(self.row_check(self.cell_by_key[k]["row"]))}
        short_n = short_sec = 0
        for pg in self.rows.values():
            for r in pg:
                if self.row_is_custom(r):
                    ck = self.row_check(r)
                    if ck and ck[3] < ck[1]:
                        short_n += 1
                        short_sec += ck[1] - ck[3]
        summary = f"추가 {len(self.entries)}건"
        if short_n:
            summary += f"  /  재원 부족 {short_n}개 행 (합계 {short_sec}초)"
        if over_rows:
            summary += f"  /  ⚠ 재원 초과 행 {len(over_rows)}개"
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        tk.Label(top, text=summary, fg="#dc2626" if over_rows else "#a16207" if short_n else "#111", font=(self.ui_font_family, 10, "bold")).pack(side="left")
        ttk.Button(top, text="🏷 표시 설정", command=lambda: self.mark_opts_dialog(parent=win, on_change=lambda: (hl.set(self.highlight_var.get()), rebuild()))).pack(side="right", padx=(6, 0))
        ttk.Checkbutton(top, text="추가한 칸 테두리 강조 (출력·저장에도 적용)", variable=hl, command=lambda: rebuild()).pack(side="right")

        nav = ttk.Frame(win, padding=(6, 0))
        nav.pack(fill="x")
        ttk.Button(nav, text="◀", width=3, command=lambda: goto(st["page"] - 1)).pack(side="left")
        ttk.Label(nav, textvariable=pv, width=10, anchor="center").pack(side="left")
        ttk.Button(nav, text="▶", width=3, command=lambda: goto(st["page"] + 1)).pack(side="left")
        ttk.Button(nav, text="🔍-", width=4, command=lambda: set_zoom(st["zoom"] / 1.25)).pack(side="left", padx=(12, 2))
        zl = ttk.Label(nav, text="", width=6, anchor="center", font=(self.ui_font_family, 9, "bold"))
        zl.pack(side="left")
        ttk.Button(nav, text="🔍+", width=4, command=lambda: set_zoom(st["zoom"] * 1.25)).pack(side="left")
        ttk.Button(nav, text="⤢ 전체 보기", command=lambda: set_zoom(fit_zoom(), "fit")).pack(side="left", padx=(10, 2))
        ttk.Button(nav, text="↔ 너비 맞춤", command=lambda: set_zoom(fit_zoom(True), "width")).pack(side="left", padx=2)
        ttk.Button(nav, text="100%", command=lambda: set_zoom(1.0)).pack(side="left", padx=2)
        ttk.Label(nav, text="Ctrl + 휠: 확대/축소", foreground="#64748b", font=(self.ui_font_family, 8)).pack(side="right")

        mid = ttk.Frame(win)
        mid.pack(fill="both", expand=True, pady=4)
        cv = tk.Canvas(mid, background="#e5e7eb", highlightthickness=0)
        vb = ttk.Scrollbar(mid, orient="vertical", command=cv.yview)
        hb = ttk.Scrollbar(mid, orient="horizontal", command=cv.xview)
        cv.configure(yscrollcommand=vb.set, xscrollcommand=hb.set)
        cv.grid(row=0, column=0, sticky="nsew")
        vb.grid(row=0, column=1, sticky="ns")
        hb.grid(row=1, column=0, sticky="ew")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        def rebuild():
            self.highlight_var.set(hl.get())
            self._save_settings()
            try:
                new = self._build_output_bytes()
                ndoc = fitz.open(stream=new, filetype="pdf")
            except Exception as exc:
                messagebox.showerror("미리보기", str(exc), parent=win)
                return
            try:
                st["doc"].close()
            except Exception:
                pass
            st["data"], st["doc"] = new, ndoc
            draw()

        def draw():
            z = st["zoom"]
            pdoc = st["doc"]
            pg = pdoc.load_page(st["page"])
            pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            st["photo"] = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"))
            cv.delete("all")
            cv.create_image(M, M, image=st["photo"], anchor="nw")
            cv.configure(scrollregion=(0, 0, pg.rect.width * z + 2 * M, pg.rect.height * z + 2 * M))
            pv.set(f"{st['page'] + 1} / {pdoc.page_count}")
            zl.config(text=f"{round(z * 100)}%")

        def goto(i):
            i = max(0, min(st["doc"].page_count - 1, i))
            if i != st["page"]:
                st["page"] = i
                draw()
                cv.yview_moveto(0)

        def zoom(d):
            cur = min(range(len(ZOOM_LEVELS)), key=lambda i: abs(ZOOM_LEVELS[i] - st["zoom"]))
            tgt = max(0, min(len(ZOOM_LEVELS) - 1, cur + d))
            if tgt != cur:
                st["zoom"] = ZOOM_LEVELS[tgt]
                draw()

        def fit_zoom(width_only=False):
            win.update_idletasks()
            pg = st["doc"].load_page(st["page"])
            cw, ch = max(100, cv.winfo_width() - 2 * M), max(100, cv.winfo_height() - 2 * M)
            zw = cw / pg.rect.width
            return zw if width_only else min(zw, ch / pg.rect.height)

        def set_zoom(z, mode=None):
            z = max(0.25, min(6.0, z))               # 25% ~ 600% (100% = PDF 실제 크기)
            self.preview_view = mode or round(z, 3)  # 마지막 보기 기억
            if abs(z - st["zoom"]) < 1e-3:
                return
            fx, fy = cv.xview()[0], cv.yview()[0]
            st["zoom"] = z
            draw()
            cv.xview_moveto(fx)
            cv.yview_moveto(fy)

        def wheel(e):
            up = e.delta > 0 or getattr(e, "num", None) == 4
            if e.state & 0x0004:                     # Ctrl + 휠: 확대/축소
                set_zoom(st["zoom"] * (1.15 if up else 1 / 1.15))
            elif e.state & 0x0001:
                cv.xview_scroll(-1 if up else 1, "units")
            else:
                cv.yview_scroll(-1 if up else 1, "units")
            return "break"

        win.bind("<Control-plus>", lambda e: set_zoom(st["zoom"] * 1.25))
        win.bind("<Control-equal>", lambda e: set_zoom(st["zoom"] * 1.25))
        win.bind("<Control-minus>", lambda e: set_zoom(st["zoom"] / 1.25))
        win.bind("<Control-0>", lambda e: set_zoom(fit_zoom(), "fit"))

        cv.bind("<MouseWheel>", wheel)
        cv.bind("<Button-4>", wheel)
        cv.bind("<Button-5>", wheel)

        def ok():
            st["ok"] = True
            win.destroy()

        bf = ttk.Frame(win, padding=8)
        bf.pack()
        ttk.Button(bf, text="✅ 이 내용으로 저장", command=ok).pack(side="left", padx=6)
        ttk.Button(bf, text="🖨 출력", command=lambda: self.print_pdf_data(st["data"], win)).pack(side="left", padx=6)
        ttk.Button(bf, text="돌아가서 수정", command=win.destroy).pack(side="left", padx=6)
        draw()
        pv0 = self.preview_view                       # 마지막에 보던 방식으로
        if pv0 == "fit":
            win.after(60, lambda: set_zoom(fit_zoom(), "fit"))
        elif pv0 == "width":
            win.after(60, lambda: set_zoom(fit_zoom(True), "width"))
        elif isinstance(pv0, (int, float)):
            win.after(60, lambda: set_zoom(float(pv0)))
        win.wait_window(win)
        self._save_settings()
        try:
            st["doc"].close()
        except Exception:
            pass
        return st["data"] if st["ok"] else None

    def print_pdf_data(self, data, parent=None):
        """미리보기 내용을 그대로 출력한다. 임시 PDF를 만들어 운영체제의 기본 PDF 프로그램으로 인쇄한다. (저장은 하지 않는다)"""
        import tempfile
        import subprocess
        import sys
        try:
            stem = f"{self.media}{(self.air_date or today_iso()).replace('-', '')}" if self.media else "운행표"
            path = os.path.join(tempfile.gettempdir(), f"{stem}_출력.pdf")
            with open(path, "wb") as f:
                f.write(data)
            if sys.platform.startswith("win"):
                try:
                    os.startfile(path, "print")          # 기본 PDF 프로그램이 인쇄 (설정에 따라 인쇄창이 뜸)
                except OSError:
                    os.startfile(path)                    # 인쇄 동작이 연결 안 된 경우: PDF를 열어 줌 -> 거기서 인쇄
                    messagebox.showinfo("출력", "PDF 프로그램에서 열었습니다.\n거기서 인쇄(Ctrl+P)해 주세요.", parent=parent)
            elif sys.platform == "darwin":
                subprocess.run(["lp", path], check=True)
            else:
                subprocess.run(["lp", path], check=True)
        except Exception as exc:
            messagebox.showerror("출력 실패", f"출력하지 못했습니다:\n{exc}\n\n[이 내용으로 저장]으로 PDF를 저장한 뒤 PDF 프로그램에서 인쇄해 주세요.", parent=parent)

    def export_pdf(self, data=None):
        if not self.doc:
            return False
        self._commit_date()
        base = os.path.splitext(os.path.basename(self.pdf_path or "output"))[0]
        stem = f"{self.media}{(self.air_date or today_iso()).replace('-', '')}" if self.media else f"{base}_입력"
        folder = getattr(self, "last_save_dir", "") or (os.path.dirname(self.pdf_path) if self.pdf_path else "")
        if not folder or not os.path.isdir(folder):
            folder = os.path.expanduser("~")
        name, n = f"{stem}.pdf", 0
        while os.path.exists(os.path.join(folder, name)):   # 같은 이름이 있으면 -1, -2 ... 로 늘린다
            n += 1
            name = f"{stem}-{n}.pdf"
        out = filedialog.asksaveasfilename(title="PDF 저장", defaultextension=".pdf", initialdir=folder, initialfile=name, filetypes=[("PDF 파일", "*.pdf")])
        if out:
            self.last_save_dir = os.path.dirname(out)
        if not out:
            return False
        try:
            if data is None:
                data = self._build_output_bytes()
            with open(out, "wb") as f:
                f.write(data)
        except Exception as exc:
            messagebox.showerror("저장 실패", f"PDF 파일 저장 실패:\n{exc}")
            return False
        if not self.save_records(state="수정완료"):
            messagebox.showwarning("기록 저장 실패", "PDF는 저장했지만 입력 기록은 저장하지 못했습니다.\n[기록만 저장]으로 다시 시도해 주세요.")
            return False
        self.dirty = False
        messagebox.showinfo("저장 완료", f"PDF 파일과 입력 기록을 저장했습니다:\n{out}")
        return True

    def _finish_exit(self):
        try:                                       # 저장하지 않고 끈 탭: 다른 PC 조회의 '현재 편집중' 표시를 지운다
            for t in list(getattr(self, "tabs", [])):
                i = self._tab_info(t)
                if i["dirty"] and i["media"] and i["air_date"]:
                    k = f"{i['media']}{i['air_date'].replace('-', '')}"
                    cur = read_preview_status().get(k) or {}
                    if cur.get("dirty"):
                        update_preview_status(k, dirty=False, state=cur.get("state_before") or cur.get("state", ""))
        except Exception:
            pass
        try:
            self.view_server.stop()
        except Exception:
            pass
        if self.backup_on:
            self.backup_error = ""
            if self.run_backup() is None and getattr(self, "backup_error", ""):
                messagebox.showwarning("자동 백업 실패", f"백업 폴더에 저장하지 못했습니다.\n{self.effective_backup_dir()}\n\n{self.backup_error}\n\n⚙️ 설정에서 백업 폴더를 확인해 주세요.")
        self.root.destroy()

    def on_closing_app(self):
        self._commit_date()
        pending = [t for t in self._sorted_tabs() if self._tab_info(t)["dirty"]]
        if pending:
            names = ", ".join(self._tab_title(self._tab_info(t)["media"], self._tab_info(t)["air_date"]) for t in pending)
            ans = messagebox.askyesnocancel("종료 확인", f"저장하지 않은 탭이 {len(pending)}개 있습니다:\n{names}\n\n"
                                            "탭마다 PDF와 기록을 저장하고 종료할까요?\n(아니오: 저장하지 않고 종료)")
            if ans is None:
                return
            if ans:
                for t in pending:
                    self._activate(t)
                    if not self.export_pdf():      # 저장 창에서 취소하면 종료도 멈춘다
                        return
        elif not messagebox.askokcancel("종료 확인", "운행표 데이터 편집기를 종료할까요?", default="cancel", parent=self.root):
            return                                 # X 를 실수로 눌러도 바로 닫히지 않게
        self._finish_exit()


if __name__ == "__main__":
    root = tk.Tk()
    app = DataEditor(root)
    root.mainloop()
