#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
엑셀 TV 편성표 편집기 (시험판) - 편집 프로그램 런처 TV 탭
MBC 편성시스템에서 내려받은 '주간편성표' 엑셀(.xlsx)로 A4 1페이지 편성표 PDF를 만든다.
이 프로그램은 아직 다른 PC·서버로 자료를 보내지 않는다 (시험용).

사용법
  - 그냥 실행하면 창이 뜬다: 엑셀 불러오기 -> 편성표를 보면서 칸을 클릭해 오른쪽에서 고치기
    (0.5초 뒤 자동으로 다시 그림) -> PDF 저장. 원본 PDF와 칸 단위 비교, 작업 시간·고친 줄 수 기록
  - 명령줄:  python tv_schedule_maker.py 주간편성표.xlsx [-o 결과.pdf] [--prev 지난주.xlsx]
             python tv_schedule_maker.py 주간편성표.xlsx --learn 같은주_원본.pdf
             python tv_schedule_maker.py --compare 기준.pdf 비교.pdf

필요한 모듈:  pip install pymupdf openpyxl  (런처로 실행하면 런처의 공용 실행 환경을 씀)
설정·이름 사전·최근 프로젝트:  런처로 실행하면 C:\\EDIT\\TV\\데이터\\엑셀편성표, 따로 실행하면 이 파일 폴더
"""
import os
import re
import sys
import json
import shutil
import datetime

import openpyxl
import pymupdf

try:
    pymupdf.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

APP_VERSION = "1.4.0"
VERSION = f"v{APP_VERSION} (시험판)"


def _detect_profile():
    """같은 프로그램을 두 가지로 쓴다: 엑셀 TV 편성표 편집기 / 엑셀 DMB 편성표 편집기.
    런처에서 'excel_dmb' 로 실행하거나, excel_dmb.py 로 실행하거나, --dmb 를 붙이면 DMB"""
    if os.environ.get("RADIO_APP_ID") == "excel_dmb" or os.environ.get("EXCEL_TV_PROFILE") == "dmb" \
            or "--dmb" in sys.argv:
        return "dmb"
    return "tv"


IS_DMB = _detect_profile() == "dmb"
APP_NAME = "엑셀 DMB 편성표 편집기" if IS_DMB else "엑셀 TV 편성표 편집기"
# 프로그램 파일이 있는 곳 (런처로 실행해도 이 파일 위치 - 런처 exe 위치가 아님)
APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """설정·이름 사전·최근 목록을 두는 곳.
    런처로 실행하면 C:\\EDIT\\TV\\데이터\\엑셀편성표 (프로그램을 업데이트해도 남음), 따로 실행하면 프로그램 폴더"""
    home = os.environ.get("RADIO_HOME")
    d = os.path.join(home, "데이터", "엑셀편성표") if home else APP_DIR
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = APP_DIR
    return d


DATA_DIR = _data_dir()
TV_CONFIG_PATH = os.path.join(DATA_DIR, "tv_schedule_config.json")
# DMB 편집기는 설정을 따로 (처음에는 TV 설정을 복사해서 시작: 글꼴·급 등은 같고 제목·빼기 체크는 따로)
CONFIG_PATH = os.path.join(DATA_DIR, "dmb_schedule_config.json") if IS_DMB else TV_CONFIG_PATH

# ---------------------------------------------------------------- 서버로 보내기 (다른 PC에서 보기)
# PDF 저장 -> C:\\EDIT\\TV\\데이터\\미리보기 에 PDF + status.json -> TV PC 런처가 사내망에 보여 줌
# -> 녹음 PC 서버가 1분 안에 가져감 (엑셀 편성표면 프로젝트·즐겨찾기·설정 백업도 그때 같이)
TV_VIEW_PORT_DEFAULT = 8081


def tv_preview_dir():
    """런처로 실행할 때만: C:\\EDIT\\TV\\데이터\\미리보기 (TV 편성표 편집기와 같은 폴더)"""
    home = os.environ.get("RADIO_HOME")
    return os.path.join(home, "데이터", "미리보기") if home else ""


def tv_view_config():
    """다른 PC에서 보기 설정 - TV 편성표 편집기·DMB 편집기·런처와 같이 쓰는 미리보기\\view.json"""
    try:
        with open(os.path.join(tv_preview_dir(), "view.json"), "r", encoding="utf-8") as f:
            c = json.load(f)
        if isinstance(c, dict):
            return {"on": bool(c.get("on")), "port": int(c.get("port", TV_VIEW_PORT_DEFAULT))}
    except Exception:
        pass
    return {"on": False, "port": TV_VIEW_PORT_DEFAULT}


def save_tv_view_config(on, port=None):
    d = tv_preview_dir()
    if not d:
        return False
    try:
        os.makedirs(d, exist_ok=True)
        port = int(port or tv_view_config()["port"])
        with open(os.path.join(d, "view.json.tmp"), "w", encoding="utf-8") as f:
            json.dump({"on": bool(on), "port": port}, f)
        os.replace(os.path.join(d, "view.json.tmp"), os.path.join(d, "view.json"))
        return True
    except Exception:
        return False


def update_tv_preview_status(key, **fields):
    """미리보기\\status.json 의 이 편성표 줄만 고침 (다른 편집기 줄은 그대로)"""
    d = tv_preview_dir()
    if not d:
        return
    try:
        os.makedirs(d, exist_ok=True)
        path = os.path.join(d, "status.json")
        try:
            with open(path, "r", encoding="utf-8") as f:
                st = json.load(f)
            if not isinstance(st, dict):
                st = {}
        except Exception:
            st = {}
        cur = st.get(key) or {}
        cur.update(fields)
        cur["updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st[key] = cur
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
        os.replace(path + ".tmp", path)
    except Exception:
        pass

DAY_START_MIN = 5 * 60          # 편성표는 05:00 부터
DAY_LEN_MIN = 24 * 60           # 다음날 05:00 까지

# ---------------------------------------------------------------- 기본 설정
DEFAULT_CONFIG = {
    "_설명": "엑셀 TV 편성표 편집기 설정. 시간은 'HH:MM' (24시 이후는 25:00 처럼).",
    "font_files": [
        "C:/Windows/Fonts/MBCNEWL.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
    ],
    "logo_file": "logo.png",
    "title_text": "MBC 주간편성표",
    "title_size": 22,
    "title_bold": False,
    "_날짜_설명": "stamp_mode: 저장 시각 / 엑셀 조회일 / 주간 범위 / 직접 입력 / 표시 안 함",
    "stamp_mode": "저장 시각",
    "stamp_text": "",
    "_프로젝트_정리_설명": "프로젝트는 데이터\\프로젝트\\주 폴더에. 한 주에 최근 몇 개를 남길지(PDF 로 낸 것은 늘 남김), 며칠 지난 주를 백업으로 옮길지",
    "proj_keep_per_week": 10,
    "proj_archive_days": 60,
    "grade_weekday": [["05:00", "07:00", "C"], ["07:00", "08:30", "B"], ["08:30", "09:30", "A"],
                      ["09:30", "12:00", "B"], ["12:00", "18:00", "C"], ["18:00", "19:00", "B"],
                      ["19:00", "19:30", "A"], ["19:30", "23:30", "SA"], ["23:30", "24:00", "A"],
                      ["24:00", "25:00", "B"], ["25:00", "29:00", "C"]],
    "grade_weekend": [["05:00", "08:00", "C"], ["08:00", "12:30", "A"], ["12:30", "17:00", "B"],
                      ["17:00", "18:00", "A"], ["18:00", "23:30", "SA"], ["23:30", "24:30", "A"],
                      ["24:30", "29:00", "C"]],
    "pink_plan2": True,
    "plan_label": False,
    "pink_changed_from_prev": True,
    "highlight": [],
    "_음영_설명": "shade_rules: 매주 자동으로 음영을 넣을 프로그램. match 는 프로그램명에 들어 있는 글자 또는 코드 앞부분",
    "shade_rules": [],
    "_끝시각_설명": "끝 시각 표시: 칸 왼쪽 아래에 끝나는 분을 표시. end_label_rules 에 프로그램명(일부)을 넣으면 매주 자동",
    "end_label_rules": [],
    "end_label_format": "{mm}",
    "_아이콘_설명": "icon_colors: 아이콘 배경색 바꾸기 (예: {\"해\": \"#c2185b\"}). icon_dark_text: 글자를 진회색으로 할 아이콘",
    "icon_colors": {},
    "icon_dark_text": ["U"],
    "_로컬_설명": "local_rules: 매주 엑셀을 불러올 때 자동으로 끼워 넣을 로컬. days 예: 월-금, 토-일, 월,수,금, 매일",
    "local_rules": [],
    "favorites": [
        {"name": "광주MBC 뉴스투데이", "shade": "#d9d9d9", "bold": True},
        {"name": "R/L...MBC 뉴스데스크...R/L"}
    ],
    "name_replace": [
        ["\u00a0", " "],
        ["엄마를부탁해", "엄마를 부탁해"],
        ["리포트결혼지옥", "리포트 결혼지옥"],
        ["스포츠축구", "스포츠 축구"],
        ["국가대표친선경기", "국가대표 친선경기 "],
        ["라이어1-6회", "라이어 1-6회"],
        ["김연경2스페셜", "김연경2 스페셜"]
    ],
    "name_regex": [
        ["^(일일드라마|금토드라마|주말드라마|수목드라마|월화드라마|일일연속극)(?=\\S)", "\\1 "],
        ["MBC(?=뉴스)", "MBC "],
        ["(?<=[^\\s])(스페셜)(?=\\(|$)", " \\1"],
        ["\\s{2,}", " "]
    ]
}


def load_config():
    if IS_DMB and not os.path.exists(CONFIG_PATH) and os.path.exists(TV_CONFIG_PATH):
        try:
            with open(TV_CONFIG_PATH, "r", encoding="utf-8") as f:
                tv = json.load(f)
            tv["title_text"] = dmb_title_auto(tv.get("title_text", ""))
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(tv, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            print(f"[경고] 설정 파일을 읽지 못해 기본값 사용: {e}")
    else:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return cfg


# ---------------------------------------------------------------- 엑셀 읽기
def hhmm_to_min(v):
    """'05:00', '25:35', 시각/날짜시각/경과시간(timedelta)/엑셀 소수(0.2083=05:00) 모두 받음"""
    if isinstance(v, datetime.datetime):
        # 엑셀이 24시 이후를 1900-01-01 01:35 처럼 바꾼 경우 -> 날짜만큼 24시간 더함
        extra = 0
        if v.year in (1899, 1900) and v.date() > datetime.date(1899, 12, 31):
            extra = (v.date() - datetime.date(1899, 12, 31)).days * 1440
        return extra + v.hour * 60 + v.minute
    if isinstance(v, datetime.time):
        return v.hour * 60 + v.minute
    if isinstance(v, datetime.timedelta):
        return int(round(v.total_seconds() / 60))
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 0 <= v < 3:              # 엑셀 시각 소수 (1 = 24시간)
            return int(round(v * 1440))
        if 0 <= v <= 3000:          # 0500, 2535 같은 숫자
            return int(v) // 100 * 60 + int(v) % 100
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})\s*[:시]\s*(\d{2})", s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m = re.fullmatch(r"(\d{1,2})(\d{2})", s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    raise ValueError(f"시각 형식 오류: {v!r}")


def to_yyyymmdd(v):
    """20261116, '2026-11-16', '2026.11.16', 날짜형 모두 '20261116' 으로"""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y%m%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 19000101 <= v <= 29991231:
            return str(int(v))
        if 30000 < v < 80000:       # 엑셀 날짜 일련번호
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime("%Y%m%d")
        return ""
    d = re.sub(r"[^0-9]", "", str(v).split(" ")[0])
    return d if len(d) == 8 else ""


HEADERS = {
    "date": "방송일", "code": "프로그램코드", "name": "프로그램명", "start": "편성시각(From)",
    "end": "편성시각(To)", "cc": "자막", "ad": "화면해설", "uhd": "UHD", "sign": "수화방송",
    "age": "연령구분", "rl": "R/L", "rerun": "재방여부", "plan": "편성안",
}


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx_raw(path):
    """openpyxl 이 못 여는 엑셀(서식 정보가 깨진 파일)용: 서식은 무시하고 값만 직접 읽음.
    반환: [시트별 [행(list)]]"""
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                # 글자 일부만 서식이 다른 경우(<r>)도 이어 붙이고, 윗주(<rPh>) 글자는 뺌
                parts = [t.text or "" for t in si.findall("m:t", ns)]
                parts += [t.text or "" for r in si.findall("m:r", ns) for t in r.findall("m:t", ns)]
                shared.append("".join(parts))
        sheet_files = []
        try:
            wbx = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            target = {r.get("Id"): r.get("Target") for r in rels}
            for sh in wbx.find("m:sheets", ns):
                t = target.get(sh.get("{%s}id" % ns["r"]), "")
                t = t.lstrip("/")
                sheet_files.append(t if t.startswith("xl/") else "xl/" + t)
        except Exception:
            pass
        if not sheet_files:
            sheet_files = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        sheets = []
        for sf in sheet_files:
            if sf not in names:
                continue
            rows = []
            for r in ET.fromstring(z.read(sf)).iter("{%s}row" % ns["m"]):
                vals = {}
                for i, c in enumerate(r.findall("m:c", ns)):
                    idx = _col_index(c.get("r", "")) if c.get("r") else i
                    t = c.get("t")
                    v = c.find("m:v", ns)
                    if t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter("{%s}t" % ns["m"]))
                    elif v is None or v.text is None:
                        val = None
                    elif t == "s":
                        val = shared[int(v.text)] if int(v.text) < len(shared) else ""
                    elif t in ("str", "e"):
                        val = v.text
                    elif t == "b":
                        val = v.text == "1"
                    else:
                        try:
                            f = float(v.text)
                            val = int(f) if f.is_integer() else f
                        except ValueError:
                            val = v.text
                    vals[idx] = val
                rows.append([vals.get(i) for i in range(max(vals) + 1)] if vals else [])
            sheets.append(rows)
        return sheets


def iter_sheet_rows(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        return [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
    except Exception as e:
        print(f"[알림] openpyxl 로 열지 못해 값만 직접 읽습니다: {type(e).__name__}: {e}")
        return read_xlsx_raw(path)


def read_query_date(path):
    """엑셀 위쪽의 '조회일 : 2026.11.16' -> '20261116' (없으면 '')"""
    try:
        for sheet in iter_sheet_rows(path):
            for row in sheet[:6]:
                for c in row:
                    if c and "조회일" in str(c):
                        m = re.search(r"(\d{4})\D?(\d{1,2})\D?(\d{1,2})", str(c))
                        if m:
                            return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"
    except Exception:
        pass
    return ""


def save_config_keys(values):
    cfg = load_config()
    cfg.update(values)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def read_schedule_xlsx(path):
    data, found, skipped = [], False, []
    for sheet in iter_sheet_rows(path):     # 머리글이 있는 시트를 찾음
        col = None
        for row in sheet:
            row = list(row)
            cells = ["" if c is None else str(c).strip() for c in row]
            if col is None:
                if "방송일" in cells and "프로그램명" in cells:
                    col = {k: cells.index(v) for k, v in HEADERS.items() if v in cells}
                    missing = [HEADERS[k] for k in ("date", "name", "start", "end") if k not in col]
                    if missing:
                        raise ValueError(f"엑셀에 필요한 열이 없습니다: {missing}")
                    found = True
                continue
            raw = lambda k: row[col[k]] if k in col and col[k] < len(row) else None
            g = lambda k: cells[col[k]] if k in col and col[k] < len(cells) else ""
            date = to_yyyymmdd(raw("date")) if raw("date") is not None else ""
            if not date:
                continue
            try:
                s, e = hhmm_to_min(raw("start")), hhmm_to_min(raw("end"))
            except (ValueError, TypeError):
                skipped.append(f"{date} {g('name')} ({g('start')}~{g('end')})")
                continue
            if e <= s:
                e += 1440 if e < DAY_START_MIN else 0
            data.append({
                "date": date, "code": g("code"), "name": g("name"),
                "s": s - DAY_START_MIN, "e": e - DAY_START_MIN,
                "uhd": g("uhd") == "Y", "cc": g("cc") == "Y", "ad": g("ad") == "Y", "sign": g("sign") == "Y",
                "age": g("age"), "rl": g("rl"), "rerun": g("rerun") == "Y", "plan": g("plan") or "1안",
            })
        if col is not None and data:
            break
    if not found:
        raise ValueError(f"엑셀에서 '방송일', '프로그램명' 머리글을 찾지 못했습니다.\n({os.path.basename(path)})")
    if not data:
        raise ValueError("머리글은 찾았지만 읽을 수 있는 편성 줄이 없습니다. 방송일·편성시각 칸 형식을 확인하세요."
                         + (f"\n예: {skipped[0]}" if skipped else ""))
    if skipped:
        print(f"[주의] 시각을 읽지 못한 줄 {len(skipped)}개: {skipped[:3]}")
    return data


# ---------------------------------------------------------------- 칸(블록) 만들기
NAME_DICT_PATH = os.path.join(DATA_DIR, "name_dict.json")
if DATA_DIR != APP_DIR and not os.path.exists(NAME_DICT_PATH) and os.path.exists(os.path.join(APP_DIR, "name_dict.json")):
    try:        # 처음 실행: 프로그램과 함께 온 기본 이름 사전을 데이터 폴더로 복사
        import shutil
        shutil.copyfile(os.path.join(APP_DIR, "name_dict.json"), NAME_DICT_PATH)
    except Exception:
        pass
_name_dict = None


def norm_key(s):
    return re.sub(r"\s+", "", s.replace(" ", " "))


def split_suffix(s):
    m = re.match(r"^(.*?)\s*(\([^()]*\))$", s.strip(), re.S)
    return (m.group(1), m.group(2)) if m else (s.strip(), "")


def name_dict():
    global _name_dict
    if _name_dict is None:
        try:
            with open(NAME_DICT_PATH, "r", encoding="utf-8") as f:
                _name_dict = json.load(f)
        except Exception:
            _name_dict = {}
    return _name_dict


def clean_name(name, cfg):
    """엑셀 이름 -> 편성표 표시 이름. 이름 사전(예전 PDF에서 배운 것)이 먼저, 없으면 규칙 적용.
    줄바꿈 위치는 '\\n' 으로 들어 있음"""
    name = re.sub("[​‌‍﻿]", "", name)
    main, suf = split_suffix(name.replace(" ", " "))
    learned = name_dict().get(norm_key(main))
    if learned:
        return re.sub(r"\s*\|\s*", "\n", learned + suf)
    for a, b in cfg.get("name_replace", []):
        name = name.replace(a, b)
    for pat, rep in cfg.get("name_regex", []):
        try:
            name = re.sub(pat, rep, name)
        except re.error:
            pass
    name = re.sub(r"^(특집|스페셜)(?=\d)", r"\1 ", name.strip())     # 특집2시 뉴스 -> 특집 2시 뉴스
    # 직접 넣은 줄바꿈: 프로그램명에 '|' 를 쓰면 그 자리에서 다음 줄로
    return re.sub(r"\s*\|\s*", "\n", name).strip()


PINK_KEYS = ((1.0, 1.0, 1.0), (1.0, 0.88, 0.88))


def extract_pdf_cells(pdf_path):
    """편성표 PDF(원본이든 이 프로그램이 만든 것이든)에서 프로그램 칸을 읽음.
    반환: [{day0, day1, half, start(05:00 기준 분), pink, lines[(글자,크기)], rect}]"""
    doc = pymupdf.open(pdf_path)
    pg = doc[0]
    dr = pg.get_drawings()
    fill = lambda d: tuple(round(c, 2) for c in d["fill"]) if d.get("fill") else None
    cells = [(d["rect"], fill(d)) for d in dr if fill(d) and fill(d) != (0.72, 0.8, 0.89)
             and fill(d) != (0.0, 0.0, 0.0) and d["rect"].width > 5 and d["rect"].height > 1.5
             and d["rect"].y0 > 118 and d["rect"].height < 400       # 요일 바탕(세로 전체) 칸은 제외
             and 55 <= d["rect"].x0 < 541]
    cells = [(c, f) for c, f in cells
             if not any(o is not c and o.contains(c) and o.get_area() > c.get_area() * 1.05 for o, _ in cells)]
    hours = sorted({(round(d["rect"].y0, 1), round(d["rect"].y1, 1)) for d in dr
                    if fill(d) == (0.72, 0.8, 0.89) and abs(d["rect"].x0 - 42.0) < 1.5 and d["rect"].y0 > 115})
    # 아이콘 글자(흰 글자, 노란 U 위 진회색 글자)는 프로그램 이름에서 뺌
    spans = [s for b in pg.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"]
             if s["text"].strip() and s.get("color", 0) not in (0xFFFFFF, 0x1A1A1A)]
    out = []
    for c, f in cells:
        days = X_COLS["days"]
        d0 = next((i for i, (a, b) in enumerate(days) if a - 1 <= c.x0 < b - 1), None)
        d1 = max((i for i, (a, b) in enumerate(days) if c.x1 > a + 1 and c.x0 < b - 1), default=d0)
        if d0 is None:
            continue
        span_w = days[d1][1] - days[d0][0]
        half = None
        if c.width < span_w * 0.7:
            half = "L" if abs(c.x0 - days[d0][0]) < 1 else "R"
        inside = [s for s in spans if c.contains(pymupdf.Point((s["bbox"][0] + s["bbox"][2]) / 2,
                                                              (s["bbox"][1] + s["bbox"][3]) / 2))]
        minute, lines = None, {}
        for s in sorted(inside, key=lambda s: (round(s["bbox"][3], 0), s["bbox"][0])):
            t = s["text"].strip()
            if minute is None and s["bbox"][0] - c.x0 < 2.5 and s["bbox"][1] - c.y0 < 4.5 and re.match(r"^\d\d(\s|$)", t):
                minute = int(t[:2])
                t = t[2:].strip()
                if not t:
                    continue
            if s["bbox"][0] - c.x0 < 2.5 and c.y1 - s["bbox"][3] < 2.5 and re.fullmatch(r"~?(\d\d:)?\d\d", t) \
                    and c.height > 9:
                continue        # 왼쪽 아래 '끝 시각' 표시는 이름이 아님
            lines.setdefault(round(s["bbox"][3], 0), []).append((t, s["size"]))
        hi = max((i for i, (y0, _) in enumerate(hours) if y0 <= c.y0 + 0.6), default=0)
        start = hi * 60 + (minute or 0)
        lines = [(re.sub(r"\s+", " ", " ".join(t for t, _ in v)).strip(), max(sz for _, sz in v))
                 for _, v in sorted(lines.items())]
        out.append(dict(day0=d0, day1=d1, half=half, start=start, pink=f == (1.0, 0.88, 0.88), bg=f,
                        lines=lines, rect=c, text=" ".join(t for t, _ in lines)))
    doc.close()
    return out


def text_sig(t):
    """줄바꿈·순서와 무관하게 같은 글자인지 보기 위한 서명"""
    t = re.sub("[\u2010-\u2015\u2212]", "-", t)
    t = re.sub("[\u200b-\u200d\ufeff]|\\s", "", t)
    return "".join(sorted(t))


def compare_pdfs(base_pdf, other_pdf, log=print):
    """두 편성표 PDF를 칸 단위로 비교 (기준 = 완성본/원본). 반환: 요약 dict"""
    A, B = extract_pdf_cells(base_pdf), extract_pdf_cells(other_pdf)
    key = lambda c: (c["day0"], c["half"], c["start"])
    bmap = {}
    for c in B:
        bmap.setdefault(key(c), c)
    wk = "월화수목금토일"
    where = lambda c: f"{wk[c['day0']]} {fmt_t(c['start'])}" + (f"({'1안' if c['half'] == 'L' else '2안'})" if c["half"] else "")
    res = {"기준 칸": len(A), "비교 칸": len(B), "일치": 0, "이름 다름": 0, "날짜 묶음 다름": 0,
           "분홍 다름": 0, "음영 다름": 0, "기준에만 있음": 0, "비교에만 있음": 0}
    detail = []
    diff_keys = []
    used = set()
    for a in A:
        b = bmap.get(key(a))
        if b is None:
            diff_keys.append(key(a))
            res["기준에만 있음"] += 1
            detail.append(f"[없음] {where(a)} {a['text']}")
            continue
        used.add(id(b))
        ok = True
        if text_sig(a["text"]) != text_sig(b["text"]):
            res["이름 다름"] += 1; ok = False
            detail.append(f"[이름] {where(a)} 기준 '{a['text']}' / 비교 '{b['text']}'")
        if a["day1"] != b["day1"]:
            res["날짜 묶음 다름"] += 1; ok = False
            detail.append(f"[묶음] {where(a)} {a['text']}: 기준 {wk[a['day0']]}~{wk[a['day1']]} / 비교 {wk[b['day0']]}~{wk[b['day1']]}")
        if a["pink"] != b["pink"]:
            res["분홍 다름"] += 1; ok = False
            detail.append(f"[분홍] {where(a)} {a['text']}: 기준 {'분홍' if a['pink'] else '흰색'} / 비교 {'분홍' if b['pink'] else '흰색'}")
        if not a["pink"] and not b["pink"] and max(abs(x - y) for x, y in zip(a["bg"], b["bg"])) > 0.03:
            res["음영 다름"] += 1; ok = False
            detail.append(f"[음영] {where(a)} {a['text']}: 기준 {a['bg']} / 비교 {b['bg']}")
        if ok:
            res["일치"] += 1
        else:
            diff_keys.append(key(a))
    for b in B:
        if id(b) not in used:
            res["비교에만 있음"] += 1
            detail.append(f"[추가] {where(b)} {b['text']}")
    log("=" * 60)
    log(f"[비교] 기준: {os.path.basename(base_pdf)}")
    log(f"       비교: {os.path.basename(other_pdf)}")
    log("       " + ", ".join(f"{k} {v}" for k, v in res.items()))
    rate = res["일치"] / max(1, res["기준 칸"]) * 100
    log(f"       칸 일치율 {rate:.1f}%  (아이콘은 비교하지 않음)")
    for line in detail:
        log("  " + line)
    res["일치율"] = rate
    res["다른 칸"] = diff_keys
    res["상세"] = detail
    return res


def learn_names(xlsx, pdf_path, log=print):
    """같은 주의 엑셀과 원본 편성표 PDF를 비교해, PDF에 쓰인 띄어쓰기·줄바꿈을 이름 사전에 저장"""
    rows = read_schedule_xlsx(xlsx)
    by_norm = {}
    for r in rows:
        by_norm.setdefault(norm_key(r["name"]), r["name"])
    doc = pymupdf.open(pdf_path)
    pg = doc[0]
    font = pymupdf.Font("korea")
    cells = [d["rect"] for d in pg.get_drawings()
             if d.get("fill") and d["rect"].width > 5 and d["rect"].y0 > 110
             and tuple(round(c, 2) for c in d["fill"]) in ((1.0, 1.0, 1.0), (1.0, 0.88, 0.88))]
    spans = [s for b in pg.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"]
             if s["text"].strip()]
    d = name_dict()
    learned = 0
    for c in cells:
        inside = [s for s in spans if c.contains(pymupdf.Point((s["bbox"][0] + s["bbox"][2]) / 2,
                                                              (s["bbox"][1] + s["bbox"][3]) / 2))]
        lines = {}
        for s in sorted(inside, key=lambda s: (round(s["bbox"][3], 0), s["bbox"][0])):
            t = s["text"]
            if s["bbox"][0] - c.x0 < 2.5 and s["bbox"][1] - c.y0 < 4.5 and re.match(r"^\d\d(\s|$)", t):
                t = t[2:].strip()     # 왼쪽 위 '분' 표시 빼기
                if not t:
                    continue
            key = round(s["bbox"][3], 0)
            lines.setdefault(key, []).append((t, s["size"]))
        if not lines:
            continue
        lines = [(re.sub(r"\s+", " ", " ".join(t for t, _ in v)).strip(), max(sz for _, sz in v))
                 for _, v in sorted(lines.items())]
        full = "".join(t for t, _ in lines)
        raw = by_norm.get(norm_key(full))
        if not raw:
            continue
        main_raw, suf_raw = split_suffix(raw)
        # PDF 줄에서 뒤의 (…) 빼기
        txt_lines = [t for t, _ in lines]
        if suf_raw:
            last = txt_lines[-1]
            if norm_key(last) == norm_key(suf_raw):
                txt_lines.pop()
            elif norm_key(last).endswith(norm_key(suf_raw)):
                txt_lines[-1] = last[: last.rfind("(")].rstrip()
        # 칸이 좁아서 저절로 넘어간 줄은 띄어쓰기로, 일부러 나눈 줄은 줄바꿈으로
        out = txt_lines[0] if txt_lines else ""
        for i in range(1, len(txt_lines)):
            prev, nxt = txt_lines[i - 1], txt_lines[i]
            size = lines[min(i, len(lines) - 1)][1]
            first_word = nxt.split(" ")[0]
            auto = font.text_length(prev + " " + first_word, fontsize=size) > c.width - 3
            out += (" " if auto else "\n") + nxt
        out = "\n".join(x.strip() for x in out.split("\n")).strip()
        if norm_key(out) != norm_key(main_raw):
            continue
        k = norm_key(main_raw)
        if out != main_raw and d.get(k) != out:
            d[k] = out
            learned += 1
    with open(NAME_DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1, sort_keys=True)
    log(f"[이름 사전] 새로 배운 표시 이름 {learned}개 (전체 {len(d)}개) -> {NAME_DICT_PATH}")
    return learned


SHADE_PRESETS = {"없음": "", "회색": "#d9d9d9", "연회색": "#eeeeee", "진회색": "#bfbfbf",
                 "노랑": "#fff2b3", "하늘": "#d6ecff", "연두": "#dff5d6", "보라": "#ece0f7"}


HL_PRESETS = {"없음": "", "하늘": "#a6d8f0", "노랑": "#fff176", "연두": "#c5e1a5", "분홍": "#f8bbd0",
              "주황": "#ffcc80", "회색": "#d9d9d9"}
FG_PRESETS = {"검정": "", "빨강": "#d32f2f", "파랑": "#1565c0", "초록": "#2e7d32", "보라": "#6a1b9a",
              "주황": "#e65100", "흰색": "#ffffff"}


def hex_to_rgb(h):
    h = (h or "").strip().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return None


def rule_shade(r, cfg):
    for rule in cfg.get("shade_rules", []):
        m = str(rule.get("match", "")).strip()
        if m and (norm_key(m) in norm_key(r.get("name", "")) or (r.get("code") or "").startswith(m)):
            return rule.get("color", "#d9d9d9")
    return ""


def rule_end(r, cfg):
    return any(m and norm_key(str(m)) in norm_key(r.get("name", "")) for m in cfg.get("end_label_rules", []))


def build_blocks(rows, cfg, prev_rows=None, log=print):
    dates = sorted({r["date"] for r in rows})
    if len(dates) != 7:
        log(f"[주의] 날짜가 {len(dates)}일입니다 (보통 7일): {dates}")
    di = {d: i for i, d in enumerate(dates)}
    ridx = {id(r): i for i, r in enumerate(rows)}     # 칸 -> 원래 편성 줄 번호 (화면에서 클릭할 때 씀)

    prev_keys = None
    if prev_rows and cfg.get("pink_changed_from_prev"):
        pd = sorted({r["date"] for r in prev_rows})
        pwd = {d: datetime.datetime.strptime(d, "%Y%m%d").weekday() for d in pd}
        prev_keys = {(pwd[r["date"]], r["s"], r["code"]) for r in prev_rows if r["plan"] == "1안"}

    hl = set()
    for item in cfg.get("highlight", []):
        m = re.match(r"(\d{4})-?(\d{2})-?(\d{2})\s+(\d{1,2}:\d{2})", str(item))
        if m:
            hl.add((m.group(1) + m.group(2) + m.group(3), hhmm_to_min(m.group(4)) - DAY_START_MIN))

    blocks = []
    for d in dates:
        day = [r for r in rows if r["date"] == d]
        p1 = sorted([r for r in day if r["plan"] != "2안"], key=lambda r: r["s"])
        p2 = sorted([r for r in day if r["plan"] == "2안"], key=lambda r: r["s"])
        # 2안 구간(이어진 묶음) 찾기 -> 그 구간과 겹치는 1안은 왼쪽 반, 2안은 오른쪽 반
        windows = []
        for r in p2:
            if windows and r["s"] <= windows[-1][1]:
                windows[-1][1] = max(windows[-1][1], r["e"])
            else:
                windows.append([r["s"], r["e"]])
        changed = True
        while changed:   # 구간에 걸친 1안까지 포함하도록 넓힘
            changed = False
            for w in windows:
                for r in p1:
                    if r["s"] < w[1] and r["e"] > w[0] and (r["s"] < w[0] or r["e"] > w[1]):
                        w[0], w[1] = min(w[0], r["s"]), max(w[1], r["e"])
                        changed = True
        in_win = lambda r: any(r["s"] >= w[0] and r["e"] <= w[1] for w in windows)
        wd = datetime.datetime.strptime(d, "%Y%m%d").weekday()
        for r in p1 + p2:
            half = None
            if in_win(r):
                half = "R" if r["plan"] == "2안" else "L"
            pink = bool(half) and cfg.get("pink_plan2", True)
            if prev_keys is not None and r["plan"] != "2안" and (wd, r["s"], r["code"]) not in prev_keys:
                pink = True
            if (d, r["s"]) in hl:
                pink = True
            fp = r.get("force_pink")        # True: 켬 / False: 끔(2안·지난주 자동 분홍도 끔) / 없음: 자동
            if fp is True:
                pink = True
            elif fp is False:
                pink = False
            blocks.append(dict(r, day0=di[d], day1=di[d], half=half, pink=pink, _ris=[ridx[id(r)]],
                               _mark=is_marked(r, cfg),
                               label=clean_name(r["name"], cfg),
                               shade=r.get("shade") or "", rule_shade=rule_shade(r, cfg),
                               show_end=bool(r.get("show_end")) or rule_end(r, cfg)))
        # 중간에 빈 시간 경고 (방송 종료 뒤는 제외)
        t = None
        for r in p1:
            if t is not None and r["s"] > t:
                log(f"[빈 시간] {d} {fmt_t(t)}~{fmt_t(r['s'])} 엑셀에 편성 없음")
            t = r["e"] if t is None else max(t, r["e"])

    # 이어진 날짜 묶기: 평일(0~4), 주말(5~6) 안에서 같은 코드·같은 시각이면 한 칸으로
    merged = []
    groups = [range(0, 5), range(5, 7)]
    used = set()
    key = lambda b: (b["code"], b["s"], b["e"])
    full = [b for b in blocks if b["half"] is None]
    by = {}
    for b in full:
        by.setdefault((b["day0"],) + key(b), b)
    for b in sorted(full, key=lambda b: (b["day0"], b["s"])):
        if id(b) in used:
            continue
        g = next((g for g in groups if b["day0"] in g), None)
        end = b["day0"]
        while g is not None and end + 1 in g:
            nb = by.get((end + 1,) + key(b))
            if nb is None or id(nb) in used or not b["code"]:
                break
            if nb.get("shade", "") != b.get("shade", "") or nb.get("rule_shade", "") != b.get("rule_shade", "") \
                    or nb.get("show_end") != b.get("show_end") or bool(nb.get("bold")) != bool(b.get("bold")) \
                    or (nb.get("hl") or "") != (b.get("hl") or "") or (nb.get("fg") or "") != (b.get("fg") or "") \
                    or bool(nb.get("top_dash")) != bool(b.get("top_dash")) \
                    or bool(nb.get("bottom_dash")) != bool(b.get("bottom_dash")):
                break
            if (b.get("name_edited") or nb.get("name_edited")) and nb["label"] != b["label"]:
                break
            used.add(id(nb))
            end += 1
            b["pink"] = b["pink"] or nb["pink"]
            b["edited"] = bool(b.get("edited") or nb.get("edited"))
            b["_mark"] = bool(b.get("_mark") or nb.get("_mark"))
            b["_ris"] = b["_ris"] + nb["_ris"]
        used.add(id(b))
        merged.append(dict(b, day1=end))
    merged += [b for b in blocks if b["half"] is not None]
    return dates, merged


def fmt_t(t):
    m = t + DAY_START_MIN
    return f"{m // 60:02d}:{m % 60:02d}"


# ---------------------------------------------------------------- 그리기
PAGE_W, PAGE_H = 595.0, 842.0
X_COLS = {  # 원본 PDF에서 잰 좌표(pt)
    "gL": (29.0, 42.0), "hL": (42.0, 55.4),
    "days": [(55.4, 121.0), (121.0, 186.6), (186.6, 252.2), (252.2, 317.8), (317.8, 383.2),
             (409.8, 475.5), (475.5, 541.1)],
    "gM": (383.2, 396.5), "hM": (396.5, 409.8), "hR": (541.1, 554.4), "gR": (554.4, 567.7),
}
TABLE_TOP, HEAD_H, TABLE_BOTTOM = 107.5, 13.1, 766.6
BLUE = (0.7216, 0.8, 0.8941)
PINK = (1.0, 0.8824, 0.8824)
LINE = (0.62, 0.62, 0.62)
LINE_HEX = "#9e9e9e"            # 칸 테두리 기본 색 (회색)
CELL_LINE_W = 0.3               # 칸 테두리 기본 굵기 (pt)
FRAME_LINE_W = 0.6              # 바깥 테두리 기본 굵기 (pt)
CELL_LINE_WS = ("0.2", "0.3", "0.4", "0.5", "0.75", "1.0")
FRAME_LINE_WS = ("0.6", "0.8", "1.0", "1.5")


EDIT_LINE_W = 0.75              # 고친 칸 테두리 기본 굵기 (pt)
EDIT_LINE_WS = ("끄기", "0.2", "0.3", "0.4", "0.5", "0.75", "1.0", "1.5")
TIME_SZ = 4.4                   # 칸 왼쪽 위 시각(분) 글자 기본 크기 (pt)
LOCAL_TIME_SZ = 6.0             # 로컬 칸 시각 글자 기본 크기 (pt)


TIME_DX = 0.3                   # 왼쪽 선 ↔ 시각 숫자 기본 간격 (pt)
TIME_DY = 0.0                   # 시각 숫자 위아래 옮김 기본값 (pt, +아래로 / -위로)
FINE_STEP = 0.05                # ⚙ 설정의 ▲▼ 한 번에 바뀌는 값 (pt)


def fine_value(v, lo, hi, default):
    """설정의 세밀한 값 (0.01pt 까지, lo~hi 밖이거나 이상하면 기본값)"""
    try:
        v = round(float(v), 2)
        return v if lo <= v <= hi else default
    except (TypeError, ValueError):
        return default


def time_size(v, default):
    """설정의 시각 글자 크기 (2~10pt, 이상하면 기본값)"""
    return fine_value(v, 2.0, 10.0, default)


def edit_mark_cfg(cfg):
    """고친 칸 따로 표시 설정 -> (켬, 테두리 굵기 또는 None(끄기), 시각 크기)  ※ 색은 edited_line_color"""
    on = bool(cfg.get("edited_mark", not IS_DMB))
    v = cfg.get("edited_line_w", EDIT_LINE_W)
    lw = None if v in (0, "0", "끄기", None, "") else line_width(v, EDIT_LINE_W)
    return on, lw, time_size(cfg.get("edited_time_sz"), TIME_SZ)


def edit_time_gap(cfg):
    """고친 칸 시각 숫자 간격 -> (왼쪽 간격, 위아래 옮김)"""
    return (fine_value(cfg.get("edited_time_dx"), 0.0, 3.0, TIME_DX),
            fine_value(cfg.get("edited_time_dy"), -1.0, 3.0, TIME_DY))


DASH_W = 0.5                    # 점선 기본 굵기 (pt)
DASH_HEX = "#595959"            # 점선 기본 색 (진회색)
DASH_STYLES = {"촘촘": (1.4, 1.1), "보통": (2.5, 1.5), "길게": (4.0, 2.0), "점": (0.0, 1.3)}   # (칠함, 빈 칸) pt
DASH_STYLE = "보통"


def dash_cfg(cfg):
    """점선 설정 -> (굵기, 색 rgb, 모양 이름, 빈 칸을 음영색으로)"""
    w = fine_value(cfg.get("dash_w"), 0.1, 3.0, DASH_W)
    c = hex_to_rgb(cfg.get("dash_color")) or hex_to_rgb(DASH_HEX)
    st = cfg.get("dash_style") if cfg.get("dash_style") in DASH_STYLES else DASH_STYLE
    return w, c, st, bool(cfg.get("dash_gap_shade", True))


def dash_pattern(style, w):
    """점선 모양 -> (PDF dashes 문자열, 선 끝 모양). 굵은 선일수록 마디도 그만큼 길게"""
    on, off = DASH_STYLES.get(style, DASH_STYLES[DASH_STYLE])
    k = max(1.0, w / 0.5)
    if style == "점":
        return f"[0 {off * k + w:.2f}] 0", 1
    return f"[{on * k:.2f} {off * k:.2f}] 0", 0


def line_width(v, default):
    """설정의 선 굵기 (0.1~3pt, 이상하면 기본값)"""
    try:
        v = float(v)
        return v if 0.1 <= v <= 3 else default
    except (TypeError, ValueError):
        return default
DARK = (0.15, 0.15, 0.15)
ICON_COLORS = {"U": (233, 208, 35), "자": (117, 148, 204), "수": (249, 86, 35), "해": (194, 24, 91),
               "R": (36, 144, 64), "L": (198, 148, 122), "AGE": (0, 0, 0)}
ICON_ORIGINAL = dict(ICON_COLORS, 해=(243, 97, 220))      # 원본 PDF 색 (해: 연분홍)
ICON_DARK_TEXT = {"U"}                                      # 이 아이콘들은 글자를 진회색으로
ICON_NAMES = (("U", "UHD"), ("자", "자막"), ("수", "한국수어"), ("해", "화면해설"), ("R", "전국릴레이"),
              ("L", "부분로컬"), ("AGE", "연령"))


def icon_palette(cfg):
    """설정의 icon_colors / icon_dark_text 를 반영한 (배경색, 글자색) 표"""
    bg, fg = {}, {}
    over = cfg.get("icon_colors", {}) or {}
    dark = set(cfg.get("icon_dark_text", sorted(ICON_DARK_TEXT)))
    for k, rgb in ICON_COLORS.items():
        bg[k] = hex_to_rgb(over.get(k)) or tuple(c / 255 for c in rgb)
        fg[k] = (0.1, 0.1, 0.1) if k in dark else (1, 1, 1)
    return bg, fg
ICON = 4.9
FONT_SIZES = [6.5, 5.8, 5.1, 4.4, 3.6, 3.3, 2.8, 2.2]


def find_font(cfg):
    for f in cfg.get("font_files", []):
        if os.path.exists(f):
            return f
    return None


_FONT_CACHE = {}


class Painter:
    """도형은 한 Shape 에, 글자는 색깔별 TextWriter 에 모았다가 finish() 때 한 번에 그림 (빠름)
    -> 도형이 먼저, 글자가 나중에 그려지므로 글자는 항상 도형 위에 보임"""
    def __init__(self, page, fontfile):
        self.page = page
        self.font = _FONT_CACHE.get(fontfile)
        if self.font is None and fontfile:
            try:
                self.font = pymupdf.Font(fontfile=fontfile)
            except Exception:
                self.font = None
        if self.font is None:
            self.font = pymupdf.Font("korea")
        _FONT_CACHE[fontfile] = self.font
        self.shape = page.new_shape()
        self.writers = {}

    def width(self, text, size):
        return self.font.text_length(text, fontsize=size)

    bold = False        # True 면 굵게 (같은 글자를 살짝 옆으로 한 번 더)

    def text(self, x, y, s, size, color=(0, 0, 0)):
        """(x, y) = 글자 왼쪽, 기준선"""
        key = tuple(round(c, 3) for c in color)
        tw = self.writers.get(key)
        if tw is None:
            tw = self.writers[key] = pymupdf.TextWriter(self.page.rect, color=key)
        tw.append((x, y), s, font=self.font, fontsize=size)
        if self.bold:
            tw.append((x + size * 0.045, y), s, font=self.font, fontsize=size)

    def text_center(self, cx, cy, s, size, color=(0, 0, 0)):
        w = self.width(s, size)
        self.text(cx - w / 2, cy + size * 0.36, s, size, color)

    def rect(self, x0, y0, x1, y1, fill=None, stroke=LINE, width=0.3):
        self.shape.draw_rect(pymupdf.Rect(x0, y0, x1, y1))
        self.shape.finish(color=stroke, fill=fill, width=width if stroke else 0)

    def line(self, x0, y0, x1, y1, color=LINE, width=0.3, dashes=None, cap=0):
        self.shape.draw_line(pymupdf.Point(x0, y0), pymupdf.Point(x1, y1))
        # closePath=False: 선은 닫지 않음 (닫으면 되돌아오는 선이 점선 틈을 메워 실선처럼 보임)
        # cap=1: 둥근 끝 (점 모양 점선)
        self.shape.finish(color=color, width=width, dashes=dashes, closePath=False, lineCap=cap)

    def finish(self):
        self.shape.commit()
        for tw in self.writers.values():
            tw.write_text(self.page)


def icon_list(b):
    out = []
    if b["uhd"]:
        out.append("U")
    if b["sign"]:
        out.append("수")
    if b["ad"]:
        out.append("해")
    if b["rl"] == "R":
        out.append("R")
    elif b["rl"] == "V":
        out.append("L")
    age = {"00": "A", "07": "7", "12": "12", "15": "15", "19": "19"}.get(b["age"])
    if age:
        out.append(("AGE", age))
    if b["cc"]:
        out.append("자")
    return out


# ---------------------------------------------------------------- DMB 변환 (같은 편성, 출력 모양만 다르게)
# 빼는 것: 키 -> 이름. 기본은 연령·분홍만 남기고 나머지 아이콘·글자색·음영은 뺌
DMB_HIDE_ITEMS = (("uhd", "UHD"), ("cc", "자막"), ("sign", "수어"), ("ad", "화면해설"), ("rl", "R/L"),
                  ("age", "연령"), ("hl", "글자 바탕색"), ("fg", "글자색"), ("bold", "글자 굵기"), ("shade", "음영"),
                  ("pink", "분홍"))
OWN_STYLE_KEYS = ("shade", "bold", "fg", "hl")
DMB_HIDE_DEFAULT = ("uhd", "cc", "sign", "ad", "rl", "hl", "fg", "bold", "shade")
ICON_KEY_OF = {"U": "uhd", "자": "cc", "수": "sign", "해": "ad", "R": "rl", "L": "rl", "AGE": "age"}


def dmb_title_auto(tv_title):
    """TV 제목에서 DMB 제목 제안: '광주MBC주간편성표[11월2주]' -> '광주 MYMBC 주간편성표[11월2주]'"""
    t = re.sub(r"(?<![A-Za-z])MBC\s*", "MYMBC ", tv_title or "MBC 주간편성표", count=1)
    return re.sub(r"(\S)MYMBC", r"\1 MYMBC", t).replace("  ", " ").strip()


def dmb_hidden(cfg):
    """DMB 로 그릴 때 뺄 것 (DMB 가 아니면 빈 집합)"""
    d = cfg.get("_dmb")
    if not d:
        return set()
    return set(d.get("hide", DMB_HIDE_DEFAULT))


def draw_icons(p, x1, top, icons, scale=1.0):
    s = ICON * scale
    x = x1 - s * len(icons)
    for ic in icons:
        if isinstance(ic, tuple):
            key, label = ic
        else:
            key, label = ic, ic
        bg, fg = getattr(p, "icon_pal", None) or icon_palette({})
        p.rect(x, top, x + s, top + s, fill=bg[key], stroke=None)
        fs = s * (0.62 if len(label) > 1 else 0.78)
        p.text_center(x + s / 2, top + s / 2, label, fs, fg[key])
        x += s
    return s * len(icons)


SUFFIX_RE = re.compile(r"^(.*?)\s*(\([^()]*\))$", re.S)


def wrap_words(p, text, size, maxw):
    words = text.split(" ")
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if p.width(t, size) <= maxw:
            cur = t
            continue
        if cur:
            lines.append(cur)
        cur = ""
        if p.width(w, size) <= maxw:
            cur = w
        else:   # 한 단어가 너무 길면 글자 단위로 자름
            for ch in w:
                if p.width(cur + ch, size) <= maxw:
                    cur += ch
                else:
                    lines.append(cur)
                    cur = ch
    if cur:
        lines.append(cur)
    return lines


def layout_label(p, label, w, h):
    """칸에 들어갈 줄 목록 [(글자, 크기)] 과 줄 높이 합"""
    m = SUFFIX_RE.match(label)
    main, suf = (m.group(1), m.group(2)) if m else (label, "")
    forced = bool(suf) and bool(re.search(r"\n\s*\([^()]*\)\s*$", label))   # '|(목포M)' 처럼 직접 줄바꿈
    for size in FONT_SIZES:
        lh = size * 1.18
        ss = round(size * 0.8, 1)
        parts = main.split("\n") if main else []
        if parts and size < 4.0 and len(parts) > 1:
            parts = [" ".join(parts)]       # 아주 작은 칸은 한 줄로
        lines = [ln for part in parts for ln in wrap_words(p, part, size, w)]
        items = [(ln, size) for ln in lines]
        if suf and forced:
            items.append((suf, size))
        elif suf:
            short = suf == "(재)"
            if items and (short or h < lh * (len(items) + 1)) and \
                    p.width(items[-1][0], size) + p.width(suf, ss) + 1 <= w:
                items[-1] = (items[-1][0], size, suf, ss)
            else:
                items.append((suf, ss))
        total = sum(it[1] * 1.18 for it in items)
        if total <= h + 0.5 and all(p.width(it[0], it[1]) + (p.width(it[2], it[3]) if len(it) > 2 else 0) <= w + 0.5
                                    for it in items):
            return items, total
    return [(label.replace("\n", " "), FONT_SIZES[-1])], FONT_SIZES[-1] * 1.18


def time_scale(blocks):
    """시간 -> y 좌표. 짧은 구간도 최소 높이를 갖도록 구간별 높이를 정한 뒤 표 높이에 맞춤"""
    pts = {0, DAY_LEN_MIN}
    pts.update(range(0, DAY_LEN_MIN, 60))
    for b in blocks:
        pts.add(max(0, min(DAY_LEN_MIN, b["s"])))
        pts.add(max(0, min(DAY_LEN_MIN, b["e"])))
    pts = sorted(pts)
    avail = TABLE_BOTTOM - (TABLE_TOP + HEAD_H)
    k = avail / DAY_LEN_MIN
    seg = [max((b - a) * k, 2.6) for a, b in zip(pts, pts[1:])]
    scale = avail / sum(seg)
    ys = [TABLE_TOP + HEAD_H]
    for s in seg:
        ys.append(ys[-1] + s * scale)
    table = dict(zip(pts, ys))

    def y_of(t):
        t = max(0, min(DAY_LEN_MIN, t))
        if t in table:
            return table[t]
        for a, b in zip(pts, pts[1:]):
            if a <= t <= b:
                return table[a] + (table[b] - table[a]) * (t - a) / (b - a)
        return ys[-1]
    return y_of


def make_stamp(cfg, printed_at, dates):
    """오른쪽 위 날짜 글자"""
    mode = cfg.get("stamp_mode", "저장 시각")
    wk = "월화수목금토일"
    if mode == "표시 안 함":
        return ""
    if mode == "직접 입력":
        return str(cfg.get("stamp_text", ""))
    if mode == "엑셀 조회일" and cfg.get("_query_date"):
        q = datetime.datetime.strptime(cfg["_query_date"], "%Y%m%d")
        return f"{q.year}년 {q.month}월 {q.day}일 ({wk[q.weekday()]})"
    if mode == "주간 범위" and dates:
        a = datetime.datetime.strptime(dates[0], "%Y%m%d")
        b = datetime.datetime.strptime(dates[-1], "%Y%m%d")
        return f"{a.year}년 {a.month}월 {a.day}일 ({wk[a.weekday()]}) ~ {b.month}월 {b.day}일 ({wk[b.weekday()]})"
    ampm = "AM" if printed_at.hour < 12 else "PM"
    h12 = printed_at.hour % 12 or 12
    return (f"{printed_at.year}년 {printed_at.month}월 {printed_at.day}일 ({wk[printed_at.weekday()]}) "
            f"{ampm} {h12:02d}:{printed_at.minute:02d}")


def draw_schedule(dates, blocks, cfg, out_path=None, printed_at=None):
    """out_path 가 있으면 저장, 없으면 문서(doc)를 그대로 돌려줌 (화면 미리보기용).
    반환: (doc 또는 None, 칸 위치 목록[{rect, block}])"""
    geom = []
    dash_lines = []
    cell_fills = []         # (x0, y0, x1, y1, 채움색): 점선 빈 칸을 위·아래 칸 색으로 칠할 때 씀
    doc = pymupdf.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    p = Painter(page, find_font(cfg))
    p.icon_pal = icon_palette(cfg)
    # 선 (⚙ 설정): 칸 테두리 굵기·색, 바깥 테두리 굵기
    cell_w = line_width(cfg.get("cell_line_w"), CELL_LINE_W)
    cell_c = hex_to_rgb(cfg.get("cell_line_color")) or LINE
    frame_w = line_width(cfg.get("frame_line_w"), FRAME_LINE_W)
    ed_on, ed_w, ed_t = edit_mark_cfg(cfg)     # 고친 칸: 테두리 굵기, 시각 크기
    ed_c = hex_to_rgb(cfg.get("edited_line_color")) or cell_c     # 고친 칸 테두리 색 (없으면 칸 테두리 색)
    ed_dx, ed_dy = edit_time_gap(cfg)                             # 고친 칸 시각 숫자 간격
    ed_lines = []
    y_of = time_scale(blocks)
    printed_at = printed_at or datetime.datetime.now()

    # 제목
    hide = dmb_hidden(cfg)
    dmb_title = (cfg.get("_dmb") or {}).get("title", "").strip()
    logo = None if dmb_title else cfg.get("logo_file")
    if logo and not os.path.isabs(logo):
        logo = next((os.path.join(d, logo) for d in (DATA_DIR, APP_DIR) if os.path.isfile(os.path.join(d, logo))), "")
    if logo and os.path.isfile(logo):
        page.insert_image(pymupdf.Rect(195.9, 33.8, 395.1, 82.8), filename=logo, keep_proportion=True)
    else:
        p.bold = bool(cfg.get("title_bold"))
        try:
            tsz = max(8.0, min(40.0, float(cfg.get("title_size", 22))))
        except (TypeError, ValueError):
            tsz = 22.0
        p.text_center(PAGE_W / 2, 56, dmb_title or cfg.get("title_text", "MBC 주간편성표"), tsz, (0.17, 0.24, 0.33))
        p.bold = False

    # 범례 줄
    p.rect(29, 86.5, 567.7, 98.5, fill=(0.9686, 0.9686, 0.9686), stroke=None)
    x = 40
    for key, label, name in (("U", "U", "UHD"), ("자", "자", "자막"), ("수", "수", "한국수어"), ("해", "해", "화면해설"),
                             ("R", "R", "전국릴레이"), ("L", "L", "부분로컬")):
        if ICON_KEY_OF[key] in hide:        # DMB: 뺀 아이콘은 범례에서도 뺌
            continue
        draw_icons(p, x + 6.5, 89.3, [(key, label)], scale=6.5 / ICON)
        p.text(x + 8, 94.8, name, 5.8)
        x += 10 + p.width(name, 5.8) + 6
    if "age" not in hide:
        for lab in ("A", "7", "12", "15", "19"):
            draw_icons(p, x + 6.5, 89.3, [("AGE", lab)], scale=6.5 / ICON)
            x += 7.6
        p.text(x + 1, 94.8, "연령", 5.8)
    stamp = make_stamp(cfg, printed_at, dates)
    if stamp:
        p.text(567.7 - p.width(stamp, 8.7), 95.5, stamp, 8.7, (0.2, 0.2, 0.2))

    # 머리글
    top, body = TABLE_TOP, TABLE_TOP + HEAD_H
    for key, txt in (("gL", "급"), ("hL", "시"), ("gM", "급"), ("hM", "시"), ("hR", "시"), ("gR", "급")):
        x0, x1 = X_COLS[key]
        p.rect(x0, top, x1, body, fill=BLUE, stroke=DARK)
        p.text_center((x0 + x1) / 2, top + HEAD_H / 2, txt, 8.2)
    for i, d in enumerate(dates[:7]):
        x0, x1 = X_COLS["days"][i]
        dt = datetime.datetime.strptime(d, "%Y%m%d")
        p.rect(x0, top, x1, body, fill=BLUE, stroke=DARK)
        p.text_center((x0 + x1) / 2, top + HEAD_H / 2, f"{dt.month}/{dt.day}({'월화수목금토일'[dt.weekday()]})", 8.7)

    # 시 열
    for key in ("hL", "hM", "hR"):
        x0, x1 = X_COLS[key]
        for hh in range(24):
            y0, y1 = y_of(hh * 60), y_of(hh * 60 + 60)
            p.rect(x0, y0, x1, y1, fill=BLUE, stroke=DARK)
            p.text_center((x0 + x1) / 2, (y0 + y1) / 2, str((hh + 5) % 12 or 12), 6.5)
    # 급 열
    for key, gk in (("gL", "grade_weekday"), ("gM", "grade_weekend"), ("gR", "grade_weekend")):
        x0, x1 = X_COLS[key]
        for s, e, g in cfg.get(gk, []):
            y0, y1 = y_of(hhmm_to_min(s) - DAY_START_MIN), y_of(hhmm_to_min(e) - DAY_START_MIN)
            p.rect(x0, y0, x1, y1, fill=BLUE, stroke=DARK)
            if g == "SA":
                p.text_center((x0 + x1) / 2, (y0 + y1) / 2 - 4.5, "S", 8.2)
                p.text_center((x0 + x1) / 2, (y0 + y1) / 2 + 4.5, "A", 8.2)
            else:
                p.text_center((x0 + x1) / 2, (y0 + y1) / 2, g, 8.2)

    # 날짜 칸 바탕 (편성 없는 시간은 흰 칸)
    for i in range(min(7, len(dates))):
        x0, x1 = X_COLS["days"][i]
        p.rect(x0, body, x1, TABLE_BOTTOM, fill=(1, 1, 1), stroke=cell_c, width=cell_w)

    # 프로그램 칸
    hide_all = hide
    own_keep = hide_all and (cfg.get("_dmb") or {}).get("keep_own", True)
    for b in sorted(blocks, key=lambda b: (b["day0"], b["s"])):
        if b["day0"] > 6:
            continue
        # DMB: DMB 편집기에서 직접 넣은 꾸밈(음영·굵기·글자색·글자 바탕)은 그대로 (TV 에서 온 꾸밈만 뺌)
        if own_keep and b.get("show") == "dmb":
            hide = hide_all - set(OWN_STYLE_KEYS)
        elif own_keep and b.get("_own"):
            hide = hide_all - set(b["_own"])
        else:
            hide = hide_all
        x0 = X_COLS["days"][b["day0"]][0]
        x1 = X_COLS["days"][min(b["day1"], 6)][1]
        if b["half"] == "L":
            x1 = (x0 + x1) / 2
        elif b["half"] == "R":
            x0 = (x0 + x1) / 2
        y0, y1 = y_of(b["s"]), y_of(b["e"])
        if y1 - y0 < 0.5:
            continue
        geom.append({"rect": (x0, y0, x1, y1), "block": b})
        fill = (None if "shade" in hide else hex_to_rgb(b.get("shade"))) \
            or (PINK if b["pink"] and "pink" not in hide else None) \
            or (None if "shade" in hide else hex_to_rgb(b.get("rule_shade"))) or (1, 1, 1)
        cell_fills.append((x0, y0, x1, y1, fill))
        if b.get("top_dash"):
            dash_lines.append((x0, y0, x1))
        if b.get("bottom_dash"):
            dash_lines.append((x0, y1, x1))
        p.rect(x0, y0, x1, y1, fill=fill, stroke=cell_c, width=cell_w)
        h = y1 - y0
        mm = f"{(b['s'] + DAY_START_MIN) % 60:02d}"
        ed = ed_on and bool(b.get("_mark", b.get("edited")))
        tsz = ed_t if ed else TIME_SZ                      # 고친 칸은 ⚙ 설정의 시각 크기
        lsz = LOCAL_TIME_SZ + (tsz - TIME_SZ)              # 로컬은 기본과의 차이만큼 같이
        if ed and ed_w:
            ed_lines.append((x0, y0, x1, y1, not b.get("top_dash"), not b.get("bottom_dash")))
        msz = tsz if h >= tsz * 1.05 else max(2.2, min(tsz, h * 0.9))
        dx, dy = (ed_dx, ed_dy) if ed else (TIME_DX, TIME_DY)     # 왼쪽 간격, 위아래 옮김
        p.bold = bool(b.get("bold")) and "bold" not in hide
        big_local = b.get("local") and h >= 9 and not ed     # 자동 로컬: 큰 회색 분 / 고친 로컬: 고친 칸 설정 그대로
        if big_local:
            p.text(x0 + 1.0 + (dx - TIME_DX), y0 + h / 2 + lsz * 0.37 + dy, mm, lsz, (0.3, 0.3, 0.3))   # 로컬은 크게
        else:
            base = y0 + msz * 0.8 + dy
            if dy > 0:      # 아래로 옮겨도 칸 밑선을 넘지 않게
                base = max(y0 + msz * 0.8, min(base, y1 - 0.3))
            p.text(x0 + dx, base, mm, msz)
        if b.get("show_end") and h >= msz * 2:      # 끝 시각: 칸 왼쪽 아래
            em = b["e"] + DAY_START_MIN
            et = cfg.get("end_label_format", "{mm}").format(hh=f"{em // 60:02d}", mm=f"{em % 60:02d}")
            p.text(x0 + dx, y1 - 0.7, et, msz, (0, 0, 0))      # 끝 시각은 시작 분처럼 검정
        icons = [ic for ic in icon_list(b) if ICON_KEY_OF.get(ic[0] if isinstance(ic, tuple) else ic) not in hide]
        isc = 1.0 if h >= ICON else max(0.5, h / ICON)
        iw = draw_icons(p, x1, y0, icons, scale=isc) if icons else 0
        if b["half"] and cfg.get("plan_label") and h >= 8:
            p.text(x0 + 0.4, y1 - 0.6, "1안" if b["half"] == "L" else "2안", 3.0, (0.75, 0.2, 0.3))
        # 글자: 칸이 낮으면 분 표시·아이콘 옆 공간에, 높으면 칸 전체 가운데에
        lw = p.width(mm, msz) + 1 + max(0.0, dx - TIME_DX)
        if big_local:       # 로컬: 왼쪽의 큰 분 표시와 오른쪽 아이콘을 피해서 가운데
            big = p.width(mm, lsz) + 2.5
            tx0, tx1, ty0, ty1 = x0 + big, x1 - max(iw, big) * (1 if iw else 0.3), y0, y1
        elif h < 11:
            tx0, tx1, ty0, ty1 = x0 + lw, x1 - iw - 0.5, y0, y1
        else:
            tx0, tx1, ty0, ty1 = x0 + 1.5, x1 - 1.5, y0 + min(ICON, h * 0.15), y1
            if msz * 0.95 + max(0.0, dy) > ty0 - y0 + 0.3:
                # 시각 숫자를 키우거나 내린 칸: 글자는 칸 가운데 그대로 두고,
                # 숫자 너비만큼 왼쪽·오른쪽을 똑같이 비워 겹치지 않게 (아래로 밀지 않음)
                tx0, tx1 = x0 + lw, x1 - lw
        items, total = layout_label(p, b["label"], tx1 - tx0, ty1 - ty0)
        cy = (ty0 + ty1) / 2 - total / 2
        fg = (None if "fg" in hide else hex_to_rgb(b.get("fg"))) or (0, 0, 0)
        hl = None if "hl" in hide else hex_to_rgb(b.get("hl"))
        for it in items:
            size = it[1]
            lh = size * 1.18
            if len(it) > 2:
                wa, wb = p.width(it[0], size), p.width(it[2], it[3])
                tw = wa + wb + 0.6
            else:
                tw = p.width(it[0], size)
            sx = (tx0 + tx1) / 2 - tw / 2
            if hl:      # 형광펜: 글자 줄 뒤에만 색
                p.rect(sx - 0.8, cy + lh * 0.08, sx + tw + 0.8, cy + lh * 0.98, fill=hl, stroke=None)
            if len(it) > 2:
                p.text(sx, cy + lh * 0.8, it[0], size, fg)
                p.text(sx + wa + 0.6, cy + lh * 0.8, it[2], it[3], fg)
            else:
                p.text(sx, cy + lh / 2 + size * 0.36, it[0], size, fg)
            cy += lh
        p.bold = False

    # 바깥 테두리
    # 위/아래 선 점선: 원래 실선을 흰 선으로 지우고 그 위에 점선 (옆 칸 테두리까지 덮도록 칸을 다 그린 뒤)
    dw, dc, dst, dgap = dash_cfg(cfg)          # ⚙ 설정: 점선 굵기·색·모양·빈 칸 색
    erase_w = max(0.55, cell_w + 0.25)

    def erase(xa, xb, y):
        """밑의 실선을 지움: 선 위쪽 반은 위 칸 색, 아래쪽 반은 아래 칸 색 (끄면 흰색)"""
        hw = erase_w / 2
        if not dgap:
            p.rect(xa, y - hw, xb, y + hw, fill=(1, 1, 1), stroke=None)
            return
        for cx0, cy0, cx1, cy1, cf in cell_fills:
            a, b_ = max(xa, cx0), min(xb, cx1)
            if b_ - a <= 0.05:
                continue
            if abs(cy1 - y) < 0.3:          # 위 칸 (가운데를 조금 겹쳐 이음매가 안 보이게)
                p.rect(a, y - hw, b_, y + 0.08, fill=cf, stroke=None)
            elif abs(cy0 - y) < 0.3:        # 아래 칸
                p.rect(a, y - 0.08, b_, y + hw, fill=cf, stroke=None)

    d_pat, d_cap = dash_pattern(dst, dw)
    for lx0, ly, lx1 in dash_lines:
        erase(lx0 + 0.2, lx1 - 0.2, ly)
        p.line(lx0, ly, lx1, ly, dc, dw, d_pat, d_cap)
    # 고친 칸 테두리: 옆 칸 선에 가리지 않게 마지막에.
    # 위·아래 변이 점선 자리면(이 칸 것이든 위·아래 칸 것이든) 실선으로 덮지 않고, 고친 칸 굵기·색의 점선으로
    def dashed_at(y, xa, xb):
        return any(abs(ly - y) < 0.3 and lx0 < xb - 0.5 and lx1 > xa + 0.5 for lx0, ly, lx1 in dash_lines)
    ed_dw = max(ed_w or 0, dw)          # 고친 칸 점선: 고친 칸 굵기와 점선 굵기 중 굵은 쪽, 색은 고친 칸 색
    ed_pat, ed_cap = dash_pattern(dst, ed_dw)
    for ex0, ey0, ex1, ey1, top_, bot_ in ed_lines:     # 1) 위·아래 변
        for y in (ey0, ey1):
            if dashed_at(y, ex0, ex1):
                erase(ex0 + ed_w / 2, ex1 - ed_w / 2, y)       # 밑의 선만 지우고 (빈 칸은 칸 색)
                p.line(ex0, y, ex1, y, ed_c, ed_dw, ed_pat, ed_cap)
            else:
                p.line(ex0 - ed_w / 2, y, ex1 + ed_w / 2, y, ed_c, ed_w)
    for ex0, ey0, ex1, ey1, top_, bot_ in ed_lines:     # 2) 왼쪽·오른쪽 변 (지운 자리 위에 다시)
        p.line(ex0, ey0, ex0, ey1, ed_c, ed_w)
        p.line(ex1, ey0, ex1, ey1, ed_c, ed_w)
    p.rect(29.0, TABLE_TOP, 567.7, TABLE_BOTTOM, fill=None, stroke=DARK, width=frame_w)
    p.finish()
    doc.set_metadata({"title": "주간편성표", "creator": f"{APP_NAME} {VERSION}"})
    if not out_path:
        return doc, geom
    save_pdf(doc, out_path)
    return None, geom


def page_png(doc_or_path, zoom=1.0):
    doc = pymupdf.open(doc_or_path) if isinstance(doc_or_path, str) else doc_or_path
    png = doc[0].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes("png")
    if isinstance(doc_or_path, str):
        doc.close()
    return png


def save_pdf(doc, out_path, subset=True):
    """PDF 저장. 글꼴을 줄여(subset) 파일을 작게 하되, 저장한 파일을 다시 그려 화면과 똑같은지 확인.
    다르면 글꼴을 줄이지 않고 다시 저장 -> 화면 = 미리보기 = 출력물 보장"""
    ref = page_png(doc)
    if subset:
        try:
            doc.subset_fonts()
        except Exception:
            subset = False
    doc.save(out_path, garbage=3, deflate=True)
    if subset and page_png(out_path) != ref:
        return False        # 글꼴 줄이기로 모양이 달라짐 -> 부른 쪽에서 다시 그려 전체 글꼴로 저장
    return True


SHOW_LABELS = {"": "모두", "tv": "TV만", "dmb": "DMB만"}
SHOW_KEYS = {v: k for k, v in SHOW_LABELS.items()}


def restore_under(rows, orig, i):
    """rows[i] 를 'DMB만'으로 바꿨을 때: 이 칸 때문에 잘려 나간 TV 편성을 엑셀(orig) 그대로 되살림.
    잘린 조각들은 다시 한 줄로 이어 붙임. 반환: 되살린 줄 수"""
    r = rows[i]
    d, plan = r["date"], r["plan"]
    visible = [x for k, x in enumerate(rows) if k != i and x["date"] == d and x["plan"] == plan
               and (x.get("show") or "") != "dmb"]
    n = 0
    for o in orig:
        if o["date"] != d or o["plan"] != plan or not (o["s"] < r["e"] and o["e"] > r["s"]):
            continue
        gaps = [(o["s"], o["e"])]
        for x in visible:
            nxt = []
            for a, b in gaps:
                if x["s"] < b and x["e"] > a:
                    if a < x["s"]:
                        nxt.append((a, x["s"]))
                    if b > x["e"]:
                        nxt.append((x["e"], b))
                else:
                    nxt.append((a, b))
            gaps = nxt
        for a, b in gaps:
            if b > a:
                rows.append(dict(o, s=a, e=b, **({"cut": True} if (a, b) != (o["s"], o["e"]) else {})))
                n += 1
    # 같은 프로그램의 잘린 조각이 이어져 있으면 한 줄로
    same = lambda x, y: all(x.get(k) == y.get(k) for k in ("date", "plan", "code", "name")) and x.get("cut") \
        and y.get("cut") and not x.get("local") and (x.get("show") or "") == (y.get("show") or "")
    changed = True
    while changed:
        changed = False
        for x in rows:
            y = next((y for y in rows if y is not x and same(x, y) and y["s"] == x["e"]), None)
            if y:
                x["e"] = y["e"]
                rows.remove(y)
                o = next((o for o in orig if o["date"] == x["date"] and o["plan"] == x["plan"]
                          and o.get("code") == x.get("code") and (o["s"], o["e"]) == (x["s"], x["e"])), None)
                if o:
                    x.pop("cut", None)
                changed = True
                break
    return n


def show_prefix(r):
    return {"tv": "[TV만] ", "dmb": "[DMB만] "}.get(r.get("show") or "", "")


def output_rows(rows, dmb=False):
    """출력용 줄 (원본은 그대로):
      TV  = 'DMB만' 칸은 뺌
      DMB = 'TV만' 칸은 빼고, 'DMB만' 칸이 덮는 시간은 다른 칸에서 잘라냄 (TV 편성은 건드리지 않음)
    줄마다 _src = rows 의 원래 번호 (화면에서 칸을 누르면 원래 줄을 고르도록)"""
    other = "tv" if dmb else "dmb"
    overlays = [r for r in rows if r.get("show") == "dmb"] if dmb else []
    out = []
    for i, r in enumerate(rows):
        sh = r.get("show") or ""
        if sh == other:
            continue
        pieces = [(r["s"], r["e"])]
        if sh != "dmb":
            for ov in overlays:
                if ov["date"] != r["date"] or ov["plan"] != r["plan"]:
                    continue
                nxt = []
                for a, b in pieces:
                    if ov["s"] < b and ov["e"] > a:
                        if a < ov["s"]:
                            nxt.append((a, ov["s"]))
                        if b > ov["e"]:
                            nxt.append((ov["e"], b))
                    else:
                        nxt.append((a, b))
                pieces = nxt
        for a, b in pieces:
            if b > a:
                x = dict(r, s=a, e=b, _src=i)
                if (a, b) != (r["s"], r["e"]):
                    x["cut"] = True
                out.append(x)
    return out


def compose(rows, prev_rows=None, cfg=None, printed_at=None, log=lambda s: None):
    """화면 보기·미리보기·PDF 저장이 모두 쓰는 단 하나의 그리기 함수"""
    cfg = cfg or load_config()
    orows = output_rows(rows, bool(cfg.get("_dmb")))
    dates, blocks = build_blocks(orows, cfg, prev_rows, log)
    for b in blocks:            # 칸 -> 원래 줄 번호
        b["_ris"] = sorted({orows[i]["_src"] for i in b["_ris"]})
    if not dates:
        raise ValueError("편성 줄이 없습니다.")
    doc, geom = draw_schedule(dates, blocks, cfg, None, printed_at)
    return doc, geom, dates, blocks


def generate(rows, out_path, prev_rows=None, log=print, printed_at=None, cfg_over=None):
    cfg = load_config()
    cfg.update(cfg_over or {})
    printed_at = printed_at or datetime.datetime.now()
    doc, _, dates, blocks = compose(rows, prev_rows, cfg, printed_at, log)
    if not save_pdf(doc, out_path):
        doc.close()
        doc, _, dates, blocks = compose(rows, prev_rows, cfg, printed_at)
        save_pdf(doc, out_path, subset=False)
        log("[알림] 글꼴 줄이기에서 모양이 달라져, 글꼴 전체를 넣어 저장했습니다 (파일이 조금 큼).")
    doc.close()
    n2 = sum(1 for r in rows if r["plan"] == "2안")
    log(f"[완료] {dates[0]}~{dates[-1]}  프로그램 {len(rows)}건(2안 {n2}건) -> 칸 {len(blocks)}개")
    log(f"[저장] {out_path}")
    if not find_font(cfg):
        log("[주의] 설정한 글꼴 파일을 찾지 못해 기본 한글 글꼴로 만들었습니다.")
    return out_path


def render_live(rows, prev_rows=None, zoom=1.4, cfg=None):
    """화면용: PDF 저장과 같은 compose() 로 그려서 PNG(바이트)와 칸 위치를 돌려줌"""
    doc, geom, dates, _ = compose(rows, prev_rows, cfg)
    png = page_png(doc, zoom)
    doc.close()
    return png, geom, dates


def make_pdf(xlsx, out_path=None, prev_xlsx=None, log=print):
    rows = read_schedule_xlsx(xlsx)
    apply_local_rules(rows, load_config(), log)
    over = {"_query_date": read_query_date(xlsx)}
    prev = read_schedule_xlsx(prev_xlsx) if prev_xlsx else None
    return generate(rows, out_path or os.path.splitext(xlsx)[0] + "_편성표.pdf", prev, log, cfg_over=over)


def add_shade_rule(match, color):
    cfg = load_config()
    rules = [r for r in cfg.get("shade_rules", []) if r.get("match") != match]
    if color:
        rules.append({"match": match, "color": color})
    cfg["shade_rules"] = rules
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


WEEKDAYS = "월화수목금토일"


def parse_days(spec):
    """'월-금', '토-일', '월,수,금', '매일' -> 요일 번호 집합(월=0)"""
    spec = str(spec).replace(" ", "")
    if spec in ("매일", "전체", ""):
        return set(range(7))
    out = set()
    for part in spec.split(","):
        if "-" in part or "~" in part:
            a, b = re.split("[-~]", part)[:2]
            if a in WEEKDAYS and b in WEEKDAYS:
                i, j = WEEKDAYS.index(a), WEEKDAYS.index(b)
                out.update(range(i, j + 1) if i <= j else list(range(i, 7)) + list(range(0, j + 1)))
        elif part in WEEKDAYS:
            out.add(WEEKDAYS.index(part))
    return out


# ---- 고친 칸 판단: 사람이 직접 고친 줄만 (로컬 때문에 잘린 줄·매주 자동 로컬은 뺌)
# 자동으로 바뀐 줄에는 그 순간의 내용 지문(auto_sig)을 남김. 나중에 사람이 고쳐 내용이 달라지면 '고친 칸'
_SIG_SKIP = ("edited", "cut", "name_edited", "auto_sig")


def row_sig(r):
    return json.dumps({k: v for k, v in r.items() if not k.startswith("_") and k not in _SIG_SKIP},
                      sort_keys=True, ensure_ascii=False, default=str)


def is_auto_row(r, cfg=None):
    """자동으로만 바뀐 줄인가 (로컬 때문에 잘린 줄, 매주 자동 로컬) - 사람이 그 뒤에 고쳤으면 False"""
    sig = r.get("auto_sig")
    if sig is not None:
        return row_sig(r) == sig
    # 예전 프로젝트 (지문 없음): 잘리기만 한 줄, 로컬 규칙 이름과 같은 로컬
    if r.get("cut") and not r.get("name_edited"):
        return True
    if r.get("local") and cfg is not None:
        names = {norm_key(x.get("name", "")) for x in cfg.get("local_rules", [])}
        return norm_key(r.get("name", "")) in names
    return False


def stamp_auto(r, was=None):
    """자동으로 바뀐 줄에 지문을 남김. was(바뀌기 전 줄)를 사람이 이미 고쳤었다면 남기지 않음(고친 칸 유지)"""
    if was is not None and was.get("edited") and not is_auto_row(was):
        r.pop("auto_sig", None)
        return r
    r["auto_sig"] = row_sig(r)
    return r


def is_marked(r, cfg):
    """⚙ 설정 '고친 칸'에 해당하는 줄인가"""
    if not r.get("edited"):
        return False
    if not is_auto_row(r, cfg):
        return True
    if r.get("local"):
        return bool(cfg.get("edited_mark_auto_local"))
    return bool(cfg.get("edited_mark_cut"))


def insert_local(rows, dates, name, s, e, shade="#d9d9d9", bold=True, code=None, extra=None, log=print,
                 auto=False):
    """고른 날짜마다 1안 편성의 s~e 를 로컬 프로그램으로 바꿈.
    겹치는 전국 프로그램은 앞/뒤 조각으로 잘라 남김(같은 코드라 요일끼리 계속 한 칸으로 묶임).
    auto=True: 매주 자동 로컬 (고친 칸으로 치지 않음). 잘린 조각은 늘 자동으로 바뀐 줄
    반환: 새로 넣은 로컬 줄 목록"""
    code = code or ("LOCAL-" + norm_key(name))
    added = []
    for d in dates:
        day = [r for r in rows if r["date"] == d and r["plan"] != "2안"]
        if any(r["date"] == d and r["plan"] == "2안" and r["s"] < e and r["e"] > s for r in rows):
            log(f"[주의] {day_label(d)} 이 시간에 2안이 있습니다. 로컬은 1안에만 넣습니다.")
        for r in day:
            if r["s"] < e and r["e"] > s:       # 겹침
                rows.remove(r)
                if r["s"] < s:
                    rows.append(stamp_auto(dict(r, e=s, edited=True, cut=True), r))
                if r["e"] > e:
                    rows.append(stamp_auto(dict(r, s=e, edited=True, cut=True), r))
                log(f"[로컬] {day_label(d)} '{r['name']}' {fmt_t(r['s'])}~{fmt_t(r['e'])} 중 "
                    f"{fmt_t(max(s, r['s']))}~{fmt_t(min(e, r['e']))} 을 로컬로")
        new = {"date": d, "code": code, "name": name, "s": s, "e": e, "uhd": False, "cc": True, "ad": False,
               "sign": False, "age": "", "rl": "", "rerun": False, "plan": "1안", "shade": shade or "",
               "bold": bool(bold), "local": True, "edited": True}
        new.update(extra or {})
        new.pop("auto_sig", None)
        if auto:
            stamp_auto(new)
        rows.append(new)
        added.append(new)
    return added


# 묶인 한 칸(월~금)을 고칠 때 요일마다 똑같이 맞추는 항목 (날짜는 빼고)
BLOCK_SYNC_KEYS = ("name", "s", "e", "plan", "age", "rl", "code", "shade", "uhd", "cc", "sign", "ad", "show_end",
                   "force_pink", "bold", "top_dash", "bottom_dash", "hl", "fg", "name_edited", "show")
STYLE_COPY_KEYS = ("shade", "hl", "fg", "bold", "top_dash", "bottom_dash")     # 꾸밈 복사/붙여넣기 항목
SHARE_CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"     # 헷갈리는 0·O·1·I 는 뺌


def share_code(rows, targets):
    """여러 칸에 같은 새 이름을 줄 때 옆 날짜와 한 칸으로 묶이도록 같은 코드(6자리)를 새로 만들어 붙임.
    이미 모두 같은 코드면 그대로. 엑셀의 원래 코드는 code_orig 에 남김(새 판 반영·되돌리기용).
    반환: 새 코드 (안 바꿨으면 None)"""
    import random
    if len(targets) < 2:
        return None
    codes = {x.get("code") or "" for x in targets}
    if len(codes) == 1 and "" not in codes:
        return None
    used = {r.get("code") for r in rows} | {r.get("code_orig") for r in rows}
    while True:
        code = "".join(random.choice(SHARE_CODE_CHARS) for _ in range(6))
        if code not in used:
            break
    for x in targets:
        if not x.get("local") and not x.get("code_orig") and x.get("code"):
            x["code_orig"] = x["code"]
        x["code"] = code
    return code


def clash(a, b):
    """두 줄이 같은 출력(TV/DMB)에서 시간이 겹치는지. 'DMB만' 칸은 다른 칸을 덮는 것이라 'DMB만'끼리만 겹침"""
    if a["date"] != b["date"] or a["plan"] != b["plan"] or not (a["s"] < b["e"] and a["e"] > b["s"]):
        return False
    sa, sb = a.get("show") or "", b.get("show") or ""
    if "dmb" in (sa, sb):
        return sa == sb == "dmb"
    return True


def find_overlaps(rows, i):
    """rows[i] 와 같은 날·같은 안에서 시간이 겹치는 줄 번호들"""
    r = rows[i]
    return [k for k, x in enumerate(rows) if k != i and clash(r, x)]


def overlap_plan(rows, i):
    """rows[i] 시간에 맞춰 겹치는 칸을 어떻게 할지: [(줄 번호, 할 일, 설명)]
      할 일: del(다 덮여서 지움) / cut_end(뒤를 잘라 앞만 남김) / cut_start(앞을 잘라 뒤만 남김) / split(앞뒤 둘로)"""
    r = rows[i]
    out = []
    for k in find_overlaps(rows, i):
        x = rows[k]
        nm = f"{fmt_t(x['s'])}~{fmt_t(x['e'])} {x['name']}".replace("\n", " ")
        if x["s"] >= r["s"] and x["e"] <= r["e"]:
            out.append((k, "del", f"{nm} → 지움"))
        elif x["s"] < r["s"] and x["e"] > r["e"]:
            out.append((k, "split", f"{nm} → {fmt_t(x['s'])}~{fmt_t(r['s'])}, {fmt_t(r['e'])}~{fmt_t(x['e'])} 둘로"))
        elif x["s"] < r["s"]:
            out.append((k, "cut_end", f"{nm} → {fmt_t(x['s'])}~{fmt_t(r['s'])} 로 줄임"))
        else:
            out.append((k, "cut_start", f"{nm} → {fmt_t(r['e'])}~{fmt_t(x['e'])} 로 줄임"))
    return out


def resolve_overlaps(rows, i, plan=None):
    """overlap_plan 대로 겹치는 칸 정리. 반환: 정리 뒤 rows[i] 의 새 줄 번호, 지운 줄 수"""
    r = rows[i]
    plan = overlap_plan(rows, i) if plan is None else plan
    gone = set()
    for k, act, _ in plan:
        x = rows[k]
        if act == "del":
            gone.add(k)
        elif act == "cut_end":
            was = dict(x)
            stamp_auto(x.update(e=r["s"], edited=True) or x, was)
        elif act == "cut_start":
            was = dict(x)
            stamp_auto(x.update(s=r["e"], edited=True) or x, was)
        else:
            was = dict(x)
            rows.append(stamp_auto(dict(x, s=r["e"], edited=True, cut=True), was))
            stamp_auto(x.update(e=r["s"], edited=True, cut=True) or x, was)
    for k in sorted(gone, reverse=True):
        rows.pop(k)
    return rows.index(r), len(gone)


# ---------------------------------------------------------------- 새 판 반영 (같은 주 엑셀 수정판)
STYLE_KEYS = ("shade", "hl", "fg", "bold", "show_end", "force_pink", "top_dash", "bottom_dash")
ICON_KEYS = ("uhd", "cc", "sign", "ad", "age", "rl")


def _exact_key(r):
    return (r["date"], r["plan"], r["s"], r["e"], r.get("code", ""), norm_key(r["name"]))


def _overlap(a0, a1, b0, b1):
    return max(0, min(a1, b1) - max(a0, b0))


def merge_new_version(cur_rows, old_orig, new_rows):
    """같은 주의 새 판 엑셀(new_rows)에, 지금 작업(cur_rows: 예전 판 + 고친 것)을 옮겨 얹는다.
      old_orig: 예전 판 엑셀 그대로의 줄 (프로젝트의 orig)
    반환: (합친 줄, 보고서)
      줄마다 _ver = same(그대로) / changed(시간·이름이 바뀜) / new(새로 생김), _check = 확인 필요
      보고서: same/changed/new/removed/kept_edits/locals/checks 목록"""
    import copy
    new_rows = copy.deepcopy(new_rows)
    # 묶으려고 새로 붙인 코드(share_code)는 엑셀 원래 코드로 짝을 찾고, 끝에 다시 붙임
    cur_rows = [dict(c, code=c["code_orig"], _share=c["code"]) if c.get("code_orig") else c for c in cur_rows]
    old_exact = {_exact_key(r) for r in old_orig}
    new_exact = {_exact_key(r) for r in new_rows}
    # 1) 새 판 줄마다: 예전 판의 어느 줄과 같은지 (똑같은 줄이 있는 예전 줄은 짝 후보에서 뺌)
    used_old = {i for i, o in enumerate(old_orig) if _exact_key(o) in new_exact}
    status = []
    for n in new_rows:
        if _exact_key(n) in old_exact:
            status.append("same")
            continue
        cands = [i for i, o in enumerate(old_orig) if i not in used_old and o["date"] == n["date"]
                 and o.get("code") and o.get("code") == n.get("code")]
        if not cands:
            cands = [i for i, o in enumerate(old_orig) if i not in used_old and o["date"] == n["date"]
                     and norm_key(split_suffix(o["name"])[0]) == norm_key(split_suffix(n["name"])[0])]
        if cands:
            i = min(cands, key=lambda i: (old_orig[i]["plan"] != n["plan"], abs(old_orig[i]["s"] - n["s"])))
            used_old.add(i)
            o = old_orig[i]
            n["_was"] = f"{fmt_t(o['s'])}~{fmt_t(o['e'])} {o['name']}" + (" (2안)" if o["plan"] == "2안" else "")
            status.append("changed")
        else:
            status.append("new")
    matched_exact = {_exact_key(n) for n, st in zip(new_rows, status) if st == "same"}
    paired = {i for i in used_old if _exact_key(old_orig[i]) not in new_exact}
    removed = [o for i, o in enumerate(old_orig) if _exact_key(o) not in matched_exact and i not in paired]

    # 2) 지금 작업의 줄을 나눔: 로컬 / 직접 추가 / 엑셀에서 온 줄(잘린 조각 포함)
    old_codes = {(o["date"], o.get("code")) for o in old_orig if o.get("code")}
    locals_ = [c for c in cur_rows if c.get("local")]
    added = [c for c in cur_rows if not c.get("local") and (c["date"], c.get("code")) not in old_codes]
    based = [c for c in cur_rows if not c.get("local") and (c["date"], c.get("code")) in old_codes]

    def orig_of(c):
        cs = [o for o in old_orig if o["date"] == c["date"] and o.get("code") == c.get("code")]
        return max(cs, key=lambda o: _overlap(o["s"], o["e"], c["s"], c["e"]) - abs(o["s"] - c["s"]) / 1000,
                   default=None)

    # 3) 고친 것 옮기기: 예전 판 엑셀(o)과 지금 작업(c)이 다른 것만 = 사용자가 고친 것
    kept, dropped, time_checks = [], [], []
    for n, st in zip(new_rows, status):
        n["_ver"] = st
        if st == "new":
            continue
        cs = [c for c in based if c["date"] == n["date"] and c.get("code") == n.get("code")]
        if not cs:
            # 예전 판에서 지운 줄이면 (로컬로 덮인 것 말고) 새 판에서도 지움
            covered = any(l["date"] == n["date"] and _overlap(l["s"], l["e"], n["s"], n["e"]) > 0 for l in locals_)
            if st == "same" and not covered:
                n["_drop"] = True
                dropped.append(n)
            continue
        c = max(cs, key=lambda c: (c["plan"] == n["plan"], c["s"] == n["s"],
                                   _overlap(c["s"], c["e"], n["s"], n["e"]), -abs(c["s"] - n["s"])))
        o = orig_of(c) or {}
        moved = []
        for k in STYLE_KEYS + ICON_KEYS:
            if k == "force_pink" and c["plan"] != n["plan"]:      # 2안 분홍 끄기는 1안으로 옮기지 않음
                continue
            nz = (lambda v: v) if k == "force_pink" else (lambda v: v or None)   # ""·False·없음은 같은 것
            if k in c and nz(c.get(k)) != nz(o.get(k)) and nz(n.get(k)) != nz(c.get(k)):
                n[k] = c.get(k)
                moved.append(k)
        if norm_key(c["name"]) != norm_key(o.get("name", c["name"])):
            n["name"] = c["name"]
            n["name_edited"] = True
            moved.append("이름")
        if o and not c.get("cut") and (c["s"], c["e"]) != (o["s"], o["e"]):     # 직접 바꾼 시간
            if st == "same":
                n["s"], n["e"] = c["s"], c["e"]
                moved.append(f"시간 {fmt_t(c['s'])}~{fmt_t(c['e'])}")
            else:
                n["_check"] = True
                n["_note"] = (f"예전 판에서 시간을 {fmt_t(c['s'])}~{fmt_t(c['e'])}(으)로 바꿨던 칸인데 "
                              f"새 판에서 편성이 바뀜 - 시간 확인")
                time_checks.append(n)
        if c.get("_share"):
            n["code_orig"], n["code"] = n.get("code"), c["_share"]
        if moved:
            n["edited"] = True
            kept.append((n, moved))
    new_rows = [n for n in new_rows if not n.get("_drop")]

    # 4) 바뀐 시간대 (로컬·직접 추가 줄이 여기에 걸리면 확인 필요)
    spans = {}
    for n in new_rows:
        if n["_ver"] != "same":
            spans.setdefault(n["date"], []).append((n["s"], n["e"]))
    for o in removed:
        spans.setdefault(o["date"], []).append((o["s"], o["e"]))
    hits = lambda d, a, b: any(_overlap(a, b, x, y) > 0 for x, y in spans.get(d, []))

    moved_rows = [(n["date"], n["s"], n["e"], n["name"]) for n in new_rows if n["_ver"] in ("changed", "new")]

    def note(d, a, b):
        ns = [f"{fmt_t(x)}~{fmt_t(y)} {nm}" for dd, x, y, nm in moved_rows if dd == d and _overlap(a, b, x, y) > 0]
        gone = [f"{fmt_t(o['s'])}~{fmt_t(o['e'])} {o['name']}" for o in removed
                if o["date"] == d and o["plan"] != "2안" and _overlap(a, b, o["s"], o["e"]) > 0]
        t = "새 판에서 바뀐 편성과 겹침: " + ", ".join(ns) if ns else "겹치던 편성이 없어짐"
        return t + (" / 없어진 것: " + ", ".join(gone) if gone else "")

    merged = new_rows
    checks = []
    for c in sorted(locals_, key=lambda c: (c["date"], c["s"])):
        extra = {k: c.get(k) for k in STYLE_KEYS + ICON_KEYS + ("name_edited",) if k in c}
        rows_added = insert_local(merged, [c["date"]], c["name"], c["s"], c["e"], c.get("shade", ""),
                                  c.get("bold", False), c.get("code"), extra=extra, log=lambda s: None,
                                  auto=c.get("auto_sig") is not None and is_auto_row(c))
        for r in rows_added:
            r["_ver"] = "local"
            if hits(c["date"], c["s"], c["e"]):
                r["_check"], r["_note"] = True, note(c["date"], c["s"], c["e"])
                checks.append(r)
    for c in added:
        r = copy.deepcopy(c)
        if r.get("_share"):
            r["code"] = r.pop("_share")
            r.pop("code_orig", None)
        r["_ver"] = "added"
        if hits(c["date"], c["s"], c["e"]):
            r["_check"], r["_note"] = True, note(c["date"], c["s"], c["e"])
            checks.append(r)
        merged.append(r)
    for r in merged:        # 잘린 조각은 원래 줄의 표시를 따라감
        r.setdefault("_ver", "same")
    checks = time_checks + checks
    # 마지막 점검: 같은 날·같은 안에서 시간이 겹치는 칸 (직접 추가한 줄과 새 판 줄 등)
    for d in sorted({r["date"] for r in merged}):
        for plan in ("1안", "2안"):
            day = sorted([r for r in merged if r["date"] == d and r["plan"] == plan], key=lambda r: (r["s"], r["e"]))
            for a, b in ((a, b) for n, a in enumerate(day) for b in day[n + 1:]):
                if clash(a, b):
                    for x, y in ((a, b), (b, a)):
                        if not x.get("_check"):
                            x["_check"], x["_ovl"] = True, True
                            x["_note"] = f"다른 칸과 시간이 겹침: {fmt_t(y['s'])}~{fmt_t(y['e'])} {y['name']}"
                            checks.append(x)
    report = {"same": [n for n in new_rows if n.get("_ver") == "same" and not n.get("local")],
              "changed": [n for n in merged if n.get("_ver") == "changed"],
              "new": [n for n in merged if n.get("_ver") == "new"],
              "removed": removed, "kept": kept, "locals": [r for r in merged if r.get("_ver") == "local"],
              "added": [r for r in merged if r.get("_ver") == "added"], "checks": checks, "dropped": dropped}
    return merged, report


def merge_keep_past(cur_rows, old_orig, new_rows, keep_before=None):
    """merge_new_version + 지난 날(keep_before 'YYYYMMDD' 앞 날짜)은 예전 판 작업 그대로.
    반환: (합친 줄, 보고서, 새 orig)  보고서['past'] = 지난 날이라 반영하지 않은 변경"""
    import copy
    merged, rep = merge_new_version(cur_rows, old_orig, new_rows)
    if not keep_before:
        rep["past"] = []
        return merged, rep, copy.deepcopy(new_rows)
    past = lambda r: r["date"] < keep_before
    kinds = (("new", "새로 생김"), ("changed", "바뀜"), ("removed", "없어짐"), ("checks", "확인"),
             ("dropped", "지운 줄"))
    rep["past"] = [(t, r) for k, t in kinds for r in rep[k] if past(r)]
    for k in ("same", "changed", "new", "removed", "locals", "added", "checks", "dropped"):
        rep[k] = [r for r in rep[k] if not past(r)]
    rep["kept"] = [(n, mv) for n, mv in rep["kept"] if not past(n)]
    keep = [copy.deepcopy(r) for r in cur_rows if past(r)]
    for r in keep:
        for k in ("_ver", "_check", "_note", "_was", "_ovl"):
            r.pop(k, None)
    merged = keep + [r for r in merged if not past(r)]
    orig = [copy.deepcopy(r) for r in old_orig if past(r)] + [copy.deepcopy(r) for r in new_rows if not past(r)]
    return merged, rep, orig


def apply_local_rules(rows, cfg, log=print):
    """설정의 local_rules 를 이번 주 편성에 적용 (엑셀 불러올 때)"""
    dates = sorted({r["date"] for r in rows})
    n = 0
    for rule in cfg.get("local_rules", []):
        try:
            s = hhmm_to_min(rule["start"]) - DAY_START_MIN
            e = hhmm_to_min(rule["end"]) - DAY_START_MIN
        except (KeyError, ValueError):
            log(f"[주의] 로컬 규칙 형식 오류: {rule}")
            continue
        wd = parse_days(rule.get("days", "월-금"))
        target = [d for d in dates if datetime.datetime.strptime(d, "%Y%m%d").weekday() in wd]
        if not target:
            continue
        insert_local(rows, target, rule.get("name", "로컬"), s, e, rule.get("shade", "#d9d9d9"),
                     rule.get("bold", True), rule.get("code"),
                     {k: rule[k] for k in ("cc", "age", "rl", "uhd", "ad", "sign") if k in rule}, log=lambda x: None,
                     auto=True)
        n += 1
        log(f"[로컬 규칙] '{rule.get('name')}' {rule.get('days', '월-금')} {rule['start']}~{rule['end']} 넣음")
    return n


def add_local_rule(rule):
    cfg = load_config()
    rules = [r for r in cfg.get("local_rules", []) if not (r.get("name") == rule["name"] and
                                                          r.get("start") == rule["start"])]
    rules.append(rule)
    cfg["local_rules"] = rules
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- 프로젝트 파일 (.tvproj)
# 프로젝트는 늘 데이터 폴더의 주 폴더에: C:\EDIT\TV\데이터\엑셀편성표\프로젝트\2026-11-16 (11월 3주)\
# (PDF 는 고른 곳에). 오래된 주는 C:\EDIT\TV\백업\엑셀편성표 로 압축해 옮김
PROJ_DIR = os.path.join(DATA_DIR, "프로젝트")
BACKUP_DIR = os.path.join(os.environ["RADIO_HOME"], "백업", "엑셀편성표") if os.environ.get("RADIO_HOME") \
    else os.path.join(DATA_DIR, "백업")
RECENT_TV_PATH = os.path.join(DATA_DIR, "recent_projects.json")       # DMB 편집기가 TV 최근 목록을 볼 때
TV_PROJ_EXT = ".tvproj"
DMB_PROJ_EXT = ".dmbproj"
PROJ_EXT = DMB_PROJ_EXT if IS_DMB else TV_PROJ_EXT
RECENT_PATH = os.path.join(DATA_DIR, "recent_dmb_projects.json" if IS_DMB else "recent_projects.json")
PROJ_FORMATS = ("tv_schedule_project", "dmb_schedule_project")
PROJ_CFG_KEYS = ("title_text", "title_size", "title_bold", "stamp_mode", "stamp_text")


def save_project(path, data):
    """편성 줄(고친 내용 포함)·엑셀 원래 내용·제목/날짜 설정을 한 파일에 저장 -> 엑셀 없이도 다시 열어 편집"""
    data = dict(data, format="dmb_schedule_project" if path.lower().endswith(DMB_PROJ_EXT) else "tv_schedule_project",
                version=VERSION, app=APP_NAME,
                saved=datetime.datetime.now().isoformat(timespec="seconds"))
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, path)
    add_recent(path)
    return path


def load_project(path, recent=True):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if data.get("format") not in PROJ_FORMATS or not isinstance(data.get("rows"), list):
        raise ValueError("편성표 프로젝트 파일이 아닙니다.")
    if recent:
        add_recent(path)
    return data


# ---------------------------------------------------------------- DMB 편집기: TV 프로젝트 가져오기
MARK_KEYS = ("_ver", "_check", "_note", "_was", "_ovl")


def tv_key(r):
    """TV 편성 줄을 알아보는 열쇠 (DMB 줄이 어느 TV 줄에서 왔는지)"""
    return f"{r['date']}|{r['plan']}|{r.get('code', '')}|{r['s']}|{r['e']}|{norm_key(r['name'])}"


def import_tv_rows(tv_rows):
    """TV 프로젝트 줄 -> DMB 작업 줄 (표시는 지우고, 어느 TV 줄에서 왔는지 _tvk 로 기억)"""
    import copy
    out = []
    for r in tv_rows:
        x = {k: v for k, v in copy.deepcopy(r).items() if k not in MARK_KEYS}
        x["_tvk"] = tv_key(r)
        out.append(x)
    return out


def refresh_from_tv(dmb_rows, old_tv, new_tv):
    """DMB 작업에 바뀐 TV 프로젝트를 다시 반영.
      TV 에서 그대로인 줄  -> DMB 에서 고친 그대로 (지운 줄은 계속 지움)
      TV 에서 바뀌거나 새로 생긴 줄 -> 새 TV 내용 (DMB 에서 넣은 음영·이름 등 꾸밈은 옮김), 주황 표시
      DMB만 칸 -> 그대로 (그 밑의 TV 편성이 바뀌었으면 빨간 표시)
    반환: (합친 줄, 보고서)"""
    import copy
    old_keys = {tv_key(o): o for o in old_tv}
    by_k = {}
    for r in dmb_rows:
        if r.get("_tvk"):
            by_k.setdefault(r["_tvk"], []).append(r)
    old_codes = {(o["date"], o.get("code")) for o in old_tv}
    merged, rep = [], {"changed": [], "new": [], "removed": [], "checks": [], "kept": 0}
    new_keys = set()
    for n in new_tv:
        k = tv_key(n)
        new_keys.add(k)
        if k in old_keys:
            for r in by_k.get(k, []):       # 그대로 (DMB 에서 지웠으면 없음)
                x = {kk: v for kk, v in copy.deepcopy(r).items() if kk not in MARK_KEYS}
                merged.append(x)
                rep["kept"] += 1
            continue
        x = {kk: v for kk, v in copy.deepcopy(n).items() if kk not in MARK_KEYS}
        x["_tvk"] = k
        # 같은 날·같은 코드의 예전 DMB 줄에서 꾸밈 옮기기
        olds = [r for kk, rs in by_k.items() for r in rs if kk.split("|")[0] == n["date"]
                and kk.split("|")[2] == (n.get("code") or "")]
        if olds:
            r = min(olds, key=lambda r: abs(r["s"] - n["s"]))
            o = old_keys.get(r["_tvk"], {})
            for kk in STYLE_KEYS:
                if (r.get(kk) or None) != (o.get(kk) or None):
                    x[kk] = r.get(kk)
        x["_ver"] = "changed" if (n["date"], n.get("code")) in old_codes else "new"
        rep[x["_ver"]].append(x)
        merged.append(x)
    new_codes = {(n["date"], n.get("code")) for n in new_tv}
    rep["removed"] = [o for k, o in old_keys.items() if k not in new_keys and (o["date"], o.get("code")) not in new_codes]
    # DMB 에서 직접 넣은 줄 (DMB만 칸 등)
    spans = {}
    for x in rep["changed"] + rep["new"] + rep["removed"]:
        spans.setdefault(x["date"], []).append((x["s"], x["e"]))
    for r in dmb_rows:
        if r.get("_tvk"):
            continue
        x = {kk: v for kk, v in copy.deepcopy(r).items() if kk not in MARK_KEYS}
        if any(a < x["e"] and b > x["s"] for a, b in spans.get(x["date"], [])):
            x["_check"] = True
            x["_note"] = "밑에 있는 TV 편성이 바뀌었습니다 - 시간 확인"
            rep["checks"].append(x)
        merged.append(x)
    return merged, rep


def week_name(dates):
    """편성 날짜들 -> 주 폴더 이름 '2026-11-16 (11월 3주)' (월요일 기준)"""
    if not dates:
        return "날짜 없음"
    d = datetime.datetime.strptime(min(dates), "%Y%m%d")
    return f"{d:%Y-%m-%d} ({d.month}월 {(d.day - 1) // 7 + 1}주)"


def week_label(dates):
    if not dates:
        return ""
    d = datetime.datetime.strptime(min(dates), "%Y%m%d")
    return f"{d.month}월 {(d.day - 1) // 7 + 1}주"


def week_dir(rows):
    return os.path.join(PROJ_DIR, week_name(sorted({r["date"] for r in rows})))


def tidy_projects(cfg=None, log=print):
    """프로젝트 정리 (켤 때·저장할 때):
      ① 데이터 폴더 바로 아래·프로젝트 폴더 바로 아래 흩어진 프로젝트 -> 주 폴더로
      ② 주 폴더마다 TV·DMB 각각 최근 N개만 (PDF 로 낸 프로젝트 kept 는 늘 남김)
      ③ 며칠 지난 주 폴더 -> 백업 폴더에 zip 으로 옮김"""
    import glob
    import shutil
    import zipfile
    cfg = cfg or load_config()
    keep_n = int(cfg.get("proj_keep_per_week", 10) or 10)
    days = int(cfg.get("proj_archive_days", 60) or 60)
    os.makedirs(PROJ_DIR, exist_ok=True)
    moved = removed = archived = 0
    for f in glob.glob(os.path.join(DATA_DIR, "*.tvproj")) + glob.glob(os.path.join(DATA_DIR, "*.dmbproj")) + \
            glob.glob(os.path.join(PROJ_DIR, "*.tvproj")) + glob.glob(os.path.join(PROJ_DIR, "*.dmbproj")):
        try:
            with open(f, "r", encoding="utf-8") as fh:
                d = json.load(fh)
            wd = os.path.join(PROJ_DIR, week_name(sorted({r["date"] for r in d.get("rows", [])})))
            os.makedirs(wd, exist_ok=True)
            dst = os.path.join(wd, os.path.basename(f))
            if not os.path.exists(dst):
                shutil.move(f, dst)
                moved += 1
                _fix_recent(f, dst)
        except Exception:
            pass
    today = datetime.date.today()
    for wd in sorted(glob.glob(os.path.join(PROJ_DIR, "*"))):
        if not os.path.isdir(wd):
            continue
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", os.path.basename(wd))
        if m and (today - datetime.date(*map(int, m.groups()))).days > days:
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True)
                zp = os.path.join(BACKUP_DIR, os.path.basename(wd) + ".zip")
                with zipfile.ZipFile(zp, "a", zipfile.ZIP_DEFLATED) as z:
                    have = set(z.namelist())
                    for f in os.listdir(wd):
                        if f not in have:
                            z.write(os.path.join(wd, f), f)
                shutil.rmtree(wd)
                archived += 1
            except Exception as e:
                log(f"[프로젝트 정리] 백업으로 옮기지 못함: {os.path.basename(wd)} ({e})")
            continue
        for ext in ("*.tvproj", "*.dmbproj"):
            files = sorted(glob.glob(os.path.join(wd, ext)), key=os.path.getmtime, reverse=True)
            normal = []
            for f in files:
                try:
                    with open(f, "r", encoding="utf-8") as fh:
                        if json.load(fh).get("kept"):
                            continue
                except Exception:
                    continue
                normal.append(f)
            for f in normal[keep_n:]:
                try:
                    os.remove(f)
                    removed += 1
                except Exception:
                    pass
    if moved or removed or archived:
        log(f"[프로젝트 정리] 주 폴더로 옮김 {moved} · 오래된 저장 지움 {removed} · "
            f"{days}일 지난 주 {archived}개 -> 백업({BACKUP_DIR})")
    return moved, removed, archived


def _fix_recent(old, new):
    for rp in {RECENT_PATH, RECENT_TV_PATH}:
        try:
            with open(rp, "r", encoding="utf-8") as f:
                lst = json.load(f)
            lst = [new if os.path.abspath(p) == os.path.abspath(old) else p for p in lst]
            with open(rp, "w", encoding="utf-8") as f:
                json.dump(lst, f, ensure_ascii=False, indent=1)
        except Exception:
            pass


def tv_project_choices(limit=12):
    """DMB 편집기: 고를 TV 프로젝트 (TV 편집기의 최근 목록 + 주 폴더의 새 것, 새 것이 위)"""
    import glob
    lst = []
    try:
        with open(RECENT_TV_PATH, "r", encoding="utf-8") as f:
            lst = [p for p in json.load(f) if os.path.isfile(p) and p.lower().endswith(TV_PROJ_EXT)]
    except Exception:
        pass
    lst += glob.glob(os.path.join(PROJ_DIR, "*", "*" + TV_PROJ_EXT))
    seen, out = set(), []
    for p in sorted(lst, key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True):
        a = os.path.abspath(p)
        if a not in seen:
            seen.add(a)
            out.append(p)
    return out[:limit]


def latest_week_dir():
    import glob
    ds = sorted(d for d in glob.glob(os.path.join(PROJ_DIR, "*")) if os.path.isdir(d))
    return ds[-1] if ds else PROJ_DIR


def recent_projects():
    try:
        with open(RECENT_PATH, "r", encoding="utf-8") as f:
            lst = json.load(f)
        return [p for p in lst if os.path.isfile(p)]
    except Exception:
        return []


def add_recent(path):
    path = os.path.abspath(path)
    lst = [p for p in recent_projects() if os.path.abspath(p) != path]
    lst.insert(0, path)
    try:
        with open(RECENT_PATH, "w", encoding="utf-8") as f:
            json.dump(lst[:10], f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def project_info(path):
    """최근 목록에 보일 글: 파일명 · 주간 · 저장 시각"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        dates = sorted({r["date"] for r in d.get("rows", [])})
        wk = f"{day_label(dates[0])}~{day_label(dates[-1])}" if dates else ""
        tag = f"[{week_label(dates)}] " if dates else ""
        pdf = " · PDF" if d.get("kept") else ""
        return f"{tag}{os.path.basename(path)}   ({wk} · 저장 {d.get('saved', '')[:16].replace('T', ' ')}{pdf})"
    except Exception:
        return os.path.basename(path)


# ---------------------------------------------------------------- 작업 파일 (고친 내용 저장)
def work_path(xlsx):
    return os.path.splitext(xlsx)[0] + "_작업.json"


def save_work(xlsx, rows, edits):
    with open(work_path(xlsx), "w", encoding="utf-8") as f:
        json.dump({"source": os.path.basename(xlsx), "saved": datetime.datetime.now().isoformat(timespec="seconds"),
                   "edits": edits, "rows": rows}, f, ensure_ascii=False, indent=0)
    return work_path(xlsx)


def load_work(xlsx):
    with open(work_path(xlsx), "r", encoding="utf-8") as f:
        w = json.load(f)
    return w["rows"], w.get("edits", 0), w.get("saved", "")


# ---------------------------------------------------------------- 창(GUI)
def day_label(d):
    dt = datetime.datetime.strptime(d, "%Y%m%d")
    return f"{dt.month}/{dt.day}({'월화수목금토일'[dt.weekday()]})"


def icon_text(r):
    return " ".join(i[1] if isinstance(i, tuple) else i for i in icon_list(r))


class App:
    COLS = (("date", "날짜", 80), ("s", "시작", 55), ("e", "끝", 55), ("name", "프로그램명", 330),
            ("plan", "안", 40), ("icons", "아이콘", 120), ("bg", "배경", 70), ("end", "끝표시", 50), ("edited", "고침", 45))

    def __init__(self):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk = tk, ttk
        self.root = root = tk.Tk()
        root.title(f"{APP_NAME} {VERSION}")
        root.geometry("1400x900")
        self.xlsx = None
        self.rows, self.orig = [], []
        self.t0 = None
        self.sel = None             # 고른 편성 줄 번호
        self.deleted = 0
        _z = load_config().get("view_zoom")
        self.zoom = float(_z) if isinstance(_z, (int, float)) and not isinstance(_z, bool) else 1.4
        self._redraw_id = None
        self.geom, self.dates = [], []
        self.diff_keys = set()
        self._loading = self._quiet = False
        self._prev_path = self._prev_rows = None
        self.cfg = load_config()
        self.img = None
        self.pink_state = {}        # 줄 번호 -> 지금 편성표에서 분홍인지
        self.undo_stack, self.redo_stack = [], []
        self.proj_path = None       # 지금 프로젝트 파일
        self.proj_dir = ""          # 새 판 반영 뒤 저장할 폴더 (예전 프로젝트가 있던 곳)
        self.history = []           # 새 판 반영 기록
        self.ver_win = None
        self.multi = []             # 여러 칸 선택 (Ctrl+클릭 / Shift+클릭)
        self.style_clip = None      # 꾸밈 복사: {shade, hl, fg, bold, top_dash, bottom_dash} + 원본 이름
        self.empty_slot = None      # 편성 없는 빈 시간을 누름: {date, s, e, plan, rect} - 이름을 넣으면 새 줄
        self._multi_warned = False
        self._last_tag, self._last_t = None, 0.0
        self.saved = True
        self.day = tk.StringVar(value="전체")
        self.v_prev = tk.StringVar()
        self.v_x = tk.StringVar()

        # 맨 위 줄 (첫 화면에도 보임): [엑셀 열기] [지난주 엑셀] [설정] 상태
        top = tk.Frame(root)
        top.pack(fill=tk.X, padx=8, pady=6)
        self.topbar = top
        if IS_DMB:      # DMB 편집기: TV 프로젝트(.tvproj)로 시작, DMB 작업은 .dmbproj 로 따로
            mb_tv = tk.Menubutton(top, text="TV 프로젝트 불러오기 ▼", relief="raised", bg="#e1bee7")
            mb_tv.pack(side=tk.LEFT)
            m_tv = tk.Menu(mb_tv, tearoff=0, postcommand=lambda: self.fill_tv_menu(m_tv, self.import_tv))
            mb_tv.config(menu=m_tv)
            tk.Button(top, text="DMB 프로젝트 열기", command=lambda: self.open_project()).pack(side=tk.LEFT,
                                                                                         padx=(6, 0))
        else:
            tk.Button(top, text="엑셀 열기", width=10, command=lambda: (self.v_x.set(""), self.open_xlsx())).pack(
                side=tk.LEFT)
            tk.Button(top, text="프로젝트 열기", command=lambda: self.open_project()).pack(side=tk.LEFT, padx=(6, 0))
        self.recent_mb = tk.Menubutton(top, text="최근 ▼", relief="raised")
        self.recent_mb.pack(side=tk.LEFT, padx=(2, 0))
        self.recent_menu = tk.Menu(self.recent_mb, tearoff=0)
        self.recent_mb.config(menu=self.recent_menu)
        tk.Button(top, text="⚙ 설정", command=self.toggle_head).pack(side=tk.LEFT, padx=6)
        self.v_view_on = tk.BooleanVar(value=tv_view_config()["on"])     # 📺 서버로 보내기 (⚙ 설정 창에서)
        self.status = tk.Label(top, text=f"[{VERSION}] 상단의 '{'TV 프로젝트 불러오기' if IS_DMB else '엑셀 열기'}'를 "
                                         f"눌러 시작하세요.")
        self.status.pack(side=tk.LEFT, padx=16)

        # 위쪽 실행 버튼 줄
        act = tk.Frame(root)
        act.pack(fill=tk.X, padx=8, pady=(0, 4))
        self.act = act
        abtn = lambda text, cmd, color=None: tk.Button(act, text=text, command=cmd,
                                                       bg=color or "#e0e0e0").pack(side=tk.LEFT, padx=2)
        self.btn_undo = tk.Button(act, text="↶ 되돌리기", command=self.undo, state=tk.DISABLED)
        self.btn_undo.pack(side=tk.LEFT, padx=2)
        self.btn_redo = tk.Button(act, text="↷ 다시 하기", command=self.redo, state=tk.DISABLED)
        self.btn_redo.pack(side=tk.LEFT, padx=(2, 10))
        abtn("분홍 모두 지우기", self.clear_all_pink, "#ffd6d6")
        abtn("PDF 저장", self.make, "#4CAF50")
        root.bind("<Escape>", lambda e: self.clear_selection())
        root.bind_all("<Control-s>", lambda e: (self.save(), "break")[1])
        root.bind_all("<Control-o>", lambda e: (self.open_project(), "break")[1])
        for k in ("<Control-z>", "<Control-Z>"):
            root.bind_all(k, lambda e: (self.undo(), "break")[1])
        for k in ("<Control-y>", "<Control-Y>", "<Control-Shift-Z>"):
            root.bind_all(k, lambda e: (self.redo(), "break")[1])
        abtn("프로젝트 저장", self.save, "#c8e6c9")
        abtn("다른 이름으로", lambda: self.save(ask=True))
        abtn("큰 창 미리보기", self.preview, "#bbdefb")
        if IS_DMB:
            mb_rf = tk.Menubutton(act, text="TV 다시 가져오기 ▼", relief="raised", bg="#e1bee7")
            mb_rf.pack(side=tk.LEFT, padx=2)
            m_rf = tk.Menu(mb_rf, tearoff=0, postcommand=lambda: self.fill_tv_menu(m_rf, self.refresh_tv))
            mb_rf.config(menu=m_rf)
        else:
            abtn("새 판 반영", self.new_version, "#ffe0b2")
        self.tv_path = ""           # DMB: 불러온 TV 프로젝트
        self.dmb_on = IS_DMB        # DMB 편집기는 늘 DMB 모양
        # DMB 줄 (DMB 편집기에만): 뺄 것 체크 - 설정은 DMB 편집기 설정에 저장
        cfgd = load_config()
        self.dmbbar = dbar = tk.Frame(root, bg="#f3e5f5")
        tk.Label(dbar, text="DMB에서 빼기:", bg="#f3e5f5", font=("", 9, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        hide0 = set(cfgd.get("dmb_hide", DMB_HIDE_DEFAULT))
        if not cfgd.get("dmb_shade_off"):       # 1.1.5: DMB 는 음영도 기본으로 뺌 (이전에 저장한 설정에도 한 번 적용)
            hide0.add("shade")
            save_config_keys({"dmb_hide": sorted(hide0), "dmb_shade_off": True})
        if not cfgd.get("dmb_bold_off"):        # 1.1.7: 글자 굵기도 기본으로 뺌 (한 번만)
            hide0.add("bold")
            save_config_keys({"dmb_hide": sorted(hide0), "dmb_bold_off": True})
        self.v_dmb_hide = {}
        for k, t in DMB_HIDE_ITEMS:
            v = tk.BooleanVar(value=k in hide0)
            v.trace_add("write", lambda *a: self.dmb_changed())
            self.v_dmb_hide[k] = v
            tk.Checkbutton(dbar, text=t, variable=v, bg="#f3e5f5").pack(side=tk.LEFT)
        self.v_dmb_keep = tk.BooleanVar(value=bool(cfgd.get("dmb_keep_own", True)))
        self.v_dmb_keep.trace_add("write", lambda *a: self.dmb_changed())
        tk.Checkbutton(dbar, text="| DMB에서 넣은 음영·굵기·글자색은 그대로", variable=self.v_dmb_keep, bg="#f3e5f5",
                       fg="#6a1b9a").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(dbar, text="   (제목은 ⚙ 설정에서)", bg="#f3e5f5", fg="#777").pack(side=tk.LEFT)

        # 제목 · 날짜 (바꾸면 바로 편성표에 반영, 설정에 저장되어 다음 주에도 그대로)
        cfg0 = load_config()
        self.head = tk.Frame(root)      # (예전 설정 줄 자리 - 이제 ⚙ 설정은 따로 창으로)
        self.head_visible = False
        self.settings_win = None
        self.v_title = tk.StringVar(value=cfg0.get("title_text", "MBC 주간편성표"))
        self.v_tsize = tk.StringVar(value=str(cfg0.get("title_size", 22)))
        self.v_tbold = tk.BooleanVar(value=bool(cfg0.get("title_bold")))
        self.v_smode = tk.StringVar(value=cfg0.get("stamp_mode", "저장 시각"))
        self.v_stext = tk.StringVar(value=cfg0.get("stamp_text", ""))
        self._head_id = None
        for v in (self.v_title, self.v_tsize, self.v_tbold, self.v_smode, self.v_stext):
            v.trace_add("write", lambda *a: self.head_changed())

        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        self.nb = nb

        # ---- 탭 1: 편성표 보면서 고치기
        live = tk.Frame(nb)
        nb.add(live, text="  편성표 보면서 고치기  ")
        pw = tk.PanedWindow(live, orient=tk.HORIZONTAL, sashwidth=6)
        pw.pack(fill=tk.BOTH, expand=True)
        left = tk.Frame(pw)
        pw.add(left, stretch="always")
        zbar = tk.Frame(left)
        zbar.pack(fill=tk.X)
        tk.Button(zbar, text="－", width=3, command=lambda: self.set_zoom(self.zoom / 1.2)).pack(side=tk.LEFT)
        self.zoom_lbl = tk.Label(zbar, text="140%", width=6)
        self.zoom_lbl.pack(side=tk.LEFT)
        tk.Button(zbar, text="＋", width=3, command=lambda: self.set_zoom(self.zoom * 1.2)).pack(side=tk.LEFT)
        tk.Button(zbar, text="한 장 맞춤", command=self.zoom_fit).pack(side=tk.LEFT, padx=4)
        tk.Button(zbar, text="폭 맞춤", command=self.zoom_fit_width).pack(side=tk.LEFT)
        self.show_marks = tk.BooleanVar(value=True)
        tk.Checkbutton(zbar, text="선택·차이 테두리 보기 (끄면 출력물 모양 그대로)", variable=self.show_marks,
                       command=self.draw_marks).pack(side=tk.LEFT, padx=6)
        self.draw_lbl = tk.Label(zbar, text="클릭: 고르기 · Ctrl+클릭: 여러 칸 · Shift+클릭: 묶인 요일 전체 · Esc: 풀기 · 오른쪽 버튼: 메뉴", fg="#555")
        self.draw_lbl.pack(side=tk.LEFT, padx=8)
        self.last_log = tk.Label(zbar, text="", fg="#1565c0", anchor="e")     # 마지막 기록 한 줄 (전체는 [기록] 탭)
        self.last_log.pack(side=tk.RIGHT, padx=6)
        cvf = tk.Frame(left)
        cvf.pack(fill=tk.BOTH, expand=True)
        self.cv = tk.Canvas(cvf, bg="#808080", highlightthickness=0)
        ys = tk.Scrollbar(cvf, orient=tk.VERTICAL, command=self.cv.yview)
        xs = tk.Scrollbar(cvf, orient=tk.HORIZONTAL, command=self.cv.xview)
        self.cv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        xs.pack(side=tk.BOTTOM, fill=tk.X)
        self.cv.pack(fill=tk.BOTH, expand=True)
        self.cv.bind("<Button-1>", self.on_canvas_click)
        self.cv.bind("<Button-3>", self.on_canvas_menu)
        self.cv.bind("<Double-1>", lambda e: self.edit_row() if self.valid_sel() else self.add_at_slot())
        self.cv.bind("<Motion>", self.on_canvas_motion)
        self.cv.bind("<Leave>", lambda e: (self.cv.delete("tip"), setattr(self, "_tip_row", -1)))
        self.cv.bind("<MouseWheel>", self.on_wheel)
        self.cv.bind("<Button-4>", lambda e: self.cv.yview_scroll(-3, "units"))
        self.cv.bind("<Button-5>", lambda e: self.cv.yview_scroll(3, "units"))

        right = tk.Frame(pw, padx=6)
        pw.add(right, width=380, stretch="never")
        self.build_form(right)

        # ---- 탭 2: 전체 표
        table = tk.Frame(nb)
        nb.add(table, text="  전체 표  ")
        self.dayfrm = tk.Frame(table)
        self.dayfrm.pack(fill=tk.X)
        mid = tk.Frame(table)
        mid.pack(fill=tk.BOTH, expand=True, pady=4)
        self.tree = ttk.Treeview(mid, columns=[c[0] for c in self.COLS], show="headings", height=18)
        for key, title, w in self.COLS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=w, anchor="w" if key == "name" else "center")
        self.tree.tag_configure("pink", background="#ffe1e1")
        self.tree.tag_configure("edited", foreground="#c62828")
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self.edit_row())
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        side = tk.Frame(mid)
        side.pack(side=tk.LEFT, fill=tk.Y, padx=8)
        btn = lambda text, cmd, color=None: tk.Button(side, text=text, command=cmd, width=22,
                                                      bg=color or "#e0e0e0").pack(pady=2)
        btn("줄 고치기 (더블클릭)", self.edit_row)
        btn("줄 추가", self.add_row)
        btn("줄 삭제", self.del_row)
        btn("분홍 켜기 / 끄기", self.toggle_pink, "#ffd6d6")
        btn("음영 넣기 / 빼기", self.shade_dialog, "#d9d9d9")
        btn("끝 시각 표시 켜기 / 끄기", self.toggle_end, "#ffcdd2")
        btn("이 줄 TV 원래대로" if IS_DMB else "이 줄 엑셀 원래대로", self.revert_row)
        btn("로컬 끼워 넣기", self.local_dialog, "#cfd8dc")

        # 탭 3: 기록 (아래에 따로 두지 않고 탭으로)
        logtab = tk.Frame(nb)
        nb.add(logtab, text="  기록  ")
        self.logbox = tk.Text(logtab, height=8, wrap="none")
        lsb = tk.Scrollbar(logtab, orient=tk.VERTICAL, command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=lsb.set)
        lsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.logbox.pack(fill=tk.BOTH, expand=True)

        # 첫 화면: 엑셀을 열기 전에는 안내만 (회색 빈 화면 + 오른쪽 빈 칸)
        ph = tk.Frame(root)
        self.placeholder = ph
        left_ph = tk.Frame(ph, bg="#e4e6ea")
        left_ph.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Frame(ph, bg="#f0f3f8", width=380).pack(side=tk.RIGHT, fill=tk.Y)
        inner = tk.Frame(left_ph, bg="#e4e6ea")
        inner.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(inner, text=("TV 프로젝트가 열리지 않았습니다." if IS_DMB else "엑셀 파일이 열리지 않았습니다."),
                 bg="#e4e6ea", fg="#7a7a7a", font=("", 15, "bold")).pack()
        tk.Label(inner, text=("상단의 [TV 프로젝트 불러오기]로 엑셀 TV 편성표 편집기에서 저장한 프로젝트(.tvproj)를 "
                              "여세요." if IS_DMB else "상단의 [엑셀 열기]를 눌러 시작하세요."),
                 bg="#e4e6ea", fg="#9a9a9a", font=("", 10)).pack(pady=(14, 0))
        self.recent_box = tk.Frame(inner, bg="#e4e6ea")
        self.recent_box.pack(pady=(26, 0))
        try:
            tidy_projects(log=self.log)
        except Exception as e:
            self.log(f"[프로젝트 정리] 건너뜀: {e}")
        self.build_recent()
        self.main_shown = True
        self.show_start()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.tick()

    # ---------- 공통
    def log(self, s):
        self.logbox.insert(self.tk.END, s + "\n")
        self.logbox.see(self.tk.END)
        try:
            first = s.strip().splitlines()[0] if s.strip() else ""
            self.last_log.config(text=(first[:70] + "…") if len(first) > 70 else first,
                                 fg="#c62828" if "오류" in first or "주의" in first or "겹침" in first else "#1565c0")
        except Exception:
            pass
        self.root.update_idletasks()

    def elapsed(self):
        if self.t0 is None:
            return "00:00"
        sec = int(datetime.datetime.now().timestamp() - self.t0)
        return f"{sec // 60:02d}:{sec % 60:02d}"

    def tick(self):
        if self.xlsx:
            self.status.config(text=f"{os.path.basename(self.xlsx)}"
                                    + f"  ·  작업 시간 {self.elapsed()}  ·  고친 줄 {self.edits}개"
                                    + ("" if self.saved else "  ·  저장 안 됨"))
        self.root.after(1000, self.tick)

    def error(self, title, e):
        import traceback
        from tkinter import messagebox
        self.log("[오류] " + traceback.format_exc())
        messagebox.showerror(title, f"{type(e).__name__}: {e}")

    def mark_edit(self, r):
        r["edited"] = True
        for k in ("_check", "_note", "_ovl"):     # 확인할 칸(빨강)을 고치면 표시를 내림
            r.pop(k, None)
        self.saved = False

    # ---------- 불러오기
    def open_xlsx(self):
        from tkinter import filedialog, messagebox
        if self.xlsx and not self.saved:
            ans = messagebox.askyesnocancel("엑셀 열기", "지금 고친 내용을 저장하지 않았습니다.\n저장하고 새 엑셀을 열까요?")
            if ans is None or (ans and not self.save()):
                return
        f = self.v_x.get().strip()
        if not f or not os.path.isfile(f):
            f = filedialog.askopenfilename(filetypes=[("엑셀", "*.xlsx")])
            if not f:
                return
        try:
            rows = read_schedule_xlsx(f)
        except Exception as e:
            self.error("엑셀 읽기 오류", e)
            return
        self.xlsx = f
        self.v_x.set(f)
        self.proj_path = None
        self.proj_dir = ""
        self.history = []
        self.last_pdf = ""
        self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(f)}")
        self.show_main()
        self.query_date = read_query_date(f)
        self.orig = json.loads(json.dumps(rows))
        self.rows = rows
        self.deleted = 0
        self.sel = None
        self.multi = []
        self.diff_keys = set()
        if os.path.exists(work_path(f)):
            try:
                wrows, wedits, saved = load_work(f)
                if messagebox.askyesno("이어서 하기", f"저장해 둔 작업이 있습니다 ({saved}).\n이어서 할까요?"):
                    self.rows = wrows
            except Exception as e:
                self.log(f"[주의] 작업 파일을 읽지 못했습니다: {e}")
        if not any(r.get("local") for r in self.rows):
            n = apply_local_rules(self.rows, load_config(), self.log)
            if n:
                self.log(f"[로컬 규칙] {n}개를 자동으로 넣었습니다. (설정 파일 local_rules)")
        self.t0 = datetime.datetime.now().timestamp()
        self.saved = True
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_buttons()
        self.build_day_buttons()
        self.refresh()
        self.root.after(300, self.apply_saved_zoom)
        self.log(f"[불러옴] {os.path.basename(f)}  편성 {len(self.rows)}줄 (작업 시간 재기 시작)")

    def pick_prev(self):
        from tkinter import filedialog
        f = filedialog.askopenfilename(title="지난주 엑셀 (달라진 칸 분홍 표시용)", filetypes=[("엑셀", "*.xlsx")])
        if f:
            self.v_prev.set(f)
            if self.xlsx:
                self.log(f"[지난주 엑셀] {os.path.basename(f)} - 지난주와 달라진 칸을 분홍으로 표시합니다.")
                self.schedule_redraw(50)

    def build_day_buttons(self):
        for w in self.dayfrm.winfo_children():
            w.destroy()
        dates = sorted({r["date"] for r in self.rows})
        self.tk.Label(self.dayfrm, text="보기:").pack(side=self.tk.LEFT)
        for d in ["전체"] + dates:
            self.tk.Radiobutton(self.dayfrm, text=d if d == "전체" else day_label(d), value=d,
                                variable=self.day, indicatoron=False, width=10,
                                command=self.refresh).pack(side=self.tk.LEFT, padx=1)
        self.day.set("전체")

    def refresh(self, keep=None):
        """줄 내용이 바뀐 뒤: 표, 오른쪽 수정란, 편성표를 모두 다시"""
        self.mark_preview_dirty()
        if keep is not None:
            self.sel = keep if 0 <= keep < len(self.rows) else None
        elif not self.valid_sel():
            self.sel = None
        self.refresh_tree()
        self.select(self.sel, keep_multi=True)
        self.schedule_redraw()

    def refresh_tree(self):
        self.cfg = load_config()
        sel = self.sel
        self.tree.delete(*self.tree.get_children())
        day = self.day.get()
        order = sorted(range(len(self.rows)), key=lambda i: (self.rows[i]["date"], self.rows[i]["plan"],
                                                             self.rows[i]["s"]))
        for i in order:
            r = self.rows[i]
            if day != "전체" and r["date"] != day:
                continue
            tags = []
            shade = r.get("shade") or rule_shade(r, self.cfg)
            if r.get("shade"):
                tag = "sh_" + r["shade"].lstrip("#")
                self.tree.tag_configure(tag, background=r["shade"])
                tags.append(tag)
            elif self.eff_pink(i):
                tags.append("pink")
            elif shade:
                tag = "sh_" + shade.lstrip("#")
                self.tree.tag_configure(tag, background=shade)
                tags.append(tag)
            if r.get("edited"):
                tags.append("edited")
            self.tree.insert("", self.tk.END, iid=str(i), tags=tags, values=(
                day_label(r["date"]), fmt_t(r["s"]), fmt_t(r["e"]), show_prefix(r) + r["name"], r["plan"], icon_text(r),
                ("음영" if r.get("shade") else
                 ("분홍(직접)" if r.get("force_pink") is True else "분홍") if self.eff_pink(i) else
                 "분홍 끔" if r.get("force_pink") is False else "음영(자동)" if shade else ""),
                "●" if (r.get("show_end") or rule_end(r, self.cfg)) else "",
                "●" if r.get("edited") else ""))
        if sel is not None and self.tree.exists(str(sel)):
            self._quiet = True
            self.tree.selection_set(str(sel))
            self.tree.see(str(sel))
            self._quiet = False

    def selected_index(self):
        return self.sel if self.valid_sel() else None

    def need_row(self):
        from tkinter import messagebox
        i = self.selected_index()
        if i is None:
            messagebox.showinfo("알림", "편성표나 표에서 칸을 먼저 고르세요.")
        return i

    # ---------- 줄 편집
    def edit_row(self, idx=None, new=False):
        tk, ttk = self.tk, self.ttk
        from tkinter import messagebox
        i = idx if idx is not None else self.need_row()
        if i is None:
            return
        r = self.rows[i]
        top = tk.Toplevel(self.root)
        top.title("편성 줄 추가" if new else "편성 줄 고치기")
        top.transient(self.root)
        f = tk.Frame(top, padx=12, pady=10)
        f.pack()
        v = {k: tk.StringVar(value=val) for k, val in (
            ("name", r["name"]), ("s", fmt_t(r["s"])), ("e", fmt_t(r["e"])), ("plan", r["plan"]),
            ("age", r["age"]), ("rl", r["rl"]), ("code", r.get("code", "")))}
        b = {k: tk.BooleanVar(value=bool(r.get(k))) for k in ("uhd", "cc", "sign", "ad", "force_pink", "show_end")}
        pink_now = self.eff_pink(i)
        b["force_pink"].set(pink_now)
        tk.Label(f, text=day_label(r["date"]), font=("", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")
        tk.Label(f, text="프로그램명").grid(row=1, column=0, sticky="w")
        e_name = tk.Entry(f, textvariable=v["name"], width=46)
        e_name.grid(row=1, column=1, columnspan=3, pady=2)
        bind_linebreak(e_name)
        tk.Label(f, text="시작 (HH:MM)").grid(row=2, column=0, sticky="w")
        tk.Entry(f, textvariable=v["s"], width=8).grid(row=2, column=1, sticky="w")
        tk.Label(f, text="끝 (25:30 처럼)").grid(row=2, column=2, sticky="w")
        tk.Entry(f, textvariable=v["e"], width=8).grid(row=2, column=3, sticky="w")
        tk.Label(f, text="편성안").grid(row=3, column=0, sticky="w")
        ttk.Combobox(f, textvariable=v["plan"], values=("1안", "2안"), width=6, state="readonly").grid(
            row=3, column=1, sticky="w")
        tk.Label(f, text="연령").grid(row=3, column=2, sticky="w")
        ttk.Combobox(f, textvariable=v["age"], values=("", "00", "07", "12", "15", "19"), width=6,
                     state="readonly").grid(row=3, column=3, sticky="w")
        tk.Label(f, text="R/L (R=릴레이, V=부분로컬)").grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Combobox(f, textvariable=v["rl"], values=("", "R", "V", "L"), width=6, state="readonly").grid(
            row=4, column=2, sticky="w")
        cf = tk.Frame(f)
        cf.grid(row=5, column=0, columnspan=4, sticky="w", pady=4)
        for k, t in (("uhd", "UHD"), ("cc", "자막"), ("sign", "수어"), ("ad", "화면해설"), ("force_pink", "분홍"),
                     ("show_end", "끝 시각 표시")):
            tk.Checkbutton(cf, text=t, variable=b[k]).pack(side=tk.LEFT, padx=3)
        tk.Label(f, text="음영(배경색)").grid(row=7, column=0, sticky="w")
        v["shade"] = tk.StringVar(value=r.get("shade", ""))
        names = list(SHADE_PRESETS)
        cur = next((k for k, c in SHADE_PRESETS.items() if c and c == v["shade"].get()), None)
        v_sn = tk.StringVar(value=cur or ("없음" if not v["shade"].get() else "직접 고른 색"))
        cb = ttk.Combobox(f, textvariable=v_sn, values=names, width=10, state="readonly")
        cb.grid(row=7, column=1, sticky="w")
        cb.bind("<<ComboboxSelected>>", lambda e: v["shade"].set(SHADE_PRESETS.get(v_sn.get(), "")))

        def pick_color():
            from tkinter import colorchooser
            c = colorchooser.askcolor(color=v["shade"].get() or "#d9d9d9", parent=top)[1]
            if c:
                v["shade"].set(c)
                v_sn.set("직접 고른 색")
        tk.Button(f, text="색 고르기", command=pick_color).grid(row=7, column=2, sticky="w")
        tk.Label(f, text="프로그램코드").grid(row=6, column=0, sticky="w")
        tk.Entry(f, textvariable=v["code"], width=16).grid(row=6, column=1, sticky="w")
        tk.Label(f, text="(같은 코드·같은 시각이면 옆 날짜와 한 칸으로 묶임)", fg="#777").grid(
            row=6, column=2, columnspan=2, sticky="w")

        # ★ 즐겨찾기: 고르면 이름 + 저장해 둔 옵션(음영·굵게·아이콘 등)을 한 번에. 이름 칸에 치면 목록이 좁혀짐
        favs = [fv for fv in load_config().get("favorites", []) if isinstance(fv, dict)]
        fav_items = [(fav_display(fv).replace("\n", " "), fv) for fv in favs]
        extra = {}          # 이 창에 칸이 없는 옵션(굵게·점선·글자 바탕·글자색)은 확인 때 넣음
        v_fav = tk.StringVar(value="")
        fav_cb = ttk.Combobox(f, textvariable=v_fav, values=[t for t, _ in fav_items], width=24, state="readonly")
        fav_cb.grid(row=1, column=4, sticky="w", padx=(6, 0))
        tk.Label(f, text="★ 즐겨찾기", fg="#b26a00").grid(row=0, column=4, sticky="sw", padx=(6, 0))
        shown = list(fav_items)

        def fav_filter(*a):
            nonlocal shown
            key = norm_key(v["name"].get().replace("|", ""))
            if not key or v["name"].get().strip() in ("새 프로그램", r["name"]):
                shown = list(fav_items)
            else:
                shown = [(t, fv) for t, fv in fav_items if key in norm_key(t)] or list(fav_items)
            fav_cb.configure(values=[t for t, _ in shown])
        v["name"].trace_add("write", fav_filter)

        def fav_pick(e=None):
            t = v_fav.get()
            fv = next((fv for tt, fv in shown if tt == t), None) or next((fv for tt, fv in fav_items if tt == t), None)
            if not fv:
                return
            name = fv.get("name") or v["name"].get()
            pre, suf = fv.get("prefix", ""), fv.get("suffix", "")
            if pre and not name.startswith(pre):
                name = pre + name
            if suf and not name.endswith(suf):
                name = name + suf
            v["name"].set(name)
            for k in ("uhd", "cc", "sign", "ad", "force_pink", "show_end"):
                if k in fv:
                    b[k].set(bool(fv[k]))
            if "age" in fv:
                v["age"].set(fv["age"] or "")
            if "shade" in fv:
                v["shade"].set(fv["shade"] or "")
                v_sn.set(next((k for k, c in SHADE_PRESETS.items() if c == (fv["shade"] or "")),
                              "직접 고른 색" if fv["shade"] else "없음"))
            extra.clear()
            extra.update({k: fv[k] for k in ("bold", "top_dash", "bottom_dash", "hl", "fg") if k in fv})
            e_name.focus_set()
            e_name.icursor("end")
        fav_cb.bind("<<ComboboxSelected>>", fav_pick)
        if new and fav_items:       # 줄 추가로 열렸으면 목록을 펼쳐서 바로 고르게
            def open_list():
                try:
                    fav_cb.focus_set()
                    fav_cb.event_generate("<Down>")
                except Exception:
                    pass
            top.after(300, open_list)

        def ok():
            try:
                s_ = hhmm_to_min(v["s"].get()) - DAY_START_MIN
                e_ = hhmm_to_min(v["e"].get()) - DAY_START_MIN
            except ValueError:
                messagebox.showwarning("입력 오류", "시각은 HH:MM 으로 입력하세요. (예: 07:50, 25:30)", parent=top)
                return
            if e_ <= s_:
                messagebox.showwarning("입력 오류", "끝 시각이 시작보다 늦어야 합니다. 자정 넘으면 25:30 처럼.",
                                       parent=top)
                return
            if not (0 <= s_ < DAY_LEN_MIN and 0 < e_ <= DAY_LEN_MIN):
                messagebox.showwarning("입력 오류", "05:00 ~ 29:00 사이로 입력하세요.", parent=top)
                return
            if not v["name"].get().strip():
                messagebox.showwarning("입력 오류", "프로그램명을 입력하세요.", parent=top)
                return
            self.push_undo("edit")
            mates = [k for k in (self.whole_block() or []) if k != i] if i == self.sel else []
            if v["name"].get().strip() != r["name"]:
                r["name_edited"] = True        # 이름을 바꾼 줄은 옆 날짜와 묶지 않음
            r.update(name=v["name"].get().strip(), s=s_, e=e_, plan=v["plan"].get(), age=v["age"].get(),
                     rl=v["rl"].get(), code=v["code"].get().strip(), shade=v["shade"].get(),
                     **{k: b[k].get() for k in b if k != "force_pink"})
            if b["force_pink"].get() != pink_now:      # 바꿨을 때만 직접 값으로
                r["force_pink"] = b["force_pink"].get()
            r.update(extra)                             # 즐겨찾기에서 고른 굵게·점선·글자 바탕·글자색
            self.mark_edit(r)
            for k in mates:             # 월~금으로 묶인 한 칸을 고쳤으면 묶인 요일 모두 똑같이
                x = self.rows[k]
                x.update({kk: r.get(kk) for kk in BLOCK_SYNC_KEYS if kk in r})
                self.mark_edit(x)
            self.log(f"[고침] {day_label(r['date'])} {fmt_t(s_)}~{fmt_t(e_)} {r['name']} ({r['plan']})")
            top.destroy()
            self.refresh(keep=i)
            self.ask_overlap(i)

        def cancel():
            if new:
                self.rows.pop(i)
                self.refresh()
            top.destroy()
        bf = tk.Frame(f)
        bf.grid(row=8, column=0, columnspan=5, pady=8)
        tk.Button(bf, text="확인", command=ok, width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="취소", command=cancel, width=10).pack(side=tk.LEFT, padx=6)
        top.protocol("WM_DELETE_WINDOW", cancel)
        top.bind("<Return>", lambda e: ok())
        top.after(150, lambda: top.grab_set())

    def ask_overlap(self, i=None, force=False):
        """시간을 바꾼 칸이 다른 칸과 겹치면: 겹치는 칸을 지우거나 줄일지 물음 (아니오면 그대로 두고 기록에 주의)"""
        from tkinter import messagebox
        if getattr(self, "_ovl_id", None):
            try:
                self.root.after_cancel(self._ovl_id)
            except Exception:
                pass
            self._ovl_id = None
        if getattr(self, "_ovl_busy", False):
            return
        if i is None:       # 수정란에서 시간을 바꾼 그 칸 (바꾼 게 없으면 묻지 않음)
            r0, self._ovl_row = getattr(self, "_ovl_row", None), None
            i = next((k for k, x in enumerate(self.rows) if x is r0), None)
        if i is None or not (0 <= i < len(self.rows)):
            return
        r = self.rows[i]
        plan = overlap_plan(self.rows, i)
        key = (r["date"], r["plan"], r["s"], r["e"], r["name"])
        if not plan or (key == getattr(self, "_ovl_no", None) and not force):
            return
        self._ovl_busy = True
        try:
            msg = (f"{day_label(r['date'])} '{r['name']}' {fmt_t(r['s'])}~{fmt_t(r['e'])} 이(가) 다른 칸과 겹칩니다.\n\n"
                   + "\n".join("· " + t for _, _, t in plan) + "\n\n이렇게 정리할까요?  (아니오: 그대로 둠)")
            ok = messagebox.askyesno("겹치는 칸 정리", msg)
        finally:
            self._ovl_busy = False
        if not ok:
            self._ovl_no = key
            self.check_overlap(r["date"])
            return
        self.push_undo("overlap")
        ni, n_del = resolve_overlaps(self.rows, i, plan)
        self.deleted += n_del
        self.saved = False
        for _, _, t in plan:
            self.log(f"[겹침 정리] {day_label(r['date'])} {t}")
        self.multi = []
        self.refresh(keep=ni)

    def set_show(self, k):
        """고른 칸들의 보이는 곳: '' 모두 / tv TV만 / dmb DMB만"""
        rows = self.selected_rows()
        if not rows:
            return
        self.push_undo("show")
        targets = [self.rows[i] for i in rows]
        cur = self.rows[self.sel] if self.valid_sel() else None
        for x in targets:
            old = x.get("show") or ""
            x["show"] = k
            self.mark_edit(x)
            self.after_show_change(x, old)
        self.log(f"[보이는 곳] {len(targets)}칸 -> {SHOW_LABELS[k]}")
        self.multi = []
        self.refresh(keep=next((n for n, x in enumerate(self.rows) if x is cur), -1))

    def after_show_change(self, x, old):
        """'DMB만'으로 바꾼 칸이 TV 편성을 잘라 두었으면(로컬 끼워 넣기) TV 쪽은 엑셀 원래대로 되살림"""
        if (x.get("show") or "") != "dmb" or old == "dmb":
            return
        n = restore_under(self.rows, self.orig, self.rows.index(x))
        if n:
            self.log(f"[DMB만] '{x['name']}' 때문에 잘렸던 TV 편성 {n}칸을 엑셀 원래대로 되살렸습니다.")

    def row_reason(self, i):
        """이 칸에 붙은 표시의 이유 한 줄: (글, 색, 겹침 정리 버튼을 보일지)"""
        if i is None or not (0 <= i < len(self.rows)):
            return "", "", False
        r = self.rows[i]
        ov = find_overlaps(self.rows, i)
        if ov:
            x = self.rows[ov[0]]
            a, b = max(r["s"], x["s"]), min(r["e"], x["e"])
            more = f" 외 {len(ov) - 1}칸" if len(ov) > 1 else ""
            return (f"🔴 겹침: {x['name']}({fmt_t(x['s'])}~{fmt_t(x['e'])}){more}과 {fmt_t(a)}~{fmt_t(b)} 겹침",
                    "#c62828", True)
        if r.get("_check") and r.get("_note"):
            return "🔴 확인: " + r["_note"], "#c62828", False
        if r.get("_ver") == "changed":
            return f"🟠 새 판에서 바뀜: 예전 {r.get('_was', '')} → 지금 {fmt_t(r['s'])}~{fmt_t(r['e'])}", "#e65100", False
        if r.get("_ver") == "new":
            return "🟠 새 판에서 새로 생김", "#e65100", False
        if r.get("show") == "dmb":
            return "🟣 DMB에서만 보이는 칸 (TV에는 안 나옴)", "#8e24aa", False
        if r.get("show") == "tv":
            return "🟣 TV에서만 보이는 칸 (DMB에는 안 나옴)", "#8e24aa", False
        return "", "", False

    def update_reason(self):
        txt, color, btn = self.row_reason(self.sel)
        if not txt:
            self.reason_bar.pack_forget()
            return
        self.reason_lbl.config(text=txt, fg=color)
        if btn:
            self.reason_btn.pack(side=self.tk.RIGHT, padx=2)
        else:
            self.reason_btn.pack_forget()
        self.reason_bar.pack(fill=self.tk.X, pady=(0, 4), after=self.form_title)

    def on_canvas_motion(self, e):
        """칸 위에 마우스: 표시 이유를 작은 말풍선으로"""
        i = self.hit(e)
        if i == getattr(self, "_tip_row", -1):
            return
        self._tip_row = i
        self.cv.delete("tip")
        txt, color, _ = self.row_reason(i)
        if not txt:
            return
        x, y = self.cv.canvasx(e.x) + 14, self.cv.canvasy(e.y) + 16
        t = self.cv.create_text(x + 4, y + 3, text=txt, anchor="nw", fill=color, font=("", 9, "bold"),
                                width=360, tags=("tip",))
        bb = self.cv.bbox(t)
        if bb:
            r = self.cv.create_rectangle(bb[0] - 4, bb[1] - 3, bb[2] + 4, bb[3] + 3, fill="#fffde7",
                                         outline=color, tags=("tip",))
            self.cv.tag_lower(r, t)

    def check_overlap(self, date):
        for plan in ("1안", "2안"):
            day = sorted([r for r in self.rows if r["date"] == date and r["plan"] == plan], key=lambda r: r["s"])
            for a, b2 in zip(day, day[1:]):
                if b2["s"] < a["e"]:
                    self.log(f"[겹침 주의] {day_label(date)} {plan}: '{a['name']}'({fmt_t(a['s'])}~{fmt_t(a['e'])}) 와 "
                             f"'{b2['name']}'({fmt_t(b2['s'])}~{fmt_t(b2['e'])}) 시간이 겹칩니다.")

    def add_row(self):
        self.push_undo("add")
        i = self.selected_index()
        if i is not None:
            base = self.rows[i]
            date, s_, plan = base["date"], base["e"], base["plan"]
        else:
            day = self.day.get()
            dates = sorted({r["date"] for r in self.rows})
            if not dates:
                return
            date, s_, plan = (day if day != "전체" else dates[0]), 0, "1안"
        r = {"date": date, "code": "", "name": "새 프로그램", "s": s_, "e": min(DAY_LEN_MIN, s_ + 30),
             "uhd": False, "cc": True, "ad": False, "sign": False, "age": "", "rl": "L", "rerun": False,
             "plan": plan}
        if self.dmb_on:         # DMB 로 보는 중에 넣는 줄은 DMB에서만
            r["show"] = "dmb"
        self.rows.append(r)
        self.refresh(keep=len(self.rows) - 1)
        self.edit_row(len(self.rows) - 1, new=True)

    def del_row(self):
        from tkinter import messagebox
        i = self.need_row()
        if i is None:
            return
        r = self.rows[i]
        sel = sorted(self.selected_rows(), reverse=True)
        msg = (f"고른 {len(sel)}줄을 지울까요?" if len(sel) > 1 else
               f"{day_label(r['date'])} {fmt_t(r['s'])} '{r['name']}' 줄을 지울까요?")
        if messagebox.askyesno("줄 삭제", msg):
            self.push_undo("del")
            for k in sel:
                x = self.rows.pop(k)
                self.deleted += 1
                self.log(f"[삭제] {day_label(x['date'])} {fmt_t(x['s'])} {x['name']}")
            self.saved = False
            self.multi = []
            self.refresh(keep=-1)

    def eff_pink(self, i):
        """이 줄이 지금 분홍인지 (직접 켠/끈 값이 먼저, 없으면 2안·지난주 비교로 정해진 자동 값)"""
        r = self.rows[i]
        if isinstance(r.get("force_pink"), bool):
            return r["force_pink"]
        if i in self.pink_state:
            return self.pink_state[i]
        return r["plan"] == "2안"

    def toggle_pink(self):
        i = self.need_row()
        if i is None:
            return
        r = self.rows[i]
        self.push_undo("pink")
        on = not self.eff_pink(i)
        for k in self.selected_rows():
            self.rows[k]["force_pink"] = on
            self.mark_edit(self.rows[k])
        n = len(self.selected_rows())
        self.log(f"[분홍] {day_label(r['date'])} {fmt_t(r['s'])} {r['name']}" + (f" 외 {n - 1}칸" if n > 1 else "")
                 + f" -> {'켬' if on else '끔'}")
        self.refresh(keep=i)

    def toggle_end(self):
        """로컬처럼 기본 프로그램보다 먼저 끝나는 칸에 끝나는 분 표시"""
        from tkinter import messagebox
        i = self.need_row()
        if i is None:
            return
        r = self.rows[i]
        on = not r.get("show_end")
        key = norm_key(split_suffix(r["name"])[0])
        same = [x for x in self.rows if x is not r and x["s"] == r["s"]
                and norm_key(split_suffix(x["name"])[0]) == key]
        targets = [r]
        if len(self.selected_rows()) > 1:       # 여러 칸을 골랐으면 고른 칸에만
            targets = [self.rows[k] for k in self.selected_rows()]
            same = []
        if same:
            ans = messagebox.askyesnocancel(
                "끝 시각 표시", f"이번 주 같은 시각의 같은 프로그램 {len(same)}줄에도 똑같이 적용할까요?\n"
                                "(예: 모두 / 아니오: 이 줄만)")
            if ans is None:
                return
            if ans:
                targets += same
        self.push_undo("end")
        for x in targets:
            x["show_end"] = on
            self.mark_edit(x)
        self.log(f"[끝 시각] '{r['name']}' {len(targets)}줄 {'표시' if on else '안 함'} "
                 f"(끝 {fmt_t(r['e'])})")
        self.refresh(keep=i)

    def affix_dialog(self):
        """고른 칸들의 프로그램명 앞뒤에 글자 붙이기/빼기 (예: R/L...MBC 뉴스데스크...R/L). 칸마다 이름은 그대로 둠"""
        tk = self.tk
        from tkinter import messagebox
        if self.need_row() is None:
            return
        picked = self.selected_rows()
        top = tk.Toplevel(self.root)
        top.title("이름 앞뒤에 글자")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        tk.Label(f, text=f"고른 {len(picked)}칸의 프로그램명 앞뒤에 붙입니다.", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w")
        v_pre, v_suf = tk.StringVar(value="R/L..."), tk.StringVar(value="...R/L")
        tk.Label(f, text="앞").grid(row=1, column=0, sticky="w")
        tk.Entry(f, textvariable=v_pre, width=14).grid(row=1, column=1, sticky="w", pady=3)
        tk.Label(f, text="뒤").grid(row=1, column=2, sticky="e")
        tk.Entry(f, textvariable=v_suf, width=14).grid(row=1, column=3, sticky="w")
        ex = self.rows[picked[0]]["name"]
        prev = tk.Label(f, fg="#1565c0")
        prev.grid(row=2, column=0, columnspan=4, sticky="w")

        def show(*a):
            prev.config(text=f"예: {v_pre.get()}{ex}{v_suf.get()}")
        v_pre.trace_add("write", show)
        v_suf.trace_add("write", show)
        show()

        def run(add):
            pre, suf = v_pre.get(), v_suf.get()
            if not pre and not suf:
                return
            self.push_undo("affix")
            n = 0
            for k in picked:
                x = self.rows[k]
                name = x["name"]
                if add:
                    if pre and not name.startswith(pre):
                        name = pre + name
                    if suf and not name.endswith(suf):
                        name = name + suf
                else:
                    if pre and name.startswith(pre):
                        name = name[len(pre):]
                    if suf and name.endswith(suf):
                        name = name[:-len(suf)]
                if name != x["name"]:
                    x["name"], x["name_edited"] = name, True
                    self.mark_edit(x)
                    n += 1
            self.log(f"[이름 앞뒤] {n}칸 {'붙임' if add else '뺌'}: '{pre}' ... '{suf}'")
            top.destroy()
            self.refresh(keep=self.sel)
            if n == 0:
                messagebox.showinfo("알림", "바뀐 칸이 없습니다.")
        bf = tk.Frame(f)
        bf.grid(row=3, column=0, columnspan=4, pady=8)
        tk.Button(bf, text="붙이기", command=lambda: run(True), width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="빼기", command=lambda: run(False), width=10).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="취소", command=top.destroy, width=10).pack(side=tk.LEFT, padx=4)
        top.after(150, lambda: top.grab_set())

    def local_dialog(self):
        """고른 전국 프로그램 시간 안에 로컬 프로그램을 끼워 넣음 (여러 요일 한꺼번에, 매주 자동 저장 가능)"""
        tk, ttk = self.tk, self.ttk
        from tkinter import messagebox
        i = self.need_row()
        if i is None:
            return
        base = self.rows[i]
        top = tk.Toplevel(self.root)
        top.title("DMB 칸 넣기 (TV 편성 위에 덮기)" if self.dmb_on else "로컬 끼워 넣기")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        tk.Label(f, text=f"기준: {day_label(base['date'])} {fmt_t(base['s'])}~{fmt_t(base['e'])} {base['name']}",
                 font=("", 10, "bold")).grid(row=0, column=0, columnspan=8, sticky="w")
        v_name = tk.StringVar(value="광주MBC ")
        v_s = tk.StringVar(value=fmt_t(base["s"]))
        v_e = tk.StringVar(value=fmt_t(base["e"]))
        tk.Label(f, text="로컬 프로그램명").grid(row=1, column=0, sticky="w")
        ent = tk.Entry(f, textvariable=v_name, width=36)
        ent.grid(row=1, column=1, columnspan=7, sticky="w", pady=3)
        tk.Label(f, text="시작").grid(row=2, column=0, sticky="w")
        tk.Entry(f, textvariable=v_s, width=7).grid(row=2, column=1, sticky="w")
        tk.Label(f, text="끝").grid(row=2, column=2, sticky="e")
        tk.Entry(f, textvariable=v_e, width=7).grid(row=2, column=3, sticky="w")
        tk.Label(f, text="요일").grid(row=3, column=0, sticky="w")
        key = norm_key(split_suffix(base["name"])[0])
        has = {datetime.datetime.strptime(r["date"], "%Y%m%d").weekday() for r in self.rows
               if r["plan"] != "2안" and norm_key(split_suffix(r["name"])[0]) == key and r["s"] == base["s"]}
        dates = sorted({r["date"] for r in self.rows})
        wd_of = {datetime.datetime.strptime(d, "%Y%m%d").weekday(): d for d in dates}
        v_days = {}
        for k in range(7):
            v_days[k] = tk.BooleanVar(value=k in has)
            tk.Checkbutton(f, text=WEEKDAYS[k], variable=v_days[k]).grid(row=3, column=1 + k, sticky="w")
        tk.Label(f, text=f"(같은 시각에 '{split_suffix(base['name'])[0]}'이 있는 요일을 미리 골라 둠)",
                 fg="#777").grid(row=4, column=0, columnspan=8, sticky="w")
        v_shade = tk.StringVar(value="회색")
        v_bold = tk.BooleanVar(value=True)
        v_cc = tk.BooleanVar(value=True)
        v_rule = tk.BooleanVar(value=False)
        tk.Label(f, text="음영").grid(row=5, column=0, sticky="w")
        ttk.Combobox(f, textvariable=v_shade, values=list(SHADE_PRESETS), width=8, state="readonly").grid(
            row=5, column=1, columnspan=2, sticky="w", pady=3)
        tk.Checkbutton(f, text="굵게", variable=v_bold).grid(row=5, column=3, columnspan=2, sticky="w")
        tk.Checkbutton(f, text="자막", variable=v_cc).grid(row=5, column=5, columnspan=2, sticky="w")
        tk.Checkbutton(f, text="매주 자동으로 넣기 (설정 파일에 규칙 저장)", variable=v_rule).grid(
            row=6, column=0, columnspan=8, sticky="w")

        def ok():
            name = v_name.get().strip()
            try:
                s_ = hhmm_to_min(v_s.get()) - DAY_START_MIN
                e_ = hhmm_to_min(v_e.get()) - DAY_START_MIN
                if e_ <= s_ or not (0 <= s_ < DAY_LEN_MIN and 0 < e_ <= DAY_LEN_MIN):
                    raise ValueError
            except ValueError:
                messagebox.showwarning("입력 오류", "시각을 확인하세요. (예: 07:20 ~ 07:50)", parent=top)
                return
            days = [k for k in range(7) if v_days[k].get() and k in wd_of]
            if not name or not days:
                messagebox.showwarning("입력 오류", "로컬 프로그램명과 요일을 고르세요.", parent=top)
                return
            shade = SHADE_PRESETS.get(v_shade.get(), "#d9d9d9")
            self.push_undo("local")
            if self.dmb_on:     # DMB 로 보는 중: TV 편성은 그대로 두고 DMB에서만 덮는 칸
                added = []
                for d in [wd_of[k] for k in days]:
                    x = {"date": d, "code": "DMB-" + norm_key(name), "name": name, "s": s_, "e": e_, "uhd": False,
                         "cc": v_cc.get(), "ad": False, "sign": False, "age": "", "rl": "", "rerun": False,
                         "plan": "1안", "shade": shade or "", "bold": bool(v_bold.get()), "local": True,
                         "edited": True, "show": "dmb"}
                    self.rows.append(x)
                    added.append(x)
                self.log(f"[DMB만] '{name}' {fmt_t(s_)}~{fmt_t(e_)} {len(added)}칸 - TV에는 안 보입니다.")
                v_rule.set(False)
            else:
                added = insert_local(self.rows, [wd_of[k] for k in days], name, s_, e_, shade, v_bold.get(),
                                     extra={"cc": v_cc.get()}, log=self.log)
            spec = "".join(WEEKDAYS[k] for k in days)
            if v_rule.get():
                add_local_rule({"name": name, "days": ",".join(WEEKDAYS[k] for k in days),
                                "start": fmt_t(s_), "end": fmt_t(e_), "shade": shade, "bold": v_bold.get(),
                                "cc": v_cc.get()})
                self.log(f"[로컬 규칙 저장] '{name}' {spec} {fmt_t(s_)}~{fmt_t(e_)} - 다음 주부터 자동")
            self.saved = False
            self.log(f"[로컬] '{name}' {spec} {fmt_t(s_)}~{fmt_t(e_)} 넣음 ({len(added)}줄)")
            top.destroy()
            self.refresh(keep=self.rows.index(added[0]) if added else None)
        bf = tk.Frame(f)
        bf.grid(row=7, column=0, columnspan=8, pady=8)
        tk.Button(bf, text="넣기", command=ok, width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="취소", command=top.destroy, width=10).pack(side=tk.LEFT, padx=6)
        ent.focus_set()
        ent.icursor(tk.END)
        top.after(150, lambda: top.grab_set())

    def copy_style(self, i):
        """이 칸의 꾸밈(음영·글자 바탕·글자색·굵게·위/아래 점선)을 복사. 아이콘은 프로그램마다 달라서 뺌"""
        r = self.rows[i]
        self.style_clip = {k: r.get(k) or ("" if k in ("shade", "hl", "fg") else False) for k in STYLE_COPY_KEYS}
        self.style_clip["_from"] = r["name"].replace("|", " ")[:16]
        self.log(f"[꾸밈 복사] {day_label(r['date'])} {fmt_t(r['s'])} {r['name']} - "
                 f"음영 {self.style_clip['shade'] or '없음'}, 굵게 {'예' if self.style_clip['bold'] else '아니오'}")

    def paste_style(self):
        """복사해 둔 꾸밈을 고른 칸(여러 칸이면 모두)에 똑같이"""
        if not self.style_clip:
            return
        sel = self.selected_rows()
        if not sel:
            return
        self.push_undo("paste_style")
        for k in sel:
            x = self.rows[k]
            for key in STYLE_COPY_KEYS:
                x[key] = self.style_clip[key]
            self.mark_edit(x)
        self.saved = False
        self.log(f"[꾸밈 붙여넣기] '{self.style_clip['_from']}' 꾸밈 -> {len(sel)}칸")
        self.refresh(keep=self.sel)

    def plan2_dialog(self, i):
        """1안 칸과 같은 시간에 2안 프로그램 넣기 -> 그날은 1안(왼쪽 반) / 2안(오른쪽 반)으로 나뉨"""
        tk, ttk = self.tk, self.ttk
        from tkinter import messagebox
        base = self.rows[i]
        top = tk.Toplevel(self.root)
        top.title("2안 넣기")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        tk.Label(f, text=f"기준(1안): {day_label(base['date'])} {fmt_t(base['s'])}~{fmt_t(base['e'])} "
                         f"{base['name'].replace('|', ' ')}", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=8, sticky="w")
        v_name = tk.StringVar(value="")
        v_s, v_e = tk.StringVar(value=fmt_t(base["s"])), tk.StringVar(value=fmt_t(base["e"]))
        tk.Label(f, text="2안 프로그램명").grid(row=1, column=0, sticky="w")
        ent = tk.Entry(f, textvariable=v_name, width=36)
        ent.grid(row=1, column=1, columnspan=7, sticky="w", pady=3)
        bind_linebreak(ent)
        # ★ 즐겨찾기: 고르면 이름 + 저장해 둔 옵션(음영·아이콘 등)
        favs = [fv for fv in load_config().get("favorites", []) if isinstance(fv, dict)]
        fav_items = [(fav_display(fv).replace("\n", " "), fv) for fv in favs]
        picked_fav = {}
        tk.Label(f, text="★ 즐겨찾기", fg="#b26a00").grid(row=2, column=0, sticky="w")
        v_fav = tk.StringVar(value="")
        fav_cb = ttk.Combobox(f, textvariable=v_fav, values=[t for t, _ in fav_items], width=34, state="readonly")
        fav_cb.grid(row=2, column=1, columnspan=7, sticky="w", pady=3)

        def fav_pick(e=None):
            fv = next((fv for t, fv in fav_items if t == v_fav.get()), None)
            if not fv:
                return
            name = fv.get("name") or v_name.get()
            pre, suf = fv.get("prefix", ""), fv.get("suffix", "")
            if pre and not name.startswith(pre):
                name = pre + name
            if suf and not name.endswith(suf):
                name = name + suf
            v_name.set(name)
            picked_fav.clear()
            picked_fav.update({k: fv[k] for k in FAV_OPT_KEYS if k in fv})
            if "force_pink" in fv:
                v_pink.set(bool(fv["force_pink"]))
            ent.focus_set()
            ent.icursor(tk.END)
        fav_cb.bind("<<ComboboxSelected>>", fav_pick)
        tk.Label(f, text="시작").grid(row=3, column=0, sticky="w")
        tk.Entry(f, textvariable=v_s, width=7).grid(row=3, column=1, sticky="w")
        tk.Label(f, text="끝").grid(row=3, column=2, sticky="e")
        tk.Entry(f, textvariable=v_e, width=7).grid(row=3, column=3, sticky="w")
        tk.Label(f, text="요일").grid(row=4, column=0, sticky="w")
        dates = sorted({r["date"] for r in self.rows})
        wd_of = {datetime.datetime.strptime(d, "%Y%m%d").weekday(): d for d in dates}
        base_wd = datetime.datetime.strptime(base["date"], "%Y%m%d").weekday()
        v_days = {}
        for k in range(7):
            v_days[k] = tk.BooleanVar(value=(k == base_wd))
            tk.Checkbutton(f, text=WEEKDAYS[k], variable=v_days[k],
                           state="normal" if k in wd_of else "disabled").grid(row=4, column=1 + k, sticky="w")
        tk.Label(f, text="(오른쪽 클릭한 요일만 미리 골라 둠 - 다른 요일도 같이 넣으려면 체크)", fg="#777").grid(
            row=5, column=0, columnspan=8, sticky="w")
        v_pink = tk.BooleanVar(value=bool(load_config().get("pink_plan2", True)))
        v_cc = tk.BooleanVar(value=True)
        tk.Checkbutton(f, text="분홍", variable=v_pink).grid(row=6, column=1, sticky="w")
        tk.Checkbutton(f, text="자막", variable=v_cc).grid(row=6, column=2, columnspan=2, sticky="w")

        def ok():
            name = v_name.get().strip()
            try:
                s_ = hhmm_to_min(v_s.get()) - DAY_START_MIN
                e_ = hhmm_to_min(v_e.get()) - DAY_START_MIN
                if e_ <= s_ or not (0 <= s_ < DAY_LEN_MIN and 0 < e_ <= DAY_LEN_MIN):
                    raise ValueError
            except (ValueError, TypeError):
                messagebox.showwarning("입력 오류", "시각을 확인하세요. (예: 07:20 ~ 07:50)", parent=top)
                return
            days = [wd_of[k] for k in range(7) if v_days[k].get() and k in wd_of]
            if not name or not days:
                messagebox.showwarning("입력 오류", "2안 프로그램명과 요일을 고르세요.", parent=top)
                return
            busy = [d for d in days if any(r["date"] == d and r["plan"] == "2안" and r["s"] < e_ and r["e"] > s_
                                           and (r.get("show") or "") == ("dmb" if self.dmb_on else "")
                                           for r in self.rows)]
            if busy and not messagebox.askyesno(
                    "2안 넣기", "이미 2안이 있는 요일: " + ", ".join(day_label(d) for d in busy) +
                    "\n그 요일은 빼고 넣을까요?", parent=top):
                return
            days = [d for d in days if d not in busy]
            if not days:
                top.destroy()
                return
            self.push_undo("plan2")
            added = []
            for d in days:
                x = {"date": d, "code": "", "name": name, "s": s_, "e": e_, "uhd": False, "cc": v_cc.get(),
                     "ad": False, "sign": False, "age": "", "rl": "", "rerun": False, "plan": "2안", "edited": True}
                x.update(picked_fav)
                x["cc"] = v_cc.get() if "cc" not in picked_fav else picked_fav["cc"]
                if v_pink.get() != bool(load_config().get("pink_plan2", True)):
                    x["force_pink"] = v_pink.get()     # 설정과 다르게 고른 경우만 직접 값
                else:
                    x.pop("force_pink", None)
                if self.dmb_on:
                    x["show"] = "dmb"
                self.rows.append(x)
                added.append(x)
            self.saved = False
            self.log(f"[2안 넣기] '{name.replace('|', ' ')}' {fmt_t(s_)}~{fmt_t(e_)} "
                     + ", ".join(day_label(d) for d in days))
            top.destroy()
            self.refresh(keep=self.rows.index(added[0]))
        bf = tk.Frame(f)
        bf.grid(row=7, column=0, columnspan=8, pady=8)
        tk.Button(bf, text="넣기", command=ok, width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="취소", command=top.destroy, width=10).pack(side=tk.LEFT, padx=6)
        ent.focus_set()
        top.after(150, lambda: top.grab_set())

    def del_plan2(self, i):
        """2안 칸 지우기 -> 그날 그 시간은 다시 1안 전체 칸 (조건이 같으면 옆 날짜와 다시 묶임)"""
        from tkinter import messagebox
        targets = [k for k in self.selected_rows() if self.rows[k]["plan"] == "2안"] or [i]
        names = ", ".join(f"{day_label(self.rows[k]['date'])} {self.rows[k]['name'].replace('|', ' ')}"
                          for k in targets[:3]) + (" ..." if len(targets) > 3 else "")
        if not messagebox.askyesno("2안 지우기", f"2안 {len(targets)}칸을 지우고 1안만 남길까요?\n{names}"):
            return
        self.push_undo("del2")
        for k in sorted(targets, reverse=True):
            x = self.rows.pop(k)
            self.deleted += 1
            self.log(f"[2안 지움] {day_label(x['date'])} {fmt_t(x['s'])} {x['name']}")
        self.saved = False
        self.multi = []
        self.refresh(keep=-1)

    def plan_partners(self, i):
        """rows[i] 의 짝: 같은 날, 시간이 겹치는 다른 편성안(1안<->2안) 줄 번호들 (TV/DMB 같은 출력끼리)"""
        r = self.rows[i]
        other = "1안" if r["plan"] == "2안" else "2안"
        return [k for k, x in enumerate(self.rows) if k != i and x["date"] == r["date"] and x["plan"] == other
                and x["s"] < r["e"] and x["e"] > r["s"] and (x.get("show") or "") == (r.get("show") or "")]

    def swap_plan(self, i):
        """1안 <-> 2안 서로 바꾸기: 누른 칸과 짝(같은 날·겹치는 시간의 다른 편성안)을 함께 바꿈.
        여러 칸을 골랐으면 고른 칸마다 짝까지. 짝이 없는 칸은 바꾸지 않음"""
        base = self.selected_rows()
        if i not in base:
            base = [i]
        sel = set()
        for k in base:
            ps = self.plan_partners(k)
            if ps:
                sel.add(k)
                sel.update(ps)
        if not sel:
            return
        self.push_undo("swap")
        for k in sorted(sel):
            x = self.rows[k]
            x["plan"] = "1안" if x["plan"] == "2안" else "2안"
            x.pop("force_pink", None)          # 분홍은 바뀐 편성안에 맞게 다시 자동으로
            self.mark_edit(x)
            self.log(f"[편성안 바꿈] {day_label(x['date'])} {fmt_t(x['s'])} {x['name']} -> {x['plan']}")
        self.saved = False
        self.refresh(keep=self.sel)

    def shade_dialog(self):
        """고른 프로그램에 음영(배경색) 넣기: 이 줄만 / 이번 주 같은 프로그램 모두 / 매주 자동"""
        tk, ttk = self.tk, self.ttk
        from tkinter import colorchooser
        i = self.need_row()
        if i is None:
            return
        r = self.rows[i]
        top = tk.Toplevel(self.root)
        top.title("음영 넣기")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        tk.Label(f, text=f"{r['name']}", font=("", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Label(f, text="색").grid(row=1, column=0, sticky="w", pady=4)
        v_name = tk.StringVar(value="회색")
        v_col = tk.StringVar(value=SHADE_PRESETS["회색"])
        cb = ttk.Combobox(f, textvariable=v_name, values=list(SHADE_PRESETS), width=10, state="readonly")
        cb.grid(row=1, column=1, sticky="w")
        sw = tk.Label(f, text="      ", bg=v_col.get(), relief="solid", bd=1)
        sw.grid(row=1, column=3, padx=4)

        def set_col(c):
            v_col.set(c)
            sw.config(bg=c or "#ffffff")
        cb.bind("<<ComboboxSelected>>", lambda e: set_col(SHADE_PRESETS.get(v_name.get(), "")))

        def pick():
            c = colorchooser.askcolor(color=v_col.get() or "#d9d9d9", parent=top)[1]
            if c:
                set_col(c)
                v_name.set("직접 고른 색")
        tk.Button(f, text="색 고르기", command=pick).grid(row=1, column=2, padx=4)
        picked = [self.rows[k] for k in self.selected_rows()]
        scope = tk.StringVar(value="picked" if len(picked) > 1 else "same")
        same = [x for x in self.rows if (r.get("code") and x.get("code") == r.get("code"))
                or norm_key(split_suffix(x["name"])[0]) == norm_key(split_suffix(r["name"])[0])]
        for k, t in ((("picked", f"고른 칸 모두 ({len(picked)}칸)"),) if len(picked) > 1 else ()) + (
                     ("one", "이 줄만"), ("same", f"이번 주 같은 프로그램 모두 ({len(same)}줄)"),
                     ("rule", "매주 자동 (설정 파일에 규칙 저장)")):
            tk.Radiobutton(f, text=t, value=k, variable=scope).grid(columnspan=4, sticky="w")
        tk.Label(f, text="'없음'을 고르면 음영을 뺍니다.", fg="#777").grid(columnspan=4, sticky="w")

        def ok():
            c = v_col.get()
            self.push_undo("shade")
            if scope.get() == "rule":
                key = split_suffix(r["name"])[0].strip()
                add_shade_rule(key, c)
                self.log(f"[음영 규칙] '{key}' -> {c or '없음'} (매주 자동, 설정 파일에 저장)")
            targets = [r] if scope.get() == "one" else picked if scope.get() == "picked" else same
            for x in targets:
                if scope.get() == "rule":
                    x.pop("shade", None)          # 규칙이 적용되도록 줄 음영은 비움
                else:
                    x["shade"] = c
                self.mark_edit(x)
            if scope.get() != "rule":
                self.log(f"[음영] '{r['name']}' {len(targets)}줄 -> {c or '없음'}")
            top.destroy()
            self.refresh(keep=i)
        bf = tk.Frame(f)
        bf.grid(columnspan=4, pady=8)
        tk.Button(bf, text="적용", command=ok, width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="취소", command=top.destroy, width=10).pack(side=tk.LEFT, padx=6)
        top.after(150, lambda: top.grab_set())

    def revert_row(self):
        from tkinter import messagebox
        i = self.need_row()
        if i is None:
            return
        r = self.rows[i]
        o = next((o for o in self.orig if o["date"] == r["date"]
                  and o.get("code") in {r.get("code"), r.get("code_orig")}
                  and o["plan"] == r["plan"] and (o["s"] == r["s"] or o["name"] == r["name"])), None)
        if o is None:
            messagebox.showinfo("알림", "엑셀에서 같은 줄을 찾지 못했습니다. (새로 추가한 줄이면 [줄 삭제]를 쓰세요)")
            return
        self.push_undo("revert")
        self.rows[i] = dict(o)
        self.saved = False
        self.log(f"[되돌림] {day_label(o['date'])} {fmt_t(o['s'])} {o['name']}")
        self.refresh(keep=i)

    # ---------- 만들기 · 확인
    def prev_rows(self):
        p = self.v_prev.get().strip()
        if not p:
            return None
        try:
            return read_schedule_xlsx(p)
        except Exception as e:
            self.log(f"[주의] 지난주 엑셀을 읽지 못해 비교 없이 만듭니다: {e}")
            return None

    def build_temp(self):
        import tempfile
        out = os.path.join(tempfile.gettempdir(), "tv_schedule_preview.pdf")
        generate(self.out_rows(), out, self.prev_rows_cached(), log=lambda s: None, cfg_over=self.cfg_over())
        return out

    def ready(self):
        from tkinter import messagebox
        if not self.rows:
            messagebox.showinfo("알림", "먼저 [불러오기]로 엑셀을 여세요.")
            return False
        return True

    def make(self):
        from tkinter import filedialog, messagebox
        if not self.ready():
            return
        base = os.path.splitext(self.default_proj_path())[0]          # 저장하는 지금 시분
        pdf_dir = os.path.dirname(getattr(self, "last_pdf", "") or "") or \
            (os.path.dirname(self.xlsx or "") if (self.xlsx or "").lower().endswith(".xlsx") else "")
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                           initialfile=os.path.basename(base) + ".pdf",
                                           initialdir=pdf_dir if pdf_dir and os.path.isdir(pdf_dir) else None)
        if not out:
            return
        try:
            generate(self.out_rows(), out, self.prev_rows_cached(), self.log, cfg_over=self.cfg_over())
        except Exception as e:
            self.error("PDF 만들기 오류", e)
            return
        self.last_pdf = out
        # 같은 이름의 프로젝트도 데이터 폴더에 함께 저장 -> 나중에 다시 열어 고칠 수 있음
        self.save(path=os.path.join(week_dir(self.rows), os.path.splitext(os.path.basename(out))[0] + PROJ_EXT),
                  kept=True)
        self.publish_preview(out)       # 서버로 보내기 (프로젝트 저장 뒤 -> 서버가 백업할 때 이번 프로젝트도 들어감)
        self.log(f"[작업 기록] 불러온 뒤 {self.elapsed()} 걸림 · 고친 줄 {self.edits}개 "
                 f"(프로젝트도 함께 저장: {os.path.basename(self.proj_path or '')})")
        if messagebox.askyesno("완료", "PDF를 만들었습니다. 열어 볼까요?"):
            open_file(out)

    def preview_key(self):
        """서버 목록의 칸 이름: 주(월요일 날짜)마다 한 줄. 'schedule-엑셀-20261116' / 'dmb-엑셀-20261116'"""
        dates = sorted({r["date"] for r in self.rows})
        if not dates:
            return ""
        d = datetime.datetime.strptime(dates[0], "%Y%m%d")
        mon = d - datetime.timedelta(days=d.weekday())
        return ("dmb" if IS_DMB else "schedule") + f"-엑셀-{mon:%Y%m%d}"

    def preview_title(self):
        t = (self.v_title.get() or "").strip() or week_label(sorted({r["date"] for r in self.rows}))
        return t

    def publish_preview(self, pdf_path):
        """PDF 저장할 때: 미리보기 폴더에 PDF를 두고 '수정완료'로 -> 녹음 PC 서버가 1분 안에 가져감"""
        d, key = tv_preview_dir(), self.preview_key()
        if not d or not key:
            return
        try:
            os.makedirs(d, exist_ok=True)
            shutil.copyfile(pdf_path, os.path.join(d, key + ".pdf.tmp"))
            os.replace(os.path.join(d, key + ".pdf.tmp"), os.path.join(d, key + ".pdf"))
        except Exception as e:
            self.log(f"[서버로 보내기] PDF를 미리보기 폴더에 두지 못했습니다: {e}")
            return
        update_tv_preview_status(key, kind="엑셀 DMB" if IS_DMB else "엑셀 TV", title=self.preview_title(),
                                 state="수정완료", dirty=False, src=os.path.basename(pdf_path),
                                 saved_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._pv_dirty_key = ""
        on = tv_view_config()["on"]
        self.log(f"[서버로 보내기] {self.preview_title()} - "
                 + ("1분 안에 녹음 PC 서버에 올라갑니다 (프로젝트·즐겨찾기·설정 백업도 함께)" if on else
                    "'서버로 보내기'가 꺼져 있어 이 PC에만 둡니다 (⚙ 설정에서 켜기)"))

    def mark_preview_dirty(self):
        """PDF 저장 뒤 다시 고치기 시작하면 서버 목록에 '✏ 현재 편집중' (바뀔 때만 파일을 고침)"""
        if not tv_preview_dir() or not self.rows or self.saved:
            return
        key = self.preview_key()
        if not key or getattr(self, "_pv_dirty_key", "") == key:
            return
        if not os.path.exists(os.path.join(tv_preview_dir(), key + ".pdf")):
            return          # 아직 이 주를 서버에 보낸 적 없으면 그대로
        self._pv_dirty_key = key
        update_tv_preview_status(key, dirty=True)

    def preview(self):
        """큰 창 미리보기: 실제로 저장되는 PDF 그대로. 확대·축소, 위아래좌우 이동"""
        if not self.ready():
            return
        try:
            pdf = self.build_temp()
            with open(pdf, "rb") as f:
                data = f.read()
        except Exception as e:
            self.error("미리보기 오류", e)
            return
        PreviewWindow(self, data)

    def compare_original(self):
        from tkinter import filedialog
        if not self.ready():
            return
        base = filedialog.askopenfilename(title="기준이 될 원본(완성본) 편성표 PDF", filetypes=[("PDF", "*.pdf")])
        if not base:
            return
        try:
            res = compare_pdfs(base, self.build_temp(), self.log)
        except Exception as e:
            self.error("비교 오류", e)
            return
        self.diff_keys = set(res.get("다른 칸", []))
        self.draw_marks()
        if self.diff_keys:
            self.log(f"[비교] 원본과 다른 칸 {len(self.diff_keys)}개를 편성표에 주황 테두리로 표시했습니다.")

    def compare_two(self):
        from tkinter import filedialog
        a = filedialog.askopenfilename(title="1) 기준 PDF (완성본/원본)", filetypes=[("PDF", "*.pdf")])
        if not a:
            return
        b2 = filedialog.askopenfilename(title="2) 비교할 PDF (편집기로 고친 것 등)", filetypes=[("PDF", "*.pdf")])
        if not b2:
            return
        try:
            compare_pdfs(a, b2, self.log)
        except Exception as e:
            self.error("비교 오류", e)

    # ---------- 기타
    def project_data(self):
        cfg = load_config()
        return {"source": self.xlsx or "", "query_date": getattr(self, "query_date", ""),
                "rows": self.rows, "orig": self.orig, "deleted": self.deleted,
                "settings": {k: cfg.get(k) for k in PROJ_CFG_KEYS}, "last_pdf": getattr(self, "last_pdf", ""),
                "history": self.history, **({"tv_project": self.tv_path} if IS_DMB else {})}

    def default_proj_path(self):
        """저장 이름: 엑셀이름_편성표 + 저장하는 지금 시분  (예: MBC_11월_3주(2)_편성표1605.tvproj)
        폴더는 늘 데이터 폴더(PROJ_DIR, 런처: C:\\EDIT\\TV\\데이터\\엑셀편성표)"""
        name = os.path.splitext(os.path.basename(self.proj_path or self.xlsx or "편성표"))[0]
        name = re.sub(r"_편성표\d{0,4}$", "", name)       # 예전 시분은 떼고 지금 시분으로
        if IS_DMB:          # DMB: DMB_MBC_11월_3주(3)_편성표1830.dmbproj
            name = "DMB_" + re.sub(r"^DMB_", "", name)
        return os.path.join(week_dir(self.rows), name + "_편성표" + datetime.datetime.now().strftime("%H%M") + PROJ_EXT)

    def save(self, ask=False, path=None, kept=False):
        """프로젝트 저장 (Ctrl+S): 지금 시분이 붙은 새 이름으로 저장 (예전 것은 기록으로 남음).
        처음 저장이거나 '다른 이름으로'면 저장 위치를 물음"""
        from tkinter import filedialog
        if not self.ready():
            return False
        p = path
        if not p and not ask:
            p = self.default_proj_path()        # 데이터 폴더, 지금 시분 (같은 분이면 덮어씀)
        if not p:
            d = self.default_proj_path()
            p = filedialog.asksaveasfilename(defaultextension=PROJ_EXT, filetypes=[("편성표 프로젝트", "*" + PROJ_EXT)],
                                             initialfile=os.path.basename(d), initialdir=os.path.dirname(d) or None)
            if not p:
                return False
        if not p.lower().endswith(PROJ_EXT):        # 엑셀·PDF 위에 덮어쓰지 않도록
            p = os.path.splitext(p)[0] + PROJ_EXT
        if not kept and os.path.exists(p):           # PDF 로 낸 프로젝트를 같은 이름으로 덮으면 '남김' 표시는 그대로
            try:
                with open(p, "r", encoding="utf-8") as f:
                    kept = bool(json.load(f).get("kept"))
            except Exception:
                pass
        try:
            save_project(p, dict(self.project_data(), kept=True) if kept else self.project_data())
        except Exception as e:
            self.error("프로젝트 저장 오류", e)
            return False
        self.proj_path = p
        self.saved = True
        self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(p)}")
        try:
            tidy_projects(log=self.log)
        except Exception:
            pass
        self.build_recent()
        self.log(f"[프로젝트 저장] {p}")
        return True

    def open_project(self, path=None):
        """프로젝트(.tvproj) 열기: 고친 상태 그대로 이어서 편집 (엑셀 파일이 없어도 됨)"""
        from tkinter import filedialog, messagebox
        if self.xlsx and not self.saved:
            ans = messagebox.askyesnocancel("프로젝트 열기", "지금 고친 내용을 저장하지 않았습니다.\n저장하고 열까요?")
            if ans is None or (ans and not self.save()):
                return
        if not path:
            path = filedialog.askopenfilename(filetypes=[("편성표 프로젝트", "*" + PROJ_EXT)] +
                                              ([("TV 프로젝트", "*" + TV_PROJ_EXT)] if IS_DMB else []),
                                              initialdir=latest_week_dir())
            if not path:
                return
        if IS_DMB and path.lower().endswith(TV_PROJ_EXT):
            return self.import_tv(path, ask_save=False)
        if not IS_DMB and path.lower().endswith(DMB_PROJ_EXT):
            from tkinter import messagebox
            messagebox.showinfo("프로젝트 열기", "DMB 프로젝트입니다. 엑셀 DMB 편성표 편집기에서 여세요.")
            return
        try:
            d = load_project(path)
        except Exception as e:
            self.error("프로젝트 열기 오류", e)
            return
        st = {k: v for k, v in (d.get("settings") or {}).items() if v is not None}
        if st:
            save_config_keys(st)
            self._loading = True
            for var, k in ((self.v_title, "title_text"), (self.v_tsize, "title_size"), (self.v_smode, "stamp_mode"),
                           (self.v_stext, "stamp_text")):
                if k in st:
                    var.set(str(st[k]))
            if "title_bold" in st:
                self.v_tbold.set(bool(st["title_bold"]))
            self._loading = False
        self.rows = d["rows"]
        self.orig = d.get("orig") or json.loads(json.dumps(d["rows"]))
        self.deleted = d.get("deleted", 0)
        self.xlsx = d.get("source") or path
        self.query_date = d.get("query_date", "")
        self.last_pdf = d.get("last_pdf", "")
        self.proj_path = path
        self.proj_dir = ""
        self.history = d.get("history") or []
        self.tv_path = d.get("tv_project", "")
        self.sel, self.multi, self.diff_keys = None, [], set()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_buttons()
        self.t0 = datetime.datetime.now().timestamp()
        self.saved = True
        self.show_main()
        self.build_day_buttons()
        self.refresh()
        self.root.after(300, self.apply_saved_zoom)
        self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(path)}")
        self.build_recent()
        self.log(f"[프로젝트 열기] {os.path.basename(path)}  편성 {len(self.rows)}줄 · 고친 줄 {self.edits}개 "
                 f"(저장 {d.get('saved', '')[:16].replace('T', ' ')})")

    # ---------- DMB 편집기: TV 프로젝트 불러오기 / 다시 가져오기
    def import_tv(self, path=None, ask_save=True):
        """엑셀 TV 편성표 편집기에서 저장한 프로젝트(.tvproj)로 DMB 작업 시작 (TV 프로젝트는 건드리지 않음)"""
        from tkinter import filedialog, messagebox
        if ask_save and self.rows and not self.saved:
            ans = messagebox.askyesnocancel("TV 프로젝트 불러오기", "지금 고친 DMB 작업을 저장하지 않았습니다.\n저장하고 불러올까요?")
            if ans is None or (ans and not self.save()):
                return
        if not path:
            path = filedialog.askopenfilename(title="TV 프로젝트 (엑셀 TV 편성표 편집기에서 저장한 .tvproj)",
                                              filetypes=[("TV 프로젝트", "*" + TV_PROJ_EXT)],
                                              initialdir=latest_week_dir())
            if not path:
                return
        try:
            d = load_project(path, recent=False)
        except Exception as e:
            self.error("TV 프로젝트 열기 오류", e)
            return
        st = {k: v for k, v in (d.get("settings") or {}).items() if v is not None}
        if st.get("title_text"):
            st["title_text"] = dmb_title_auto(st["title_text"])     # 광주MBC... -> 광주 MYMBC...
        if st:
            save_config_keys(st)
            self._loading = True
            for var, k in ((self.v_title, "title_text"), (self.v_tsize, "title_size"), (self.v_smode, "stamp_mode"),
                           (self.v_stext, "stamp_text")):
                if k in st:
                    var.set(str(st[k]))
            if "title_bold" in st:
                self.v_tbold.set(bool(st["title_bold"]))
            self._loading = False
        self.rows = import_tv_rows(d["rows"])
        self.orig = import_tv_rows(d["rows"])
        self.tv_path = path
        self.xlsx = path
        self.query_date = d.get("query_date", "")
        self.last_pdf = ""
        self.deleted = 0
        self.proj_path = None
        self.proj_dir = os.path.dirname(path)
        self.history = [{"tv": os.path.basename(path), "at": datetime.datetime.now().isoformat(timespec="seconds"),
                         "what": "TV 프로젝트 불러오기"}]
        self.sel, self.multi, self.diff_keys = None, [], set()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_buttons()
        self.t0 = datetime.datetime.now().timestamp()
        self.saved = False
        self.show_main()
        self.build_day_buttons()
        self.refresh()
        self.root.after(300, self.apply_saved_zoom)
        self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(path)} (DMB 작업, 저장 전)")
        self.log(f"[TV 프로젝트 불러오기] {os.path.basename(path)}  편성 {len(self.rows)}줄 - "
                 f"TV 프로젝트는 그대로 두고, DMB 작업은 [프로젝트 저장] 때 {DMB_PROJ_EXT} 로 따로 저장합니다.")

    def fill_tv_menu(self, menu, action):
        """TV 프로젝트 고르기 메뉴: 최근 TV 프로젝트(새 것이 위) + 다른 파일 찾기"""
        menu.delete(0, "end")
        lst = tv_project_choices()
        for p in lst:
            menu.add_command(label=project_info(p), command=lambda p=p: action(p))
        if not lst:
            menu.add_command(label="(TV 프로젝트가 없습니다 - 엑셀 TV 편성표 편집기에서 저장하세요)", state="disabled")
        menu.add_separator()
        menu.add_command(label="다른 파일 찾기 …", command=lambda: action(None))

    def refresh_tv(self, path=None):
        """TV 프로젝트가 바뀐 뒤(새 판 반영 등): DMB 에서 고친 것은 두고 바뀐 TV 편성만 다시 가져오기"""
        from tkinter import filedialog, messagebox
        if not self.ready():
            return
        if not path:
            path = filedialog.askopenfilename(title="다시 가져올 TV 프로젝트 (.tvproj)",
                                              filetypes=[("TV 프로젝트", "*" + TV_PROJ_EXT)],
                                              initialdir=latest_week_dir())
            if not path:
                return
        try:
            d = load_project(path, recent=False)
        except Exception as e:
            self.error("TV 프로젝트 열기 오류", e)
            return
        od, nd = sorted({r["date"] for r in self.orig}), sorted({r["date"] for r in d["rows"]})
        if od != nd and not messagebox.askyesno("TV 다시 가져오기", "날짜가 다릅니다. 같은 주의 TV 프로젝트가 아닌 것 같습니다.\n"
                                                              "그래도 가져올까요?"):
            return
        old_tv = [{k: v for k, v in r.items() if k != "_tvk"} for r in self.orig]
        new_tv = [{k: v for k, v in r.items() if k not in MARK_KEYS} for r in d["rows"]]
        merged, rep = refresh_from_tv(self.rows, old_tv, new_tv)
        msg = (f"바뀜 {len(rep['changed'])} · 새로 {len(rep['new'])} · 없어짐 {len(rep['removed'])} · "
               f"DMB 고친 것 그대로 {rep['kept']}줄 · 확인 필요 {len(rep['checks'])}")
        if not rep["changed"] and not rep["new"] and not rep["removed"]:
            messagebox.showinfo("TV 다시 가져오기", "TV 편성에서 바뀐 것이 없습니다.")
            self.tv_path = path
            return
        if not messagebox.askyesno("TV 다시 가져오기", msg + "\n\n반영할까요? (바뀐 칸은 주황, 확인할 칸은 빨간 테두리)"):
            return
        self.push_undo()
        self.rows = merged
        self.orig = import_tv_rows(new_tv)
        self.tv_path = self.xlsx = path
        self.query_date = d.get("query_date", self.query_date)
        self.history = list(self.history) + [{"tv": os.path.basename(path), "what": "TV 다시 가져오기", "summary": msg,
                                              "at": datetime.datetime.now().isoformat(timespec="seconds")}]
        self.sel, self.multi = None, []
        self.saved = False
        self.build_day_buttons()
        self.refresh()
        self.log(f"[TV 다시 가져오기] {os.path.basename(path)}  {msg}")
        for x in rep["removed"]:
            self.log(f"   - 없어짐: {day_label(x['date'])} {fmt_t(x['s'])}~{fmt_t(x['e'])} {x['name']}")
        for x in rep["checks"]:
            self.log(f"   ⚠ {day_label(x['date'])} {fmt_t(x['s'])}~{fmt_t(x['e'])} {x['name']} - {x.get('_note', '')}")

    def build_recent(self):
        """첫 화면의 최근 프로젝트 목록 + 위쪽 [최근 ▼] 메뉴"""
        tk = self.tk
        lst = recent_projects()
        try:
            self.recent_menu.delete(0, "end")
            for p in lst:
                self.recent_menu.add_command(label=project_info(p), command=lambda p=p: self.open_project(p))
            if not lst:
                self.recent_menu.add_command(label="(최근 프로젝트 없음)", state="disabled")
        except Exception:
            pass
        box = getattr(self, "recent_box", None)
        if box is None:
            return
        for w_ in box.winfo_children():
            w_.destroy()
        if not lst:
            return
        tk.Button(box, text=f"▶ 마지막 프로젝트 이어서 편집:  {os.path.basename(lst[0])}", bg="#4CAF50", fg="white",
                  font=("", 11, "bold"), padx=10, pady=4, command=lambda: self.open_project(lst[0])).pack(pady=(0, 10))
        tk.Label(box, text="최근 프로젝트", bg="#e4e6ea", fg="#666", font=("", 9, "bold")).pack(anchor="w")
        for p in lst[:6]:
            tk.Button(box, text=project_info(p), anchor="w", relief="flat", bg="#f4f5f7", width=70,
                      command=lambda p=p: self.open_project(p)).pack(fill=tk.X, pady=1)

    def learn(self):
        global _name_dict
        from tkinter import filedialog, messagebox
        if not self.xlsx:
            messagebox.showwarning("알림", "먼저 엑셀을 불러오세요. 같은 주의 원본 PDF와 비교합니다.")
            return
        pdf = filedialog.askopenfilename(title="같은 주의 원본 편성표 PDF", filetypes=[("PDF", "*.pdf")])
        if not pdf:
            return
        try:
            _name_dict = None
            learn_names(self.xlsx, pdf, self.log)
        except Exception as e:
            self.error("이름 배우기 오류", e)

    def open_config_folder(self):
        """설정 파일이 있는 폴더를 탐색기로 (백업·복사할 때)"""
        load_config()
        open_file(os.path.dirname(CONFIG_PATH))
        self.log(f"[설정 폴더] {os.path.dirname(CONFIG_PATH)}")

    def on_close(self):
        from tkinter import messagebox
        if self.xlsx and not self.saved:
            ans = messagebox.askyesnocancel("종료 확인", "고친 내용을 프로젝트로 저장하지 않았습니다.\n저장하고 끝낼까요?")
            if ans is None:
                return
            if ans and not self.save():
                return
        elif not messagebox.askokcancel("종료 확인", "프로그램을 종료할까요?"):
            return
        if getattr(self, "_zoom_save_id", None):    # 끄기 직전에 바꾼 보기 크기도 저장
            self.save_zoom_pref()
        self.root.destroy()

    # ================================================================ 편성표 보면서 고치기
    @property
    def edits(self):
        return sum(1 for r in self.rows if r.get("edited")) + self.deleted

    def build_form(self, parent):
        tk, ttk = self.tk, self.ttk
        self.form_title = tk.Label(parent, text="칸을 고르세요", font=("", 12, "bold"), anchor="w")
        self.form_title.pack(fill=tk.X, pady=(2, 2))
        # 묶인 칸(월~금): 지금 고치는 범위 [묶음 전체] / [이 요일만] - 묶인 칸을 골랐을 때만 보임
        self.scope_bar = tk.Frame(parent)
        self.scope_all_btn = tk.Button(self.scope_bar, text="묶음 전체", width=10,
                                       command=lambda: self.valid_sel() and self.select_block(self.sel))
        self.scope_all_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.scope_one_btn = tk.Button(self.scope_bar, text="이 요일만", width=10,
                                       command=lambda: self.valid_sel() and self.select(self.sel))
        self.scope_one_btn.pack(side=tk.LEFT)
        tk.Label(self.scope_bar, text="(같은 칸을 한 번 더 누르면 그 요일만)", fg="#777").pack(side=tk.LEFT, padx=6)
        self._btn_bg = self.scope_all_btn.cget("bg")
        # 주황·빨강 표시가 있는 칸: 이유 한 줄 + [겹침 정리]
        self.reason_bar = tk.Frame(parent)
        self.reason_lbl = tk.Label(self.reason_bar, text="", anchor="w", justify="left", wraplength=380,
                                   font=("", 9, "bold"))
        self.reason_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.reason_btn = tk.Button(self.reason_bar, text="겹침 정리", bg="#ffcdd2",
                                    command=lambda: self.ask_overlap(self.sel, force=True))
        f = tk.Frame(parent)
        f.pack(fill=tk.X)
        self.fv = {k: tk.StringVar() for k in ("name", "s", "e", "plan", "age", "rl", "code", "shade_name",
                                               "hl_name", "fg_name", "show")}
        self.f_hl = self.f_fg = ""
        self.fb = {k: tk.BooleanVar() for k in ("uhd", "cc", "sign", "ad", "force_pink", "show_end", "bold",
                                                "top_dash", "bottom_dash")}
        self.f_shade = ""
        tk.Label(f, text="프로그램명").grid(row=0, column=0, sticky="w")
        e_name = tk.Entry(f, textvariable=self.fv["name"], width=34)
        e_name.grid(row=0, column=1, columnspan=3, sticky="we", pady=2)
        bind_linebreak(e_name)
        tk.Label(f, text="줄바꿈: | 또는 Shift+Enter  (예: 특집 다큐|선비의 정신 필암서원)", fg="#777",
                 font=("", 8)).grid(row=10, column=0, columnspan=4, sticky="w")
        self.v_name_all = tk.BooleanVar(value=False)
        tk.Checkbutton(f, text="여러 칸 선택 시 이름도 모두 똑같이", variable=self.v_name_all,
                       font=("", 8)).grid(row=11, column=0, columnspan=4, sticky="w")
        tk.Label(f, text="시작").grid(row=1, column=0, sticky="w")
        self.ent_s = tk.Entry(f, textvariable=self.fv["s"], width=7)
        self.ent_s.grid(row=1, column=1, sticky="w")
        tk.Label(f, text="끝").grid(row=1, column=2, sticky="e")
        self.ent_e = tk.Entry(f, textvariable=self.fv["e"], width=7)
        self.ent_e.grid(row=1, column=3, sticky="w")
        for ent in (self.ent_s, self.ent_e):
            ent.bind("<Return>", lambda e: self.ask_overlap())
            ent.bind("<FocusOut>", lambda e: self.ask_overlap())
        tk.Label(f, text="편성안").grid(row=2, column=0, sticky="w")
        ttk.Combobox(f, textvariable=self.fv["plan"], values=("1안", "2안"), width=6,
                     state="readonly").grid(row=2, column=1, sticky="w", pady=2)
        tk.Label(f, text="연령").grid(row=2, column=2, sticky="e")
        ttk.Combobox(f, textvariable=self.fv["age"], values=("", "00", "07", "12", "15", "19"), width=6,
                     state="readonly").grid(row=2, column=3, sticky="w")
        tk.Label(f, text="R/L").grid(row=3, column=0, sticky="w")
        ttk.Combobox(f, textvariable=self.fv["rl"], values=("", "R", "V", "L"), width=6,
                     state="readonly").grid(row=3, column=1, sticky="w", pady=2)
        tk.Label(f, text="코드").grid(row=3, column=2, sticky="e")
        tk.Entry(f, textvariable=self.fv["code"], width=12).grid(row=3, column=3, sticky="w")
        cf = tk.Frame(f)
        cf.grid(row=4, column=0, columnspan=4, sticky="w", pady=2)
        for k, t in (("uhd", "UHD"), ("cc", "자막"), ("sign", "수어"), ("ad", "화면해설")):
            tk.Checkbutton(cf, text=t, variable=self.fb[k]).pack(side=tk.LEFT)
        cf2 = tk.Frame(f)
        cf2.grid(row=5, column=0, columnspan=4, sticky="w")
        tk.Checkbutton(cf2, text="분홍", variable=self.fb["force_pink"]).pack(side=tk.LEFT)
        tk.Checkbutton(cf2, text="끝 시각 표시", variable=self.fb["show_end"]).pack(side=tk.LEFT)
        tk.Checkbutton(cf2, text="굵게", variable=self.fb["bold"]).pack(side=tk.LEFT)
        tk.Checkbutton(cf2, text="위 선 점선", variable=self.fb["top_dash"]).pack(side=tk.LEFT)
        tk.Checkbutton(cf2, text="아래 선 점선", variable=self.fb["bottom_dash"]).pack(side=tk.LEFT)
        tk.Label(f, text="음영").grid(row=6, column=0, sticky="w")
        ttk.Combobox(f, textvariable=self.fv["shade_name"], values=list(SHADE_PRESETS), width=10,
                     state="readonly").grid(row=6, column=1, sticky="w", pady=2)
        tk.Button(f, text="색 고르기", command=self.form_pick_color).grid(row=6, column=2, columnspan=2, sticky="w")
        tk.Label(f, text="글자 바탕").grid(row=7, column=0, sticky="w")
        ttk.Combobox(f, textvariable=self.fv["hl_name"], values=list(HL_PRESETS), width=10,
                     state="readonly").grid(row=7, column=1, sticky="w", pady=2)
        tk.Label(f, text="글자색").grid(row=7, column=2, sticky="e")
        ttk.Combobox(f, textvariable=self.fv["fg_name"], values=list(FG_PRESETS), width=8,
                     state="readonly").grid(row=7, column=3, sticky="w")
        tk.Button(f, text="바탕색 고르기", command=lambda: self.form_pick_color("hl")).grid(row=8, column=1, sticky="w")
        tk.Button(f, text="글자색 고르기", command=lambda: self.form_pick_color("fg")).grid(row=8, column=2, columnspan=2,
                                                                                    sticky="w")
        tk.Button(f, text="★ 이 칸을 즐겨찾기에 저장 (이름 + 음영·글자 바탕·아이콘 등 옵션)", bg="#fff8e1",
                  command=self.fav_add_from_cell).grid(row=9, column=0, columnspan=4, sticky="we", pady=(4, 0))
        self.form_msg = tk.Label(parent, text="", fg="#c62828", anchor="w")
        self.form_msg.pack(fill=tk.X)
        bf = tk.Frame(parent)
        bf.pack(fill=tk.X, pady=4)
        tk.Button(bf, text="같은 프로그램 모두에 적용", command=self.apply_same, bg="#e3f2fd").pack(side=tk.LEFT)
        tk.Button(bf, text="앞뒤 글자 ...", command=self.affix_dialog).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text=("TV 원래대로" if IS_DMB else "엑셀 원래대로"), command=self.revert_row).pack(
            side=tk.LEFT, padx=4)
        bf2 = tk.Frame(parent)
        bf2.pack(fill=tk.X)
        tk.Button(bf2, text="뒤에 줄 추가", command=self.add_row).pack(side=tk.LEFT)
        tk.Button(bf2, text="줄 삭제", command=self.del_row).pack(side=tk.LEFT, padx=4)
        tk.Button(bf2, text="자세히 고치기", command=lambda: self.edit_row()).pack(side=tk.LEFT)
        tk.Button(parent, text="로컬 끼워 넣기 ...", command=self.local_dialog, bg="#cfd8dc").pack(fill=tk.X, pady=(6, 0))
        # 오른쪽 아래: 탭 [즐겨찾기 | 그날 편성 줄]
        rnb = ttk.Notebook(parent)
        rnb.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.rnb = rnb
        tab_fav = tk.Frame(rnb)
        tab_day = tk.Frame(rnb)
        rnb.add(tab_fav, text="  즐겨찾기  ")
        rnb.add(tab_day, text="  그날 편성 줄  ")
        fh = tk.Frame(tab_fav)
        fh.pack(fill=tk.X, pady=(4, 2))
        tk.Label(fh, text="☆ 누르면 고른 칸에 넣기", fg="#777", font=("", 8)).pack(side=tk.LEFT)
        tk.Button(fh, text="+ 추가", command=self.fav_add_dialog).pack(side=tk.RIGHT)
        tk.Button(fh, text="고른 칸 추가", command=self.fav_add_from_cell).pack(side=tk.RIGHT, padx=2)
        tk.Button(fh, text="창으로 ⧉", command=self.open_fav_window).pack(side=tk.RIGHT, padx=2)
        favf = tk.Frame(tab_fav, relief="sunken", bd=1)
        favf.pack(fill=tk.BOTH, expand=True)
        self.favf = favf
        # 즐겨찾기를 창으로 열면 여기 목록은 숨기고 안내만 (목록이 두 군데 나오지 않게)
        self.fav_note = tk.Frame(tab_fav)
        tk.Label(self.fav_note, text="즐겨찾기를 따로 창에서 보는 중입니다.", fg="#666").pack(pady=(20, 6))
        tk.Button(self.fav_note, text="여기로 되돌리기", command=self.close_fav_window).pack()
        self.fav_cv = tk.Canvas(favf, height=180, highlightthickness=0, bg="#f7f7f7")
        fsb = tk.Scrollbar(favf, orient=tk.VERTICAL, command=self.fav_cv.yview)
        self.fav_cv.configure(yscrollcommand=fsb.set)
        fsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.fav_cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.fav_frame = tk.Frame(self.fav_cv, bg="#f7f7f7")
        self._fav_win = self.fav_cv.create_window(0, 0, window=self.fav_frame, anchor="nw")
        self.fav_frame.bind("<Configure>", lambda e: self.fav_cv.configure(scrollregion=self.fav_cv.bbox("all")))
        self.fav_cv.bind("<Configure>", lambda e: self.fav_cv.itemconfigure(self._fav_win, width=e.width))
        self.fav_cv.bind("<MouseWheel>", lambda e: self.fav_cv.yview_scroll(int(-e.delta / 120), "units"))
        self.build_fav_buttons()
        self.dl_title = tk.Label(tab_day, text="편성표에서 고른 칸의 요일", fg="#777", font=("", 8), anchor="w")
        self.dl_title.pack(fill=tk.X, pady=(4, 2))
        self.daylist = ttk.Treeview(tab_day, columns=("s", "e", "name", "plan"), show="headings", height=14)
        for k, t, w in (("s", "시작", 48), ("e", "끝", 48), ("name", "프로그램명", 210), ("plan", "안", 36)):
            self.daylist.heading(k, text=t)
            self.daylist.column(k, width=w, anchor="w" if k == "name" else "center")
        self.daylist.tag_configure("edited", foreground="#c62828")
        self.daylist.pack(fill=tk.BOTH, expand=True)
        self.daylist.bind("<<TreeviewSelect>>", self.on_daylist_select)
        for v in list(self.fv.values()) + list(self.fb.values()):
            v.trace_add("write", lambda *a: self.form_changed())

    def valid_sel(self):
        return self.sel is not None and 0 <= self.sel < len(self.rows)

    def select(self, i, source=None, keep_multi=False):
        """줄 하나를 고름: 편성표 테두리, 오른쪽 수정란, 표 선택을 함께 맞춤"""
        self.sel = i if (i is not None and 0 <= i < len(self.rows)) else None
        self.empty_slot = None
        if not keep_multi:
            self.multi = [self.sel] if self.sel is not None else []
        self.multi = [k for k in self.multi if 0 <= k < len(self.rows)]
        if self.sel is not None and self.sel not in self.multi:
            self.multi.append(self.sel)
        self.load_form()
        self.fill_daylist()
        if source != "tree" and self.valid_sel() and self.tree.exists(str(self.sel)):
            self._quiet = True
            self.tree.selection_set([str(k) for k in self.multi if self.tree.exists(str(k))])
            self.tree.see(str(self.sel))
            self._quiet = False
        self.draw_marks()

    def whole_block(self):
        """지금 고른 칸들이 월~금(주말)으로 묶인 한 칸 전체이면 그 줄 번호들, 아니면 None"""
        sel = set(self.selected_rows())
        if len(sel) < 2 or not self.valid_sel():
            return None
        blk = self.block_of(self.sel)
        if blk and set(blk["_ris"]) == sel:
            return sorted(sel)
        return None

    def select_block(self, i):
        """칸 하나를 누름: 월~금으로 묶인 칸이면 묶인 요일 전체를 한 칸처럼 고름"""
        blk = self.block_of(i)
        if blk and len(blk["_ris"]) > 1:
            self.multi = [k for k in blk["_ris"] if 0 <= k < len(self.rows)]
            self.select(i, keep_multi=True)
        else:
            self.select(i)

    def selected_rows(self):
        """선택한 줄 번호들 (여러 칸 선택이면 전부, 아니면 지금 줄 하나)"""
        rows = [k for k in self.multi if 0 <= k < len(self.rows)]
        if not rows and self.valid_sel():
            rows = [self.sel]
        return rows

    def clear_selection(self):
        self.multi = []
        self.select(None)

    def load_form(self):
        self._loading = True
        try:
            if not self.valid_sel():
                self.form_title.config(text="칸을 고르세요", fg="black")
                self.update_scope_bar()
                for v in self.fv.values():
                    v.set("")
                for v in self.fb.values():
                    v.set(False)
                return
            r = self.rows[self.sel]
            self.load_form_title()
            for k in ("name", "plan", "age", "rl", "code"):
                self.fv[k].set(r.get(k, "") or "")
            self.fv["show"].set(SHOW_LABELS.get(r.get("show") or "", "모두"))
            self.fv["s"].set(fmt_t(r["s"]))
            self.fv["e"].set(fmt_t(r["e"]))
            for k in self.fb:
                self.fb[k].set(bool(r.get(k)))
            self.fb["force_pink"].set(self.eff_pink(self.sel))
            self.f_shade = r.get("shade", "") or ""
            name = next((k for k, c in SHADE_PRESETS.items() if c == self.f_shade), "직접 고른 색")
            self.fv["shade_name"].set(name)
            self.f_hl, self.f_fg = r.get("hl", "") or "", r.get("fg", "") or ""
            self.fv["hl_name"].set(next((k for k, c in HL_PRESETS.items() if c == self.f_hl), "직접 고른 색"))
            self.fv["fg_name"].set(next((k for k, c in FG_PRESETS.items() if c == self.f_fg), "직접 고른 색"))
            self.form_msg.config(text="")
        finally:
            self._loading = False

    def form_changed(self):
        if self._loading:
            return
        if not self.valid_sel():
            name = self.fv["name"].get().strip()
            if not (self.empty_slot and name) or self.new_row_at_slot(name) is None:
                return
        r = self.rows[self.sel]
        try:
            s_ = hhmm_to_min(self.fv["s"].get()) - DAY_START_MIN
            e_ = hhmm_to_min(self.fv["e"].get()) - DAY_START_MIN
            if e_ <= s_ or not (0 <= s_ < DAY_LEN_MIN and 0 < e_ <= DAY_LEN_MIN):
                raise ValueError
        except (ValueError, TypeError):
            self.form_msg.config(text="시각을 확인하세요 (예: 07:50, 자정 넘으면 25:30)")
            return
        name = self.fv["name"].get().strip()
        if not name:
            self.form_msg.config(text="프로그램명이 비어 있습니다")
            return
        self.form_msg.config(text="")
        sn = self.fv["shade_name"].get()
        if sn in SHADE_PRESETS:
            self.f_shade = SHADE_PRESETS[sn]
        if self.fv["hl_name"].get() in HL_PRESETS:
            self.f_hl = HL_PRESETS[self.fv["hl_name"].get()]
        if self.fv["fg_name"].get() in FG_PRESETS:
            self.f_fg = FG_PRESETS[self.fv["fg_name"].get()]
        new = dict(name=name, s=s_, e=e_, plan=self.fv["plan"].get() or "1안", age=self.fv["age"].get(),
                   rl=self.fv["rl"].get(), code=self.fv["code"].get().strip(), shade=self.f_shade,
                   hl=self.f_hl, fg=self.f_fg, show=SHOW_KEYS.get(self.fv["show"].get(), ""),
                   **{k: v.get() for k, v in self.fb.items()})
        fpv = new.pop("force_pink")
        changed = any((r.get(k) or "") != (v or "") if not isinstance(v, bool) else bool(r.get(k)) != v
                      for k, v in new.items())
        if fpv != self.eff_pink(self.sel):
            new["force_pink"] = fpv
            changed = True
        if not changed:
            return
        self.push_undo(("form", tuple(self.selected_rows())))  # 같은 칸을 이어서 고치는 동안은 한 번만 기록
        old_shows = {id(self.rows[k]): self.rows[k].get("show") or "" for k in self.selected_rows()}
        # 여러 칸 선택: 바뀐 항목 중 아이콘·음영·분홍·끝 시각·굵게·편성안은 고른 칸 모두에,
        # 프로그램명·시각·코드는 기준 칸(마지막에 누른 칸)에만
        others = [k for k in self.selected_rows() if k != self.sel]
        if others:
            per_row = ("s", "e", "code") if self.v_name_all.get() else ("name", "s", "e", "code")
            if self.whole_block():      # 묶인 한 칸: 이름·시각·코드도 묶인 요일 모두 똑같이
                per_row = ()
            diff = {k: v for k, v in new.items() if k not in per_row and
                    ((r.get(k) or "") != (v or "") if not isinstance(v, bool) else bool(r.get(k)) != v)}
            for k in others:
                x = self.rows[k]
                x.update(diff)
                self.mark_edit(x)
            if "name" not in per_row and name != r["name"]:
                for k in others:
                    self.rows[k]["name_edited"] = True
            if any(new[k] != r.get(k) for k in per_row if k in new) and not self._multi_warned:
                self.log("[여러 칸] 프로그램명·시각·코드는 기준 칸에만 적용됩니다.")
                self._multi_warned = True
        if name != r["name"]:
            r["name_edited"] = True
        time_changed = (s_, e_) != (r["s"], r["e"])
        r.update(new)
        if others and self.v_name_all.get():     # 여러 칸에 같은 이름 -> 같은 코드로 한 칸에 묶이게
            code = share_code(self.rows, [r] + [self.rows[k] for k in others])
            if code:
                self._loading = True
                self.fv["code"].set(code)
                self._loading = False
                self.log(f"[묶음 코드] {len(others) + 1}칸에 같은 코드 {code} 붙임")
        if any((x.get("show") or "") != old_shows.get(id(x), "") for x in [r] + [self.rows[k] for k in others]):
            for x in [r] + [self.rows[k] for k in others]:
                self.after_show_change(x, old_shows.get(id(x), ""))
            self.sel = self.rows.index(r)
            self.multi = []
            self.refresh(keep=self.sel)
            return
        if time_changed:        # 입력이 끝나면(잠깐 멈추거나 Enter) 겹치는 칸 정리를 물음
            if getattr(self, "_ovl_id", None):
                self.root.after_cancel(self._ovl_id)
            self._ovl_row = r
            self._ovl_id = self.root.after(1500, self.ask_overlap)
        self.mark_edit(r)
        self.load_form_title()
        self.refresh_tree()
        self.fill_daylist()
        self.schedule_redraw()

    def load_form_title(self):
        if not self.valid_sel():
            return
        r = self.rows[self.sel]
        n = len(self.selected_rows())
        wb = self.whole_block()
        self.update_scope_bar(wb)
        if wb:
            ds = sorted({self.rows[k]["date"] for k in wb})
            self.form_title.config(text=f"묶인 칸 {len(wb)}일 ({day_label(ds[0])}~{day_label(ds[-1])})  "
                                        f"{fmt_t(r['s'])}~{fmt_t(r['e'])}  ({r['plan']})", fg="#1565c0")
            return
        if n == 1 and len((self.block_of(self.sel) or {}).get("_ris", [])) > 1:
            self.form_title.config(text=f"{day_label(r['date'])}만 (묶인 칸 중)  {fmt_t(r['s'])}~{fmt_t(r['e'])}"
                                        f"  ({r['plan']})", fg="#c62828")
            return
        self.form_title.config(text=(f"{n}칸 선택 · 기준: " if n > 1 else "") +
                               f"{day_label(r['date'])}  {fmt_t(r['s'])}~{fmt_t(r['e'])}  ({r['plan']})",
                               fg="#1565c0" if n > 1 else "black")

    def update_scope_bar(self, wb=None):
        """묶인 칸을 골랐을 때만 [묶음 전체]/[이 요일만] 을 보이고, 지금 범위 버튼을 눌린 모양으로"""
        bar = getattr(self, "scope_bar", None)
        if bar is None:
            return
        blk = self.block_of(self.sel) if self.valid_sel() else None
        if not blk or len(blk["_ris"]) < 2:
            bar.pack_forget()
            return
        if not bar.winfo_ismapped():
            bar.pack(fill="x", pady=(0, 4), after=self.form_title)
        on, off = dict(relief="sunken", bg="#bbdefb"), dict(relief="raised", bg=self._btn_bg)
        self.scope_all_btn.config(**(on if wb else off))
        self.scope_one_btn.config(**(off if wb else on))

    def form_pick_color(self, what="shade"):
        """what: shade(칸 음영) / hl(글자 바탕) / fg(글자색)"""
        from tkinter import colorchooser
        if not self.valid_sel():
            return
        attr, var, default = {"shade": ("f_shade", "shade_name", "#d9d9d9"), "hl": ("f_hl", "hl_name", "#a6d8f0"),
                              "fg": ("f_fg", "fg_name", "#000000")}[what]
        c = colorchooser.askcolor(color=getattr(self, attr) or default, parent=self.root)[1]
        if c:
            setattr(self, attr, c)
            self._loading = True
            self.fv[var].set("직접 고른 색")
            self._loading = False
            self.form_changed()

    def apply_same(self):
        """지금 칸의 음영·분홍·끝 시각·아이콘을 이번 주 같은 프로그램 모두에 똑같이"""
        from tkinter import messagebox
        if not self.valid_sel():
            return
        r = self.rows[self.sel]
        key = norm_key(split_suffix(r["name"])[0])
        same = [x for x in self.rows if x is not r and ((r.get("code") and x.get("code") == r.get("code"))
                                                         or norm_key(split_suffix(x["name"])[0]) == key)]
        if not same:
            messagebox.showinfo("알림", "이번 주에 같은 프로그램이 없습니다.")
            return
        if not messagebox.askyesno("같은 프로그램 모두", f"'{r['name']}'와 같은 프로그램 {len(same)}줄에\n"
                                   "음영·분홍·끝 시각 표시·아이콘(UHD·자막·수어·화면해설·연령·R/L)을 똑같이 맞출까요?"):
            return
        self.push_undo("same")
        for x in same:
            for k in ("shade", "force_pink", "show_end", "uhd", "cc", "sign", "ad", "age", "rl", "hl", "fg", "bold",
                      "top_dash", "bottom_dash"):
                x[k] = r.get(k)
            self.mark_edit(x)
        self.log(f"[같은 프로그램] '{r['name']}' 설정을 {len(same)}줄에 적용")
        self.refresh(keep=self.sel)

    def fill_daylist(self):
        self.daylist.delete(*self.daylist.get_children())
        if not self.valid_sel():
            return
        d = self.rows[self.sel]["date"]
        idx = sorted((i for i, r in enumerate(self.rows) if r["date"] == d),
                     key=lambda i: (self.rows[i]["plan"], self.rows[i]["s"]))
        for i in idx:
            r = self.rows[i]
            self.daylist.insert("", self.tk.END, iid=str(i), values=(fmt_t(r["s"]), fmt_t(r["e"]), show_prefix(r) + r["name"],
                                                                     r["plan"]),
                                tags=("edited",) if r.get("edited") else ())
        if self.daylist.exists(str(self.sel)):
            self._quiet = True
            self.daylist.selection_set(str(self.sel))
            self.daylist.see(str(self.sel))
            self._quiet = False

    # ---------- 즐겨찾기 (자주 쓰는 꾸밈을 버튼으로)
    FAV_KEYS = (("name", "프로그램명 통째로"), ("prefix", "이름 앞에 붙일 글자"), ("suffix", "이름 뒤에 붙일 글자"),
                ("shade", "칸 음영"), ("hl", "글자 바탕"), ("fg", "글자색"), ("bold", "굵게"),
                ("show_end", "끝 시각 표시"), ("force_pink", "분홍"))

    def build_fav_buttons(self):
        """즐겨찾기 목록: [☆ 프로그램명] [X] [배경색] - 오른쪽 탭과 즐겨찾기 창(열려 있으면) 둘 다"""
        targets = [(self.fav_frame, 9)]
        if getattr(self, "fav_win", None) is not None:
            targets = [(self.fav_win_frame, 11)]     # 창에서 보는 동안 탭 목록은 숨겨져 있음
        for frame, fsize in targets:
            try:
                self._fill_fav_list(frame, fsize)
            except Exception:
                if frame is not self.fav_frame:
                    self.fav_win = None

    def open_fav_window(self):
        """즐겨찾기를 따로 창으로: 크게 보고, 옆에 띄워 두고 편성표에서 칸을 고르며 ☆ 누르기"""
        tk = self.tk
        if getattr(self, "fav_win", None) is not None:
            try:
                self.fav_win.deiconify()
                self.fav_win.lift()
                return
            except Exception:
                self.fav_win = None
        w = tk.Toplevel(self.root)
        w.title("즐겨찾기")
        w.geometry("460x760")
        self.fav_win = w
        bar = tk.Frame(w)
        bar.pack(fill=tk.X, padx=6, pady=4)
        tk.Label(bar, text="편성표에서 칸을 고르고 ☆ 누르기", fg="#555").pack(side=tk.LEFT)
        tk.Button(bar, text="+ 추가", command=self.fav_add_dialog).pack(side=tk.RIGHT)
        tk.Button(bar, text="고른 칸 추가", command=self.fav_add_from_cell).pack(side=tk.RIGHT, padx=2)
        v_top = tk.BooleanVar(value=True)
        w.attributes("-topmost", True)
        tk.Checkbutton(bar, text="항상 위", variable=v_top,
                       command=lambda: w.attributes("-topmost", bool(v_top.get()))).pack(side=tk.RIGHT, padx=4)
        box = tk.Frame(w, relief="sunken", bd=1)
        box.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        cv = tk.Canvas(box, highlightthickness=0, bg="#f7f7f7")
        sb = tk.Scrollbar(box, orient=tk.VERTICAL, command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.fav_win_frame = tk.Frame(cv, bg="#f7f7f7")
        win = cv.create_window(0, 0, window=self.fav_win_frame, anchor="nw")
        self.fav_win_frame.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
        w.bind("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units"))

        w.protocol("WM_DELETE_WINDOW", self.close_fav_window)
        self.favf.pack_forget()                     # 오른쪽 탭 목록은 숨김
        self.fav_note.pack(fill=tk.BOTH, expand=True)
        self.build_fav_buttons()

    def close_fav_window(self):
        """즐겨찾기 창을 닫고 오른쪽 탭에 목록을 다시 보이기"""
        w, self.fav_win = getattr(self, "fav_win", None), None
        if w is not None:
            try:
                w.destroy()
            except Exception:
                pass
        self.fav_note.pack_forget()
        self.favf.pack(fill=self.tk.BOTH, expand=True)
        self.build_fav_buttons()

    def _fill_fav_list(self, frame, fsize=9):
        tk = self.tk
        for w_ in frame.winfo_children():
            w_.destroy()
        favs = load_config().get("favorites", [])
        if not favs:
            tk.Label(frame, text="(없음)  [+ 추가] 로 프로그램 이름을 미리 넣어 두세요.", fg="#999",
                     bg="#f7f7f7").pack(anchor="w", padx=4, pady=4)
        for n, fv in enumerate(favs):
            row = tk.Frame(frame, bg="#f7f7f7", bd=1, relief="groove")
            row.pack(fill=tk.X, padx=2, pady=1)
            text = fav_display(fv)
            summ = fav_opt_summary(fv)
            if summ:
                text += "\n    ⚙ " + summ
            bg = fv.get("shade") or fv.get("hl") or "#ffffff"
            b = tk.Button(row, text="☆ " + text, anchor="w", justify="left", relief="flat", bg=bg,
                          fg=fv.get("fg") or "black", font=("", fsize, "bold"), command=lambda fv=fv: self.fav_apply(fv))
            b.pack(side=tk.LEFT, fill=tk.X, expand=True)
            b.bind("<Button-3>", lambda e, n=n: self.fav_menu(e, n))
            tk.Button(row, text="배경색", font=("", 8), command=lambda n=n: self.fav_color(n)).pack(side=tk.RIGHT, padx=1)
            tk.Button(row, text="X", fg="#d32f2f", font=("", 8, "bold"), relief="flat",
                      command=lambda n=n: self.fav_delete(n)).pack(side=tk.RIGHT)

    def fav_color(self, n):
        from tkinter import colorchooser, messagebox
        favs = load_config().get("favorites", [])
        if not 0 <= n < len(favs):
            return
        c = colorchooser.askcolor(color=favs[n].get("shade") or "#d9d9d9", parent=self.root,
                                  title="이 즐겨찾기의 칸 배경색 (취소하면 그대로)")[1]
        if c:
            favs[n]["shade"] = c
        elif favs[n].get("shade") and messagebox.askyesno("배경색", "배경색을 없앨까요?"):
            favs[n]["shade"] = ""
        else:
            return
        save_config_keys({"favorites": favs})
        self.build_fav_buttons()

    def fav_add_from_cell(self):
        """지금 고른 칸의 프로그램명과 옵션(아이콘·분홍·끝 시각·굵게·음영·글자 바탕·글자색)을 즐겨찾기에 저장.
        같은 이름이 이미 있으면 그 즐겨찾기의 옵션을 지금 칸 것으로 바꿈"""
        if self.need_row() is None:
            return
        r = self.rows[self.sel]
        favs = load_config().get("favorites", [])
        fv = dict(name=r["name"], **self.cell_options(self.sel))
        same = next((n for n, f in enumerate(favs) if f.get("name") == r["name"]), None)
        if same is not None:        # 같은 이름이 이미 있으면 옵션을 지금 칸 것으로 새로 저장
            favs[same] = fv
            msg = "[즐겨찾기 옵션 새로 저장]"
        else:
            favs.append(fv)
            msg = "[즐겨찾기 추가]"
        save_config_keys({"favorites": favs})
        self.build_fav_buttons()
        self.log(f"{msg} {r['name']}  ({fav_opt_summary(fv) or '옵션 없음'})")

    def cell_options(self, i):
        """칸의 옵션 값 (즐겨찾기에 함께 저장하는 것)"""
        r = self.rows[i]
        d = {k: bool(r.get(k)) for k in ("uhd", "cc", "sign", "ad", "show_end", "bold", "top_dash", "bottom_dash")}
        d["force_pink"] = self.eff_pink(i)
        d["age"] = r.get("age", "") or ""
        for k in ("shade", "hl", "fg"):
            d[k] = r.get(k, "") or ""
        d["opts"] = True
        return d

    def fav_menu(self, e, n):
        m = self.tk.Menu(self.root, tearoff=0)
        m.add_command(label="고치기 ...", command=lambda: self.fav_add_dialog(edit_index=n))
        m.add_command(label="위로", command=lambda: self.fav_move(n, -1))
        m.add_command(label="아래로", command=lambda: self.fav_move(n, 1))
        m.add_command(label="지우기", command=lambda: self.fav_delete(n))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def fav_delete(self, n):
        from tkinter import messagebox
        favs = load_config().get("favorites", [])
        if 0 <= n < len(favs) and messagebox.askyesno("즐겨찾기",
                                                      f"'{fav_display(favs[n]).replace(chr(10), ' ')}' 을(를) 지울까요?"):
            favs.pop(n)
            save_config_keys({"favorites": favs})
            self.build_fav_buttons()

    def fav_move(self, n, d):
        favs = load_config().get("favorites", [])
        if 0 <= n + d < len(favs):
            favs[n], favs[n + d] = favs[n + d], favs[n]
            save_config_keys({"favorites": favs})
            self.build_fav_buttons()

    def fav_apply(self, fv):
        from tkinter import messagebox
        picked = self.selected_rows()
        if not picked:
            messagebox.showinfo("알림", "편성표에서 칸을 먼저 고르세요. (Ctrl/Shift+클릭으로 여러 칸)")
            return
        self.push_undo(("fav", fav_display(fv)))
        for k in picked:
            x = self.rows[k]
            if "name" in fv and fv["name"]:
                x["name"], x["name_edited"] = fv["name"], True
            pre, suf = fv.get("prefix", ""), fv.get("suffix", "")
            if pre and not x["name"].startswith(pre):
                x["name"], x["name_edited"] = pre + x["name"], True
            if suf and not x["name"].endswith(suf):
                x["name"], x["name_edited"] = x["name"] + suf, True
            for key in FAV_OPT_KEYS:     # 즐겨찾기에 저장된 옵션만 덮어씀
                if key in fv:
                    x[key] = fv[key]
            self.mark_edit(x)
        if fv.get("name"):          # 여러 칸에 같은 이름 -> 같은 코드로 한 칸에 묶이게
            code = share_code(self.rows, [self.rows[k] for k in picked])
            if code:
                self.log(f"[묶음 코드] {len(picked)}칸에 같은 코드 {code} 붙임")
        self.log(f"[즐겨찾기] '{fav_display(fv).replace(chr(10), ' ')}' -> {len(picked)}칸")
        self.refresh(keep=self.sel)

    def fav_add_dialog(self, edit_index=None):
        """즐겨찾기 추가/고치기: 프로그램 이름(여러 줄) + 옵션(아이콘·분홍·끝 시각·굵게·음영·글자 바탕·글자색)"""
        tk, ttk = self.tk, self.ttk
        from tkinter import messagebox, colorchooser
        favs = load_config().get("favorites", [])
        base = dict(favs[edit_index]) if edit_index is not None else {}
        top = tk.Toplevel(self.root)
        top.title("즐겨찾기 고치기" if edit_index is not None else "즐겨찾기 추가")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        tk.Label(f, text="프로그램 이름 (Enter 로 줄바꿈)", font=("", 10, "bold")).grid(row=0, column=0, columnspan=4,
                                                                                 sticky="w")
        txt = tk.Text(f, width=40, height=4, font=("", 11))
        txt.grid(row=1, column=0, columnspan=4, pady=4)
        txt.insert("1.0", (base.get("name") or "").replace("|", "\n"))
        txt.focus_set()

        has_opts = bool(base.get("opts")) or any(k in base for k in ("uhd", "cc", "sign", "ad"))
        v_opts = tk.BooleanVar(value=has_opts or edit_index is None and self.valid_sel())
        tk.Checkbutton(f, text="옵션도 함께 저장 (☆ 누르면 아래 옵션까지 그대로 적용)", variable=v_opts,
                       font=("", 9, "bold")).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 2))
        src = base if (has_opts or edit_index is not None) else (self.cell_options(self.sel) if self.valid_sel() else {})
        vb = {}
        of = tk.Frame(f, relief="groove", bd=1, padx=6, pady=4)
        of.grid(row=3, column=0, columnspan=4, sticky="we")
        for n, (k, t) in enumerate((("uhd", "UHD"), ("cc", "자막"), ("sign", "수어"), ("ad", "화면해설"))):
            vb[k] = tk.BooleanVar(value=bool(src.get(k)))
            tk.Checkbutton(of, text=t, variable=vb[k]).grid(row=0, column=n, sticky="w")
        for n, (k, t) in enumerate((("force_pink", "분홍"), ("show_end", "끝 시각 표시"), ("bold", "굵게"),
                                    ("top_dash", "위 선 점선"), ("bottom_dash", "아래 선 점선"))):
            vb[k] = tk.BooleanVar(value=bool(src.get(k)))
            tk.Checkbutton(of, text=t, variable=vb[k]).grid(row=1, column=n, sticky="w")
        vc = {}
        for row, (k, t, presets) in enumerate((("shade", "음영", SHADE_PRESETS), ("hl", "글자 바탕", HL_PRESETS),
                                                ("fg", "글자색", FG_PRESETS)), start=2):
            vc[k] = tk.StringVar(value=src.get(k, "") or "")
            tk.Label(of, text=t).grid(row=row, column=0, sticky="w")
            name0 = next((nm for nm, c in presets.items() if c == vc[k].get()), "직접 고른 색")
            vn = tk.StringVar(value=name0)
            cb = ttk.Combobox(of, textvariable=vn, values=list(presets), width=10, state="readonly")
            cb.grid(row=row, column=1, sticky="w", pady=1)
            cb.bind("<<ComboboxSelected>>", lambda e, k=k, vn=vn, pr=presets: vc[k].set(pr.get(vn.get(), "")))

            def pick(k=k, vn=vn):
                c = colorchooser.askcolor(color=vc[k].get() or "#d9d9d9", parent=top)[1]
                if c:
                    vc[k].set(c)
                    vn.set("직접 고른 색")
            tk.Button(of, text="색 고르기", command=pick).grid(row=row, column=2, sticky="w", padx=2)
        # 연령: 안 바꿈(☆ 눌러도 칸의 연령 그대로) / 없음 / 00·07·12·15·19
        tk.Label(of, text="연령").grid(row=5, column=0, sticky="w")
        v_age = tk.StringVar(value=(src.get("age") or FAV_AGE_NONE) if "age" in src else FAV_AGE_KEEP)
        ttk.Combobox(of, textvariable=v_age, values=(FAV_AGE_KEEP, FAV_AGE_NONE, "00", "07", "12", "15", "19"),
                     width=10, state="readonly").grid(row=5, column=1, sticky="w", pady=1)
        tk.Label(of, text="(안 바꿈: 칸의 연령 그대로)", fg="#777").grid(row=5, column=2, columnspan=3, sticky="w")

        def from_cell():
            if not self.valid_sel():
                messagebox.showinfo("알림", "편성표에서 칸을 먼저 고르세요.", parent=top)
                return
            o = self.cell_options(self.sel)
            for k in vb:
                vb[k].set(bool(o.get(k)))
            for k in vc:
                vc[k].set(o.get(k, ""))
            v_age.set(o.get("age") or FAV_AGE_NONE)
            v_opts.set(True)
            self.log("[즐겨찾기] 고른 칸의 옵션을 가져왔습니다. (색 이름 칸은 저장하면 맞춰집니다)")
        tk.Button(f, text="지금 고른 칸의 옵션 가져오기", command=from_cell).grid(row=4, column=0, columnspan=4,
                                                                         sticky="w", pady=4)
        tk.Label(f, text="☆ 누르면 고른 칸의 프로그램명이 이 이름으로 바뀌고, 옵션도 저장했으면 함께 바뀝니다.",
                 fg="#777").grid(row=5, column=0, columnspan=4, sticky="w")

        def ok():
            name = "|".join(x.strip() for x in txt.get("1.0", "end").strip().split("\n") if x.strip())
            if not name:
                messagebox.showwarning("입력", "프로그램 이름을 넣으세요.", parent=top)
                return
            fv = {"name": name}
            if v_opts.get():
                fv["opts"] = True
                for k in vb:
                    fv[k] = bool(vb[k].get())
                for k in vc:
                    fv[k] = vc[k].get()
                if v_age.get() != FAV_AGE_KEEP:
                    fv["age"] = "" if v_age.get() == FAV_AGE_NONE else v_age.get()
            favs2 = load_config().get("favorites", [])
            if edit_index is not None and edit_index < len(favs2):
                favs2[edit_index] = fv
            else:
                favs2.append(fv)
            save_config_keys({"favorites": favs2})
            self.build_fav_buttons()
            self.log(f"[즐겨찾기 저장] {name}  ({fav_opt_summary(fv) or '옵션 없음'})")
            top.destroy()
        bf = tk.Frame(f)
        bf.grid(row=6, column=0, columnspan=4, pady=8)
        tk.Button(bf, text="저장", command=ok, width=10, bg="#4CAF50").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="취소", command=top.destroy, width=10).pack(side=tk.LEFT, padx=6)
        top.after(150, lambda: top.grab_set())

    def on_daylist_select(self, e=None):
        if self._quiet:
            return
        s = self.daylist.selection()
        if not s or int(s[0]) == self.sel or int(s[0]) in self.selected_rows():
            return      # 프로그램이 다시 고른 것이거나 이미 고른 칸
        self.select(int(s[0]), source="daylist")

    def on_tree_select(self, e=None):
        if self._quiet:
            return
        s = [int(x) for x in self.tree.selection()]
        if not s:
            return
        # 프로그램이 표를 새로 고치며 다시 고른 것(늦게 오는 신호)이면 무시 -> 고치던 칸이 바뀌지 않게
        if set(s) == set(self.selected_rows()):
            return
        self.multi = s
        if self.sel in s:
            base = self.sel
        else:
            foc = self.tree.focus()
            base = int(foc) if foc and int(foc) in s else s[-1]
        self.select(base, source="tree", keep_multi=True)

    # ---- 편성표 그리기
    def schedule_redraw(self, delay=500):
        if self._redraw_id:
            try:
                self.root.after_cancel(self._redraw_id)
            except Exception:
                pass
        self._redraw_id = self.root.after(delay, self.redraw)

    def prev_rows_cached(self):
        p = self.v_prev.get().strip()
        if p != self._prev_path:
            self._prev_path = p
            self._prev_rows = self.prev_rows() if p else None
        return self._prev_rows

    def redraw(self):
        import base64
        self._redraw_id = None
        if not self.rows:
            return
        try:
            cfg = load_config()
            cfg.update(self.cfg_over())
            png, geom, dates = render_live(self.out_rows(), self.prev_rows_cached(), self.zoom, cfg)
        except Exception as e:
            self.draw_lbl.config(text=f"그리기 오류: {e}", fg="#c62828")
            self.log(f"[그리기 오류] {type(e).__name__}: {e}")
            return
        self.geom, self.dates = geom, dates
        self.pink_state = {ri: g["block"]["pink"] for g in geom for ri in g["block"]["_ris"]}
        self.img = self.tk.PhotoImage(data=base64.b64encode(png).decode("ascii"))
        self.cv.delete("all")
        self._tip_row = -1
        self.cv.create_image(0, 0, image=self.img, anchor="nw", tags=("page",))
        self.cv.configure(scrollregion=(0, 0, self.img.width(), self.img.height()))
        self.draw_lbl.config(text=("[DMB] " if self.dmb_on else "") +
                             f"{datetime.datetime.now():%H:%M:%S} 다시 그림 · 칸 {len(geom)}개",
                             fg="#8e24aa" if self.dmb_on else "#555")
        self.draw_marks()

    def day_x(self, date):
        try:
            return X_COLS["days"][self.dates.index(date)]
        except (ValueError, IndexError):
            return None

    def draw_marks(self):
        """선택한 칸(파란 점선)과 원본과 다른 칸(주황 점선) 테두리 - 화면에만 보이고 PDF에는 안 나옴"""
        self.cv.delete("mark")
        self.update_reason()
        if not self.show_marks.get():       # 출력 모양 그대로 보기
            return
        z = self.zoom
        for g in self.geom:
            b = g["block"]
            if (b["day0"], b["half"], b["s"]) in self.diff_keys:
                x0, y0, x1, y1 = g["rect"]
                self.cv.create_rectangle(x0 * z, y0 * z, x1 * z, y1 * z, outline="#ff9800", width=3, dash=(6, 3),
                                         tags=("mark",))
        ris = {}
        for g in self.geom:
            for ri in g["block"]["_ris"]:
                ris.setdefault(ri, g)
        for k, r in enumerate(self.rows):      # 새 판 반영: 바뀐·새 칸(주황), 확인할 칸(빨강)
            ovl = r.get("_ovl") or str(r.get("_note", "")).startswith("다른 칸과 시간이 겹침")
            if ovl and not find_overlaps(self.rows, k):     # 겹침이 풀렸으면 빨강도 내림
                for x in ("_check", "_note", "_ovl"):
                    r.pop(x, None)
            if not (r.get("_check") or r.get("_ver") in ("changed", "new")) or k not in ris:
                continue
            x0, y0, x1, y1 = ris[k]["rect"]
            dx = self.day_x(r["date"])
            if dx:
                x0, x1 = max(x0, dx[0]), min(x1, dx[1])
            self.cv.create_rectangle(x0 * z + 2, y0 * z + 2, x1 * z - 2, y1 * z - 2,
                                     outline="#e53935" if r.get("_check") else "#fb8c00", width=3, tags=("mark",))
        if self.empty_slot:         # 빈 시간: 초록 점선
            x0, y0, x1, y1 = self.empty_slot["rect"]
            self.cv.create_rectangle(x0 * z + 1, y0 * z + 1, x1 * z - 1, y1 * z - 1, outline="#2e7d32",
                                     width=3, dash=(4, 2), tags=("mark",))
        picked = set(self.selected_rows())
        drawn = set()
        for g in self.geom:        # 묶인 칸의 요일이 모두 골라졌으면 칸 전체에 테두리 하나
            ris = g["block"]["_ris"]
            if len(ris) > 1 and set(ris) <= picked:
                x0, y0, x1, y1 = g["rect"]
                main = self.sel in ris
                self.cv.create_rectangle(x0 * z + 1, y0 * z + 1, x1 * z - 1, y1 * z - 1, outline="#1e88e5",
                                         width=3 if main else 2, dash=(4, 2) if main else (2, 2), tags=("mark",))
                drawn.update(ris)
        for k in self.selected_rows():
            if k in drawn:
                continue
            r = self.rows[k]
            dx = self.day_x(r["date"])
            g = next((g for g in self.geom if k in g["block"]["_ris"]), None)
            if g is None:
                continue
            x0, y0, x1, y1 = g["rect"]
            if dx:      # 월~금으로 묶인 칸이면 그 요일 부분만
                x0, x1 = max(x0, dx[0]), min(x1, dx[1])
            main = k == self.sel
            self.cv.create_rectangle(x0 * z + 1, y0 * z + 1, x1 * z - 1, y1 * z - 1, outline="#1e88e5",
                                     width=3 if main else 2, dash=(4, 2) if main else (2, 2), tags=("mark",))

    def hit(self, e):
        z = self.zoom
        px, py = self.cv.canvasx(e.x) / z, self.cv.canvasy(e.y) / z
        area = lambda g: (g["rect"][2] - g["rect"][0]) * (g["rect"][3] - g["rect"][1])
        for g in sorted(self.geom, key=area):      # 칸이 겹쳐 있으면 작은 칸(위에 그려진 칸)부터
            x0, y0, x1, y1 = g["rect"]
            if x0 <= px <= x1 and y0 <= py <= y1:
                b = g["block"]
                day = next((i for i, (a, c) in enumerate(X_COLS["days"]) if a <= px <= c), b["day0"])
                date = self.dates[day] if day < len(self.dates) else None
                for ri in b["_ris"]:
                    if ri < len(self.rows) and self.rows[ri]["date"] == date:
                        return ri
                return b["_ris"][0]
        return None

    def on_canvas_click(self, e):
        i = self.hit(e)
        ctrl, shift = bool(e.state & 0x0004), bool(e.state & 0x0001)
        if i is None:
            if not ctrl and not shift:
                self.clear_selection()
                self.pick_empty_slot(e)
            return
        if ctrl:                    # Ctrl+클릭: 칸 하나씩 더하기 / 빼기
            if i in self.multi and len(self.multi) > 1:
                self.multi.remove(i)
                self.select(self.multi[-1], keep_multi=True)
            else:
                self.multi.append(i)
                self.select(i, keep_multi=True)
        elif shift:                 # Shift+클릭: 월~금으로 묶인 칸의 모든 요일
            blk = self.block_of(i)
            for k in (blk["_ris"] if blk else [i]):
                if k not in self.multi:
                    self.multi.append(k)
            self.select(i, keep_multi=True)
        elif e.state & (0x20000 if sys.platform == "win32" else 0x0008):    # Alt+클릭: 묶인 칸에서 그 요일만
            self.select(i)
        elif self.whole_block() and i in self.selected_rows():     # 묶음 전체를 고른 채 한 번 더: 그 요일만
            self.select(i)
        else:                       # 그냥 클릭: 묶인 칸이면 한 칸으로 통째로
            self.select_block(i)

    def empty_slot_at(self, e):
        """누른 곳이 편성 없는 빈 시간이면 {date, s, e, plan, rect}. 위·아래 칸 사이 시간, 옆이 2안(1안)이면 반쪽"""
        z = self.zoom
        px, py = self.cv.canvasx(e.x) / z, self.cv.canvasy(e.y) / z
        top, bottom = TABLE_TOP + HEAD_H, TABLE_BOTTOM
        if not (top <= py <= bottom):
            return None
        day = next((k for k, (a, c) in enumerate(X_COLS["days"]) if a <= px <= c), None)
        if day is None or day >= len(self.dates):
            return None
        cx0, cx1 = X_COLS["days"][day]
        mid = (cx0 + cx1) / 2
        col = [g for g in self.geom if g["rect"][0] <= px <= g["rect"][2]]
        up = max((g for g in col if g["rect"][3] <= py + 0.01), key=lambda g: g["rect"][3], default=None)
        lo = min((g for g in col if g["rect"][1] >= py - 0.01), key=lambda g: g["rect"][1], default=None)
        s_ = up["block"]["e"] if up else 0
        e_ = lo["block"]["s"] if lo else DAY_LEN_MIN
        if e_ <= s_:
            return None
        y0 = up["rect"][3] if up else top
        y1 = lo["rect"][1] if lo else bottom
        # 같은 날 같은 시간에 반쪽(1안/2안) 칸이 있으면 비어 있는 쪽 반만
        side = {g["block"]["half"] for g in self.geom if g["block"]["half"] and g["rect"][0] < cx1 and g["rect"][2] > cx0
                and g["block"]["s"] < e_ and g["block"]["e"] > s_}
        plan, x0, x1 = "1안", cx0, cx1
        if "R" in side and px < mid:
            x1 = mid
        elif "L" in side and px >= mid:
            plan, x0 = "2안", mid
        return {"date": self.dates[day], "s": s_, "e": e_, "plan": plan, "rect": (x0, y0, x1, y1)}

    def pick_empty_slot(self, e):
        """빈 시간을 누름: 초록 점선 + 오른쪽 수정란에 시각·편성안을 채워 둠 -> 프로그램명을 치면 새 줄"""
        slot = self.empty_slot_at(e)
        if slot is None:
            return
        self.empty_slot = slot
        self._loading = True
        try:
            self.fv["s"].set(fmt_t(slot["s"]))
            self.fv["e"].set(fmt_t(slot["e"]))
            self.fv["plan"].set(slot["plan"])
        finally:
            self._loading = False
        self.form_title.config(text=f"빈 시간 {day_label(slot['date'])} {fmt_t(slot['s'])}~{fmt_t(slot['e'])}"
                                    f" ({slot['plan']}) - 프로그램명을 넣으면 새 줄", fg="#2e7d32")
        self.draw_marks()

    def new_row_at_slot(self, name):
        """빈 시간(empty_slot)에 새 줄을 만들고 고름. 시각·편성안은 오른쪽 수정란 값(없으면 빈 시간 그대로)"""
        slot = self.empty_slot
        if not slot:
            return None
        try:
            s_ = hhmm_to_min(self.fv["s"].get()) - DAY_START_MIN
            e_ = hhmm_to_min(self.fv["e"].get()) - DAY_START_MIN
            if e_ <= s_ or not (0 <= s_ < DAY_LEN_MIN and 0 < e_ <= DAY_LEN_MIN):
                raise ValueError
        except (ValueError, TypeError):
            s_, e_ = slot["s"], slot["e"]
        self.push_undo("add")
        r = {"date": slot["date"], "code": "", "name": name, "s": s_, "e": e_,
             "uhd": False, "cc": True, "ad": False, "sign": False, "age": "", "rl": "", "rerun": False,
             "plan": self.fv["plan"].get() or slot["plan"], "edited": True}
        if self.dmb_on:         # DMB 로 보는 중에 넣는 줄은 DMB에서만
            r["show"] = "dmb"
        self.rows.append(r)
        self.saved = False
        self.log(f"[빈 시간에 추가] {day_label(r['date'])} {fmt_t(s_)}~{fmt_t(e_)} {name} ({r['plan']})")
        self.empty_slot = None
        self.refresh(keep=len(self.rows) - 1)
        return len(self.rows) - 1

    def add_at_slot(self):
        """빈 시간을 두 번 누름: 그 자리에 새 줄을 만들고 [자세히 고치기] 창"""
        i = self.new_row_at_slot("새 프로그램")
        if i is not None:
            self.edit_row(i, new=True)

    def block_of(self, i):
        return next((g["block"] for g in self.geom if i in g["block"]["_ris"]), None)

    def on_canvas_menu(self, e):
        i = self.hit(e)
        if i is None:
            return
        if i in self.selected_rows():
            self.select(i, keep_multi=True)
        else:
            self.select_block(i)
        r = self.rows[i]
        m = self.tk.Menu(self.root, tearoff=0)
        m.add_command(label=f"{day_label(r['date'])} {fmt_t(r['s'])} {r['name']}", state="disabled")
        m.add_separator()
        m.add_command(label="음영 넣기 / 빼기 ...", command=self.shade_dialog)
        m.add_command(label="꾸밈 복사 (음영·글자 바탕·글자색·굵게·점선)", command=lambda: self.copy_style(i))
        if self.style_clip:
            n = len(self.selected_rows())
            m.add_command(label=f"꾸밈 붙여넣기 ← '{self.style_clip['_from']}'" + (f" ({n}칸)" if n > 1 else ""),
                          command=self.paste_style)
        m.add_command(label=("분홍 끄기" if self.eff_pink(i) else "분홍 켜기"), command=self.toggle_pink)
        m.add_command(label=("끝 시각 표시 끄기" if r.get("show_end") else "끝 시각 표시 켜기"), command=self.toggle_end)
        m.add_separator()
        favs = load_config().get("favorites", [])
        if favs:
            fm = self.tk.Menu(m, tearoff=0)
            for fv in favs:
                fm.add_command(label=fav_display(fv).replace("\n", " "), command=lambda fv=fv: self.fav_apply(fv))
            m.add_cascade(label="즐겨찾기 적용", menu=fm)
        m.add_command(label="즐겨찾기에 추가 ...", command=self.fav_add_dialog)
        m.add_command(label="이름 앞뒤에 글자 붙이기 ...", command=self.affix_dialog)
        m.add_command(label=("DMB 칸 넣기 (TV 편성 위에 덮기) ..." if self.dmb_on else "로컬 끼워 넣기 ..."),
                      command=self.local_dialog)
        if r["plan"] == "2안":
            m.add_command(label="2안 지우기 (1안만 남기기)", command=lambda: self.del_plan2(i))
        else:
            m.add_command(label="2안 넣기 (이 시간을 1안/2안으로 나누기) ...", command=lambda: self.plan2_dialog(i))
        if self.plan_partners(i):     # 같은 날·겹치는 시간에 짝(1안/2안)이 있을 때만
            m.add_command(label="1안 ↔ 2안 서로 바꾸기", command=lambda: self.swap_plan(i))
        m.add_command(label="자세히 고치기 ...", command=lambda: self.edit_row())
        if self.whole_block():
            m.add_command(label="이 요일만 고르기 (Alt+클릭)", command=lambda: self.select(i))
        m.add_command(label="이 칸 뒤에 줄 추가 ...", command=self.add_row)
        if r.get("_check") or r.get("_ver") in ("changed", "new"):
            m.add_separator()
            if r.get("_note"):
                m.add_command(label="⚠ " + r["_note"][:60], state="disabled")
            m.add_command(label="이 칸 새 판 표시(주황·빨강) 지우기", command=lambda: self.clear_row_mark(i))
        if any(x.get("_check") or x.get("_ver") in ("changed", "new") for x in self.rows):
            m.add_command(label="새 판 표시 모두 지우기", command=self.clear_version_marks)
        m.add_command(label="줄 삭제", command=self.del_row)
        m.add_command(label=("TV 원래대로" if IS_DMB else "엑셀 원래대로"), command=self.revert_row)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def on_wheel(self, e):
        if e.state & 0x0004:        # Ctrl + 휠 = 확대/축소
            self.set_zoom(self.zoom * (1.1 if e.delta > 0 else 1 / 1.1))
        else:
            self.cv.yview_scroll(int(-e.delta / 120), "units")

    def set_zoom(self, z, pref=None):
        """보기 크기 바꾸기. pref: 기억할 값 (None 이면 이 배율, 'fit' 이면 한 장 맞춤)"""
        self.zoom = max(0.6, min(4.0, z))
        self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")
        self.schedule_redraw(100)
        self.remember_zoom(pref if pref in ("fit", "fitw") else round(self.zoom, 3))

    def zoom_fit(self):
        h = self.cv.winfo_height() or 700
        self.set_zoom((h - 4) / PAGE_H, pref="fit")

    def zoom_fit_width(self):
        """폭 맞춤: 편성표 가로가 창 너비에 꽉 차게 (세로는 스크롤)"""
        w = self.cv.winfo_width() or 900
        self.set_zoom((w - 24) / PAGE_W, pref="fitw")

    def remember_zoom(self, pref):
        """마지막 보기 크기를 설정 파일에 기억 (1초 동안 더 안 바뀌면 저장)"""
        self._zoom_pref = pref
        if getattr(self, "_zoom_save_id", None):
            try:
                self.root.after_cancel(self._zoom_save_id)
            except Exception:
                pass
        self._zoom_save_id = self.root.after(1000, self.save_zoom_pref)

    def save_zoom_pref(self):
        self._zoom_save_id = None
        pref = getattr(self, "_zoom_pref", None)
        if pref is None or load_config().get("view_zoom") == pref:
            return
        try:
            save_config_keys({"view_zoom": pref})
        except Exception as e:
            self.log(f"[주의] 보기 크기를 기억하지 못했습니다: {e}")

    def apply_saved_zoom(self):
        """파일을 열 때: 마지막 보기 크기 그대로 (처음이거나 '한 장 맞춤'이었으면 창에 맞춤)"""
        pref = load_config().get("view_zoom", "fit")
        if isinstance(pref, (int, float)) and not isinstance(pref, bool):
            self.set_zoom(float(pref))
        elif pref == "fitw":
            self.zoom_fit_width()
        else:
            self.zoom_fit()

    def icon_color_dialog(self):
        """아이콘(U·자·수·해·R·L·연령) 배경색과 글자색 바꾸기 - 바로 편성표에 반영, 설정에 저장"""
        tk = self.tk
        from tkinter import colorchooser
        cfg = load_config()
        colors = dict(cfg.get("icon_colors", {}) or {})
        dark = set(cfg.get("icon_dark_text", ["U"]))
        top = tk.Toplevel(self.root)
        top.title("아이콘 색")
        top.transient(self.root)
        f = tk.Frame(top, padx=14, pady=10)
        f.pack()
        to_hex = lambda rgb: "#%02x%02x%02x" % tuple(int(round(c * 255)) for c in rgb)
        sw = {}

        def cur_bg(k):
            return colors.get(k) or to_hex(tuple(c / 255 for c in ICON_COLORS[k]))

        def paint(k):
            sw[k].config(bg=cur_bg(k), fg="#1a1a1a" if k in dark else "#ffffff")

        def save():
            save_config_keys({"icon_colors": colors, "icon_dark_text": sorted(dark)})
            self.schedule_redraw(50)

        def pick(k):
            c = colorchooser.askcolor(color=cur_bg(k), parent=top)[1]
            if c:
                colors[k] = c
                paint(k)
                save()

        def flip(k):
            dark.symmetric_difference_update({k})
            paint(k)
            save()

        def reset(k, original=False):
            colors.pop(k, None)
            if original and ICON_ORIGINAL[k] != ICON_COLORS[k]:
                colors[k] = to_hex(tuple(c / 255 for c in ICON_ORIGINAL[k]))
            paint(k)
            save()

        for row, (k, name) in enumerate(ICON_NAMES):
            label = "12" if k == "AGE" else k
            sw[k] = tk.Label(f, text=label, width=3, font=("", 11, "bold"), relief="solid", bd=1)
            sw[k].grid(row=row, column=0, pady=2)
            paint(k)
            tk.Label(f, text=name, width=10, anchor="w").grid(row=row, column=1)
            tk.Button(f, text="배경색", command=lambda k=k: pick(k)).grid(row=row, column=2, padx=2)
            tk.Button(f, text="글자 흰/검정", command=lambda k=k: flip(k)).grid(row=row, column=3, padx=2)
            tk.Button(f, text="기본값", command=lambda k=k: reset(k)).grid(row=row, column=4, padx=2)
            if ICON_ORIGINAL[k] != ICON_COLORS[k]:
                tk.Button(f, text="원본 PDF 색", command=lambda k=k: reset(k, True)).grid(row=row, column=5, padx=2)
        tk.Label(f, text="바꾸면 바로 편성표·범례에 반영되고 다음에도 그대로 씁니다.", fg="#777").grid(
            row=10, column=0, columnspan=6, sticky="w", pady=(8, 0))
        tk.Button(f, text="닫기", command=top.destroy, width=10).grid(row=11, column=0, columnspan=6, pady=6)

    # ---------- 되돌리기 / 다시 하기
    # ---------- 새 판 반영 (같은 주 엑셀 (2)판 -> (3)판)
    def new_version(self):
        """같은 주의 새 판 엑셀을 열어, 지금 고친 것(로컬·음영·이름 등)을 그대로 옮겨 얹기"""
        from tkinter import filedialog, messagebox
        if not self.ready():
            return
        f = filedialog.askopenfilename(title="같은 주의 새 판 엑셀 (예: (3)판)", filetypes=[("엑셀", "*.xlsx")],
                                       initialdir=os.path.dirname(self.xlsx or "") or None)
        if not f:
            return
        try:
            new = read_schedule_xlsx(f)
        except Exception as e:
            self.error("엑셀 읽기 오류", e)
            return
        od, nd = sorted({r["date"] for r in self.orig}), sorted({r["date"] for r in new})
        if od != nd:
            span = lambda ds: f"{day_label(ds[0])} ~ {day_label(ds[-1])}" if ds else "-"
            if not messagebox.askyesno("새 판 반영", f"날짜가 다릅니다.\n지금: {span(od)}\n새 엑셀: {span(nd)}\n\n"
                                       "같은 주의 다른 판이 아닌 것 같습니다. 그래도 반영할까요?"):
                return
        cur = [{k: v for k, v in r.items() if k not in ("_ver", "_check", "_note", "_was", "_ovl")} for r in self.rows]
        # 기준 날짜: 새 엑셀 조회일 (그 앞 날은 이미 방송이 끝난 '지난 날')
        qd = read_query_date(f)
        base = qd if qd and nd and nd[0] < qd <= nd[-1] else None
        self.show_version_report(f, new, cur, base)

    def version_entries(self, rep):
        """보고 창 목록: (구분, 줄, 비고, 편성표에 있는 줄인지)"""
        ent = [("+ 새로 생김", r, "", True) for r in rep["new"]]
        ent += [("~ 바뀜", r, "예전: " + r.get("_was", ""), True) for r in rep["changed"]]
        ent += [("- 없어짐", r, "새 판에 없음", False) for r in rep["removed"]]
        ent += [("⚠ 확인", r, r.get("_note", ""), True) for r in rep["checks"]]
        ent += [("× 지운 줄", r, "예전 판에서 지운 줄이라 새 판에서도 뺌", False) for r in rep["dropped"]]
        ent += [("● 옮김", n, "고친 것 그대로: " + ", ".join(mv), True) for n, mv in rep["kept"]]
        ent += [("· 지난 날", r, f"반영 안 함 (새 판: {t})", False) for t, r in rep.get("past", [])]
        return sorted(ent, key=lambda e: (e[1]["date"], e[1]["s"], e[0]))

    def show_version_report(self, path, new, cur, base=None):
        tk, ttk = self.tk, self.ttk
        self.close_version_window()
        w = tk.Toplevel(self.root)
        self.ver_win = w
        w.title(f"새 판 반영 - {os.path.basename(path)}")
        w.geometry("940x600")
        self.ver_summ = tk.Label(w, anchor="w", font=("맑은 고딕", 10, "bold"))
        self.ver_summ.pack(fill=tk.X, padx=8, pady=(8, 2))
        # 지난 날은 그대로
        dates = sorted({r["date"] for r in new})
        opt = tk.Frame(w)
        opt.pack(fill=tk.X, padx=8, pady=2)
        self.v_keep_past = tk.BooleanVar(value=bool(base))
        labels = [day_label(d) for d in dates[1:]]
        self.v_keep_base = tk.StringVar(value=day_label(base) if base else (labels[0] if labels else ""))
        tk.Checkbutton(opt, text="지난 날은 예전 판 그대로 두기 -", variable=self.v_keep_past,
                       command=lambda: self.version_recalc()).pack(side=tk.LEFT)
        cb = ttk.Combobox(opt, textvariable=self.v_keep_base, values=labels, width=11, state="readonly")
        cb.pack(side=tk.LEFT, padx=2)
        cb.bind("<<ComboboxSelected>>", lambda e: (self.v_keep_past.set(True), self.version_recalc()))
        tk.Label(opt, text="부터 새 판 반영", fg="#333").pack(side=tk.LEFT)
        self.ver_keep_lbl = tk.Label(opt, fg="#777")
        self.ver_keep_lbl.pack(side=tk.LEFT, padx=8)
        self.ver_hint = tk.Label(w, anchor="w", fg="#555", justify="left",
                                 text="[반영]을 누르면 편성표가 새 판으로 바뀝니다. 바뀐 칸은 주황, 확인할 칸은 빨간 테두리"
                                      "(화면에만 보이고 PDF에는 안 나옴). 마음에 안 들면 [되돌리기].")
        self.ver_hint.pack(fill=tk.X, padx=8)
        box = tk.Frame(w)
        box.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        cols = (("kind", "구분", 90), ("day", "요일", 80), ("time", "시간", 95), ("name", "프로그램", 260),
                ("note", "비고", 380))
        tv = ttk.Treeview(box, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for c, t, wd in cols:
            tv.heading(c, text=t)
            tv.column(c, width=wd, anchor="w", stretch=c == "note")
        sb = ttk.Scrollbar(box, orient=tk.VERTICAL, command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for tag, color in (("new", "#ef6c00"), ("del", "#757575"), ("chk", "#c62828"), ("keep", "#2e7d32"),
                           ("past", "#9e9e9e")):
            tv.tag_configure(tag, foreground=color)
        self.ver_tree = tv
        tv.bind("<<TreeviewSelect>>", lambda e: self.on_version_pick())
        self.ver_bar = bar = tk.Frame(w)
        bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        tk.Button(bar, text="반영", width=12, bg="#ffb74d", command=self.apply_new_version).pack(side=tk.RIGHT, padx=2)
        tk.Button(bar, text="취소", width=10, command=self.close_version_window).pack(side=tk.RIGHT, padx=2)
        w.protocol("WM_DELETE_WINDOW", self.close_version_window)
        self.ver_applied = False
        self.ver_src = (path, new, cur, dates)
        self.version_recalc()

    def version_recalc(self):
        """지난 날 기준을 바꾸면 다시 계산해서 목록을 새로"""
        if self.ver_applied:        # 반영한 뒤에는 바꾸지 않음 (되돌리기 후 다시 [새 판 반영])
            return
        path, new, cur, dates = self.ver_src
        base = None
        if self.v_keep_past.get():
            base = next((d for d in dates if day_label(d) == self.v_keep_base.get()), None)
        try:
            merged, rep, orig = merge_keep_past(cur, self.orig, new, base)
        except Exception as e:
            self.error("새 판 반영 오류", e)
            return
        self.ver_result = (path, new, merged, rep, orig, base)
        past_days = [day_label(d) for d in dates if base and d < base]
        self.ver_keep_lbl.config(text=("그대로: " + ", ".join(past_days)) if past_days else "모든 날 새 판 반영")
        self.ver_summ.config(text=f"그대로 {len(rep['same'])}줄 · 바뀜 {len(rep['changed'])} · 새로 {len(rep['new'])} · "
                                  f"없어짐 {len(rep['removed'])} · 고친 것 옮김 {len(rep['kept'])}줄 · "
                                  f"로컬 {len(rep['locals'])}칸 · 직접 추가 {len(rep['added'])}줄 · "
                                  f"확인 필요 {len(rep['checks'])}" +
                                  (f" · 지난 날 반영 안 함 {len(rep['past'])}" if rep["past"] else ""))
        tv = self.ver_tree
        tv.delete(*tv.get_children())
        tagof = {"+": "new", "~": "new", "-": "del", "⚠": "chk", "×": "del", "●": "keep", "·": "past"}
        self.ver_entries = ent = self.version_entries(rep)
        for i, (kind, r, note, _live) in enumerate(ent):
            tv.insert("", "end", iid=str(i), tags=(tagof[kind[0]],),
                      values=(kind, day_label(r["date"]), f"{fmt_t(r['s'])}~{fmt_t(r['e'])}",
                              r["name"].replace("\n", " ") + (" (2안)" if r["plan"] == "2안" else ""), note))

    def apply_new_version(self):
        tk = self.tk
        path, new, merged, rep, orig, base = self.ver_result
        self.push_undo()
        old_name = os.path.basename(self.proj_path or self.xlsx or "")
        self.proj_dir = os.path.dirname(self.proj_path or "") or self.proj_dir or os.path.dirname(self.xlsx or "")
        self.rows = merged
        self.orig = orig
        self.xlsx = path
        self.v_x.set(path)
        self.query_date = read_query_date(path)
        self.proj_path = None               # 저장하면 새 판 이름으로 (예: ..(3)_편성표1710.tvproj)
        self.history = list(self.history) + [{
            "from": old_name, "to": os.path.basename(path), "at": datetime.datetime.now().isoformat(timespec="seconds"),
            "keep_before": base or "",
            "summary": {k: len(rep[k]) for k in ("changed", "new", "removed", "kept", "locals", "added", "checks",
                                                  "past")}}]
        self.sel, self.multi, self.diff_keys = None, [], set()
        self.pink_state = {}
        self.saved = False
        self.build_day_buttons()
        self.refresh()
        self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(path)} (새 판 반영, 저장 전)")
        self.log(f"[새 판 반영] {old_name} -> {os.path.basename(path)}  바뀜 {len(rep['changed'])} · 새로 "
                 f"{len(rep['new'])} · 없어짐 {len(rep['removed'])} · 고친 것 옮김 {len(rep['kept'])} · "
                 f"로컬 {len(rep['locals'])} · 확인 필요 {len(rep['checks'])}" +
                 (f" · {day_label(base)} 앞은 예전 판 그대로 (반영 안 한 변경 {len(rep['past'])})" if base else ""))
        for r in rep["checks"]:
            self.log(f"   ⚠ {day_label(r['date'])} {fmt_t(r['s'])}~{fmt_t(r['e'])} {r['name']} - {r.get('_note', '')}")
        self.ver_applied = True
        if self.ver_win is not None:
            for wdg in self.ver_bar.winfo_children():
                wdg.destroy()
            self.ver_hint.config(text="반영했습니다. 목록을 누르면 그 칸으로 갑니다. 확인이 끝나면 [표시 지우기]. "
                                      "저장하면 새 판 이름으로 저장됩니다.", fg="#2e7d32")
            tk.Button(self.ver_bar, text="닫기", width=10, command=self.close_version_window).pack(side=tk.RIGHT, padx=2)
            tk.Button(self.ver_bar, text="표시 지우기", width=12, command=self.clear_version_marks).pack(
                side=tk.RIGHT, padx=2)

    def find_row(self, row):
        """보고 창의 줄 -> 지금 편성표의 줄 번호 (되돌리기 뒤에도 날짜·시각·이름으로 찾음)"""
        for i, r in enumerate(self.rows):
            if r is row:
                return i
        key = lambda r: (r["date"], r["s"], r["e"], r["plan"], r["name"])
        return next((i for i, r in enumerate(self.rows) if key(r) == key(row)), None)

    def on_version_pick(self):
        sel = self.ver_tree.selection()
        if not sel or not self.ver_applied:
            return
        _kind, row, _note, live = self.ver_entries[int(sel[0])]
        i = self.find_row(row) if live else None
        if i is None:       # 없어진 줄: 그날 그 시각 근처 칸
            day = [k for k, r in enumerate(self.rows) if r["date"] == row["date"]]
            i = min(day, key=lambda k: abs(self.rows[k]["s"] - row["s"]), default=None)
        if i is not None:
            self.select(i)

    def clear_row_mark(self, i):
        if 0 <= i < len(self.rows):
            for k in ("_ver", "_check", "_note", "_was", "_ovl"):
                self.rows[i].pop(k, None)
            self.saved = False
            self.draw_marks()

    def clear_version_marks(self):
        for r in self.rows:
            for k in ("_ver", "_check", "_note", "_was", "_ovl"):
                r.pop(k, None)
        self.saved = False
        self.draw_marks()
        self.log("[새 판 반영] 주황·빨간 표시를 지웠습니다.")
        self.close_version_window()

    def close_version_window(self):
        if self.ver_win is not None:
            try:
                self.ver_win.destroy()
            except Exception:
                pass
        self.ver_win = None

    def snapshot(self):
        return (json.dumps(self.rows, ensure_ascii=False), self.deleted, self.sel,
                json.dumps({"orig": self.orig, "xlsx": self.xlsx, "query_date": getattr(self, "query_date", ""),
                            "proj_path": self.proj_path, "proj_dir": self.proj_dir, "history": self.history,
                            "tv_path": self.tv_path},
                           ensure_ascii=False))

    def push_undo(self, tag=None):
        """바꾸기 직전 상태를 기록. 같은 칸을 2초 안에 이어서 고치면(글자 입력 등) 한 번으로 묶음"""
        import time
        now = time.time()
        if tag is not None and tag == self._last_tag and now - self._last_t < 2.0 and self.undo_stack:
            self._last_t = now
            return
        self._last_tag, self._last_t = tag, now
        self.undo_stack.append(self.snapshot())
        del self.undo_stack[:-100]          # 최대 100단계
        self.redo_stack.clear()
        self.update_undo_buttons()

    def restore(self, snap):
        rows, deleted, sel = snap[:3]
        if len(snap) > 3:       # 새 판 반영 되돌리기: 엑셀 판도 함께
            st = json.loads(snap[3])
            if st["xlsx"] != self.xlsx:
                self.orig, self.xlsx, self.query_date = st["orig"], st["xlsx"], st["query_date"]
                self.proj_path, self.proj_dir, self.history = st["proj_path"], st["proj_dir"], st["history"]
                self.tv_path = st.get("tv_path", "")
                self.root.title(f"{APP_NAME} {VERSION} - {os.path.basename(self.proj_path or self.xlsx or '')}")
        self.rows = json.loads(rows)
        self.deleted = deleted
        self.saved = False
        self._last_tag = None
        self.pink_state = {}
        self.multi = []
        self.refresh(keep=sel if sel is not None else -1)

    def undo(self):
        if not self.undo_stack:
            return
        self.redo_stack.append(self.snapshot())
        self.restore(self.undo_stack.pop())
        self.log(f"[되돌리기] (남은 {len(self.undo_stack)}단계)")
        self.update_undo_buttons()

    def redo(self):
        if not self.redo_stack:
            return
        self.undo_stack.append(self.snapshot())
        self.restore(self.redo_stack.pop())
        self.log(f"[다시 하기] (남은 {len(self.redo_stack)}단계)")
        self.update_undo_buttons()

    def update_undo_buttons(self):
        n, m = len(self.undo_stack), len(self.redo_stack)
        self.btn_undo.config(text=f"↶ 되돌리기 ({n})" if n else "↶ 되돌리기",
                             state=self.tk.NORMAL if n else self.tk.DISABLED)
        self.btn_redo.config(text=f"↷ 다시 하기 ({m})" if m else "↷ 다시 하기",
                             state=self.tk.NORMAL if m else self.tk.DISABLED)

    def clear_all_pink(self):
        """지금 분홍인 칸을 모두 흰색으로 (2안·지난주 비교로 자동 칠해진 것 포함). 되돌리기 가능"""
        from tkinter import messagebox
        if not self.ready():
            return
        pinks = [i for i in range(len(self.rows)) if self.eff_pink(i)]
        if not pinks:
            messagebox.showinfo("알림", "분홍 칸이 없습니다.")
            return
        if not messagebox.askyesno("분홍 모두 지우기", f"분홍 칸 {len(pinks)}줄을 모두 흰색으로 바꿀까요?\n"
                                   "(되돌리기로 다시 살릴 수 있습니다)"):
            return
        self.push_undo("clear_pink")
        for i in pinks:
            self.rows[i]["force_pink"] = False
            self.mark_edit(self.rows[i])
        self.log(f"[분홍 모두 지우기] {len(pinks)}줄")
        self.refresh(keep=self.sel)

    # ---------- 첫 화면 / 작업 화면
    def show_start(self):
        """엑셀을 열기 전: 맨 위 줄과 안내 화면만"""
        for w_ in (self.act, self.head, self.nb, self.dmbbar):
            w_.pack_forget()
        self.placeholder.pack(fill=self.tk.BOTH, expand=True)
        self.main_shown = False

    def show_main(self):
        """엑셀을 연 뒤: 실행 버튼 줄, 편성표·수정란, 기록 창"""
        if self.main_shown:
            return
        tk = self.tk
        self.placeholder.pack_forget()
        self.act.pack(fill=tk.X, padx=8, pady=(0, 4), after=self.topbar)
        last = self.act
        if IS_DMB:
            self.dmbbar.pack(fill=tk.X, padx=8, pady=(0, 4), after=self.act)
            last = self.dmbbar
        if self.head_visible:
            self.head.pack(fill=tk.X, padx=8, pady=(0, 4), after=self.act)
            last = self.head
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(2, 6), after=last)
        self.main_shown = True

    def toggle_head(self):
        """⚙ 설정: 따로 창 (제목 · 오른쪽 위 날짜 · 아이콘 색 · 서버로 보내기). 이미 열려 있으면 앞으로"""
        tk, ttk = self.tk, self.ttk
        w = self.settings_win
        if w is not None and w.winfo_exists():
            w.deiconify()
            w.lift()
            w.focus_force()
            return
        w = tk.Toplevel(self.root)
        self.settings_win = w
        w.title(("엑셀 DMB" if IS_DMB else "엑셀 TV") + " 편성표 설정")
        w.transient(self.root)
        w.resizable(False, False)
        # 화면이 낮은 PC(노트북)에서도 [닫기]까지 보이도록: 내용이 화면보다 길면 세로 스크롤
        sc = tk.Canvas(w, highlightthickness=0)
        sb = tk.Scrollbar(w, orient="vertical", command=sc.yview)
        sc.configure(yscrollcommand=sb.set)
        sc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        f = tk.Frame(sc, padx=16, pady=12)
        sc.create_window(0, 0, window=f, anchor="nw")
        tk.Label(f, text="⚙ 편성표 설정", font=("", 12, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")
        tk.Label(f, text="바꾸면 편성표에 바로 반영되고, 다음 주에도 그대로 씁니다.", fg="#777").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(0, 8))

        # 제목
        g1 = tk.LabelFrame(f, text=" 제목 ", padx=10, pady=6)
        g1.grid(row=2, column=0, columnspan=4, sticky="we", pady=4)
        tk.Label(g1, text="제목").grid(row=0, column=0, sticky="w")
        tk.Entry(g1, textvariable=self.v_title, width=34).grid(row=0, column=1, columnspan=3, sticky="w", padx=4)
        tk.Label(g1, text="크기").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Combobox(g1, textvariable=self.v_tsize, values=("16", "18", "20", "22", "24", "26", "28", "32"),
                     width=5).grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        tk.Checkbutton(g1, text="굵게", variable=self.v_tbold).grid(row=1, column=2, sticky="w", pady=(4, 0))

        # 오른쪽 위 날짜
        g2 = tk.LabelFrame(f, text=" 오른쪽 위 날짜 ", padx=10, pady=6)
        g2.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)
        tk.Label(g2, text="표시").grid(row=0, column=0, sticky="w")
        ttk.Combobox(g2, textvariable=self.v_smode, width=12, state="readonly",
                     values=("저장 시각", "엑셀 조회일", "주간 범위", "직접 입력", "표시 안 함")).grid(
            row=0, column=1, sticky="w", padx=4)
        tk.Label(g2, text="직접 입력").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ent = tk.Entry(g2, textvariable=self.v_stext, width=34)
        ent.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))

        def stext_state(*a):
            ent.config(state="normal" if self.v_smode.get() == "직접 입력" else "disabled")
        stext_state()
        tid = self.v_smode.trace_add("write", stext_state)

        # 아이콘 색
        g3 = tk.LabelFrame(f, text=" 아이콘 ", padx=10, pady=6)
        g3.grid(row=4, column=0, columnspan=4, sticky="we", pady=4)
        tk.Button(g3, text="아이콘 색 ...", command=self.icon_color_dialog).grid(row=0, column=0, sticky="w")
        tk.Label(g3, text="UHD·자막·연령 등 아이콘 바탕색", fg="#777").grid(row=0, column=1, sticky="w", padx=6)

        # 선 (칸 테두리)
        cfg_l = load_config()
        g5 = tk.LabelFrame(f, text=" 선 (칸 테두리) ", padx=10, pady=6)
        g5.grid(row=5, column=0, columnspan=4, sticky="we", pady=4)
        v_cw = tk.StringVar(value=f"{line_width(cfg_l.get('cell_line_w'), CELL_LINE_W):g}")
        v_fw = tk.StringVar(value=f"{line_width(cfg_l.get('frame_line_w'), FRAME_LINE_W):g}")
        col = {"c": cfg_l.get("cell_line_color") or LINE_HEX}
        tk.Label(g5, text="칸 테두리 굵기").grid(row=0, column=0, sticky="w")
        ttk.Combobox(g5, textvariable=v_cw, values=CELL_LINE_WS, width=6).grid(row=0, column=1, sticky="w", padx=4)
        tk.Label(g5, text="pt   색").grid(row=0, column=2, sticky="w")
        sw = tk.Label(g5, text="      ", bg=col["c"], relief="solid", bd=1)
        sw.grid(row=0, column=3, sticky="w", padx=4)
        tk.Label(g5, text="바깥 테두리 굵기").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Combobox(g5, textvariable=v_fw, values=FRAME_LINE_WS, width=6).grid(row=1, column=1, sticky="w",
                                                                             padx=4, pady=(4, 0))
        tk.Label(g5, text="pt").grid(row=1, column=2, sticky="w", pady=(4, 0))
        pv = tk.Canvas(g5, width=200, height=56, bg="white", highlightthickness=0)
        pv.grid(row=2, column=0, columnspan=5, sticky="w", pady=(6, 2))
        tk.Label(g5, text="(4배 확대 미리보기)", fg="#777").grid(row=2, column=4, sticky="e")

        dir_ed = {}         # 고친 칸 미리보기 (아래에서 만듦)

        def draw_pv():
            pv.delete("all")
            cw = line_width(v_cw.get(), CELL_LINE_W)
            fw = line_width(v_fw.get(), FRAME_LINE_W)
            k = 4                   # 미리보기는 4배로
            pv.create_rectangle(8, 8, 192, 50, outline="#262626", width=max(1, fw * k))
            for x in (8 + 184 / 3, 8 + 184 * 2 / 3):
                pv.create_line(x, 8, x, 50, fill=col["c"], width=max(1, cw * k))
            pv.create_line(8, 29, 8 + 184 / 3, 29, fill=col["c"], width=max(1, cw * k))

        def save_lines(*a):
            cw = line_width(v_cw.get(), None)
            fw = line_width(v_fw.get(), None)
            if cw is None or fw is None:
                return
            save_config_keys({"cell_line_w": cw, "frame_line_w": fw, "cell_line_color": col["c"]})
            draw_pv()
            if "draw_epv" in dir_ed:
                dir_ed["draw_epv"]()
            self.schedule_redraw(50)

        def pick_color():
            from tkinter import colorchooser
            c = colorchooser.askcolor(color=col["c"], parent=w, title="칸 테두리 색")[1]
            if c:
                col["c"] = c.lower()
                sw.config(bg=col["c"])
                save_lines()

        def reset_lines():
            col["c"] = LINE_HEX
            sw.config(bg=LINE_HEX)
            v_cw.set(f"{CELL_LINE_W:g}")
            v_fw.set(f"{FRAME_LINE_W:g}")
            save_lines()
        tk.Button(g5, text="색 고르기", command=pick_color).grid(row=0, column=4, sticky="w", padx=4)
        tk.Button(g5, text="기본값으로", command=reset_lines).grid(row=3, column=0, sticky="w", pady=(2, 0))
        tk.Label(g5, text=f"(기본: 칸 {CELL_LINE_W:g}pt 회색 · 바깥 {FRAME_LINE_W:g}pt)", fg="#777").grid(
            row=3, column=1, columnspan=4, sticky="w", pady=(2, 0))
        line_tids = [(v_cw, v_cw.trace_add("write", save_lines)), (v_fw, v_fw.trace_add("write", save_lines))]
        draw_pv()

        # 고친 칸 (수정된 프로그램)
        g6 = tk.LabelFrame(f, text=" 고친 칸 (수정된 프로그램) ", padx=10, pady=6)
        g6.grid(row=6, column=0, columnspan=4, sticky="we", pady=4)
        e_on, e_w, e_t = edit_mark_cfg(cfg_l)
        v_eon = tk.BooleanVar(value=e_on)
        v_ew = tk.StringVar(value="끄기" if e_w is None else f"{e_w:g}")
        v_et = tk.StringVar(value=f"{e_t:.2f}")
        e_dx, e_dy = edit_time_gap(cfg_l)
        v_edx = tk.StringVar(value=f"{e_dx:.2f}")
        v_edy = tk.StringVar(value=f"{e_dy:.2f}")
        tk.Checkbutton(g6, text="고친 칸 따로 표시", variable=v_eon).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Label(g6, text="테두리 굵기").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Combobox(g6, textvariable=v_ew, values=EDIT_LINE_WS, width=6).grid(row=1, column=1, sticky="w",
                                                                            padx=4, pady=(4, 0))
        ecol = {"c": cfg_l.get("edited_line_color") or ""}      # "" = 칸 테두리 색과 같게
        tk.Label(g6, text="pt   색").grid(row=1, column=2, sticky="w", pady=(4, 0))
        esw = tk.Label(g6, text="      ", relief="solid", bd=1)
        esw.grid(row=1, column=3, sticky="w", padx=4, pady=(4, 0))
        ecf = tk.Frame(g6)
        ecf.grid(row=1, column=4, sticky="w", pady=(4, 0))

        def edit_color():
            return ecol["c"] or col["c"]

        def pick_ecolor():
            from tkinter import colorchooser
            c = colorchooser.askcolor(color=edit_color(), parent=w, title="고친 칸 테두리 색")[1]
            if c:
                ecol["c"] = c.lower()
                save_edit()

        def same_ecolor():
            ecol["c"] = ""
            save_edit()
        tk.Button(ecf, text="색 고르기", command=pick_ecolor).pack(side=tk.LEFT)
        tk.Button(ecf, text="칸과 같게", command=same_ecolor).pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(g6, text="시각 크기").grid(row=2, column=0, sticky="w", pady=(4, 0))
        def fine_spin(row, var, lo, hi, label, note):
            tk.Label(g6, text=label).grid(row=row, column=0, sticky="w", pady=(4, 0))
            tk.Spinbox(g6, textvariable=var, from_=lo, to=hi, increment=FINE_STEP, format="%.2f", width=6).grid(
                row=row, column=1, sticky="w", padx=4, pady=(4, 0))
            tk.Label(g6, text=note).grid(row=row, column=2, columnspan=3, sticky="w", pady=(4, 0))
        fine_spin(2, v_et, 2.0, 10.0, "시각 크기", "pt  (0.05pt씩)")
        fine_spin(3, v_edx, 0.0, 3.0, "왼쪽 간격", "pt  (선 ↔ 숫자)")
        fine_spin(4, v_edy, -1.0, 3.0, "위아래 옮김", "pt  (+아래로 · -위로)")
        epv = tk.Canvas(g6, width=220, height=84, bg="white", highlightthickness=0)
        epv.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 2))
        tk.Label(g6, text="(8배 확대 미리보기)", fg="#777").grid(row=5, column=4, sticky="e")

        def draw_epv():
            """칸 왼쪽 위 모서리를 8배로: 왼쪽은 보통 칸, 오른쪽은 고친 칸"""
            epv.delete("all")
            k = 8
            cw = line_width(v_cw.get(), CELL_LINE_W)
            ew = None if v_ew.get().strip() in ("끄기", "0", "") else line_width(v_ew.get(), None)
            ed = v_eon.get()
            et = time_size(v_et.get(), TIME_SZ) if ed else TIME_SZ
            edx = fine_value(v_edx.get(), 0.0, 3.0, TIME_DX) if ed else TIME_DX
            edy = fine_value(v_edy.get(), -1.0, 3.0, TIME_DY) if ed else TIME_DY
            c = col["c"]
            esw.config(bg=edit_color())
            top = 10
            for cx, sz, dx, dy, lc, lw_, txt in ((6, TIME_SZ, TIME_DX, TIME_DY, c, cw, "05"),
                                                 (116, et, edx, edy, edit_color() if (ed and ew) else c,
                                                  ew if (ed and ew) else cw, "50")):
                w_ = max(1, round(lw_ * k))
                epv.create_line(cx, top, cx + 100, top, fill=lc, width=w_)       # 위쪽 선
                epv.create_line(cx, top, cx, 82, fill=lc, width=w_)              # 왼쪽 선
                fpx = max(6, round(sz * k))
                base = top + (sz * 0.8 + dy) * k                                 # PDF 와 같은 글자 바닥선
                epv.create_text(cx + dx * k, base + fpx * 0.12, text=txt, anchor="sw", font=("Arial", -fpx))
            epv.create_text(56, 76, text="보통 칸", fill="#777")
            epv.create_text(166, 76, text="고친 칸", fill="#c62828")

        def save_edit(*a):
            ewv = v_ew.get().strip()
            vals = {"edited_mark": bool(v_eon.get()), "edited_line_color": ecol["c"],
                    "edited_mark_cut": bool(v_ecut.get()), "edited_mark_auto_local": bool(v_eauto.get()),
                    "edited_line_w": 0 if ewv in ("끄기", "0", "") else line_width(ewv, None),
                    "edited_time_sz": time_size(v_et.get(), None),
                    "edited_time_dx": fine_value(v_edx.get(), 0.0, 3.0, None),
                    "edited_time_dy": fine_value(v_edy.get(), -1.0, 3.0, None)}
            # 입력 중이거나 범위 밖인 칸은 건너뛰고, 올바른 칸만 저장 (한 칸 때문에 다른 칸이 막히지 않게)
            save_config_keys({k: v for k, v in vals.items() if v is not None})
            draw_epv()
            self.schedule_redraw(50)

        def reset_edit():
            v_eon.set(not IS_DMB)
            v_ew.set(f"{EDIT_LINE_W:g}")
            v_et.set(f"{TIME_SZ:.2f}")
            v_edx.set(f"{TIME_DX:.2f}")
            v_edy.set(f"{TIME_DY:.2f}")
            v_ecut.set(False)
            v_eauto.set(False)
            ecol["c"] = ""
            save_edit()
        tk.Button(g6, text="기본값으로", command=reset_edit).grid(row=6, column=0, sticky="w", pady=(2, 0))
        tk.Label(g6, text=f"(기본: {EDIT_LINE_W:g}pt · 칸과 같은 색 · 시각 {TIME_SZ:g} · 왼쪽 {TIME_DX:g} · 옮김 0)",
                 fg="#777").grid(row=6, column=1, columnspan=4, sticky="w", pady=(2, 0))
        tk.Label(g6, text="직접 고친 칸에만 · 화면과 PDF에 같이 · 직접 치면 0.01pt까지",
                 fg="#777").grid(row=7, column=0, columnspan=5, sticky="w")
        v_ecut = tk.BooleanVar(value=bool(cfg_l.get("edited_mark_cut")))
        v_eauto = tk.BooleanVar(value=bool(cfg_l.get("edited_mark_auto_local")))
        tk.Checkbutton(g6, text="로컬 때문에 잘린 칸도 표시", variable=v_ecut).grid(
            row=8, column=0, columnspan=3, sticky="w")
        tk.Checkbutton(g6, text="자동 로컬 칸도 표시 (매주 로컬 규칙)", variable=v_eauto).grid(
            row=9, column=0, columnspan=4, sticky="w")
        for v in (v_eon, v_ew, v_et, v_edx, v_edy, v_ecut, v_eauto):
            v.trace_add("write", save_edit)
        for v in (v_cw,):       # 칸 테두리 굵기를 바꾸면 고친 칸 미리보기도 다시
            v.trace_add("write", lambda *a: draw_epv())
        dir_ed["draw_epv"] = draw_epv
        draw_epv()

        # 서버로 보내기
        g4 = tk.LabelFrame(f, text=" 📺 서버로 보내기 (다른 PC에서 보기) ", padx=10, pady=6)
        g4.grid(row=7, column=0, columnspan=4, sticky="we", pady=4)
        if tv_preview_dir():
            def view_toggle():
                if not save_tv_view_config(self.v_view_on.get()):
                    self.log("[주의] 서버로 보내기 설정을 저장하지 못했습니다.")
                    return
                self.log("[서버로 보내기] " + ("켬 - PDF 저장할 때마다 녹음 PC 서버로 (TV 편성표 편집기와 같은 설정)"
                                          if self.v_view_on.get() else "끔 - 다른 PC·서버에서 볼 수 없습니다"))
            tk.Checkbutton(g4, text="PDF 저장할 때마다 녹음 PC 서버로 보내기", variable=self.v_view_on,
                           command=view_toggle).grid(row=0, column=0, sticky="w")
            tk.Label(g4, text="TV 편성표 편집기·런처와 같은 설정 · 프로젝트·즐겨찾기·설정 백업도 함께",
                     fg="#777").grid(row=1, column=0, sticky="w")
        else:
            tk.Label(g4, text="런처로 실행할 때만 쓸 수 있습니다.", fg="#999").grid(row=0, column=0, sticky="w")

        def close():
            try:
                self.v_smode.trace_remove("write", tid)
            except Exception:
                pass
            self.settings_win = None
            w.destroy()
        bf = tk.Frame(f)
        bf.grid(row=8, column=0, columnspan=4, sticky="we", pady=(10, 0))
        tk.Button(bf, text="설정 폴더 열기", command=self.open_config_folder).pack(side=tk.LEFT)
        tk.Button(bf, text="닫기", width=10, command=close).pack(side=tk.RIGHT)
        w.protocol("WM_DELETE_WINDOW", close)
        w.bind("<Escape>", lambda e: close())

        def fit(*a):
            f.update_idletasks()
            fw, fh = f.winfo_reqwidth(), f.winfo_reqheight()
            room = max(300, w.winfo_screenheight() - 120)      # 작업 표시줄·제목줄 자리 빼고
            sc.configure(width=fw, height=min(fh, room), scrollregion=(0, 0, fw, fh))
            if fh > room:
                if not sb.winfo_ismapped():
                    sb.pack(side=tk.RIGHT, fill=tk.Y)
                    w.bind("<MouseWheel>", lambda e: sc.yview_scroll(-1 if e.delta > 0 else 1, "units"))
                if w.winfo_y() + min(fh, room) + 60 > w.winfo_screenheight():
                    w.geometry(f"+{max(0, w.winfo_x())}+0")
            elif sb.winfo_ismapped():
                sb.pack_forget()
        fit()
        f.bind("<Configure>", lambda e: w.after_idle(fit))

    def cfg_over(self):
        o = {"_query_date": getattr(self, "query_date", "")}
        if self.dmb_on:
            o["_dmb"] = {"hide": [k for k, v in self.v_dmb_hide.items() if v.get()],
                         "keep_own": bool(self.v_dmb_keep.get()), "title": self.v_title.get().strip()}
        return o

    def out_rows(self):
        """그리기용 줄. DMB 편집기: 줄마다 DMB에서 직접 넣은 꾸밈(_own)을 표시 -> 'DMB에서 넣은 꾸밈은 그대로'"""
        if not IS_DMB:
            return self.rows
        tv = {r["_tvk"]: r for r in self.orig if r.get("_tvk")}
        out = []
        for r in self.rows:
            o = tv.get(r.get("_tvk")) if r.get("_tvk") else None
            own = [k for k in OWN_STYLE_KEYS if r.get(k) and (o is None or (r.get(k) or None) != (o.get(k) or None))]
            out.append(dict(r, _own=own) if own else r)
        return out

    def dmb_changed(self):
        if self._loading:
            return
        save_config_keys({"dmb_hide": [k for k, v in self.v_dmb_hide.items() if v.get()],
                          "dmb_keep_own": bool(self.v_dmb_keep.get())})
        if self.dmb_on:
            self.schedule_redraw(300)

    def head_changed(self):
        """제목·날짜 입력이 바뀌면 0.6초 뒤 설정에 저장하고 다시 그림"""
        if self._head_id:
            try:
                self.root.after_cancel(self._head_id)
            except Exception:
                pass

        def apply():
            self._head_id = None
            try:
                size = float(self.v_tsize.get())
            except ValueError:
                size = 22
            if self.v_smode.get() == "직접 입력" and not self.v_stext.get().strip() and self.rows:
                self._loading = True
                self.v_stext.set(make_stamp({"stamp_mode": "저장 시각"}, datetime.datetime.now(), self.dates))
                self._loading = False
            save_config_keys({"title_text": self.v_title.get(), "title_size": size,
                              "title_bold": self.v_tbold.get(), "stamp_mode": self.v_smode.get(),
                              "stamp_text": self.v_stext.get()})
            self.schedule_redraw(50)
        self._head_id = self.root.after(600, apply)

    def clear_diff(self):
        self.diff_keys = set()
        self.draw_marks()

    def run(self):
        self.root.mainloop()


def bind_linebreak(entry):
    """프로그램명 입력칸에서 Shift+Enter / Ctrl+Enter 를 누르면 줄바꿈 표시 '|' 를 넣음"""
    def ins(e):
        entry.insert("insert", "|")
        return "break"
    entry.bind("<Shift-Return>", ins)
    entry.bind("<Control-Return>", ins)


FAV_AGE_KEEP, FAV_AGE_NONE = "안 바꿈", "없음"     # 즐겨찾기 연령: 안 바꿈(저장 안 함) / 없음("")
FAV_OPT_KEYS = ("uhd", "cc", "sign", "ad", "age", "force_pink", "show_end", "bold", "top_dash", "bottom_dash",
                "shade", "hl", "fg")


def fav_opt_summary(fv):
    """즐겨찾기 옵션을 짧게: 'UHD·자막·수어 / 분홍·굵게 / 음영·바탕'"""
    if not any(k in fv for k in FAV_OPT_KEYS):
        return ""
    a = [t for k, t in (("uhd", "UHD"), ("cc", "자막"), ("sign", "수어"), ("ad", "화면해설")) if fv.get(k)]
    if "age" in fv:
        a.append(f"연령 {fv['age'] or '없음'}")
    b = [t for k, t in (("force_pink", "분홍"), ("show_end", "끝 시각"), ("bold", "굵게"),
                        ("top_dash", "위 점선"), ("bottom_dash", "아래 점선")) if fv.get(k)]
    c = [t for k, t in (("shade", "음영"), ("hl", "바탕"), ("fg", "글자색")) if fv.get(k)]
    return " / ".join("·".join(x) for x in (a, b, c) if x) or "옵션 모두 끔"


def fav_display(fv):
    """즐겨찾기 목록에 보일 글자: 프로그램명(줄바꿈 그대로), 없으면 예전 방식 이름표"""
    if fv.get("name"):
        return fv["name"].replace("|", "\n")
    if fv.get("prefix") or fv.get("suffix"):
        return f"{fv.get('prefix', '')}(이름){fv.get('suffix', '')}"
    return fv.get("label", "?")


class PreviewWindow:
    """미리보기 창.  확대/축소: ＋/－ 버튼, Ctrl+휠(마우스 위치 기준), + - 키, 0 = 한 장 맞춤
    이동: 마우스로 끌기, 휠(위아래), Shift+휠(좌우), 방향키, PageUp/PageDown, Home/End"""
    ZMIN, ZMAX = 0.5, 5.0

    def __init__(self, app, pdf_bytes):
        import tkinter as tk
        self.tk = tk
        self.doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        self.z = 1.0
        self._job = None
        self._anchor = None
        top = self.top = tk.Toplevel(app.root)
        top.title("미리보기 - 실제로 저장되는 PDF 그대로")
        top.geometry("1100x950")
        bar = tk.Frame(top)
        bar.pack(fill=tk.X, padx=6, pady=4)
        tk.Button(bar, text="－", width=3, command=lambda: self.zoom_by(1 / 1.25)).pack(side=tk.LEFT)
        self.lbl = tk.Label(bar, text="", width=6)
        self.lbl.pack(side=tk.LEFT)
        tk.Button(bar, text="＋", width=3, command=lambda: self.zoom_by(1.25)).pack(side=tk.LEFT)
        tk.Button(bar, text="한 장 맞춤", command=self.fit_page).pack(side=tk.LEFT, padx=(8, 2))
        tk.Button(bar, text="폭 맞춤", command=self.fit_width).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text="100%", command=lambda: self.set_zoom(1.0)).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text="200%", command=lambda: self.set_zoom(2.0)).pack(side=tk.LEFT, padx=2)
        tk.Label(bar, text="끌기: 이동 · 휠: 위아래 · Shift+휠: 좌우 · Ctrl+휠: 확대/축소 · 방향키 · 0: 한 장 맞춤",
                 fg="#666").pack(side=tk.LEFT, padx=10)
        fr = tk.Frame(top)
        fr.pack(fill=tk.BOTH, expand=True)
        cv = self.cv = tk.Canvas(fr, bg="#808080", highlightthickness=0, cursor="fleur")
        ys = tk.Scrollbar(fr, orient=tk.VERTICAL, command=cv.yview)
        xs = tk.Scrollbar(fr, orient=tk.HORIZONTAL, command=cv.xview)
        cv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        xs.pack(side=tk.BOTTOM, fill=tk.X)
        cv.pack(fill=tk.BOTH, expand=True)
        # 끌어서 이동
        for btn in ("<ButtonPress-1>", "<ButtonPress-2>"):
            cv.bind(btn, lambda e: cv.scan_mark(e.x, e.y))
        for btn in ("<B1-Motion>", "<B2-Motion>"):
            cv.bind(btn, lambda e: cv.scan_dragto(e.x, e.y, gain=1))
        cv.bind("<MouseWheel>", self.on_wheel)
        cv.bind("<Button-4>", lambda e: cv.yview_scroll(-3, "units"))
        cv.bind("<Button-5>", lambda e: cv.yview_scroll(3, "units"))
        keys = {"<Up>": (cv.yview_scroll, -3, "units"), "<Down>": (cv.yview_scroll, 3, "units"),
                "<Left>": (cv.xview_scroll, -3, "units"), "<Right>": (cv.xview_scroll, 3, "units"),
                "<Prior>": (cv.yview_scroll, -1, "pages"), "<Next>": (cv.yview_scroll, 1, "pages")}
        for k, (fn, n, what) in keys.items():
            top.bind(k, lambda e, fn=fn, n=n, what=what: fn(n, what))
        top.bind("<Home>", lambda e: cv.yview_moveto(0))
        top.bind("<End>", lambda e: cv.yview_moveto(1))
        for k in ("<plus>", "<equal>", "<KP_Add>"):
            top.bind(k, lambda e: self.zoom_by(1.25))
        for k in ("<minus>", "<KP_Subtract>"):
            top.bind(k, lambda e: self.zoom_by(1 / 1.25))
        top.bind("0", lambda e: self.fit_page())
        top.bind("<Escape>", lambda e: self.close())
        top.protocol("WM_DELETE_WINDOW", self.close)
        top.focus_set()
        top.after(120, self.fit_page)

    def close(self):
        try:
            self.doc.close()
        except Exception:
            pass
        self.top.destroy()

    def on_wheel(self, e):
        if e.state & 0x0004:            # Ctrl+휠: 마우스 있는 곳을 기준으로 확대/축소
            self.zoom_by(1.15 if e.delta > 0 else 1 / 1.15, e)
        elif e.state & 0x0001:          # Shift+휠: 좌우
            self.cv.xview_scroll(int(-e.delta / 120) * 3, "units")
        else:
            self.cv.yview_scroll(int(-e.delta / 120) * 3, "units")

    def fit_page(self):
        w, h = max(100, self.cv.winfo_width()), max(100, self.cv.winfo_height())
        self.set_zoom(min((w - 10) / PAGE_W, (h - 10) / PAGE_H))

    def fit_width(self):
        self.set_zoom((max(100, self.cv.winfo_width()) - 20) / PAGE_W)

    def zoom_by(self, f, e=None):
        self.set_zoom(self.z * f, e)

    def set_zoom(self, z, e=None):
        z = max(self.ZMIN, min(self.ZMAX, z))
        cv = self.cv
        if e is not None:       # 마우스 아래의 점이 그대로 있도록
            self._anchor = (cv.canvasx(e.x) / self.z, cv.canvasy(e.y) / self.z, e.x, e.y)
        else:                   # 화면 가운데가 그대로 있도록
            w, h = cv.winfo_width(), cv.winfo_height()
            self._anchor = (cv.canvasx(w / 2) / self.z, cv.canvasy(h / 2) / self.z, w / 2, h / 2)
        self.z = z
        self.lbl.config(text=f"{int(round(z * 100))}%")
        if self._job:
            self.top.after_cancel(self._job)
        self._job = self.top.after(60, self.render)

    def render(self):
        import base64
        self._job = None
        z = self.z
        pix = self.doc[0].get_pixmap(matrix=pymupdf.Matrix(z, z))
        self.img = self.tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"))
        cv = self.cv
        cv.delete("all")
        W, H = self.img.width(), self.img.height()
        cw, ch = cv.winfo_width(), cv.winfo_height()
        ox, oy = max(0, (cw - W) // 2), max(0, (ch - H) // 2)       # 화면보다 작으면 가운데
        cv.create_image(ox, oy, image=self.img, anchor="nw")
        cv.configure(scrollregion=(0, 0, max(W, cw), max(H, ch)))
        if self._anchor:
            px, py, sx, sy = self._anchor
            tw, th = max(W, cw), max(H, ch)
            if tw > cw:
                cv.xview_moveto(max(0.0, (px * z + ox - sx) / tw))
            if th > ch:
                cv.yview_moveto(max(0.0, (py * z + oy - sy) / th))


def open_file(path):
    try:
        os.startfile(path)      # 윈도우
    except AttributeError:
        import subprocess
        subprocess.Popen(["xdg-open", path])


_APP_MUTEX_HANDLE = None


def _already_running():
    """같은 편집기(엑셀 TV 또는 엑셀 DMB)가 이미 켜져 있으면 그 창을 앞으로 가져오고 True. (윈도우 전용, 실패하면 False)
    엑셀 TV 와 엑셀 DMB 는 서로 따로 셈한다 (둘은 함께 켤 수 있다)."""
    global _APP_MUTEX_HANDLE
    try:
        import ctypes
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateMutexW.restype = ctypes.c_void_p
        _APP_MUTEX_HANDLE = k32.CreateMutexW(None, False, "EditExcel_%s_Running" % ("DMB" if IS_DMB else "TV"))
        if ctypes.get_last_error() != 183:          # 183 = ERROR_ALREADY_EXISTS
            return False
    except Exception:
        return False
    try:                                            # 이미 켜진 창을 찾아 앞으로
        u32 = ctypes.WinDLL("user32")
        found = []
        proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def _cb(hwnd, _):
            n = u32.GetWindowTextLengthW(hwnd)
            if n and u32.IsWindowVisible(hwnd):
                buf = ctypes.create_unicode_buffer(n + 1)
                u32.GetWindowTextW(hwnd, buf, n + 1)
                if buf.value.startswith(APP_NAME + " "):
                    found.append(hwnd)
                    return False
            return True
        u32.EnumWindows(proto(_cb), 0)
        if found:
            if u32.IsIconic(found[0]):
                u32.ShowWindow(found[0], 9)         # SW_RESTORE: 최소화돼 있으면 되살리기
            u32.SetForegroundWindow(found[0])
    except Exception:
        pass
    return True


def run_gui():
    if _already_running():      # 두 번 눌러도 하나만: 이미 켜진 창을 앞으로 가져오고 끝낸다
        sys.exit(0)
    App().run()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--dmb"]
    if not args:
        run_gui()
    else:
        out = prev = None
        if "-o" in args:
            i = args.index("-o"); out = args[i + 1]; del args[i:i + 2]
        if "--prev" in args:
            i = args.index("--prev"); prev = args[i + 1]; del args[i:i + 2]
        if "--compare" in args:  # python tv_schedule_maker.py --compare 기준.pdf 비교.pdf
            i = args.index("--compare"); compare_pdfs(args[i + 1], args[i + 2]); sys.exit(0)
        if "--learn" in args:   # python tv_schedule_maker.py 주간.xlsx --learn 같은주_원본.pdf
            i = args.index("--learn"); learn_names(args[0], args[i + 1]); sys.exit(0)
        make_pdf(args[0], out, prev)
