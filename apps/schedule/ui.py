import os
import sys
import glob
import re
import time
import datetime
import threading
import json
import tempfile
import io
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
import base64
import copy
try:
    import winreg
except ImportError:
    winreg = None

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError:
    Image = ImageDraw = ImageFont = ImageTk = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

# 상단/하단 시간 글자 크기 초기값(pt)
DEFAULT_TIME_FONT_SIZE = 4

# PyMuPDF 는 여러 스레드에서 동시에 쓰면 안전하지 않다 -> PDF 작업(읽기·그리기·저장)은 이 잠금으로 한 번에 하나씩
PDF_LOCK = threading.RLock()


def schedule_badge_dir():
    """뱃지 PNG 폴더. 런처로 실행하면 데이터 폴더(C:\\EDIT\\TV\\데이터\\편성표\\badges) - 업데이트해도 남는다.
    따로 실행하면 예전처럼 실행 위치의 badges 폴더."""
    home = os.environ.get("RADIO_HOME")
    if home:
        return os.path.join(home, "데이터", "편성표", "badges")
    return os.path.abspath("badges")

APP_VERSION = "3.12.7"


def _detect_profile():
    """같은 프로그램을 두 가지로 쓴다: 편성표 편집기(schedule) / DMB 편성 편집기(dmb).
    런처에서 'dmb' 로 실행하거나, dmb.py 로 실행하거나, --dmb 를 붙이면 DMB."""
    if os.environ.get("RADIO_APP_ID") == "dmb" or os.environ.get("EDIT_PROFILE") == "dmb" or "--dmb" in sys.argv:
        return "dmb"
    return "schedule"


PROFILE = _detect_profile()
IS_DMB = PROFILE == "dmb"
APP_TITLE = (f"DMB 편성 편집기 v{APP_VERSION}" if IS_DMB
             else f"TV 편성표 편집기 v{APP_VERSION}")
# 편성표 편집기와 DMB 편집기가 같이 쓰는 것 (나머지 설정·자동 저장·머리글은 따로)
SHARED_KEYS = ("title_presets", "favorite_presets", "custom_badge_paths", "global_font_path")
DMB_REGIONS = ["광주", "전주", "여수", "목포"]


def schedule_settings_path():
    """설정(자동 저장) 파일 위치. 편집 프로그램 런처로 실행하면 C:\\EDIT\\TV\\데이터\\편성표\\settings.json,
    따로 실행하면 예전처럼 사용자 폴더. 새 위치가 비어 있으면 예전 설정을 한 번 복사해 이어받는다."""
    old = os.path.join(os.path.expanduser("~"), ".schedule_editor_settings.json")
    home = os.environ.get("RADIO_HOME")
    if IS_DMB:   # DMB 는 따로: C:\\EDIT\\TV\\데이터\\DMB\\settings.json
        if not home:
            return os.path.join(os.path.expanduser("~"), ".dmb_editor_settings.json")
        d = os.path.join(home, "데이터", "DMB")
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            return os.path.join(os.path.expanduser("~"), ".dmb_editor_settings.json")
        return os.path.join(d, "settings.json")
    if not home:
        return old
    new_dir = os.path.join(home, "데이터", "편성표")
    new = os.path.join(new_dir, "settings.json")
    try:
        os.makedirs(new_dir, exist_ok=True)
        if not os.path.exists(new):
            # 이어받을 예전 설정: RADIO 쪽에 잠깐 설치했던 것(C:\\EDIT\\radio\\...) -> 사용자 폴더의 것 순서
            cands = []
            if os.environ.get("EDIT_HOME"):
                cands.append(os.path.join(os.environ["EDIT_HOME"], "radio", "데이터", "편성표", "settings.json"))
            cands.append(old)
            src = next((c for c in cands if os.path.exists(c) and os.path.abspath(c) != os.path.abspath(new)), None)
            if src:
                import shutil
                shutil.copy2(src, new)
    except OSError:
        return old
    return new


def shared_settings_path():
    """두 편집기가 같이 쓰는 것(★ 칸 이름, 즐겨찾기 색, 뱃지, 글꼴) - C:\\EDIT\\TV\\데이터\\공용\\shared.json"""
    home = os.environ.get("RADIO_HOME")
    if not home:
        return os.path.join(os.path.expanduser("~"), ".edit_tv_shared.json")
    return os.path.join(home, "데이터", "공용", "shared.json")


def read_shared_settings():
    """같이 쓰는 설정. 아직 없으면 편성표 편집기 설정에서 이어받는다."""
    cands = [shared_settings_path()]
    home = os.environ.get("RADIO_HOME")
    cands.append(os.path.join(home, "데이터", "편성표", "settings.json") if home
                 else os.path.join(os.path.expanduser("~"), ".schedule_editor_settings.json"))
    for path in cands:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return {k: data[k] for k in SHARED_KEYS if k in data}
        except (OSError, ValueError):
            continue
    return {}


def write_shared_settings(values):
    path = shared_settings_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        text = json.dumps(values, ensure_ascii=False)
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(path + ".tmp", path)
    except OSError:
        pass


def _rgb_to_hex(c):
    """PyMuPDF 색(0~1, 회색 1값 / RGB 3값 / CMYK 4값) -> '#rrggbb'. 없거나 변환 실패 시 ''.
    (엑셀에서 인쇄한 PDF는 회색 칸을 회색 1값(예: 0.85)으로 쓴다)"""
    try:
        if c is None:
            return ""
        if isinstance(c, (int, float)):
            c = (c,)
        c = [float(v) for v in c]
        if len(c) == 1:
            c = [c[0]] * 3
        elif len(c) == 4:
            k = c[3]
            c = [(1 - c[0]) * (1 - k), (1 - c[1]) * (1 - k), (1 - c[2]) * (1 - k)]
        elif len(c) != 3:
            return ""
        return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(v * 255)))) for v in c)
    except Exception:
        return ""


# 원본 채우기 처리 범위: "pink" = 분홍 계열만 / "all" = 색이 있는 모든 칸 / "off" = 처리 안 함
CLEAR_FILL_MODE = "pink"

# 테두리 정보를 못 읽은 블록에 새로 그릴 때 쓰는 기본 테두리 색 (원본 편성표의 회색 #bfbfbf)
DEFAULT_BORDER_COLOR = (0.75, 0.75, 0.75)


def _is_pink_rgb(r, g, b):
    """0~1 RGB가 분홍~자홍 계열이면 True (연분홍 #ffe1e1, 핫핑크, 자홍 등)."""
    import colorsys
    hh, ss, vv = colorsys.rgb_to_hsv(r, g, b)
    hue = hh * 360
    if ss < 0.08 or vv < 0.6:      # 회색/흰색, 너무 어두운 색 제외
        return False
    if 295 <= hue < 360:           # 분홍 ~ 자홍
        return True
    if hue < 15 and ss < 0.65:     # 연한 빨강(분홍빛)
        return True
    return False


def _is_colored_rgb(r, g, b):
    """흰색과 어두운 색(글씨)을 뺀 모든 채우기 색."""
    return max(r, g, b) >= 0.6 and min(r, g, b) < 0.99


def _is_pink(hex_color):
    try:
        h = hex_color.lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:
        return False
    return _is_pink_rgb(r, g, b)


def _fill_targets(blocks, mode):
    out = []
    for b in blocks:
        f = (b.get("orig_fill") or "").lower()
        if not f or f == "#ffffff":
            continue
        if mode == "all" or (mode == "pink" and _is_pink(f)):
            out.append(b)
    return out


def _fill_report(blocks, mode):
    """인식된 채우기 색 목록(색값, 칸 수, 처리 여부)을 문자열로."""
    if mode == "off":
        return ""
    counts = {}
    for b in blocks:
        f = (b.get("orig_fill") or "").lower()
        if f:
            counts[f] = counts.get(f, 0) + 1
    if not counts:
        return "인식된 채우기 색: 없음"
    lines = ["인식된 채우기 색:"]
    for f, n in sorted(counts.items(), key=lambda kv: -kv[1])[:8]:
        done = f != "#ffffff" and (mode == "all" or _is_pink(f))
        lines.append(f"  {f}  {n}칸  {'→ 투명 처리' if done else '(유지)'}")
    return "\n".join(lines)


_NUM = rb'[-+]?(?:\d+\.?\d*|\.\d+)'
_FILL_OP_RE = re.compile(rb'(?<![\w.+\-])((?:' + _NUM + rb'\s+){1,4})(rg|scn|sc|k)(?![\w])')


def _recolor_stream(data, pred):
    """콘텐츠 스트림의 채우기 색 명령(rg/k/sc/scn) 중 pred에 걸리는 색을 흰색으로 바꾼다.
    반환: (새 바이트, 바꾼 개수)"""
    count = [0]

    def repl(m):
        try:
            vals = [float(x) for x in m.group(1).split()]
        except ValueError:
            return m.group(0)
        op = m.group(2)
        if op == b"rg" and len(vals) == 3:
            rgb, white = vals, b"1 1 1 rg"
        elif op == b"k" and len(vals) == 4:
            c, mm, y, k = vals
            rgb, white = ((1 - c) * (1 - k), (1 - mm) * (1 - k), (1 - y) * (1 - k)), b"0 0 0 0 k"
        elif op in (b"sc", b"scn") and len(vals) == 3:
            rgb, white = vals, b"1 1 1 " + op
        elif op in (b"sc", b"scn") and len(vals) == 4:
            c, mm, y, k = vals
            rgb, white = ((1 - c) * (1 - k), (1 - mm) * (1 - k), (1 - y) * (1 - k)), b"0 0 0 0 " + op
        else:
            return m.group(0)
        if pred(*rgb):
            count[0] += 1
            return white
        return m.group(0)

    return _FILL_OP_RE.sub(repl, data), count[0]


def _clear_orig_fill_pdf(path, page_idx, blocks):
    """원본 PDF의 해당 페이지에서 대상 채우기 색(기본: 분홍)을 흰색으로 바꾼 PDF 바이트를 만든다.

    도형을 찾아 지우는 대신 페이지 그리기 명령의 '채우기 색'만 흰색으로 바꾸므로
    글씨, 테두리, 이미지는 그대로 남는다.
    반환: (bytes 또는 None, 안내 문구 또는 None)
    """
    if CLEAR_FILL_MODE == "off":
        return None, None
    pred = _is_pink_rgb if CLEAR_FILL_MODE == "pink" else _is_colored_rgb

    doc = fitz.open(path)
    try:
        page = doc.load_page(page_idx)
        try:
            page.clean_contents()
        except Exception:
            pass

        total = 0
        contents = page.get_contents()
        if contents:
            new, n = _recolor_stream(page.read_contents(), pred)
            if n:
                doc.update_stream(contents[0], new)
                for x in contents[1:]:
                    doc.update_stream(x, b" ")
                total += n

        try:  # 폼 XObject 안에 그려진 도형
            for item in page.get_xobjects():
                xref = item[0]
                if doc.xref_get_key(xref, "Subtype")[1] != "/Form":
                    continue
                raw = doc.xref_stream(xref)
                if not raw:
                    continue
                new, n = _recolor_stream(raw, pred)
                if n:
                    doc.update_stream(xref, new)
                    total += n
        except Exception:
            pass

        if total == 0:
            if _fill_targets(blocks, CLEAR_FILL_MODE):
                return None, ("분홍색 칸은 인식됐지만 PDF 안에서 채우기 색 명령을 찾지 못해 변환하지 못했습니다.\n"
                              "(채우기가 이미지나 패턴으로 들어 있을 수 있습니다.)")
            return None, None
        return doc.tobytes(deflate=True), None
    finally:
        doc.close()


def _collect_line_segments(drawings):
    """원본 페이지의 가로/세로 선을 모은다. (칸 채우기가 선을 덮었을 때 복구하기 위한 데이터)
    항목: (종류, x0, y0, x1, y1, 색, 굵기)  종류 's'=선(stroke), 'f'=선처럼 얇은 채움 사각형"""
    segs = []
    for d in drawings:
        col = d.get("color")
        fcol = d.get("fill")
        w = d.get("width") or 0.5
        for it in d.get("items", []):
            op = it[0]
            if op == "l" and col is not None:
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.01 or abs(p1.x - p2.x) < 0.01:
                    segs.append(("s", min(p1.x, p2.x), min(p1.y, p2.y), max(p1.x, p2.x), max(p1.y, p2.y), tuple(col), w))
            elif op == "re":
                rr = it[1]
                if col is not None:
                    c = tuple(col)
                    segs += [("s", rr.x0, rr.y0, rr.x1, rr.y0, c, w), ("s", rr.x0, rr.y1, rr.x1, rr.y1, c, w),
                             ("s", rr.x0, rr.y0, rr.x0, rr.y1, c, w), ("s", rr.x1, rr.y0, rr.x1, rr.y1, c, w)]
                if fcol is not None and min(rr.width, rr.height) <= 2.0 and max(rr.width, rr.height) >= 3.0:
                    segs.append(("f", rr.x0, rr.y0, rr.x1, rr.y1, tuple(fcol), 0))
    return segs


def _restore_lines(page, segs, r, tol=2.5):
    """블록 r의 '테두리 근처'(가장자리 tol pt 이내)에 있던 원본 선만 다시 그린다.
    블록 안쪽의 선(병합된 칸의 내부 경계 등)은 복구하지 않는다."""
    for kind, x0, y0, x1, y1, col, w in segs:
        horizontal = abs(y0 - y1) < 0.01 if kind == "s" else (y1 - y0) <= (x1 - x0)
        if horizontal:
            pos = y0 if kind == "s" else (y0 + y1) / 2
            if abs(pos - r.y0) > tol and abs(pos - r.y1) > tol:
                continue
            a, b = max(x0, r.x0 - 1), min(x1, r.x1 + 1)
            if b <= a:
                continue
            if kind == "s":
                page.draw_line(fitz.Point(a, y0), fitz.Point(b, y0), color=col, width=w)
            else:
                page.draw_rect(fitz.Rect(a, y0, b, y1), color=None, fill=col)
        else:
            pos = x0 if kind == "s" else (x0 + x1) / 2
            if abs(pos - r.x0) > tol and abs(pos - r.x1) > tol:
                continue
            a, b = max(y0, r.y0 - 1), min(y1, r.y1 + 1)
            if b <= a:
                continue
            if kind == "s":
                page.draw_line(fitz.Point(x0, a), fitz.Point(x0, b), color=col, width=w)
            else:
                page.draw_rect(fitz.Rect(x0, a, x1, b), color=None, fill=col)


# ============================================================================
# 블록(칸) 찾기 — 테두리에서 칸을 만든다
#   · 크롬에서 만든 PDF: 테두리가 0.5pt 짜리 검은 가는 사각형, 칸마다 배경 사각형이 따로 있음
#   · 엑셀에서 인쇄한 PDF: 테두리는 칸마다 있지만 배경은 줄 전체를 가로 띠로 칠함
#   -> 배경이 아니라 '테두리'로 칸을 만들면 두 형식 모두 칸이 정확히 나온다.
# ============================================================================
DETECT_MIN_W = 20.0          # 블록으로 볼 최소 너비(pt)
DETECT_MIN_H = 3.0           # 블록으로 볼 최소 높이(pt) - 10~25분짜리 얇은 칸까지
BORDER_MAX_THICK = 2.2       # 이보다 얇은 사각형은 '테두리 선'으로 본다
_GRID_TOL = 0.7


def _lum(c):
    """색(0~1, 회색/RGB/CMYK) 밝기. 읽을 수 없으면 1(흰색)."""
    try:
        c = [float(v) for v in c]
    except Exception:
        return 1.0
    if len(c) == 1:
        return c[0]
    if len(c) == 3:
        return sum(c) / 3
    if len(c) == 4:
        return (1 - c[3]) * (1 - sum(c[:3]) / 3)
    return 1.0


def _is_header_color(c):
    """표 머리(시간·날짜 칸)의 연한 파랑 (#B8CCE4 등)."""
    try:
        r, g, b = [float(v) for v in c][:3]
    except Exception:
        return False
    return b > r + 0.1 and b > 0.75 and r > 0.55


def _drawing_parts(drawings):
    """get_drawings() 결과에서 (테두리 선분들, 채움 사각형들)을 뽑는다.
    선분: ('h', y, x0, x1, 색, 굵기) / ('v', x, y0, y1, 색, 굵기)
    채움: (Rect, 색)"""
    segs, fills = [], []
    for d in drawings:
        fill, stroke = d.get("fill"), d.get("color")
        width = d.get("width") or 0.5
        items = d.get("items", [])
        rects = []
        for it in items:
            if it[0] == "re":
                rects.append(fitz.Rect(it[1]))
            elif it[0] == "qu":
                rects.append(fitz.Rect(it[1].rect))
        if not rects and items and all(it[0] == "l" for it in items) and fill is not None:
            rects.append(fitz.Rect(d["rect"]))          # 선 4개로 그린 채움 사각형
        if fill is not None:
            for r in rects:
                if min(r.width, r.height) <= BORDER_MAX_THICK and _lum(fill) < 0.6:
                    if r.width >= r.height:
                        segs.append(("h", (r.y0 + r.y1) / 2, r.x0, r.x1, tuple(fill), max(r.height, 0.1)))
                    else:
                        segs.append(("v", (r.x0 + r.x1) / 2, r.y0, r.y1, tuple(fill), max(r.width, 0.1)))
                elif r.width > BORDER_MAX_THICK and r.height > BORDER_MAX_THICK:
                    fills.append((r, tuple(fill)))
        if stroke is not None and _lum(stroke) < 0.75:
            col = tuple(stroke)
            for it in items:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.3:
                        segs.append(("h", (p1.y + p2.y) / 2, min(p1.x, p2.x), max(p1.x, p2.x), col, width))
                    elif abs(p1.x - p2.x) < 0.3:
                        segs.append(("v", (p1.x + p2.x) / 2, min(p1.y, p2.y), max(p1.y, p2.y), col, width))
            for r in rects:
                segs += [("h", r.y0, r.x0, r.x1, col, width), ("h", r.y1, r.x0, r.x1, col, width),
                         ("v", r.x0, r.y0, r.y1, col, width), ("v", r.x1, r.y0, r.y1, col, width)]
    return segs, fills


def _cells_from_segments(segs, tol=_GRID_TOL, min_w=DETECT_MIN_W, min_h=DETECT_MIN_H):
    """테두리 선분들로 둘러싸인 직사각형 칸들. 칸 안의 뱃지 상자 같은 작은 구멍은 무시한다."""
    import bisect
    hs = [(s[1], s[2], s[3]) for s in segs if s[0] == "h"]
    vs = [(s[1], s[2], s[3]) for s in segs if s[0] == "v"]
    if not hs or not vs:
        return []

    def cluster(vals):
        out = []
        for v in sorted(vals):
            if out and v - out[-1][-1] <= tol:
                out[-1].append(v)
            else:
                out.append([v])
        return [sum(g) / len(g) for g in out]

    xs = cluster([v[0] for v in vs] + [h[1] for h in hs] + [h[2] for h in hs])
    ys = cluster([h[0] for h in hs] + [v[1] for v in vs] + [v[2] for v in vs])
    nx, ny = len(xs) - 1, len(ys) - 1
    if nx < 1 or ny < 1 or nx * ny > 400000:
        return []

    def nearest(arr, v):
        i = bisect.bisect_left(arr, v)
        return min((j for j in (i - 1, i) if 0 <= j < len(arr)), key=lambda j: abs(arr[j] - v))

    vblock, hblock = set(), set()
    ymids = [(ys[j] + ys[j + 1]) / 2 for j in range(ny)]
    xmids = [(xs[i] + xs[i + 1]) / 2 for i in range(nx)]
    for x, y0, y1 in vs:
        xi = nearest(xs, x)
        for j in range(bisect.bisect_left(ymids, y0 - tol), bisect.bisect_right(ymids, y1 + tol)):
            vblock.add((xi, j))
    for y, x0, x1 in hs:
        yj = nearest(ys, y)
        for i in range(bisect.bisect_left(xmids, x0 - tol), bisect.bisect_right(xmids, x1 + tol)):
            hblock.add((i, yj))
    parent = list(range(nx * ny))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    for j in range(ny):
        for i in range(nx):
            k = j * nx + i
            if i + 1 < nx and (i + 1, j) not in vblock:
                ra, rb = find(k), find(k + 1)
                if ra != rb:
                    parent[ra] = rb
            if j + 1 < ny and (i, j + 1) not in hblock:
                ra, rb = find(k), find(k + nx)
                if ra != rb:
                    parent[ra] = rb
    groups = {}
    for j in range(ny):
        for i in range(nx):
            groups.setdefault(find(j * nx + i), []).append((i, j))
    cells = []
    for g in groups.values():
        i0 = min(a for a, _ in g); i1 = max(a for a, _ in g)
        j0 = min(b for _, b in g); j1 = max(b for _, b in g)
        if i0 == 0 or j0 == 0 or i1 == nx - 1 or j1 == ny - 1:
            continue                                   # 표 바깥으로 이어진 영역
        x0, x1, y0, y1 = xs[i0], xs[i1 + 1], ys[j0], ys[j1 + 1]
        if x1 - x0 < min_w or y1 - y0 < min_h:
            continue
        area = sum((xs[a + 1] - xs[a]) * (ys[b + 1] - ys[b]) for a, b in g)
        if area < 0.8 * (x1 - x0) * (y1 - y0):
            continue                                   # ㄱ자 등 직사각형이 아닌 영역
        cells.append((x0, y0, x1, y1))
    return cells


def detect_schedule_blocks(drawings, page_rect=None, min_w=DETECT_MIN_W, min_h=DETECT_MIN_H):
    """편성표 페이지의 칸들. 반환: [{"rect", "border_color", "border_width", "orig_fill"}], 방식('테두리'/'배경')
    · 테두리로 칸을 만들고, 표 머리(파란 시간·날짜 칸)와 표 위쪽(제목·범례)은 뺀다.
    · 테두리로 칸이 거의 안 나오는 PDF 는 예전처럼 배경 사각형으로 찾는다."""
    segs, fills = _drawing_parts(drawings)
    pw = (page_rect[2] - page_rect[0]) if page_rect else 1e9
    ph = (page_rect[3] - page_rect[1]) if page_rect else 1e9

    def fill_at(x, y):
        best = None
        for r, c in fills:
            if r.x0 - 0.5 <= x <= r.x1 + 0.5 and r.y0 - 0.5 <= y <= r.y1 + 0.5:
                if best is None or r.width * r.height < best[0].width * best[0].height:
                    best = (r, c)
        return best[1] if best else None

    head_tops = [r.y0 for r, c in fills if _is_header_color(c)]
    table_top = min(head_tops) - 1 if head_tops else None
    out = []
    for (x0, y0, x1, y1) in _cells_from_segments(segs, min_w=min_w, min_h=min_h):
        if x1 - x0 > pw * 0.85 or y1 - y0 > ph * 0.85:
            continue
        if table_top is not None and y0 < table_top:
            continue
        fc = fill_at((x0 + x1) / 2, (y0 + y1) / 2)
        if fc is not None and _is_header_color(fc):
            continue
        # 테두리 색·굵기: 이 칸 윗변에 걸친 선
        bc, bw = None, None
        for s in segs:
            if s[0] == "h" and abs(s[1] - y0) <= 1.2 and s[2] - 1 <= (x0 + x1) / 2 <= s[3] + 1:
                bc, bw = s[4], s[5]
                break
        out.append({"rect": (x0, y0, x1, y1), "border_color": bc, "border_width": bw,
                    "orig_fill": _rgb_to_hex(fc) if fc else ""})
    if len(out) >= 10:
        return out, "테두리"
    # --- 테두리가 거의 없는 PDF: 배경 사각형 (작은 칸까지)
    out = []
    for r, c in fills:
        if r.width < max(min_w, 25) or r.height < min_h or r.width * r.height < 80:
            continue
        if r.width > pw * 0.85 or r.height > ph * 0.85 or _is_header_color(c) or _lum(c) < 0.3:
            continue
        box = (r.x0, r.y0, r.x1, r.y1)
        if any(all(abs(box[i] - o["rect"][i]) < 1 for i in range(4)) for o in out):
            continue
        out.append({"rect": box, "border_color": None, "border_width": None, "orig_fill": _rgb_to_hex(c)})
    return out, "배경"


# ============================================================================
# 시간 축(5분 단위로 나누기용) · 머리글(제목·날짜) 찾기
# ============================================================================
def _undouble(s):
    """굵은 글씨를 두 번 찍은 PDF('1100' = '10' 두 번)에서 원래 글자로."""
    if len(s) >= 2 and len(s) % 2 == 0 and all(s[i] == s[i + 1] for i in range(0, len(s), 2)):
        return s[::2]
    return s


def detect_time_rows(words, segs):
    """편성표 왼쪽의 시간 숫자 열(5,6,7…12,1,2…)과 가로 테두리로 1시간 줄들을 찾는다.
    · 시간 숫자가 가장 많이 세로로 늘어선 열을 시간 열로 본다.
    · 그 열을 가로지르는 테두리 사이가 1시간. 숫자가 겹쳐 읽히지 않는 줄도 순서대로 시각을 매긴다.
    words: [(x0,y0,x1,y1,text,...)], segs: _drawing_parts 의 선분
    반환: [(y0, y1, hour24)] 위에서부터. 못 찾으면 []"""
    cand = []
    for w in words:
        t = _undouble(str(w[4]).strip())
        if t.isdigit() and len(t) <= 2 and 1 <= int(t) <= 24:
            cand.append(((w[0] + w[2]) / 2, (w[1] + w[3]) / 2, int(t)))
    cols = []
    for cx, cy, h in cand:
        col = next((c for c in cols if abs(c["x"] - cx) <= 4), None)
        if col is None:
            col = {"x": cx, "items": []}
            cols.append(col)
        col["items"].append((cy, h))
    cols = [c for c in cols if len(c["items"]) >= 6]
    if not cols:
        return []
    # 같은 시간 열이 여러 개(표 양쪽)면 가장 왼쪽. 숫자가 가장 많은 열들 중에서
    most = max(len(c["items"]) for c in cols)
    col = min((c for c in cols if len(c["items"]) >= most * 0.7), key=lambda c: c["x"])
    cx = col["x"]
    ys = sorted({round(sg[1], 1) for sg in segs if sg[0] == "h" and sg[2] - 1 <= cx <= sg[3] + 1})
    bounds = []
    for y in ys:
        if bounds and y - bounds[-1] <= 1.5:
            continue
        bounds.append(y)
    labels = sorted(col["items"])
    lo, hi = labels[0][0], labels[-1][0]
    rows = []
    for y0, y1 in zip(bounds, bounds[1:]):
        if y1 - y0 < 8 or y1 < lo - 1 or y0 > hi + 1:
            continue
        lab = [h for cy, h in labels if y0 <= cy <= y1]
        rows.append([y0, y1, lab[0] if len(lab) == 1 else None])
    if not rows:
        return []
    # 시각 매기기: 읽힌 숫자 중 첫 번째를 기준으로 한 줄에 1시간씩
    first = next((i for i, r in enumerate(rows) if r[2] is not None), None)
    if first is None:
        return []
    base = rows[first][2] - first
    out = []
    for i, (y0, y1, _) in enumerate(rows):
        out.append((y0, y1, (base + i) % 24))
    return out


def time_ticks(rows, step=5):
    """1시간 줄마다 step 분 간격의 (y, '시:분')."""
    ticks = []
    for y0, y1, h in rows:
        for m in range(0, 60, step):
            ticks.append((y0 + (y1 - y0) * m / 60.0, f"{h:02d}:{m:02d}"))
    if rows:
        y0, y1, h = rows[-1]
        ticks.append((y1, f"{(h + 1) % 24:02d}:00"))
    return ticks


_DATE_RE = re.compile(r"\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일")


def detect_header(text_dict, images, table_top, page_rect):
    """표 위쪽의 제목(로고 그림 또는 가장 큰 글자 줄)과 날짜 줄을 찾는다.
    반환: {"title": {"rect", "text"} 또는 None, "date": {"rect", "text", "size"} 또는 None}"""
    top_limit = table_top if table_top else (page_rect[1] + (page_rect[3] - page_rect[1]) * 0.15)
    lines = []
    for b in (text_dict or {}).get("blocks", []):
        for ln in b.get("lines", []):
            spans = [s for s in ln.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            x0 = min(s["bbox"][0] for s in spans); y0 = min(s["bbox"][1] for s in spans)
            x1 = max(s["bbox"][2] for s in spans); y1 = max(s["bbox"][3] for s in spans)
            if y1 > top_limit:
                continue
            lines.append({"rect": (x0, y0, x1, y1), "text": "".join(s["text"] for s in spans).strip(),
                          "size": max(s.get("size", 0) for s in spans)})
    date = next((ln for ln in lines if _DATE_RE.search(ln["text"])), None)
    title = None
    imgs = [im for im in (images or []) if im["bbox"][3] <= top_limit and (im["bbox"][2] - im["bbox"][0]) > 60]
    if imgs:
        bb = max(imgs, key=lambda im: (im["bbox"][2] - im["bbox"][0]) * (im["bbox"][3] - im["bbox"][1]))["bbox"]
        title = {"rect": tuple(bb), "text": ""}
    else:
        big = [ln for ln in lines if ln is not date and ln["size"] >= 12]
        if big:
            ln = max(big, key=lambda l: l["size"])
            title = {"rect": ln["rect"], "text": ln["text"]}
    return {"title": title, "date": date}


_DAY_RE = re.compile(r"^(\d{1,2})/(\d{1,2})\s*\(([월화수목금토일])\)")


def detect_day_headers(text_dict, fills):
    """표 머리의 요일 칸 글자('11/2(월)')를 찾는다 (DMB 에서 '11/2(월)전주' 로 바꾸는 자리).
    반환: 왼쪽부터 [{"rect", "text", "size", "cx", "cy", "cell_w", "fill"}] (3개 미만이면 [])"""
    found = []
    for b in (text_dict or {}).get("blocks", []):
        for ln in b.get("lines", []):
            spans = [sp for sp in ln.get("spans", []) if sp.get("text", "").strip()]
            if not spans:
                continue
            text = "".join(sp["text"] for sp in spans).strip()
            if not _DAY_RE.match(text) or len(text) > 14:
                continue
            x0 = min(sp["bbox"][0] for sp in spans); y0 = min(sp["bbox"][1] for sp in spans)
            x1 = max(sp["bbox"][2] for sp in spans); y1 = max(sp["bbox"][3] for sp in spans)
            found.append({"rect": (x0, y0, x1, y1), "text": text, "size": max(sp.get("size", 0) for sp in spans)})
    if len(found) < 3:
        return []
    # 같은 높이에 가장 많이 모인 줄만 (표 머리)
    rows = {}
    for d in found:
        rows.setdefault(round((d["rect"][1] + d["rect"][3]) / 4), []).append(d)
    row = sorted(max(rows.values(), key=len), key=lambda d: d["rect"][0])
    if len(row) < 3:
        return []
    cxs = [(d["rect"][0] + d["rect"][2]) / 2 for d in row]
    for i, d in enumerate(row):
        gaps = [abs(cxs[i] - cxs[j]) for j in (i - 1, i + 1) if 0 <= j < len(row)]
        d["cx"] = cxs[i]
        d["cy"] = (d["rect"][1] + d["rect"][3]) / 2
        d["cell_w"] = max(d["rect"][2] - d["rect"][0] + 4, min(gaps))
        fill = None
        best = None
        for r, c in fills or []:
            hx = _rgb_to_hex(c)
            if hx and r.x0 - 0.5 <= d["cx"] <= r.x1 + 0.5 and r.y0 - 0.5 <= d["cy"] <= r.y1 + 0.5:
                area = r.width * r.height
                if best is None or area < best:
                    best, fill = area, hx
        d["fill"] = fill or "#B8CCE4"
    return row



# ---------------------------------------------------------------- 새 판 비교 · 고친 내용 옮기기
def block_source_texts(words, rects):
    """원본 PDF 글자 중 각 칸 안에 들어 있는 것을 이어 붙인다 (판끼리 같은 칸인지 비교할 때 씀)."""
    out = []
    for r in rects:
        x0, y0, x1, y1 = r
        ws = [w for w in words if x0 - 0.5 <= (w[0] + w[2]) / 2 <= x1 + 0.5 and y0 - 0.5 <= (w[1] + w[3]) / 2 <= y1 + 0.5]
        ws.sort(key=lambda w: (round(w[1]), w[0]))
        out.append(" ".join(w[4] for w in ws).strip())
    return out


def _norm_text(t):
    return re.sub(r"\s+", "", t or "")


def match_original_blocks(old_initial, new_initial, tol=2.5):
    """이전 판과 새 판의 '원본 칸'을 짝짓는다 (위치가 tol pt 안이고 원본 글자가 같으면 같은 칸).
    반환: {이전 칸 id: 새 칸}, 새 판에서 바뀐(짝 없는) 칸 id 집합"""
    pairs, used = {}, set()
    for ob in old_initial:
        ot = _norm_text(ob.get("src_text"))
        best = None
        for nb in new_initial:
            if nb["id"] in used:
                continue
            if all(abs(a - c) <= tol for a, c in zip(ob["rect"], nb["rect"])) and _norm_text(nb.get("src_text")) == ot:
                d = sum(abs(a - c) for a, c in zip(ob["rect"], nb["rect"]))
                if best is None or d < best[0]:
                    best = (d, nb)
        if best:
            pairs[ob["id"]] = best[1]
            used.add(best[1]["id"])
    changed = {nb["id"] for nb in new_initial if nb["id"] not in used}
    return pairs, changed


def _overlap(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def _area(r):
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


_EDIT_KEYS = (("title", ""), ("color", ""), ("text_color", "#000000"), ("font_size_offset", 0), ("text_dx", 0), ("text_dy", 0))


def _is_edited(b, initial):
    if b.get("shade"):
        return True
    for ob in initial:
        if all(abs(a - c) <= 0.5 for a, c in zip(b["rect"], ob["rect"])):
            return any(b.get(k, d) != ob.get(k, d) for k, d in _EDIT_KEYS) or b.get("options", {}) != ob.get("options", {})
    return True


def transplant_work(old_initial, old_blocks, new_initial, next_id):
    """이전 판에서 고친 칸을 새 판으로 옮긴다.
    고친 칸마다 그 자리의 '이전 판 원본 칸'을 찾고, 그 원본 칸이 새 판에도 그대로(위치·글자) 있을 때만 옮긴다.
    반환: (새 판 칸 목록, 옮긴 칸 수, 못 옮긴 칸 이름들, 새 판에서 바뀐 칸 id 집합, 다음 id)"""
    pairs, changed = match_original_blocks(old_initial, new_initial)
    edited = [b for b in old_blocks if _is_edited(b, old_initial)]
    # 고친 칸 -> 그 자리 원본 칸들
    src_of = {}
    for e in edited:
        srcs = [ob for ob in old_initial
                if _overlap(e["rect"], ob["rect"]) >= 0.25 * max(1e-6, min(_area(e["rect"]), _area(ob["rect"])))]
        src_of[e["id"]] = srcs
    # 원본 칸을 같이 쓰는 고친 칸끼리 한 묶음 (나눈 칸 둘 = 같은 원본)
    comps, seen = [], set()
    for e in edited:
        if e["id"] in seen:
            continue
        comp_e, comp_s, stack = [], {}, [e]
        seen.add(e["id"])
        while stack:
            cur = stack.pop()
            comp_e.append(cur)
            for ob in src_of[cur["id"]]:
                comp_s[ob["id"]] = ob
                for other in edited:
                    if other["id"] not in seen and any(o["id"] == ob["id"] for o in src_of[other["id"]]):
                        seen.add(other["id"])
                        stack.append(other)
        comps.append((comp_e, list(comp_s.values())))
    new_blocks = [copy.deepcopy(nb) for nb in new_initial]
    moved, skipped = 0, []
    gid_map = {}
    for comp_e, comp_s in comps:
        if not comp_s or any(ob["id"] not in pairs for ob in comp_s):
            skipped += [(b.get("title") or "").replace("\n", " ")[:20] or "(이름 없는 칸)" for b in comp_e]
            continue
        o = [min(ob["rect"][0] for ob in comp_s), min(ob["rect"][1] for ob in comp_s),
             max(ob["rect"][2] for ob in comp_s), max(ob["rect"][3] for ob in comp_s)]
        ns = [pairs[ob["id"]] for ob in comp_s]
        n = [min(b["rect"][0] for b in ns), min(b["rect"][1] for b in ns), max(b["rect"][2] for b in ns), max(b["rect"][3] for b in ns)]

        def mx(x):
            return n[0] + (x - o[0]) * ((n[2] - n[0]) / (o[2] - o[0]) if o[2] > o[0] else 1.0)

        def my(y):
            return n[1] + (y - o[1]) * ((n[3] - n[1]) / (o[3] - o[1]) if o[3] > o[1] else 1.0)

        drop = {b["id"] for b in ns}
        new_blocks = [b for b in new_blocks if b["id"] not in drop]
        for e in comp_e:
            c = copy.deepcopy(e)
            c["id"] = next_id
            next_id += 1
            x0, y0, x1, y1 = e["rect"]
            c["rect"] = (round(mx(x0), 2), round(my(y0), 2), round(mx(x1), 2), round(my(y1), 2))
            c["orig_rect"] = c["rect"]
            if e.get("group_id") is not None:
                if e["group_id"] not in gid_map:
                    gid_map[e["group_id"]] = next_id
                    next_id += 1
                c["group_id"] = gid_map[e["group_id"]]
            # 같은 자리 원본의 테두리 모양은 새 판 것을 쓴다
            if ns:
                c["border_color"] = ns[0].get("border_color", c.get("border_color"))
                c["border_width"] = ns[0].get("border_width", c.get("border_width"))
            new_blocks.append(c)
            moved += 1
    new_blocks.sort(key=lambda b: (b["rect"][1], b["rect"][0]))
    return new_blocks, moved, skipped, changed, next_id


def get_system_korean_font_path(custom_path=None):
    if custom_path and os.path.exists(custom_path) and custom_path.lower().endswith(('.ttf', '.otf', '.ttc')):
        return custom_path
        
    try:
        reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                val_name, val_data, _ = winreg.EnumValue(key, i)
                if any(k in val_name.lower() for k in ["malgun", "맑은 고딕", "nanum", "나눔고딕", "gulim", "굴림", "batang", "바탕"]):
                    if val_data.lower().endswith(('.ttf', '.ttc', '.otf')):
                        if not os.path.isabs(val_data):
                            full_path = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", val_data)
                        else:
                            full_path = val_data
                        if os.path.exists(full_path):
                            return full_path
    except Exception:
        pass

    candidates = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/malgunbd.ttf",
        "C:/Windows/Fonts/gulim.ttc",
        "C:/Windows/Fonts/batang.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf"
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
            
    return None


def get_tk_font_family(font_path):
    if not font_path or not os.path.exists(font_path):
        return "맑은 고딕"
    
    font_name = os.path.splitext(os.path.basename(font_path))[0].lower()
    if "malgun" in font_name:
        return "맑은 고딕"
    elif "gulim" in font_name:
        return "굴림"
    elif "batang" in font_name:
        return "바탕"
    elif "nanum" in font_name:
        return "나눔고딕"
    elif "gothic" in font_name:
        return "맑은 고딕"
    
    return "맑은 고딕"


class ScheduleEditor:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1480x1020")
        root.minsize(1000, 700)
        
        self.pdf_path = None
        self.page_count = 0
        self.current_page = 0
        self.zoom = 1.0  
        self.zoom_version = 0
        self.saved_view = (0.0, 0.0)
        self.png_data = None
        self.page_rect = None
        
        self.zoom_label = tk.StringVar(value="100%")
        self.manual_draw_mode = tk.BooleanVar(value=False)
        self.manual_rect_start = None
        self.temp_draw_id = None
        
        self.page_blocks = {}         
        self.initial_blocks = {} 
        self.blocks = [] 
        
        self.selected = set()    
        self.history = []
        self.copied_blocks = [] 
        
        self.h_split_count = 1
        self.v_split_count = 1
        self.split_ratios = [] 
        self.split_titles = {}
        self.split_title_vars = {}
        self._split_preview_job = None 
        
        self.badge_folder = schedule_badge_dir()
        if not os.path.exists(self.badge_folder):
            try:
                os.makedirs(self.badge_folder)
            except Exception:
                pass

        self.custom_badge_paths = {
            "uhd": os.path.join(self.badge_folder, "uhd.png"),
            "all": os.path.join(self.badge_folder, "all.png"),
            "7": os.path.join(self.badge_folder, "7.png"),
            "12": os.path.join(self.badge_folder, "12.png"),
            "15": os.path.join(self.badge_folder, "15.png"),
            "19": os.path.join(self.badge_folder, "19.png"),
            "sub": os.path.join(self.badge_folder, "sub.png"),
            "sign": os.path.join(self.badge_folder, "sign.png"),
            "des": os.path.join(self.badge_folder, "des.png"),
            "local": os.path.join(self.badge_folder, "local.png"),
        }

        self.global_font_path = get_system_korean_font_path()
        self.ui_font_family = get_tk_font_family(self.global_font_path)

        self.split_styles = []          # 작업 스타일: [{"name","keyword","axis","cuts","parts"}]
        self.split_styles_auto = False
        self.favorite_presets = [
            {
                "title": "일반프로그램", 
                "color": "", 
                "text_color": "#000000", 
                "h": 50,
                "font_size_offset": 0,
                "options": {
                    "top_time_enabled": False, "top_time_val": "00", "top_time_font_size": DEFAULT_TIME_FONT_SIZE,
                    "bottom_time_enabled": False, "bottom_time_val": "00", "bottom_time_font_size": DEFAULT_TIME_FONT_SIZE,
                    "badge_size": 12, "uhd": False, "sub": False, "sign": False, "des": False, "local": False, "age": "none"
                }
            },
            {
                "title": "주요뉴스", 
                "color": "", 
                "text_color": "#000000", 
                "h": 40,
                "font_size_offset": 0,
                "options": {
                    "top_time_enabled": False, "top_time_val": "00", "top_time_font_size": DEFAULT_TIME_FONT_SIZE,
                    "bottom_time_enabled": False, "bottom_time_val": "00", "bottom_time_font_size": DEFAULT_TIME_FONT_SIZE,
                    "badge_size": 12, "uhd": False, "sub": False, "sign": False, "des": False, "local": False, "age": "none"
                }
            },
        ]
        
        self.drag = None
        self.drag_start_pos = None 
        self.item_map = {}
        self.next_id = 1
        
        # 수정된 블록 표시 설정
        self.mod_shade_enabled = True          # 반투명 음영 표시
        self.mod_shade_color = "#facc15"       # 음영 색 (사용자 선택)
        self.mod_shade_alpha = 0.35            # 음영 진하기 (0~1)
        self.mod_shade_in_export = False       # 저장 PDF에도 음영 포함 (기본: 미리보기에만)
        self.shade_target = "mod"              # 음영 넣을 칸: mod=수정한 칸 / unmod=수정 안 한 칸 / both=둘 다(색 따로)
        self.unmod_shade_color = "#60a5fa"     # 수정 안 한 칸 음영 색
        self.pick_shade_last = "#fde047"       # 고른 칸 음영: 마지막에 쓴 색
        self.shade_color_counts = {}           # 칸 음영 색별로 쓴 횟수 -> '자주 쓴 색' 줄
        self.pick_shade_in_export = True       # 고른 칸 음영을 저장 PDF에도 넣기
        self.mod_border_width = None           # 수정 블록 공통 테두리 두께(pt), None = 원본 유지
        self._mod_ids = set()

        # 블록 위치 고정: 칸은 옮기지 않고 합치기·나누기·내용 수정만 한다 (기본 켜짐)
        self.lock_move = tk.BooleanVar(value=True)
        self.lock_move.trace_add("write", lambda *a: self._on_lock_changed())
        self._lock_hint_shown = False
        # 나누기(2분할, 5분 단위) · 칸 이름 · 머리글
        self.split_mode_var = tk.BooleanVar(value=False)
        self.split_mode_var.trace_add("write", lambda *a: self._on_split_mode_var())
        self.time_rows_by_page = {}
        self.time_ticks_by_page = {}
        self.header_found_by_page = {}
        self.header_edit = {}               # 페이지 -> 제목·날짜 바꾸기 설정 (작업 내용, 자동 저장)
        self.header_title_presets = []      # 저장해 둔 머리글 제목
        self.title_presets = []             # 저장해 둔 칸 이름
        self.recent_titles = []             # 최근에 넣은 칸 이름
        self.dmb_regions = list(DMB_REGIONS)    # DMB 요일 칸 지역 목록
        self.dmb_day_regions = []               # DMB 요일 칸마다 지난번에 넣은 지역 (다음 주에 그대로 제안)
        self.last_save_dir = ""                 # 마지막으로 PDF 를 저장한 폴더 (편성표·DMB 따로)
        self._shared_last = None
        self.changed_ids_by_page = {}           # 새 판에서 원본이 바뀐 칸 (이전 판과 비교) - 빨간 점선
        self.show_changed = True
        self._offered_carry_for = None
        # 블록 찾기 최소 크기(pt)
        self.detect_min_w = DETECT_MIN_W
        self.detect_min_h = DETECT_MIN_H

        self.backup_file_path = schedule_settings_path()
        
        self._recovery = None
        self._clean_exit = False
        self.load_autosave()
        
        self._build_ui()
        self.draw_blank_canvas()
        if self._recovery:
            self.root.after(400, self._offer_recovery)

    def _build_ui(self):
        self.toolbar = ttk.Frame(self.root, padding=(8, 6))
        self.toolbar.pack(side="top", fill="x")
        self._build_default_toolbar()

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

        self.root.bind("<Delete>", lambda e: self.delete_selected_blocks())
        self.root.bind("<Control-z>", lambda e: None if self._typing() else self.undo_action())
        self.root.bind("<Control-c>", lambda e: None if self._typing() else self.copy_selected_blocks())
        self.root.bind("<Control-v>", lambda e: None if self._typing() else self.paste_blocks())
        self.root.bind("<Control-a>", lambda e: None if self._typing() else self.select_all_blocks())
        self.root.bind("<KeyPress-s>", lambda e: None if self._typing() else self.toggle_split_mode())
        self.root.bind("<KeyPress-S>", lambda e: None if self._typing() else self.toggle_split_mode())
        self.root.bind("<KeyPress-h>", lambda e: None if self._typing() else self.toggle_pick_shade())
        self.root.bind("<KeyPress-H>", lambda e: None if self._typing() else self.toggle_pick_shade())
        self.root.bind("<Escape>", lambda e: self.toggle_split_mode(False) if self.split_mode_var.get() else None)
        
        self.root.bind("<Left>", lambda e: self.nudge_selected_block(-1, 0, 0, 0))
        self.root.bind("<Right>", lambda e: self.nudge_selected_block(1, 0, 0, 0))
        self.root.bind("<Up>", lambda e: self.nudge_selected_block(0, -1, 0, 0))
        self.root.bind("<Down>", lambda e: self.nudge_selected_block(0, 1, 0, 0))
        self.root.bind("<Shift-Left>", lambda e: self.nudge_selected_block(0, 0, -1, 0))
        self.root.bind("<Shift-Right>", lambda e: self.nudge_selected_block(0, 0, 1, 0))
        self.root.bind("<Shift-Up>", lambda e: self.nudge_selected_block(0, 0, 0, -1))
        self.root.bind("<Shift-Down>", lambda e: self.nudge_selected_block(0, 0, 0, 1))

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing_app)

    def _build_sidebar_scroll_wrapper(self):
        for widget in self.side.winfo_children():
            widget.destroy()

        self.side_canvas = tk.Canvas(self.side, highlightthickness=0, background="#f1f5f9")
        self.side_scrollbar = ttk.Scrollbar(self.side, orient="vertical", command=self.side_canvas.yview)
        
        self.sidebar_content_frame = ttk.Frame(self.side_canvas, padding=10)
        
        self.sidebar_content_frame.bind(
            "<Configure>",
            lambda e: self.side_canvas.configure(scrollregion=self.side_canvas.bbox("all"))
        )
        
        self.side_window_id = self.side_canvas.create_window((0, 0), window=self.sidebar_content_frame, anchor="nw")
        self.side_canvas.configure(yscrollcommand=self.side_scrollbar.set)
        
        def on_side_canvas_configure(event):
            self.side_canvas.itemconfig(self.side_window_id, width=event.width)
        self.side_canvas.bind("<Configure>", on_side_canvas_configure)

        self.side_canvas.pack(side="left", fill="both", expand=True)
        self.side_scrollbar.pack(side="right", fill="y")

        def _on_side_mousewheel(event):
            if event.delta:
                self.side_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self.side_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.side_canvas.yview_scroll(1, "units")

        self.side_canvas.bind_all("<MouseWheel>", _on_side_mousewheel, add="+")
        self.side_canvas.bind_all("<Button-4>", _on_side_mousewheel, add="+")
        self.side_canvas.bind_all("<Button-5>", _on_side_mousewheel, add="+")

    def _typing(self):
        w = self.root.focus_get()
        return isinstance(w, (tk.Entry, ttk.Entry, tk.Text, ttk.Combobox, ttk.Spinbox, ttk.Scale))

    def clear_sidebar(self):
        for widget in self.sidebar_content_frame.winfo_children():
            widget.destroy()

    def _on_lock_changed(self):
        self.drag = None
        if not hasattr(self, "toolbar"):          # 시작하면서 설정을 읽는 중에는 저장·그리기를 하지 않는다
            return
        try:
            if self.pdf_path:
                self.draw_pdf(refresh_sidebar=True)
            self.save_autosave()
        except Exception:
            pass

    def _move_locked(self, show=True):
        """위치 고정 중이면 True (처음 한 번만 안내)."""
        if not self.lock_move.get():
            return False
        if show and not self._lock_hint_shown:
            self._lock_hint_shown = True
            messagebox.showinfo("블록 위치 고정", "블록 위치 고정이 켜져 있어 옮기거나 크기를 바꿀 수 없습니다.\n"
                                "(합치기·나누기·내용 수정은 그대로 됩니다)\n\n위치를 바꿔야 할 때만 도구 모음의 [🔒 블록 위치 고정]을 끄세요.")
        return True

    def nudge_selected_block(self, dx, dy, dw, dh):
        if len(self.selected) != 1:
            return
        if self._typing() or self._move_locked():
            return
        if self._typing():
            return
            
        sel_id = list(self.selected)[0]
        block = next((b for b in self.blocks if b["id"] == sel_id), None)
        if not block:
            return
            
        self._push_coalesced("nudge")
        gid = block.get("group_id")
        group_blocks = [gb for gb in self.blocks if gid is not None and gb.get("group_id") == gid]
        if group_blocks:   # 분할 블록: 그룹 전체가 같이 이동 / 크기 조절
            gx0 = min(gb["rect"][0] for gb in group_blocks)
            gy0 = min(gb["rect"][1] for gb in group_blocks)
            gx1 = max(gb["rect"][2] for gb in group_blocks)
            gy1 = max(gb["rect"][3] for gb in group_blocks)
        else:
            gx0, gy0, gx1, gy1 = block["rect"]
        if dw != 0 or dh != 0:
            new_rect = (gx0, gy0, max(gx0 + 10, gx1 + dw), max(gy0 + 10, gy1 + dh))
            if group_blocks:
                self._apply_group_rect(group_blocks, new_rect)
            else:
                block["rect"] = new_rect
        elif group_blocks:
            for gb in group_blocks:
                bx0, by0, bx1, by1 = gb["rect"]
                gb["rect"] = (bx0 + dx, by0 + dy, bx1 + dx, by1 + dy)
        else:
            block["rect"] = (gx0 + dx, gy0 + dy, gx1 + dx, gy1 + dy)
            
        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()

    def _exit_clean(self):
        self._clean_exit = True
        self.save_autosave()
        self.root.destroy()

    def _work_sig(self):
        """지금 작업 상태(칸·음영·제목·날짜·요일 칸)를 한 줄로 - 마지막 저장 때와 비교해 바뀐 게 있는지 본다."""
        try:
            return json.dumps([self.pdf_path, self._snapshot_blocks(), self.header_edit], sort_keys=True, default=str)
        except Exception:
            return None

    def _mark_saved(self):
        self._saved_sig = self._work_sig()

    def _has_unsaved_work(self):
        if not self.pdf_path:
            return False
        sig = self._work_sig()
        return sig is None or sig != getattr(self, "_saved_sig", None)

    def on_closing_app(self):
        """X 를 눌러도 바로 닫지 않고 물어본다. 저장 안 한 작업이 있으면 저장할지도 묻는다."""
        self.save_autosave()
        name = "DMB 편성 편집기" if IS_DMB else "TV 편성표 편집기"
        if self._has_unsaved_work():
            ans = messagebox.askyesnocancel(
                "종료 확인",
                "저장하지 않은 작업이 있습니다.\n\nPDF로 저장하고 종료할까요?\n\n"
                "  예 : 저장하고 종료\n  아니오 : 저장하지 않고 종료\n  취소 : 계속 작업",
                icon="warning", default="cancel", parent=self.root)
            if ans is True:
                if self.export_pdf():
                    self._exit_clean()
            elif ans is False:
                self._exit_clean()
            return
        if messagebox.askokcancel("종료 확인", f"{name}를 종료할까요?", default="cancel", parent=self.root):
            self._exit_clean()

    def push_history(self):
        self._coal = None
        self._push_snapshot(self._snapshot_blocks())

    def _push_coalesced(self, key, window=1.0):
        """같은 종류의 변경이 window초 안에 이어지면(키 반복, 연속 클릭, 타이핑) 되돌리기 기록을 한 번만 남긴다."""
        now = time.time()
        last = getattr(self, "_coal", None)
        if last and last[0] == key and now - last[1] < window:
            self._coal = (key, now)
            return
        self.push_history()
        self._coal = (key, now)

    def _push_snapshot(self, snap):
        self.history.append(snap)
        if len(self.history) > 30:
            self.history.pop(0)
        self.save_autosave()

    def _snapshot_blocks(self):
        state_snapshot = [
            {
                "id": b["id"],
                "title": b["title"],
                "color": b["color"],
                "text_color": b.get("text_color", "#000000"),
                "font_size_offset": b.get("font_size_offset", 0),
                "text_dx": b.get("text_dx", 0),
                "text_dy": b.get("text_dy", 0),
                "options": copy.deepcopy(b.get("options", {})),
                "rect": tuple(b["rect"]),
                "orig_rect": tuple(b["orig_rect"]),
                "group_id": b.get("group_id"),
                "rel_box": b.get("rel_box"),
                "border_color": b.get("border_color"),
                "border_width": b.get("border_width"),
                "orig_fill": b.get("orig_fill", ""),
                "shade": b.get("shade"),
                "split_dash": b.get("split_dash", False),
            }
            for b in self.blocks
        ]
        return state_snapshot

    def save_autosave(self):
        try:
            clean = getattr(self, "_clean_exit", False)
            data = {
                "pdf_path": None if clean else self.pdf_path,
                "current_page": self.current_page,
                "page_blocks": {} if clean else {str(k): v for k, v in self.page_blocks.items()},
                "initial_blocks": {} if clean else {str(k): v for k, v in self.initial_blocks.items()},
                "next_id": self.next_id,
                "custom_badge_paths": self.custom_badge_paths,
                "global_font_path": self.global_font_path,
                "favorite_presets": self.favorite_presets,
                "split_styles": self.split_styles,
                "split_styles_auto": self.split_styles_auto,
                "mod_shade_enabled": self.mod_shade_enabled,
                "mod_shade_color": self.mod_shade_color,
                "mod_shade_alpha": self.mod_shade_alpha,
                "mod_shade_in_export": self.mod_shade_in_export,
                "shade_target": self.shade_target,
                "unmod_shade_color": self.unmod_shade_color,
                "pick_shade_last": self.pick_shade_last,
                "shade_color_counts": self.shade_color_counts,
                "pick_shade_in_export": self.pick_shade_in_export,
                "mod_border_width": self.mod_border_width,
                "lock_move": bool(self.lock_move.get()),
                "header_edit": {} if clean else {str(k): v for k, v in self.header_edit.items()},
                "header_title_presets": self.header_title_presets,
                "title_presets": self.title_presets,
                "recent_titles": self.recent_titles,
                "detect_min_w": self.detect_min_w,
                "detect_min_h": self.detect_min_h,
                "dmb_regions": self.dmb_regions,
                "dmb_day_regions": self.dmb_day_regions,
                "last_save_dir": self.last_save_dir,
            }
            text = json.dumps(data, ensure_ascii=False)   # 먼저 문자열로 만들어 두어야 직렬화 실패 시 기존 파일이 깨지지 않는다
            tmp_path = self.backup_file_path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(text)
            os.replace(tmp_path, self.backup_file_path)
        except Exception:
            pass
        self._save_shared()
        self._save_work_record()

    # ------------------------------------------------------------ PDF 별 작업 기록 (다음 판으로 옮길 때 씀)
    def _work_dir(self):
        base = os.path.dirname(os.path.abspath(self.backup_file_path))
        name = "작업기록" if os.environ.get("RADIO_HOME") else f".{PROFILE}_작업기록"
        return os.path.join(base, name)

    def _save_work_record(self):
        try:
            if getattr(self, "_clean_exit", False) or not self.pdf_path or not self.page_blocks:
                return
            pages = {}
            for pg, blocks in self.page_blocks.items():
                init = self.initial_blocks.get(pg) or []
                if any(_is_edited(b, init) for b in blocks):
                    pages[str(pg)] = {"initial": init, "blocks": blocks}
            hdr = {str(k): v for k, v in self.header_edit.items()}
            if not pages:
                return
            d = self._work_dir()
            os.makedirs(d, exist_ok=True)
            fn = re.sub(r'[\\/:*?"<>|]', "_", os.path.basename(self.pdf_path)) + ".json"
            data = {"pdf_name": os.path.basename(self.pdf_path), "pdf_path": self.pdf_path,
                    "saved_at": datetime.datetime.now().isoformat(timespec="seconds"), "pages": pages, "header_edit": hdr}
            text = json.dumps(data, ensure_ascii=False, default=list)
            with open(os.path.join(d, fn + ".tmp"), "w", encoding="utf-8") as f:
                f.write(text)
            os.replace(os.path.join(d, fn + ".tmp"), os.path.join(d, fn))
            olds = sorted(glob.glob(os.path.join(d, "*.json")), key=os.path.getmtime)
            for old in olds[:-80]:
                os.remove(old)
        except Exception:
            pass

    def _list_work_records(self):
        out = []
        for path in glob.glob(os.path.join(self._work_dir(), "*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and isinstance(data.get("pages"), dict):
                    data["_file"] = path
                    out.append(data)
            except Exception:
                continue
        out.sort(key=lambda d: d.get("saved_at", ""), reverse=True)
        return out

    @staticmethod
    def _week_key(name):
        m = re.search(r"_(\d{1,2})__(\d)", name or "")
        return (int(m.group(1)), int(m.group(2))) if m else None

    def _carry_candidate(self):
        """지금 연 PDF 에 옮겨 올 만한 이전 판 작업: 같은 주(파일 이름의 월·주) -> 없으면 최근 3주 안의 가장 최근 것."""
        me = os.path.basename(self.pdf_path or "")
        recs = [r for r in self._list_work_records() if str(self.current_page) in r["pages"]]
        same = [r for r in recs if r.get("pdf_name") != me and self._week_key(r.get("pdf_name")) and self._week_key(r.get("pdf_name")) == self._week_key(me)]
        if same:
            return same[0]
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=21)).isoformat()
        recent = [r for r in recs if r.get("pdf_name") != me and r.get("saved_at", "") >= cutoff]
        return recent[0] if recent else None

    def apply_work_record(self, rec, quiet=False):
        """이전 판 작업 기록을 지금 페이지에 옮긴다."""
        pg = str(self.current_page)
        page = (rec.get("pages") or {}).get(pg)
        init_new = self.initial_blocks.get(self.current_page) or []
        if not page or not init_new:
            if not quiet:
                messagebox.showinfo("이전 판 가져오기", "옮길 작업이 없습니다.", parent=self.root)
            return None
        for b in page["initial"] + page["blocks"]:
            b["rect"] = tuple(b["rect"])
        self.push_history()
        new_blocks, moved, skipped, changed, self.next_id = transplant_work(page["initial"], page["blocks"], init_new, self.next_id)
        self.blocks = new_blocks
        self.page_blocks[self.current_page] = self.blocks
        for gid in {b.get("group_id") for b in self.blocks if b.get("group_id") is not None}:
            self._recompute_rel_boxes(gid)
        self.changed_ids_by_page[self.current_page] = set(changed)
        self.show_changed = True
        # 머리글: 지역·제목 크기·굵게는 이어받고, 날짜는 새 판 것을 쓴다
        old_h = (rec.get("header_edit") or {}).get(pg) or {}
        st = self._header_state()
        for k in ("title_size", "title_bold", "date_bold"):
            if k in old_h:
                st[k] = old_h[k]
        if isinstance(old_h.get("days"), list) and len(old_h["days"]) == len(st.get("days") or []):
            for v, ov in zip(st["days"], old_h["days"]):
                if isinstance(ov, dict) and ov.get("region"):
                    v["region"] = ov["region"]
            st["days_on"] = bool(old_h.get("days_on", st.get("days_on")))
        self.selected = set()
        self.draw_pdf()
        try:
            self._build_pdf_sidebar()          # 머리글 지역·바뀐 칸 표시 줄 다시 그리기
        except Exception:
            pass
        self.save_autosave()
        msg = (f"'{rec.get('pdf_name')}' 에서 고친 칸 {moved}개를 옮겼습니다.\n"
               f"이 판에서 원본이 바뀐 칸: {len(changed)}개 (빨간 점선으로 표시)")
        if skipped:
            msg += (f"\n\n원본이 바뀌어 옮기지 않은 칸 {len(skipped)}개 - 다시 확인해 주세요:\n  "
                    + "\n  ".join(skipped[:12]) + ("\n  …" if len(skipped) > 12 else ""))
        if not quiet:
            messagebox.showinfo("이전 판 가져오기", msg, parent=self.root)
        return moved, skipped, changed

    def offer_carry_over(self):
        if not self.pdf_path or self._offered_carry_for == (self.pdf_path, self.current_page):
            return
        self._offered_carry_for = (self.pdf_path, self.current_page)
        if any(_is_edited(b, self.initial_blocks.get(self.current_page) or []) for b in self.blocks):
            return                               # 이미 작업 중(복구한 것 등)이면 묻지 않는다
        rec = self._carry_candidate()
        if not rec:
            return
        n = sum(1 for b in rec["pages"][str(self.current_page)]["blocks"] if _is_edited(b, rec["pages"][str(self.current_page)]["initial"]))
        when = rec.get("saved_at", "").replace("T", " ")[:16]
        if messagebox.askyesno("이전 판 작업 가져오기",
                               f"이전 판 '{rec.get('pdf_name')}' ({when} 작업)에서 고친 칸 {n}개가 있습니다.\n\n"
                               "이 판으로 옮길까요?\n(원본이 그대로인 칸만 옮기고, 이 판에서 바뀐 칸은 빨간 점선으로 표시합니다)",
                               parent=self.root):
            self.apply_work_record(rec)

    def choose_work_record(self):
        """툴바 [📥 이전 판 가져오기]: 작업 기록 목록에서 골라 옮긴다."""
        if not self.pdf_path:
            return
        me = os.path.basename(self.pdf_path)
        recs = [r for r in self._list_work_records() if r.get("pdf_name") != me and str(self.current_page) in r["pages"]]
        if not recs:
            messagebox.showinfo("이전 판 가져오기", "옮길 수 있는 작업 기록이 아직 없습니다.\n(PDF를 고치면 파일마다 자동으로 기록됩니다)", parent=self.root)
            return
        win = tk.Toplevel(self.root)
        win.title("이전 판 작업 가져오기")
        win.transient(self.root)
        win.geometry("520x360")
        ttk.Label(win, text="어느 PDF에서 고친 내용을 이 판으로 옮길까요? (위가 최근)", padding=(10, 8, 10, 2)).pack(anchor="w")
        lb = tk.Listbox(win, font=(self.ui_font_family, 10), activestyle="dotbox")
        lb.pack(fill="both", expand=True, padx=10, pady=4)
        wk = self._week_key(me)
        for r in recs[:40]:
            tag = "  ★ 같은 주" if wk and self._week_key(r.get("pdf_name")) == wk else ""
            lb.insert("end", f"{r.get('saved_at', '').replace('T', ' ')[:16]}   {r.get('pdf_name')}{tag}")
        lb.selection_set(0)

        def go(*_a):
            sel = lb.curselection()
            if not sel:
                return
            win.destroy()
            if any(_is_edited(b, self.initial_blocks.get(self.current_page) or []) for b in self.blocks):
                if not messagebox.askyesno("이전 판 가져오기", "지금 화면에서 고친 내용이 있습니다.\n이전 판 작업으로 바꿀까요? (Ctrl+Z 로 되돌릴 수 있음)", parent=self.root):
                    return
                self.blocks = copy.deepcopy(self.initial_blocks.get(self.current_page) or [])
                self.page_blocks[self.current_page] = self.blocks
            self.apply_work_record(recs[sel[0]])

        lb.bind("<Double-Button-1>", go)
        bf = ttk.Frame(win, padding=(10, 0, 10, 10))
        bf.pack(fill="x")
        ttk.Button(bf, text="옮기기", command=go).pack(side="right")
        ttk.Button(bf, text="닫기", command=win.destroy).pack(side="right", padx=6)

    def _shared_values(self):
        return {k: copy.deepcopy(getattr(self, k)) for k in SHARED_KEYS}

    def _save_shared(self):
        """같이 쓰는 설정은 이 프로그램에서 바뀐 것만 공용 파일에 덮어쓴다 (다른 편집기가 바꾼 것은 그대로 둔다)."""
        try:
            cur = self._shared_values()
            last = self._shared_last or {}
            changed = {k: v for k, v in cur.items() if k not in last or last[k] != v}
            if not changed and self._shared_last is not None:
                return
            on_disk = {}
            try:
                with open(shared_settings_path(), "r", encoding="utf-8") as f:
                    on_disk = json.load(f)
                if not isinstance(on_disk, dict):
                    on_disk = {}
            except (OSError, ValueError):
                on_disk = {}
            merged = {k: on_disk[k] for k in SHARED_KEYS if k in on_disk}
            merged.update(changed if on_disk else cur)
            write_shared_settings(merged)
            self._shared_last = cur
        except Exception:
            pass

    def load_autosave(self):
        data = {}
        try:
            if os.path.exists(self.backup_file_path):
                with open(self.backup_file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
        except Exception:
            data = {}
        if not isinstance(data, dict):
            data = {}
        data.update(read_shared_settings())     # 같이 쓰는 것은 공용 파일이 우선
        try:
            if data:
                _pb = data.get("page_blocks")
                if data.get("pdf_path") and isinstance(_pb, dict) and any(_pb.values()):
                    self._recovery = data
                if "custom_badge_paths" in data and isinstance(data["custom_badge_paths"], dict):
                    for k, v in data["custom_badge_paths"].items():
                        if v and not os.path.isabs(v):
                            v = os.path.join(self.badge_folder, os.path.basename(v))   # 예전 'badges\\xx.png' -> 뱃지 폴더
                        self.custom_badge_paths[k] = v
                if "global_font_path" in data and data["global_font_path"]:
                    self.global_font_path = data["global_font_path"]
                    self.ui_font_family = get_tk_font_family(self.global_font_path)
                if "next_id" in data:
                    self.next_id = data["next_id"]
                if "favorite_presets" in data and isinstance(data["favorite_presets"], list):
                    self.favorite_presets = data["favorite_presets"]
                if isinstance(data.get("mod_shade_color"), str) and re.fullmatch(r"#[0-9a-fA-F]{6}", data["mod_shade_color"]):
                    self.mod_shade_color = data["mod_shade_color"]
                if isinstance(data.get("mod_shade_alpha"), (int, float)):
                    self.mod_shade_alpha = max(0.05, min(0.8, float(data["mod_shade_alpha"])))
                if isinstance(data.get("mod_shade_enabled"), bool):
                    self.mod_shade_enabled = data["mod_shade_enabled"]
                if isinstance(data.get("mod_shade_in_export"), bool):
                    self.mod_shade_in_export = data["mod_shade_in_export"]
                if isinstance(data.get("shade_color_counts"), dict):
                    self.shade_color_counts = {k.lower(): int(v) for k, v in data["shade_color_counts"].items()
                                               if isinstance(k, str) and re.fullmatch(r"#[0-9a-fA-F]{6}", k) and isinstance(v, (int, float))}
                if isinstance(data.get("pick_shade_last"), str) and re.fullmatch(r"#[0-9a-fA-F]{6}", data["pick_shade_last"]):
                    self.pick_shade_last = data["pick_shade_last"]
                if isinstance(data.get("pick_shade_in_export"), bool):
                    self.pick_shade_in_export = data["pick_shade_in_export"]
                if data.get("shade_target") in ("mod", "unmod", "both"):
                    self.shade_target = data["shade_target"]
                if isinstance(data.get("unmod_shade_color"), str) and re.fullmatch(r"#[0-9a-fA-F]{6}", data["unmod_shade_color"]):
                    self.unmod_shade_color = data["unmod_shade_color"]
                for key in ("header_title_presets", "title_presets", "recent_titles"):
                    v = data.get(key)
                    if isinstance(v, list):
                        setattr(self, key, [str(x) for x in v if isinstance(x, str)][:200])
                if isinstance(data.get("split_styles"), list):
                    self.split_styles = [x for x in data["split_styles"] if isinstance(x, dict) and x.get("parts")]
                if isinstance(data.get("split_styles_auto"), bool):
                    self.split_styles_auto = data["split_styles_auto"]
                if isinstance(data.get("lock_move"), bool):
                    self.lock_move.set(data["lock_move"])
                for key, lo, hi in (("detect_min_w", 5.0, 200.0), ("detect_min_h", 1.0, 100.0)):
                    v = data.get(key)
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        setattr(self, key, max(lo, min(hi, float(v))))
                mbw = data.get("mod_border_width")
                if isinstance(mbw, (int, float)) and not isinstance(mbw, bool):
                    self.mod_border_width = max(0.25, min(6.0, float(mbw)))
                if isinstance(data.get("last_save_dir"), str):
                    self.last_save_dir = data["last_save_dir"]
                for key in ("dmb_regions", "dmb_day_regions"):
                    v = data.get(key)
                    if isinstance(v, list):
                        setattr(self, key, [str(x) for x in v if isinstance(x, str)][:50])
        except Exception:
            pass
        self._shared_last = self._shared_values()

    def undo_action(self):
        if not self.history:
            messagebox.showinfo("되돌리기", "더 이상 되돌릴 작업이 없습니다.")
            return
        self.blocks = self.history.pop()
        self._coal = None
        self.page_blocks[self.current_page] = self.blocks
        self.selected.clear()
        self.draw_pdf()
        self.save_autosave()

    def copy_selected_blocks(self):
        if not self.selected:
            return
        self.copied_blocks = [copy.deepcopy(b) for b in self.blocks if b["id"] in self.selected]
        messagebox.showinfo("복사", f"{len(self.copied_blocks)}개의 블록이 클립보드에 복사되었습니다.")

    def paste_blocks(self):
        if not self.copied_blocks:
            return
        ans = messagebox.askyesno("붙여넣기 확인", "복사한 블록을 현재 페이지에 붙여넣으시겠습니까?")
        if not ans:
            return
            
        self.push_history()
        self.selected.clear()
        
        newly_selected = set()
        for cb in self.copied_blocks:
            cb_copy = copy.deepcopy(cb)
            cb_copy["id"] = self.next_id
            rx0, ry0, rx1, ry1 = cb_copy["rect"]
            offset = 15
            cb_copy["rect"] = (rx0 + offset, ry0 + offset, rx1 + offset, ry1 + offset)
            cb_copy["orig_rect"] = cb_copy["rect"]
            cb_copy["group_id"] = None
            cb_copy["rel_box"] = None
            
            self.blocks.append(cb_copy)
            newly_selected.add(self.next_id)
            self.next_id += 1
            
        self.selected = newly_selected
        self.draw_pdf()
        self.save_autosave()

    def select_all_blocks(self):
        if not self.blocks:
            return
        self.selected = {b["id"] for b in self.blocks}
        self.draw_pdf()

    def clear_toolbar(self):
        for widget in self.toolbar.winfo_children():
            widget.destroy()

    def _build_default_toolbar(self):
        self.clear_toolbar()
        ttk.Button(self.toolbar, text="PDF 열기", command=self.open_pdf).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="⚙️ 설정 (폰트/뱃지)", command=self.open_settings_window).pack(side="left", padx=3)
        ttk.Label(self.toolbar, text=f"  [v{APP_VERSION}] 상단의 'PDF 열기'를 눌러 시작하세요.").pack(side="left", padx=12)

    def _build_pdf_toolbar(self):
        self.clear_toolbar()
        ttk.Button(self.toolbar, text="PDF 열기", command=self.open_pdf).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="⚙️ 설정 (폰트/뱃지)", command=self.open_settings_window).pack(side="left", padx=3)
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        
        ttk.Checkbutton(self.toolbar, text="🔒 블록 위치 고정", variable=self.lock_move).pack(side="left", padx=5)
        ttk.Checkbutton(self.toolbar, text="✂ 나누기 (S)", variable=self.split_mode_var).pack(side="left", padx=5)
        chk_draw = ttk.Checkbutton(self.toolbar, text="✏️ 수동 블록 그리기", variable=self.manual_draw_mode)
        chk_draw.pack(side="left", padx=5)
        
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(self.toolbar, text="↩ 되돌리기 (Ctrl+Z)", command=self.undo_action).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="선택 블록 합치기", command=self.merge_selected).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="삭제 (Del)", command=self.delete_selected_blocks).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="작업 초기화", command=self.clear_all_blocks).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="📥 이전 판 가져오기", command=self.choose_work_record).pack(side="left", padx=3)
        
        ttk.Button(self.toolbar, text="🔍 저장 전 미리보기", command=self.preview_pdf).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text=("💾 DMB PDF 저장" if IS_DMB else "PDF 저장 (다운로드)"), command=self.export_pdf).pack(side="left", padx=3)
        
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(self.toolbar, text="🔍+", command=lambda: self.change_zoom(1)).pack(side="left", padx=2)
        ttk.Label(self.toolbar, textvariable=self.zoom_label, width=5, anchor="center").pack(side="left")
        ttk.Button(self.toolbar, text="🔍-", command=lambda: self.change_zoom(-1)).pack(side="left", padx=2)

    def open_settings_window(self):
        sett_win = tk.Toplevel(self.root)
        sett_win.title("⚙️ 프로그램 설정 (기본 폰트 및 뱃지 지정)")
        sett_win.geometry("560x840")
        sett_win.transient(self.root)
        sett_win.grab_set()

        ttk.Label(sett_win, text="PDF 출력 및 텍스트 렌더링 환경 설정", font=(self.ui_font_family, 11, "bold")).pack(pady=10)

        font_frame = ttk.LabelFrame(sett_win, text=" 기본 한글 폰트 설정 (.ttf 파일) ", padding=10)
        font_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(font_frame, text="편성표 PDF 내보내기 및 화면 표시용 기본 폰트 파일:", font=(self.ui_font_family, 9), foreground="#333").pack(anchor="w", pady=(0, 4))
        
        f_row = ttk.Frame(font_frame)
        f_row.pack(fill="x", pady=2)
        
        self.font_path_var = tk.StringVar(value=self.global_font_path if self.global_font_path else "")
        font_entry = ttk.Entry(f_row, textvariable=self.font_path_var, font=(self.ui_font_family, 8))
        font_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        def browse_font_file():
            path = filedialog.askopenfilename(title="트루타입 폰트 파일 선택", filetypes=[("TrueType Fonts", "*.ttf"), ("Font Files", "*.ttf;*.ttc;*.otf"), ("모든 파일", "*.*")])
            if path:
                self.font_path_var.set(path)
                
        ttk.Button(f_row, text="찾아보기", width=8, command=browse_font_file).pack(side="right")
        ttk.Label(font_frame, text="※ 한글 깨짐 방지를 위해 반드시 .ttf 또는 .ttc 폰트 파일을 지정해 주세요.", font=(self.ui_font_family, 8), foreground="#666").pack(anchor="w", pady=(4, 0))

        badge_frame = ttk.LabelFrame(sett_win, text=" 연령고지 및 방송 옵션 뱃지 이미지 경로 (.png) ", padding=10)
        badge_frame.pack(fill="both", expand=True, padx=15, pady=5)

        self.badge_path_vars = {}
        badges_to_conf = [
            ("UHD 뱃지", "uhd"), ("전연령 뱃지", "all"), ("7세 뱃지", "7"), 
            ("12세 뱃지", "12"), ("15세 뱃지", "15"), ("19세 뱃지", "19"), 
            ("자막해설(sub)", "sub"), ("한국수어(sign)", "sign"), ("화면해설(des)", "des"), ("부분로컬(local)", "local")
        ]
        
        for label_text, key in badges_to_conf:
            row_f = ttk.Frame(badge_frame)
            row_f.pack(fill="x", pady=2)
            
            ttk.Label(row_f, text=label_text, width=13, anchor="w", font=(self.ui_font_family, 8)).pack(side="left")
            current_val = self.custom_badge_paths.get(key, "")
            var = tk.StringVar(value=current_val)
            self.badge_path_vars[key] = var
            
            entry = ttk.Entry(row_f, textvariable=var, font=(self.ui_font_family, 8))
            entry.pack(side="left", fill="x", expand=True, padx=4)
            
            btn = ttk.Button(row_f, text="찾기", width=6, command=lambda k=key, v=var: self.browse_badge_image(k, v))
            btn.pack(side="right")

        det_frame = ttk.LabelFrame(sett_win, text=" 블록 찾기 (PDF를 열 때 칸을 블록으로 만드는 기준) ", padding=8)
        det_frame.pack(fill="x", padx=15, pady=5)
        det_row = ttk.Frame(det_frame)
        det_row.pack(fill="x")
        ttk.Label(det_row, text="최소 너비(pt):", font=(self.ui_font_family, 8)).pack(side="left")
        min_w_var = tk.StringVar(value=f"{self.detect_min_w:g}")
        ttk.Spinbox(det_row, from_=5, to=200, increment=1, width=6, textvariable=min_w_var).pack(side="left", padx=(2, 12))
        ttk.Label(det_row, text="최소 높이(pt):", font=(self.ui_font_family, 8)).pack(side="left")
        min_h_var = tk.StringVar(value=f"{self.detect_min_h:g}")
        ttk.Spinbox(det_row, from_=1, to=100, increment=0.5, width=6, textvariable=min_h_var).pack(side="left", padx=2)
        ttk.Label(det_frame, text=f"※ 기본 너비 {DETECT_MIN_W:g} · 높이 {DETECT_MIN_H:g}. 작은 칸이 빠지면 줄이고, 뱃지 등이 블록으로 잡히면 늘리세요. 다음에 PDF를 열 때 적용됩니다.",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=510).pack(anchor="w", pady=(4, 0))

        def save_and_close():
            for var, key, lo, hi in ((min_w_var, "detect_min_w", 5.0, 200.0), (min_h_var, "detect_min_h", 1.0, 100.0)):
                try:
                    setattr(self, key, max(lo, min(hi, float(var.get()))))
                except ValueError:
                    pass
            self.global_font_path = self.font_path_var.get()
            self.ui_font_family = get_tk_font_family(self.global_font_path)
            for k, v in self.badge_path_vars.items():
                self.custom_badge_paths[k] = v.get()
            self._badge_cache = {}
            self.save_autosave()
            if self.pdf_path:
                self.draw_pdf(refresh_sidebar=True)
            messagebox.showinfo("설정 완료", "기본 폰트 및 설정이 성공적으로 저장되었습니다.")
            sett_win.destroy()

        bf_row = ttk.Frame(sett_win)
        bf_row.pack(fill="x", padx=15)
        ttk.Label(bf_row, text=f"뱃지 폴더: {self.badge_folder}", font=(self.ui_font_family, 8), foreground="#555").pack(side="left")

        def open_badge_folder():
            try:
                os.makedirs(self.badge_folder, exist_ok=True)
                os.startfile(self.badge_folder)
            except Exception:
                pass

        ttk.Button(bf_row, text="폴더 열기", command=open_badge_folder).pack(side="right")
        ttk.Label(sett_win, text="※ 뱃지 PNG를 이 폴더에 uhd.png, all.png, 7.png … 이름으로 넣으면 바로 쓰입니다. (업데이트해도 지워지지 않습니다)",
                  font=(self.ui_font_family, 8), foreground="#666", wraplength=520).pack(anchor="w", padx=15)
        ttk.Button(sett_win, text="설정 저장", command=save_and_close).pack(pady=12)

    def browse_badge_image(self, key, var):
        path = filedialog.askopenfilename(title=f"'{key}' 뱃지 이미지 선택", filetypes=[("PNG 이미지", "*.png"), ("모든 파일", "*.*")])
        if path:
            var.set(path)

    def _build_default_sidebar(self):
        self.clear_sidebar()
        ttk.Label(self.sidebar_content_frame, text="블록", font=(self.ui_font_family, 13, "bold")).pack(anchor="w")
        ttk.Label(self.sidebar_content_frame, text="PDF 파일을 열면 블록 편집 및 파렛트 패널이 활성화됩니다.", foreground="#555", wraplength=340).pack(anchor="w", pady=(2, 8))

    def _build_pdf_sidebar(self):
        self.clear_sidebar()
        ttk.Label(self.sidebar_content_frame, text="📦 블록 파렛트", font=(self.ui_font_family, 13, "bold")).pack(anchor="w")
        ttk.Label(self.sidebar_content_frame, text="파렛트 항목을 클릭하면 선택된 블록에 적용됩니다.", foreground="#555", wraplength=340).pack(anchor="w", pady=(2, 6))
        
        ttk.Separator(self.sidebar_content_frame).pack(fill="x", pady=4)
        
        self._build_header_panel()
        if IS_DMB:
            self._build_days_panel()
        fav_frame = ttk.LabelFrame(self.sidebar_content_frame, text=" ⭐ 블록 파렛트 (관리 및 적용) ", padding=8)
        fav_frame.pack(fill="x", pady=4)
        
        self.fav_container = ttk.Frame(fav_frame)
        self.fav_container.pack(fill="x", pady=2)
        self.refresh_favorite_ui()
        self._build_style_panel()
        
        ttk.Separator(self.sidebar_content_frame).pack(fill="x", pady=6)
        
        self._build_mod_panel()
        ttk.Separator(self.sidebar_content_frame).pack(fill="x", pady=6)

        self.info_frame = ttk.LabelFrame(self.sidebar_content_frame, text=" 🎨 선택된 블록 편집 패널 ", padding=10)
        self.info_frame.pack(fill="both", expand=True, pady=4)
        
        self.update_sidebar_info()

    # ---------------- 수정된 블록 표시 (음영 / 테두리 두께) ----------------
    def _orig_match(self, b):
        r = tuple(b["rect"])
        init = self.initial_blocks.get(self.current_page, [])
        key = (self.current_page, id(init), len(init))
        idx = getattr(self, "_orig_idx", None)
        if idx is None or idx[0] != key:             # 원본 칸을 왼쪽 x 로 묶어 둔다 (매번 전체를 훑지 않게)
            buckets = {}
            for ob in init:
                buckets.setdefault(int(round(ob["rect"][0])), []).append(ob)
            idx = self._orig_idx = (key, buckets)
        bx = int(round(r[0]))
        for k in (bx - 1, bx, bx + 1):
            for ob in idx[1].get(k, ()):
                if all(abs(a - c) <= 0.5 for a, c in zip(r, ob["rect"])):
                    return ob
        return None

    def _cover_color(self, b):
        """배경색이 없는(투명) 블록 중 원본 글자를 덮어야 하는 블록이면 덮을 색, 아니면 None.
        나눈 칸·합친 칸·직접 그린 칸(원본 칸과 자리가 다름)과 이름을 새로 넣은 칸은 원본 글자가 비쳐 보이므로
        원본 바탕색(분홍 등 지운 색이면 흰색)으로 먼저 덮고 그 위에 새 글자를 쓴다."""
        if len((b.get("color") or "").lstrip("#")) == 6:
            return None
        if b.get("group_id") is None and self._orig_match(b) is not None and not (b.get("title") or "").strip():
            return None
        f = (b.get("orig_fill") or "").lower()
        if not f or f == "#ffffff" or CLEAR_FILL_MODE == "all" or (CLEAR_FILL_MODE == "pink" and _is_pink(f)):
            return "#ffffff"
        return f

    def _cover_rects(self, b):
        """덮어야 하는 블록 안의 원본 글자 자리들 (PDF 좌표). 뱃지 그림·테두리 선은 그대로 둔다."""
        x0, y0, x1, y1 = b["rect"]
        same_place = b.get("group_id") is None and self._orig_match(b) is not None
        out = []
        for w in getattr(self, "page_words", {}).get(self.current_page, []):
            wx0, wy0, wx1, wy1 = w[:4]
            if same_place and len(w) > 4 and re.fullmatch(r"\d{1,2}", _undouble(w[4])):
                continue                            # 자리 그대로인 칸: 원본의 분 숫자는 그대로 맞으므로 남긴다
            if min(x1, wx1) - max(x0, wx0) > 0.5 and min(y1, wy1) - max(y0, wy0) > 0.5:   # 걸친 글자도 이 칸 안쪽 부분은 덮는다
                out.append((max(x0 + 0.6, wx0 - 0.8), max(y0 + 0.6, wy0 - 0.6), min(x1 - 0.6, wx1 + 0.8), min(y1 - 0.6, wy1 + 0.6)))
        return out

    def _is_modified(self, b):
        """원본에서 감지된 블록과 위치/크기가 다르거나(이동·크기변경·분할·병합·수동추가),
        내용(명칭·색·옵션·글자 조절)이 바뀐 블록이면 True."""
        ob = self._orig_match(b)
        if ob is None:
            return True
        for k, d in (("title", ""), ("color", ""), ("text_color", "#000000"),
                     ("font_size_offset", 0), ("text_dx", 0), ("text_dy", 0)):
            if b.get(k, d) != ob.get(k, d):
                return True
        return b.get("options", {}) != ob.get("options", {})

    def _eff_border_width(self, b, modified):
        if modified and self.mod_border_width is not None:
            return self.mod_border_width
        return b.get("border_width") or 0.5

    def _shade_color_for(self, b, modified):
        """이 칸에 넣을 음영 색 (없으면 None). 직접 고른 칸 음영이 먼저."""
        if b.get("shade"):
            return b["shade"]
        if not self.mod_shade_enabled:
            return None
        if modified and self.shade_target in ("mod", "both"):
            return self.mod_shade_color
        if not modified and self.shade_target in ("unmod", "both"):
            return self.unmod_shade_color
        return None

    def _draw_shade(self, c, x0, y0, x1, y1, color=None):
        color = color or self.mod_shade_color
        if Image and ImageTk:
            w = max(1, int(round(x1 - x0)))
            h = max(1, int(round(y1 - y0)))
            cache = self.__dict__.setdefault("_shade_cache", {})
            key = (w, h, color, self.mod_shade_alpha)
            img = cache.get(key)
            if img is None:
                h6 = color.lstrip("#")
                rgb = tuple(int(h6[i:i + 2], 16) for i in (0, 2, 4))
                img = ImageTk.PhotoImage(Image.new("RGBA", (w, h), rgb + (int(255 * self.mod_shade_alpha),)))
                cache[key] = img
            c.create_image(x0, y0, image=img, anchor="nw")
        else:
            c.create_rectangle(x0, y0, x1, y1, fill=color, outline="", stipple="gray25")

    PICK_SHADE_COLORS = ("#fde047", "#86efac", "#93c5fd", "#f9a8d4", "#fdba74", "#cbd5e1")

    def _build_pick_shade_row(self):
        """선택한 칸(여러 개 가능)에 반투명 음영 넣기/빼기."""
        n = len(self.selected)
        f = ttk.LabelFrame(self.info_frame, text=f" 🖍 선택한 칸 음영 ({n}개) ", padding=6)
        f.pack(fill="x", pady=(0, 6))
        row = ttk.Frame(f)
        row.pack(fill="x")
        for col in self.PICK_SHADE_COLORS:
            tk.Button(row, bg=col, activebackground=col, width=2, relief="raised" if col != self.pick_shade_last else "sunken",
                      command=lambda c=col: self.shade_selected(c)).pack(side="left", padx=1)

        def custom():
            c = colorchooser.askcolor(title="음영 색 선택", initialcolor=self.pick_shade_last)
            if c[1]:
                self.shade_selected(c[1])

        ttk.Button(row, text="색…", width=4, command=custom).pack(side="left", padx=(3, 0))
        ttk.Button(row, text="빼기", width=5, command=lambda: self.shade_selected(None)).pack(side="right")
        freq = self._frequent_shades()
        if freq:                                   # 자주 쓴 색 (많이 쓴 순서, 최대 8개). 오른쪽 클릭: 목록에서 빼기
            rowf = ttk.Frame(f)
            rowf.pack(fill="x", pady=(3, 0))
            ttk.Label(rowf, text="자주 쓴 색:", font=(self.ui_font_family, 8)).pack(side="left", padx=(0, 3))
            for col in freq:
                bt = tk.Button(rowf, bg=col, activebackground=col, width=2,
                               relief="sunken" if col == self.pick_shade_last.lower() else "raised",
                               command=lambda c=col: self.shade_selected(c))
                bt.pack(side="left", padx=1)
                bt.bind("<Button-3>", lambda e, c=col: self._forget_shade(c))
                bt.bind("<Button-2>", lambda e, c=col: self._forget_shade(c))
            ttk.Label(rowf, text="(오른쪽 클릭: 빼기)", font=(self.ui_font_family, 7), foreground="#94a3b8").pack(side="right")
        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=(3, 0))
        ex = tk.BooleanVar(value=self.pick_shade_in_export)

        def set_export():
            self.pick_shade_in_export = bool(ex.get())
            self.save_autosave()

        ttk.Checkbutton(row2, text="저장 PDF에도 넣기", variable=ex, command=set_export).pack(side="left")
        ttk.Label(row2, text="H: 넣기/빼기", font=(self.ui_font_family, 8), foreground="#64748b").pack(side="right")

    def shade_selected(self, color):
        """선택한 칸 모두에 음영 색을 넣는다 (None 이면 뺀다). 진하기는 '음영 진하기'를 같이 쓴다."""
        targets = [b for b in self.blocks if b["id"] in self.selected]
        if not targets:
            return
        self.push_history()
        for b in targets:
            if color:
                b["shade"] = color
            else:
                b.pop("shade", None)
        if color:
            self.pick_shade_last = color
            k = color.lower()
            self.shade_color_counts[k] = self.shade_color_counts.get(k, 0) + 1     # 자주 쓴 색 세기
        self.draw_pdf()
        self.save_autosave()

    def _frequent_shades(self, n=8):
        """많이 쓴 음영 색 (기본 6색은 위 줄에 있으니 빼고), 많이 쓴 순서."""
        base = {c.lower() for c in self.PICK_SHADE_COLORS}
        items = [(c, k) for c, k in self.shade_color_counts.items() if c not in base and k > 0]
        items.sort(key=lambda t: -t[1])
        return [c for c, _k in items[:n]]

    def _forget_shade(self, color):
        self.shade_color_counts.pop(color.lower(), None)
        self.save_autosave()
        self.update_sidebar_info()

    def toggle_pick_shade(self):
        targets = [b for b in self.blocks if b["id"] in self.selected]
        if not targets:
            return
        self.shade_selected(None if all(b.get("shade") for b in targets) else self.pick_shade_last)

    def _build_mod_panel(self):
        f = ttk.LabelFrame(self.sidebar_content_frame, text=" 🟨 수정 블록 표시 ", padding=8)
        f.pack(fill="x", pady=4)
        fnt = (self.ui_font_family, 8, "bold")

        self.mod_shade_var = tk.BooleanVar(value=self.mod_shade_enabled)
        ttk.Checkbutton(f, text="블록에 반투명 음영 표시", variable=self.mod_shade_var,
                        command=self.on_mod_style_changed).pack(anchor="w")
        row_t = ttk.Frame(f)
        row_t.pack(fill="x", pady=(1, 2))
        self.shade_target_var = tk.StringVar(value=self.shade_target)
        for val, txt in (("mod", "수정한 칸"), ("unmod", "수정 안 한 칸"), ("both", "둘 다")):
            ttk.Radiobutton(row_t, text=txt, value=val, variable=self.shade_target_var,
                            command=self.on_mod_style_changed).pack(side="left", padx=(0, 6))
        self.mod_export_var = tk.BooleanVar(value=self.mod_shade_in_export)
        ttk.Checkbutton(f, text="저장 PDF에도 음영 포함 (끄면 미리보기에만)", variable=self.mod_export_var,
                        command=self.on_mod_style_changed).pack(anchor="w")

        row_c = ttk.Frame(f)
        row_c.pack(fill="x", pady=2)
        ttk.Label(row_c, text="음영 색:", font=fnt).pack(side="left")
        self.unmod_color_btn = tk.Button(row_c, text=" 안 한 칸 ", bg=self.unmod_shade_color,
                                         font=(self.ui_font_family, 8), command=self.pick_unmod_shade_color)
        self.unmod_color_btn.pack(side="right")
        self.mod_color_btn = tk.Button(row_c, text=" 수정한 칸 ", bg=self.mod_shade_color,
                                       font=(self.ui_font_family, 8), command=self.pick_mod_shade_color)
        self.mod_color_btn.pack(side="right", padx=2)

        row_a = ttk.Frame(f)
        row_a.pack(fill="x", pady=2)
        ttk.Label(row_a, text="음영 진하기:", font=fnt).pack(side="left")
        ttk.Button(row_a, text="+", width=2, command=lambda: self.adjust_mod_alpha(0.05)).pack(side="right", padx=1)
        self.mod_alpha_lbl = ttk.Label(row_a, text="", width=5, anchor="center", font=fnt, background="#f1f5f9")
        self.mod_alpha_lbl.pack(side="right", padx=1)
        ttk.Button(row_a, text="-", width=2, command=lambda: self.adjust_mod_alpha(-0.05)).pack(side="right", padx=1)

        row_b = ttk.Frame(f)
        row_b.pack(fill="x", pady=2)
        ttk.Label(row_b, text="테두리 두께(수정 블록 전체):", font=fnt).pack(side="left")
        ttk.Button(row_b, text="원본", width=4, command=self.reset_mod_border).pack(side="right", padx=1)
        ttk.Button(row_b, text="+", width=2, command=lambda: self.adjust_mod_border(0.25)).pack(side="right", padx=1)
        self.mod_border_lbl = ttk.Label(row_b, text="", width=6, anchor="center", font=fnt, background="#f1f5f9")
        self.mod_border_lbl.pack(side="right", padx=1)
        ttk.Button(row_b, text="-", width=2, command=lambda: self.adjust_mod_border(-0.25)).pack(side="right", padx=1)

        self.mod_count_lbl = ttk.Label(f, text="수정된 블록: 0개", font=(self.ui_font_family, 8), foreground="#2563eb")
        self.mod_count_lbl.pack(anchor="w", pady=(2, 0))
        ch = self.changed_ids_by_page.get(self.current_page) or set()
        if ch:
            row_ch = ttk.Frame(f)
            row_ch.pack(fill="x", pady=(2, 0))
            show = tk.BooleanVar(value=self.show_changed)

            def toggle_changed():
                self.show_changed = bool(show.get())
                self.draw_pdf(refresh_sidebar=False)

            ttk.Checkbutton(row_ch, text=f"이전 판과 달라진 칸 {len(ch)}개 표시 (빨간 점선)", variable=show,
                            command=toggle_changed).pack(side="left")
        self._refresh_mod_labels()

    def _refresh_mod_labels(self):
        try:
            self.mod_alpha_lbl.config(text=f"{int(round(self.mod_shade_alpha * 100))}%")
            self.mod_border_lbl.config(text="원본" if self.mod_border_width is None else f"{self.mod_border_width:.2f}pt")
            self.mod_color_btn.config(bg=self.mod_shade_color)
            self.unmod_color_btn.config(bg=self.unmod_shade_color)
        except (AttributeError, tk.TclError):
            pass

    def _update_mod_count(self):
        try:
            if self.mod_count_lbl.winfo_exists():
                self.mod_count_lbl.config(text=f"수정된 블록: {len(self._mod_ids)}개")
        except (AttributeError, tk.TclError):
            pass

    def _mod_style_apply(self):
        self._refresh_mod_labels()
        self.save_autosave()
        self.draw_pdf(refresh_sidebar=False)

    def on_mod_style_changed(self):
        self.mod_shade_enabled = bool(self.mod_shade_var.get())
        self.mod_shade_in_export = bool(self.mod_export_var.get())
        try:
            self.shade_target = self.shade_target_var.get()
        except AttributeError:
            pass
        self._mod_style_apply()

    def pick_mod_shade_color(self):
        c = colorchooser.askcolor(title="수정 블록 음영 색 선택", initialcolor=self.mod_shade_color)
        if c[1]:
            self.mod_shade_color = c[1]
            self._mod_style_apply()

    def pick_unmod_shade_color(self):
        c = colorchooser.askcolor(title="수정 안 한 칸 음영 색 선택", initialcolor=self.unmod_shade_color)
        if c[1]:
            self.unmod_shade_color = c[1]
            self._mod_style_apply()

    def adjust_mod_alpha(self, delta):
        self.mod_shade_alpha = max(0.05, min(0.8, round(self.mod_shade_alpha + delta, 2)))
        self._mod_style_apply()

    def adjust_mod_border(self, delta):
        cur = self.mod_border_width if self.mod_border_width is not None else 0.5
        self.mod_border_width = max(0.25, min(6.0, round(cur + delta, 2)))
        self._mod_style_apply()

    def reset_mod_border(self):
        self.mod_border_width = None
        self._mod_style_apply()

    def refresh_favorite_ui(self):
        for widget in self.fav_container.winfo_children():
            widget.destroy()
            
        for idx, preset in enumerate(self.favorite_presets):
            f = tk.Frame(self.fav_container, bg="#f8fafc", bd=1, relief="raised")
            f.pack(fill="x", pady=2, ipady=2)
            
            lbl = tk.Label(f, text=f"⭐ {preset['title']}", bg="#f8fafc", fg=preset.get("text_color", "#000000"), font=(self.ui_font_family, 9, "bold"))
            lbl.pack(side="left", padx=6)
            
            col_btn = tk.Button(f, text="배경색", bg="#e2e8f0", font=(self.ui_font_family, 7), bd=1, command=lambda i=idx: self.pick_preset_color(i))
            col_btn.pack(side="right", padx=2)
            
            del_btn = tk.Button(f, text="X", bg="#fee2e2", fg="#dc2626", font=(self.ui_font_family, 7, "bold"), bd=0, command=lambda i=idx: self.remove_favorite_preset(i))
            del_btn.pack(side="right", padx=2)
            
            for widget in (f, lbl):
                widget.bind("<Button-1>", lambda e, p=preset: self.apply_favorite_preset_to_selected(p))

    def pick_preset_color(self, index):
        preset = self.favorite_presets[index]
        initial = preset["color"] if preset["color"] else "#ffffff"
        c = colorchooser.askcolor(title="파렛트 배경색 선택", initialcolor=initial)
        if c[1]:
            preset["color"] = c[1]
            self.refresh_favorite_ui()
            self.save_autosave()

    def remove_favorite_preset(self, index):
        if len(self.favorite_presets) <= 1:
            messagebox.showwarning("삭제 불가", "최소 1개 이상의 파렛트 항목이 유지되어야 합니다.")
            return
        removed = self.favorite_presets.pop(index)
        self.refresh_favorite_ui()
        self.save_autosave()
        messagebox.showinfo("파렛트 삭제", f"'{removed['title']}' 스타일이 파렛트에서 제거되었습니다.")

    def apply_favorite_preset_to_selected(self, preset):
        if not self.selected:
            messagebox.showinfo("파렛트", "스타일을 적용할 블록을 먼저 선택해주세요.")
            return
        
        self.push_history()
        for b_id in self.selected:
            block = next((b for b in self.blocks if b["id"] == b_id), None)
            if block:
                block["title"] = preset["title"]
                block["color"] = preset["color"]
                block["text_color"] = preset.get("text_color", "#000000")
                block["font_size_offset"] = preset.get("font_size_offset", 0)
                block["text_dx"] = preset.get("text_dx", 0)
                block["text_dy"] = preset.get("text_dy", 0)
                if "options" in preset:
                    block["options"] = copy.deepcopy(preset["options"])
                self._fill_auto_times(block)       # 시간 값은 파렛트 것이 아니라 이 블록 위치의 분으로 (보이기는 체크 상태를 따른다)
                
        self.draw_pdf()
        self.save_autosave()
        messagebox.showinfo("파렛트 적용", f"선택한 블록에 '{preset['title']}' 양식이 적용되었습니다!")

    def save_to_favorites(self):
        if len(self.selected) != 1:
            messagebox.showinfo("파렛트", "파렛트에 저장할 블록을 1개만 선택해주세요.")
            return
        sel_id = list(self.selected)[0]
        block = next((b for b in self.blocks if b["id"] == sel_id), None)
        if not block: return
        
        title = block["title"] if block["title"] else "새블록"
        color = block["color"] if block["color"] else ""
        txt_color = block.get("text_color", "#000000")
        font_offset = block.get("font_size_offset", 0)
        options = copy.deepcopy(block.get("options", {}))
        
        r = fitz.Rect(block["rect"])
        h = max(int(r.height), 35)
        
        new_fav = {
            "title": title, 
            "color": color, 
            "text_color": txt_color, 
            "h": h,
            "font_size_offset": font_offset,
            "text_dx": block.get("text_dx", 0),
            "text_dy": block.get("text_dy", 0),
            "options": options
        }
        
        existing = next((p for p in self.favorite_presets if p["title"] == title), None)
        if existing:
            existing.update(new_fav)
            messagebox.showinfo("파렛트 갱신", f"'{title}' 스타일이 파렛트에 업데이트되었습니다!")
        else:
            self.favorite_presets.append(new_fav)
            messagebox.showinfo("파렛트 등록", f"'{title}' 스타일이 블록 파렛트에 추가되었습니다!")
            
        self.refresh_favorite_ui()
        self.save_autosave()

    def draw_blank_canvas(self):
        c = self.canvas
        c.delete("all")
        c.configure(scrollregion=(0, 0, 1000, 800))
        c.create_text(500, 350, text="PDF 파일이 열리지 않았습니다.", font=(self.ui_font_family, 16, "bold"), fill="#888888")
        c.create_text(500, 390, text="상단의 [PDF 열기]를 눌러 시작하세요.", font=(self.ui_font_family, 11), fill="#aaaaaa")

    @staticmethod
    def _normalize_blocks(blocks):
        """JSON으로 읽으면 튜플이 리스트가 되므로 좌표/색 항목을 다시 튜플로 되돌린다."""
        out = []
        for b in blocks:
            b = dict(b)
            for k in ("rect", "orig_rect", "rel_box", "border_color"):
                if isinstance(b.get(k), list):
                    b[k] = tuple(b[k])
            out.append(b)
        return out

    def _offer_recovery(self):
        data, self._recovery = self._recovery, None
        if not data or not fitz:
            return
        path = data.get("pdf_path")
        if not path or not os.path.exists(path):
            return
        if not messagebox.askyesno("작업 복구", f"저장하지 못하고 종료된 이전 작업이 있습니다.\n\n{os.path.basename(path)}\n\n이어서 작업하시겠습니까?"):
            return
        try:
            pb = {int(k): self._normalize_blocks(v) for k, v in data["page_blocks"].items()}
            ib = {int(k): self._normalize_blocks(v) for k, v in (data.get("initial_blocks") or {}).items()}
            page_idx = int(data.get("current_page", 0))
            if page_idx not in pb:
                page_idx = min(pb)
        except Exception as exc:
            messagebox.showerror("복구 실패", f"이전 작업을 불러오지 못했습니다:\n{exc}")
            return

        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("불러오는 중")
        self.loading_win.geometry("340x140")
        self.loading_win.resizable(False, False)
        self.loading_win.transient(self.root)
        self.loading_win.grab_set()
        self.load_status_lbl = ttk.Label(self.loading_win, text="이전 작업 복구 중...", font=(self.ui_font_family, 10))
        self.load_status_lbl.pack(pady=(20, 10))
        self.progress = ttk.Progressbar(self.loading_win, mode="indeterminate", length=280)
        self.progress.pack(pady=4)
        self.progress.start(12)

        self.pdf_path = path
        self.current_page = page_idx
        self.page_blocks = pb
        self.initial_blocks = ib
        try:
            self.header_edit = {int(k): dict(v) for k, v in (data.get("header_edit") or {}).items() if isinstance(v, dict)}
        except Exception:
            self.header_edit = {}
        ids = [b["id"] for blist in pb.values() for b in blist if isinstance(b.get("id"), int)]
        if ids:
            self.next_id = max(self.next_id, max(ids) + 1)
        self.base_pdf_bytes = None
        self.orig_lines = {}
        self.history.clear()
        threading.Thread(target=self._load_pdf_worker, args=(path, page_idx), daemon=True).start()

    def open_pdf(self):
        if not fitz:
            messagebox.showerror("PDF 기능 필요", "PyMuPDF가 설치되어 있지 않습니다.\npip install pymupdf 명령어로 설치해주세요.")
            return
        path = filedialog.askopenfilename(title="편성표 PDF 열기", filetypes=[("PDF 파일", "*.pdf")])
        if not path:
            return
        
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("불러오는 중")
        self.loading_win.geometry("340x140")
        self.loading_win.resizable(False, False)
        self.loading_win.transient(self.root)
        self.loading_win.grab_set()
        
        self.load_status_lbl = ttk.Label(self.loading_win, text="PDF 문서 로딩 및 블록 분석 중...", font=(self.ui_font_family, 10))
        self.load_status_lbl.pack(pady=(20, 10))
        self.progress = ttk.Progressbar(self.loading_win, mode="indeterminate", length=280)
        self.progress.pack(pady=4)
        self.progress.start(12)
        
        self.pdf_path = path
        self.current_page = 0
        self.page_blocks = {}
        self.initial_blocks = {}
        self.header_edit = {}
        self.changed_ids_by_page = {}
        self.split_mode_var.set(False)
        self.base_pdf_bytes = None
        self.orig_lines = {}
        self.history.clear()
        threading.Thread(target=self._load_pdf_worker, args=(path, self.current_page), daemon=True).start()

    def _load_pdf_worker(self, path, page_idx):
        with PDF_LOCK:
            self._load_pdf_worker_locked(path, page_idx)

    def _load_pdf_worker_locked(self, path, page_idx):
        try:
            start_time = time.time()
            doc = fitz.open(path)
            if not doc.page_count:
                raise ValueError("PDF에 페이지가 없습니다.")
            self.page_count = doc.page_count
            page = doc.load_page(page_idx)
            self.page_rect = tuple(page.rect)
            
            pw = self.page_rect[2] - self.page_rect[0]
            self.zoom = round(min(max(1100.0 / pw, 0.8), 2.5), 2) if pw > 0 else 1.25

            pixmap = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=True)
            png_data = pixmap.tobytes("png")

            all_drawings = page.get_drawings()
            try:   # 5분 단위 나누기용 시간 줄, 제목·날짜 자리
                _segs, _fills = _drawing_parts(all_drawings)
                _words = page.get_text("words")
                self.__dict__.setdefault("page_words", {})[page_idx] = [tuple(w[:4]) + (str(w[4]).strip(),) for w in _words if str(w[4]).strip()]
                rows = detect_time_rows(_words, _segs)
                self.time_rows_by_page[page_idx] = rows
                self.time_ticks_by_page[page_idx] = time_ticks(rows)
                _heads = [r.y0 for r, c in _fills if _is_header_color(c) and r.height >= 8]
                _td = page.get_text("dict")
                _hf = detect_header(_td, page.get_image_info(), (min(_heads) - 1) if _heads else None, tuple(page.rect))
                try:
                    _hf["days"] = detect_day_headers(_td, _fills)
                except Exception:
                    _hf["days"] = []
                self.header_found_by_page[page_idx] = _hf
            except Exception:
                self.time_rows_by_page[page_idx] = []
                self.time_ticks_by_page[page_idx] = []
                self.header_found_by_page[page_idx] = {}

            if page_idx not in self.page_blocks:
                drawings = all_drawings
                found, method = detect_schedule_blocks(drawings, tuple(page.rect),
                                                       min_w=self.detect_min_w, min_h=self.detect_min_h)
                detected_blocks = []
                for fb in found:
                    r = fb["rect"]
                    detected_blocks.append({
                        "id": self.next_id,
                        "title": "",
                        "color": "",
                        "text_color": "#000000",
                        "font_size_offset": 0,
                        "options": {
                            "top_time_enabled": False, "top_time_val": "00", "top_time_font_size": DEFAULT_TIME_FONT_SIZE,
                            "bottom_time_enabled": False, "bottom_time_val": "00", "bottom_time_font_size": DEFAULT_TIME_FONT_SIZE,
                            "badge_size": 12,
                            "uhd": False, "sub": False, "sign": False, "des": False, "local": False, "age": "none"
                        },
                        "rect": tuple(r),
                        "orig_rect": tuple(r),
                        "border_color": fb["border_color"],
                        "border_width": fb["border_width"],
                        "orig_fill": fb["orig_fill"],
                        "group_id": None,
                        "rel_box": None,
                    })
                    self.next_id += 1
                try:   # 판 비교용: 칸 안의 원본 글자
                    for _b, _t in zip(detected_blocks, block_source_texts(page.get_text("words"), [d["rect"] for d in detected_blocks])):
                        _b["src_text"] = _t
                except Exception:
                    pass
                self._detect_note = f"블록 {len(detected_blocks)}개를 찾았습니다 ({method} 기준"\
                                    f", 높이 10pt 미만 작은 칸 {sum(1 for x in found if x['rect'][3] - x['rect'][1] < 10)}개 포함)."
                self.__dict__.setdefault("orig_lines", {})[page_idx] = _collect_line_segments(drawings)
                self.page_blocks[page_idx] = detected_blocks
                self.initial_blocks[page_idx] = copy.deepcopy(detected_blocks)

            _init = self.initial_blocks.get(page_idx) or []
            if _init and any("src_text" not in b for b in _init):     # 예전 자동 저장에서 이어받은 원본 칸: 판 비교용 글자 채우기
                try:
                    for _b, _t in zip(_init, block_source_texts(page.get_text("words"), [tuple(b["rect"]) for b in _init])):
                        _b["src_text"] = _t
                except Exception:
                    pass

            if page_idx not in self.__dict__.setdefault("orig_lines", {}):
                self.orig_lines[page_idx] = _collect_line_segments(page.get_drawings())

            doc.close()

            self._clear_fill_warning = None
            self.base_pdf_bytes = None
            self._fill_report = _fill_report(self.page_blocks.get(page_idx, []), CLEAR_FILL_MODE)
            if CLEAR_FILL_MODE != "off" and page_idx in self.page_blocks:
                try:
                    base_bytes, warn = _clear_orig_fill_pdf(path, page_idx, self.page_blocks[page_idx])
                except Exception as exc:
                    base_bytes, warn = None, f"원본 채우기 제거에 실패해 원본 그대로 불러왔습니다:\n{exc}"
                if base_bytes:
                    cdoc = fitz.open(stream=base_bytes, filetype="pdf")
                    png_data = cdoc.load_page(page_idx).get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=True).tobytes("png")
                    cdoc.close()
                    self.base_pdf_bytes = base_bytes
                self._clear_fill_warning = warn

            elapsed_time = round(time.time() - start_time, 2)
            self.root.after(0, self._load_pdf_finished, (png_data, self.page_rect, self.page_count, self.page_blocks[page_idx], elapsed_time), None)
        except Exception as exc:
            self.root.after(0, self._load_pdf_finished, None, exc)

    def _load_pdf_finished(self, payload, error):
        if hasattr(self, "loading_win") and self.loading_win.winfo_exists():
            self.progress.stop()
            self.loading_win.destroy()
            
        if error:
            messagebox.showerror("PDF 로드 실패", str(error))
            return
            
        png_data, page_rect, page_count, current_page_blocks, elapsed_time = payload
        self.png_data = png_data
        self.page_rect = page_rect
        self.page_count = page_count
        self.blocks = current_page_blocks
        
        self.selected = set()
        self.history.clear()
        self.save_autosave()
        
        self.zoom_label.set(f"{round(self.zoom * 100)}%")
        self._build_pdf_toolbar()
        self._build_pdf_sidebar()
        self.bind_pdf_events()
        self.draw_pdf()
        
        load_msg = f"현재 페이지의 블록을 불러왔습니다. (소요 시간: {elapsed_time}초)"
        note = getattr(self, "_detect_note", "")
        if note:
            load_msg += "\n" + note
            self._detect_note = ""
        fill_rep = getattr(self, "_fill_report", "")
        if fill_rep:
            load_msg += "\n\n" + fill_rep
        if not any(self._is_modified(b) or b.get("shade") for b in self.blocks):
            self.root.after(300, self._mark_saved)     # 막 연 그대로면 '저장할 것 없음' (머리글 기본값이 채워진 뒤)
        else:
            self._saved_sig = None                     # 이어서 하던 작업(복구)은 저장 전
        messagebox.showinfo("페이지 로드 완료", load_msg)
        warn = getattr(self, "_clear_fill_warning", None)
        if warn:
            messagebox.showwarning("채우기 제거 안내", warn)
        try:
            self.offer_carry_over()
        except Exception:
            pass
        if self.split_styles_auto and self.split_styles:
            try:
                self.apply_split_styles(quiet=True)
            except Exception:
                pass
        self.refresh_style_ui()

    def bind_pdf_events(self):
        self.canvas.bind("<Button-1>", self.on_press_pdf)
        self.canvas.bind("<Double-Button-1>", self.on_double_click_pdf)
        self.canvas.bind("<B1-Motion>", self.on_motion_pdf)
        self.canvas.bind("<ButtonRelease-1>", self.on_release_pdf)
        self.canvas.bind("<Motion>", self.on_hover_pdf)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)
        self.canvas.bind("<Button-3>", self.on_right_click_pdf)        # 오른쪽 마우스: 칸 메뉴
        self.canvas.bind("<Button-2>", self.on_right_click_pdf)        # (맥 트랙패드)

    def on_right_click_pdf(self, event):
        """칸 위에서 오른쪽 클릭: 작업 스타일 저장·적용, 파렛트, 여기서 나누기."""
        if not self.pdf_path:
            return
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        block = self._block_at(cx, cy)
        if not block:
            return
        if block["id"] not in self.selected:          # 오른쪽 클릭한 칸을 선택한다
            self.selected = {block["id"]}
            self.draw_pdf()
        fnt = (self.ui_font_family, 9)
        m = tk.Menu(self.root, tearoff=0, font=fnt)
        m.add_command(label="🧩 이 칸을 작업 스타일로 저장…", command=self.save_split_style)
        sm = tk.Menu(m, tearoff=0, font=fnt)
        fresh = block.get("group_id") is None and self._orig_match(block) is not None
        for st in self.split_styles:
            hit = self._style_fits(block, st)[0]
            sm.add_command(label=("✔ " if hit else "    ") + st.get("name", ""),
                           state="normal" if fresh else "disabled",
                           command=lambda s_=st: self._apply_style_here(block, s_))
        if not self.split_styles:
            sm.add_command(label="(저장된 스타일 없음)", state="disabled")
        m.add_cascade(label="🧩 이 칸에 스타일 적용" + ("" if fresh else "  (이미 나눈 칸)"), menu=sm)
        m.add_command(label="▶ 이 페이지에 스타일 모두 적용", command=lambda: self.apply_split_styles(),
                      state="normal" if self.split_styles else "disabled")
        m.add_separator()
        tgt = self._split_target(block, *self._to_doc(cx, cy), False)
        m.add_command(label="✂ 여기서 나누기" + (f"  ({tgt[1]})" if tgt and tgt[1] else ""),
                      state="normal" if tgt else "disabled",
                      command=lambda: tgt and self._do_split(block, tgt[0], False))
        if block.get("group_id") is not None:
            m.add_command(label=("━ 나눈 경계선을 실선으로" if block.get("split_dash") else "┄ 나눈 경계선을 점선으로"),
                          command=lambda: self.set_split_dash(block))
        m.add_separator()
        pm = tk.Menu(m, tearoff=0, font=fnt)
        for pr in self.favorite_presets:
            pm.add_command(label=pr.get("title", ""), command=lambda p_=pr: self.apply_favorite_preset_to_selected(p_))
        m.add_cascade(label="⭐ 파렛트 적용", menu=pm)
        m.add_command(label="⭐ 파렛트로 저장", command=self.save_to_favorites)
        m.add_separator()
        m.add_command(label="↩ 되돌리기 (Ctrl+Z)", command=self.undo_action)
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _apply_style_here(self, block, style):
        """오른쪽 클릭한 그 칸 하나에만 스타일 적용 (찾을 글자가 달라도 적용)."""
        if block not in self.blocks:
            return
        self.push_history()
        new = self._apply_style_to(block, style)
        if not new:
            self.history.pop() if self.history else None
            messagebox.showinfo("작업 스타일", "이 칸은 짧아서 같은 자리에서 나눌 수 없습니다.")
            return
        self.page_blocks[self.current_page] = self.blocks
        self.selected = {new[-1]["id"]}
        self.draw_pdf()
        self.save_autosave()
        self.refresh_style_ui()

    def _page_origin(self):
        """PDF 페이지 원점(CropBox 여백이 있으면 0이 아님). 화면 좌표 <-> PDF 좌표 변환에 쓴다."""
        return (self.page_rect[0], self.page_rect[1]) if self.page_rect else (0, 0)

    def _xy(self, rect):
        r = fitz.Rect(rect)
        x_off = self.page_rect[0] if self.page_rect else 0
        y_off = self.page_rect[1] if self.page_rect else 0
        return (18 + (r.x0 - x_off) * self.zoom, 
                18 + (r.y0 - y_off) * self.zoom, 
                18 + (r.x1 - x_off) * self.zoom, 
                18 + (r.y1 - y_off) * self.zoom)

    def _format_smart_text(self, text, block_width, block_height, size_offset=0):
        if not text:
            return "", 8
        
        if "\n" in text:
            lines = text.split("\n")
            max_line_len = max(len(l) for l in lines) if lines else 1
            if block_height < 35 or block_width < 50:
                font_size = 9
            elif block_height < 50 or block_width < 60:
                font_size = 11
            elif max_line_len <= 4:
                font_size = 13
            elif max_line_len <= 10:
                font_size = 11
            else:
                font_size = 9
            font_size = max(4, int(font_size + size_offset))
            return text, font_size

        char_count = len(text)
        if block_height < 35 or block_width < 50:
            font_size = 9
        elif block_height < 50 or block_width < 60:
            font_size = 11
        elif char_count <= 4:
            font_size = 14
        elif char_count <= 10:
            font_size = 12
        else:
            font_size = 10
            
        font_size = max(4, int(font_size + size_offset))
        char_w = 0.55 if text.isascii() else 1.0
        max_chars_per_line = max(int(block_width / (font_size * char_w)), 3)

        if char_count <= max_chars_per_line:
            return text, font_size

        fs = font_size
        while True:
            per_line = max(int(block_width / (fs * char_w)), 3)
            lines = [text[i:i + per_line] for i in range(0, char_count, per_line)]
            if fs <= 4 or len(lines) * fs * 1.15 <= block_height:
                break
            fs -= 1
        return "\n".join(lines), fs

    def _selected_block(self):
        if len(self.selected) != 1:
            return None
        sel_id = next(iter(self.selected))
        return next((b for b in self.blocks if b["id"] == sel_id), None)

    def _schedule_split_preview(self):
        job = getattr(self, "_split_preview_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self._split_preview_job = self.root.after(80, self._refresh_split_preview)

    def _refresh_split_preview(self):
        self._split_preview_job = None
        block = self._selected_block()
        if not block:
            return
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def _paint_shape_preview(self, block):
        if not hasattr(self, "shape_canvas"):
            return
        try:
            if not self.shape_canvas.winfo_exists():
                return
        except tk.TclError:
            return

        self.shape_canvas.delete("all")
        r = fitz.Rect(block["rect"])
        preview_canvas_w, preview_canvas_h = 270, 110
        bw, bh = r.width, r.height
        if bw <= 0 or bh <= 0:
            return

        max_w = preview_canvas_w - 30
        max_h = preview_canvas_h - 20
        scale = min(max_w / bw, max_h / bh)
        dw, dh = bw * scale, bh * scale
        dx0 = (preview_canvas_w - dw) / 2
        dy0 = (preview_canvas_h - dh) / 2
        fill_c = block["color"] if block["color"] else ""
        text_c = block.get("text_color", "#000000")
        _tdx = block.get("text_dx", 0) * scale
        _tdy = block.get("text_dy", 0) * scale
        total_splits = max(self.h_split_count, self.v_split_count)

        if total_splits > 1 and self.h_split_count > 1:
            curr_dy = dy0
            for i in range(self.h_split_count):
                ratio = self.split_ratios[i] if i < len(self.split_ratios) else (1.0 / self.h_split_count)
                sub_h = dh * ratio
                sy0, sy1 = curr_dy, curr_dy + sub_h
                self.shape_canvas.create_rectangle(dx0, sy0, dx0 + dw, sy1, fill=fill_c if fill_c else "", outline="#2563eb", width=1)
                if i > 0:
                    self.shape_canvas.create_line(dx0, sy0, dx0 + dw, sy0, fill="#ef4444", dash=(2, 2))
                sub_t = self.split_titles.get(i, block["title"] if i == 0 else "")
                if sub_t:
                    formatted_sub, fs = self._format_smart_text(sub_t, dw / scale, sub_h / scale, size_offset=block.get("font_size_offset", 0))
                    preview_sub_fs = max(4, int(fs * scale * 0.75))
                    self.shape_canvas.create_text(dx0 + dw / 2 + _tdx, (sy0 + sy1) / 2 + _tdy, text=formatted_sub, font=(self.ui_font_family, preview_sub_fs, "bold"), fill=text_c, anchor="center", justify="center")
                curr_dy += sub_h
        elif total_splits > 1 and self.v_split_count > 1:
            curr_dx = dx0
            for i in range(self.v_split_count):
                ratio = self.split_ratios[i] if i < len(self.split_ratios) else (1.0 / self.v_split_count)
                sub_w = dw * ratio
                sx0, sx1 = curr_dx, curr_dx + sub_w
                self.shape_canvas.create_rectangle(sx0, dy0, sx1, dy0 + dh, fill=fill_c if fill_c else "", outline="#2563eb", width=1)
                if i > 0:
                    self.shape_canvas.create_line(sx0, dy0, sx0, dy0 + dh, fill="#10b981", dash=(2, 2))
                sub_t = self.split_titles.get(i, block["title"] if i == 0 else "")
                if sub_t:
                    formatted_sub, fs = self._format_smart_text(sub_t, sub_w / scale, dh / scale, size_offset=block.get("font_size_offset", 0))
                    preview_sub_fs = max(4, int(fs * scale * 0.75))
                    self.shape_canvas.create_text((sx0 + sx1) / 2 + _tdx, dy0 + dh / 2 + _tdy, text=formatted_sub, font=(self.ui_font_family, preview_sub_fs, "bold"), fill=text_c, anchor="center", justify="center")
                curr_dx += sub_w
        else:
            if fill_c:
                self.shape_canvas.create_rectangle(dx0, dy0, dx0 + dw, dy0 + dh, fill=fill_c, outline="#2563eb", width=2)
            else:
                self.shape_canvas.create_rectangle(dx0, dy0, dx0 + dw, dy0 + dh, fill="", outline="#2563eb", width=2)
            if block["title"]:
                formatted_p_text, p_size = self._format_smart_text(block["title"], bw, bh, size_offset=block.get("font_size_offset", 0))
                preview_font_size = max(4, int(p_size * scale * 0.75))
                self.shape_canvas.create_text(
                    dx0 + dw / 2 + _tdx, dy0 + dh / 2 + _tdy,
                    text=formatted_p_text,
                    font=(self.ui_font_family, preview_font_size, "bold"),
                    fill=text_c,
                    anchor="center",
                    justify="center",
                )

    def _draw_live_split_texts(self, c, b, x0, y0, x1, y1):
        last_id = None
        fill = b.get("text_color", "#000000")
        offset = b.get("font_size_offset", 0)
        dxp = b.get("text_dx", 0) * self.zoom
        dyp = b.get("text_dy", 0) * self.zoom
        font_name = self.ui_font_family
        if self.h_split_count > 1:
            curr_y = y0
            total_h = y1 - y0
            for i in range(self.h_split_count):
                ratio = self.split_ratios[i] if i < len(self.split_ratios) else (1.0 / self.h_split_count)
                sub_h = total_h * ratio
                sub_t = self.split_titles.get(i, b["title"] if i == 0 else "")
                if sub_t:
                    formatted, f_size = self._format_smart_text(sub_t, (x1 - x0) / self.zoom, sub_h / self.zoom, size_offset=offset)
                    last_id = c.create_text(
                        x0 + (x1 - x0) / 2 + dxp, curr_y + sub_h / 2 + dyp,
                        text=formatted, anchor="center",
                        font=(font_name, -max(4, int(round(f_size * self.zoom))), "bold"),
                        justify="center", fill=fill,
                    )
                curr_y += sub_h
        elif self.v_split_count > 1:
            curr_x = x0
            total_w = x1 - x0
            for i in range(self.v_split_count):
                ratio = self.split_ratios[i] if i < len(self.split_ratios) else (1.0 / self.v_split_count)
                sub_w = total_w * ratio
                sub_t = self.split_titles.get(i, b["title"] if i == 0 else "")
                if sub_t:
                    formatted, f_size = self._format_smart_text(sub_t, sub_w / self.zoom, (y1 - y0) / self.zoom, size_offset=offset)
                    last_id = c.create_text(
                        curr_x + sub_w / 2 + dxp, y0 + (y1 - y0) / 2 + dyp,
                        text=formatted, anchor="center",
                        font=(font_name, -max(4, int(round(f_size * self.zoom))), "bold"),
                        justify="center", fill=fill,
                    )
                curr_x += sub_w
        return last_id

    def _render_single_block(self, b, is_selected):
        c = self.canvas
        x0, y0, x1, y1 = self._xy(b["rect"])
        
        fill_color = b["color"] if b["color"] else ""
        outline_color = "#2563eb" if is_selected else "#3b82f6"
        width = 3 if is_selected else 1
        
        cover = self._cover_color(b) if not fill_color else None
        if cover:                                   # 원본 글자가 비치지 않게 그 자리만 바탕색으로 덮는다 (뱃지 그림·선은 그대로)
            for r_ in self._cover_rects(b):
                a_ = self._to_canvas(r_[0], r_[1])
                b_ = self._to_canvas(r_[2], r_[3])
                c.create_rectangle(a_[0], a_[1], b_[0], b_[1], fill=cover, outline="")
        if is_selected:
            hit_id = c.create_rectangle(x0, y0, x1, y1, fill=fill_color if fill_color else "#dbeafe", outline="", stipple="gray25" if not fill_color else "")
        else:
            hit_id = c.create_rectangle(x0, y0, x1, y1, fill=fill_color, outline="")
            
        _sc = self._shade_color_for(b, b["id"] in self._mod_ids)
        if _sc:
            self._draw_shade(c, x0, y0, x1, y1, _sc)

        rect_id = c.create_rectangle(x0, y0, x1, y1, fill="", outline=outline_color, width=width)
        _rb = b.get("rel_box")
        if b.get("split_dash") and b.get("group_id") is not None and _rb:   # 점선으로 바꾼 나눈 경계선 (출력과 같은 모양)
            if _rb[1] > 0.001:
                c.create_line(x0, y0, x1, y0, fill="#111111", dash=(4, 3), width=max(1, round(self.zoom * 0.6)))
            if _rb[0] > 0.001:
                c.create_line(x0, y0, x0, y1, fill="#111111", dash=(4, 3), width=max(1, round(self.zoom * 0.6)))
        
        opts = b.get("options", {})
        top_enabled = opts.get("top_time_enabled", opts.get("time_enabled", False))
        top_val = opts.get("top_time_val", opts.get("time_val", "00"))
        bottom_enabled = opts.get("bottom_time_enabled", False)
        bottom_val = opts.get("bottom_time_val", "00")
        
        top_time_fs = opts.get("top_time_font_size", opts.get("time_font_size", DEFAULT_TIME_FONT_SIZE))
        bottom_time_fs = opts.get("bottom_time_font_size", opts.get("time_font_size", DEFAULT_TIME_FONT_SIZE))
        badge_sz = opts.get("badge_size", 12)
        
        top_fs_render = max(4, int(top_time_fs * self.zoom))
        bottom_fs_render = max(4, int(bottom_time_fs * self.zoom))
        badge_sz_render = max(2, int(badge_sz * self.zoom))

        if top_enabled and top_val:
            c.create_text(x0 + 6, y0 + (top_fs_render / 2 + 3), text=top_val, anchor="w", font=(self.ui_font_family, -top_fs_render, "bold"), fill="#334155")
        
        if bottom_enabled and bottom_val:
            c.create_text(x0 + 6, y1 - (bottom_fs_render / 2 + 3), text=bottom_val, anchor="w", font=(self.ui_font_family, -bottom_fs_render, "bold"), fill="#334155")

        right_x = x1 - 4
        badge_y = y0 + 3
        active_badge_keys = []
        if opts.get("uhd", False): active_badge_keys.append("uhd")
        if opts.get("local", False): active_badge_keys.append("local")
        if opts.get("sign", False): active_badge_keys.append("sign")
        if opts.get("sub", False): active_badge_keys.append("sub")
        if opts.get("des", False): active_badge_keys.append("des")
        age = opts.get("age", "none")
        if age in ["all", "7", "12", "15", "19"]: active_badge_keys.append(age)

        for bk in active_badge_keys:
            img_path = self.custom_badge_paths.get(bk, os.path.join(self.badge_folder, f"{bk}.png"))
            if os.path.exists(img_path) and ImageTk:
                try:
                    cache = self.__dict__.setdefault("_badge_cache", {})
                    cache_key = (img_path, badge_sz_render)
                    tk_img = cache.get(cache_key)
                    if tk_img is None:
                        pil_img = Image.open(img_path)
                        pil_img = pil_img.resize((badge_sz_render, badge_sz_render), Image.Resampling.LANCZOS)
                        tk_img = ImageTk.PhotoImage(pil_img)
                        cache[cache_key] = tk_img
                    c.create_image(right_x - (badge_sz_render / 2), badge_y + (badge_sz_render / 2), image=tk_img, anchor="center")
                    right_x -= (badge_sz_render + 3)
                except Exception:
                    pass
            else:
                if bk == "uhd": fallback_txt = "UHD"
                elif bk == "all": fallback_txt = "ALL"
                elif bk == "local": fallback_txt = "L"
                elif bk == "sign": fallback_txt = "수"
                elif bk == "sub": fallback_txt = "자"
                elif bk == "des": fallback_txt = "해"
                else: fallback_txt = bk
                
                c.create_text(right_x - 6, badge_y + 4, text=fallback_txt, anchor="center", font=(self.ui_font_family, max(4, int(badge_sz_render * 0.6)), "bold"), fill="#dc2626")
                right_x -= (badge_sz_render + 4)

        text_id = None
        previewing_split = is_selected and (self.h_split_count > 1 or self.v_split_count > 1)
        if previewing_split:
            text_id = self._draw_live_split_texts(c, b, x0, y0, x1, y1)
        elif b["title"]:
            bw = x1 - x0
            bh = y1 - y0
            formatted_text, f_size = self._format_smart_text(b["title"], bw / self.zoom, bh / self.zoom, size_offset=b.get("font_size_offset", 0))
            
            cx = x0 + bw / 2
            offset_y = 0
            if top_enabled: offset_y += top_fs_render / 2
            if bottom_enabled: offset_y -= bottom_fs_render / 2
            cx += b.get("text_dx", 0) * self.zoom
            cy = y0 + bh / 2 + offset_y + b.get("text_dy", 0) * self.zoom
            
            text_id = c.create_text(
                cx, cy, 
                text=formatted_text, 
                anchor="center", 
                font=(self.ui_font_family, -max(4, int(round(f_size * self.zoom))), "bold"), 
                justify="center",
                fill=b.get("text_color", "#000000")
            )
        
        grip_id = None
        edge_handles = []
        if is_selected and not self.lock_move.get():
            grip_id = c.create_rectangle(x1-10, y1-10, x1, y1, fill="#2563eb", outline="white")
            if True:
                gid_h = b.get("group_id")
                if gid_h is None:
                    hx0, hy0, hx1, hy1 = x0, y0, x1, y1
                else:
                    grs = [gb["rect"] for gb in self.blocks if gb.get("group_id") == gid_h]
                    hx0, hy0, hx1, hy1 = self._xy((min(r[0] for r in grs), min(r[1] for r in grs),
                                                   max(r[2] for r in grs), max(r[3] for r in grs)))
                mx, my = (hx0 + hx1) / 2, (hy0 + hy1) / 2
                hlist = [("nw", hx0, hy0), ("n", mx, hy0), ("ne", hx1, hy0), ("w", hx0, my),
                         ("e", hx1, my), ("sw", hx0, hy1), ("s", mx, hy1)]
                if gid_h is not None:
                    hlist.append(("se", hx1, hy1))
                for hm, hx, hy in hlist:
                    hid = c.create_rectangle(hx - 4, hy - 4, hx + 4, hy + 4, fill="#2563eb", outline="white")
                    edge_handles.append((hid, "rs_" + hm))
            
        self.item_map[hit_id] = (b, "move")
        self.item_map[rect_id] = (b, "move")
        if text_id:
            self.item_map[text_id] = (b, "move")
        if grip_id:
            self.item_map[grip_id] = (b, "resize")
        for hid, hmode in edge_handles:
            self.item_map[hid] = (b, hmode)

    def draw_pdf(self, refresh_sidebar=True):
        if not self.png_data: return
        c = self.canvas
        c.delete("all")
        if getattr(self, "_photo_src", None) is not self.png_data or getattr(self, "photo", None) is None:
            self.photo = tk.PhotoImage(data=base64.b64encode(self.png_data).decode("ascii"))   # 페이지 그림은 바뀔 때만 다시 만든다
            self._photo_src = self.png_data
        
        pw = (self.page_rect[2] - self.page_rect[0]) if self.page_rect else 1000
        ph = (self.page_rect[3] - self.page_rect[1]) if self.page_rect else 800
        page_width, page_height = pw * self.zoom, ph * self.zoom
        
        c.create_rectangle(18, 18, 18 + page_width, 18 + page_height, fill="#ffffff", outline="")   # 종이 바탕(투명한 곳이 회색으로 보이지 않게)
        c.create_image(18, 18, image=self.photo, anchor="nw")
        c.configure(scrollregion=(0, 0, page_width + 36, page_height + 36))

        self.item_map.clear()
        
        self._mod_ids = {b["id"] for b in self.blocks if self._is_modified(b)}
        if len(self.__dict__.get("_shade_cache", {})) > 400:
            self._shade_cache = {}

        unselected_blocks = [b for b in self.blocks if b["id"] not in self.selected]
        for b in unselected_blocks:
            self._render_single_block(b, is_selected=False)
            
        selected_blocks = [b for b in self.blocks if b["id"] in self.selected]
        for b in selected_blocks:
            self._render_single_block(b, is_selected=True)
            
        for b in selected_blocks:
            x0, y0, x1, y1 = self._xy(b["rect"])
            if self.h_split_count > 1:
                if len(self.split_ratios) == self.h_split_count:
                    curr_y = y0
                    total_h = y1 - y0
                    for i in range(1, self.h_split_count):
                        curr_y += total_h * self.split_ratios[i-1]
                        c.create_line(x0, curr_y, x1, curr_y, fill="#ef4444", dash=(4, 4), width=2)
                else:
                    step_h = (y1 - y0) / self.h_split_count
                    for i in range(1, self.h_split_count):
                        sy = y0 + step_h * i
                        c.create_line(x0, sy, x1, sy, fill="#ef4444", dash=(4, 4), width=2)
            elif self.v_split_count > 1:
                if len(self.split_ratios) == self.v_split_count:
                    curr_x = x0
                    total_w = x1 - x0
                    for i in range(1, self.v_split_count):
                        curr_x += total_w * self.split_ratios[i-1]
                        c.create_line(curr_x, y0, curr_x, y1, fill="#10b981", dash=(4, 4), width=2)
                else:
                    step_w = (x1 - x0) / self.v_split_count
                    for i in range(1, self.v_split_count):
                        sx = x0 + step_w * i
                        c.create_line(sx, y0, sx, y1, fill="#10b981", dash=(4, 4), width=2)

        if self.show_changed:
            ch = self.changed_ids_by_page.get(self.current_page) or set()
            for b in self.blocks:
                if b["id"] in ch:
                    x0, y0, x1, y1 = self._xy(b["rect"])
                    c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y1 - 1, outline="#dc2626", width=2, dash=(5, 3))
        self._paint_header_on_canvas(c)       # 머리글(제목·날짜·요일 칸)은 맨 위에 - 블록에 가려지지 않게
        self._update_mod_count()
        if refresh_sidebar:
            self.update_sidebar_info()
        self._schedule_autosave()

    def _schedule_autosave(self, delay=800):
        """바뀐 뒤 상태가 꼭 자동 저장되도록, 다시 그릴 때마다 잠시 뒤 한 번 저장한다 (드래그 중에는 묶어서 한 번)."""
        job = getattr(self, "_autosave_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self._autosave_job = self.root.after(delay, self._run_autosave)

    def _run_autosave(self):
        self._autosave_job = None
        self.save_autosave()

    def update_sidebar_info(self):
        if not hasattr(self, "info_frame"): return
        
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        sel_count = len(self.selected)
        
        if sel_count == 0:
            ttk.Label(self.info_frame, text="블록을 선택하면 편집 패널이 활성화됩니다.", font=(self.ui_font_family, 9), foreground="#64748b").pack(pady=20)
            return
        self._build_pick_shade_row()
            
        if sel_count == 1:
            sel_id = list(self.selected)[0]
            block = next((b for b in self.blocks if b["id"] == sel_id), None)
            if not block: return
            
            group_txt = f" (연동 그룹: #{block['group_id']})" if block.get("group_id") else ""
            ttk.Label(self.info_frame, text=f"선택 블록 ID: #{block['id']}{group_txt}", font=(self.ui_font_family, 9, "bold"), foreground="#2563eb").pack(anchor="w", pady=(0, 4))
            
            preview_canvas_w, preview_canvas_h = 270, 110
            self.shape_canvas = tk.Canvas(self.info_frame, width=preview_canvas_w, height=preview_canvas_h, bg="#f8fafc", highlightthickness=1, highlightbackground="#cbd5e1")
            self.shape_canvas.pack(pady=2)
            self._paint_shape_preview(block)

            ttk.Label(self.info_frame, text="프로그램 명칭 입력 (엔터키로 줄바꿈 가능):", font=(self.ui_font_family, 8, "bold")).pack(anchor="w", pady=(4, 1))
            
            text_frame = ttk.Frame(self.info_frame)
            text_frame.pack(fill="x", pady=2)
            
            self.text_widget = tk.Text(text_frame, height=3, font=(self.ui_font_family, 9), wrap="word")
            self.text_widget.pack(side="left", fill="x", expand=True)
            self.text_widget.insert("1.0", block["title"])
            
            t_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text_widget.yview)
            t_scroll.pack(side="right", fill="y")
            self.text_widget.configure(yscrollcommand=t_scroll.set)
            
            def on_text_modified(event=None):
                self.confirm_title_input(block)

            self.text_widget.bind("<KeyRelease>", on_text_modified)
            self.text_widget.bind("<FocusOut>", on_text_modified)
            
            total_splits = max(self.h_split_count, self.v_split_count)
            if total_splits <= 1:
                self.text_widget.focus_set()
                self.text_widget.mark_set("insert", "end")
            
            col_row = ttk.Frame(self.info_frame)
            col_row.pack(fill="x", pady=4)
            tk.Button(col_row, text=" 배경색 설정(투명/색상) ", bg=block["color"] if block["color"] else "#f8fafc", font=(self.ui_font_family, 8), command=lambda: self.pick_single_color(block, 'bg')).pack(side="left", fill="x", expand=True, padx=(0, 2))
            tk.Button(col_row, text=" 글씨색 변경 ", bg=block.get("text_color", "#000000"), fg="white", font=(self.ui_font_family, 8), command=lambda: self.pick_single_color(block, 'txt')).pack(side="right", fill="x", expand=True, padx=(2, 0))
            
            opt_frame = ttk.LabelFrame(self.info_frame, text=" 📺 방송 옵션 및 뱃지 설정 ", padding=6)
            opt_frame.pack(fill="x", pady=4)
            
            if "options" not in block:
                block["options"] = {}
            opts = block["options"]
            if self._is_modified(block):
                self._sync_auto_times(block)          # 나누거나 옮긴 블록: 시간 값을 지금 자리에 맞춘다
            
            self.opt_top_time_enabled_var = tk.BooleanVar(value=opts.get("top_time_enabled", opts.get("time_enabled", False)))
            self.top_time_val_var = tk.StringVar(value=opts.get("top_time_val", opts.get("time_val", "00")))
            
            self.opt_bottom_time_enabled_var = tk.BooleanVar(value=opts.get("bottom_time_enabled", False))
            self.bottom_time_val_var = tk.StringVar(value=opts.get("bottom_time_val", "00"))

            row_top_time = ttk.Frame(opt_frame)
            row_top_time.pack(fill="x", pady=2)
            
            ttk.Checkbutton(row_top_time, text="상단시간(위):", variable=self.opt_top_time_enabled_var, command=lambda: self._on_time_check(block, "top")).pack(side="left", padx=(0, 2))
            top_entry = ttk.Entry(row_top_time, textvariable=self.top_time_val_var, width=4, font=(self.ui_font_family, 8))
            top_entry.pack(side="left", padx=(0, 4))
            top_entry.bind("<KeyRelease>", lambda e: self.on_option_changed(block))
            top_entry.bind("<FocusOut>", lambda e: self.on_option_changed(block))

            top_fs_val = opts.get("top_time_font_size", opts.get("time_font_size", DEFAULT_TIME_FONT_SIZE))
            ttk.Button(row_top_time, text="+", width=2, command=lambda: self.modify_time_font_size(block, 'top', 1)).pack(side="right", padx=1)
            self.top_time_size_lbl = ttk.Label(row_top_time, text=f"{top_fs_val} pt", width=4, anchor="center", font=(self.ui_font_family, 8, "bold"), background="#f1f5f9")
            self.top_time_size_lbl.pack(side="right", padx=1)
            ttk.Button(row_top_time, text="-", width=2, command=lambda: self.modify_time_font_size(block, 'top', -1)).pack(side="right", padx=1)
            ttk.Label(row_top_time, text="│ 크기:", font=(self.ui_font_family, 7, "bold")).pack(side="right", padx=(2, 4))

            row_bottom_time = ttk.Frame(opt_frame)
            row_bottom_time.pack(fill="x", pady=2)
            
            ttk.Checkbutton(row_bottom_time, text="하단시간(아래):", variable=self.opt_bottom_time_enabled_var, command=lambda: self._on_time_check(block, "bottom")).pack(side="left", padx=(0, 2))
            bottom_entry = ttk.Entry(row_bottom_time, textvariable=self.bottom_time_val_var, width=4, font=(self.ui_font_family, 8))
            bottom_entry.pack(side="left", padx=(0, 4))
            bottom_entry.bind("<KeyRelease>", lambda e: self.on_option_changed(block))
            bottom_entry.bind("<FocusOut>", lambda e: self.on_option_changed(block))

            bot_fs_val = opts.get("bottom_time_font_size", opts.get("time_font_size", DEFAULT_TIME_FONT_SIZE))
            ttk.Button(row_bottom_time, text="+", width=2, command=lambda: self.modify_time_font_size(block, 'bottom', 1)).pack(side="right", padx=1)
            self.bottom_time_size_lbl = ttk.Label(row_bottom_time, text=f"{bot_fs_val} pt", width=4, anchor="center", font=(self.ui_font_family, 8, "bold"), background="#f1f5f9")
            self.bottom_time_size_lbl.pack(side="right", padx=1)
            ttk.Button(row_bottom_time, text="-", width=2, command=lambda: self.modify_time_font_size(block, 'bottom', -1)).pack(side="right", padx=1)
            ttk.Label(row_bottom_time, text="│ 크기:", font=(self.ui_font_family, 7, "bold")).pack(side="right", padx=(2, 4))

            badge_size_row = ttk.Frame(opt_frame)
            badge_size_row.pack(fill="x", pady=(4, 2))
            ttk.Label(badge_size_row, text="뱃지(옵션) 크기 조절:", font=(self.ui_font_family, 8, "bold")).pack(side="left", padx=2)
            
            badge_sz_val = opts.get("badge_size", 12)
            ttk.Button(badge_size_row, text="+", width=2, command=lambda: self.modify_badge_size(block, 1)).pack(side="right", padx=1)
            self.badge_size_lbl = ttk.Label(badge_size_row, text=f"{badge_sz_val} px", width=5, anchor="center", font=(self.ui_font_family, 8, "bold"), background="#f1f5f9")
            self.badge_size_lbl.pack(side="right", padx=1)
            ttk.Button(badge_size_row, text="-", width=2, command=lambda: self.modify_badge_size(block, -1)).pack(side="right", padx=1)

            row_opt1 = ttk.Frame(opt_frame)
            row_opt1.pack(fill="x", pady=1)
            self.opt_uhd_var = tk.BooleanVar(value=opts.get("uhd", False))
            self.opt_sub_var = tk.BooleanVar(value=opts.get("sub", False))
            self.opt_sign_var = tk.BooleanVar(value=opts.get("sign", False))
            ttk.Checkbutton(row_opt1, text="UHD", variable=self.opt_uhd_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Checkbutton(row_opt1, text="자막(자)", variable=self.opt_sub_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Checkbutton(row_opt1, text="한국수어(수)", variable=self.opt_sign_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            
            row_opt2 = ttk.Frame(opt_frame)
            row_opt2.pack(fill="x", pady=1)
            self.opt_des_var = tk.BooleanVar(value=opts.get("des", False))
            self.opt_local_var = tk.BooleanVar(value=opts.get("local", False))
            ttk.Checkbutton(row_opt2, text="화면해설(해)", variable=self.opt_des_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Checkbutton(row_opt2, text="부분로컬(L)", variable=self.opt_local_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)

            age_frame = ttk.Frame(opt_frame)
            age_frame.pack(fill="x", pady=(3, 1))
            ttk.Label(age_frame, text="연령고지:", font=(self.ui_font_family, 8, "bold")).pack(anchor="w", pady=(0, 2))
            
            self.age_var = tk.StringVar(value=opts.get("age", "none"))
            
            age_row1 = ttk.Frame(age_frame)
            age_row1.pack(fill="x", pady=1)
            ttk.Radiobutton(age_row1, text="없음", value="none", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Radiobutton(age_row1, text="전연령", value="all", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Radiobutton(age_row1, text="7세", value="7", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)

            age_row2 = ttk.Frame(age_frame)
            age_row2.pack(fill="x", pady=1)
            ttk.Radiobutton(age_row2, text="12세", value="12", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Radiobutton(age_row2, text="15세", value="15", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)
            ttk.Radiobutton(age_row2, text="19세", value="19", variable=self.age_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=2)

            font_row = ttk.Frame(self.info_frame, padding=(0, 2))
            font_row.pack(fill="x", pady=2)
            ttk.Label(font_row, text="글자 크기 미세 조절:", font=(self.ui_font_family, 8, "bold")).pack(side="left")
            
            current_offset = block.get("font_size_offset", 0)
            ttk.Button(font_row, text="-", width=2, command=lambda: self.modify_font_offset(block, -1)).pack(side="right", padx=1)
            self.font_offset_lbl = ttk.Label(font_row, text=f"{'+' if current_offset > 0 else ''}{current_offset}", width=5, anchor="center", font=(self.ui_font_family, 8, "bold"), background="#f1f5f9")
            self.font_offset_lbl.pack(side="right", padx=1)
            ttk.Button(font_row, text="+", width=2, command=lambda: self.modify_font_offset(block, 1)).pack(side="right", padx=1)

            pos_rows = (
                ("y", "글자 위치 상하(pt):", "text_dy_lbl", "▲", "▼", -block.get("text_dy", 0)),
                ("x", "글자 위치 좌우(pt):", "text_dx_lbl", "◀", "▶", block.get("text_dx", 0)),
            )
            for axis, label_txt, attr, minus_txt, plus_txt, shown in pos_rows:
                pos_row = ttk.Frame(self.info_frame, padding=(0, 2))
                pos_row.pack(fill="x", pady=2)
                ttk.Label(pos_row, text=label_txt, font=(self.ui_font_family, 8, "bold")).pack(side="left")
                # 상하: ▲ = 위로(dy -1), ▼ = 아래로(dy +1) / 좌우: ◀ = 왼쪽(dx -1), ▶ = 오른쪽(dx +1)
                ttk.Button(pos_row, text=plus_txt, width=3, command=lambda a=axis: self.modify_text_pos(block, a, 1)).pack(side="right", padx=1)
                lbl = ttk.Label(pos_row, text=f"{shown:+d}" if shown else "0", width=5, anchor="center", font=(self.ui_font_family, 8, "bold"), background="#f1f5f9")
                lbl.pack(side="right", padx=1)
                setattr(self, attr, lbl)
                ttk.Button(pos_row, text=minus_txt, width=3, command=lambda a=axis: self.modify_text_pos(block, a, -1)).pack(side="right", padx=1)

            if block is not None:
                edge_frame = ttk.LabelFrame(self.info_frame, text=" 📐 블록 크기 조절 (상·하·좌·우 변) ", padding=6)
                edge_frame.pack(fill="x", pady=4)
                step_row = ttk.Frame(edge_frame)
                step_row.pack(fill="x", pady=(0, 2))
                ttk.Label(step_row, text="이동 단위(pt):", font=(self.ui_font_family, 8, "bold")).pack(side="left")
                if not hasattr(self, "edge_step_var"):
                    self.edge_step_var = tk.IntVar(value=1)
                for v in (1, 2, 5, 10):
                    ttk.Radiobutton(step_row, text=str(v), value=v, variable=self.edge_step_var).pack(side="left", padx=2)
                ttk.Label(edge_frame, text="화살표 방향으로 해당 변이 이동합니다.", font=(self.ui_font_family, 7), foreground="#666").pack(anchor="w")
                for edge, label_txt, minus_txt, plus_txt in (("top", "위쪽 변", "▲", "▼"), ("bottom", "아래쪽 변", "▲", "▼"),
                                                             ("left", "왼쪽 변", "◀", "▶"), ("right", "오른쪽 변", "◀", "▶")):
                    erow = ttk.Frame(edge_frame)
                    erow.pack(fill="x", pady=1)
                    ttk.Label(erow, text=label_txt, width=9, font=(self.ui_font_family, 8, "bold")).pack(side="left")
                    ttk.Button(erow, text=plus_txt, width=3, command=lambda e=edge: self.modify_block_edge(block, e, 1)).pack(side="right", padx=1)
                    ttk.Button(erow, text=minus_txt, width=3, command=lambda e=edge: self.modify_block_edge(block, e, -1)).pack(side="right", padx=1)

            self.split_title_vars = {}             # (예전 '선택적 블록 분할' 패널은 툴바 '✂ 나누기 (S)' 로 하나로 합침)
            if block.get("group_id") is not None:
                _dv = tk.BooleanVar(value=bool(block.get("split_dash")))
                ttk.Checkbutton(self.info_frame, text="┄ 나눈 경계선을 점선으로", variable=_dv,
                                command=lambda: self.set_split_dash(block, _dv.get())).pack(anchor="w", pady=(4, 0))
            ttk.Label(self.info_frame, text="✂ 칸 나누기: 툴바의 [✂ 나누기 (S)]를 켜거나 S 키 → 칸 안을 클릭 (5분 눈금에 맞춤, Shift+클릭 = 좌우)",
                      font=(self.ui_font_family, 8), foreground="#64748b", wraplength=330, justify="left").pack(anchor="w", pady=(4, 0))

            btn_row = ttk.Frame(self.info_frame)
            btn_row.pack(fill="x", pady=4)
            ttk.Button(btn_row, text="✔ 적용 확인", command=self.apply_sidebar_changes).pack(side="left", fill="x", expand=True, padx=(0, 2))
            ttk.Button(btn_row, text="⭐ 파렛트 저장", command=self.save_to_favorites).pack(side="right", fill="x", expand=True, padx=(2, 0))

        else:
            ttk.Label(self.info_frame, text=f"📦 다중 선택됨 ({sel_count}개 블록)", font=(self.ui_font_family, 9, "bold"), foreground="#2563eb").pack(anchor="w", pady=(0, 4))
            container = ttk.Frame(self.info_frame)
            container.pack(fill="both", expand=True, pady=2)
            
            canvas_multi = tk.Canvas(container, borderwidth=0, background="#f8fafc", highlightthickness=0)
            scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas_multi.yview)
            scroll_frame = ttk.Frame(canvas_multi)
            
            scroll_frame.bind("<Configure>", lambda e: canvas_multi.configure(scrollregion=canvas_multi.bbox("all")))
            canvas_multi.create_window((0, 0), window=scroll_frame, anchor="nw")
            canvas_multi.configure(yscrollcommand=scrollbar.set)
            
            canvas_multi.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            self.multi_title_vars = {}
            for b_id in sorted(list(self.selected)):
                block = next((b for b in self.blocks if b["id"] == b_id), None)
                if not block: continue
                
                row_f = ttk.Frame(scroll_frame, padding=2)
                row_f.pack(fill="x", pady=2)
                ttk.Label(row_f, text=f"#{block['id']}", font=(self.ui_font_family, 8, "bold"), width=4).pack(side="left")
                
                t_var = tk.StringVar(value=block["title"])
                self.multi_title_vars[block["id"]] = t_var
                ttk.Entry(row_f, textvariable=t_var, font=(self.ui_font_family, 8)).pack(side="left", fill="x", expand=True, padx=2)
                
            ttk.Button(self.info_frame, text="✔ 다중 블록 변경 적용", command=self.apply_multi_changes).pack(fill="x", pady=(6, 2))
            ttk.Button(self.info_frame, text="선택 블록 합치기", command=self.merge_selected).pack(fill="x", pady=(2, 2))

    def modify_time_font_size(self, block, target, delta):
        if "options" not in block:
            block["options"] = {}
        
        key = "top_time_font_size" if target == 'top' else "bottom_time_font_size"
        current = block["options"].get(key, block["options"].get("time_font_size", DEFAULT_TIME_FONT_SIZE))
        new_val = max(4, min(16, current + delta))
        if new_val != current:
            self._push_coalesced("opt")
        block["options"][key] = new_val
        
        if target == 'top':
            if hasattr(self, "top_time_size_lbl") and self.top_time_size_lbl.winfo_exists():
                self.top_time_size_lbl.config(text=f"{new_val} pt")
        else:
            if hasattr(self, "bottom_time_size_lbl") and self.bottom_time_size_lbl.winfo_exists():
                self.bottom_time_size_lbl.config(text=f"{new_val} pt")
                
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def modify_badge_size(self, block, delta):
        if "options" not in block:
            block["options"] = {}
        current = block["options"].get("badge_size", 12)
        new_val = max(2, current + delta)
        if new_val != current:
            self._push_coalesced("opt")
        block["options"]["badge_size"] = new_val
        
        if hasattr(self, "badge_size_lbl") and self.badge_size_lbl.winfo_exists():
            self.badge_size_lbl.config(text=f"{new_val} px")
            
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def modify_font_offset(self, block, delta):
        current = block.get("font_size_offset", 0)
        new_val = max(-4, min(20, current + delta))
        if new_val != current:
            self._push_coalesced("opt")
        block["font_size_offset"] = new_val
        
        if hasattr(self, "font_offset_lbl") and self.font_offset_lbl.winfo_exists():
            disp = f"{'+' if new_val > 0 else ''}{new_val}"
            self.font_offset_lbl.config(text=disp)
            
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def _apply_group_rect(self, group_blocks, rect):
        """분할 그룹의 전체 외곽을 rect로 바꾸고, 각 조각은 rel_box 비율대로 다시 배치한다."""
        gx0, gy0, gx1, gy1 = rect
        gw, gh = gx1 - gx0, gy1 - gy0
        for gb in group_blocks:
            rx0, ry0, rx1, ry1 = (gb.get("rel_box") or (0, 0, 1, 1))
            gb["rect"] = (gx0 + gw * rx0, gy0 + gh * ry0, gx0 + gw * rx1, gy0 + gh * ry1)

    def modify_block_edge(self, block, edge, sign):
        """선택 블록의 한 변(위/아래/왼쪽/오른쪽)을 sign(-1/+1) x 이동 단위(pt)만큼 옮긴다."""
        if self._move_locked():
            return
        try:
            step = int(self.edge_step_var.get())
        except Exception:
            step = 1
        d = sign * step
        gid = block.get("group_id")
        group_blocks = [gb for gb in self.blocks if gid is not None and gb.get("group_id") == gid]
        if group_blocks:   # 분할 블록: 그룹 전체 외곽 기준
            x0 = min(gb["rect"][0] for gb in group_blocks)
            y0 = min(gb["rect"][1] for gb in group_blocks)
            x1 = max(gb["rect"][2] for gb in group_blocks)
            y1 = max(gb["rect"][3] for gb in group_blocks)
        else:
            x0, y0, x1, y1 = block["rect"]
        cur_rect = (x0, y0, x1, y1)
        if edge == "top":
            y0 = min(y0 + d, y1 - 10)
        elif edge == "bottom":
            y1 = max(y1 + d, y0 + 10)
        elif edge == "left":
            x0 = min(x0 + d, x1 - 20)
        elif edge == "right":
            x1 = max(x1 + d, x0 + 20)
        new_rect = (x0, y0, x1, y1)
        if new_rect == cur_rect:
            return
        self.push_history()
        if group_blocks:
            self._apply_group_rect(group_blocks, new_rect)
        else:
            block["rect"] = new_rect
        self.page_blocks[self.current_page] = self.blocks
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)
        self.save_autosave()

    def modify_text_pos(self, block, axis, delta):
        key = "text_dy" if axis == 'y' else "text_dx"
        new_val = max(-60, min(60, block.get(key, 0) + delta))
        if new_val == block.get(key, 0):
            return
        self.push_history()
        block[key] = new_val

        if axis == 'y':
            lbl = getattr(self, "text_dy_lbl", None)
            shown = -new_val
        else:
            lbl = getattr(self, "text_dx_lbl", None)
            shown = new_val
        if lbl is not None and lbl.winfo_exists():
            lbl.config(text=f"{shown:+d}" if shown else "0")

        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def on_option_changed(self, block):
        if not hasattr(self, "opt_top_time_enabled_var"):
            return
        top_fs = block.get("options", {}).get("top_time_font_size", DEFAULT_TIME_FONT_SIZE)
        bottom_fs = block.get("options", {}).get("bottom_time_font_size", DEFAULT_TIME_FONT_SIZE)
        badge_sz = block.get("options", {}).get("badge_size", 12)
        new_opts = {
            "top_time_enabled": self.opt_top_time_enabled_var.get(),
            "top_time_val": self.top_time_val_var.get(),
            "top_time_font_size": top_fs,
            "bottom_time_enabled": self.opt_bottom_time_enabled_var.get(),
            "bottom_time_val": self.bottom_time_val_var.get(),
            "bottom_time_font_size": bottom_fs,
            "badge_size": badge_sz,
            "uhd": self.opt_uhd_var.get(),
            "sub": self.opt_sub_var.get(),
            "sign": self.opt_sign_var.get(),
            "des": self.opt_des_var.get(),
            "local": self.opt_local_var.get(),
            "age": self.age_var.get()
        }
        self._keep_auto_flags(block, new_opts)
        if new_opts != block.get("options"):
            self._push_coalesced("opt")
            block["options"] = new_opts
        self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def on_ratio_slider_change(self, val):
        r1 = float(val)
        r2 = 1.0 - r1
        self.split_ratios = [r1, r2]
        block = self._selected_block()
        if block:
            self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def adjust_v_ratio(self, delta):
        if len(self.split_ratios) != 2:
            self.split_ratios = [0.5, 0.5]
        new_r1 = max(0.1, min(0.9, round(self.split_ratios[0] + delta, 2)))
        new_r2 = round(1.0 - new_r1, 2)
        self.split_ratios = [new_r1, new_r2]
        if hasattr(self, "ratio_scale"):
            self.ratio_scale.set(new_r1)
        block = self._selected_block()
        if block:
            self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def on_h_ratio_slider_change(self, val):
        r1 = float(val)
        r2 = 1.0 - r1
        self.split_ratios = [r1, r2]
        block = self._selected_block()
        if block:
            self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def adjust_h_ratio(self, delta):
        if len(self.split_ratios) != 2:
            self.split_ratios = [0.5, 0.5]
        new_r1 = max(0.1, min(0.9, round(self.split_ratios[0] + delta, 2)))
        new_r2 = round(1.0 - new_r1, 2)
        self.split_ratios = [new_r1, new_r2]
        if hasattr(self, "h_ratio_scale"):
            self.h_ratio_scale.set(new_r1)
        block = self._selected_block()
        if block:
            self._paint_shape_preview(block)
        self.draw_pdf(refresh_sidebar=False)

    def confirm_title_input(self, block):
        if hasattr(self, "text_widget"):
            new_title = self.text_widget.get("1.0", "end-1c")
            if block["title"] != new_title:
                self._push_coalesced("title", window=3.0)
                block["title"] = new_title
                self._paint_shape_preview(block)
                self.draw_pdf(refresh_sidebar=False)

    def confirm_split_title(self, index, text, immediate=False):
        self.split_titles[index] = text
        if immediate:
            job = getattr(self, "_split_preview_job", None)
            if job is not None:
                try:
                    self.root.after_cancel(job)
                except Exception:
                    pass
                self._split_preview_job = None
            self._refresh_split_preview()
        else:
            self._schedule_split_preview()

    def apply_multi_changes(self):
        self.push_history()
        for b_id, var in getattr(self, "multi_title_vars", {}).items():
            block = next((b for b in self.blocks if b["id"] == b_id), None)
            if block:
                block["title"] = var.get()
        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()
        messagebox.showinfo("적용 완료", "다중 선택된 블록들에 변경사항이 적용되었습니다.")

    def pick_single_color(self, block, target):
        if target == 'bg':
            initial = block["color"] if block["color"] else "#ffffff"
        else:
            initial = block.get("text_color", "#000000")
            
        c = colorchooser.askcolor(title="색상 선택", initialcolor=initial)
        if c[1]:
            self.push_history()
            if target == 'bg':
                block["color"] = c[1]
            else:
                block["text_color"] = c[1]
            self.draw_pdf()
            self.save_autosave()

    def apply_sidebar_changes(self):
        if len(self.selected) != 1:
            return
        sel_id = list(self.selected)[0]
        block = next((b for b in self.blocks if b["id"] == sel_id), None)
        if not block: return
        
        self.push_history()
        
        if hasattr(self, "text_widget"):
            block["title"] = self.text_widget.get("1.0", "end-1c")
        if hasattr(self, "opt_top_time_enabled_var"):
            top_fs = block.get("options", {}).get("top_time_font_size", DEFAULT_TIME_FONT_SIZE)
            bottom_fs = block.get("options", {}).get("bottom_time_font_size", DEFAULT_TIME_FONT_SIZE)
            badge_sz = block.get("options", {}).get("badge_size", 12)
            block["options"] = self._keep_auto_flags(block, {
                "top_time_enabled": self.opt_top_time_enabled_var.get(),
                "top_time_val": self.top_time_val_var.get(),
                "top_time_font_size": top_fs,
                "bottom_time_enabled": self.opt_bottom_time_enabled_var.get(),
                "bottom_time_val": self.bottom_time_val_var.get(),
                "bottom_time_font_size": bottom_fs,
                "badge_size": badge_sz,
                "uhd": self.opt_uhd_var.get(),
                "sub": self.opt_sub_var.get(),
                "sign": self.opt_sign_var.get(),
                "des": self.opt_des_var.get(),
                "local": self.opt_local_var.get(),
                "age": self.age_var.get()
            })
        for i, var in getattr(self, "split_title_vars", {}).items():
            try:
                self.split_titles[i] = var.get()
            except tk.TclError:
                pass
        
        if self.h_split_count > 1 or self.v_split_count > 1:
            x0, y0, x1, y1 = block["rect"]
            color = block["color"]
            text_color = block.get("text_color", "#000000")
            options = copy.deepcopy(block["options"])
            
            self.blocks = [b for b in self.blocks if b["id"] != sel_id]
            new_blocks = []
            
            group_id = self.next_id
            self.next_id += 1
            
            if self.h_split_count > 1:
                total_h = y1 - y0
                if len(self.split_ratios) != self.h_split_count:
                    self.split_ratios = [1.0 / self.h_split_count] * self.h_split_count
                
                curr_y = y0
                for i in range(self.h_split_count):
                    h_ratio = self.split_ratios[i]
                    sy0 = curr_y
                    sy1 = curr_y + total_h * h_ratio
                    r = (x0, sy0, x1, sy1)
                    sub_title = self.split_titles.get(i, "")
                    
                    rel_start = (sy0 - y0) / total_h
                    rel_end = (sy1 - y0) / total_h
                    
                    new_blocks.append({
                        "id": self.next_id,
                        "title": sub_title,
                        "color": color,
                        "text_color": text_color,
                        "font_size_offset": block.get("font_size_offset", 0),
                        "text_dx": block.get("text_dx", 0),
                        "text_dy": block.get("text_dy", 0),
                        "border_color": block.get("border_color"),
                        "border_width": block.get("border_width"),
                        "orig_fill": block.get("orig_fill", ""),
                        "options": copy.deepcopy(options),
                        "rect": r,
                        "orig_rect": r,
                        "group_id": group_id,
                        "rel_box": (0.0, rel_start, 1.0, rel_end)
                    })
                    self.next_id += 1
                    curr_y = sy1
            elif self.v_split_count > 1:
                total_w = x1 - x0
                if len(self.split_ratios) != self.v_split_count:
                    self.split_ratios = [1.0 / self.v_split_count] * self.v_split_count
                
                curr_x = x0
                for i in range(self.v_split_count):
                    w_ratio = self.split_ratios[i]
                    sx0 = curr_x
                    sx1 = curr_x + total_w * w_ratio
                    r = (sx0, y0, sx1, y1)
                    sub_title = self.split_titles.get(i, "")
                    
                    rel_start = (sx0 - x0) / total_w
                    rel_end = (sx1 - x0) / total_w
                    
                    new_blocks.append({
                        "id": self.next_id,
                        "title": sub_title,
                        "color": color,
                        "text_color": text_color,
                        "font_size_offset": block.get("font_size_offset", 0),
                        "text_dx": block.get("text_dx", 0),
                        "text_dy": block.get("text_dy", 0),
                        "border_color": block.get("border_color"),
                        "border_width": block.get("border_width"),
                        "orig_fill": block.get("orig_fill", ""),
                        "options": copy.deepcopy(options),
                        "rect": r,
                        "orig_rect": r,
                        "group_id": group_id,
                        "rel_box": (rel_start, 0.0, rel_end, 1.0)
                    })
                    self.next_id += 1
                    curr_x = sx1
                    
            self.blocks.extend(new_blocks)
            self.selected = {b["id"] for b in new_blocks}
            
            self.h_split_count = 1
            self.v_split_count = 1
            self.split_ratios.clear()
            self.split_titles.clear()
            messagebox.showinfo("적용 완료", "블록이 분할되고 변경사항이 반영되었습니다.")
        else:
            messagebox.showinfo("적용 완료", "변경사항이 성공적으로 반영되었습니다.")

        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()

    def adjust_split_count(self, axis, delta):
        if axis == 'h':
            self.h_split_count = max(1, min(6, self.h_split_count + delta))
            if self.h_split_count > 1:
                self.v_split_count = 1
                self.split_ratios = [1.0 / self.h_split_count] * self.h_split_count
            else:
                self.split_ratios.clear()
        else:
            self.v_split_count = max(1, min(6, self.v_split_count + delta))
            if self.v_split_count > 1:
                self.h_split_count = 1
                self.split_ratios = [1.0 / self.v_split_count] * self.v_split_count
            else:
                self.split_ratios.clear()
        self.update_sidebar_info()
        self.draw_pdf(refresh_sidebar=False)

    def change_zoom(self, direction):
        if not self.pdf_path: return
        levels = (0.60, 0.75, 0.90, 1.00, 1.25, 1.50, 1.75, 2.00, 2.50, 3.00)
        current = min(range(len(levels)), key=lambda i: abs(levels[i] - self.zoom))
        target = max(0, min(len(levels) - 1, current + direction))
        if target == current: return
        self.zoom = levels[target]
        self.zoom_label.set(f"{round(self.zoom * 100)}%")
        self.saved_view = (self.canvas.xview()[0], self.canvas.yview()[0])
        self.zoom_version += 1
        threading.Thread(target=self._render_zoom_worker, args=(self.zoom, self.zoom_version), daemon=True).start()

    def on_mouse_wheel(self, event):
        if event.state & 0x0004:
            self.change_zoom(1 if event.delta > 0 or event.num == 4 else -1)
            return "break"
        move = -1 if (event.delta > 0 or event.num == 4) else 1 if (event.delta < 0 or event.num == 5) else 0
        self.canvas.yview_scroll(move, "units")
        return "break"

    def _open_base_doc(self):
        data = getattr(self, "base_pdf_bytes", None)
        if data:
            return fitz.open(stream=data, filetype="pdf")
        return fitz.open(self.pdf_path)

    def _render_zoom_worker(self, zoom, version):
        try:
            with PDF_LOCK:
                doc = self._open_base_doc()
                page = doc.load_page(self.current_page)
                png_data = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=True).tobytes("png")
                doc.close()
            self.root.after(0, self._zoom_finished, version, png_data, None)
        except Exception as exc:
            self.root.after(0, self._zoom_finished, version, None, exc)

    def _zoom_finished(self, version, png_data, error):
        if version != self.zoom_version or not self.pdf_path or error: return
        self.png_data = png_data
        self.draw_pdf(refresh_sidebar=False)
        try:
            self.canvas.xview_moveto(self.saved_view[0])
            self.canvas.yview_moveto(self.saved_view[1])
        except Exception:
            pass

    def clear_all_blocks(self):
        if not self.pdf_path: return
        if messagebox.askyesno("작업 초기화", "이 페이지의 모든 작업 내역을 취소하고 원본 블록 상태로 되돌리시겠습니까?"):
            self.push_history()
            self.blocks = copy.deepcopy(self.initial_blocks.get(self.current_page, []))
            self.page_blocks[self.current_page] = self.blocks
            self.selected.clear()
            self.split_titles.clear()
            self.draw_pdf()
            self.save_autosave()
            messagebox.showinfo("초기화 완료", "원본 감지 블록 상태로 초기화되었습니다.")

    def delete_selected_blocks(self):
        if not self.selected:
            return
        if self._typing():
            return
            
        self.push_history()
        self.blocks = [b for b in self.blocks if b["id"] not in self.selected]
        self.selected.clear()
        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()

    def merge_selected(self):
        if len(self.selected) < 2:
            messagebox.showinfo("블록 병합", "합칠 블록을 2개 이상 선택해주세요. (Ctrl+클릭으로 다중 선택)")
            return
            
        self.push_history()
        selected_blocks = [b for b in self.blocks if b["id"] in self.selected]
        
        min_x0 = min(b["rect"][0] for b in selected_blocks)
        min_y0 = min(b["rect"][1] for b in selected_blocks)
        max_x1 = max(b["rect"][2] for b in selected_blocks)
        max_y1 = max(b["rect"][3] for b in selected_blocks)
        
        titles = [b["title"] for b in selected_blocks if b["title"]]
        merged_title = "\n".join(titles)
        
        first_b = selected_blocks[0]
        merged_color = first_b["color"]
        merged_text_color = first_b.get("text_color", "#000000")
        merged_options = copy.deepcopy(first_b.get("options", {}))
        
        self.blocks = [b for b in self.blocks if b["id"] not in self.selected]
        
        new_block = {
            "id": self.next_id,
            "title": merged_title,
            "color": merged_color,
            "text_color": merged_text_color,
            "font_size_offset": 0,
            "options": merged_options,
            "text_dx": first_b.get("text_dx", 0),
            "text_dy": first_b.get("text_dy", 0),
            "border_color": next((b.get("border_color") for b in selected_blocks if b.get("border_color")), None),
            "border_width": next((b.get("border_width") for b in selected_blocks if b.get("border_width")), None),
            "orig_fill": first_b.get("orig_fill", ""),
            "rect": (min_x0, min_y0, max_x1, max_y1),
            "orig_rect": (min_x0, min_y0, max_x1, max_y1),
            "group_id": None,
            "rel_box": None
        }
        self.next_id += 1
        self.blocks.append(new_block)
        
        self.selected = {new_block["id"]}
        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()
        messagebox.showinfo("병합 완료", "선택된 블록들이 하나로 병합되었습니다.")

    def preview_pdf(self):
        if not self.pdf_path:
            return
        preview_win = tk.Toplevel(self.root)
        preview_win.title("🔍 PDF 출력 최종 미리보기")
        preview_win.geometry("900x800")
        
        lbl = ttk.Label(preview_win, text="최종 PDF 렌더링 생성 중...", font=(self.ui_font_family, 11))
        lbl.pack(expand=True)
        
        blocks_snap = copy.deepcopy(self.blocks)      # 미리보기를 만드는 동안 블록을 고쳐도 섞이지 않게
        page_no = self.current_page

        def render_preview():
            try:
                with PDF_LOCK:
                    out_bytes = self._generate_pdf_bytes(for_preview=True, blocks=blocks_snap)
                    doc = fitz.open(stream=out_bytes, filetype="pdf")
                    page = doc.load_page(min(page_no, doc.page_count - 1))
                    pix = page.get_pixmap(dpi=150)
                    img_data = pix.tobytes("png")
                    doc.close()
                
                self.root.after(0, show_img, preview_win, lbl, img_data)
            except Exception as e:
                self.root.after(0, show_err, preview_win, str(e))

        def show_img(win, old_lbl, img_bytes):
            old_lbl.destroy()
            photo = tk.PhotoImage(data=base64.b64encode(img_bytes).decode("ascii"))
            
            c = tk.Canvas(win, bg="#64748b", yscrollincrement=30, xscrollincrement=30)
            v = ttk.Scrollbar(win, orient="vertical", command=c.yview)
            h = ttk.Scrollbar(win, orient="horizontal", command=c.xview)
            c.configure(yscrollcommand=v.set, xscrollcommand=h.set)
            
            c.grid(row=0, column=0, sticky="nsew")
            v.grid(row=0, column=1, sticky="ns")
            h.grid(row=1, column=0, sticky="ew")
            win.rowconfigure(0, weight=1)
            win.columnconfigure(0, weight=1)
            
            c.create_image(20, 20, image=photo, anchor="nw")
            c.image = photo
            c.configure(scrollregion=(0, 0, photo.width() + 40, photo.height() + 40))

            def _on_preview_wheel(event):
                if getattr(event, "num", None) == 4:
                    step = -1
                elif getattr(event, "num", None) == 5:
                    step = 1
                else:
                    step = -1 if event.delta > 0 else 1
                if event.state & 0x0001:
                    c.xview_scroll(step * 3, "units")
                else:
                    c.yview_scroll(step * 3, "units")
                return "break"

            win.bind("<MouseWheel>", _on_preview_wheel)
            win.bind("<Button-4>", _on_preview_wheel)
            win.bind("<Button-5>", _on_preview_wheel)

        def show_err(win, err_msg):
            messagebox.showerror("미리보기 실패", err_msg, parent=win)
            win.destroy()

        threading.Thread(target=render_preview, daemon=True).start()

    def export_pdf(self):
        if not self.pdf_path:
            return
        src_dir = os.path.dirname(os.path.abspath(self.pdf_path))
        init_dir = self.last_save_dir if self.last_save_dir and os.path.isdir(self.last_save_dir) else src_dir
        out_path = filedialog.asksaveasfilename(title=("DMB PDF 저장" if IS_DMB else "PDF 저장"), defaultextension=".pdf",
                                                filetypes=[("PDF 파일", "*.pdf")], initialdir=init_dir,
                                                initialfile=self._default_save_name())
        if not out_path:
            return False
        if os.path.abspath(out_path) == os.path.abspath(self.pdf_path):
            messagebox.showwarning("PDF 저장", "받은 원본 PDF 위에는 저장할 수 없습니다.\n다른 이름으로 저장해 주세요.")
            return False
        self.last_save_dir = os.path.dirname(out_path)
            
        try:
            with PDF_LOCK:
                pdf_bytes = self._generate_pdf_bytes()
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            self._mark_saved()
            messagebox.showinfo("저장 완료", f"PDF 파일이 성공적으로 저장되었습니다:\n{out_path}")
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"PDF 파일 저장 실패:\n{str(e)}")
            return False

    def _default_save_name(self):
        """저장 이름 제안. DMB: 'MBC_11__1__2.pdf' -> 'DMB_11__1__2.pdf' / 편성표: 원래 이름 그대로"""
        base = os.path.basename(self.pdf_path or "편성표.pdf")
        if not IS_DMB:
            return base
        if re.match(r"(?i)^dmb[_\s-]", base):
            return base
        if re.match(r"^[A-Za-z]+_", base):
            return re.sub(r"^[A-Za-z]+_", "DMB_", base)
        return "DMB_" + base

    def _generate_pdf_bytes(self, for_preview=False, blocks=None):
        blocks = self.blocks if blocks is None else blocks
        doc = self._open_base_doc()
        page = doc.load_page(self.current_page)
        
        font_path = self.global_font_path
        font_name = "custom_korean_font"
        if font_path and os.path.exists(font_path):
            try:
                page.insert_font(fontname=font_name, fontfile=font_path)
            except Exception:
                font_name = "helv"
        else:
            font_name = "helv"

        _hst = self.header_edit.get(self.current_page) or {}
        _htexts = [_hst.get("title", "") if _hst.get("title_on") else "", _hst.get("date", "") if _hst.get("date_on") else ""]
        if _hst.get("days_on"):
            _htexts += [str(v.get("date", "")) + str(v.get("region", "")) for v in (_hst.get("days") or []) if isinstance(v, dict)]
        if font_name == "helv" and (any(not (b.get("title") or "").isascii() for b in blocks) or any(not t.isascii() for t in _htexts)):
            doc.close()
            raise RuntimeError("한글 폰트를 불러오지 못했습니다.\n설정(⚙️)에서 .ttf/.ttc 한글 폰트 파일을 지정해 주세요.")

        try:
            measure_font = fitz.Font(fontfile=font_path) if font_name == "custom_korean_font" else fitz.Font("helv")
        except Exception:
            measure_font = None

        mod_ids = {b["id"] for b in blocks if self._is_modified(b)}

        # 1) 채우기: 원본 테두리를 덮지 않도록 선 굵기의 절반만큼 안쪽으로 줄여서 그린다.
        for b in blocks:                  # 0) 나눈 칸 등(투명): 원본 글자 자리만 바탕색으로 덮는다 (뱃지 그림·선은 그대로)
            cov = (self._cover_color(b) or "").lstrip("#")
            if len(cov) == 6:
                crgb = tuple(int(cov[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
                for r_ in self._cover_rects(b):
                    page.draw_rect(fitz.Rect(r_), color=None, fill=crgb)
        covered = set()
        for b in blocks:
            hex_c = (b.get("color") or "").lstrip("#")
            if len(hex_c) != 6:
                continue
            covered.add(b["id"])
            rgb = tuple(int(hex_c[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
            r = fitz.Rect(b["rect"])
            hw = self._eff_border_width(b, b["id"] in mod_ids) / 2
            inner = fitz.Rect(r.x0 + hw, r.y0 + hw, r.x1 - hw, r.y1 - hw)
            if inner.width <= 0 or inner.height <= 0:
                inner = r
            page.draw_rect(inner, color=None, fill=rgb)

        # 1-2) 채우기가 덮어버린 원본 선(칸 가장자리)을 원래 위치에 다시 그린다.
        segs = getattr(self, "orig_lines", {}).get(self.current_page, [])
        if segs:
            for b in blocks:
                if b["id"] in covered:
                    _restore_lines(page, segs, fitz.Rect(b["rect"]))

        # 2) 테두리: 원본에서 찾은 칸과 위치/크기가 다른 블록(이동, 크기 변경, 분할, 병합, 수동 추가)은
        #    원본 테두리가 없거나 어긋나 있으므로 같은 스타일로 새로 그린다.
        # 1-3) 수정된 블록 반투명 음영 (미리보기, 또는 '저장 PDF에도 음영 포함'을 켠 경우)
        for b in blocks:
            sc = None
            if b.get("shade"):
                if for_preview or self.pick_shade_in_export:
                    sc = b["shade"]
            elif self.mod_shade_enabled and (for_preview or self.mod_shade_in_export):
                sc = self._shade_color_for(b, b["id"] in mod_ids)
            if sc:
                sh = sc.lstrip("#")
                sh_rgb = tuple(int(sh[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
                page.draw_rect(fitz.Rect(b["rect"]), color=None, fill=sh_rgb, fill_opacity=self.mod_shade_alpha)

        orig_rects = [tuple(ob["rect"]) for ob in self.initial_blocks.get(self.current_page, [])]
        for b in blocks:
            r = fitz.Rect(b["rect"])
            rt = (r.x0, r.y0, r.x1, r.y1)
            modified = b["id"] in mod_ids
            custom_w = modified and self.mod_border_width is not None
            if not custom_w and any(all(abs(a - c) <= 0.5 for a, c in zip(rt, orr)) for orr in orig_rects):
                continue
            col = tuple(b.get("border_color") or DEFAULT_BORDER_COLOR)
            rb = b.get("rel_box")
            if b.get("split_dash") and b.get("group_id") is not None and rb:
                # 나눈 경계선을 점선으로: 바깥 테두리만 실선, 안쪽 경계는 아래에서 점선으로 그린다
                wd = self._eff_border_width(b, modified)
                for (p1, p2), inner in (((r.tl, r.tr), rb[1] > 0.001), ((r.bl, r.br), rb[3] < 0.999),
                                        ((r.tl, r.bl), rb[0] > 0.001), ((r.tr, r.br), rb[2] < 0.999)):
                    if not inner:
                        page.draw_line(p1, p2, color=col, width=wd)
                continue
            page.draw_rect(r, color=col, fill=None, width=self._eff_border_width(b, modified))

        for b in blocks:
            r = fitz.Rect(b["rect"])

            opts = b.get("options", {})
            top_enabled = opts.get("top_time_enabled", False)
            top_val = opts.get("top_time_val", "00")
            bottom_enabled = opts.get("bottom_time_enabled", False)
            bottom_val = opts.get("bottom_time_val", "00")
            
            top_time_fs = opts.get("top_time_font_size", DEFAULT_TIME_FONT_SIZE)
            bottom_time_fs = opts.get("bottom_time_font_size", DEFAULT_TIME_FONT_SIZE)
            badge_sz = opts.get("badge_size", 12)

            if top_enabled and top_val:
                p_top = fitz.Point(r.x0 + 3, r.y0 + top_time_fs + 2)
                page.insert_text(p_top, top_val, fontsize=top_time_fs, fontname=font_name, color=(0.2, 0.25, 0.33))
                
            if bottom_enabled and bottom_val:
                p_bot = fitz.Point(r.x0 + 3, r.y1 - 3)
                page.insert_text(p_bot, bottom_val, fontsize=bottom_time_fs, fontname=font_name, color=(0.2, 0.25, 0.33))

            right_x = r.x1 - 2
            badge_y = r.y0 + 2
            active_badge_keys = []
            if opts.get("uhd", False): active_badge_keys.append("uhd")
            if opts.get("local", False): active_badge_keys.append("local")
            if opts.get("sign", False): active_badge_keys.append("sign")
            if opts.get("sub", False): active_badge_keys.append("sub")
            if opts.get("des", False): active_badge_keys.append("des")
            age = opts.get("age", "none")
            if age in ["all", "7", "12", "15", "19"]: active_badge_keys.append(age)

            for bk in active_badge_keys:
                img_path = self.custom_badge_paths.get(bk, os.path.join(self.badge_folder, f"{bk}.png"))
                badge_rect = fitz.Rect(right_x - badge_sz, badge_y, right_x, badge_y + badge_sz)
                if os.path.exists(img_path):
                    try:
                        page.insert_image(badge_rect, filename=img_path)
                    except Exception:
                        pass
                right_x -= (badge_sz + 2)

            if b["title"]:
                hex_txt = b.get("text_color", "#000000").lstrip("#")
                txt_rgb = (0, 0, 0)
                if len(hex_txt) == 6:
                    txt_rgb = tuple(int(hex_txt[i:i+2], 16)/255.0 for i in (0, 2, 4))
                    
                bw, bh = r.width, r.height
                formatted_text, f_size = self._format_smart_text(b["title"], bw, bh, size_offset=b.get("font_size_offset", 0))
                
                lines = formatted_text.split("\n")
                total_text_h = len(lines) * (f_size * 1.15)
                
                offset_y = 0
                if top_enabled: offset_y += top_time_fs / 2
                if bottom_enabled: offset_y -= bottom_time_fs / 2
                
                start_y = r.y0 + (bh - total_text_h) / 2 + f_size * 0.85 + offset_y + b.get("text_dy", 0)
                
                for idx, line in enumerate(lines):
                    line_y = start_y + idx * (f_size * 1.15)
                    try:
                        line_w = measure_font.text_length(line, fontsize=f_size)
                    except Exception:
                        line_w = sum(f_size if ord(ch) > 127 else f_size * 0.5 for ch in line)
                        
                    line_x = r.x0 + (bw - line_w) / 2 + b.get("text_dx", 0)
                    page.insert_text(fitz.Point(line_x, line_y), line, fontsize=f_size, fontname=font_name, color=txt_rgb)

        for b in blocks:
            rb = b.get("rel_box")
            if b.get("group_id") is None or not rb:
                continue
            r = fitz.Rect(b["rect"])
            col = tuple(b.get("border_color") or DEFAULT_BORDER_COLOR)
            wid = self._eff_border_width(b, b["id"] in mod_ids)
            dash = {"dashes": "[2 1.5] 0"} if b.get("split_dash") else {}   # 나눈 경계선: 실선 / 점선
            if rb[1] > 0.001:
                page.draw_line(fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y0), color=col, width=wid, **dash)
            if rb[0] > 0.001:
                page.draw_line(fitz.Point(r.x0, r.y0), fitz.Point(r.x0, r.y1), color=col, width=wid, **dash)

        self._draw_header_pdf(page, font_name, measure_font)
        out_bytes = doc.tobytes() if for_preview else doc.tobytes(garbage=3, deflate=True)   # 저장본은 압축해서 용량을 줄인다
        doc.close()
        return out_bytes

    # ================================================================ 화면 좌표 <-> PDF 좌표
    def _to_doc(self, cx, cy):
        ox, oy = self._page_origin()
        return (cx - 18) / self.zoom + ox, (cy - 18) / self.zoom + oy

    def _to_canvas(self, x, y):
        ox, oy = self._page_origin()
        return 18 + (x - ox) * self.zoom, 18 + (y - oy) * self.zoom

    def _block_by_id(self, bid):
        return next((b for b in self.blocks if b["id"] == bid), None)

    def _block_at(self, cx, cy):
        hits = self.canvas.find_overlapping(cx - 1, cy - 1, cx + 1, cy + 1)
        found = next((self.item_map[i] for i in reversed(hits) if i in self.item_map), None)
        return found[0] if found else None

    # ================================================================ 나누기 (2분할, 5분 단위)
    def _ticks(self):
        return self.time_ticks_by_page.get(self.current_page, [])

    def _time_label(self, y):
        """PDF y 위치의 시각 '10:25' (시간 줄을 못 찾았으면 '')."""
        for y0, y1, h in self.time_rows_by_page.get(self.current_page, []):
            if y0 - 0.5 <= y <= y1 + 0.5 and y1 > y0:
                m = int(round((y - y0) / (y1 - y0) * 60))
                hh, mm = (h + m // 60) % 24, m % 60
                return f"{hh:02d}:{mm:02d}"
        return ""

    def _auto_minutes(self, block):
        """블록 위·아래 끝의 시각에서 '분'만 ('20'). 시간 줄을 못 찾으면 None.
        1분 안쪽으로 어긋난 것은 가까운 5분 눈금에 맞춘다 (선 굵기 때문에 19분·21분으로 읽히는 것 막기)."""
        out = []
        for y in (block["rect"][1], block["rect"][3]):
            lab = self._time_label(y)
            if not lab:
                out.append(None)
                continue
            m = int(lab[3:])
            r5 = int(round(m / 5.0) * 5) % 60
            if abs(m - r5) <= 1 or abs(m - r5) >= 59:
                m = r5
            out.append(f"{m:02d}")
        return tuple(out)

    def _fill_auto_times(self, block, top=True, bottom=True):
        """블록 옵션의 위·아래 시간 값을 위치에서 읽은 분으로 채운다. 채운 게 있으면 True."""
        tm, bm = self._auto_minutes(block)
        opts = block.setdefault("options", {})
        done = False
        if top and tm is not None:
            opts["top_time_val"] = tm
            done = True
        if bottom and bm is not None:
            opts["bottom_time_val"] = bm
            done = True
        return done

    def _on_time_check(self, block, which):
        """상단/하단 시간 체크: 켜면 블록 위치에서 분을 자동으로 넣고, 끄면 숨긴다(값은 남겨 둔다)."""
        var = self.opt_top_time_enabled_var if which == "top" else self.opt_bottom_time_enabled_var
        if var.get():
            tm, bm = self._auto_minutes(block)
            val = tm if which == "top" else bm
            if val is not None:
                (self.top_time_val_var if which == "top" else self.bottom_time_val_var).set(val)
                o = block.setdefault("options", {})
                o[f"{which}_time_val"] = val
                o[f"{which}_time_auto"] = True        # 자동 값 (블록 자리가 바뀌면 따라 바뀐다)
        self.on_option_changed(block)

    def set_split_dash(self, block, on=None):
        """나눈 칸(같은 묶음) 사이 경계선을 실선/점선으로."""
        gid = block.get("group_id")
        if gid is None:
            messagebox.showinfo("나눈 경계선", "나눈 칸이 아닙니다. 칸을 먼저 나눠 주세요.")
            return
        group = [b for b in self.blocks if b.get("group_id") == gid]
        on = (not block.get("split_dash")) if on is None else bool(on)
        self.push_history()
        for b in group:
            b["split_dash"] = on
        self.page_blocks[self.current_page] = self.blocks
        self.draw_pdf()
        self.save_autosave()

    def _sync_auto_times(self, block):
        """나누거나 고친 블록: 상단/하단 시간 값을 지금 자리의 분으로 맞춘다.
        - 체크가 꺼져 있으면(안 보임) 늘 맞춘다.
        - 체크가 켜져 있으면 자동으로 넣은 값일 때만 맞춘다 (손으로 고친 값은 그대로)."""
        tm, bm = self._auto_minutes(block)
        o = block.setdefault("options", {})
        for which, val in (("top", tm), ("bottom", bm)):
            if val is None:
                continue
            on = o.get(f"{which}_time_enabled", o.get("time_enabled", False) if which == "top" else False)
            if not on or o.get(f"{which}_time_auto"):
                if o.get(f"{which}_time_val") != val:
                    o[f"{which}_time_val"] = val
                if on:
                    o[f"{which}_time_auto"] = True

    def _keep_auto_flags(self, block, new_opts):
        """옵션을 새로 만들 때 '자동 값' 표시를 이어 붙인다 (값을 손으로 바꾸면 자동이 아님)."""
        prev = block.get("options") or {}
        for which in ("top", "bottom"):
            k = f"{which}_time_auto"
            if k in prev:
                new_opts[k] = bool(prev.get(k)) and new_opts.get(f"{which}_time_val") == prev.get(f"{which}_time_val")
        return new_opts

    def toggle_split_mode(self, on=None):
        if not self.pdf_path:
            return
        on = (not self.split_mode_var.get()) if on is None else bool(on)
        if on and self.manual_draw_mode.get():
            self.manual_draw_mode.set(False)
        self.split_mode_var.set(on)
        self.canvas.delete("splitprev")
        self.canvas.config(cursor="crosshair" if on else "")

    def _on_split_mode_var(self):
        try:
            self.canvas.delete("splitprev")
            self.canvas.config(cursor="crosshair" if self.split_mode_var.get() else "")
        except Exception:
            pass

    def _split_target(self, block, x, y, vertical):
        """나눌 위치 (PDF 좌표, 표시 글). 가로 나누기는 5분 눈금·옆 칸 경계에 붙는다."""
        x0, y0, x1, y1 = block["rect"]
        edge_tol = 5 / self.zoom
        if not vertical:
            lo, hi = y0 + 2.5, y1 - 2.5
            if hi <= lo:
                return None
            edges = [e for b in self.blocks if b is not block      # 옆 요일 칸의 경계에 맞추기
                     for e in (b["rect"][1], b["rect"][3]) if lo <= e <= hi and abs(e - y) <= edge_tol]
            if edges:
                pos = min(edges, key=lambda e: abs(e - y))
                return pos, self._time_label(pos)
            cands = [(t, lab) for t, lab in self._ticks() if lo <= t <= hi]
            if cands:
                return min(cands, key=lambda c: abs(c[0] - y))
            pos = round(min(max(y, lo), hi) * 2) / 2
            return pos, self._time_label(pos)
        lo, hi = x0 + 8, x1 - 8
        if hi <= lo:
            return None
        mid = (x0 + x1) / 2
        if abs(mid - x) <= 10 / self.zoom:
            return mid, "가운데"
        edges = [e for b in self.blocks if b is not block for e in (b["rect"][0], b["rect"][2])
                 if lo <= e <= hi and abs(e - x) <= edge_tol]
        if edges:
            return min(edges, key=lambda e: abs(e - x)), ""
        return round(min(max(x, lo), hi) * 2) / 2, ""

    def _draw_split_label(self, cx, cy, text):
        c = self.canvas
        t = c.create_text(cx + 14, cy - 14, text=text, anchor="w", font=(self.ui_font_family, 9, "bold"), fill="#b91c1c", tags="splitprev")
        bb = c.bbox(t)
        if bb:
            c.create_rectangle(bb[0] - 3, bb[1] - 2, bb[2] + 3, bb[3] + 2, fill="#fff7ed", outline="#f97316", tags="splitprev")
            c.tag_raise(t)

    def on_hover_pdf(self, event):
        """마우스를 움직일 때: 나누기 모드면 자를 선 미리보기, 아니면 나눈 경계선 위에서 커서 모양 바꾸기."""
        if not self.pdf_path or self.drag:
            return
        c = self.canvas
        c.delete("splitprev")
        cx, cy = c.canvasx(event.x), c.canvasy(event.y)
        if self.split_mode_var.get():
            block = self._block_at(cx, cy)
            if not block:
                return
            vertical = bool(event.state & 0x0001)
            tgt = self._split_target(block, *self._to_doc(cx, cy), vertical)
            if not tgt:
                return
            pos, lab = tgt
            x0, y0, x1, y1 = block["rect"]
            if vertical:
                (ax, ay), (bx, by) = self._to_canvas(pos, y0), self._to_canvas(pos, y1)
            else:
                (ax, ay), (bx, by) = self._to_canvas(x0, pos), self._to_canvas(x1, pos)
            c.create_line(ax, ay, bx, by, fill="#dc2626", width=2, dash=(6, 3), tags="splitprev")
            hint = ("세로로 나누기" if vertical else "가로로 나누기") + (f"  {lab}" if lab else "")
            self._draw_split_label(cx, cy, hint + "   (클릭: 나누기 · Shift: 세로 · Esc: 끝)")
            return
        bd = self._boundary_at(*self._to_doc(cx, cy))
        c.config(cursor=("sb_v_double_arrow" if bd[0] == "h" else "sb_h_double_arrow") if bd else "")

    # ================================================================ 작업 스타일 (매일 반복하는 나누기 + 이름 넣기)
    def _text_in_rect(self, rect):
        """이 자리의 원본 PDF 글자 (굵은 글씨 겹침 정리)."""
        x0, y0, x1, y1 = rect
        ws = [w for w in getattr(self, "page_words", {}).get(self.current_page, [])
              if x0 - 0.5 <= (w[0] + w[2]) / 2 <= x1 + 0.5 and y0 - 0.5 <= (w[1] + w[3]) / 2 <= y1 + 0.5]
        ws.sort(key=lambda w: (round(w[1]), w[0]))
        return " ".join(_undouble(w[4]) for w in ws if len(w) > 4)

    @staticmethod
    def _norm_kw(t):
        return re.sub(r"\s+", "", t or "").lower()

    def _minutes_at(self, y):
        lab = self._time_label(y)
        if not lab:
            return None
        return int(lab[:2]) * 60 + int(lab[3:])

    def _y_at_offset(self, y_top, minutes):
        """블록 위 끝(y_top)에서 minutes 분 뒤의 y. 시간 줄을 못 찾으면 None."""
        rows = self.time_rows_by_page.get(self.current_page, [])
        m0 = self._minutes_at(y_top)
        if m0 is None or not rows:
            return None
        # 위 끝이 들어 있는 줄부터 아래로 1시간씩 따라간다
        start = next((i for i, (a, b, _h) in enumerate(rows) if a - 0.5 <= y_top <= b + 0.5), None)
        if start is None:
            return None
        a, b, h = rows[start]
        left = (m0 - h * 60) % 1440 + minutes      # 그 줄의 정각부터 몇 분
        i = start
        while i < len(rows):
            a, b, h = rows[i]
            if left <= 60:
                return a + (b - a) * left / 60.0
            left -= 60
            i += 1
        return None

    def _style_from_selection(self):
        """선택한 칸(나눈 묶음이면 묶음 전체)에서 스타일을 만든다."""
        if len(self.selected) != 1:
            return None, "스타일로 저장할 칸을 1개만 선택해 주세요. (나눈 칸이면 그중 아무 칸이나)"
        b0 = self._block_by_id(next(iter(self.selected)))
        if not b0:
            return None, "선택한 칸을 찾지 못했습니다."
        gid = b0.get("group_id")
        parts = [b for b in self.blocks if gid is not None and b.get("group_id") == gid] or [b0]
        gx0 = min(b["rect"][0] for b in parts); gy0 = min(b["rect"][1] for b in parts)
        gx1 = max(b["rect"][2] for b in parts); gy1 = max(b["rect"][3] for b in parts)
        axis = "h"
        if len(parts) > 1:
            same_x = all(abs(b["rect"][0] - gx0) < 0.6 and abs(b["rect"][2] - gx1) < 0.6 for b in parts)
            same_y = all(abs(b["rect"][1] - gy0) < 0.6 and abs(b["rect"][3] - gy1) < 0.6 for b in parts)
            if same_x:
                axis = "h"
            elif same_y:
                axis = "v"
            else:
                return None, "가로·세로가 섞여 나뉜 칸은 스타일로 저장할 수 없습니다. (한 방향으로만 나눈 칸)"
        parts.sort(key=lambda b: b["rect"][1] if axis == "h" else b["rect"][0])
        cuts = []
        m_top = self._minutes_at(gy0)
        for b in parts[1:]:
            if axis == "h":
                pos = b["rect"][1]
                m = self._minutes_at(pos)
                cut = {"ratio": round((pos - gy0) / max(gy1 - gy0, 1e-6), 4)}
                if m_top is not None and m is not None:
                    cut["min"] = (m - m_top) % 1440       # 시작에서 몇 분 뒤
                    cut["clock"] = m                      # 몇 시 몇 분 (하루 중 분)
            else:
                cut = {"ratio": round((b["rect"][0] - gx0) / max(gx1 - gx0, 1e-6), 4)}
            cuts.append(cut)
        keep = ("title",)                                 # 나누기 + 프로그램 이름만 (시간·뱃지는 사용자가 따로)
        style = {"axis": axis, "cuts": cuts, "parts": [{k: copy.deepcopy(b.get(k)) for k in keep if k in b} for b in parts],
                 "dash": bool(parts[0].get("split_dash")) if len(parts) > 1 else False}
        src = self._text_in_rect((gx0, gy0, gx1, gy1))
        toks = [w for w in src.split() if not re.fullmatch(r"\d{1,2}|UHD|ALL|L|자|수|해|재|\d{1,2}세?", w)]
        style["keyword"] = " ".join(toks).strip() or (parts[0].get("title") or "")
        m_bot = self._minutes_at(gy1)
        if m_top is not None and m_bot is not None:
            style["dur"] = (m_bot - m_top) % 1440 or 1440      # 원래 칸 길이(분) - 길이가 다른 칸에는 적용하지 않으려고
        style["height"] = round((gy1 - gy0) if axis == "h" else (gx1 - gx0), 2)
        style["name"] = " / ".join((p.get("title") or "").replace("\n", " ") for p in style["parts"] if p.get("title")) or style["keyword"]
        return style, ""

    def save_split_style(self):
        style, err = self._style_from_selection()
        if style is None:
            messagebox.showinfo("작업 스타일", err)
            return
        win = tk.Toplevel(self.root)
        win.title("🧩 작업 스타일 저장")
        win.transient(self.root)
        win.grab_set()
        f = ttk.Frame(win, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="스타일 이름", font=(self.ui_font_family, 9, "bold")).grid(row=0, column=0, sticky="w")
        nv = tk.StringVar(value=style["name"][:60])
        ttk.Entry(f, textvariable=nv, width=46).grid(row=0, column=1, sticky="we", pady=2)
        ttk.Label(f, text="찾을 원본 글자", font=(self.ui_font_family, 9, "bold")).grid(row=1, column=0, sticky="w")
        kv = tk.StringVar(value=style["keyword"][:60])
        ttk.Entry(f, textvariable=kv, width=46).grid(row=1, column=1, sticky="we", pady=2)
        ttk.Label(f, text="원본 편성표 칸에 이 글자가 들어 있으면 적용합니다 (띄어쓰기는 무시). 짧고 확실한 말로 줄여 주세요.\n예: '뉴스투데이 2부', '뉴스데스크'",
                  font=(self.ui_font_family, 8), foreground="#64748b", justify="left").grid(row=2, column=1, sticky="w")
        if style["cuts"]:
            how = ", ".join(f"{round(c['ratio'] * 100)}% 자리" for c in style["cuts"])
            desc = f"{'위아래' if style['axis'] == 'h' else '좌우'} {len(style['parts'])}칸으로 나눔 ({how})"
        else:
            desc = "나누지 않음 (이름만)"
        lines = [desc] + [f"  #{i + 1} {(p.get('title') or '(이름 없음)').replace(chr(10), ' ')}" for i, p in enumerate(style["parts"])]
        ttk.Label(f, text="\n".join(lines), justify="left").grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 4))

        def ok(_e=None):
            name, kw = nv.get().strip(), kv.get().strip()
            if not name or not self._norm_kw(kw):
                messagebox.showwarning("작업 스타일", "이름과 찾을 원본 글자를 넣어 주세요.", parent=win)
                return
            style["name"], style["keyword"] = name, kw
            old = next((i for i, s_ in enumerate(self.split_styles) if s_.get("name") == name), None)
            if old is not None:
                if not messagebox.askyesno("작업 스타일", f"'{name}' 스타일이 이미 있습니다. 바꿀까요?", parent=win):
                    return
                self.split_styles[old] = style
            else:
                self.split_styles.append(style)
            self.save_autosave()
            self.refresh_style_ui()
            win.destroy()

        b = ttk.Frame(f)
        b.grid(row=4, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(b, text="저장", command=ok).pack(side="left", padx=4)
        ttk.Button(b, text="취소", command=win.destroy).pack(side="left", padx=4)
        win.bind("<Return>", ok)

    def _program_text(self, rect):
        """칸 안 원본 글자에서 분 숫자·뱃지 글자를 뺀 것 (띄어쓰기 없이)."""
        toks = [w for w in self._text_in_rect(rect).split() if not re.fullmatch(r"\d{1,2}|UHD|ALL|L|자|수|해|재|\d{1,2}세?", w)]
        return self._norm_kw(" ".join(toks))

    def _style_fits(self, b, style):
        """이 칸에 스타일을 걸어도 되는지. (맞음, 이유)"""
        kw = self._norm_kw(style.get("keyword"))
        if not kw:
            return False, "찾을 글자 없음"
        txt = self._program_text(b["rect"])
        if kw not in txt:
            return False, "글자 다름"
        if len(txt) > len(kw) * 2 + 6:                   # 칸 안에 다른 프로그램 글자가 많이 섞여 있음 (여러 칸이 한 칸으로 잡힌 경우)
            return False, "다른 프로그램이 섞인 칸"
        return True, ""

    def _style_targets(self, style, skipped=None):
        out = []
        for b in self.blocks:
            if b.get("group_id") is not None or self._orig_match(b) is None:
                continue                                  # 이미 나눈 칸·옮긴 칸은 건드리지 않는다
            ok, why = self._style_fits(b, style)
            if ok:
                out.append(b)
            elif skipped is not None and why != "글자 다름":
                skipped.append((b, style, why))
        return out

    def _block_when(self, b):
        """'10/24(토) 20:00~20:40' 같은 칸 설명."""
        t0, t1 = self._time_label(b["rect"][1]), self._time_label(b["rect"][3])
        day = ""
        for d in (self.header_found_by_page.get(self.current_page) or {}).get("days") or []:
            if abs(d["cx"] - (b["rect"][0] + b["rect"][2]) / 2) <= d.get("cell_w", 60) / 2:
                day = d["text"] + " "
                break
        return f"{day}{t0}~{t1}" if t0 else day.strip() or "칸"

    def _apply_style_to(self, block, style):
        """한 칸에 스타일 적용. 성공하면 새 칸 목록, 못 하면 None."""
        x0, y0, x1, y1 = block["rect"]
        axis, cuts, parts = style.get("axis", "h"), style.get("cuts") or [], style.get("parts") or []
        if not parts or len(parts) != len(cuts) + 1:
            return None
        if axis == "h":
            lo, hi = y0, y1
        else:
            lo, hi = x0, x1
        pos = [lo + (hi - lo) * float(c.get("ratio", 0.5)) for c in cuts]   # 같은 비율로 나눈다 (시간 계산 없음)
        bounds = [lo] + pos + [hi]
        if any(bounds[i + 1] - bounds[i] < 2.5 for i in range(len(bounds) - 1)):
            return None                                   # 이 칸은 너무 짧아 같은 자리에서 나눌 수 없음
        gid = None
        if len(parts) > 1:
            gid = self.next_id
            self.next_id += 1
        new = []
        for i, sp in enumerate(parts):
            nb = copy.deepcopy(block)
            nb["id"] = self.next_id
            self.next_id += 1
            if len(parts) > 1:
                r = (x0, bounds[i], x1, bounds[i + 1]) if axis == "h" else (bounds[i], y0, bounds[i + 1], y1)
                nb["rect"] = nb["orig_rect"] = r
                nb["group_id"] = gid
            nb["title"] = sp.get("title", "")             # 프로그램 이름만 넣는다 (시간·뱃지 등은 사용자가 따로)
            if len(parts) > 1:
                nb["split_dash"] = bool(style.get("dash"))    # 나눈 경계선 실선/점선도 스타일대로
            new.append(nb)
        i = self.blocks.index(block)
        self.blocks[i:i + 1] = new
        if gid is not None:
            self._recompute_rel_boxes(gid)
        return new

    def apply_split_styles(self, styles=None, quiet=False):
        """스타일을 이 페이지의 맞는 칸 모두에 적용한다. styles=None 이면 저장된 스타일 전부."""
        styles = list(self.split_styles) if styles is None else styles
        if not self.pdf_path or not styles:
            return 0
        skipped = []
        plan = [(st, b) for st in styles for b in self._style_targets(st, skipped)]
        seen, todo = set(), []
        for st, b in plan:                                # 한 칸에 스타일 두 개가 걸리면 먼저 저장한 것만
            if b["id"] not in seen:
                seen.add(b["id"])
                todo.append((st, b))
        skip_txt = ""
        if skipped:
            skip_txt = "\n\n건너뛴 칸 (글자는 맞지만):\n" + "\n".join(f"  · {st['name']} — {self._block_when(b)}: {why}" for b, st, why in skipped[:12])
        if not todo:
            if not quiet:
                messagebox.showinfo("작업 스타일", "이 페이지에서 적용할 칸을 찾지 못했습니다.\n(이미 나눈 칸은 건너뜁니다. '찾을 원본 글자'를 확인해 주세요.)" + skip_txt)
            return 0
        if not quiet:
            lst = "\n".join(f"  · {st['name']} — {self._block_when(b)}" for st, b in todo[:20]) + ("\n  …" if len(todo) > 20 else "")
            if not messagebox.askyesno("작업 스타일 적용", f"아래 {len(todo)}칸에 적용합니다.\n\n{lst}{skip_txt}\n\n계속할까요?"):
                return 0
        self.push_history()
        done, skip, names = 0, 0, {}
        for st, b in todo:
            if self._apply_style_to(b, st):
                done += 1
                names[st["name"]] = names.get(st["name"], 0) + 1
            else:
                skip += 1
        self.page_blocks[self.current_page] = self.blocks
        self.selected = set()
        self.draw_pdf()
        self.save_autosave()
        if not quiet or done:
            msg = "\n".join(f"· {n}: {c}칸" for n, c in names.items()) or "적용한 칸 없음"
            if quiet and skipped:
                msg += skip_txt
            if skip:
                msg += f"\n\n칸이 짧아 나누지 못한 {skip}칸은 건너뛰었습니다."
            messagebox.showinfo("작업 스타일 적용", msg + "\n\nCtrl+Z 로 한 번에 되돌릴 수 있습니다.")
        return done

    def delete_split_style(self, idx):
        if 0 <= idx < len(self.split_styles):
            name = self.split_styles[idx].get("name", "")
            if messagebox.askyesno("작업 스타일", f"'{name}' 스타일을 지울까요?"):
                self.split_styles.pop(idx)
                self.save_autosave()
                self.refresh_style_ui()

    def _build_style_panel(self):
        f = ttk.LabelFrame(self.sidebar_content_frame, text=" 🧩 작업 스타일 (매일 반복: 나누기 + 이름) ", padding=8)
        f.pack(fill="x", pady=4)
        self.style_container = ttk.Frame(f)
        self.style_container.pack(fill="x")
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(4, 0))
        ttk.Button(b, text="▶ 모두 적용", command=lambda: self.apply_split_styles()).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(b, text="＋ 선택한 칸을 스타일로 저장", command=self.save_split_style).pack(side="left", fill="x", expand=True, padx=(2, 0))
        av = tk.BooleanVar(value=self.split_styles_auto)

        def set_auto():
            self.split_styles_auto = bool(av.get())
            self.save_autosave()
        ttk.Checkbutton(f, text="PDF를 열 때 자동으로 모두 적용", variable=av, command=set_auto).pack(anchor="w", pady=(4, 0))
        self.refresh_style_ui()

    def refresh_style_ui(self):
        box = getattr(self, "style_container", None)
        if box is None or not box.winfo_exists():
            return
        for w in box.winfo_children():
            w.destroy()
        if not self.split_styles:
            ttk.Label(box, text="저장된 스타일이 없습니다. 한 칸을 나누고 이름을 넣은 뒤\n그 칸을 선택하고 [＋ 선택한 칸을 스타일로 저장]을 누르세요.",
                      font=(self.ui_font_family, 8), foreground="#64748b", justify="left").pack(anchor="w")
            return
        for i, st in enumerate(self.split_styles):
            row = tk.Frame(box, bg="#f8fafc", bd=1, relief="groove")
            row.pack(fill="x", pady=1)
            n = len(self._style_targets(st)) if self.pdf_path else 0
            tk.Label(row, text=f"🧩 {st.get('name', '')}", bg="#f8fafc", font=(self.ui_font_family, 8, "bold"),
                     anchor="w").pack(side="left", padx=4)
            tk.Button(row, text="X", bg="#fee2e2", fg="#dc2626", bd=0, font=(self.ui_font_family, 7, "bold"),
                      command=lambda k=i: self.delete_split_style(k)).pack(side="right", padx=2)
            tk.Button(row, text=f"적용 ({n})", bg="#dbeafe", bd=1, font=(self.ui_font_family, 7),
                      state="normal" if n else "disabled",
                      command=lambda s_=st: self.apply_split_styles([s_])).pack(side="right", padx=2)
            tk.Label(row, text=f"찾기: {st.get('keyword', '')}", bg="#f8fafc", fg="#64748b",
                     font=(self.ui_font_family, 7)).pack(side="right", padx=4)

    def _recompute_rel_boxes(self, gid):
        gbs = [b for b in self.blocks if b.get("group_id") == gid]
        if not gbs:
            return
        gx0 = min(b["rect"][0] for b in gbs); gy0 = min(b["rect"][1] for b in gbs)
        gx1 = max(b["rect"][2] for b in gbs); gy1 = max(b["rect"][3] for b in gbs)
        gw, gh = max(gx1 - gx0, 1e-6), max(gy1 - gy0, 1e-6)
        for b in gbs:
            x0, y0, x1, y1 = b["rect"]
            b["rel_box"] = ((x0 - gx0) / gw, (y0 - gy0) / gh, (x1 - gx0) / gw, (y1 - gy0) / gh)

    def _do_split(self, block, pos, vertical):
        """블록을 pos 에서 둘로 나눈다. 위(왼쪽)는 원래 이름, 아래(오른쪽)는 새 칸 -> 이름 입력 창."""
        self.push_history()
        x0, y0, x1, y1 = block["rect"]
        gid = block.get("group_id")
        if gid is None:
            gid = self.next_id
            self.next_id += 1
        a, b = copy.deepcopy(block), copy.deepcopy(block)
        a["id"] = self.next_id
        b["id"] = self.next_id + 1
        self.next_id += 2
        if vertical:
            a["rect"], b["rect"] = (x0, y0, pos, y1), (pos, y0, x1, y1)
        else:
            a["rect"], b["rect"] = (x0, y0, x1, pos), (x0, pos, x1, y1)
        a["orig_rect"], b["orig_rect"] = a["rect"], b["rect"]
        if not (a.get("title") or "").strip():     # 나누면 원본 글자를 덮으므로, 위(왼쪽) 칸에는 원본 이름을 넣어 둔다
            a["title"] = self._source_title(block)
        b["title"] = ""
        if True:                                   # 원본의 분 숫자도 덮이므로, 정시가 아니면 위 시간(분)을 자동으로 켜 둔다
            for part in (a, b):
                tm, _bm = self._auto_minutes(part)
                if tm is not None and tm != "00":
                    part.setdefault("options", {})["top_time_enabled"] = True
                    part["options"]["top_time_val"] = tm
                    part["options"]["top_time_auto"] = True
                self._sync_auto_times(part)
        a["group_id"] = b["group_id"] = gid
        idx = self.blocks.index(block)
        self.blocks[idx:idx + 1] = [a, b]
        self._recompute_rel_boxes(gid)
        self.page_blocks[self.current_page] = self.blocks
        self.selected = {b["id"]}
        self.toggle_split_mode(False)          # 2분할: 한 번 나누면 나누기 모드 끝
        self._last_split_at = time.time()
        self.draw_pdf()
        self.save_autosave()
        self.root.after(60, lambda: self.open_title_editor(b, hint="나눈 아래(오른쪽) 칸의 이름"))

    def _source_title(self, block):
        """블록 자리의 원본 글자에서 이름만 (분 숫자·뱃지 글자는 뺀다)."""
        t = block.get("src_text")
        if t is None:
            ob = self._orig_match(block)
            t = (ob or {}).get("src_text") or ""
        toks = [w for w in str(t).split() if not re.fullmatch(r"\d{1,2}|UHD|ALL|L|자|수|해|\d{1,2}세?", w)]
        return " ".join(toks).strip()

    def _boundary_at(self, x, y):
        """나눈 두 칸(같은 묶음)이 맞닿은 경계선 근처면 ('h'|'v', 위/왼쪽 칸, 아래/오른쪽 칸)."""
        tol = 4 / self.zoom
        groups = {}
        for b in self.blocks:
            if b.get("group_id") is not None:
                groups.setdefault(b["group_id"], []).append(b)
        for gbs in groups.values():
            for a in gbs:
                ax0, ay0, ax1, ay1 = a["rect"]
                for b in gbs:
                    if a is b:
                        continue
                    bx0, by0, bx1, by1 = b["rect"]
                    if abs(ay1 - by0) < 0.6:
                        o0, o1 = max(ax0, bx0), min(ax1, bx1)
                        if o1 > o0 and o0 <= x <= o1 and abs(y - ay1) <= tol:
                            return ("h", a, b)
                    if abs(ax1 - bx0) < 0.6:
                        o0, o1 = max(ay0, by0), min(ay1, by1)
                        if o1 > o0 and o0 <= y <= o1 and abs(x - ax1) <= tol:
                            return ("v", a, b)
        return None

    def _drag_boundary(self, cx, cy):
        a, b = self._block_by_id(self.drag["a"]), self._block_by_id(self.drag["b"])
        if not a or not b:
            return
        X, Y = self._to_doc(cx, cy)
        if self.drag["axis"] == "h":
            lo, hi = a["rect"][1] + 2.5, b["rect"][3] - 2.5
            cands = [t for t, _ in self._ticks() if lo <= t <= hi]
            pos = min(cands, key=lambda t: abs(t - Y)) if cands else round(min(max(Y, lo), hi) * 2) / 2
            a["rect"] = (a["rect"][0], a["rect"][1], a["rect"][2], pos)
            b["rect"] = (b["rect"][0], pos, b["rect"][2], b["rect"][3])
            lab = self._time_label(pos)
        else:
            lo, hi = a["rect"][0] + 8, b["rect"][2] - 8
            mid = (a["rect"][0] + b["rect"][2]) / 2
            pos = mid if abs(X - mid) <= 8 / self.zoom else round(min(max(X, lo), hi) * 2) / 2
            a["rect"] = (a["rect"][0], a["rect"][1], pos, a["rect"][3])
            b["rect"] = (pos, b["rect"][1], b["rect"][2], b["rect"][3])
            lab = "가운데" if pos == mid else ""
        a["orig_rect"], b["orig_rect"] = a["rect"], b["rect"]
        self._recompute_rel_boxes(a.get("group_id"))
        self.drag["moved"] = True
        self._draw_soon((lambda: self._draw_split_label(cx, cy, lab)) if lab else None)

    # ================================================================ 칸 이름 입력 창 (더블클릭)
    def _remember_title(self, text):
        t = (text or "").strip()
        if not t:
            return
        self.recent_titles = [t] + [x for x in self.recent_titles if x != t]
        del self.recent_titles[20:]

    def open_title_editor(self, block, hint="칸 이름"):
        """칸 위에 작은 창: 이름을 바로 고치거나, 저장해 둔 제목·최근 제목에서 골라 넣는다.
        Enter: 확인 / Shift+Enter: 줄바꿈 / Esc: 취소"""
        old = getattr(self, "_title_win", None)
        if old is not None:
            try:
                old.destroy()
            except tk.TclError:
                pass
        win = tk.Toplevel(self.root)
        self._title_win = win
        win.title("칸 이름")
        win.transient(self.root)
        win.resizable(False, False)
        try:
            cx, cy = self._to_canvas(block["rect"][2], block["rect"][1])
            px = self.canvas.winfo_rootx() + int(cx - self.canvas.canvasx(0)) + 8
            py = self.canvas.winfo_rooty() + int(cy - self.canvas.canvasy(0))
            win.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass
        ttk.Label(win, text=hint, font=(self.ui_font_family, 9, "bold")).pack(anchor="w", padx=8, pady=(8, 2))
        txt = tk.Text(win, width=30, height=3, font=(self.ui_font_family, 10), wrap="word")
        txt.pack(padx=8)
        txt.insert("1.0", block.get("title", ""))
        txt.tag_add("sel", "1.0", "end-1c")
        txt.focus_set()
        ttk.Label(win, text="저장한 제목 · 파렛트 · 최근 (클릭: 넣기, 더블클릭: 넣고 확인)", font=(self.ui_font_family, 8), foreground="#555").pack(anchor="w", padx=8, pady=(6, 1))
        lf = ttk.Frame(win)
        lf.pack(fill="both", padx=8)
        lb = tk.Listbox(lf, height=9, width=34, font=(self.ui_font_family, 9), activestyle="none", exportselection=False)
        sb = ttk.Scrollbar(lf, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both")
        sb.pack(side="right", fill="y")
        items = []

        def fill_list():
            lb.delete(0, "end")
            items.clear()
            seen = set()
            for kind, mark, src in (("saved", "★ ", self.title_presets),
                                    ("palette", "◆ ", [p.get("title", "") for p in self.favorite_presets]),
                                    ("recent", "· ", self.recent_titles)):
                for t in src:
                    if t and t not in seen:
                        seen.add(t)
                        items.append((kind, t))
                        lb.insert("end", mark + t.replace("\n", " / "))

        fill_list()

        def put(event=None):
            sel = lb.curselection()
            if sel:
                txt.delete("1.0", "end")
                txt.insert("1.0", items[sel[0]][1])

        def ok(event=None):
            if event is not None and getattr(event, "state", 0) & 0x0001:
                return None                         # Shift+Enter: 줄바꿈
            new = txt.get("1.0", "end-1c").strip()
            if new != block.get("title", ""):
                self.push_history()
                block["title"] = new
                self.page_blocks[self.current_page] = self.blocks
            self._remember_title(new)
            win.destroy()
            self.selected = {block["id"]}
            self.draw_pdf()
            self.save_autosave()
            return "break"

        def save_preset():
            t = txt.get("1.0", "end-1c").strip()
            if t and t not in self.title_presets:
                self.title_presets.insert(0, t)
                self.save_autosave()
                fill_list()

        def del_preset():
            sel = lb.curselection()
            if sel and items[sel[0]][0] == "saved":
                self.title_presets.remove(items[sel[0]][1])
                self.save_autosave()
                fill_list()

        lb.bind("<<ListboxSelect>>", put)
        lb.bind("<Double-Button-1>", lambda e: (put(), ok()))
        txt.bind("<Return>", ok)
        win.bind("<Escape>", lambda e: win.destroy())
        bf = ttk.Frame(win)
        bf.pack(fill="x", padx=8, pady=8)
        ttk.Button(bf, text="확인 (Enter)", command=ok).pack(side="left")
        ttk.Button(bf, text="취소", command=win.destroy).pack(side="left", padx=4)
        ttk.Button(bf, text="★ 제목 저장", command=save_preset).pack(side="right")
        ttk.Button(bf, text="저장 삭제", command=del_preset).pack(side="right", padx=4)

    # ================================================================ 머리글 (제목 · 날짜) 바꾸기
    def _header_state(self):
        st = self.header_edit.setdefault(self.current_page, {})
        found = self.header_found_by_page.get(self.current_page) or {}
        st.setdefault("title_on", bool(IS_DMB and found.get("title")))   # DMB 는 제목을 늘 바꾼다
        st.setdefault("title", self._suggest_header_title())
        st.setdefault("title_size", 0)            # 0 = 자동
        st.setdefault("title_bold", True)
        st.setdefault("date_bold", False)
        st.setdefault("date_on", False)
        st.setdefault("date", ((found.get("date") or {}).get("text") or "").strip())
        days = found.get("days") or []
        cur = st.get("days")
        if not isinstance(cur, list) or len(cur) != len(days):
            prev = self.dmb_day_regions
            st["days"] = [{"date": d["text"], "region": (prev[i] if i < len(prev) else "")} for i, d in enumerate(days)]
        st.setdefault("days_on", bool(IS_DMB and any(v.get("region") for v in st["days"])))
        return st

    def _suggest_header_title(self):
        """파일 이름 'MBC_10__3_2' -> '광주MBC 주간편성표 [ 10월 03주 확정 ]'.
        DMB: 'MBC_11__1__2' -> '광주 MYMBC 주간편성표 [ 11월 01주 2 ]' (뒤 번호가 없으면 '기본')"""
        name = os.path.basename(self.pdf_path or "")
        if IS_DMB:
            m = re.search(r"_(\d{1,2})__(\d)(?:_+([0-9A-Za-z가-힣]+))?", name)
            if m:
                rev = m.group(3) or "기본"
                return f"광주 MYMBC 주간편성표 [ {int(m.group(1))}월 {int(m.group(2)):02d}주 {rev} ]"
            return self.header_title_presets[0] if self.header_title_presets else "광주 MYMBC 주간편성표"
        m = re.search(r"_(\d{1,2})__(\d)", name)
        if m:
            return f"광주MBC 주간편성표 [ {int(m.group(1))}월 {int(m.group(2)):02d}주 확정 ]"
        return self.header_title_presets[0] if self.header_title_presets else "광주MBC 주간편성표"

    def _now_date_text(self, like=""):
        """원래 날짜와 같은 모양으로 지금 시각. (AM/PM 이 있으면 12시간제)"""
        now = datetime.datetime.now()
        wd = "월화수목금토일"[now.weekday()]
        if re.search(r"\b(AM|PM)\b", like or ""):
            h = now.hour % 12 or 12
            return f"{now.year}년 {now.month}월 {now.day}일 ({wd}) {'AM' if now.hour < 12 else 'PM'} {h:02d}:{now.minute:02d}"
        return f"{now.year}년 {now.month}월 {now.day}일({wd}) {now.hour:02d}:{now.minute:02d}"

    def _header_layout(self, measure):
        """머리글을 그릴 자리. measure(text, size) -> 너비(pt).
        반환: [(덮을 사각형, 글자, 크기, (x, 기준선 y), 굵게)]"""
        out = []
        found = self.header_found_by_page.get(self.current_page) or {}
        st = self._header_state()
        pr = self.page_rect or (0, 0, 595, 842)
        tf = found.get("title")
        if st["title_on"] and tf and st["title"].strip():
            x0, y0, x1, y1 = tf["rect"]
            text = st["title"].strip()
            size = float(st["title_size"] or 0) or min((y1 - y0) * 0.55, 20.0)
            avail = (pr[2] - pr[0]) - 60
            while size > 6 and measure(text, size) > avail:
                size -= 0.5
            w = measure(text, size)
            cxm = (x0 + x1) / 2
            tx = min(max(cxm - w / 2, pr[0] + 30), pr[2] - 30 - w)
            out.append(((x0 - 1, y0 - 1, x1 + 1, y1 + 1), text, size, (tx, (y0 + y1) / 2 + size * 0.36), bool(st["title_bold"]), "#ffffff"))
        df = found.get("date")
        if st["date_on"] and df and st["date"].strip():
            x0, y0, x1, y1 = df["rect"]
            text = st["date"].strip()
            size = float(df.get("size") or (y1 - y0) * 0.8)
            w = measure(text, size)
            if st.get("date_bold"):
                w *= 1.06                          # 굵은 글자는 조금 넓다 (원래 글자를 다 덮고 오른쪽 끝을 맞추려고)
            out.append(((min(x0, x1 - w) - 1, y0 - 1, x1 + 1, y1 + 1), text, size, (x1 - w, (y0 + y1) / 2 + size * 0.36), bool(st.get("date_bold")), "#ffffff"))
        days = found.get("days") or []
        if st.get("days_on") and days:
            for d, v in zip(days, st["days"]):
                text = (str(v.get("date", "")).strip() + str(v.get("region", "")).strip())
                if not text or text == d["text"]:
                    continue
                x0, y0, x1, y1 = d["rect"]
                size = float(d.get("size") or (y1 - y0) * 0.8)
                avail = d["cell_w"] - 4
                while size > 4 and measure(text, size) > avail:
                    size -= 0.25
                w = measure(text, size)
                half = max(d["cell_w"] / 2 - 1.5, (x1 - x0) / 2 + 0.5)
                cover = (d["cx"] - half, y0 - 0.6, d["cx"] + half, y1 + 0.6)
                out.append((cover, text, size, (d["cx"] - w / 2, d["cy"] + size * 0.36), False, d.get("fill") or "#B8CCE4"))
        return out

    def _paint_header_on_canvas(self, c):
        """기본 화면에 머리글 바꾼 내용을 바로 보여 준다. 실패하면 한 번만 이유를 알려 준다."""
        c.delete("hdrprev")
        try:
            self._draw_header_preview(c)
            c.tag_raise("hdrprev")
        except Exception as exc:
            import traceback
            traceback.print_exc()
            if not getattr(self, "_hdr_err_shown", False):
                self._hdr_err_shown = True
                try:
                    messagebox.showwarning("머리글 표시", f"기본 화면에 머리글을 그리지 못했습니다.\n(저장 PDF·미리보기에는 그대로 들어갑니다)\n\n{type(exc).__name__}: {exc}")
                except Exception:
                    pass

    def _draw_header_preview(self, c):
        try:
            from tkinter import font as tkfont
        except Exception:          # 예전 런처(실행 환경)에 tkinter.font 가 빠져 있으면 글자 폭을 어림으로 잰다
            tkfont = None
        cache = {}

        def measure(text, size):
            if tkfont is None:
                return sum(size if ord(ch) > 127 else size * 0.55 for ch in text)
            px = max(4, int(round(size * self.zoom)))
            f = cache.get(px) or tkfont.Font(family=self.ui_font_family, size=-px)
            cache[px] = f
            return f.measure(text) / self.zoom

        for cover, text, size, (tx, ty), bold, bg in self._header_layout(measure):
            a = self._to_canvas(cover[0], cover[1])
            b = self._to_canvas(cover[2], cover[3])
            c.create_rectangle(a[0], a[1], b[0], b[1], fill=bg, outline="", tags=("hdrprev",))
            px = max(4, int(round(size * self.zoom)))
            x, y = self._to_canvas(tx, ty)
            c.create_text(x, y, text=text, anchor="sw", fill="#111111", tags=("hdrprev",),
                          font=(self.ui_font_family, -px, "bold") if bold else (self.ui_font_family, -px))

    def _draw_header_pdf(self, page, font_name, measure_font):
        def measure(text, size):
            try:
                return measure_font.text_length(text, fontsize=size)
            except Exception:
                return sum(size if ord(ch) > 127 else size * 0.5 for ch in text)

        for cover, text, size, (tx, ty), bold, bg in self._header_layout(measure):
            bg = bg.lstrip("#")
            page.draw_rect(fitz.Rect(cover), color=None, fill=tuple(int(bg[i:i + 2], 16) / 255 for i in (0, 2, 4)))
            kw = {"render_mode": 2, "border_width": 0.03} if bold else {}
            try:
                page.insert_text(fitz.Point(tx, ty), text, fontsize=size, fontname=font_name, color=(0, 0, 0), fill=(0, 0, 0), **kw)
            except TypeError:
                page.insert_text(fitz.Point(tx, ty), text, fontsize=size, fontname=font_name, color=(0, 0, 0))

    def _build_header_panel(self):
        f = ttk.LabelFrame(self.sidebar_content_frame, text=" 📝 머리글 (제목 · 날짜) ", padding=8)
        f.pack(fill="x", pady=4)
        found = self.header_found_by_page.get(self.current_page) or {}
        st = self._header_state()
        fnt = (self.ui_font_family, 8, "bold")

        def changed(*_a):
            st["title_on"] = bool(t_on.get())
            st["title"] = t_var.get()
            st["date_on"] = bool(d_on.get())
            st["date"] = d_var.get()
            st["title_bold"] = bool(t_bold.get())
            st["date_bold"] = bool(d_bold.get())
            self.draw_pdf(refresh_sidebar=False)

        t_on = tk.BooleanVar(value=st["title_on"])
        t_var = tk.StringVar(value=st["title"])
        t_bold = tk.BooleanVar(value=st["title_bold"])
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Checkbutton(row, text="제목 바꾸기", variable=t_on, command=changed).pack(side="left")
        ttk.Checkbutton(row, text="굵게", variable=t_bold, command=changed).pack(side="left", padx=6)
        if not found.get("title"):
            ttk.Label(row, text="(제목 자리를 찾지 못함)", foreground="#b45309", font=(self.ui_font_family, 8)).pack(side="left")
        values = [self._suggest_header_title()] + [x for x in self.header_title_presets if x != self._suggest_header_title()]
        cb = ttk.Combobox(f, textvariable=t_var, values=values, font=(self.ui_font_family, 9))
        cb.pack(fill="x", pady=2)
        cb.bind("<<ComboboxSelected>>", lambda e: (t_on.set(True), changed()))
        t_var.trace_add("write", lambda *a: (t_on.set(True) if t_var.get().strip() else None, changed()))
        r2 = ttk.Frame(f)
        r2.pack(fill="x")

        def size(d):
            cur = float(st["title_size"] or 0) or 20.0
            st["title_size"] = max(6.0, min(48.0, cur + d))
            sz_lbl.config(text=f"{st['title_size']:.0f}pt")
            self.draw_pdf(refresh_sidebar=False)

        def auto_size():
            st["title_size"] = 0
            sz_lbl.config(text="자동")
            self.draw_pdf(refresh_sidebar=False)

        def save_title():
            t = t_var.get().strip()
            if t and t not in self.header_title_presets:
                self.header_title_presets.insert(0, t)
                cb.config(values=[self._suggest_header_title()] + [x for x in self.header_title_presets if x != self._suggest_header_title()])
                self.save_autosave()

        def del_title():
            t = t_var.get().strip()
            if t in self.header_title_presets:
                self.header_title_presets.remove(t)
                cb.config(values=[self._suggest_header_title()] + self.header_title_presets)
                self.save_autosave()

        ttk.Label(r2, text="크기:", font=fnt).pack(side="left")
        ttk.Button(r2, text="-", width=2, command=lambda: size(-1)).pack(side="left")
        sz_lbl = ttk.Label(r2, text=(f"{st['title_size']:.0f}pt" if st["title_size"] else "자동"), width=5, anchor="center")
        sz_lbl.pack(side="left")
        ttk.Button(r2, text="+", width=2, command=lambda: size(1)).pack(side="left")
        ttk.Button(r2, text="자동", width=4, command=auto_size).pack(side="left", padx=2)
        ttk.Button(r2, text="삭제", width=4, command=del_title).pack(side="right")
        ttk.Button(r2, text="★ 저장", width=6, command=save_title).pack(side="right", padx=2)

        ttk.Separator(f).pack(fill="x", pady=5)
        d_on = tk.BooleanVar(value=st["date_on"])
        d_var = tk.StringVar(value=st["date"])
        r3 = ttk.Frame(f)
        r3.pack(fill="x")
        d_bold = tk.BooleanVar(value=bool(st.get("date_bold")))
        ttk.Checkbutton(r3, text="날짜 바꾸기", variable=d_on, command=changed).pack(side="left")
        ttk.Checkbutton(r3, text="굵게", variable=d_bold, command=lambda: (d_on.set(True), changed())).pack(side="left", padx=6)
        if not found.get("date"):
            ttk.Label(r3, text="(날짜 자리를 찾지 못함)", foreground="#b45309", font=(self.ui_font_family, 8)).pack(side="left")
        orig = ((found.get("date") or {}).get("text") or "").strip()
        ttk.Button(r3, text="지금 시각", command=lambda: (d_var.set(self._now_date_text(orig)), d_on.set(True), changed())).pack(side="right")
        ttk.Button(r3, text="원래", command=lambda: (d_var.set(orig), changed())).pack(side="right", padx=2)
        de = ttk.Entry(f, textvariable=d_var, font=(self.ui_font_family, 9))
        de.pack(fill="x", pady=2)
        d_var.trace_add("write", lambda *a: (d_on.set(True) if d_var.get().strip() != orig else None, changed()))

    def _build_days_panel(self):
        """DMB: 표 머리 요일 칸 '11/2(월)' -> '11/2(월)전주' (날짜는 고칠 수 있고, 지역은 목록에서 고른다)."""
        f = ttk.LabelFrame(self.sidebar_content_frame, text=" 📅 요일 칸 (날짜 · 지역) ", padding=8)
        f.pack(fill="x", pady=4)
        found = self.header_found_by_page.get(self.current_page) or {}
        days = found.get("days") or []
        st = self._header_state()
        small = (self.ui_font_family, 8)
        on = tk.BooleanVar(value=bool(st.get("days_on")))
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Checkbutton(top, text="요일 칸 바꾸기", variable=on, command=lambda: changed()).pack(side="left")
        if not days:
            ttk.Label(top, text="(요일 칸을 찾지 못함)", foreground="#b45309", font=small).pack(side="left")
            return
        if not self.dmb_regions:
            self.dmb_regions = list(DMB_REGIONS)
        rows = []
        busy = [True]

        def changed(*_a):
            if busy[0]:
                return
            st["days"] = [{"date": dv.get(), "region": rv.get().strip()} for dv, rv, _cb in rows]
            st["days_on"] = bool(on.get())
            regions = [d["region"] for d in st["days"]]
            if any(regions):
                self.dmb_day_regions = regions
            self.draw_pdf(refresh_sidebar=False)
            self._schedule_autosave()

        def turn_on(*_a):
            if not busy[0] and not on.get():
                on.set(True)
            changed()

        def remember_region(rv):
            r = rv.get().strip()
            if r and r not in self.dmb_regions:
                self.dmb_regions.append(r)
                for _dv, _rv, cb in rows:
                    cb.config(values=self.dmb_regions)
                self._schedule_autosave()

        grid = ttk.Frame(f)
        grid.pack(fill="x", pady=(3, 0))
        for i, (d, v) in enumerate(zip(days, st["days"])):
            dv = tk.StringVar(value=v.get("date", d["text"]))
            rv = tk.StringVar(value=v.get("region", ""))
            ttk.Entry(grid, textvariable=dv, width=10, font=(self.ui_font_family, 9)).grid(row=i, column=0, sticky="we", pady=1)
            cb = ttk.Combobox(grid, textvariable=rv, values=self.dmb_regions, width=6, font=(self.ui_font_family, 9))
            cb.grid(row=i, column=1, sticky="we", padx=(3, 0), pady=1)
            cb.bind("<<ComboboxSelected>>", lambda e, _rv=rv: (remember_region(_rv), turn_on()))
            cb.bind("<FocusOut>", lambda e, _rv=rv: remember_region(_rv))
            cb.bind("<Return>", lambda e, _rv=rv: remember_region(_rv))
            dv.trace_add("write", turn_on)
            rv.trace_add("write", turn_on)
            rows.append((dv, rv, cb))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        def set_all(dates=None, regions=None):
            busy[0] = True
            for i, (dv, rv, _cb) in enumerate(rows):
                if dates is not None:
                    dv.set(dates[i] if i < len(dates) else "")
                if regions is not None:
                    rv.set(regions[i] if i < len(regions) else "")
            busy[0] = False
            on.set(True)
            changed()

        def edit_regions():
            from tkinter import simpledialog
            txt = simpledialog.askstring("지역 목록", "쉼표(,)로 나눠 적어 주세요.", initialvalue=", ".join(self.dmb_regions), parent=self.root)
            if txt is None:
                return
            self.dmb_regions = [x.strip() for x in txt.split(",") if x.strip()] or list(DMB_REGIONS)
            for _dv, _rv, cb in rows:
                cb.config(values=self.dmb_regions)
            self._schedule_autosave()

        b = ttk.Frame(f)
        b.pack(fill="x", pady=(4, 0))
        ttk.Button(b, text="원래 날짜", command=lambda: set_all(dates=[d["text"] for d in days])).pack(side="left")
        ttk.Button(b, text="지역 비우기", command=lambda: set_all(regions=[])).pack(side="left", padx=2)
        ttk.Button(b, text="지역 목록…", command=edit_regions).pack(side="right")
        ttk.Label(f, text="지역은 요일마다 지난번 것을 기억해 다음 PDF에 미리 넣습니다.", font=small, foreground="#64748b").pack(anchor="w", pady=(3, 0))
        busy[0] = False

    def on_press_pdf(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        
        if self.manual_draw_mode.get():
            self.manual_rect_start = (x, y)
            self.temp_draw_id = self.canvas.create_rectangle(x, y, x, y, outline="#9333ea", width=2, dash=(4, 4))
            return

        if self.split_mode_var.get():       # 나누기 모드: 클릭한 자리에서 둘로
            block = self._block_at(x, y)
            if block:
                vertical = bool(event.state & 0x0001)
                tgt = self._split_target(block, *self._to_doc(x, y), vertical)
                if tgt:
                    self._do_split(block, tgt[0], vertical)
            return

        bd = self._boundary_at(*self._to_doc(x, y))
        if bd:                              # 나눈 두 칸 사이 경계선 끌기 (위치 고정 중에도 됨)
            self.canvas.delete("splitprev")
            self.drag = {"type": "boundary", "axis": bd[0], "a": bd[1]["id"], "b": bd[2]["id"],
                         "snap": self._snapshot_blocks(), "moved": False}
            self.selected = {bd[1]["id"], bd[2]["id"]}
            return

        hits = self.canvas.find_overlapping(x - 2, y - 2, x + 2, y + 2)
        found = next((self.item_map[i] for i in reversed(hits) if i in self.item_map), None)
        
        if not found:
            if not (event.state & 0x0004):
                self.selected.clear()
                self.h_split_count = 1
                self.v_split_count = 1
                self.split_ratios.clear()
                self.split_titles.clear()
            self.drag = None
            self.draw_pdf()
            return
            
        block, mode = found
        clicked_id = block["id"]
        
        if clicked_id not in self.selected:
            self.h_split_count = 1
            self.v_split_count = 1
            self.split_ratios.clear()
            self.split_titles.clear()

        if event.state & 0x0004:
            if clicked_id in self.selected: self.selected.remove(clicked_id)
            else: self.selected.add(clicked_id)
        else:
            if clicked_id not in self.selected: self.selected = {clicked_id}
            
        if self.lock_move.get():          # 위치 고정: 클릭은 선택만 (끌기·크기 조절 없음)
            self.drag = None
            self.drag_start_pos = None
            self.draw_pdf()
            return

        snap = self._snapshot_blocks()   # 실제로 움직였을 때만 release에서 기록
        gid = block.get("group_id")
        group_rects = {gb["id"]: tuple(gb["rect"]) for gb in self.blocks if gid is not None and gb.get("group_id") == gid}
        self.drag = {"id": clicked_id, "mode": mode, "start_x": x, "start_y": y, "rect": tuple(block["rect"]), "group_rects": group_rects, "snap": snap}
        self.drag_start_pos = tuple(block["rect"])
        self.draw_pdf()

    def on_double_click_pdf(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        hits = self.canvas.find_overlapping(x - 2, y - 2, x + 2, y + 2)
        found = next((self.item_map[i] for i in reversed(hits) if i in self.item_map), None)
        if found:
            block, _ = found
            self.h_split_count = 1
            self.v_split_count = 1
            self.split_ratios.clear()
            self.split_titles.clear()
            self.selected = {block["id"]}
            self.drag = None
            self.draw_pdf()
            if not self.split_mode_var.get() and time.time() - getattr(self, "_last_split_at", 0) > 0.8:
                self.open_title_editor(block)

    def get_snapped_rect(self, current_rect, exclude_id=None, snap_dist=10):
        x0, y0, x1, y1 = current_rect
        doc_snap_dist = snap_dist / self.zoom
        tx0, ty0, tx1, ty1 = x0, y0, x1, y1
        
        for b in self.blocks:
            if exclude_id is not None:
                ex = exclude_id if isinstance(exclude_id, (set, frozenset)) else {exclude_id}
                if b["id"] in ex:
                    continue
            bx0, by0, bx1, by1 = b["rect"]
            
            if abs(tx0 - bx0) < doc_snap_dist: tx0 = bx0
            elif abs(tx0 - bx1) < doc_snap_dist: tx0 = bx1
                
            if abs(ty0 - by0) < doc_snap_dist: ty0 = by0
            elif abs(ty0 - by1) < doc_snap_dist: ty0 = by1
                
            if abs(tx1 - bx0) < doc_snap_dist: tx1 = bx0
            elif abs(tx1 - bx1) < doc_snap_dist: tx1 = bx1
                
            if abs(ty1 - by0) < doc_snap_dist: ty1 = by0
            elif abs(ty1 - by1) < doc_snap_dist: ty1 = by1
                
        return (tx0, ty0, tx1, ty1)

    def _draw_soon(self, after=None):
        """끄는 동안에는 마우스가 움직일 때마다 다 그리지 않고, 화면 한 번(약 1/60초)에 한 번만 그린다."""
        self._draw_after = after
        if getattr(self, "_draw_job", None):
            return

        def run():
            self._draw_job = None
            self.draw_pdf(refresh_sidebar=False)
            cb, self._draw_after = getattr(self, "_draw_after", None), None
            if cb:
                cb()
        self._draw_job = self.root.after(16, run)

    def on_motion_pdf(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        
        if self.manual_draw_mode.get() and self.manual_rect_start:
            sx, sy = self.manual_rect_start
            ox, oy = self._page_origin()
            raw_x0 = (min(sx, x) - 18) / self.zoom + ox
            raw_y0 = (min(sy, y) - 18) / self.zoom + oy
            raw_x1 = (max(sx, x) - 18) / self.zoom + ox
            raw_y1 = (max(sy, y) - 18) / self.zoom + oy
            
            snapped = self.get_snapped_rect((raw_x0, raw_y0, raw_x1, raw_y1), snap_dist=12)
            csx0, csy0, csx1, csy1 = self._xy(snapped)
            
            self.canvas.coords(self.temp_draw_id, csx0, csy0, csx1, csy1)
            return

        if self.drag and self.drag.get("type") == "boundary":
            self._drag_boundary(x, y)
            return

        if not self.drag: return
        block = next((b for b in self.blocks if b["id"] == self.drag["id"]), None)
        if not block: return
        # 클릭할 때 손이 조금 흔들려도 블록이 움직이거나 옆 칸에 붙지 않도록: 5픽셀 이상 끌어야 움직인다
        if not self.drag.get("started"):
            if abs(x - self.drag["start_x"]) < 5 and abs(y - self.drag["start_y"]) < 5:
                return
            self.drag["started"] = True
        
        dx = (x - self.drag["start_x"]) / self.zoom
        dy = (y - self.drag["start_y"]) / self.zoom
        
        ox0, oy0, ox1, oy1 = self.drag["rect"]
        
        if str(self.drag["mode"]).startswith("rs_"):
            m = self.drag["mode"][3:]
            ex0, ey0, ex1, ey1 = ox0, oy0, ox1, oy1
            if "w" in m: ex0 = min(ox0 + dx, ox1 - 20)
            if "e" in m: ex1 = max(ox1 + dx, ox0 + 20)
            if "n" in m: ey0 = min(oy0 + dy, oy1 - 10)
            if "s" in m: ey1 = max(oy1 + dy, oy0 + 10)
            new_rect = (ex0, ey0, ex1, ey1)
        elif self.drag["mode"] == "resize":
            raw_x1 = max(ox0 + 20, ox1 + dx)
            raw_y1 = max(oy0 + 10, oy1 + dy)
            new_rect = (ox0, oy0, raw_x1, raw_y1)
        else:
            w_val = ox1 - ox0
            h_val = oy1 - oy0
            nx0, ny0 = ox0 + dx, oy0 + dy
            new_rect = (nx0, ny0, nx0 + w_val, ny0 + h_val)
            
        target_rect = self.get_snapped_rect(new_rect, exclude_id=block["id"])
        if str(self.drag["mode"]).startswith("rs_"):
            m = self.drag["mode"][3:]
            tx0, ty0, tx1, ty1 = target_rect
            target_rect = (tx0 if "w" in m else ox0, ty0 if "n" in m else oy0,
                           tx1 if "e" in m else ox1, ty1 if "s" in m else oy1)
            if target_rect[2] - target_rect[0] < 20 or target_rect[3] - target_rect[1] < 10:
                target_rect = new_rect
        
        group_id = block.get("group_id")
        if group_id is not None:
            start_rects = self.drag.get("group_rects") or {}
            group_blocks = [b for b in self.blocks if b.get("group_id") == group_id]
            group_ids = {b["id"] for b in group_blocks}
            _start = lambda gb: start_rects.get(gb["id"], tuple(gb["rect"]))
            g_min_x = min(_start(b)[0] for b in group_blocks)
            g_min_y = min(_start(b)[1] for b in group_blocks)
            g_max_x = max(_start(b)[2] for b in group_blocks)
            g_max_y = max(_start(b)[3] for b in group_blocks)
            
            if str(self.drag["mode"]).startswith("rs_"):
                m = self.drag["mode"][3:]
                ex0, ey0, ex1, ey1 = g_min_x, g_min_y, g_max_x, g_max_y
                if "w" in m: ex0 = min(g_min_x + dx, g_max_x - 20)
                if "e" in m: ex1 = max(g_max_x + dx, g_min_x + 20)
                if "n" in m: ey0 = min(g_min_y + dy, g_max_y - 10)
                if "s" in m: ey1 = max(g_max_y + dy, g_min_y + 10)
                sx0, sy0, sx1, sy1 = self.get_snapped_rect((ex0, ey0, ex1, ey1), exclude_id=group_ids)
                tgt = (sx0 if "w" in m else g_min_x, sy0 if "n" in m else g_min_y,
                       sx1 if "e" in m else g_max_x, sy1 if "s" in m else g_max_y)
                if tgt[2] - tgt[0] < 20 or tgt[3] - tgt[1] < 10:
                    tgt = (ex0, ey0, ex1, ey1)
                self._apply_group_rect(group_blocks, tgt)
            elif self.drag["mode"] == "resize":
                new_g_max_x = max(g_min_x + 20, g_max_x + dx)
                new_g_max_y = max(g_min_y + 10, g_max_y + dy)
                target_g_rect = self.get_snapped_rect((g_min_x, g_min_y, new_g_max_x, new_g_max_y), exclude_id=group_ids)
                
                gw = target_g_rect[2] - target_g_rect[0]
                gh = target_g_rect[3] - target_g_rect[1]
                
                for gb in group_blocks:
                    rx0, ry0, rx1, ry1 = (gb.get("rel_box") or (0, 0, 1, 1))
                    gb["rect"] = (
                        target_g_rect[0] + gw * rx0,
                        target_g_rect[1] + gh * ry0,
                        target_g_rect[0] + gw * rx1,
                        target_g_rect[1] + gh * ry1
                    )
            else:
                gw = g_max_x - g_min_x
                gh = g_max_y - g_min_y
                gnx0, gny0 = g_min_x + dx, g_min_y + dy
                target_g_rect = self.get_snapped_rect((gnx0, gny0, gnx0 + gw, gny0 + gh), exclude_id=group_ids)
                
                dx_actual = target_g_rect[0] - g_min_x
                dy_actual = target_g_rect[1] - g_min_y
                
                for gb in group_blocks:
                    bx0, by0, bx1, by1 = _start(gb)
                    gb["rect"] = (bx0 + dx_actual, by0 + dy_actual, bx1 + dx_actual, by1 + dy_actual)
        else:
            block["rect"] = target_rect
            
        self._draw_soon()

    def on_release_pdf(self, event):
        job = getattr(self, "_draw_job", None)
        if job:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
            self._draw_job = None
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        
        if self.manual_draw_mode.get() and self.manual_rect_start:
            sx, sy = self.manual_rect_start
            ox, oy = self._page_origin()
            x0 = (min(sx, x) - 18) / self.zoom + ox
            y0 = (min(sy, y) - 18) / self.zoom + oy
            x1 = (max(sx, x) - 18) / self.zoom + ox
            y1 = (max(sy, y) - 18) / self.zoom + oy
            
            if self.temp_draw_id:
                self.canvas.delete(self.temp_draw_id)
                self.temp_draw_id = None
            self.manual_rect_start = None
            
            if (x1 - x0) > 10 and (y1 - y0) > 10:
                snapped = self.get_snapped_rect((x0, y0, x1, y1), snap_dist=12)
                self.push_history()
                new_b = {
                    "id": self.next_id,
                    "title": "",
                    "color": "",
                    "text_color": "#000000",
                    "font_size_offset": 0,
                    "options": {
                        "top_time_enabled": False, "top_time_val": "00", "top_time_font_size": DEFAULT_TIME_FONT_SIZE,
                        "bottom_time_enabled": False, "bottom_time_val": "00", "bottom_time_font_size": DEFAULT_TIME_FONT_SIZE,
                        "badge_size": 12, "uhd": False, "sub": False, "sign": False, "des": False, "local": False, "age": "none"
                    },
                    "rect": snapped,
                    "orig_rect": snapped,
                    "group_id": None,
                    "rel_box": None
                }
                self.next_id += 1
                self.blocks.append(new_b)
                self.selected = {new_b["id"]}
                self.manual_draw_mode.set(False)   # 블록을 그리면 수동 그리기 모드 자동 해제
                self.page_blocks[self.current_page] = self.blocks
                self.save_autosave()
                self.draw_pdf()
            return

        if self.drag and self.drag.get("type") == "boundary":
            for _bid in (self.drag.get("a"), self.drag.get("b")):      # 경계를 옮기면 두 칸의 시간 값도 새 자리에 맞춘다
                _bb = self._block_by_id(_bid)
                if _bb:
                    self._sync_auto_times(_bb)
            if self.drag.get("moved"):
                self._push_snapshot(self.drag["snap"])
                self.page_blocks[self.current_page] = self.blocks
                self.save_autosave()
            self.drag = None
            self.draw_pdf()
            return

        if self.drag:
            block = next((b for b in self.blocks if b["id"] == self.drag["id"]), None)
            if block and self.drag_start_pos != tuple(block["rect"]):
                if self.drag.get("snap") is not None:
                    self._push_snapshot(self.drag["snap"])
                self.page_blocks[self.current_page] = self.blocks
                self.save_autosave()
            self.drag = None
            self.draw_pdf()


if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleEditor(root)
    root.mainloop()