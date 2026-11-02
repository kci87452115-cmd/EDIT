import os
import re
import time
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

APP_VERSION = "3.9.31"
APP_TITLE = f"편성표 PDF 스마트 편집기 v{APP_VERSION} (Full Integrated Edition)"


def schedule_settings_path():
    """설정(자동 저장) 파일 위치. 편집 프로그램 런처로 실행하면 C:\\EDIT\\TV\\데이터\\편성표\\settings.json,
    따로 실행하면 예전처럼 사용자 폴더. 새 위치가 비어 있으면 예전 설정을 한 번 복사해 이어받는다."""
    old = os.path.join(os.path.expanduser("~"), ".schedule_editor_settings.json")
    home = os.environ.get("RADIO_HOME")
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


def _rgb_to_hex(c):
    """PyMuPDF 색(0~1 RGB 튜플) -> '#rrggbb'. 없거나 변환 실패 시 ''."""
    try:
        if not c:
            return ""
        return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(v * 255)))) for v in c[:3])
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
        self.mod_border_width = None           # 수정 블록 공통 테두리 두께(pt), None = 원본 유지
        self._mod_ids = set()

        # 블록 위치 고정: 칸은 옮기지 않고 합치기·나누기·내용 수정만 한다 (기본 켜짐)
        self.lock_move = tk.BooleanVar(value=True)
        self.lock_move.trace_add("write", lambda *a: self._on_lock_changed())
        self._lock_hint_shown = False
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

    def on_closing_app(self):
        self.save_autosave()
        if self.pdf_path and any(self._is_modified(b) for b in self.blocks):
            ans = messagebox.askyesnocancel("종료 확인", "작업 중인 내용이 있습니다.\nPDF로 저장하고 종료하시겠습니까?\n(아니오: 저장하지 않고 종료)")
            if ans is True:
                if self.export_pdf():
                    self._exit_clean()
            elif ans is False:
                self._exit_clean()
            else:
                return
        else:
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
                "orig_fill": b.get("orig_fill", "")
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
                "mod_shade_enabled": self.mod_shade_enabled,
                "mod_shade_color": self.mod_shade_color,
                "mod_shade_alpha": self.mod_shade_alpha,
                "mod_shade_in_export": self.mod_shade_in_export,
                "mod_border_width": self.mod_border_width,
                "lock_move": bool(self.lock_move.get()),
                "detect_min_w": self.detect_min_w,
                "detect_min_h": self.detect_min_h,
            }
            text = json.dumps(data, ensure_ascii=False)   # 먼저 문자열로 만들어 두어야 직렬화 실패 시 기존 파일이 깨지지 않는다
            tmp_path = self.backup_file_path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(text)
            os.replace(tmp_path, self.backup_file_path)
        except Exception:
            pass

    def load_autosave(self):
        if not os.path.exists(self.backup_file_path):
            return
        try:
            with open(self.backup_file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
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
                if isinstance(data.get("lock_move"), bool):
                    self.lock_move.set(data["lock_move"])
                for key, lo, hi in (("detect_min_w", 5.0, 200.0), ("detect_min_h", 1.0, 100.0)):
                    v = data.get(key)
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        setattr(self, key, max(lo, min(hi, float(v))))
                mbw = data.get("mod_border_width")
                if isinstance(mbw, (int, float)) and not isinstance(mbw, bool):
                    self.mod_border_width = max(0.25, min(6.0, float(mbw)))
        except Exception:
            pass

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
        chk_draw = ttk.Checkbutton(self.toolbar, text="✏️ 수동 블록 그리기", variable=self.manual_draw_mode)
        chk_draw.pack(side="left", padx=5)
        
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(self.toolbar, text="↩ 되돌리기 (Ctrl+Z)", command=self.undo_action).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="선택 블록 합치기", command=self.merge_selected).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="삭제 (Del)", command=self.delete_selected_blocks).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="작업 초기화", command=self.clear_all_blocks).pack(side="left", padx=3)
        
        ttk.Button(self.toolbar, text="🔍 저장 전 미리보기", command=self.preview_pdf).pack(side="left", padx=3)
        ttk.Button(self.toolbar, text="PDF 저장 (다운로드)", command=self.export_pdf).pack(side="left", padx=3)
        
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
        
        fav_frame = ttk.LabelFrame(self.sidebar_content_frame, text=" ⭐ 블록 파렛트 (관리 및 적용) ", padding=8)
        fav_frame.pack(fill="x", pady=4)
        
        self.fav_container = ttk.Frame(fav_frame)
        self.fav_container.pack(fill="x", pady=2)
        self.refresh_favorite_ui()
        
        ttk.Separator(self.sidebar_content_frame).pack(fill="x", pady=6)
        
        self._build_mod_panel()
        ttk.Separator(self.sidebar_content_frame).pack(fill="x", pady=6)

        self.info_frame = ttk.LabelFrame(self.sidebar_content_frame, text=" 🎨 선택된 블록 편집 패널 ", padding=10)
        self.info_frame.pack(fill="both", expand=True, pady=4)
        
        self.update_sidebar_info()

    # ---------------- 수정된 블록 표시 (음영 / 테두리 두께) ----------------
    def _orig_match(self, b):
        r = tuple(b["rect"])
        for ob in self.initial_blocks.get(self.current_page, []):
            if all(abs(a - c) <= 0.5 for a, c in zip(r, ob["rect"])):
                return ob
        return None

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

    def _draw_shade(self, c, x0, y0, x1, y1):
        color = self.mod_shade_color
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

    def _build_mod_panel(self):
        f = ttk.LabelFrame(self.sidebar_content_frame, text=" 🟨 수정 블록 표시 ", padding=8)
        f.pack(fill="x", pady=4)
        fnt = (self.ui_font_family, 8, "bold")

        self.mod_shade_var = tk.BooleanVar(value=self.mod_shade_enabled)
        ttk.Checkbutton(f, text="수정된 블록에 반투명 음영 표시", variable=self.mod_shade_var,
                        command=self.on_mod_style_changed).pack(anchor="w")
        self.mod_export_var = tk.BooleanVar(value=self.mod_shade_in_export)
        ttk.Checkbutton(f, text="저장 PDF에도 음영 포함 (끄면 미리보기에만)", variable=self.mod_export_var,
                        command=self.on_mod_style_changed).pack(anchor="w")

        row_c = ttk.Frame(f)
        row_c.pack(fill="x", pady=2)
        ttk.Label(row_c, text="음영 색:", font=fnt).pack(side="left")
        self.mod_color_btn = tk.Button(row_c, text="  색 선택  ", bg=self.mod_shade_color,
                                       font=(self.ui_font_family, 8), command=self.pick_mod_shade_color)
        self.mod_color_btn.pack(side="right")

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
        self._refresh_mod_labels()

    def _refresh_mod_labels(self):
        try:
            self.mod_alpha_lbl.config(text=f"{int(round(self.mod_shade_alpha * 100))}%")
            self.mod_border_lbl.config(text="원본" if self.mod_border_width is None else f"{self.mod_border_width:.2f}pt")
            self.mod_color_btn.config(bg=self.mod_shade_color)
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
        self._mod_style_apply()

    def pick_mod_shade_color(self):
        c = colorchooser.askcolor(title="수정 블록 음영 색 선택", initialcolor=self.mod_shade_color)
        if c[1]:
            self.mod_shade_color = c[1]
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

            if page_idx not in self.page_blocks:
                drawings = page.get_drawings()
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
                self._detect_note = f"블록 {len(detected_blocks)}개를 찾았습니다 ({method} 기준"\
                                    f", 높이 10pt 미만 작은 칸 {sum(1 for x in found if x['rect'][3] - x['rect'][1] < 10)}개 포함)."
                self.__dict__.setdefault("orig_lines", {})[page_idx] = _collect_line_segments(drawings)
                self.page_blocks[page_idx] = detected_blocks
                self.initial_blocks[page_idx] = copy.deepcopy(detected_blocks)

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
        messagebox.showinfo("페이지 로드 완료", load_msg)
        warn = getattr(self, "_clear_fill_warning", None)
        if warn:
            messagebox.showwarning("채우기 제거 안내", warn)

    def bind_pdf_events(self):
        self.canvas.bind("<Button-1>", self.on_press_pdf)
        self.canvas.bind("<Double-Button-1>", self.on_double_click_pdf)
        self.canvas.bind("<B1-Motion>", self.on_motion_pdf)
        self.canvas.bind("<ButtonRelease-1>", self.on_release_pdf)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)

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
        
        if is_selected:
            hit_id = c.create_rectangle(x0, y0, x1, y1, fill=fill_color if fill_color else "#dbeafe", outline="", stipple="gray25" if not fill_color else "")
        else:
            hit_id = c.create_rectangle(x0, y0, x1, y1, fill=fill_color, outline="")
            
        if self.mod_shade_enabled and b["id"] in self._mod_ids:
            self._draw_shade(c, x0, y0, x1, y1)

        rect_id = c.create_rectangle(x0, y0, x1, y1, fill="", outline=outline_color, width=width)
        
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
        self.photo = tk.PhotoImage(data=base64.b64encode(self.png_data).decode("ascii"))
        
        pw = (self.page_rect[2] - self.page_rect[0]) if self.page_rect else 1000
        ph = (self.page_rect[3] - self.page_rect[1]) if self.page_rect else 800
        page_width, page_height = pw * self.zoom, ph * self.zoom
        
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
            
            self.opt_top_time_enabled_var = tk.BooleanVar(value=opts.get("top_time_enabled", opts.get("time_enabled", False)))
            self.top_time_val_var = tk.StringVar(value=opts.get("top_time_val", opts.get("time_val", "00")))
            
            self.opt_bottom_time_enabled_var = tk.BooleanVar(value=opts.get("bottom_time_enabled", False))
            self.bottom_time_val_var = tk.StringVar(value=opts.get("bottom_time_val", "00"))

            row_top_time = ttk.Frame(opt_frame)
            row_top_time.pack(fill="x", pady=2)
            
            ttk.Checkbutton(row_top_time, text="상단시간(위):", variable=self.opt_top_time_enabled_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=(0, 2))
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
            
            ttk.Checkbutton(row_bottom_time, text="하단시간(아래):", variable=self.opt_bottom_time_enabled_var, command=lambda: self.on_option_changed(block)).pack(side="left", padx=(0, 2))
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

            split_frame = ttk.LabelFrame(self.info_frame, text=" ✂ 선택적 블록 분할 ", padding=6)
            split_frame.pack(fill="x", pady=4)
            
            h_row = ttk.Frame(split_frame)
            h_row.pack(fill="x", pady=1)
            ttk.Button(h_row, text="-", width=2, command=lambda: self.adjust_split_count('h', -1)).pack(side="left")
            self.h_split_lbl = ttk.Label(h_row, text=f"가로 {self.h_split_count}분할", anchor="center", font=(self.ui_font_family, 8, "bold"))
            self.h_split_lbl.pack(side="left", fill="x", expand=True)
            ttk.Button(h_row, text="+", width=2, command=lambda: self.adjust_split_count('h', 1)).pack(side="left")

            if self.h_split_count == 2:
                h_ratio_row = ttk.Frame(split_frame)
                h_ratio_row.pack(fill="x", pady=2)
                ttk.Label(h_ratio_row, text="상하 비율:", font=(self.ui_font_family, 7, "bold")).pack(side="left")
                ttk.Button(h_ratio_row, text="-", width=2, command=lambda: self.adjust_h_ratio(-0.01)).pack(side="left", padx=1)
                curr_h_ratio = self.split_ratios[0] if len(self.split_ratios) > 0 else 0.5
                self.h_ratio_scale = ttk.Scale(h_ratio_row, from_=0.1, to=0.9, value=curr_h_ratio, command=lambda v: self.on_h_ratio_slider_change(v))
                self.h_ratio_scale.pack(side="left", fill="x", expand=True, padx=2)
                ttk.Button(h_ratio_row, text="+", width=2, command=lambda: self.adjust_h_ratio(0.01)).pack(side="left", padx=1)
            
            v_row = ttk.Frame(split_frame)
            v_row.pack(fill="x", pady=1)
            ttk.Button(v_row, text="-", width=2, command=lambda: self.adjust_split_count('v', -1)).pack(side="left")
            self.v_split_lbl = ttk.Label(v_row, text=f"세로 {self.v_split_count}분할", anchor="center", font=(self.ui_font_family, 8, "bold"))
            self.v_split_lbl.pack(side="left", fill="x", expand=True)
            ttk.Button(v_row, text="+", width=2, command=lambda: self.adjust_split_count('v', 1)).pack(side="left")

            if self.v_split_count == 2:
                v_ratio_row = ttk.Frame(split_frame)
                v_ratio_row.pack(fill="x", pady=2)
                ttk.Label(v_ratio_row, text="좌우 비율:", font=(self.ui_font_family, 7, "bold")).pack(side="left")
                ttk.Button(v_ratio_row, text="-", width=2, command=lambda: self.adjust_v_ratio(-0.01)).pack(side="left", padx=1)
                curr_ratio = self.split_ratios[0] if len(self.split_ratios) > 0 else 0.5
                self.ratio_scale = ttk.Scale(v_ratio_row, from_=0.1, to=0.9, value=curr_ratio, command=lambda v: self.on_ratio_slider_change(v))
                self.ratio_scale.pack(side="left", fill="x", expand=True, padx=2)
                ttk.Button(v_ratio_row, text="+", width=2, command=lambda: self.adjust_v_ratio(0.01)).pack(side="left", padx=1)

            total_splits = max(self.h_split_count, self.v_split_count)
            self.split_title_vars = {}
            if total_splits > 1:
                sub_input_frame = ttk.Frame(split_frame)
                sub_input_frame.pack(fill="x", pady=(4, 2))
                ttk.Label(sub_input_frame, text=f"분할된 {total_splits}개 영역별 명칭 입력:", font=(self.ui_font_family, 7, "bold"), foreground="#2563eb").pack(anchor="w", pady=(1, 2))
                
                first_split_entry = None
                for idx in range(total_splits):
                    if idx not in self.split_titles:
                        self.split_titles[idx] = block["title"] if idx == 0 else ""
                        
                    s_row = ttk.Frame(sub_input_frame)
                    s_row.pack(fill="x", pady=1)
                    ttk.Label(s_row, text=f"#{idx+1}", width=3, font=(self.ui_font_family, 7)).pack(side="left")
                    
                    st_var = tk.StringVar(value=self.split_titles[idx])
                    self.split_title_vars[idx] = st_var
                    sub_entry = ttk.Entry(s_row, textvariable=st_var, font=(self.ui_font_family, 8))
                    sub_entry.pack(side="left", fill="x", expand=True, padx=2)
                    if first_split_entry is None:
                        first_split_entry = sub_entry
                    
                    sub_entry.bind("<KeyRelease>", lambda e, i=idx, v=st_var: self.confirm_split_title(i, v.get()))
                    sub_entry.bind("<Return>", lambda e, i=idx, v=st_var: self.confirm_split_title(i, v.get()))
                    sub_entry.bind("<FocusOut>", lambda e, i=idx, v=st_var: self.confirm_split_title(i, v.get(), immediate=True))

                if first_split_entry is not None:
                    first_split_entry.focus_set()
                    first_split_entry.icursor(tk.END)

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
            block["options"] = {
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
        out_path = filedialog.asksaveasfilename(title="PDF 저장", defaultextension=".pdf", filetypes=[("PDF 파일", "*.pdf")])
        if not out_path:
            return False
            
        try:
            with PDF_LOCK:
                pdf_bytes = self._generate_pdf_bytes()
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            messagebox.showinfo("저장 완료", f"PDF 파일이 성공적으로 저장되었습니다:\n{out_path}")
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"PDF 파일 저장 실패:\n{str(e)}")
            return False

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

        if font_name == "helv" and any(not (b.get("title") or "").isascii() for b in blocks):
            doc.close()
            raise RuntimeError("한글 폰트를 불러오지 못했습니다.\n설정(⚙️)에서 .ttf/.ttc 한글 폰트 파일을 지정해 주세요.")

        try:
            measure_font = fitz.Font(fontfile=font_path) if font_name == "custom_korean_font" else fitz.Font("helv")
        except Exception:
            measure_font = None

        mod_ids = {b["id"] for b in blocks if self._is_modified(b)}

        # 1) 채우기: 원본 테두리를 덮지 않도록 선 굵기의 절반만큼 안쪽으로 줄여서 그린다.
        for b in blocks:
            hex_c = (b.get("color") or "").lstrip("#")
            if len(hex_c) != 6:
                continue
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
                if len((b.get("color") or "").lstrip("#")) == 6:
                    _restore_lines(page, segs, fitz.Rect(b["rect"]))

        # 2) 테두리: 원본에서 찾은 칸과 위치/크기가 다른 블록(이동, 크기 변경, 분할, 병합, 수동 추가)은
        #    원본 테두리가 없거나 어긋나 있으므로 같은 스타일로 새로 그린다.
        # 1-3) 수정된 블록 반투명 음영 (미리보기, 또는 '저장 PDF에도 음영 포함'을 켠 경우)
        if self.mod_shade_enabled and (for_preview or self.mod_shade_in_export):
            sh = self.mod_shade_color.lstrip("#")
            sh_rgb = tuple(int(sh[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
            for b in blocks:
                if b["id"] in mod_ids:
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
            if rb[1] > 0.001:
                page.draw_line(fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y0), color=col, width=wid)
            if rb[0] > 0.001:
                page.draw_line(fitz.Point(r.x0, r.y0), fitz.Point(r.x0, r.y1), color=col, width=wid)

        out_bytes = doc.tobytes() if for_preview else doc.tobytes(garbage=3, deflate=True)   # 저장본은 압축해서 용량을 줄인다
        doc.close()
        return out_bytes

    def on_press_pdf(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        
        if self.manual_draw_mode.get():
            self.manual_rect_start = (x, y)
            self.temp_draw_id = self.canvas.create_rectangle(x, y, x, y, outline="#9333ea", width=2, dash=(4, 4))
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
            self.draw_pdf()

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
            
        self.draw_pdf(refresh_sidebar=False)

    def on_release_pdf(self, event):
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