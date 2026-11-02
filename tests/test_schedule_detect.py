"""편성표 PDF에서 블록(칸)을 제대로 찾는지 점검 (GitHub 빌드 때 실제 PyMuPDF 로 실행).
기대값은 PDF를 분석해 확인한 칸 수 근처의 최소값이다."""
import importlib.util
import os
import sys

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("schedule_ui", os.path.join(HERE, "..", "apps", "schedule", "ui.py"))
ui = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ui)

EXPECT = {   # 파일: (최소 블록 수, 작은 칸(높이 10pt 미만) 최소 수)
    "mbc_1_6.pdf": (72, 5),
    "mbc_1_7.pdf": (70, 5),
    "mbc_4_1.pdf": (113, 15),
    "mbc_11_1_2.pdf": (122, 15),
    "mbc_2_1_1.pdf": (110, 10),
    "mbc_3_2.pdf": (112, 10),
    "mbc_4_3.pdf": (110, 10),
    "mbc_5_1_3.pdf": (115, 15),
    "mbc_5_2.pdf": (115, 15),
    "tv_3.pdf": (105, 8),
}

fail = []
for fn, (min_n, min_small) in EXPECT.items():
    doc = pymupdf.open(os.path.join(HERE, "pdfs", fn))
    page = doc.load_page(0)
    blocks, method = ui.detect_schedule_blocks(page.get_drawings(), tuple(page.rect))
    small = sum(1 for b in blocks if b["rect"][3] - b["rect"][1] < 10)
    def rgb(h):
        return [int(h[i:i + 2], 16) / 255 for i in (1, 3, 5)] if h and len(h) == 7 else (1, 1, 1)
    head = sum(1 for b in blocks if ui._is_header_color(rgb(b["orig_fill"])))   # 파란 시간·날짜 칸이 섞이면 안 됨
    ok = method == "테두리" and len(blocks) >= min_n and small >= min_small and head == 0
    gray = sum(1 for b in blocks if b["orig_fill"] in ("#d9d9d9", "#d8d8d8", "#dadada"))
    if fn.startswith("tv_") and gray < 20:     # 엑셀 PDF의 회색 칸(회색 1값 색)이 색으로 기록되는지
        ok = False
    print(f"{'OK  ' if ok else 'FAIL'} {fn:16s} 방식={method} 블록={len(blocks)} (최소 {min_n}) 작은칸={small} (최소 {min_small}) 머리칸={head} 회색칸={gray}")
    if not ok:
        fail.append(fn)
    doc.close()
if fail:
    sys.exit(f"블록 찾기 점검 실패: {fail}")
print("블록 찾기 점검 통과")


# ---------------------------------------------------------------- 시간 줄 · 머리글 · 나누기
print()
fail2 = []
for fn in EXPECT:
    doc = pymupdf.open(os.path.join(HERE, "pdfs", fn))
    page = doc.load_page(0)
    dr = page.get_drawings()
    segs, fills = ui._drawing_parts(dr)
    rows = ui.detect_time_rows(page.get_text("words"), segs)
    heads = [r.y0 for r, c in fills if ui._is_header_color(c) and r.height >= 8]
    hdr = ui.detect_header(page.get_text("dict"), page.get_image_info(), (min(heads) - 1) if heads else None, tuple(page.rect))
    ok = len(rows) == 24 and rows[0][2] == 5 and hdr.get("title") is not None and (fn.startswith("tv_") or hdr.get("date") is not None)
    # 5분 눈금이 실제 칸 경계와 맞는지
    ticks = ui.time_ticks(rows)
    blocks, _ = ui.detect_schedule_blocks(dr, tuple(page.rect))
    near = sum(1 for b in blocks if ticks and min(abs(b["rect"][1] - t[0]) for t in ticks) <= 1.5)
    ok = ok and near >= len(blocks) * 0.9
    print(f"{'OK  ' if ok else 'FAIL'} {fn:16s} 시간줄={len(rows)} ({rows[0][2] if rows else '-'}시~) 칸 윗변이 5분 눈금에 맞음={near}/{len(blocks)} "
          f"제목={'있음' if hdr.get('title') else '없음'} 날짜={(hdr.get('date') or {}).get('text', '없음')}")
    if not ok:
        fail2.append(fn)
    doc.close()

# 나누기 동작 (화면 없이 계산만)
class _Canvas:
    def __getattr__(self, name):
        return lambda *a, **k: (0, 0, 1, 1) if name == "bbox" else 1
class _Var:
    def __init__(self, v=False): self.v = v
    def get(self): return self.v
    def set(self, v): self.v = v
doc = pymupdf.open(os.path.join(HERE, "pdfs", "mbc_3_2.pdf"))
page = doc.load_page(0)
dr = page.get_drawings()
segs, _f = ui._drawing_parts(dr)
E = ui.ScheduleEditor.__new__(ui.ScheduleEditor)
E.zoom, E.page_rect, E.current_page, E.ui_font_family = 1.5, tuple(page.rect), 0, "x"
rows = ui.detect_time_rows(page.get_text("words"), segs)
E.time_rows_by_page, E.time_ticks_by_page = {0: rows}, {0: ui.time_ticks(rows)}
found, _ = ui.detect_schedule_blocks(dr, E.page_rect)
E.blocks = [{"id": i + 1, "title": f"t{i}", "rect": f["rect"], "orig_rect": f["rect"], "group_id": None, "rel_box": None} for i, f in enumerate(found)]
E.next_id, E.page_blocks, E.selected, E.drag, E.pdf_path = 1000, {0: E.blocks}, set(), None, "x"
E.canvas, E.split_mode_var, E.manual_draw_mode = _Canvas(), _Var(), _Var()
for n in ("draw_pdf", "save_autosave", "push_history", "open_title_editor"):
    setattr(E, n, lambda *a, **k: None)
E.root = type("R", (), {"after": lambda *a: None})()
blk = max(E.blocks, key=lambda b: b["rect"][3] - b["rect"][1])
x0, y0, x1, y1 = blk["rect"]
pos, lab = E._split_target(blk, (x0 + x1) / 2, y0 + (y1 - y0) / 3, False)
on_tick = any(abs(pos - t[0]) < 0.01 for t in E.time_ticks_by_page[0])
E._do_split(blk, pos, False)
parts = [b for b in E.blocks if b.get("group_id") is not None]
ok = on_tick and len(parts) == 2 and lab.endswith(("0", "5")) and E._boundary_at((x0 + x1) / 2, pos)
print(f"{'OK  ' if ok else 'FAIL'} 나누기: {lab} 에서 둘로 (5분 눈금 위={on_tick}), 경계선 끌기 대상 찾음={bool(E._boundary_at((x0 + x1) / 2, pos))}")
if not ok:
    fail2.append("split")
if fail2:
    sys.exit(f"시간 줄·머리글·나누기 점검 실패: {fail2}")
print("시간 줄·머리글·나누기 점검 통과")


# ---------------------------------------------------------------- DMB: 요일 칸 · 제목
print()
fail3 = []
for fn in EXPECT:
    doc = pymupdf.open(os.path.join(HERE, "pdfs", fn))
    page = doc.load_page(0)
    _s, fills = ui._drawing_parts(page.get_drawings())
    days = ui.detect_day_headers(page.get_text("dict"), fills)
    ok = (len(days) == 7 and [d["text"][-2] for d in days] == list("월화수목금토일")
          and all(30 < d["cell_w"] < 120 for d in days) and all(d["fill"].lower() != "#ffffff" for d in days))
    print(f"{'OK  ' if ok else 'FAIL'} 요일 칸 {fn:16s} {len(days)}개 {[d['text'] for d in days][:2]}… 폭 {round(days[0]['cell_w']) if days else 0} 색 {days[0]['fill'] if days else ''}")
    if not ok:
        fail3.append(fn)
    doc.close()

ui.IS_DMB = True
doc = pymupdf.open(os.path.join(HERE, "pdfs", "mbc_11_1_2.pdf"))
page = doc.load_page(0)
_s, fills = ui._drawing_parts(page.get_drawings())
heads = [r.y0 for r, c in fills if ui._is_header_color(c) and r.height >= 8]
td = page.get_text("dict")
hdr = ui.detect_header(td, page.get_image_info(), (min(heads) - 1) if heads else None, tuple(page.rect))
hdr["days"] = ui.detect_day_headers(td, fills)
D = ui.ScheduleEditor.__new__(ui.ScheduleEditor)
D.current_page, D.page_rect, D.pdf_path = 0, tuple(page.rect), r"C:\x\MBC_11__1__2.pdf"
D.header_found_by_page, D.header_edit, D.header_title_presets = {0: hdr}, {}, []
D.dmb_day_regions = ["전주", "전주", "여수", "목포", "여수", "광주", "광주"]
st = D._header_state()
font = pymupdf.Font("cjk")
page.insert_font(fontname="kr", fontbuffer=font.buffer)
D._draw_header_pdf(page, "kr", font)
text = page.get_text()
ok = (st["title_on"] and st["days_on"] and st["title"] == "광주 MYMBC 주간편성표 [ 11월 01주 2 ]"
      and all(w in text for w in ("전주", "여수", "목포", "광주", "MYMBC")))
print(f"{'OK  ' if ok else 'FAIL'} DMB 제목 '{st['title']}' · 요일 칸 그리기")
if not ok:
    fail3.append("dmb-draw")
D.pdf_path = "DMB_11__1.pdf"
D.header_edit = {}
if D._suggest_header_title() != "광주 MYMBC 주간편성표 [ 11월 01주 기본 ]":
    fail3.append("dmb-title-basic")
for src, want in (("MBC_11__1__2.pdf", "DMB_11__1__2.pdf"), ("편성표.pdf", "DMB_편성표.pdf"), ("DMB_11__1.pdf", "DMB_11__1.pdf")):
    D.pdf_path = os.path.join(HERE, src)
    if D._default_save_name() != want:
        fail3.append(f"save-name {src}")
ui.IS_DMB = False
doc.close()
if fail3:
    sys.exit(f"DMB 점검 실패: {fail3}")
print("DMB 요일 칸·제목 점검 통과")


# ---------------------------------------------------------------- 새 판 비교 · 고친 내용 옮기기
print()
fail4 = []


def load_initial(fn):
    doc = pymupdf.open(os.path.join(HERE, "pdfs", fn))
    page = doc.load_page(0)
    found, _ = ui.detect_schedule_blocks(page.get_drawings(), tuple(page.rect))
    texts = ui.block_source_texts(page.get_text("words"), [f["rect"] for f in found])
    out = [{"id": i + 1, "title": "", "color": "", "text_color": "#000000", "font_size_offset": 0, "options": {},
            "rect": tuple(f["rect"]), "orig_rect": tuple(f["rect"]), "group_id": None, "rel_box": None,
            "border_color": f["border_color"], "border_width": f["border_width"], "orig_fill": f["orig_fill"], "src_text": t}
           for i, (f, t) in enumerate(zip(found, texts))]
    doc.close()
    return out


for old_fn, new_fn in (("mbc_4_1.pdf", "mbc_4_3.pdf"), ("mbc_1_6.pdf", "mbc_1_7.pdf"), ("mbc_5_1_3.pdf", "mbc_5_2.pdf")):
    old_init, new_init = load_initial(old_fn), load_initial(new_fn)
    pairs, changed = ui.match_original_blocks(old_init, new_init)
    same = [ob for ob in old_init if ob["id"] in pairs and (ob["rect"][3] - ob["rect"][1]) > 12]
    diff = [ob for ob in old_init if ob["id"] not in pairs and (ob["rect"][3] - ob["rect"][1]) > 12]
    blocks = [dict(b) for b in old_init]
    byid = {b["id"]: b for b in blocks}
    # 고친 것: 같은 칸 3개 이름 바꾸기, 1개 음영, 1개 둘로 나누기, 바뀐 칸 1개 이름 바꾸기(옮기면 안 됨)
    for k, ob in enumerate(same[:3]):
        byid[ob["id"]]["title"] = f"바꾼이름{k}"
    byid[same[3]["id"]]["shade"] = "#86efac"
    sp = byid[same[4]["id"]]
    x0, y0, x1, y1 = sp["rect"]
    mid = round((y0 + y1) / 2, 2)
    a, b = dict(sp, id=9001, rect=(x0, y0, x1, mid), group_id=777, title="위칸"), dict(sp, id=9002, rect=(x0, mid, x1, y1), group_id=777, title="아래칸")
    blocks = [x for x in blocks if x["id"] != sp["id"]] + [a, b]
    want_skip = 0
    if diff:
        next(x for x in blocks if x["id"] == diff[0]["id"])["title"] = "옮기면안됨"
        want_skip = 1
    res, moved, skipped, ch, _nid = ui.transplant_work(old_init, blocks, new_init, 10000)
    titles = {x.get("title") for x in res}
    ok = (moved == 6 and len(skipped) == want_skip and {"바꾼이름0", "바꾼이름1", "바꾼이름2", "위칸", "아래칸"} <= titles
          and "옮기면안됨" not in titles and sum(1 for x in res if x.get("shade")) == 1
          and len(res) == len(new_init) + 1 and 0 < len(ch) < len(new_init))
    print(f"{'OK  ' if ok else 'FAIL'} 판 옮기기 {old_fn} -> {new_fn}: 같은 칸 {len(pairs)}/{len(old_init)}, 새 판에서 바뀐 칸 {len(ch)}, "
          f"옮김 {moved}, 못 옮김 {len(skipped)} (기대 {want_skip}), 칸 수 {len(res)} = {len(new_init)}+1")
    if not ok:
        fail4.append(old_fn)
if fail4:
    sys.exit(f"판 옮기기 점검 실패: {fail4}")
print("새 판 비교·옮기기 점검 통과")
