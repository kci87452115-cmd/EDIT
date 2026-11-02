"""화면 점검 (Windows 빌드에서 실제 Tk 로 실행): DMB 로 PDF 를 열고
제목·날짜·요일 칸이 편집 화면(캔버스)에 바로 보이는지, 음영이 그려지는지 확인한다."""
import base64
import importlib.util
import io
import os
import sys
import tempfile
import traceback

os.environ["RADIO_HOME"] = tempfile.mkdtemp()
os.environ["EDIT_PROFILE"] = "dmb"
import tkinter as tk

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("schedule_ui", os.path.join(HERE, "..", "apps", "schedule", "ui.py"))
ui = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ui)
for n in ("showinfo", "showwarning", "showerror"):
    setattr(ui.messagebox, n, lambda *a, **k: None)
ui.messagebox.askyesno = lambda *a, **k: False
ui.messagebox.askokcancel = lambda *a, **k: False
ui.messagebox.askyesnocancel = lambda *a, **k: None

fail = []
root = tk.Tk()
root.geometry("1480x1000+0+0")
E = ui.ScheduleEditor(root)
path = os.path.join(HERE, "pdfs", "mbc_11_1_2.pdf")
E.pdf_path, E.current_page = path, 0
E._load_pdf_worker_locked(path, 0)
for _ in range(30):
    root.update()
print("DMB =", ui.IS_DMB, "블록", len(E.blocks), "요일", len((E.header_found_by_page.get(0) or {}).get("days") or []))
E.dmb_day_regions = ["전주", "전주", "여수", "목포", "여수", "광주", "광주"]
E.header_edit = {}
st = E._header_state()
st["date_on"], st["date"] = True, "2026년 11월 2일 (월) PM 06:00"
try:
    E.update_sidebar_info()
except Exception:
    traceback.print_exc()
    fail.append("sidebar")
E.draw_pdf()
c = E.canvas
try:
    E._draw_header_preview(c)
except Exception:
    traceback.print_exc()
    fail.append("header-preview")
texts = [c.itemcget(i, "text") for i in c.find_all() if c.type(i) == "text"]
for want in ("MYMBC", "11/2(월)전주", "PM 06:00"):
    hit = [t for t in texts if want in t]
    print(f"{'OK  ' if hit else 'FAIL'} 화면 글자 '{want}': {hit[:1]}")
    if not hit:
        fail.append(want)

# 그 자리 맨 위에 무엇이 있나 (가려지는지)
from tkinter import font as tkfont
def measure(t, s):
    return tkfont.Font(family=E.ui_font_family, size=-max(4, int(round(s * E.zoom)))).measure(t) / E.zoom
for cover, text, size, (tx, ty), bold, bg in E._header_layout(measure)[:3]:
    x, y = E._to_canvas(tx + 2, ty - size / 3)
    top = c.find_overlapping(x, y, x + 1, y + 1)
    print("  ", text[:20], "맨 위:", [(c.type(i), (c.itemcget(i, "text") if c.type(i) == "text" else "")[:12]) for i in top[-3:]])

# 음영: 수정 안 한 칸
E.mod_shade_enabled, E.shade_target = True, "unmod"
E.draw_pdf()
n_img = sum(1 for i in c.find_all() if c.type(i) == "image")
print(f"{'OK  ' if n_img > 50 else 'FAIL'} 수정 안 한 칸 음영 그림 {n_img - 1}개")
if n_img <= 50:
    fail.append("shade-unmod")
E.shade_target = "both"
E.draw_pdf()

# 닫기 확인: 막 저장한 상태 -> 저장할 것 없음, 음영을 넣으면 -> 저장 안 한 작업
E._mark_saved()
clean_before = not E._has_unsaved_work()
E.on_closing_app()          # askokcancel=False 로 바꿔 두었으므로 닫히지 않아야 한다
alive = bool(root.winfo_exists())

# 여러 칸 골라 음영 넣기 / 빼기 (H)
E.mod_shade_enabled = False
E.selected = {b["id"] for b in E.blocks[:5]}
E.shade_selected("#86efac")
n1 = sum(1 for b in E.blocks if b.get("shade") == "#86efac")
E.draw_pdf()
n_img1 = sum(1 for i in c.find_all() if c.type(i) == "image") - 1
E.toggle_pick_shade()
n2 = sum(1 for b in E.blocks if b.get("shade"))
E.toggle_pick_shade()
ok = n1 == 5 and n_img1 == 5 and n2 == 0 and sum(1 for b in E.blocks if b.get("shade")) == 5
dirty = E._has_unsaved_work()
print(f"{'OK  ' if (clean_before and alive and dirty) else 'FAIL'} 닫기 확인: 저장 직후 깨끗={clean_before}, X 눌러도 안 닫힘={alive}, 음영 뒤 저장할 것 있음={dirty}")
if not (clean_before and alive and dirty):
    fail.append("close-confirm")
print(f"{'OK  ' if ok else 'FAIL'} 선택 칸 음영: 넣기 {n1}칸, 화면 {n_img1}개, H로 빼기 후 {n2}칸")
if not ok:
    fail.append("pick-shade")
try:
    E._generate_pdf_bytes()
except Exception:
    traceback.print_exc()
    fail.append("export-with-shade")
E.selected = set()
E.mod_shade_enabled = True
E.draw_pdf()

# 이전 판 작업 옮기기: 4주 1판에서 고치고 -> 3판을 열면 묻고 옮긴다
def open_pdf(fn):
    E.pdf_path, E.current_page = os.path.join(HERE, "pdfs", fn), 0
    E.page_blocks, E.initial_blocks, E.header_edit, E.changed_ids_by_page = {}, {}, {}, {}
    E._load_pdf_worker_locked(E.pdf_path, 0)
    for _ in range(30):
        root.update()

ui.messagebox.askyesno = lambda *a, **k: True
open_pdf("mbc_4_1.pdf")
import pymupdf
_d = pymupdf.open(os.path.join(HERE, "pdfs", "mbc_4_3.pdf"))
_nt = ui.block_source_texts(_d[0].get_text("words"), [b["rect"] for b in E.blocks])
_d.close()
big = [b for b, nt in zip(E.blocks, _nt) if b.get("src_text") and nt == b["src_text"] and b["rect"][3] - b["rect"][1] > 15][:3]
E.selected = {big[0]["id"]}
E.shade_selected("#fde047")
big[1]["title"] = "옮겨질 이름"
E.save_autosave()
recs = E._list_work_records()
open_pdf("mbc_4_3.pdf")
E._offered_carry_for = None
E.offer_carry_over()
titles = [b.get("title") for b in E.blocks]
ok = (len(recs) >= 1 and "옮겨질 이름" in titles and any(b.get("shade") for b in E.blocks)
      and len(E.changed_ids_by_page.get(0) or ()) > 0)
print(f"{'OK  ' if ok else 'FAIL'} 이전 판 옮기기 (화면): 기록 {len(recs)}개, 이름 옮김={'옮겨질 이름' in titles}, "
      f"음영 옮김={any(b.get('shade') for b in E.blocks)}, 바뀐 칸 표시 {len(E.changed_ids_by_page.get(0) or ())}개")
if not ok:
    fail.append("carry-over")

# 화면 사진 (로그에 작게)
try:
    from PIL import ImageGrab
    root.update()
    x, y = c.winfo_rootx(), c.winfo_rooty()
    img = ImageGrab.grab((x, y, x + c.winfo_width(), y + c.winfo_height())).convert("RGB")
    img.thumbnail((900, 900))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=60)
    print("SCREENSHOT_B64_BEGIN")
    print(base64.b64encode(buf.getvalue()).decode())
    print("SCREENSHOT_B64_END")
except Exception as exc:
    print("화면 사진 실패:", exc)

root.destroy()
if fail:
    sys.exit(f"화면 점검 실패: {fail}")
print("화면 점검 통과")
