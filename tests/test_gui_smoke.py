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
