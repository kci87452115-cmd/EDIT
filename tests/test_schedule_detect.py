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
    print(f"{'OK  ' if ok else 'FAIL'} {fn:16s} 방식={method} 블록={len(blocks)} (최소 {min_n}) 작은칸={small} (최소 {min_small}) 머리칸={head}")
    if not ok:
        fail.append(fn)
    doc.close()
if fail:
    sys.exit(f"블록 찾기 점검 실패: {fail}")
print("블록 찾기 점검 통과")
