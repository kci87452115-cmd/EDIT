"""빌드 점검: 만든 업데이트 폴더를 런처와 똑같은 방법으로 확인한다 (서명 + 모든 꾸러미 지문).
또 서명을 한 글자 바꾸거나 파일 내용을 바꾸면 런처가 거부하는지도 본다."""
import importlib.util
import json
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
spec = importlib.util.spec_from_file_location("launcher", os.path.join(ROOT, "launcher", "launcher.py"))
L = importlib.util.module_from_spec(spec)
spec.loader.exec_module(L)

folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Output", "update")
cat = L.read_catalog(folder)
if cat is None:
    sys.exit(f"서명 확인 실패: {L.CATALOG_ERROR or 'catalog.json 을 읽지 못함'}")
files = [cat["launcher"]["file"]] + [e["file"] for e in cat["apps"].values()]
for fn in files:
    L.check_file(cat, os.path.join(folder, fn))
print(f"OK   서명 확인, 꾸러미 {len(files)}개 지문 일치")

# 바뀐 목록·파일은 거부해야 한다
tmp = tempfile.mkdtemp()
try:
    for fn in ["catalog.json", "catalog.sig"] + [e["file"] for e in cat["apps"].values()]:
        shutil.copy2(os.path.join(folder, fn), tmp)
    data = json.load(open(os.path.join(tmp, "catalog.json"), encoding="utf-8"))
    first = next(iter(data["apps"]))
    data["apps"][first]["version"] = "99.0"
    json.dump(data, open(os.path.join(tmp, "catalog.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    bad_cat = L.read_catalog(tmp) is None and "서명이 맞지 않" in L.CATALOG_ERROR
    shutil.copy2(os.path.join(folder, "catalog.json"), tmp)
    good = L.read_catalog(tmp)
    zf = os.path.join(tmp, good["apps"][first]["file"])
    with open(zf, "ab") as f:
        f.write(b"tampered")
    try:
        L.check_file(good, zf)
        bad_file = False
    except L.UpdateSecurityError:
        bad_file = True
    os.remove(os.path.join(tmp, "catalog.sig"))
    no_sig = L.read_catalog(tmp) is None and "catalog.sig" in L.CATALOG_ERROR
finally:
    shutil.rmtree(tmp, ignore_errors=True)
ok = bad_cat and bad_file and no_sig
print(f"{'OK  ' if ok else 'FAIL'} 바뀐 목록 거부={bad_cat}, 바뀐 파일 거부={bad_file}, 서명 없는 꾸러미 거부={no_sig}")
if not ok:
    sys.exit("바뀐 꾸러미를 거부하지 못했습니다.")
