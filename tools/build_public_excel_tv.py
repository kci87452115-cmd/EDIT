"""배포판(엑셀 TV 편성표 편집기) 정보 파일 만들기 — GitHub Actions 에서 PyInstaller 다음에 실행

입력:  dist/ExcelTVEditor_public.exe
출력:  Output/public/ExcelTVEditor_public.exe
       Output/public/excel_tv_public.json   (버전·날짜·크기·지문·바뀐 내용)
       Output/public/excel_tv_public.sig    (EDIT_SIGN_KEY 서명 — 배포판이 업데이트 전에 확인)
녹음 PC 서버는 이 json 을 보고 웹 '다운로드' 파일을 바꾸고, 설치된 배포판은 이것으로 새 버전을 안다.
"""
import base64
import datetime
import hashlib
import json
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "Output", "public")
EXE = os.path.join(ROOT, "dist", "ExcelTVEditor_public.exe")


def main():
    msg = sys.argv[1] if len(sys.argv) > 1 else ""
    src = open(os.path.join(ROOT, "apps", "excel_tv", "tv_schedule_maker.py"), encoding="utf-8").read()
    ver = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', src, flags=re.M).group(1)
    if not os.path.exists(EXE):
        sys.exit("배포판 exe 가 없습니다: " + EXE)
    os.makedirs(OUT, exist_ok=True)
    dst = os.path.join(OUT, os.path.basename(EXE))
    shutil.copyfile(EXE, dst)
    data = open(dst, "rb").read()
    info = {"id": "excel_tv", "name": "엑셀 TV 편성표 편집기", "edition": "배포판", "version": ver,
            "date": (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).strftime("%Y-%m-%d"),
            "size": len(data), "sha256": hashlib.sha256(data).hexdigest(),
            "changes": [msg[:200]] if msg else []}
    raw = json.dumps(info, ensure_ascii=False, indent=1).encode("utf-8")
    open(os.path.join(OUT, "excel_tv_public.json"), "wb").write(raw)
    key_pem = os.environ.get("EDIT_SIGN_KEY", "").strip()
    if not key_pem:
        sys.exit("서명 키(EDIT_SIGN_KEY)가 없습니다.")
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import padding
    key = serialization.load_pem_private_key(key_pem.encode(), password=None)
    sig = key.sign(raw, padding.PKCS1v15(), hashes.SHA256())
    open(os.path.join(OUT, "excel_tv_public.sig"), "w", encoding="ascii").write(base64.b64encode(sig).decode())
    print(f"배포판 v{ver} · {len(data):,} 바이트 · 서명 완료")


if __name__ == "__main__":
    main()
