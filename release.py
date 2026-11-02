"""
운행표 데이터 편집기 - 설치 파일 만들기 + 업데이트 폴더에 올리기

준비 (이 폴더에 함께 둔다)
  - pdf_data_editor24.py (또는 최신 버전 .py)  ← 가장 번호가 큰 파일을 자동으로 고른다
  - setup.iss
  - 기본데이터.db   (선택) 프로그램 ⚙️ 설정 > [설치용 기본 데이터 파일 만들기]로 만든 파일
  - icon.ico        (선택)
  - Inno Setup 6 설치 (https://jrsoftware.org/isdl.php)

실행: release.bat 더블클릭 (또는  python release.py)
  1) 필요한 파이썬 패키지 설치
  2) PyInstaller 로 프로그램 폴더 만들기 (dist\\RunSheetEditor)
  3) Inno Setup 으로 설치 파일 만들기 (Output\\RunSheetEditor_setup_버전.exe)
  4) (원하면) 업데이트 폴더에 설치 파일 + version.json 올리기 -> 각 PC가 켤 때 새 버전을 알려 줌
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "RunSheetEditor"


def step(t):
    print("\n" + "=" * 60 + f"\n {t}\n" + "=" * 60)


def find_source():
    files = glob.glob(os.path.join(HERE, "pdf_data_editor*.py"))
    if not files:
        sys.exit("pdf_data_editor*.py 파일이 이 폴더에 없습니다.")

    def num(p):
        m = re.search(r"pdf_data_editor(\d*)\.py$", os.path.basename(p))
        return int(m.group(1)) if m and m.group(1) else 0
    return max(files, key=num)


def read_version(src):
    text = open(src, encoding="utf-8").read()
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', text, flags=re.M)
    if not m:
        sys.exit(f"{os.path.basename(src)} 안에서 APP_VERSION 을 찾지 못했습니다.")
    return m.group(1)


def find_iscc():
    cands = [os.path.join(os.environ.get(v, ""), "Inno Setup 6", "ISCC.exe")
             for v in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]
    cands.append(os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"))
    for c in cands:
        if c and os.path.exists(c):
            return c
    return shutil.which("ISCC") or shutil.which("iscc")


def run(cmd):
    print(">", " ".join(f'"{c}"' if " " in c else c for c in cmd))
    r = subprocess.run(cmd, cwd=HERE)
    if r.returncode != 0:
        sys.exit(f"\n실패했습니다 (코드 {r.returncode}). 위 메시지를 확인해 주세요.")


def main():
    os.chdir(HERE)
    src = find_source()
    ver = read_version(src)
    print(f"소스: {os.path.basename(src)}   버전: {ver}")

    step("1/4 필요한 패키지 설치")
    run([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller", "pymupdf", "openpyxl"])

    step("2/4 프로그램 폴더 만들기 (PyInstaller)")
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--onedir",
           "--name", NAME, "--collect-all", "pymupdf"]
    if os.path.exists("icon.ico"):
        cmd += ["--icon", "icon.ico"]
    run(cmd + [src])

    step("3/4 설치 파일 만들기 (Inno Setup)")
    iscc = find_iscc()
    if not iscc:
        sys.exit("Inno Setup 6 을 찾지 못했습니다. https://jrsoftware.org/isdl.php 에서 설치한 뒤 다시 실행해 주세요.")
    if os.path.exists("기본데이터.db"):
        print("기본데이터.db 를 설치 파일에 넣습니다 (처음 설치하는 PC에만 들어갑니다).")
    else:
        print("기본데이터.db 가 없어 빈 데이터로 설치됩니다. (프로그램 ⚙️ 설정에서 만들 수 있습니다)")
    run([iscc, f"/DAppVer={ver}", f"/DSrcDir=dist\\{NAME}", "setup.iss"])
    setup = os.path.join(HERE, "Output", f"{NAME}_setup_{ver}.exe")
    print(f"\n설치 파일: {setup}")

    step("4/4 업데이트 폴더에 올리기 (선택)")
    default_dir = r"C:\RADIO\업데이트"
    folder = input(f"올릴 업데이트 폴더 (Enter = {default_dir}, 건너뛰려면 n): ").strip().strip('"')
    if folder.lower() == "n":
        print("올리지 않았습니다. 설치 파일을 직접 각 PC에서 실행하면 됩니다.")
        return
    folder = folder or default_dir
    notes = input("바뀐 내용 (한 줄로, 업데이트 안내창에 보입니다): ").strip()
    os.makedirs(folder, exist_ok=True)
    shutil.copy2(setup, os.path.join(folder, os.path.basename(setup)))
    tmp = os.path.join(folder, "version.json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:            # 설치 파일을 다 올린 뒤에 version.json 을 바꾼다
        json.dump({"version": ver, "setup": os.path.basename(setup), "notes": notes}, f, ensure_ascii=False, indent=1)
    os.replace(tmp, os.path.join(folder, "version.json"))
    # 예전 설치 파일은 최근 3개만 남긴다
    olds = sorted(glob.glob(os.path.join(folder, f"{NAME}_setup_*.exe")), key=os.path.getmtime)
    for old in olds[:-3]:
        try:
            os.remove(old)
        except OSError:
            pass
    print(f"\n올렸습니다: {folder}\n각 PC에서 프로그램을 켜면 새 버전 {ver} 을 안내합니다.")


if __name__ == "__main__":
    try:
        main()
    finally:
        if os.name == "nt":
            input("\nEnter 를 누르면 창이 닫힙니다.")
