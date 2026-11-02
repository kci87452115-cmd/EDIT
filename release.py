"""
운행표 데이터 편집기 - 설치 꾸러미 만들기 + 업데이트 폴더에 올리기
(GitHub 가입·Inno Setup 없이도 됩니다)

준비 (이 폴더에 함께 둔다)
  - pdf_data_editor*.py  ← 번호가 가장 큰 파일을 자동으로 고른다
  - release.py, release.bat
  - 파이썬: 둘 중 하나
      ① WinPython (설치 없이 압축만 풀면 되는 파이썬) 을 이 폴더 안에 풀어 둔다  (예: WPy64-31290\\...)
      ② 또는 PC에 설치된 파이썬
  - (선택) 기본데이터.db : 프로그램 ⚙️ 설정 > [기본 데이터 파일 만들기]
  - (선택) icon.ico
  - (선택) Inno Setup 6 이 설치돼 있으면 setup.exe 도 같이 만든다 (없어도 됨)

결과 (Output 폴더)
  - RunSheetEditor_버전.zip  ← 각 PC: 압축 풀고 '설치.bat' 더블클릭 -> C:\\RADIO 설치 + 바탕화면 바로가기
  - version.json            ← 업데이트 폴더에 zip 과 같이 두면 각 PC가 켤 때 새 버전을 안내
  - (Inno 가 있으면) RunSheetEditor_setup_버전.exe
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "RunSheetEditor"

INSTALL_BAT = r"""@echo off
rem RunSheetEditor installer - run after extracting the zip
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
pause
"""

INSTALL_PS1 = r"""# 운행표 편집기 설치: C:\RADIO 에 넣고 바탕화면 바로가기를 만든다 (관리자 권한 필요 없음)
$ErrorActionPreference = 'Stop'
$root = 'C:\RADIO'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host ''
Write-Host '  운행표 데이터 편집기 설치  ->  ' $root
Write-Host ''
if (-not (Test-Path (Join-Path $src 'RunSheetEditor\RunSheetEditor.exe'))) {
    Write-Host '  [오류] 압축을 먼저 풀어 주세요. (zip 안에서 바로 실행하면 설치할 수 없습니다)'
    exit 1
}
if (Get-Process RunSheetEditor -ErrorAction SilentlyContinue) {
    Write-Host '  운행표 편집기가 켜져 있습니다. 프로그램을 닫고 다시 실행해 주세요.'
    exit 1
}
foreach ($d in '프로그램', '데이터', '백업', '업데이트') {
    New-Item -ItemType Directory -Force -Path (Join-Path $root $d) | Out-Null
}
$prog = Join-Path $root '프로그램'
robocopy (Join-Path $src 'RunSheetEditor') $prog /MIR /R:5 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw '프로그램 파일을 복사하지 못했습니다.' }

# 기본 데이터: 이 PC에 데이터가 아직 없을 때만 넣는다 (있으면 절대 덮어쓰지 않음)
$seed = Join-Path $src '기본데이터.db'
$data = Join-Path $root '데이터'
if ((Test-Path $seed) -and -not (Test-Path (Join-Path $data 'pdf_data_editor.db')) -and -not (Test-Path (Join-Path $data '기본데이터.db'))) {
    Copy-Item $seed (Join-Path $data '기본데이터.db')
    Write-Host '  기본 데이터를 넣었습니다.'
}

$exe = Join-Path $prog 'RunSheetEditor.exe'
$ws = New-Object -ComObject WScript.Shell
foreach ($place in [Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs')) {
    $lnk = $ws.CreateShortcut((Join-Path $place '운행표 편집기.lnk'))
    $lnk.TargetPath = $exe
    $lnk.WorkingDirectory = $prog
    $lnk.IconLocation = "$exe,0"
    $lnk.Save()
}
Write-Host '  설치를 마쳤습니다. 바탕화면의 [운행표 편집기]로 실행하세요.'
Start-Process $exe -WorkingDirectory $prog
"""

INSTALL_TXT = """운행표 데이터 편집기 설치 방법 (파이썬 필요 없음)

1. 이 zip 파일의 압축을 풉니다. (마우스 오른쪽 > 압축 풀기)
2. 풀린 폴더의 '설치.bat' 를 더블클릭합니다.
3. C:\\RADIO 에 설치되고 바탕화면에 [운행표 편집기] 바로가기가 생깁니다.

· 이미 설치된 PC에서 다시 실행하면 프로그램만 새 버전으로 바뀌고, 기록·설정은 그대로입니다.
· 윈도우가 '알 수 없는 게시자' 경고를 띄우면 [추가 정보] > [실행] 을 누르세요.
· 폴더: C:\\RADIO\\프로그램(프로그램) / 데이터(기록·설정) / 백업(자동 백업) / 업데이트
"""


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
    cands = [os.path.join(os.environ.get(v, ""), "Inno Setup 6", "ISCC.exe") for v in ("ProgramFiles(x86)", "ProgramFiles")]
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


def make_zip(ver):
    out = os.path.join(HERE, "Output")
    os.makedirs(out, exist_ok=True)
    path = os.path.join(out, f"{NAME}_{ver}.zip")
    dist = os.path.join(HERE, "dist", NAME)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for base, _dirs, files in os.walk(dist):
            for fn in files:
                full = os.path.join(base, fn)
                z.write(full, os.path.join(NAME, os.path.relpath(full, dist)))
        z.writestr("설치.bat", INSTALL_BAT.replace("\n", "\r\n"))
        z.writestr("install.ps1", ("\ufeff" + INSTALL_PS1).replace("\n", "\r\n").encode("utf-8"))
        z.writestr("설치방법.txt", ("\ufeff" + INSTALL_TXT).replace("\n", "\r\n").encode("utf-8"))
        if os.path.exists(os.path.join(HERE, "기본데이터.db")):
            z.write(os.path.join(HERE, "기본데이터.db"), "기본데이터.db")
    return path


def main():
    os.chdir(HERE)
    src = find_source()
    ver = read_version(src)
    print(f"소스: {os.path.basename(src)}   버전: {ver}   파이썬: {sys.executable}")

    step("1/4 필요한 패키지 설치 (인터넷 필요)")
    run([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller", "pymupdf", "openpyxl"])

    step("2/4 프로그램 만들기 (PyInstaller)")
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--onedir",
           "--name", NAME, "--collect-all", "pymupdf"]
    if os.path.exists("icon.ico"):
        cmd += ["--icon", "icon.ico"]
    run(cmd + [src])

    step("3/4 설치 꾸러미 만들기")
    if os.path.exists("기본데이터.db"):
        print("기본데이터.db 를 넣습니다 (처음 설치하는 PC에만 들어갑니다).")
    else:
        print("기본데이터.db 가 없습니다. (설치한 뒤 ⚙️ 설정 > [기본 데이터 가져오기] 로 넣을 수도 있습니다)")
    zpath = make_zip(ver)
    print(f"설치 zip: {zpath}")
    setup = ""
    iscc = find_iscc()
    if iscc and os.path.exists("setup.iss"):
        run([iscc, f"/DAppVer={ver}", f"/DSrcDir=dist\\{NAME}", "setup.iss"])
        setup = os.path.join(HERE, "Output", f"{NAME}_setup_{ver}.exe")
        print(f"설치 exe: {setup}")
    else:
        print("(Inno Setup 이 없어 setup.exe 는 만들지 않았습니다. zip 으로 충분합니다)")
    info = {"version": ver, "package": os.path.basename(zpath), "notes": ""}
    if setup:
        info["setup"] = os.path.basename(setup)

    step("4/4 업데이트 폴더에 올리기 (선택)")
    default_dir = r"C:\RADIO\업데이트"
    folder = input(f"올릴 업데이트 폴더 (Enter = {default_dir}, 건너뛰려면 n): ").strip().strip('"')
    info["notes"] = "" if folder.lower() == "n" else input("바뀐 내용 (한 줄, 업데이트 안내창에 보입니다): ").strip()
    with open(os.path.join(HERE, "Output", "version.json"), "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=1)
    if folder.lower() == "n":
        print("\n올리지 않았습니다. Output 폴더의 zip 을 각 PC에서 풀고 '설치.bat' 를 실행하면 됩니다.")
        return
    folder = folder or default_dir
    os.makedirs(folder, exist_ok=True)
    for f in [zpath] + ([setup] if setup else []):
        shutil.copy2(f, os.path.join(folder, os.path.basename(f)))
    tmp = os.path.join(folder, "version.json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:            # 파일을 다 올린 뒤에 version.json 을 바꾼다
        json.dump(info, f, ensure_ascii=False, indent=1)
    os.replace(tmp, os.path.join(folder, "version.json"))
    for pat in (f"{NAME}_[0-9]*.zip", f"{NAME}_setup_*.exe"):     # 예전 파일은 최근 3개만 남긴다
        olds = sorted(glob.glob(os.path.join(folder, pat)), key=os.path.getmtime)
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
