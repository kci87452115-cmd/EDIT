"""
꾸러미 만들기 (GitHub Actions 에서 PyInstaller 다음에 실행)

입력:  dist/RADIOLauncher/ (PyInstaller 결과), apps/*/
출력:  Output/update/             ← 업데이트 폴더(공유 폴더·USB)에 그대로 복사하는 것
           catalog.json
           launcher_<ver>.zip      (런처 + 공용 실행 환경)
           app_<id>_<ver>.zip      (프로그램 파일)
       Output/RADIO_update_<tag>.zip   (위 update 폴더를 하나로 묶은 것 - USB 용)
       Output/RADIO_install_<tag>.zip  (설치 도구 없이 설치: 압축 풀고 설치.bat)
"""
import datetime
import json
import os
import re
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "Output")
UPD = os.path.join(OUT, "update")

APPS = {   # id: (폴더, 실행 파일, 이름, 설명, 필요한 런처 실행 환경 세대)
    "runsheet": ("apps/runsheet", "pdf_data_editor.py", "운행표 데이터 편집기", "운행표 PDF에 광고주 데이터 입력 · 정산 · 일정표", 1),
    "schedule": ("apps/schedule", "ui.py", "편성표 PDF 편집기", "편성표 PDF 편집 (칸 색 · 글자 · 분할선)", 1),
}


def const(path, name):
    text = open(os.path.join(ROOT, path), encoding="utf-8").read()
    m = re.search(rf'^{name}\s*=\s*"?([^"\n]+)"?', text, flags=re.M)
    if not m:
        sys.exit(f"{path} 에서 {name} 를 찾지 못했습니다.")
    return m.group(1).strip()


def zip_dir(src_dir, zpath, arc_prefix=""):
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for base, _d, files in os.walk(src_dir):
            for fn in files:
                if fn.endswith((".pyc",)) and "__pycache__" in base:
                    continue
                full = os.path.join(base, fn)
                z.write(full, os.path.join(arc_prefix, os.path.relpath(full, src_dir)))


INSTALL_BAT = r"""@echo off
rem RADIO launcher installer - extract the zip first, then run this file
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
pause
"""

INSTALL_PS1 = r"""# RADIO 런처 설치: C:\RADIO 에 넣고 바탕화면에 [RADIO 프로그램] 바로가기 하나를 만든다 (관리자 권한 필요 없음)
$ErrorActionPreference = 'Stop'
$root = 'C:\RADIO'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host ''
Write-Host '  RADIO 프로그램 (런처) 설치  ->  ' $root
Write-Host ''
if (-not (Test-Path (Join-Path $src 'RADIOLauncher\RADIOLauncher.exe'))) {
    Write-Host '  [오류] 압축을 먼저 풀어 주세요. (zip 안에서 바로 실행하면 설치할 수 없습니다)'
    exit 1
}
if (Get-Process RADIOLauncher, RunSheetEditor -ErrorAction SilentlyContinue) {
    Write-Host '  RADIO 프로그램이 켜져 있습니다. 모두 닫고 다시 실행해 주세요.'
    exit 1
}
foreach ($d in '런처', '앱', '데이터', '백업', '업데이트') {
    New-Item -ItemType Directory -Force -Path (Join-Path $root $d) | Out-Null
}
robocopy (Join-Path $src 'RADIOLauncher') (Join-Path $root '런처') /MIR /R:5 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw '런처 파일을 복사하지 못했습니다.' }
# 업데이트 꾸러미(프로그램들)를 이 PC의 업데이트 폴더에 넣어 둔다 -> 런처에서 필요한 것만 설치
Copy-Item (Join-Path $src 'update\*') (Join-Path $root '업데이트') -Force
# 기본 데이터 (이 PC에 데이터가 아직 없을 때만)
$seed = Join-Path $src '기본데이터.db'
$data = Join-Path $root '데이터'
if ((Test-Path $seed) -and -not (Test-Path (Join-Path $data 'pdf_data_editor.db')) -and -not (Test-Path (Join-Path $data '기본데이터.db'))) {
    Copy-Item $seed (Join-Path $data '기본데이터.db')
}
# 예전 단독 설치(운행표 편집기) 정리 - 데이터는 그대로
$old = Join-Path $root '프로그램'
if (Test-Path $old) { Remove-Item $old -Recurse -Force -ErrorAction SilentlyContinue }
$desk = [Environment]::GetFolderPath('Desktop')
$progs = [Environment]::GetFolderPath('Programs')
foreach ($p in (Join-Path $desk '운행표 편집기.lnk'), (Join-Path $progs '운행표 편집기.lnk')) {
    if (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
}
# 바로가기는 런처 하나만
$exe = Join-Path $root '런처\RADIOLauncher.exe'
$ws = New-Object -ComObject WScript.Shell
foreach ($place in $desk, $progs) {
    $lnk = $ws.CreateShortcut((Join-Path $place 'RADIO 프로그램.lnk'))
    $lnk.TargetPath = $exe
    $lnk.WorkingDirectory = (Join-Path $root '런처')
    $lnk.IconLocation = "$exe,0"
    $lnk.Save()
}
Write-Host '  설치를 마쳤습니다. 바탕화면의 [RADIO 프로그램]을 실행해 필요한 프로그램을 설치하세요.'
Start-Process $exe -WorkingDirectory (Join-Path $root '런처')
"""

INSTALL_TXT = """RADIO 프로그램 설치 방법 (파이썬 필요 없음)

1. 이 zip 의 압축을 풉니다. (마우스 오른쪽 > 압축 풀기)
2. '설치.bat' 를 더블클릭합니다.
3. C:\\RADIO 에 설치되고 바탕화면에 [RADIO 프로그램] 바로가기가 생깁니다.
4. RADIO 프로그램(런처)에서 필요한 프로그램만 [설치] 를 누릅니다.
   - 운행표 데이터 편집기
   - 편성표 PDF 편집기

· 다시 설치해도 기록·설정은 그대로입니다.
· 윈도우가 '알 수 없는 게시자' 경고를 띄우면 [추가 정보] > [실행].
"""


def crlf_bytes(text, bom=False):
    return (("\ufeff" if bom else "") + text).replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8")


def main():
    tag = sys.argv[1] if len(sys.argv) > 1 else datetime.datetime.now().strftime("%Y%m%d")
    notes = sys.argv[2] if len(sys.argv) > 2 else ""
    shutil.rmtree(UPD, ignore_errors=True)
    os.makedirs(UPD)
    lver = const("launcher/launcher.py", "LAUNCHER_VERSION")
    runtime = int(const("launcher/launcher.py", "RUNTIME").split()[0])
    dist = os.path.join(ROOT, "dist", "RADIOLauncher")
    if not os.path.exists(os.path.join(dist, "RADIOLauncher.exe")) and os.name == "nt":
        sys.exit("dist/RADIOLauncher/RADIOLauncher.exe 가 없습니다. PyInstaller 를 먼저 실행하세요.")
    cat = {"built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "tag": tag,
           "launcher": {"version": lver, "runtime": runtime, "file": f"launcher_{lver}.zip"}, "apps": {}}
    zip_dir(dist, os.path.join(UPD, f"launcher_{lver}.zip"), "RADIOLauncher")
    for app_id, (folder, entry, name, desc, need_rt) in APPS.items():
        ver = const(f"{folder}/{entry}", "APP_VERSION")
        fn = f"app_{app_id}_{ver}.zip"
        zip_dir(os.path.join(ROOT, folder), os.path.join(UPD, fn))
        cat["apps"][app_id] = {"name": name, "desc": desc, "version": ver, "file": fn, "entry": entry,
                               "runtime": need_rt, "notes": notes}
        print(f"  {name} v{ver} -> {fn}")
    with open(os.path.join(UPD, "catalog.json"), "w", encoding="utf-8") as f:
        json.dump(cat, f, ensure_ascii=False, indent=1)
    print(f"  런처 v{lver} (실행 환경 {runtime})")

    # USB·공유 폴더용: update 폴더를 하나로
    upd_zip = os.path.join(OUT, f"RADIO_update_{tag}.zip")
    zip_dir(UPD, upd_zip)
    # 설치 도구 없이 설치하는 zip
    inst_zip = os.path.join(OUT, f"RADIO_install_{tag}.zip")
    with zipfile.ZipFile(inst_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for base, _d, files in os.walk(dist):
            for fn in files:
                full = os.path.join(base, fn)
                z.write(full, os.path.join("RADIOLauncher", os.path.relpath(full, dist)))
        for fn in os.listdir(UPD):
            if not fn.startswith("launcher_"):           # 런처는 이미 RADIOLauncher 폴더로 들어 있다
                z.write(os.path.join(UPD, fn), f"update/{fn}")
        z.writestr("설치.bat", crlf_bytes(INSTALL_BAT))
        z.writestr("install.ps1", crlf_bytes(INSTALL_PS1, bom=True))
        z.writestr("설치방법.txt", crlf_bytes(INSTALL_TXT, bom=True))
    print(f"  {upd_zip}\n  {inst_zip}")
    with open(os.path.join(OUT, "versions.txt"), "w", encoding="utf-8") as f:
        f.write(f"launcher={lver}\n" + "".join(f"{a}={e['version']}\n" for a, e in cat["apps"].items()))


if __name__ == "__main__":
    main()
