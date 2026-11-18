#!/usr/bin/env python3
"""녹음 PC 서버 백업 작업 (systemd 타이머·관리자 화면 버튼이 부름)

  python3 backup_job.py server   설정 백업: 이 PC를 다시 꾸밀 때 필요한 것을 zip 하나로 (최근 7개)
                                 + 예비 서버면 본 서버로도 복사 (서로 백업)
  python3 backup_job.py sync     예비 서버(FM PC)만: 본 서버(AM PC)의 녹음 최근 N일 + 본 서버 설정 백업을 가져옴

설정: ~/edit-mirror/config.json
  "role": "backup"          예비 서버 (없으면 본 서버)
  "peer": "192.168.0.107"   상대 PC 주소 (예비 서버에서: 본 서버 주소)
  "peer_user": "rec"        상대 PC 계정
  "am_keep_days": 2         가져와 둘 AM 녹음 날짜 수
  "backup_keep": 7          설정 백업 보관 개수
녹음에 영향이 없도록 가장 낮은 우선순위(nice 19, ionice idle)로, 복사는 초당 5MB 로 제한한다.
"""
import datetime
import fcntl
import glob
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import zipfile

HOME_DIR = os.path.expanduser("~")
MIRROR = os.environ.get("EDIT_MIRROR_HOME") or os.path.join(HOME_DIR, "edit-mirror")
CONFIG_FILE = os.path.join(MIRROR, "config.json")
BK = os.path.join(MIRROR, "backup")
STATUS_FILE = os.path.join(BK, "status-jobs.json")
AM_COPY = os.environ.get("EDIT_AM_COPY") or os.path.join(HOME_DIR, "Backup", "AM")
LOW = ["nice", "-n", "19", "ionice", "-c3"]
BWLIMIT = "5000"           # KB/s
SSH_OPTS = ["-o", "BatchMode=yes", "-o", "ConnectTimeout=8", "-o", "StrictHostKeyChecking=accept-new"]


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, type(default)) else default
    except Exception:
        return default


def save(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def status_set(key, **kw):
    st = load(STATUS_FILE, {})
    st.setdefault(key, {}).update(kw)
    hist = st.setdefault("history", [])
    if "ok" in kw:
        hist.insert(0, {"t": now(), "job": key, "ok": kw["ok"], "msg": kw.get("msg") or kw.get("err") or "",
                        "sec": kw.get("sec", 0), "size": kw.get("size", 0)})
        cut = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
        st["history"] = [h for h in hist if h["t"] >= cut][:300]
    save(STATUS_FILE, st)


def cfg():
    return load(CONFIG_FILE, {})


def lock(name):
    os.makedirs(BK, exist_ok=True)
    f = open(os.path.join(BK, f".{name}.lock"), "w")
    try:
        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        print("이미 실행 중입니다.")
        sys.exit(0)
    return f


def run(cmd, timeout=3600):
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout or f"코드 {r.returncode}").strip().splitlines()[-1][:300])
    return r.stdout


def keep_newest(folder, pattern, n):
    fs = sorted(glob.glob(os.path.join(folder, pattern)), key=os.path.getmtime, reverse=True)
    for f in fs[n:]:
        try:
            os.remove(f)
        except OSError:
            pass


# ---------------------------------------------------------------- 설정 백업
SERVER_FILES = [
    ("edit-mirror/config.json", os.path.join(MIRROR, "config.json")),
    ("edit-mirror/admin.json", os.path.join(MIRROR, "admin.json")),
    ("edit-mirror/state.json", os.path.join(MIRROR, "state.json")),
    ("system/fstab", "/etc/fstab"),
]
SERVER_GLOBS = [
    ("nginx/sites-available", "/etc/nginx/sites-available/*"),
    ("nginx/edit-extra", "/etc/nginx/edit-extra/*.conf"),
    ("systemd", "/etc/systemd/system/edit-*"),
    ("autostart", os.path.join(HOME_DIR, ".config", "autostart", "*.desktop")),
    ("radio", os.path.join(HOME_DIR, "radio", "*.py")),
    ("radio", os.path.join(HOME_DIR, "radio", "*.json")),
    ("radio", os.path.join(HOME_DIR, "radio", "*.conf")),
    ("radio", os.path.join(HOME_DIR, "radio", "*.ini")),
    ("radio", os.path.join(HOME_DIR, "radio", "*.txt")),
    ("radio", os.path.join(HOME_DIR, "Desktop", "*.desktop")),
]


def _restore_txt(host):
    return f"""녹음 PC 설정 백업 ({host}, {now()})

새 SSD·새 PC에 다시 꾸밀 때 (Debian 13 설치, 같은 계정 이름으로):
 1. git clone https://github.com/kci87452115-cmd/EDIT.git ~/edit
 2. 이 zip 을 풀어 둔다:  mkdir ~/복원 && cd ~/복원 && unzip 이zip파일.zip
 3. 녹음 프로그램:  mkdir -p ~/radio && cp radio/* ~/radio/   (자동 실행: autostart/*.desktop -> ~/.config/autostart/)
 4. 자료 모으기 설정:  mkdir -p ~/edit-mirror && cp edit-mirror/*.json ~/edit-mirror/
    운행표 사본 DB:     mkdir -p ~/edit-mirror/radio/데이터 && cp edit-mirror/radio/pdf_data_editor.db ~/edit-mirror/radio/데이터/
 5. 웹 서버:  bash ~/edit/server/install.sh   (예비 서버면 install-backup.sh 본서버IP)
    FM 연결:  nginx/edit-extra/fm.conf 를 보고 bash ~/edit/server/add-fm.sh FM_IP
 6. 디스크 자동 연결: system/fstab 의 녹음 디스크 줄을 /etc/fstab 에 옮겨 적기 (UUID 확인: lsblk -f)
 7. 크론: system/crontab.txt 를 보고 crontab -e
편집기 설정 백업(editors/…zip)은 편집기 PC의 C:\\EDIT\\radio\\데이터, C:\\EDIT\\TV\\데이터 에 풀면 됩니다.
"""


def job_server():
    t0 = time.time()
    host = socket.gethostname()
    out_dir = os.path.join(BK, "server")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"{host}_{datetime.datetime.now():%Y%m%d-%H%M}.zip")
    tmp = out + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for arc, src in SERVER_FILES:
                if os.path.isfile(src) and os.access(src, os.R_OK):
                    z.write(src, arc)
            for arc_dir, pat in SERVER_GLOBS:
                for src in glob.glob(pat):
                    if os.path.isfile(src) and os.access(src, os.R_OK) and os.path.getsize(src) < 20 * 1024 * 1024:
                        z.write(src, f"{arc_dir}/{os.path.basename(src)}")
            db = os.path.join(MIRROR, "radio", "데이터", "pdf_data_editor.db")
            if os.path.isfile(db):
                snap = db + ".backup-tmp"
                try:                                        # 쓰는 중이어도 안전하게 복사 (sqlite backup)
                    import sqlite3
                    src_c, dst_c = sqlite3.connect(db), sqlite3.connect(snap)
                    with dst_c:
                        src_c.backup(dst_c)
                    src_c.close()
                    dst_c.close()
                    z.write(snap, "edit-mirror/radio/pdf_data_editor.db")
                finally:
                    if os.path.exists(snap):
                        os.remove(snap)
            for kind in ("radio", "tv"):                    # 편집기 설정 백업 (가장 최근 것)
                fs = sorted(glob.glob(os.path.join(BK, "editors", kind, "*.zip")), key=os.path.getmtime)
                if fs:
                    z.write(fs[-1], f"editors/{kind}/{os.path.basename(fs[-1])}")
            try:
                z.writestr("system/crontab.txt", subprocess.run(["crontab", "-l"], capture_output=True, text=True,
                                                                timeout=10).stdout)
            except Exception:
                pass
            z.writestr("복원방법.txt", _restore_txt(host))
        os.replace(tmp, out)
    except Exception as exc:
        if os.path.exists(tmp):
            os.remove(tmp)
        status_set("server", at=now(), ok=False, err=str(exc), sec=round(time.time() - t0, 1))
        raise
    keep = int(cfg().get("backup_keep", 7) or 7)
    keep_newest(out_dir, "*.zip", keep)
    size = os.path.getsize(out)
    msg = f"설정 백업 {os.path.basename(out)} ({size // 1024} KB)"
    # 예비 서버: 본 서버에도 복사 (서로 백업)
    c = cfg()
    if c.get("role") == "backup" and c.get("peer"):
        try:
            peer = f"{c.get('peer_user', 'rec')}@{c['peer']}"
            run(["ssh"] + SSH_OPTS + [peer, "mkdir -p ~/edit-mirror/backup/peer"], 60)
            run(LOW + ["rsync", "-a", f"--bwlimit={BWLIMIT}", "-e", "ssh " + " ".join(SSH_OPTS),
                       out_dir + "/", f"{peer}:edit-mirror/backup/peer/"], 600)
            run(["ssh"] + SSH_OPTS + [peer, f"ls -1t ~/edit-mirror/backup/peer/{host}_*.zip 2>/dev/null | "
                                            f"tail -n +{keep + 1} | xargs -r rm -f"], 60)
            msg += " · 본 서버에도 복사"
        except Exception as exc:
            msg += f" · 본 서버로 복사 못함 ({exc})"
    status_set("server", at=now(), ok=True, file=os.path.basename(out), size=size, msg=msg,
               sec=round(time.time() - t0, 1))
    print(msg)


# ---------------------------------------------------------------- 예비 서버: AM 녹음 최근 N일 가져오기
def job_sync():
    t0 = time.time()
    c = cfg()
    if c.get("role") != "backup" or not c.get("peer"):
        print("예비 서버가 아닙니다 (config.json role/peer).")
        return
    days = max(1, min(30, int(c.get("am_keep_days", 2) or 2)))
    peer = f"{c.get('peer_user', 'rec')}@{c['peer']}"
    ssh = ["ssh"] + SSH_OPTS + [peer]
    try:
        out = run(ssh + ["ls -1d ~/Recordings/20[0-9][0-9]-[0-9][0-9]-[0-9][0-9] 2>/dev/null"], 60)
        dates = sorted(os.path.basename(x.strip()) for x in out.splitlines()
                       if re.fullmatch(r".*/\d{4}-\d{2}-\d{2}", x.strip()))[-days:]
        if not dates:
            raise RuntimeError("본 서버에 녹음 날짜 폴더가 없습니다")
        # 다 녹음된 파일만 (2분 넘게 바뀌지 않은 mp3) - 지금 녹음 중인 파일은 건너뜀
        lst = run(ssh + ["cd ~/Recordings && find " + " ".join(dates) + " -maxdepth 1 -type f -name '*.mp3' -mmin +2"],
                  120)
        files = sorted(x.strip() for x in lst.splitlines() if x.strip())
        os.makedirs(AM_COPY, exist_ok=True)
        flist = os.path.join(BK, ".sync-files.txt")
        with open(flist, "w", encoding="utf-8") as f:
            f.write("\n".join(files) + "\n")
        before = _tree_size(AM_COPY)
        run(LOW + ["rsync", "-a", "--partial", f"--bwlimit={BWLIMIT}", "--files-from=" + flist,
                   "-e", "ssh " + " ".join(SSH_OPTS), f"{peer}:Recordings/", AM_COPY + "/"], 3 * 3600)
        # 가장 최근 N일만 남김 (본 서버가 멈춰도 마지막 N일은 지우지 않음)
        local = sorted(d for d in os.listdir(AM_COPY) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d))
        for d in local[:-days] if len(local) > days else []:
            shutil.rmtree(os.path.join(AM_COPY, d), ignore_errors=True)
        player = os.path.join(os.path.dirname(os.path.abspath(__file__)), "player.html")
        if os.path.exists(player):
            shutil.copy(player, os.path.join(AM_COPY, "player.html"))
        # 본 서버 설정 백업도 가져옴
        try:
            os.makedirs(os.path.join(BK, "peer"), exist_ok=True)
            run(LOW + ["rsync", "-a", f"--bwlimit={BWLIMIT}", "-e", "ssh " + " ".join(SSH_OPTS),
                       f"{peer}:edit-mirror/backup/server/", os.path.join(BK, "peer") + "/"], 600)
            keep_newest(os.path.join(BK, "peer"), "*.zip", int(c.get("backup_keep", 7) or 7) * 2)
        except Exception:
            pass
        got = _tree_size(AM_COPY) - before
        per = {d: [len([f for f in os.listdir(os.path.join(AM_COPY, d)) if f.endswith(".mp3")]),
                   _tree_size(os.path.join(AM_COPY, d))]
               for d in sorted(os.listdir(AM_COPY)) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)}
        newest = max((os.path.getmtime(os.path.join(AM_COPY, f)) for f in files
                      if os.path.exists(os.path.join(AM_COPY, f))), default=0)
        status_set("sync", at=now(), ok=True, last_ok=now(), dates=per, days=days, got=max(0, got),
                   newest_file=datetime.datetime.fromtimestamp(newest).strftime("%Y-%m-%d %H:%M") if newest else "",
                   msg=f"AM 녹음 {len(files)}개 확인 · 새로 {max(0, got) // (1024 * 1024)} MB", err="",
                   sec=round(time.time() - t0, 1), size=max(0, got))
        print("복사 끝", per)
    except Exception as exc:
        status_set("sync", at=now(), ok=False, err=str(exc), msg="", sec=round(time.time() - t0, 1))
        print("복사 실패:", exc)
        sys.exit(1)


def _tree_size(d):
    n = 0
    for base, _ds, fs in os.walk(d):
        for f in fs:
            try:
                n += os.path.getsize(os.path.join(base, f))
            except OSError:
                pass
    return n


if __name__ == "__main__":
    what = sys.argv[1] if len(sys.argv) > 1 else ""
    if what not in ("server", "sync"):
        sys.exit("사용법: backup_job.py server | sync")
    _lk = lock(what)
    (job_server if what == "server" else job_sync)()
