#!/bin/bash
# FM 녹음 PC를 '예비 웹서버'로 만들기 (FM PC에서 실행)
#   bash ~/edit/server/install-backup.sh 본서버IP [본서버계정]
#   예) bash ~/edit/server/install-backup.sh 192.168.0.107 rec
# 하는 일
#   1. 본 서버(AM PC)에 비밀번호 없이 접속할 열쇠(SSH 키) 등록 (본 서버 비밀번호를 한 번 물음)
#   2. 이 PC에 본 서버와 같은 웹 서버 설치 (운행표 · TV 편성표 · 관리자 화면)
#      /rec/    = 본 서버에서 가져온 AM 녹음 (최근 2일)   /rec-fm/ = 이 PC의 FM 녹음
#   3. 매시 10분: AM 녹음 최근 2일 복사 (녹음 중인 파일 빼고, 초당 5MB, 가장 낮은 우선순위)
#      매일 04:00: 설정 백업 (서로 백업)
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
AM="$1"; AMU="${2:-rec}"
if ! [[ "$AM" =~ ^[0-9]{1,3}(\.[0-9]{1,3}){3}$ ]]; then echo "사용법: bash install-backup.sh 본서버IP [계정(rec)]"; exit 1; fi

echo "== 1. 본 서버($AMU@$AM) 접속 열쇠"
sudo apt-get install -y rsync openssh-client >/dev/null
[ -f "$HOME/.ssh/id_ed25519" ] || ssh-keygen -t ed25519 -N "" -f "$HOME/.ssh/id_ed25519" -C "edit-backup@$(hostname)" >/dev/null
if ! ssh -o BatchMode=yes -o ConnectTimeout=8 -o StrictHostKeyChecking=accept-new "$AMU@$AM" true 2>/dev/null; then
    echo "   본 서버 '$AMU' 계정 비밀번호를 한 번 넣어 주세요."
    ssh-copy-id -o StrictHostKeyChecking=accept-new -i "$HOME/.ssh/id_ed25519.pub" "$AMU@$AM"
fi
ssh -o BatchMode=yes "$AMU@$AM" true && echo "   열쇠 OK"
if ! ssh -o BatchMode=yes "$AMU@$AM" "command -v rsync" >/dev/null 2>&1; then
    echo "   본 서버에 rsync 설치 (본 서버 sudo 비밀번호)"
    ssh -t "$AMU@$AM" "sudo apt-get install -y rsync"
fi

echo "== 2. 예비 서버 설정 (~/edit-mirror/config.json)"
mkdir -p "$HOME/edit-mirror"
AMCFG="$(ssh -o BatchMode=yes "$AMU@$AM" 'cat ~/edit-mirror/config.json 2>/dev/null' || true)"
AMCFG="$AMCFG" AM="$AM" AMU="$AMU" python3 - <<'PY'
import json, os
p = os.path.expanduser("~/edit-mirror/config.json")
try:
    c = json.load(open(p, encoding="utf-8"))
except Exception:
    c = {}
try:
    am = json.loads(os.environ.get("AMCFG") or "{}")
except Exception:
    am = {}
for k in ("radio", "tv", "links", "ip_names"):      # 편집기 주소 · 바로가기 · IP 이름은 본 서버와 같게
    if k in am and k not in c:
        c[k] = am[k]
c.update({"role": "backup", "peer": os.environ["AM"], "peer_user": os.environ["AMU"]})
c.setdefault("am_keep_days", 2)
c.setdefault("backup_keep", 7)
json.dump(c, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("   ", {k: c[k] for k in ("role", "peer", "peer_user", "am_keep_days")})
PY

echo "== 3. 웹 서버 설치 (예비 서버)"
EDIT_ROLE=backup bash "$HERE/install.sh"

echo "== 4. 첫 AM 녹음 복사 시작 (뒤에서 천천히: 2일치 약 2.6GB, 10분쯤)"
sudo systemctl start --no-block edit-backup-sync.service
sudo systemctl start --no-block edit-backup-server.service
echo
echo "끝. 진행 보기:  journalctl -u edit-backup-sync -f"
echo
echo "!! 본 서버(AM PC)에서 한 번:  cd ~/edit && git pull && bash server/install.sh && bash server/add-fm.sh $(hostname -I | awk '{print $1}')"
echo "   (FM 녹음 주소가 /rec-fm/ 으로 바뀌었습니다)"
