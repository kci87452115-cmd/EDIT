#!/bin/bash
# 녹음 PC 서버 설치 / 업데이트 (Debian)
#   본 서버(AM PC):   bash ~/edit/server/install.sh
#   예비 서버(FM PC): bash ~/edit/server/install-backup.sh 본서버IP   (이 파일을 EDIT_ROLE=backup 으로 부름)
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(dirname "$HERE")"
ME="$(whoami)"
ROLE="${EDIT_ROLE:-main}"

if [ "$REPO" != "$HOME/edit" ]; then
    echo "이 저장소는 $HOME/edit 에 있어야 합니다 (지금: $REPO)"; exit 1
fi
if [ "$ROLE" = "backup" ]; then
    REC_DIR="$HOME/Backup/AM"            # 예비 서버: /rec/ = 본 서버에서 가져온 AM 녹음 (최근 며칠)
else
    REC_DIR="$HOME/Recordings"
fi

echo "== 1. 필요한 프로그램"
sudo apt-get install -y python3 python3-tk git nginx rsync >/dev/null

echo "== 2. 녹음 플레이어 위치 확인"
mkdir -p "$REC_DIR" "$HOME/Recordings"
cp "$HERE/player.html" "$REC_DIR/player.html"      # 늘 저장소의 최신 플레이어로
cp "$HERE/player.html" "$HOME/Recordings/player.html"
chmod o+x "$HOME"
[ "$ROLE" = "backup" ] && chmod o+x "$HOME/Backup"

echo "== 3. 접속 기록 (관리자 화면이 nginx 기록을 읽을 수 있게)"
sudo usermod -aG adm "$ME"
sudo touch /var/log/nginx/edit-access.log
sudo chown www-data:adm /var/log/nginx/edit-access.log
sudo chmod 640 /var/log/nginx/edit-access.log

echo "== 3-1. 자료 모으기 서비스 (edit-mirror)"
sed -e "s#^User=rec#User=$ME#" -e "s#/home/rec/#$HOME/#g" "$HERE/edit-mirror.service" \
    | sed "/^Environment=PYTHONUNBUFFERED=1/a Environment=EDIT_REC_DIR=$REC_DIR" \
    | sudo tee /etc/systemd/system/edit-mirror.service >/dev/null

echo "== 3-2. 백업 예약 (설정 백업: 매일 04:00$( [ "$ROLE" = "backup" ] && echo ' · AM 녹음 복사: 매시 10분'))"
for JOB in server sync; do
    sudo tee /etc/systemd/system/edit-backup-$JOB.service >/dev/null <<UNIT
[Unit]
Description=녹음 PC 서버 백업 ($JOB)
After=network-online.target
[Service]
Type=oneshot
User=$ME
ExecStart=/usr/bin/python3 $HERE/backup_job.py $JOB
Nice=19
IOSchedulingClass=idle
TimeoutStartSec=4h
UNIT
done
sudo tee /etc/systemd/system/edit-backup-server.timer >/dev/null <<UNIT
[Unit]
Description=녹음 PC 설정 백업 (매일 04:00)
[Timer]
OnCalendar=*-*-* 04:00:00
RandomizedDelaySec=300
Persistent=true
[Install]
WantedBy=timers.target
UNIT
sudo tee /etc/systemd/system/edit-backup-sync.timer >/dev/null <<UNIT
[Unit]
Description=AM 녹음 최근 며칠 가져오기 (예비 서버, 매시 10분)
[Timer]
OnCalendar=*-*-* *:10:00
Persistent=true
[Install]
WantedBy=timers.target
UNIT
sudo systemctl daemon-reload
sudo systemctl enable edit-mirror >/dev/null 2>&1
sudo systemctl enable --now edit-backup-server.timer >/dev/null 2>&1
if [ "$ROLE" = "backup" ]; then
    sudo systemctl enable --now edit-backup-sync.timer >/dev/null 2>&1
else
    sudo systemctl disable --now edit-backup-sync.timer >/dev/null 2>&1 || true
fi
sudo systemctl restart edit-mirror

echo "== 4. 웹 서버 (nginx)"
sudo mkdir -p /etc/nginx/edit-extra
sed "s#/home/rec/Recordings/#$REC_DIR/#g" "$HERE/nginx-edit.conf" | sudo tee /etc/nginx/sites-available/edit >/dev/null
if [ "$ROLE" = "backup" ]; then
    # 예비 서버: /rec-fm/ = 이 PC의 FM 녹음. 본 서버(AM PC)도 이 주소로 FM 녹음을 가져간다
    sudo rm -f /etc/nginx/edit-extra/fm.conf
    sudo tee /etc/nginx/edit-extra/fm-local.conf >/dev/null <<CONF
# 예비 서버(FM PC)의 FM 녹음 — install.sh (EDIT_ROLE=backup) 가 만듦
location = /rec-fm { return 301 /rec-fm/; }
location = /rec-fm/player.html {
    alias $HOME/Recordings/player.html;
    add_header Cache-Control "no-cache";
}
location ~ "^/rec-fm/(lost\+found|recording_logs|\.)" { return 404; }
location /rec-fm/ {
    alias $HOME/Recordings/;
    autoindex on;
    autoindex_exact_size off;
    autoindex_localtime on;
}
CONF
    # 예전 FM 전용 웹(fm-rec)은 끈다 (같은 80번 포트)
    if [ -e /etc/nginx/sites-enabled/fm-rec ]; then
        sudo mkdir -p /etc/nginx/disabled-sites
        sudo mv /etc/nginx/sites-enabled/fm-rec /etc/nginx/disabled-sites/fm-rec.$(date +%Y%m%d%H%M%S)
    fi
fi
if [ -e /etc/nginx/sites-enabled/recordings ]; then
    sudo mkdir -p /etc/nginx/disabled-sites
    sudo mv /etc/nginx/sites-enabled/recordings /etc/nginx/disabled-sites/recordings.$(date +%Y%m%d%H%M%S)
fi
sudo rm -f /etc/nginx/sites-enabled/default
sudo ln -sf /etc/nginx/sites-available/edit /etc/nginx/sites-enabled/edit
if sudo nginx -t 2>/dev/null; then
    sudo systemctl reload nginx || sudo systemctl restart nginx
else
    sudo nginx -t
    echo "!! nginx 설정 오류 — 위 메시지를 확인해 주세요"; exit 1
fi

sleep 3
echo
echo "== 확인 ($( [ "$ROLE" = "backup" ] && echo '예비 서버' || echo '본 서버'))"
systemctl is-active edit-mirror nginx
curl -s -o /dev/null -w "첫 화면: HTTP %{http_code}\n" http://localhost/
curl -s -o /dev/null -w "녹음 플레이어: HTTP %{http_code}\n" http://localhost/rec/player.html
curl -s -o /dev/null -w "구간 저장: HTTP %{http_code}\n" "http://localhost/api/clip?info=1"
curl -s -o /dev/null -w "관리자 화면: HTTP %{http_code}\n" http://localhost/admin/
[ "$ROLE" = "backup" ] && curl -s -o /dev/null -w "FM 녹음(이 PC): HTTP %{http_code}\n" http://localhost/rec-fm/
if [ ! -f "$HOME/edit-mirror/admin.json" ]; then
    echo
    echo "관리자 비밀번호가 아직 없습니다. 정하려면:  bash ~/edit/server/admin-password.sh"
fi
IP=$(hostname -I | awk '{print $1}')
echo
echo "끝. 사내 PC 브라우저에서:  http://$IP/   (또는 http://$(hostname).local/)"
echo "상태 기록 보기:  journalctl -u edit-mirror -f"
