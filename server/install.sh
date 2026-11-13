#!/bin/bash
# 녹음 PC 서버 설치 / 업데이트 (Debian)
#   bash ~/edit/server/install.sh
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(dirname "$HERE")"

if [ "$REPO" != "/home/rec/edit" ]; then
    echo "이 저장소는 /home/rec/edit 에 있어야 합니다 (지금: $REPO)"; exit 1
fi

echo "== 1. 필요한 프로그램"
sudo apt-get install -y python3 python3-tk git nginx >/dev/null

echo "== 2. 녹음 플레이어 위치 확인"
mkdir -p /home/rec/Recordings
if [ ! -f /home/rec/Recordings/player.html ] && [ -f "$HERE/player.html" ]; then
    cp "$HERE/player.html" /home/rec/Recordings/
fi
chmod o+x /home/rec

echo "== 3. 자료 모으기 서비스 (edit-mirror)"
sudo cp "$HERE/edit-mirror.service" /etc/systemd/system/edit-mirror.service
sudo systemctl daemon-reload
sudo systemctl enable edit-mirror >/dev/null 2>&1
sudo systemctl restart edit-mirror

echo "== 4. 웹 서버 (nginx)"
sudo cp "$HERE/nginx-edit.conf" /etc/nginx/sites-available/edit
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
echo "== 확인"
systemctl is-active edit-mirror nginx
curl -s -o /dev/null -w "첫 화면: HTTP %{http_code}\n" http://localhost/
curl -s -o /dev/null -w "녹음 플레이어: HTTP %{http_code}\n" http://localhost/rec/player.html
IP=$(hostname -I | awk '{print $1}')
echo
echo "끝. 사내 PC 브라우저에서:  http://$IP/   (또는 http://$(hostname).local/)"
echo "상태 기록 보기:  journalctl -u edit-mirror -f"
