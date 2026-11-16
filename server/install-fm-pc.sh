#!/bin/bash
# FM 녹음 PC에서 실행: 녹음 폴더를 AM 녹음 PC(웹 서버)에 보여 주는 작은 웹 서버
#   bash ~/edit/server/install-fm-pc.sh
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
REC="$HOME/Recordings"
sudo apt-get install -y nginx >/dev/null
mkdir -p "$REC"
chmod o+x "$HOME"
cp "$HERE/player.html" "$REC/player.html"
sudo tee /etc/nginx/sites-available/fm-rec >/dev/null <<CONF
# FM 녹음 폴더 — 공유기 안쪽(AM 녹음 PC)에서만
server {
    listen 80 default_server;
    listen [::]:80 default_server;
    charset utf-8;
    allow 127.0.0.1; allow ::1; allow 192.168.0.0/24; deny all;
    location = / { return 302 /rec/player.html; }
    location ~ "^/rec/(lost\+found|recording_logs|\.)" { return 404; }
    location /rec/ {
        alias $REC/;
        autoindex on;
        autoindex_exact_size off;
        autoindex_localtime on;
    }
}
CONF
sudo rm -f /etc/nginx/sites-enabled/default
sudo ln -sf /etc/nginx/sites-available/fm-rec /etc/nginx/sites-enabled/fm-rec
sudo nginx -t && (sudo systemctl reload nginx || sudo systemctl restart nginx)
curl -s -o /dev/null -w "FM 녹음 폴더: HTTP %{http_code}\n" http://localhost/rec/
echo "IP: $(hostname -I | awk '{print $1}')  → AM 녹음 PC에서: bash ~/edit/server/add-fm.sh $(hostname -I | awk '{print $1}')"
