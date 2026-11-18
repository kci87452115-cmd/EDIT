#!/bin/bash
# AM 녹음 PC(본 서버)에서 실행: FM 녹음 PC의 녹음을 /rec-fm/ 으로 연결
#   bash ~/edit/server/add-fm.sh 192.168.0.106
# FM PC가 예비 서버(install-backup.sh)면 FM 녹음은 FM PC의 /rec-fm/, 아니면 /rec/ 에 있다 (자동으로 고름)
set -e
IP="$1"
if ! [[ "$IP" =~ ^[0-9]{1,3}(\.[0-9]{1,3}){3}$ ]]; then echo "사용법: bash add-fm.sh FM_PC_IP"; exit 1; fi
if [ "$(curl -s -o /dev/null -w '%{http_code}' "http://$IP/rec-fm/" --max-time 5)" = "200" ]; then
    SRC="rec-fm"; echo "FM PC는 예비 서버입니다 (FM 녹음: /rec-fm/)"
else
    SRC="rec"
fi
sudo mkdir -p /etc/nginx/edit-extra
sudo tee /etc/nginx/edit-extra/fm.conf >/dev/null <<CONF
# FM 녹음 PC ($IP) — add-fm.sh 가 만듦
location = /rec-fm { return 301 /rec-fm/; }
location = /rec-fm/player.html {
    alias $HOME/Recordings/player.html;
    add_header Cache-Control "no-cache";
}
location /rec-fm/ {
    proxy_pass http://$IP/$SRC/;
    proxy_connect_timeout 3s;
    proxy_read_timeout 120s;
    proxy_buffering off;
}
CONF
sudo nginx -t && sudo systemctl reload nginx
curl -s -o /dev/null -w "FM 녹음 연결: HTTP %{http_code}\n" http://localhost/rec-fm/
