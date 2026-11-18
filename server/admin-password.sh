#!/bin/bash
# 녹음 PC 웹 관리자(http://녹음PC/admin/) 비밀번호 정하기 / 바꾸기
#   bash ~/edit/server/admin-password.sh
python3 "$(cd "$(dirname "$0")" && pwd)/admin.py" --set-password
