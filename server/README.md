# 녹음 PC 서버 (자료 모으기)

라디오 운행표 편집기·TV 편성표 편집기에서 **저장한 내용을 녹음 PC(Debian)에 자동으로 모아**,
편집기 PC가 꺼져 있어도 사내 어느 PC에서나 웹으로 볼 수 있게 합니다. (모두 읽기 전용)

```
[라디오 편집기 PC] 조회 웹 ──┐   사내망 알림(UDP 50808)으로 서로 찾음 — IP가 바뀌어도 됨
[TV 편성표 PC]     조회 웹 ──┤   1분마다 바뀐 것만 가져옴
                              ▼
        녹음 PC  http://녹음PC/
          /        첫 화면
          /radio/  운행표 달력 · 정산 · 검색 · 미리보기 (편집기와 같은 화면)
          /tv/     TV 주간 편성표 PDF
          /rec/    녹음 듣기 (player.html)
```

## 편집기 쪽 준비
- 운행표 편집기 **2.12.0 이상** + ⚙️ 설정에서 **다른 PC 조회 허용** 켜기 (편집기 또는 런처가 켜져 있어야 함)
- TV 편성표 편집기 설정에서 **다른 PC에서 보기** 켜기 (TV PC 런처가 켜져 있어야 함)

## 설치 (녹음 PC, 처음 한 번)
```bash
git clone https://github.com/kci87452115-cmd/EDIT.git ~/edit
bash ~/edit/server/install.sh
```
- 6시간마다 저장소를 새로 받아, 조회 화면이 편집기와 같은 버전으로 유지됩니다.
- 상태 보기: `journalctl -u edit-mirror -f`
- 데이터 위치: `~/edit-mirror/` (radio / tv / state.json)

## 공유기 안쪽에 있거나 사내망 알림이 막힌 경우
알림(UDP)은 공유기를 넘지 못합니다. `~/edit-mirror/config.json` 에 편집기 주소를 직접 적고 `sudo systemctl restart edit-mirror`.
```json
{"radio": ["10.10.5.31:8080"], "tv": ["10.10.5.42:8081"], "public_host": "10.10.5.20"}
```
- `public_host`: 사무실 PC가 녹음 PC 웹에 들어오는 주소 (공유기의 회사망 쪽 주소). 운행표 화면의 'TV 주간 편성표' 탭에 씁니다.
- 공유기에서 녹음 PC에 **DHCP 고정 할당** + **포트 포워딩 80 → 녹음 PC:80** 을 해 두세요.

## 왼쪽 패널 '바로가기' (새 창으로 열림)
`~/edit-mirror/config.json` 에 `links` 를 넣으면 1분 안에 패널에 나타납니다 (서버 다시 시작 필요 없음).
```json
{"links": [{"name": "송출 달력", "url": "http://218.237.3.199:8000/calendar", "icon": "📅"}]}
```

## 업데이트 (Ctrl+F5 필요 없음)
- 녹음 PC 서버가 **1시간마다** GitHub에서 새 버전을 받아 스스로 다시 시작하고, 녹음 플레이어도 새 것으로 바꿉니다.
- 열려 있는 브라우저는 1분 안에 새 버전을 알아채고 **스스로 새로 고칩니다**.
  녹음을 듣는 중이거나 멈춘 지 15분이 안 됐으면 기다렸다가 새로 고칩니다.
- 바로 받고 싶으면: `cd ~/edit && git pull && bash server/install.sh` (nginx 설정이 바뀐 경우에도 이것)

## 관리자 · 접속 기록 (http://녹음PC/admin/)
누가(PC·IP) 언제 무엇을 봤는지 확인합니다. 첫 화면 왼쪽 아래 **🔐 관리자 · 접속 기록**.
1. 처음 한 번 (녹음 PC): `cd ~/edit && git pull && bash server/install.sh` — nginx 가 접속 기록을 남기도록 바꿉니다.
2. 비밀번호 정하기·바꾸기 (녹음 PC): `bash ~/edit/server/admin-password.sh` — 웹에서는 비밀번호를 만들 수 없습니다.
- 보이는 것: 첫 화면 · 라디오 운행표 · 운행표 PDF · TV 편성표 · AM/FM 녹음 듣기 · 녹음 구간 저장 · 관리자 로그인
  (화면이 1분마다 묻는 상태 확인 등은 숨김 — '모든 요청'을 켜면 보임)
- 녹음 듣기는 같은 PC·같은 파일을 30분 안에 다시 받으면 한 줄로 묶어 횟수로 보여 줍니다.
- IP를 누르면 이름(예: 라디오 편집 PC)을 붙일 수 있습니다 (`config.json` 의 `ip_names`).
- 비밀번호를 10분에 5번 틀리면 그 PC는 10분 동안 잠깁니다. 로그인은 12시간 유지.
- 기록 파일: `/var/log/nginx/edit-access.log` (nginx 기본 정리로 14일 보관), 관리자 기록: `~/edit-mirror/admin-events.log`

## 예비 서버 (FM PC) · 백업
FM 녹음 PC를 **예비 웹서버**로: 운행표 · TV 편성표 · 관리자 화면이 본 서버와 같고, AM 녹음은 최근 2일만 들어 있습니다.
1. FM PC: `cd ~/edit && git pull && bash server/install-backup.sh 본서버IP` (본 서버 계정 비밀번호를 한 번 물음)
2. AM PC: `cd ~/edit && git pull && bash server/install.sh && bash server/add-fm.sh FM_PC_IP` (FM 녹음 주소가 /rec-fm/ 으로 바뀜)
- 매시 10분: AM 녹음 최근 2일을 FM PC `~/Backup/AM` 으로 (녹음 중인 파일 빼고, 초당 5MB, 가장 낮은 우선순위)
- 매일 04:00: 설정 백업 zip (웹 설정 · 녹음 프로그램 · 자동 실행 · 운행표 사본 DB · 편집기 설정) — 최근 7개, 서로 백업
- 하루 한 번: 라디오·TV 편집기 PC의 설정(즐겨찾기 · 설정 · 뱃지 · 엑셀 편성표 프로젝트)을 가져와 보관 (바뀐 것만, 14개)
  — 운행표 편집기 2.14.4 · 런처 1.2.4 이상
- 관리자 화면 **💾 백업** 탭: 상태 · 문제 알림 · [지금 복사] [지금 백업] [지금 가져오기] · 백업 파일 받기 · 보관 설정 · 기록
- 백업 zip 의 `복원방법.txt` 에 새 SSD·새 PC에 다시 꾸미는 순서가 있습니다.

## 녹음 구간 저장 (플레이어의 ✂ 구간 저장)
- 시작·끝을 정하면 그 부분만 새 MP3로 내려받습니다. **원본 녹음 파일은 건드리지 않습니다.**
- 구간 고르기: 시간 입력 · [지금 위치로] · −1초/+1초 · 진행바에서 Shift+드래그 · 노란 구간 양 끝 끌기
- 앞뒤 부드럽게(0.3초 페이드)가 기본으로 켜져 있습니다. 듣는 볼륨을 100%보다 올리면 '소리 크기'가 비슷한 칸으로 맞춰집니다.
- 다른 파일을 열면 구간 저장 창은 닫히고 처음 상태로 돌아갑니다.
- 소리 크기: 그대로(다시 압축 안 함, 1초 안에 끝남) / 조금 크게(+6dB) / 크게(+12dB) / 자동 맞춤(방송 기준 -16 LUFS)
- 녹음에 영향이 없도록 가장 낮은 우선순위로 처리하고, 길이·동시 처리 수를 제한합니다.
  `~/edit-mirror/config.json` 에서 바꿀 수 있습니다 (서버 다시 시작 필요 없음).
```json
{"clip_max_min": 30, "clip_workers": 1}
```
- 같은 구간을 10분 안에 다시 받으면 만들어 둔 것을 바로 줍니다.

## FM 녹음 PC 연결 (AM 녹음 PC 웹에서 FM 녹음도 듣기)
1. FM 녹음 PC: `git clone https://github.com/kci87452115-cmd/EDIT.git ~/edit && bash ~/edit/server/install-fm-pc.sh`
2. AM 녹음 PC(웹 서버): `cd ~/edit && git pull && bash server/install.sh && bash server/add-fm.sh FM_PC_IP`
3. 첫 화면 왼쪽에 **📻 FM 녹음 듣기**가 생깁니다 (FM PC가 꺼져 있으면 메뉴가 숨겨짐).
