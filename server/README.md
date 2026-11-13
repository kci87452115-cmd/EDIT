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

## 사내망 알림이 막힌 경우
`~/edit-mirror/config.json` 에 편집기 주소를 직접 적습니다.
```json
{"radio": ["192.168.0.50:8080"], "tv": ["192.168.0.60:8081"]}
```
