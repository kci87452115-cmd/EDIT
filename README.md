# 편집 프로그램 (RADIO / TV)

런처 하나로 RADIO·TV 프로그램 중 필요한 것만 설치·업데이트·실행합니다.

| 폴더 | 내용 |
|---|---|
| `apps/runsheet/` | 운행표 데이터 편집기 (`pdf_data_editor.py`, `APP_VERSION`) — RADIO |
| `apps/schedule/` | 편성표 PDF 편집기 (`ui.py`, `APP_VERSION`) — RADIO |
| `launcher/` | 편집 프로그램 런처 (`LAUNCHER_VERSION`, 공용 실행 환경 `RUNTIME`) |
| `installer/` | Inno Setup 설치 스크립트 |
| `tools/build_packages.py` | catalog.json · 프로그램 꾸러미 · 설치 zip (앱 목록·그룹은 여기 `APPS`) |

`main` 에 올리면 GitHub Actions(윈도우)가 빌드해서 Releases 에 올립니다.

- 처음 설치: `EDIT_setup_rN.exe` 또는 `EDIT_install_rN.zip`(압축 풀고 `설치.bat`)
- 업데이트: `EDIT_update_rN.zip` 을 업데이트 폴더(공유 폴더)에 풀거나 런처의 [업데이트 파일 가져오기]
- 저장소가 공개면 런처 업데이트 위치에 이 저장소 주소를 넣어 인터넷으로 바로 받을 수 있음

설치 결과 (`C:\EDIT`)
```
런처\        EditLauncher.exe + 공용 실행 환경
업데이트\    프로그램 꾸러미
radio\       앱 · 데이터 · 백업
TV\          앱 · 데이터 · 백업
```
