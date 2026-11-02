"""DMB 편성 편집기 - TV 편성표 편집기(ui.py)와 같은 프로그램을 DMB 설정으로 실행한다.
설정·자동 저장은 따로(데이터\\DMB), ★ 칸 이름·즐겨찾기 색·뱃지·글꼴은 편성표 편집기와 같이 쓴다."""
import os
import runpy

os.environ["EDIT_PROFILE"] = "dmb"
runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui.py"), run_name="__main__")
