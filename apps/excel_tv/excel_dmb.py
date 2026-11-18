"""엑셀 DMB 편성표 편집기 - 엑셀 TV 편성표 편집기(tv_schedule_maker.py)와 같은 프로그램을 DMB 설정으로 실행한다.
TV 프로젝트(.tvproj)를 불러와 DMB 모양(연령만 · 음영/굵기 빼기 등)으로 고치고, DMB 작업은 .dmbproj 로 따로 저장.
설정·최근 목록은 따로(dmb_schedule_config.json), 이름 사전·글꼴은 TV 편집기와 같이 쓴다."""
import os
import runpy

os.environ["EXCEL_TV_PROFILE"] = "dmb"
runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "tv_schedule_maker.py"), run_name="__main__")
