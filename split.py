import sys
import os
import shutil
import requests
import pandas as pd
import time
from pathlib import Path
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

from dotenv import load_dotenv
import gmail

load_dotenv()

PROJECT_PATH = os.getenv("PROJECT_PATH")

MAPIA_KEYWORDS_STR = os.getenv("MAPIA_KEYWORDS_STR")

load_dotenv()

PROJECT_PATH = os.getenv("PROJECT_PATH")

MAPIA_KEYWORDS_STR = os.getenv("MAPIA_KEYWORDS_STR")
# 문자열
# MAPIA_KEYWORDS_STR = "롬앤\n에뛰드\n클리오\n헤라\n페리페라\n에스쁘아\n투쿨포스쿨\n3CE\n릴리바이레드\n바닐라코\n퓌\n어뮤즈\n힌스\n하트퍼센트\n무지개맨션\n라카\n파넬\nVDL\n노베브\n삐아\n이니스프리\n컬러그램\n티핏\n루나\n루나컨실러\n루나컨실러팔레트\n루나롱래스팅팁컨실러\n웨이크메이크\n웨이크메이크컨실러\n데이지크\n데이지크컨실러\n데이지크컨실러팔레트\n정샘물\n정샘물컨실러\n정샘물컨실러팔레트\n티핏컨실러\n삐아컨실러\n글로우컨실러\n글로우낫드라이컨실러\n디어에이컨실러\n클리오컨실러\n바비브라운컨실러\n마루빌츠컨실러\n헤라컨실러\n더샘\n더샘컨실러\n더샘\n더샘컨실러\n더샘컨실러펜슬\n더샘팟컨실러\n더샘커버퍼펙션컨실러\n더샘커버퍼펙션팁컨실러\n더샘커버퍼펙션트리플팟컨실러\n더샘커버퍼펙션컨실러펜슬\n더샘커버퍼펙션컨실러쿠션\n더샘트리플팟컨실러\n더샘컨실러팔레트\n더샘파운데이션\n더샘컨실러쿠션\n더샘팁컨실러\n더샘립펜슬\n더샘쿠션\n더샘블러셔\n더샘젤리블러셔\n더샘코렉트베이지\n더샘코렉트업베이지\n더샘선크림\n더샘하이라이터\n더샘입덕주의화이트\n더샘클렌징워터\n더샘세일\n더샘파운데이션밤\n더샘섀도우\n더샘멜팅밤\n더샘프라이머립밤\n더샘입주화\n올리브영더샘\n더샘올리브영\n올리브영더샘컨실러\n올리브영다크서클컨실러\n더샘컨실러팟\n더샘다크서클컨실러\n더샘브라이트너\n더샘피치베이지\n더샘컨실러1.5\n더샘컨실러피치베이지\n더샘컨투어베이지\n더샘펜슬컨실러\n더샘트리플컨실러\n더샘마스크팩\n더샘커버퍼펙션립펜슬\n더샘코렉터\n더샘컨실커버쿠션\n더샘커버쿠션\n더샘컨실러파운데이션\n더샘새미스\n더샘틴트\n더샘에이드샷틴트\n더샘새미스시럽샷멜팅밤\n더샘립밤\n더샘핸드크림\n더샘컬러코렉터\n더샘커버퍼펙션트리플파운데이션밤\n더샘새미스에이드샷틴트\n더샘새미스멜팅밤\n더샘망고피치\n더샘아이라이너\n더샘하라케케\n더샘필링젤\n더샘핑크선크림\n더샘커버퍼펙션\n더샘립글로스\n더샘데저트샌드\n더샘립스틱\n더샘키스홀릭\n입덕주의화이트\n더샘골드볼륨라이트\n더샘샘물싱글섀도우\n더샘입덕주의\n더샘네이키드피치\n더샘오키드루머\n더샘리치캐모마일\n더샘바이올렛진\n더샘향수\n더샘이준호\n이준호더샘\n더샘준호\n더샘새미스시럽샷피치콧\n더샘새미스시럽샷멜팅밤\n더샘멜팅밤피치콧\n더샘피치콧\n더샘립밤\n더샘소프트블러링프라이머립밤\n더샘프라이머립\n더샘포토카드\n이준호포토카드\n이준호키링\n더샘키링\n더샘\n더샘컨실러\n더샘컨실러펜슬\n더샘팟컨실러\n더샘커버퍼펙션컨실러\n더샘커버퍼펙션팁컨실러\n더샘커버퍼펙션트리플팟컨실러\n더샘커버퍼펙션컨실러펜슬\n더샘커버퍼펙션컨실러쿠션\n더샘트리플팟컨실러\n더샘컨실러팔레트\n더샘컨실러쿠션\n더샘파운데이션\n더샘팁컨실러\n더샘립펜슬\n더샘쿠션\n더샘블러셔\n더샘젤리블러셔\n더샘코렉트베이지\n더샘코렉트업베이지\n더샘선크림\n더샘하이라이터\n입덕주의화이트\n더샘싱글섀도우\n더샘새미스에이드샷틴트\n더샘섀도우\n더샘브라이트너\n더샘새미스시럽샷멜팅밤\n더샘멜팅밤\n더샘프라이머립밤\n더샘맨즈케어\n올리브영더샘\n더샘올리브영\n올리브영더샘컨실러\n올리브영다크서클컨실러\n더샘컨실러팟\n더샘다크서클컨실러\n더샘피치베이지\n더샘컨실러1.5\n더샘컨실러피치베이지\n더샘컨투어베이지\n더샘펜슬컨실러\n더샘트리플컨실러\n더샘마스크팩\n더샘파운데이션밤\n더샘커버퍼펙션컨실러팔레트\n더샘커버퍼펙션립펜슬\n더샘코렉터\n더샘입덕주의화이트\n더샘클렌징워터\n더샘핑크선크림\n더샘커버쿠션\n더샘컨실쿠션\n더샘컨실커버쿠션\n더샘컨실러파운데이션\n더샘새미스\n더샘틴트\n더샘에이드샷틴트\n더샘립밤\n더샘핸드크림\n더샘세일\n더샘망고피치\n더샘아이라이너\n더샘하라케케\n더샘필링젤\n더샘할인\n더샘수분크림\n더샘더마플랜\n더샘립앤아이리무버\n더샘립글로스\n더샘데저트샌드\n더샘립스틱\n더샘키스홀릭\n더샘골드볼륨라이트\n더샘샘물싱글섀도우\n더샘네이키드피치\n더샘오키드루머\n더샘리치캐모마일\n더샘바이올렛진\n더샘향수\n더샘이준호\n이준호더샘\n더샘준호\n더샘캔디틴트\n더샘베이비라벤더\n더샘새미스시럽샷피치콧\n더샘멜팅밤피치콧\n더샘피치콧\n더샘커버퍼펙션\n더샘포토카드\n이준호포토카드\n이준호키링\n더샘키링\n더샘립밤\n더샘프라이머립\n더샘소프트블러링립밤\n더샘소프트블러링프라이머립밤\n"

# \n와 , 둘 다 처리하도록 ,를 \n로 변환한 뒤 split
keywords_list = [keyword.strip() for keyword in MAPIA_KEYWORDS_STR.replace(",", "\n").replace("\\n", "\n").split("\n") if keyword.strip()]

# 결과 출력
print(keywords_list)
