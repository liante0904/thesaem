import sys
import os
import shutil
import pandas as pd
import time
from pathlib import Path
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from openpyxl.styles import Border, Side, Alignment
from dotenv import load_dotenv

import gmail

load_dotenv()

PROJECT_PATH = os.getenv("PROJECT_PATH")

def generate_naver_keyword_excel(page):
    print("*************네이버 키워드 광고 엑셀(마피아닷컴)*************")
    
    page.goto("https://www.ma-pia.net/keyword/keyword.php")
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").click()
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").fill("롬앤\n에뛰드\n클리오\n헤라\n페리페라\n에스쁘아\n투쿨포스쿨\n3CE\n릴리바이레드\n바닐라코\n퓌\n어뮤즈\n힌스\n하트퍼센트\n무지개맨션\n라카\n파넬\nVDL\n노베브\n삐아\n이니스프리\n컬러그램\n티핏\n루나\n루나컨실러\n루나컨실러팔레트\n루나롱래스팅팁컨실러\n웨이크메이크\n웨이크메이크컨실러\n데이지크\n데이지크컨실러\n데이지크컨실러팔레트\n정샘물\n정샘물컨실러\n정샘물컨실러팔레트\n티핏컨실러\n삐아컨실러\n글로우컨실러\n글로우낫드라이컨실러\n디어에이컨실러\n클리오컨실러\n바비브라운컨실러\n마루빌츠컨실러\n헤라컨실러\n더샘\n더샘컨실러\n더샘\n더샘컨실러\n더샘컨실러펜슬\n더샘팟컨실러\n더샘커버퍼펙션컨실러\n더샘커버퍼펙션팁컨실러\n더샘커버퍼펙션트리플팟컨실러\n더샘커버퍼펙션컨실러펜슬\n더샘커버퍼펙션컨실러쿠션\n더샘트리플팟컨실러\n더샘컨실러팔레트\n더샘파운데이션\n더샘컨실러쿠션\n더샘팁컨실러\n더샘립펜슬\n더샘쿠션\n더샘블러셔\n더샘젤리블러셔\n더샘코렉트베이지\n더샘코렉트업베이지\n더샘선크림\n더샘하이라이터\n더샘입덕주의화이트\n더샘클렌징워터\n더샘세일\n더샘파운데이션밤\n더샘섀도우\n더샘멜팅밤\n더샘프라이머립밤\n더샘입주화\n올리브영더샘\n더샘올리브영\n올리브영더샘컨실러\n올리브영다크서클컨실러\n더샘컨실러팟\n더샘다크서클컨실러\n더샘브라이트너\n더샘피치베이지\n더샘컨실러1.5\n더샘컨실러피치베이지\n더샘컨투어베이지\n더샘펜슬컨실러\n더샘트리플컨실러\n더샘마스크팩\n")
    page.get_by_role("button", name="조회하기").click()
    page.locator("div").filter(has_text="조회중입니다. 잠시만 기다려주세요").nth(1).click()
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").click()
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").fill("더샘커버퍼펙션립펜슬\n더샘코렉터\n더샘컨실커버쿠션\n더샘커버쿠션\n더샘컨실러파운데이션\n더샘새미스\n더샘틴트\n더샘에이드샷틴트\n더샘새미스시럽샷멜팅밤\n더샘립밤\n더샘핸드크림\n더샘컬러코렉터\n더샘커버퍼펙션트리플파운데이션밤\n더샘새미스에이드샷틴트\n더샘새미스멜팅밤\n더샘망고피치\n더샘아이라이너\n더샘하라케케\n더샘필링젤\n더샘핑크선크림\n더샘커버퍼펙션\n더샘립글로스\n더샘데저트샌드\n더샘립스틱\n더샘키스홀릭\n입덕주의화이트\n더샘골드볼륨라이트\n더샘샘물싱글섀도우\n더샘입덕주의\n더샘네이키드피치\n더샘오키드루머\n더샘리치캐모마일\n더샘바이올렛진\n더샘향수\n더샘이준호\n이준호더샘\n더샘준호\n더샘새미스시럽샷피치콧\n더샘새미스시럽샷멜팅밤\n더샘멜팅밤피치콧\n더샘피치콧\n더샘립밤\n더샘소프트블러링프라이머립밤\n더샘프라이머립\n더샘포토카드\n이준호포토카드\n이준호키링\n더샘키링\n더샘\n더샘컨실러\n더샘컨실러펜슬\n더샘팟컨실러\n더샘커버퍼펙션컨실러\n더샘커버퍼펙션팁컨실러\n더샘커버퍼펙션트리플팟컨실러\n더샘커버퍼펙션컨실러펜슬\n더샘커버퍼펙션컨실러쿠션\n더샘트리플팟컨실러\n더샘컨실러팔레트\n더샘컨실러쿠션\n더샘파운데이션\n더샘팁컨실러\n더샘립펜슬\n더샘쿠션\n더샘블러셔\n더샘젤리블러셔\n더샘코렉트베이지\n더샘코렉트업베이지\n더샘선크림\n더샘하이라이터\n입덕주의화이트\n더샘싱글섀도우\n더샘새미스에이드샷틴트\n더샘섀도우\n더샘브라이트너\n더샘새미스시럽샷멜팅밤\n더샘멜팅밤\n더샘프라이머립밤\n더샘맨즈케어\n올리브영더샘\n더샘올리브영\n올리브영더샘컨실러\n올리브영다크서클컨실러\n더샘컨실러팟\n더샘다크서클컨실러\n더샘피치베이지\n더샘컨실러1.5\n더샘컨실러피치베이지\n더샘컨투어베이지\n더샘펜슬컨실러\n")
    page.get_by_role("button", name="조회하기").click()
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").click()
    page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)").fill("더샘트리플컨실러\n더샘마스크팩\n더샘파운데이션밤\n더샘커버퍼펙션컨실러팔레트\n더샘커버퍼펙션립펜슬\n더샘코렉터\n더샘입덕주의화이트\n더샘클렌징워터\n더샘핑크선크림\n더샘커버쿠션\n더샘컨실쿠션\n더샘컨실커버쿠션\n더샘컨실러파운데이션\n더샘새미스\n더샘틴트\n더샘에이드샷틴트\n더샘립밤\n더샘핸드크림\n더샘세일\n더샘망고피치\n더샘아이라이너\n더샘하라케케\n더샘필링젤\n더샘할인\n더샘수분크림\n더샘더마플랜\n더샘립앤아이리무버\n더샘립글로스\n더샘데저트샌드\n더샘립스틱\n더샘키스홀릭\n더샘골드볼륨라이트\n더샘샘물싱글섀도우\n더샘네이키드피치\n더샘오키드루머\n더샘리치캐모마일\n더샘바이올렛진\n더샘향수\n더샘이준호\n이준호더샘\n더샘준호\n더샘캔디틴트\n더샘베이비라벤더\n더샘새미스시럽샷피치콧\n더샘멜팅밤피치콧\n더샘피치콧\n더샘커버퍼펙션\n더샘포토카드\n이준호포토카드\n이준호키링\n더샘키링\n더샘립밤\n더샘프라이머립\n더샘소프트블러링립밤\n더샘소프트블러링프라이머립밤\n")
    
    # 입력 필드 클릭 및 값 입력
    input_field = page.get_by_placeholder("한 줄에 하나씩 입력하세요.\r\n(최대100개까지)")
    input_field.click()
    print("입력 필드를 클릭했습니다.")
    
    # 조회 버튼 클릭
    page.get_by_role("button", name="조회하기").click()
    print("조회하기 버튼을 클릭했습니다.")

    # 로딩 완료까지 대기
    wait_for_loading_to_complete(page)
    
    # 최대 10번 반복
    max_attempts = 10
    attempts = 0

    # 다운로드 및 엑셀 폴더 경로
    downloads_folder = os.path.join(PROJECT_PATH, "downloads")

    # 다운로드 폴더 확인 및 생성
    ensure_directory_exists(downloads_folder)

    attempts = 0
    max_attempts=30
    interval=10
    while attempts < max_attempts:
        count = count_rows_in_table(page)  # 테이블의 레코드 수를 계산하는 함수
        print(f'Table has {count} rows.')

        # 200개가 넘으면 전체 체크 후 파일 다운로드
        if count > 200:
            # 전체 체크 후 파일 다운로드
            check_all_rows(page)

            # 다운로드 버튼 클릭
            with page.expect_download() as download_info:
                download_button = page.locator('xpath=/html/body/div[2]/section[1]/div[4]/div[1]/button[1]')
                download_button.click()
                print("다운로드 버튼을 클릭했습니다.")
            download = download_info.value

            # 다운로드된 파일 경로
            downloaded_file_path = download.path()  # 다운로드된 파일의 경로
            print(f"다운로드된 파일 경로: {downloaded_file_path}")

            # 새 파일 이름 생성
            new_file_name = f"더샘_브랜드검색_키워드쿼리_{datetime.now().strftime('%Y%m%d')}.xlsx"
            new_file_path = os.path.join(downloads_folder, new_file_name)

            # 파일 이름 변경 및 이동
            os.rename(downloaded_file_path, new_file_path)
            print(f"파일 이름이 '{new_file_name}'으로 변경되어 '{downloads_folder}'로 이동되었습니다.")

            # 저장할 경로 설정
            save_directory = os.path.join(PROJECT_PATH, "excel")  # 엑셀 폴더 경로

            # 디렉토리가 존재하지 않으면 생성
            os.makedirs(save_directory, exist_ok=True)

            # 파일 복사
            copy_file_path = os.path.join(save_directory, new_file_name)
            shutil.copy(new_file_path, copy_file_path)  # 파일 복사
            print(f"파일이 '{save_directory}'로 복사되었습니다.")
            break  # 다운로드가 완료되면 루프 종료
        
        # 200개가 안 되면 10초 대기 후 다시 호출
        print(f"Not enough rows. Waiting for {interval} seconds...")
        time.sleep(interval)
        attempts += 1

    if attempts == max_attempts:
        print("최대 대기 시간 5분이 초과되었습니다. 다운로드를 수행할 수 없습니다.")

def wait_for_loading_to_complete(page):
    # 로딩 표시를 감지할 CSS 선택자 설정
    loading_element = page.locator('div[style*="display: block"]')

    # 로딩이 시작될 때까지 대기 (display: block)
    print("로딩 시작을 기다리는 중...")
    loading_element.wait_for(state="visible", timeout=30000)  # 최대 10초 대기
    print("로딩 표시가 보였습니다.")

    # 로딩이 끝나서 사라질 때까지 대기 (display: none)
    loading_element.wait_for(state="hidden", timeout=30000)  # 최대 30초 대기
    print("로딩이 완료되었습니다.")

def check_all_rows(page):
    # 전체 체크박스를 선택
    checkbox = page.locator('xpath=/html/body/div[2]/section[1]/div[4]/div[2]/table/thead/tr[1]/th[1]/input')
    try:
        checkbox.wait_for(timeout=5000)  # 최대 5초 대기
        checkbox.check()
        print("전체 체크 박스를 클릭했습니다.")
    except TimeoutError:
        print("체크박스 클릭 실패. 재시도 중...")
        time.sleep(2)
        check_all_rows(page)

def count_rows_in_table(page):
    # XPath를 사용하여 tbody의 하위 tr 요소 선택
    rows = page.locator('xpath=//*[@id="mytable2"]/tbody/tr')
    
    # tr 요소의 개수 반환
    row_count = rows.count()
    return row_count

def ensure_directory_exists(directory):
    """주어진 디렉토리가 존재하지 않으면 생성합니다."""
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"디렉토리 '{directory}'가 생성되었습니다.")

def get_files_from_directory(directory, extension=None):
    """
    특정 디렉토리에서 모든 파일을 가져오거나 특정 확장자의 파일만 가져옴.
    파일이 없는 경우 빈 리스트를 반환함.
    
    :param directory: 검색할 디렉토리 경로
    :param extension: 특정 파일 확장자 (예: '.txt'). None일 경우 모든 파일을 반환.
    :return: 해당 디렉토리의 파일 경로 리스트. 파일이 없으면 빈 리스트.
    """
    files = []
    
    # 디렉토리 내 파일 목록 확인
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        
        # 파일이면서 확장자가 조건에 맞는 경우 리스트에 추가
        if os.path.isfile(file_path) and (extension is None or file.endswith(extension)):
            files.append(file_path)
    
    # 파일이 없는 경우 빈 리스트 반환
    return files  
        
def setup_directories(project_path):
    """
    필요한 폴더들을 생성하고, 기준일자 폴더 내 파일을 확인합니다.
    기준일자 폴더에 파일이 있으면 프로그램을 종료합니다.
    
    :param project_path: 프로젝트 경로
    """
    # 전송 폴더 및 기준일자 폴더 생성 및 파일 확인
    send_directory = os.path.join(project_path, "send")
    date_folder_name = datetime.now().strftime('%Y%m%d')
    send_date_folder_path = os.path.join(send_directory, date_folder_name)

    # 기준일자 폴더 확인 및 생성
    ensure_directory_exists(send_date_folder_path)

    # 기준일자 폴더에 이미 전송된 파일이 있는지 확인
    send_files = get_files_from_directory(send_date_folder_path, extension=".xlsx")
    if send_files:
        print(f"{datetime.now().strftime('%Y%m%d')} 이미 이메일이 발송되어 종료합니다.")
        sys.exit(0)

    # 다운로드, 엑셀 폴더 경로 생성
    downloads_folder = os.path.join(project_path, "downloads")
    excel_folder = os.path.join(project_path, "excel")
    send_folder = os.path.join(project_path, "send")

    # 폴더들 확인 및 생성
    ensure_directory_exists(downloads_folder)
    ensure_directory_exists(excel_folder)
    ensure_directory_exists(send_folder)

def setup_browser(playwright: Playwright):
    """
    Playwright 브라우저 설정을 처리합니다. 환경 변수에 따라 headless 모드를 설정합니다.
    
    :param playwright: Playwright 인스턴스
    :return: 생성된 페이지 객체
    """
    env = os.getenv('ENV')
    print(f"현재 환경: {env}")
    
    headless = False
    if env == 'production':
        headless = True

    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(locale="ko-KR")
    page = context.new_page()

    return browser, context, page

def run(playwright: Playwright) -> None:
    
    # 폴더 생성 및 파일 확인
    setup_directories(PROJECT_PATH)    
    
    # 브라우저 설정 및 실행
    browser, context, page = setup_browser(playwright)
    
    # 네이버 키워드 광고 엑셀(마피아닷컴)
    generate_naver_keyword_excel(page)
    
    # ---------------------
    context.close()
    browser.close()

    print("************* 이메일 발송 *************")
    # 이메일 전송
    gmail.main()

with sync_playwright() as playwright:
    run(playwright)
