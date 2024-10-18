import requests
import time
import pandas as pd
import asyncio
import aiohttp
import os
from datetime import datetime

from dotenv import load_dotenv


async def generate_naver_keyword_excel():
    load_dotenv()
    MAPIA_KEYWORDS_STR = os.getenv("MAPIA_KEYWORDS_STR")

    PROJECT_PATH = os.getenv("PROJECT_PATH")

    # 현재 시간을 밀리초로 변환
    current_time = int(time.time() * 1000)

    # 결과를 저장할 리스트
    data = {
        "Keyword": [],
        "POST_result_1": [],
        "POST_result_2": [],
        "POST_result_3": [],
        "POST_result_4": [],
        "POST_result_5": [],
        "POST_result_6": [],
        "POST_result_7": [],
        "POST_result_8": [],
        "POST_sum_1_2": [],  # POST_result_1 + POST_result_2 항목 추가
        "GET_shopCategory": [],
        "GET_monthBlog": [],
        "GET_blogSaturation": [],
    }

    # 헤더 설정
    headers = {
        "referer": "https://www.ma-pia.net/",
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    }

    mapia_splited_keywords = [
        keyword.strip()
        for keyword in MAPIA_KEYWORDS_STR.replace(",", "\n")
        .replace("\\n", "\n")
        .split("\n")
        if keyword.strip()
    ]

    print(f"****변환 키워드 리스트*****\n{mapia_splited_keywords}")
    print(f"****변환 키워드 건수*****\n{len(mapia_splited_keywords)}")


    # 100개 단위로 키워드를 나누기
    keyword_chunks = [
        mapia_splited_keywords[i : i + 100]
        for i in range(0, len(mapia_splited_keywords), 100)
    ]

    # POST 요청 처리 (기존 동기 방식 유지)
    for keyword_chunk in keyword_chunks:
        print(f"\n\n현재 키워드 세트: {keyword_chunk}")

        # POST 요청을 위한 키워드들을 쉼표로 구분한 문자열로 변환
        keyword_str = ",".join(keyword_chunk)

        # POST 요청
        post_url = "https://www.ma-pia.net/keyword/keywordQ.php"
        post_data = {"DataQ": keyword_str}

        retry_count = 0
        max_retries = 10
        success = False

        while retry_count < max_retries and not success:
            print(f"POST 요청 중... 키워드: {keyword_str}")
            post_response = requests.post(post_url, data=post_data, headers=headers)

            if post_response.status_code == 200:
                print(
                    f"POST 요청 성공! 응답: {post_response.text[:200]}..."
                )  # 응답의 첫 200자를 출력
                if "success :" in post_response.text:
                    post_results = post_response.text.split("success :")[1].split("|||")

                    print(f"POST 결과 처리 중... 총 {len(post_results)}개의 키워드 결과")

                    # 각 키워드별로 결과 처리
                    for post_result in post_results:
                        result = post_result.split("///")

                        # 빈 문자열 또는 유효하지 않은 데이터를 건너뛰기
                        if len(result) < 9 or not result[0]:
                            print(f"경고: 결과가 유효하지 않습니다. 건너뜀. 결과: {result}")
                            continue

                        print(f"결과 분리: {result}")  # 결과 디버깅용 출력

                        # POST 결과 추가
                        data["Keyword"].append(result[0])
                        data["POST_result_1"].append(result[1])
                        data["POST_result_2"].append(result[2])
                        data["POST_result_3"].append(result[3])
                        data["POST_result_4"].append(result[4])
                        data["POST_result_5"].append(result[5])
                        data["POST_result_6"].append(result[6])
                        data["POST_result_7"].append(result[7])
                        data["POST_result_8"].append(result[8])

                        # POST_result_1 + POST_result_2 합산 결과 추가
                        sum_1_2 = int(result[1]) + int(result[2])
                        data["POST_sum_1_2"].append(sum_1_2)

                    success = True
                else:
                    print(
                        f"POST 응답에 'success :'가 없습니다. 재시도 중... ({retry_count + 1}/{max_retries})"
                    )
            else:
                print(f"POST 요청 실패: 상태 코드 {post_response.status_code}")

            retry_count += 1
            time.sleep(1)  # 재시도 전에 1초 대기

        if not success:
            print(f"POST 요청 실패: {keyword_str} (최대 재시도 횟수 초과)")
            continue


    # 비동기 GET 요청 함수
    async def fetch_get_data(session, keyword, sum_1_2, current_time, headers):
        get_url = "https://uy3w6h3mzi.execute-api.ap-northeast-2.amazonaws.com/Prod/hello"
        params = {"keyword": keyword, "totalSum": sum_1_2, "time": current_time}

        async with session.get(get_url, params=params, headers=headers) as response:
            if response.status == 200:
                print(f"GET 요청 성공! 키워드: {keyword}")
                get_data = await response.json()
                return (
                    get_data["result"]["shopCategory"],
                    get_data["result"]["monthBlog"],
                    get_data["result"]["blogSaturation"],
                )
            else:
                print(f"GET 요청 실패: 키워드 {keyword}, 상태 코드 {response.status}")
                return None, None, None


    # 비동기 GET 요청 처리
    async def fetch_all_get_data():
        async with aiohttp.ClientSession() as session:
            tasks = []
            for i, keyword in enumerate(data["Keyword"]):
                sum_1_2 = data["POST_sum_1_2"][i]
                task = fetch_get_data(session, keyword, sum_1_2, current_time, headers)
                tasks.append(task)
            results = await asyncio.gather(*tasks)

            # GET 결과 추가
            for shop_category, month_blog, blog_saturation in results:
                data["GET_shopCategory"].append(shop_category)
                data["GET_monthBlog"].append(month_blog)
                data["GET_blogSaturation"].append(blog_saturation)

    # 비동기 GET 요청 실행
    await fetch_all_get_data()  # await 추가

    # 수집된 데이터를 CSV 파일로 저장
    df = pd.DataFrame(data)
    print("현재 데이터프레임 열 목록:", df.columns.tolist())

    # 정확한 열 순서로 데이터프레임을 재정렬
    df = df[
        [
            "Keyword",
            "POST_result_1",
            "POST_result_2",
            "POST_sum_1_2",
            "GET_shopCategory",
            "GET_monthBlog",
            "GET_blogSaturation",
            "POST_result_3",
            "POST_result_4",
            "POST_result_5",
            "POST_result_6",
            "POST_result_7",
            "POST_result_8",
        ]
    ]
    df.columns = [
        "키워드",
        "월간검색수\nPC",
        "월간검색수\n모바일",
        "검색수\n합계",
        "월간 블로그 발행\n수량",
        "월간 블로그 발행\n포화도",
        "네이버쇼핑\n카테고리",
        "월평균클릭수\nPC",
        "월평균클릭수\n모바일",
        "월평균클릭율(%)\nPC",
        "월평균클릭율(%)\n모바일",
        "경쟁정도",
        "월평균\n노출광고수",
    ]



    # 저장할 경로 설정
    save_directory = os.path.join(PROJECT_PATH, "excel")

    # CSV 파일로 저장
    csv_file_name = f"더샘_브랜드검색_키워드쿼리_{datetime.now().strftime('%Y%m%d')}.csv"

    output_path = os.path.join(save_directory, csv_file_name)

    df.to_csv(csv_file_name, index=False, encoding="utf-8-sig")

    print(f"데이터가 CSV 파일로 저장되었습니다: {csv_file_name}")

if __name__ == "__main__":
    asyncio.run(generate_naver_keyword_excel())