import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import time
import random
import os
import re
import sys
from datetime import datetime
from urllib.parse import urljoin


# ============================================================
# [고정 설정]
# ============================================================

GALLERY_ID = "example_gallery"  # Change to the target public gallery ID

BASE_FILENAME = "community_crawling"

EXCEL_EXTENSION = ".xlsx"


# ============================================================
# [크롤링 설정]
# ============================================================

POST_DELAY_MIN = 1.0
POST_DELAY_MAX = 2.0

PAGE_DELAY_MIN = 1.0
PAGE_DELAY_MAX = 2.0

SAVE_INTERVAL = 20

RETRY_COUNT = 2


# ============================================================
# [실행 파일 위치]
# ============================================================

def get_program_folder():

    if getattr(sys, "frozen", False):

        return os.path.dirname(
            sys.executable
        )

    return os.path.dirname(
        os.path.abspath(__file__)
    )


PROGRAM_FOLDER = get_program_folder()


# ============================================================
# [Excel 파일명 자동 생성]
# ============================================================

def get_next_output_file():

    # 최초 파일
    first_file = os.path.join(
        PROGRAM_FOLDER,
        BASE_FILENAME + EXCEL_EXTENSION
    )

    if not os.path.exists(first_file):

        return first_file


    number = 2

    while True:

        filename = (
            f"{BASE_FILENAME}_{number}"
            f"{EXCEL_EXTENSION}"
        )

        filepath = os.path.join(
            PROGRAM_FOLDER,
            filename
        )

        if not os.path.exists(filepath):

            return filepath

        number += 1


# ============================================================
# [기존 Excel 중 가장 최신 파일 찾기]
# ============================================================

def find_latest_existing_file():

    files = []


    if not os.path.exists(
        PROGRAM_FOLDER
    ):

        return None


    for filename in os.listdir(
        PROGRAM_FOLDER
    ):

        if not filename.lower().endswith(
            ".xlsx"
        ):

            continue


        if not filename.startswith(
            BASE_FILENAME
        ):

            continue


        filepath = os.path.join(
            PROGRAM_FOLDER,
            filename
        )


        # 임시 Excel 파일 제외
        if filename.startswith(
            "~$"
        ):

            continue


        files.append(
            filepath
        )


    if not files:

        return None


    # 수정 시간이 가장 최신인 파일
    files.sort(
        key=os.path.getmtime,
        reverse=True
    )


    return files[0]


# ============================================================
# [URL]
# ============================================================

BASE_URL = "https://gall.dcinside.com"

LIST_URL = (
    "https://gall.dcinside.com/"
    "mgallery/board/lists/"
)


# ============================================================
# [HTTP 설정]
# ============================================================

HEADERS = {

    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),

    "Accept-Language": (
        "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    )

}


session = requests.Session()

session.headers.update(
    HEADERS
)


# ============================================================
# [텍스트 정리]
# ============================================================

def clean_text(text):

    if not text:

        return ""

    text = text.replace(
        "\xa0",
        " "
    )

    text = re.sub(
        r"\r\n",
        "\n",
        text
    )

    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text
    )

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# [페이지 요청]
# ============================================================

def request_page(url):

    for attempt in range(
        1,
        RETRY_COUNT + 1
    ):

        try:

            response = session.get(
                url,
                timeout=15
            )


            if response.status_code in [
                403,
                429
            ]:

                print()
                print(
                    f"[접근 제한 감지] "
                    f"HTTP {response.status_code}"
                )

                return None, "BLOCKED"


            response.raise_for_status()


            response.encoding = (
                response.apparent_encoding
            )


            return response.text, "OK"


        except requests.RequestException as e:

            print(
                f"요청 실패 "
                f"({attempt}/{RETRY_COUNT}) : {e}"
            )


            if attempt < RETRY_COUNT:

                time.sleep(
                    random.uniform(
                        2,
                        4
                    )
                )


    return None, "ERROR"


# ============================================================
# [게시판 목록 수집]
# ============================================================

def get_posts_from_page(page):

    url = (
        f"{LIST_URL}"
        f"?id={GALLERY_ID}"
        f"&page={page}"
    )


    print()
    print(
        f"[페이지 {page}] 목록 수집 중..."
    )


    html, status = request_page(
        url
    )


    if status == "BLOCKED":

        return None, "BLOCKED"


    if not html:

        return [], "ERROR"


    soup = BeautifulSoup(
        html,
        "html.parser"
    )


    posts = soup.select(
        "tr.ub-content"
    )


    results = []


    for post in posts:

        # 글번호
        number_tag = post.select_one(
            ".gall_num"
        )


        if not number_tag:

            continue


        number = clean_text(
            number_tag.get_text()
        )


        # 공지 / 광고 제외
        if not number.isdigit():

            continue


        # 추천 수
        recommend_tag = post.select_one(".gall_recommend")
        recommend_count = 0
        if recommend_tag:
            m = re.search(r"\d+", clean_text(recommend_tag.get_text(" ", strip=True)))
            if m:
                recommend_count = int(m.group())

        # 댓글 수
        comment_tag = post.select_one(".reply_num")
        comment_count = 0
        if comment_tag:
            m = re.search(r"\d+", clean_text(comment_tag.get_text(" ", strip=True)))
            if m:
                comment_count = int(m.group())

        # 제목
        title_tag = post.select_one(
            ".gall_tit a"
        )


        if not title_tag:

            continue


        title = clean_text(
            title_tag.get_text(
                " ",
                strip=True
            )
        )


        # 링크
        href = title_tag.get(
            "href",
            ""
        )


        if not href:

            continue


        post_url = urljoin(
            BASE_URL,
            href
        )


        # 작성자
        writer_tag = post.select_one(
            ".gall_writer"
        )


        writer = ""


        if writer_tag:

            writer = clean_text(
                writer_tag.get_text(
                    " ",
                    strip=True
                )
            )


        # 작성자 정보
        writer_info = ""


        if writer_tag:

            writer_info = writer_tag.get(
                "data-ip",
                ""
            )


            if not writer_info:

                ip_tag = writer_tag.select_one(
                    "[data-ip]"
                )


                if ip_tag:

                    writer_info = ip_tag.get(
                        "data-ip",
                        ""
                    )


        writer_info = clean_text(
            writer_info
        )


        # 작성일시
        date_tag = post.select_one(
            ".gall_date"
        )


        date = ""


        if date_tag:

            date = date_tag.get(
                "title",
                ""
            )


            if not date:

                date = date_tag.get_text(
                    strip=True
                )


        date = clean_text(
            date
        )


        results.append({

            "작성일시": date,

            "제목": title,

            "내용": "",

            "작성자": writer,

            "작성자정보": writer_info,

            "글번호": number,

            "링크": post_url,

            "_추천수": recommend_count,

            "_댓글수": comment_count

        })


    print(
        f"게시글 {len(results)}개 발견"
    )


    return results, "OK"


# ============================================================
# [본문 수집]
# ============================================================

def get_post_content(post):

    html, status = request_page(
        post["링크"]
    )


    if status == "BLOCKED":

        return None, "BLOCKED"


    if not html:

        return "", "ERROR"


    soup = BeautifulSoup(
        html,
        "html.parser"
    )


    content = soup.select_one(
        ".write_div"
    )


    if not content:

        content = soup.select_one(
            ".writing_view_box"
        )


    if not content:

        return "", "OK"


    # 불필요한 태그 제거
    for tag in content.select(
        "script, style, iframe, button"
    ):

        tag.decompose()


    text = content.get_text(
        "\n",
        strip=True
    )


    return clean_text(
        text
    ), "OK"


# ============================================================
# [기존 데이터 불러오기]
# ============================================================

def load_existing_data():

    latest_file = (
        find_latest_existing_file()
    )


    if not latest_file:

        print()
        print(
            "기존 Excel 파일이 없습니다."
        )

        return pd.DataFrame(
            columns=[
                "작성일시",
                "제목",
                "내용",
                "작성자",
                "작성자정보",
                "글번호",
                "링크"
            ]
        )


    print()
    print(
        "기존 결과 파일:"
    )

    print(
        os.path.basename(
            latest_file
        )
    )


    try:

        df = pd.read_excel(
            latest_file,
            engine="openpyxl"
        )


        print(
            f"기존 데이터 "
            f"{len(df)}개 발견"
        )


        return df


    except Exception as e:

        print()
        print(
            f"Excel 읽기 실패: {e}"
        )


        return pd.DataFrame(
            columns=[
                "작성일시",
                "제목",
                "내용",
                "작성자",
                "작성자정보",
                "글번호",
                "링크"
            ]
        )


# ============================================================
# [Excel 저장]
# ============================================================

def save_data(
    data,
    output_file
):

    if not data:

        return


    df = pd.DataFrame(
        data,
        columns=[
            "작성일시",
            "제목",
            "내용",
            "작성자",
            "작성자정보",
            "글번호",
            "링크"
        ]
    )


    # 글번호 숫자 변환
    df["글번호"] = pd.to_numeric(
        df["글번호"],
        errors="coerce"
    )


    # 중복 제거
    df = df.drop_duplicates(
        subset=["글번호"],
        keep="last"
    )


    # 글번호 높은 순
    df = df.sort_values(
        by="글번호",
        ascending=False
    )


    display_columns = [
        "작성일시",
        "제목",
        "내용",
        "작성자",
        "작성자정보",
        "글번호",
        "링크"
    ]

    df[display_columns].to_excel(
        output_file,
        index=False,
        engine="openpyxl"
    )

    # 추천 10개 이상 OR 댓글 5개 이상인 게시글 전체 행 노란색 표시
    try:
        workbook = load_workbook(output_file)
        worksheet = workbook.active

        yellow_fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC"
        )

        for excel_row, (_, row) in enumerate(df.iterrows(), start=2):
            recommend_count = int(row.get("_추천수", 0) or 0)
            comment_count = int(row.get("_댓글수", 0) or 0)

            if recommend_count >= 10 or comment_count >= 5:
                for cell in worksheet[excel_row]:
                    cell.fill = yellow_fill

        workbook.save(output_file)

    except Exception as e:
        print(f"[Excel 서식] 노란색 표시 실패: {e}")


# ============================================================
# [페이지 입력]
# ============================================================

def get_page_range():

    print()
    print("=" * 70)
    print("       디시인사이드 갤러리 크롤러")
    print("=" * 70)

    print()
    print(
        f"수집 대상 갤러리 : "
        f"{GALLERY_ID}"
    )

    print()
    print(
        "수집할 페이지 범위를 입력하세요."
    )

    print(
        "예시 : 1 ~ 19페이지"
    )

    print()


    while True:

        try:

            start_page = int(
                input(
                    "시작 페이지 : "
                ).strip()
            )


            if start_page < 1:

                print(
                    "페이지는 1 이상이어야 합니다."
                )

                continue


            break


        except ValueError:

            print(
                "숫자만 입력해주세요."
            )


    while True:

        try:

            end_page = int(
                input(
                    "마지막 페이지 : "
                ).strip()
            )


            if end_page < start_page:

                print(
                    "마지막 페이지는 "
                    "시작 페이지보다 작을 수 없습니다."
                )

                continue


            break


        except ValueError:

            print(
                "숫자만 입력해주세요."
            )


    return start_page, end_page


# ============================================================
# [실행 확인]
# ============================================================

def confirm_start(
    start_page,
    end_page
):

    print()
    print("-" * 70)

    print(
        f"갤러리 : {GALLERY_ID}"
    )

    print(
        f"수집 범위 : "
        f"{start_page} ~ {end_page} 페이지"
    )

    print(
        "기존 Excel은 덮어쓰지 않습니다."
    )

    print(
        "새로운 번호의 Excel 파일로 저장됩니다."
    )

    print("-" * 70)


    while True:

        answer = input(
            "\n크롤링을 시작하시겠습니까? "
            "(Y/N) : "
        ).strip().lower()


        if answer in [
            "y",
            "yes"
        ]:

            return True


        if answer in [
            "n",
            "no"
        ]:

            return False


        print(
            "Y 또는 N을 입력해주세요."
        )


# ============================================================
# [메인]
# ============================================================

def main():

    start_time = datetime.now()


    # 페이지 입력
    start_page, end_page = (
        get_page_range()
    )


    # 실행 확인
    if not confirm_start(
        start_page,
        end_page
    ):

        print()
        print(
            "크롤링을 취소했습니다."
        )

        input(
            "\n종료하려면 Enter를 누르세요."
        )

        return


    # --------------------------------------------------------
    # 기존 데이터
    # --------------------------------------------------------

    existing_df = (
        load_existing_data()
    )


    existing_numbers = set()


    if not existing_df.empty:

        for number in existing_df[
            "글번호"
        ].dropna():

            try:

                existing_numbers.add(
                    str(int(number))
                )

            except:

                pass


    print()
    print(
        f"기존 글번호 : "
        f"{len(existing_numbers)}개"
    )


    # --------------------------------------------------------
    # 새 결과 파일명 결정
    # --------------------------------------------------------

    output_file = (
        get_next_output_file()
    )


    print()
    print(
        "이번 실행 결과:"
    )

    print(
        os.path.basename(
            output_file
        )
    )


    # --------------------------------------------------------
    # 목록 수집
    # --------------------------------------------------------

    all_posts = []


    for page in range(
        start_page,
        end_page + 1
    ):


        posts, status = (
            get_posts_from_page(
                page
            )
        )


        if status == "BLOCKED":

            print()
            print(
                "접근 제한이 감지되었습니다."
            )

            print(
                "현재까지의 데이터를 저장합니다."
            )

            break


        if status == "ERROR":

            print(
                f"페이지 {page} "
                f"수집 실패"
            )

            continue


        all_posts.extend(
            posts
        )


        if page < end_page:

            delay = random.uniform(
                PAGE_DELAY_MIN,
                PAGE_DELAY_MAX
            )


            print(
                f"다음 페이지까지 "
                f"{delay:.1f}초 대기..."
            )


            time.sleep(
                delay
            )


    # --------------------------------------------------------
    # 게시글 중복 제거
    # --------------------------------------------------------

    unique_posts = {}


    for post in all_posts:

        number = post["글번호"]


        if number not in unique_posts:

            unique_posts[number] = post


    all_posts = list(
        unique_posts.values()
    )


    # --------------------------------------------------------
    # 신규 게시글만
    # --------------------------------------------------------

    new_posts = []


    for post in all_posts:

        if post["글번호"] not in existing_numbers:

            new_posts.append(
                post
            )


    print()
    print("=" * 70)

    print(
        f"전체 게시글 : "
        f"{len(all_posts)}개"
    )

    print(
        f"신규 게시글 : "
        f"{len(new_posts)}개"
    )

    print(
        f"기존 게시글 : "
        f"{len(all_posts) - len(new_posts)}개"
    )

    print("=" * 70)


    # --------------------------------------------------------
    # 본문 수집
    # --------------------------------------------------------

    completed = 0


    try:

        for post in new_posts:

            print()
            print(
                f"[{completed + 1}/"
                f"{len(new_posts)}]"
            )

            print(
                f"글번호 : "
                f"{post['글번호']}"
            )

            print(
                f"제목 : "
                f"{post['제목']}"
            )


            content, status = (
                get_post_content(
                    post
                )
            )


            if status == "BLOCKED":

                print()
                print(
                    "접근 제한이 감지되었습니다."
                )

                print(
                    "현재까지 수집한 데이터를 저장합니다."
                )

                break


            if status == "OK":

                post["내용"] = content

                print(
                    "본문 수집 완료"
                )

            else:

                print(
                    "본문 수집 실패"
                )


            completed += 1


            # ------------------------------------------------
            # 중간 저장
            # ------------------------------------------------

            if (
                completed
                % SAVE_INTERVAL
                == 0
            ):

                combined_data = (
                    existing_df.to_dict(
                        "records"
                    )
                    +
                    new_posts[:completed]
                )


                save_data(
                    combined_data,
                    output_file
                )


                print()
                print(
                    "[중간 저장 완료]"
                )


            # ------------------------------------------------
            # 게시글 사이 대기
            # ------------------------------------------------

            if completed < len(new_posts):

                delay = random.uniform(
                    POST_DELAY_MIN,
                    POST_DELAY_MAX
                )


                print(
                    f"다음 게시글까지 "
                    f"{delay:.1f}초 대기..."
                )


                time.sleep(
                    delay
                )


    except KeyboardInterrupt:

        print()
        print(
            "사용자가 크롤링을 중단했습니다."
        )


    finally:

        # ----------------------------------------------------
        # 최종 저장
        # ----------------------------------------------------

        combined_data = (
            existing_df.to_dict(
                "records"
            )
            +
            new_posts[:completed]
        )


        save_data(
            combined_data,
            output_file
        )


    # --------------------------------------------------------
    # 완료
    # --------------------------------------------------------

    elapsed = (
        datetime.now()
        - start_time
    )


    print()
    print("=" * 70)

    print(
        "크롤링 완료"
    )

    print(
        f"이번 실행 신규 수집 : "
        f"{completed}개"
    )

    print(
        f"최종 데이터 : "
        f"{len(combined_data)}개"
    )

    print(
        f"소요 시간 : "
        f"{elapsed}"
    )

    print(
        f"저장 파일 : "
        f"{os.path.basename(output_file)}"
    )

    print("=" * 70)

    print()

    input(
        "종료하려면 Enter를 누르세요."
    )


# ============================================================
# [실행]
# ============================================================

if __name__ == "__main__":

    main()
