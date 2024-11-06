import imaplib
import email
from email.policy import default
import os
from dotenv import load_dotenv, set_key
import json
from datetime import datetime
from send_error import send_message_to_shell  # send_message_to_shell 함수 가져오기

load_dotenv()

# 환경 변수에서 이메일 정보 가져오기
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
RECEIVER_EMAIL = os.getenv('RECEIVER_EMAIL')
IMAP_SERVER = 'imap.gmail.com'

# .env 파일 경로 설정
ENV_PATH = '.env'
# JSON 변경 내역 파일 경로 설정
HISTORY_JSON_PATH = 'mapia_keywords_history.json'

def fetch_unread_emails_from_receiver():
    load_dotenv()
    try:
        # IMAP 서버에 연결하여 로그인
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(SENDER_EMAIL, SENDER_PASSWORD)
        mail.select('inbox')

        # 특정 수신자로부터 온 읽지 않은 이메일 검색
        status, messages = mail.search(None, f'(UNSEEN FROM "{RECEIVER_EMAIL}")')
        email_ids = messages[0].split()

        if not email_ids:
            print(f"[{RECEIVER_EMAIL}]로부터 받은 새로운 이메일이 없습니다.")
            mail.logout()
            return

        for e_id in email_ids:
            # 이메일 데이터 가져오기
            status, msg_data = mail.fetch(e_id, '(RFC822)')
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email, policy=default)

            # 이메일 본문 추출
            if msg.is_multipart():
                for part in msg.iter_parts():
                    if part.get_content_type() == 'text/plain':
                        email_content = part.get_payload(decode=True).decode('utf-8').strip()
                        print("이메일 본문:", email_content)
                        update_env_variable(email_content)  # 환경변수 업데이트
            else:
                email_content = msg.get_payload(decode=True).decode('utf-8').strip()
                print("이메일 본문:", email_content)
                update_env_variable(email_content)  # 환경변수 업데이트

        # 연결 종료
        mail.logout()

    except Exception as e:
        error_message = f"오류 발생: {str(e)}"
        print(error_message)
        send_message_to_shell(error_message)  # 관리자에게 오류 메시지 전송

def update_env_variable(new_value):
    """
    .env 파일의 MAPIA_KEYWORDS_STR 값을 이메일 본문 내용으로 업데이트하고,
    변경 내역을 JSON 파일에 기록
    """
    try:
        # .env 파일에 MAPIA_KEYWORDS_STR 업데이트
        new_value = new_value.strip()  # 공백 제거
        set_key(ENV_PATH, "MAPIA_KEYWORDS_STR", new_value)
        
        # 업데이트된 환경변수 값 출력
        updated_value = os.getenv('MAPIA_KEYWORDS_STR')
        print("기존 MAPIA_KEYWORDS_STR:", updated_value)
        print("변경된 MAPIA_KEYWORDS_STR:", new_value)
        
        # 변경 내역을 JSON 파일에 기록
        save_to_json_history(new_value)

    except Exception as e:
        error_message = f"환경변수 업데이트 오류: {str(e)}"
        print(error_message)
        send_message_to_shell(error_message)  # 관리자에게 오류 메시지 전송

def save_to_json_history(new_value):
    """
    MAPIA_KEYWORDS_STR 변경 내역을 JSON 파일에 저장
    """
    try:
        # 현재 시간과 새로운 값 기록
        update_entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "MAPIA_KEYWORDS_STR": new_value
        }

        # 기존 기록 로드 또는 빈 리스트 초기화
        if os.path.exists(HISTORY_JSON_PATH):
            with open(HISTORY_JSON_PATH, 'r', encoding='utf-8') as file:
                history_data = json.load(file)
        else:
            history_data = []

        # 새로운 변경 내역 추가
        history_data.append(update_entry)

        # JSON 파일에 기록
        with open(HISTORY_JSON_PATH, 'w', encoding='utf-8') as file:
            json.dump(history_data, file, ensure_ascii=False, indent=4)
        print(f"변경 내역이 {HISTORY_JSON_PATH}에 저장되었습니다.")

    except Exception as e:
        error_message = f"JSON 파일 저장 오류: {str(e)}"
        print(error_message)
        send_message_to_shell(error_message)  # 관리자에게 오류 메시지 전송

# 실행
if __name__ == "__main__":
    fetch_unread_emails_from_receiver()
