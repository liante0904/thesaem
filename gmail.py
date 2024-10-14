from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
import os
from dotenv import load_dotenv  # .env 파일을 로드하기 위한 모듈

# .env 파일 로드
load_dotenv()

# .env에서 이메일 정보 불러오기
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
RECEIVER_EMAIL = os.getenv('RECEIVER_EMAIL')

# 프로젝트 경로 가져오기
PROJECT_PATH = os.getenv('PROJECT_PATH')

def get_files_from_directory(directory, extension=None):
    """
    특정 디렉토리에서 모든 파일을 가져오거나 특정 확장자의 파일만 가져옴.
    """
    files = []
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        # 파일이면서 확장자가 조건에 맞는 경우 리스트에 추가
        if os.path.isfile(file_path) and (extension is None or file.endswith(extension)):
            files.append(file_path)
    return files

def send_email(attachment_paths=None):
    # 이메일 객체 생성
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"W더샘 데일리자료_{datetime.now().strftime('%Y%m%d')}"

    # 이메일 본문 추가
    body = " "
    msg.attach(MIMEText(body, 'plain'))

    # 첨부파일 추가
    if attachment_paths:
        for attachment_path in attachment_paths:
            # 파일 이름 추출
            filename = os.path.basename(attachment_path)
            try:
                # 파일을 바이너리 모드로 열기
                with open(attachment_path, "rb") as attachment:
                    # MIMEBase 객체 생성
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                # 파일을 base64로 인코딩
                encoders.encode_base64(part)

                # 파일 이름이 깨지지 않도록 따옴표로 감싸기
                
                # 파일 이름이 깨지지 않도록 UTF-8 인코딩 적용
                part.add_header(
                    "Content-Disposition",
                    'attachment; filename="%s"' % Header(filename, 'utf-8').encode()
                )
                # 이메일에 파일 첨부
                msg.attach(part)

            except Exception as e:
                print(f"첨부파일 '{filename}' 추가 중 오류 발생: {e}")

    try:
        # SMTP 서버에 연결
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()  # 서버에 인사 (필수)
        server.starttls()  # TLS 사용 시작 (보안)
        server.ehlo()  # TLS 후 다시 인사 (필수)
        
        # Gmail 로그인
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        
        # 이메일 전송
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
        print("이메일이 성공적으로 전송되었습니다.")
        
        # 연결 종료
        server.quit()
    
    except Exception as e:
        print(f"이메일 전송 중 오류 발생: {e}")


# 사용 예시
if __name__ == "__main__":
    # 파일이 저장된 디렉토리 (예: '/root/dev/thesaem/excel')
    directory = os.path.join(PROJECT_PATH, "excel")

    # 디렉토리에서 모든 .xlsx 파일을 가져옴
    attachment_paths = get_files_from_directory(directory, extension=".xlsx")
    
    # 이메일 보내기
    send_email(attachment_paths)
