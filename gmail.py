from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
import os
import shutil
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

def send_email(attachment_paths=None):
    # 이메일 객체 생성
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"더샘 데일리자료_{datetime.now().strftime('%Y%m%d')}"

    # 이메일 본문 추가
    body = "첨부파일을 확인해 주세요."
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
        
        print('='*30)
        print("이메일이 성공적으로 전송되었습니다.")
        print(f"발신자:{SENDER_EMAIL}")
        print(f"수신자:{RECEIVER_EMAIL}")
        print(f"첨부파일정보:{attachment_paths}")
        print('='*30)
        # 연결 종료
        server.quit()
        return True  # 이메일 전송 성공 시 True 반환

    except Exception as e:
        print(f"이메일 전송 중 오류 발생: {e}")
        return False  # 오류 발생 시 False 반환

def main():
    print("************* 이메일 발송 *************")
    directory = os.path.join(PROJECT_PATH, "excel")
    
    result = ''
    
    # 디렉토리가 없으면 생성
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"{directory} 디렉토리가 생성되었습니다.")
        
    # 디렉토리에서 모든 .xlsx 파일을 가져옴
    attachment_paths = get_files_from_directory(directory, extension=".xlsx")
    
    # send 디렉토리 경로 설정
    send_directory = os.path.join(PROJECT_PATH, "send")
    if not os.path.exists(send_directory):
        os.makedirs(send_directory)
        print(f"{send_directory} 디렉토리가 생성되었습니다.")
    
    # 이메일 보내기
    if attachment_paths:
        print('전송 대기 파일.')
        print(attachment_paths)
        print('='*50)
        result = send_email(attachment_paths)
    else:
        print('전송 대기 파일이 없습니다.')

    if result:
        print('='*50)
        print('전송 완료 폴더로 이동 합니다.')
        # 기준일자 폴더 이름 생성
        date_folder_name = datetime.now().strftime('%Y%m%d')
        date_folder_path = os.path.join(send_directory, date_folder_name)
        
        # 기준일자 폴더가 없으면 생성
        if not os.path.exists(date_folder_path):
            os.makedirs(date_folder_path)
            print(f"{date_folder_path} 폴더가 생성되었습니다.")

        # 파일을 기준일자 폴더로 이동
        for file_path in attachment_paths:
            file_name = os.path.basename(file_path)  # 파일 이름 추출
            destination = os.path.join(date_folder_path, file_name)
            shutil.move(file_path, destination)  # 파일 이동
            print(f"{file_name}이(가) {date_folder_path}로 이동되었습니다.")    

# 사용 예시
if __name__ == "__main__":
    main()