import subprocess
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

def get_files_from_directory(directory, extensions=None):
    """
    특정 디렉토리에서 모든 파일을 가져오거나 특정 확장자의 파일만 가져옴.
    파일이 없는 경우 빈 리스트를 반환함.

    :param directory: 검색할 디렉토리 경로
    :param extensions: 특정 파일 확장자 또는 확장자 리스트 (예: '.txt' 또는 ['.txt', '.xlsx']). None일 경우 모든 파일을 반환.
    :return: 해당 디렉토리의 파일 경로 리스트. 파일이 없으면 빈 리스트.
    """
    files = []

    # extensions가 문자열이면 리스트로 변환
    if isinstance(extensions, str):
        extensions = [extensions]

    # 디렉토리 내 파일 목록 확인
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)

        # 파일이면서 확장자가 조건에 맞는 경우 리스트에 추가
        if os.path.isfile(file_path) and (extensions is None or any(file.endswith(ext) for ext in extensions)):
            files.append(file_path)

        if files:
            # 파일명 기준으로 정렬
            files = sorted(files, key=lambda x: x.split('/')[-1])
        
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
        print('='*30)
        print("================================첨부파일정보================================")
 
        result_message = (
            "=========이메일 전송 성공=========\n"
            f"발신자={SENDER_EMAIL}\n"
            f"수신자={RECEIVER_EMAIL}\n"
            "=============================\n"
        )

        # 경로별로 파일을 저장할 딕셔너리 생성
        path_dict = {}

        # 파일 경로에서 경로와 파일명 분리
        for attachment_path in attachment_paths:
            directory = '/'.join(attachment_path.split('/')[:-1])  # 디렉토리 경로 추출
            filename = attachment_path.split('/')[-1]  # 파일명 추출

            # 경로가 이미 딕셔너리에 있으면 파일명을 추가, 없으면 새 리스트 생성
            if directory in path_dict:
                path_dict[directory].append(filename)
            else:
                path_dict[directory] = [filename]

        # 경로와 해당 경로의 파일명을 출력
        for directory, filenames in path_dict.items():
            result_message += f"첨부파일 경로: {directory}/\n\n"
            for filename in filenames:
                result_message += f"{filename}\n"
            result_message += "\n"

        result_message += "============================="
        send_message_to_shell(result_message)

        # 연결 종료
        server.quit()
        return True  # 이메일 전송 성공 시 True 반환

    except Exception as e:
        print(f"이메일 전송 중 오류 발생: {e}")
        
        # 실패 메시지 생성 및 쉘로 전송
        result_message = f"이메일 전송 실패: {e}"
        # send_message_to_shell(result_message)
        return False  # 오류 발생 시 False 반환

def send_message_to_shell(result_message):
    """
    쉘 스크립트로 메시지를 전송하는 함수
    :param result_message: 쉘로 보낼 메시지 내용
    """
    try:
        subprocess.run(['/home/ubuntu/sh/sendmsg.sh', result_message], check=True)
        print(f"메시지가 쉘 스크립트로 전송되었습니다: {result_message}")
    except subprocess.CalledProcessError as e:
        print(f"쉘 스크립트 실행 중 오류 발생: {e}")

def main():
    print("************* 이메일 발송 *************")
    directory = os.path.join(PROJECT_PATH, "excel")
    
    result = ''
    
    # 디렉토리가 없으면 생성
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"{directory} 디렉토리가 생성되었습니다.")
        
    # 확장자 리스트로 엑셀과 CSV 파일을 함께 가져옴
    attachment_paths = get_files_from_directory(directory, extensions=[".xlsx", ".csv"])

    # send 디렉토리 경로 설정
    send_directory = os.path.join(PROJECT_PATH, "send")
    if not os.path.exists(send_directory):
        os.makedirs(send_directory)
        print(f"{send_directory} 디렉토리가 생성되었습니다.")
    
    # 이메일 보내기
    if attachment_paths:
        print('='*50)
        print('전송 대기 파일.')
        for attachment_path in attachment_paths:
            print(attachment_path)
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