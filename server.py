#!/usr/bin/env python3
"""
다정 & 선준 가계부 대시보드 - 맥북 로컬 개발 & 아이폰 실시간 연동 서버
실행 방법: 
  python3 server.py 
  또는 start.command 더블 클릭
"""

import os
import sys
import socket
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler

DEFAULT_PORT = 8080

def get_local_ip():
    """현재 맥북의 Wi-Fi 로컬 IP 주소 추출"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def find_available_port(start_port=DEFAULT_PORT):
    """사용 가능한 포트 탐색"""
    port = start_port
    while port < start_port + 100:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(('0.0.0.0', port)) != 0:
                return port
        port += 1
    return start_port

def print_banner(local_url, iphone_ip_url, port):
    swdjsj_url = f"http://swdjsj.local:{port}"
    mac_host_url = f"http://gimseonjun-ui-MacBookAir.local:{port}"

    print("\n" + "=" * 68)
    print("      ✨ 다정 & 선준 가계부 대시보드 로컬 서버가 시작되었습니다! ✨")
    print("=" * 68)
    print(f" 💻 맥북 접속 주소      : \033[1;36m{local_url}\033[0m")
    print(f" 📱 아이폰 간편 주소 (1) : \033[1;32m{swdjsj_url}\033[0m  \033[0;33m(추천: swdjsj 적용시)\033[0m")
    print(f" 📱 아이폰 호스트 주소 (2): \033[1;32m{mac_host_url}\033[0m  \033[0;32m(지금 바로 가능)\033[0m")
    print(f" 📱 아이폰 숫자 IP 주소  : \033[1;30m{iphone_ip_url}\033[0m")
    print("-" * 68)
    
    # QR코드 터미널 출력 (swdjsj_url 기본)
    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(swdjsj_url)
        qr.make(fit=True)
        print(f" 📷 [아이폰 카메라로 아래 QR 코드를 비추면 즉시 열립니다] ({swdjsj_url})\n")
        qr.print_ascii(invert=True)
        print()
    except Exception as e:
        print(f" (QR코드 생성 안내: 아이폰 사파리 주소창에 {swdjsj_url} 입력)")

    print("=" * 68)
    print(" 💡 아이폰에서 접속이 안 될 때 체크리스트:")
    print("  1. [가장 중요] 아이폰의 Wi-Fi가 켜져 있고 맥북과 같은 Wi-Fi에 연결되어 있어야 합니다.")
    print("     (아이폰이 LTE/5G 데이터로 켜져 있으면 접속되지 않습니다)")
    print("  2. 주소창 입력 시 'http://' 를 반드시 포함해 주세요: http://swdjsj.local:8080")
    print("  3. 맥북에서 'swdjsj.local' 주소를 활성화하는 초간단 방법:")
    print("     맥북 [시스템 설정] ➔ [일반] ➔ [공유] ➔ 맨 밑 [로컬 호스트 이름]을 'swdjsj'로 변경!")
    print("  4. 서버를 종료하려면 터미널에서 [Ctrl + C]를 누르세요.")
    print("=" * 68 + "\n")

class CustomHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # 개발 중 캐시 방지 헤더 추가
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        # 보기 편한 로그 포맷 (요청 주소 및 응답 코드만 간결히 출력)
        sys.stderr.write(f"[{self.log_date_time_string()}] {self.address_string()} - {args[0]} {args[1]}\n")

def main():
    # 스크립트가 위치한 디렉토리로 이동
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    ip = get_local_ip()
    port = find_available_port(DEFAULT_PORT)
    
    local_url = f"http://localhost:{port}"
    iphone_url = f"http://{ip}:{port}"
    
    print_banner(local_url, iphone_url, port)
    
    # 맥북 브라우저 자동 오픈
    try:
        webbrowser.open(local_url)
    except Exception:
        pass
        
    server = HTTPServer(('0.0.0.0', port), CustomHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n👋 가계부 대시보드 서버를 안전하게 종료했습니다.")
        sys.exit(0)

if __name__ == '__main__':
    main()
