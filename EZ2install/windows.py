import subprocess
import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
from mega import Mega


class WindowsInstaller:
    def __init__(self):
        # Windows 버전에 따른 ISO 다운로드 링크 설정
        self.download_links = {
            #"Windows 10 x64": "https://drive.google.com/uc?id=1OyA527cvaXC9k6WQ3k1OPTX6Prqe2OZm",
            #"Windows 11 x64": "https://drive.google.com/uc?id=1aGnNeDCE-2TrvcXEqwNpdRZjVPHzfGF2"
            "Windows 10 x64": "https://mega.nz/file/0P8WwA6b#pWjCQGRYn5xooSqnauHY6j2KGyMCJit5JOeasJNJUb4",
            "Windows 11 x64": "https://mega.nz/file/dekiUBaK#sOQQ_EFBIGnRwVltznt9sFfmcKQYkgWPNhH1Gyt5FwU"

        }

    # ISO 파일 다운로드 및 폴더 복사 (메소드 정의)
    def install_windows(self, selected_option):
        # 선택된 옵션에 따라 zip 다운로드 링크 가져오기
        url = self.download_links.get(selected_option)
        if not url:
            print("올바른 버전을 선택하세요.")
            return
        try:
            # zip 파일 다운로드
            temp_folder = tempfile.gettempdir()
            win_file_path = os.path.join(temp_folder, "windows.zip")

            mega = Mega()
            m = mega.login("seonghoonxg@gmail.com", "kosh1108k!")
            m.download_url(url)
            m.download(file, temp_folder, "Windows.zip")

            # 실행 파일이 위치한 드라이브의 최상위 루트를 가져옴
            exe_drive_root = os.path.abspath(os.path.dirname(__file__))[:3]  # 실행 파일이 있는 드라이브 루트 (예: C:\)

            # zip 파일을 실행 파일이 위치한 드라이브의 최상위 루트로 압축 해제
            with zipfile.ZipFile(win_file_path, 'r') as zip_ref:
                zip_ref.extractall(exe_drive_root)

            # zip 파일 삭제 (압축 해제 후에는 zip 파일이 필요 없으므로 삭제)
            os.remove(win_file_path)

            print("설치 파일이 설치 위치로 압축 해제되었습니다.")

        except Exception as e:
            print(f"설치 중 오류가 발생했습니다: {e}")

    # 부팅 메뉴 설정 (메소드 정의)
    def setup_bootmenu(self):
        # bcdedit 명령어를 창을 띄우지 않고 실행하기 위해 shell=True 사용
        subprocess.run(
            "bcdedit /create {ramdiskoptions} /d 'Setup Windows 10' || bcdedit /set {ramdiskoptions} description 'Setup Windows 10'",
            shell=True)
        subprocess.run("bcdedit /set {ramdiskoptions} ramdisksdidevice partition=%~d0", shell=True)
        subprocess.run("bcdedit /set {ramdiskoptions} ramdisksdipath \\boot\\boot.sdi", shell=True)

        setlocal_command = [
            "setlocal enabledelayedexpansion",
            "for /f 'tokens=1-5 usebackq delims=-' %%a in (`bcdedit /create /d 'Setup Windows 10' /application osloader`) do (",
            "set first=%%a",
            "set last=%%e",
            "set guid=!first:~-9!-%%b-%%c-%%d-!last:~0,13!",
            ")"
        ]
        # 리스트 형태의 명령어를 subprocess.run()에 전달하고 shell=True로 실행
        subprocess.run(setlocal_command, shell=True)

        subprocess.run("bcdedit /set !guid! device ramdisk=[%~d0]\\sources\\boot.wim,{ramdiskoptions}", shell=True)
        subprocess.run("bcdedit /set !guid! osdevice ramdisk=[%~d0]\\sources\\boot.wim,{ramdiskoptions}", shell=True)
        subprocess.run("set bios=exe", shell=True)
        subprocess.run("findstr /c:'Detected boot environment: BIOS' 'C:\\Windows\\Panther\\setupact.log'", shell=True)
        subprocess.run("if errorlevel 1 set bios=efi", shell=True)
        subprocess.run("bcdedit /set !guid! path \\windows\\system32\\winload.!bios!", shell=True)
        subprocess.run("bcdedit /set !guid! systemroot \\windows", shell=True)
        subprocess.run("bcdedit /set !guid! winpe yes", shell=True)
        subprocess.run("bcdedit /set !guid! detecthal yes", shell=True)
        subprocess.run("bcdedit /displayorder !guid! /addlast", shell=True)
        subprocess.run("bcdedit /timeout 10", shell=True)


# 윈도우 인증 버튼 연결    
class WindowsActivator:
    def __init__(self):
        pass

    # 윈도우 인증
    def activate_windows(self):
        activation_cmds = [
            'slmgr /ipk "W269N-WFGWX-YVC9B-4J6C9-T83GX"',
            'slmgr /skms kms8.msguides.com',
            'slmgr /ato'
        ]

        for cmd in activation_cmds:
            # 각 명령어를 shell=True로 실행하여 창을 띄우지 않고 실행
            subprocess.run(cmd, shell=True)





