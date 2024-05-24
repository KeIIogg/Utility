import subprocess
import os
import sys
import zipfile
import requests
import tempfile
import ctypes
import psutil
from mega import Mega
from PyQt5.QtCore import QObject, pyqtSignal

class WindowsInstaller(QObject):
    progress_callback = pyqtSignal(int, str)

    def __init__(self, selected_option):
        QObject.__init__(self)
        self.selected_option = selected_option
    def install_windows(self):
        temp_folder = tempfile.gettempdir()
        exe_drive_root = os.path.abspath(os.path.dirname(__file__))[:3]

        self.progress_callback.emit(20, "윈도우 파일 다운로드 중...")
        # Check if boot and sources folders exist, if so, skip installation
        if all(folder in os.listdir(exe_drive_root) for folder in ["boot", "sources"]):
            print("boot 및 sources 폴더가 이미 존재합니다. Windows 파일다운로드를 건너뜁니다.")
            return

        if self.selected_option == "Windows 10 x64":
            url1 = "https://mega.nz/file/RKUEgTKR#sJLxvH03M8eLqfxDv1Dc87AJON2H4D31e5gz0Qj9ve4"
            url2 = "https://mega.nz/file/5akAWRDT#DUflQI3yXkngpuKZib0gnvTtF59WszulHO5dY-g03U0"
            url3 = "https://mega.nz/file/cLtHXbja#73uvjAY32Q5x31SXT-jWEygA2-qmkQ1wTzqjAd5BGyQ"
        elif self.selected_option == "Windows 11 x64":
            url1 = "https://mega.nz/file/EKtwELiS#mEfVQlNUM2OzSte0P3LSmXy_sSVx4Wb2KJwHSsJamOo"
            url2 = "https://mega.nz/file/NbtHSKhZ#7RlSa5DSCHrIiMIrqO0L-CZL7aFqUUVBU_qKaDgPjwI"
            url3 = "https://mega.nz/file/sDEW2IBD#QRMCK5t3zpr1WNrbu0KCH1qRXaSXrRM0pC_HdQPvsHw"
        else:
            raise ValueError("올바른 버전을 선택하세요.")

        # Cloudflare WARP 설치 경로가 이미 존재하는지 확인
        warp_installed = os.path.exists(r"C:\Program Files\Cloudflare\Cloudflare WARP\Cloudflare WARP.exe")

        warp_url = "https://1111-releases.cloudflareclient.com/windows/Cloudflare_WARP_Release-x64.msi"
        warp_file_path = os.path.join(temp_folder, "Cloudflare_WARP_Release-x64.msi")

        self.progress_callback.emit(25, "윈도우 파일 다운로드 툴 설치 중...")
        # Cloudflare WARP 설치 경로가 이미 존재하지 않는 경우에만 다운로드 및 설치 수행
        if not warp_installed:
            with requests.get(warp_url, stream=True) as r:
                with open(warp_file_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)

            # Cloudflare WARP 설치 및 실행
            subprocess.run(["msiexec", "/i", warp_file_path, "/quiet", "/passive"])
            subprocess.Popen([r"C:\Program Files\Cloudflare\Cloudflare WARP\Cloudflare WARP.exe"])

        else:
            print("Cloudflare WARP가 이미 설치되어 있습니다. 설치 과정을 생략합니다.")
            subprocess.Popen([r"C:\Program Files\Cloudflare\Cloudflare WARP\Cloudflare WARP.exe"])

        self.progress_callback.emit(30, "윈도우 파일 다운로드 툴 설치 완료")

        # Mega 객체 생성
        mega = Mega()
        m = mega.login("seonghoonxg@gmail.com", "kosh1108k!")
        file1 = m.find('downloaded_files.zip')

        self.progress_callback.emit(35, "윈도우 파일 (1) 다운로드 중...")

        # 파일1 다운로드
        m.download_url(url1)
        m.download(file1, temp_folder, "downloaded_files.zip")

        self.progress_callback.emit(45, "윈도우 파일 (1) 다운로드 완료")

        # Cloudflare WARP 프로세스 종료 및 재실행
        for proc in psutil.process_iter():
            try:
                if "Cloudflare WARP.exe" in proc.name():
                    proc.kill()
            except psutil.AccessDenied:
                pass
        subprocess.Popen(["C:/Program Files/Cloudflare/Cloudflare WARP/Cloudflare WARP.exe"])

        self.progress_callback.emit(50, "윈도우 파일 (2) 다운로드 중...")

        file2 = m.find('downloaded_files.z02')
        # 파일2 다운로드
        m.download_url(url2)
        m.download(file2, temp_folder, "downloaded_files.z01")

        self.progress_callback.emit(60, "윈도우 파일 (2) 다운로드 완료")

        # Cloudflare WARP 프로세스 종료 및 재실행
        for proc in psutil.process_iter():
            try:
                if "Cloudflare WARP.exe" in proc.name():
                    proc.kill()
            except psutil.AccessDenied:
                pass
        subprocess.Popen([r"C:\Program Files\Cloudflare\Cloudflare WARP\Cloudflare WARP.exe"])

        self.progress_callback.emit(65, "윈도우 파일 (3) 다운로드 중...")

        file3 = m.find('downloaded_files.z02')
        # 파일3 다운로드
        m.download_url(url3)
        m.download(file3, temp_folder, "downloaded_files.z02")

        # Cloudflare WARP 프로세스 종료 및 재실행
        for proc in psutil.process_iter():
            try:
                if "Cloudflare WARP.exe" in proc.name():
                    proc.kill()
            except psutil.AccessDenied:
                pass

        self.progress_callback.emit(75, "윈도우 파일 (3) 다운로드 완료")

        # 받은 분할 압축 파일을 통합
        split_files = [f"downloaded_files.{ext}" for ext in ["zip", "z01", "z02"]]
        with open(os.path.join(temp_folder, "downloaded_files.zip"), 'wb') as f:
            for split_file in split_files:
                split_file_path = os.path.join(temp_folder, split_file)
                with open(split_file_path, 'rb') as part_file:
                    f.write(part_file.read())

        self.progress_callback.emit(80, "윈도우 파일 통합 완료")

        with zipfile.ZipFile(os.path.join(temp_folder, "downloaded_files.zip"), 'r') as zip_ref:
            zip_ref.extractall(exe_drive_root)

        self.progress_callback.emit(90, "윈도우 파일 복사 완료")

    # 부팅 메뉴 설정 (메소드 정의)
    def setup_bootmenu(self):
        try:
            # bcdedit 명령어를 창을 띄우지 않고 실행하기 위해 shell=True 사용
            self.progress_callback.emit(95, "윈도우 시작 메뉴 추가 중...")

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
            subprocess.run("bcdedit /set !guid! osdevice ramdisk=[%~d0]\\sources\\boot.wim,{ramdiskoptions}",
                           shell=True)
            subprocess.run("set bios=exe", shell=True)
            subprocess.run("findstr /c:'Detected boot environment: BIOS' 'C:\\Windows\\Panther\\setupact.log'",
                           shell=True)
            subprocess.run("if errorlevel 1 set bios=efi", shell=True)
            subprocess.run("bcdedit /set !guid! path \\windows\\system32\\winload.!bios!", shell=True)
            subprocess.run("bcdedit /set !guid! systemroot \\windows", shell=True)
            subprocess.run("bcdedit /set !guid! winpe yes", shell=True)
            subprocess.run("bcdedit /set !guid! detecthal yes", shell=True)
            subprocess.run("bcdedit /displayorder !guid! /addlast", shell=True)
            subprocess.run("bcdedit /timeout 10", shell=True)

            self.progress_callback.emit(98, "윈도우 시작 메뉴 추가완료.")
        except Exception as e:
            print(f"부팅 메뉴 설정 중 오류가 발생했습니다: {e}")

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