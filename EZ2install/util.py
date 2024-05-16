import os
import subprocess
import urllib.request
import zipfile
import tempfile
from mega import Mega

#파일 정리하기

def download_and_install_hnc2022():
    #url = "https://drive.google.com/uc?id=1Ca-XiHtkDD_X_obpRoTB8meYquuTEKGG"
    url = "https://mega.nz/file/ZKkiFJ5C#Va51ThSM0d4NBlJmjpJGPcLuh_SLC3ucgZIfpAziov8"
    temp_folder = tempfile.gettempdir()
    download_path = os.path.join(temp_folder, "HNC2022.zip")

    try:
        #gdown.download(url, download_path, quiet=False)

        mega = Mega()
        m = mega.login("seonghoonxg@gmail.com", "kosh1108k!")
        m.download_url(url)
        m.download(file, temp_folder, "HNC2022.zip")


        print(f"HNC2022 설치 파일이 다운로드되었습니다.")

        # 압축 해제
        extract_folder = os.path.join(temp_folder, "HNC2022")
        with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)

        # 사일런트 모드로 설치 실행
        setup_exe_path = os.path.join(extract_folder, "install", "setup.exe")
        subprocess.Popen([setup_exe_path, "/S"])

        # 파일 삭제
        os.remove(zip_filepath)
        os.remove(extract_folder)

    except Exception as e:
        print(f"Error occurred: {e}")


def download_and_install_bandizip():
    # 다운로드 URL 및 파일명 지정
    url = "https://dl.bandisoft.com/bandizip.std/BANDIZIP-SETUP-STD-X64.EXE"
    filename = "Bandizip-Setup.exe"

    # 시스템의 기본 다운로드 폴더 경로 가져오기
    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])
        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")





def download_and_install_everything():
    url = "https://www.voidtools.com/Everything-1.4.1.1024.x64-Setup.exe"
    filename = "Everything-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_ccleaner():
    url = "https://bits.avcdn.net/productfamily_CCLEANER/insttype_FREE/platform_WIN_PIR/installertype_ONLINE/build_RELEASE/cookie_mmm_ccl_003_999_a8e_m"
    filename = "CCleaner-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_chrome():
    url = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B0B26EAD3-2E55-C5EF-B295-1C60269E6DF5%7D%26lang%3Dko%26browser%3D4%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"
    filename = "Chrome-Setup.exe"

    temp_folder = tempfile.gettempdir()
    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")


def download_and_install_nespdf():
    url = "http://download.nespdf.com/nespdf_free_x64_1.3.EXE"
    filename = "NESpdf-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_pdanet():
    url = "https://pdanet.co/bin/PdaNetA5232b.exe"
    filename = "PDANet-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")


def download_and_install_samsung_usb():
    url = "https://downloadcenter.samsung.com/content/SW/202401/20240111141255930/SAMSUNG_USB_Driver_for_Mobile_Phones.exe"
    filename = "Samsung-USB-Driver-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_adb_usb():
    try:
        # 다운로드 URL 및 파일명 지정
        url = "https://dl-ssl.google.com/android/repository/latest_usb_driver_windows.zip"

        temp_folder = tempfile.gettempdir()
        # 다운로드 파일 경로
        zip_file_path = os.path.join(temp_folder, "ADB-USB-Driver.zip")

        # 압축 해제 폴더 경로
        extract_folder = os.path.join(temp_folder, "ADB-USB-Driver")

        # Android USB 드라이버 설치 스크립트 파일 경로
        inf_file_path = os.path.join(extract_folder, "android_winusb.inf")

        # 파일 다운로드
        urllib.request.urlretrieve(url, zip_file_path)

        # 압축 해제
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)

        # Android USB 드라이버 설치
        subprocess.Popen(["pnputil", "/i", "/a", inf_file_path])

        # 파일 삭제
        os.remove(zip_file_path)
        os.remove(extract_folder)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_samsung_switch():
    url = "https://downloadcenter.samsung.com/content/SW/202312/20231219105000341/Smart.Switch.PC_setup.exe"
    filename = "Samsung-Switch-Setup.exe"

    temp_folder = tempfile.gettempdir()
    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")

def download_and_install_bandiview():
    url = "https://dl.bandisoft.com/bandiview/BANDIVIEW-SETUP-X64.EXE"
    filename = "Bandiview-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")

def download_and_install_mobaxterm():
    try:
        # 다운로드 URL 및 파일명 지정
        url = "https://download.mobatek.net/2412024041614011/MobaXterm_Installer_v24.1.zip"
        zip_filename = "MobaXterm-Setup.zip"
        installer_filename = "MobaXterm/MobaXterm_installer_24.1.msi"

        temp_folder = tempfile.gettempdir()

        # 파일 다운로드 경로 지정
        zip_file_path = os.path.join(temp_folder, zip_filename)

        # 압축 해제 폴더 경로 지정
        extract_folder = os.path.join(temp_folder, "MobaXterm")

        # MobaXterm 설치 파일 경로 지정
        installer_file_path = os.path.join(temp_folder, installer_filename)

        # 파일 다운로드
        urllib.request.urlretrieve(url, zip_file_path)

        # 압축 해제
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)

        # MobaXterm 설치
        subprocess.Popen(["msiexec", "/i", installer_file_path, "/quiet", "/qn"])

        # 파일 삭제
        os.remove(zip_file_path)
        os.remove(extract_folder)

    except Exception as e:
        print(f"Error occurred: {e}")

def download_and_install_anydesk():
    url = "https://anydesk.com/ko/downloads/thank-you?dv=win_exe"
    filename = "AnyDesk-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")

def download_and_install_kakaotalk():
    url = "https://app-pc.kakaocdn.net/talk/win32/KakaoTalk_Setup.exe"
    filename = "KakaoTalk-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")

def download_and_install_lg_gram():
    try:
        # 다운로드 및 설치에 필요한 정보
        url = "https://gscs-b2c.lge.com/downloadFile?fileId=oSxUlNAKD1EOIDjwbhBEQ"
        temp_folder = tempfile.gettempdir()
        download_path = os.path.join(temp_folder, "LG-Gram-Setup.zip")
        extract_folder = os.path.join(temp_folder, "LG-Gram-Setup")
        install_exe = "LG Update Installer.exe"

        # 파일 다운로드
        urllib.request.urlretrieve(url, download_path)

        # 압축 해제
        with zipfile.ZipFile(download_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)

        # 설치 파일 경로
        install_path = os.path.join(extract_folder, install_exe)

        # 설치 실행 (사일런트 모드)
        subprocess.Popen([install_path, "/S"])

        # 파일 삭제
        os.remove(download_path)
        os.remove(extract_folder)

    except Exception as e:
        print(f"Error occurred: {e}")



def download_and_install_mop_printer():
    url = "https://mop.pirodown.win/ebp332_x64.exe"
    filename = "MopPrinter-Setup.exe"

    temp_folder = tempfile.gettempdir()

    # 다운로드 경로 지정
    util_download_path = os.path.join(temp_folder, filename)

    try:
        # 다운로드
        urllib.request.urlretrieve(url, util_download_path)

        # 실행
        subprocess.Popen([util_download_path, "/S"])

        # 파일 삭제
        os.remove(util_download_path)

    except Exception as e:
        print(f"Error occurred: {e}")


# 윈도우 업데이트 끄기
def disable_windows_update():
    # Windows 업데이트 서비스 및 관련 프로세스를 중지하고 비활성화합니다.
    subprocess.run(['sc', 'config', 'wuauserv', 'start=disabled'], shell=True)
    subprocess.run(['sc', 'config', 'UsoSvc', 'start=disabled'], shell=True)
    subprocess.run(['sc', 'stop', 'wuauserv'], shell=True)
    subprocess.run(['sc', 'stop', 'UsoSvc'], shell=True)

    # WaaSMedicAgent.exe 파일에 대한 권한 변경 작업을 수행합니다.
    subprocess.run(['takeown', '/f', 'WaaSMedicAgent.exe'], shell=True)
    subprocess.run(['icacls', 'WaaSMedicAgent.exe', '/deny', 'system:RX'], shell=True)
    subprocess.run(['icacls', 'WaaSMedicAgent.exe', '/grant', 'administrators:F'], shell=True)
    subprocess.run(['icacls', 'WaaSMedicAgent.exe', '/setowner', 'NT Service\\TrustedInstaller'], shell=True)
    subprocess.run(['icacls', 'WaaSMedicAgent.exe', '/grant:r', 'administrators:RX'], shell=True)

# Onedrive 삭제
def remove_onedrive():
    # OneDrive 프로세스 강제 종료
    subprocess.run(['taskkill', '/f', '/im', 'OneDrive.exe'], shell=True)

    # OneDrive 제거 프로그램 실행하여 OneDrive 제거
    subprocess.run([os.path.join(os.environ['SystemRoot'], 'SysWOW64', 'OneDriveSetup.exe'), '/uninstall'], shell=True)

    # 사용자 OneDrive 폴더 및 관련 폴더 삭제
    subprocess.run(['rd', os.path.expanduser('~') + '\\OneDrive', '/s', '/q'], shell=True)
    subprocess.run(['rd', os.path.expandvars('%LocalAppData%\\Microsoft\\OneDrive'), '/s', '/q'], shell=True)
    subprocess.run(['rd', os.path.expandvars('%ProgramData%\\Microsoft OneDrive'), '/s', '/q'], shell=True)
    subprocess.run(['rd', 'C:\\OneDriveTemp', '/s', '/q'], shell=True)

    # OneDrive 바로 가기 파일 삭제
    subprocess.run(['del', os.path.expandvars('%USERPROFILE%\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\OneDrive.lnk'), '/s', '/f', '/q'], shell=True)

    # 레지스트리 항목 제거
    subprocess.run(['REG', 'Delete', 'HKEY_CLASSES_ROOT\\CLSID\\{018D5C66-4533-4307-9B53-224DE2ED1FE6}', '/f'], shell=True)
    subprocess.run(['REG', 'Delete', 'HKEY_CLASSES_ROOT\\Wow6432Node\\CLSID\\{018D5C66-4533-4307-9B53-224DE2ED1FE6}', '/f'], shell=True)
    subprocess.run(['REG', 'ADD', 'HKEY_CLASSES_ROOT\\CLSID\\{018D5C66-4533-4307-9B53-224DE2ED1FE6}', '/v', 'System.IsPinnedToNameSpaceTree', '/d', '0', '/t', 'REG_DWORD', '/f'], shell=True)

def windows_defender_off():
    # Windows Defender 비활성화 관련 레지스트리 설정
    subprocess.run(['reg', 'add', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows Defender', '/v', 'DisableAntiSpyware', '/t', 'REG_DWORD', '/d', '1', '/f'], shell=True)
    subprocess.run(['REG', 'DELETE', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows Defender', '/v', 'DisableAntiSpyware', '/f'], shell=True)
    subprocess.run(['reg', 'add', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\CloudContent', '/v', 'DisableWindowsSpotlightFeatures', '/t', 'REG_DWORD', '/d', '1', '/f'], shell=True)
    subprocess.run(['reg', 'add', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\CloudContent', '/v', 'DisableThirdPartySuggestions', '/t', 'REG_DWORD', '/d', '1', '/f'], shell=True)

def office_update_off():
    # Office 자동 업데이트 비활성화 관련 레지스트리 설정
    subprocess.run(['reg', 'add', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Office\\16.0\\common\\officeupdate', '/v', 'enableautomaticupdates', '/t', 'REG_DWORD', '/d', '0', '/f'], shell=True)
    subprocess.run(['reg', 'add', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Office\\17.0\\common\\officeupdate', '/v', 'enableautomaticupdates', '/t', 'REG_DWORD', '/d', '0', '/f'], shell=True)
    subprocess.run(['reg', 'add', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Office\\18.0\\common\\officeupdate', '/v', 'enableautomaticupdates', '/t', 'REG_DWORD', '/d', '0', '/f'], shell=True)

# 다운로드 및 설치 테스트

# Bandizip 다운로드 및 설치 테스트
#download_and_install_bandizip()

# Everything 다운로드 및 설치 테스트
#download_and_install_everything()

# CCleaner 다운로드 및 설치 테스트
#download_and_install_ccleaner()

# Chrome 다운로드 및 설치 테스트
#download_and_install_chrome()

# HNC2022 다운로드 및 설치 테스트
#download_and_install_hnc2022()

# NESpdf 다운로드 및 설치 테스트
#download_and_install_nespdf()

# PDANet 다운로드 및 설치 테스트
#download_and_install_pdanet()

# Samsung USB Driver 다운로드 및 설치 테스트
#download_and_install_samsung_usb()

# ADB USB Driver 다운로드 및 설치 테스트
#download_and_install_adb_usb()

# Samsung Switch 다운로드 및 설치 테스트
#download_and_install_samsung_switch()

# Bandiview 다운로드 및 설치 테스트
#download_and_install_bandiview()

# MobaXterm 다운로드 및 설치 테스트
#download_and_install_mobaxterm()

# AnyDesk 다운로드 및 설치 테스트
#download_and_install_anydesk()

# KakaoTalk 다운로드 및 설치 테스트
#download_and_install_kakaotalk()

# LG Gram 다운로드 및 설치 테스트
#download_and_install_lg_gram()

# Mop Printer 다운로드 및 설치 테스트
#download_and_install_mop_printer()

# 윈도우 업데이트 끄기 테스트
#disable_windows_update()

# Onedrive 삭제 테스트
#remove_onedrive()
