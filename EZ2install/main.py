import sys
import os
import shutil
import logging
import tempfile
import subprocess
import winreg as reg
import resources
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QLineEdit, QVBoxLayout, QPushButton, QMessageBox, QProgressBar, QLabel
from windows import WindowsInstaller, WindowsActivator
from msoffice import MSOfficeInstaller
from util import (
    download_and_install_bandizip, download_and_install_everything, download_and_install_ccleaner,
    download_and_install_chrome, download_and_install_hnc2022, download_and_install_nespdf,
    download_and_install_pdanet, download_and_install_samsung_usb, download_and_install_adb_usb,
    download_and_install_samsung_switch, download_and_install_bandiview, download_and_install_mobaxterm,
    download_and_install_anydesk, download_and_install_kakaotalk, download_and_install_lg_gram, download_and_install_mop_printer, remove_onedrive, disable_windows_update, windows_defender_off, office_update_off
)
class WindowsInstallThread(QThread):
    finished = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):
        selected_option = self.parent().win_drop.currentText()
        try:
            installer = WindowsInstaller(selected_option)
            installer.progress_callback.connect(self.parent().update_progress)  # 프로그레스 바 업데이트를 위한 시그널 연결
            installer.install_windows()
            installer.setup_bootmenu()
            QMessageBox.information(None, "작업 완료", "Windows 설치가 완료.재부팅 후 설치메뉴를 선택해주세요.")
        except Exception as e:
            print(f"윈도우 설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"윈도우 설치 중 오류가 발생했습니다: {str(e)}")

        finally:
            self.finished.emit()
class OfficeInstallThread(QThread):
    finished = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):

        version = self.parent().office_drop.currentText()
        selected_apps = [
            "Word" if self.parent().chk_word.isChecked() else None,
            "Teams" if self.parent().chk_teams.isChecked() else None,
            "Excel" if self.parent().chk_excel.isChecked() else None,
            "Outlook" if self.parent().chk_outlook.isChecked() else None,
            "PowerPoint" if self.parent().chk_powerpoint.isChecked() else None,
            "OneDrive" if self.parent().chk_onedrive.isChecked() else None,
            "Access" if self.parent().chk_access.isChecked() else None,
            "OneNote" if self.parent().chk_onenote.isChecked() else None
        ]
        selected_apps = [app for app in selected_apps if app]

        try:
            office_installer = MSOfficeInstaller(version, selected_apps)
            office_installer.generate_config_xml()
            office_installer.execute_setup()
            QMessageBox.information(None, "작업 완료", "Office 설치가 완료되었습니다.")
        except Exception as e:
            print(f"설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"설치 중 오류가 발생했습니다: {str(e)}")

        finally:
            self.finished.emit()
class UtilInstallThread(QThread):
    finished = pyqtSignal()

    def run(self):
        try:
            if self.parent().chk_bandizip.isChecked():
                download_and_install_bandizip()
            if self.parent().chk_everything.isChecked():
                download_and_install_everything()
            if self.parent().chk_ccleaner.isChecked():
                download_and_install_ccleaner()
            if self.parent().chk_chrome.isChecked():
                download_and_install_chrome()
            if self.parent().chk_hnc2022.isChecked():
                download_and_install_hnc2022()
            if self.parent().chk_nespdf.isChecked():
                download_and_install_nespdf()
            if self.parent().chk_pdanet.isChecked():
                download_and_install_pdanet()
            if self.parent().chk_samusb.isChecked():
                download_and_install_samsung_usb()
            if self.parent().chk_adbdriver.isChecked():
                download_and_install_adb_usb()
            if self.parent().chk_samswitch.isChecked():
                download_and_install_samsung_switch()
            if self.parent().chk_honeyview.isChecked():
                download_and_install_bandiview()
            if self.parent().chk_mobaxterm.isChecked():
                download_and_install_mobaxterm()
            if self.parent().chk_anydesk.isChecked():
                download_and_install_anydesk()
            if self.parent().chk_kakao.isChecked():
                download_and_install_kakaotalk()
            if self.parent().chk_printpdf.isChecked():
                download_and_install_mop_printer()
            if self.parent().chk_lggram.isChecked():
                download_and_install_lg_gram()

            QMessageBox.information(None, "작업 완료", "util 설치가 완료되었습니다.")

        except Exception as e:
            print(f"유틸리티 설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"유틸리티 설치 중 오류가 발생했습니다: {str(e)}")

        finally:
            self.finished.emit()
class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("EZ2install - login")
        self.setGeometry(100, 100, 300, 150)

        layout = QVBoxLayout()

        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setMaxLength(4)
        self.password_input.setPlaceholderText('비밀번호 입력(4자리 숫자)')

        layout.addWidget(self.password_input)
        login_button = QPushButton("Login", self)

        login_button.clicked.connect(self.check_password)
        layout.addWidget(login_button)

        self.setLayout(layout)

        screen_geometry = QApplication.desktop().availableGeometry()
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())
        self.move(window_geometry.topLeft())

    def check_password(self):
        entered_password = self.password_input.text()
        correct_password = "0308"

        if entered_password == correct_password:
            self.accept()
        else:
            QMessageBox.warning(self, "경고", "비밀번호가 틀렸습니다. 프로그램을 종료합니다.", QMessageBox.Ok)
            self.reject()
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        ui_path = os.path.join(os.path.dirname(__file__), 'e2i_main.ui')
        uic.loadUi(ui_path, self)

        self.setFixedSize(705, 570)


        pixmap = QPixmap(':/tan.png')
        self.tani.setPixmap(pixmap)
        self.tani.setScaledContents(True)


        self.extract_additional_files()

        self.win_install.clicked.connect(self.windows_install_clicked)
        self.win_auth.clicked.connect(self.windows_activate_clicked)
        self.office_install.clicked.connect(self.office_install_clicked)
        self.office_auth.clicked.connect(self.office_activate_clicked)
        self.office_uninstall.clicked.connect(self.office_uninstall_clicked)
        self.util_install.clicked.connect(self.util_install_clicked)
        self.delete_onedrive.clicked.connect(self.onedrive_delete_clicked)
        self.disable_windows_update.clicked.connect(self.windows_update_off_clicked)
        self.windows_defender_off.clicked.connect(self.windows_defender_off_clicked)
        self.office_update_off.clicked.connect(self.office_update_off_clicked)
        self.turbo_off.clicked.connect(self.turbo_off_clicked)

        self.progress_bar = QProgressBar(self)
        self.progress_label = QLabel(self)
        self.init_progress_bar()

    def init_progress_bar(self):
        self.progress_bar.setGeometry(15, 545, 690, 20)
        self.progress_label.setGeometry(310, 515, 290, 20)
        self.progress_bar.hide()

    def update_progress(self, value, text):
        self.progress_bar.setValue(value)
        self.progress_label.setText(text)

        if value > 0 and value < 100:
            self.progress_bar.show()  # 작업이 진행중일 때만 보이도록 설정
        else:
            self.progress_bar.hide()  # 작업이 완료되면 숨김

    def extract_additional_files(self):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        temp_folder = tempfile.gettempdir()

        setup_src = os.path.join(base_path, "KeIIog", "setup.exe")
        uninstall_src = os.path.join(base_path, "KeIIog", "uninstall.exe")


        setup_dest = os.path.join(temp_folder, "setup.exe")
        uninstall_dest = os.path.join(temp_folder, "uninstall.exe")

        if os.path.exists(setup_src) and os.path.exists(uninstall_src):
            try:
                shutil.copyfile(setup_src, setup_dest)
                shutil.copyfile(uninstall_src, uninstall_dest)

            except FileNotFoundError as e:
                logging.error(f"파일을 복사하는 동안 오류가 발생했습니다: {e}")
        else:
            logging.error("KeIIog 에 파일이 존재하지 않습니다.")


    def closeEvent(self, event):
        super().closeEvent(event)

        temp_folder = tempfile.gettempdir()
        setup_path = os.path.join(temp_folder, "setup.exe")
        uninstall_path = os.path.join(temp_folder, "uninstall.exe")

        try:
            if os.path.exists(setup_path):
                os.remove(setup_path)
                print(f"{setup_path} 파일이 삭제되었습니다.")
            else:
                print(f"{setup_path} 파일이 존재하지 않습니다.")
        except Exception as e:
            print(f"파일 삭제 중 오류 발생: {str(e)}")

        try:
            if os.path.exists(uninstall_path):
                os.remove(uninstall_path)
                print(f"{uninstall_path} 파일이 삭제되었습니다.")
            else:
                print(f"{uninstall_path} 파일이 존재하지 않습니다.")
        except Exception as e:
            print(f"파일 삭제 중 오류 발생: {str(e)}")

    def handle_windows_install_finished(self):
        self.update_progress(100, "Windows 설치 완료.")

    def handle_office_install_finished(self):
        self.update_progress(100, "Microsoft office 설치완료.")

    def handle_util_install_finished(self):
        self.update_progress(100, "유틸 설치완료.")

    def windows_install_clicked(self):
        print("Windows Install Clicked")
        thread = WindowsInstallThread(self)
        thread.finished.connect(self.handle_windows_install_finished)
        thread.start()
        self.update_progress(10, "Windows 설치 중...")

    def windows_activate_clicked(self):
        print("Windows Activate Clicked")
        try:
            self.update_progress(0, "윈도우 인증 진행 중...")
            activator = WindowsActivator()
            activator.activate_windows()
            self.update_progress(100, "윈도우 인증이 완료되었습니다.")
            QMessageBox.information(None, "작업 완료", "윈도우 인증이 완료되었습니다.")
        except Exception as e:
            print(f"윈도우 인증 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"윈도우 인증 중 오류가 발생했습니다: {str(e)}")

    def office_install_clicked(self):
        print("Office Install Clicked")

        try:
            thread = OfficeInstallThread(self)
            thread.finished.connect(self.handle_office_install_finished)
            thread.start()
            self.update_progress(10, "MS office 설치 중...")
            QMessageBox.information(None, "작업 완료", "office 설치가 완료되었습니다.")
        except Exception as e:
            print(f"설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"설치 중 오류가 발생했습니다: {str(e)}")

    def office_activate_clicked(self):
        print("Office Activation Clicked")
        version = self.office_drop.currentText()

        try:
            self.update_progress(0, "오피스 인증 진행 중...")
            office_installer = MSOfficeInstaller(version)
            office_installer.activate_office()
            self.update_progress(100, "오피스 인증완료")
            QMessageBox.information(None, "작업 완료", "오피스 인증완료")
        except Exception as e:
            print(f"인증 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"인증 중 오류가 발생했습니다: {str(e)}")


    def office_uninstall_clicked(self):
        print("Office Uninstall Clicked")
        temp_folder = tempfile.gettempdir()
        uninstall_path = os.path.join(temp_folder, "uninstall.exe")

        if os.path.exists(uninstall_path):
            subprocess.Popen(uninstall_path, shell=True)
        else:
            QMessageBox.warning(None, "오류", "제거툴을 찾을 수 없습니다.")

    def util_install_clicked(self):
        print("Util Install Clicked")
        thread = UtilInstallThread(self)
        thread.finished.connect(self.handle_util_install_finished)
        thread.start()
        self.update_progress(10, "유틸 설치 중...")

    def onedrive_delete_clicked(self):
        print("Onedrive Delete Clicked")
        self.update_progress(0, "원드라이브 삭제 진행 중...")
        remove_onedrive()
        self.update_progress(100, "원드라이브 삭제 완료.")
        QMessageBox.information(None, "작업 완료", "원드라이브 삭제 완료.")

    def windows_update_off_clicked(self):
        print("Windows Update Off Clicked")
        self.update_progress(0, "윈도우 업데이트 중지 중...")
        disable_windows_update()
        self.update_progress(100, "윈도우 업데이트 중지 완료.")
        QMessageBox.information(None, "작업 완료", "윈도우 업데이트 중지 완료.")

    def windows_defender_off_clicked(self):
        print("Windows Defender Off Clicked")
        self.update_progress(0, "윈도우 디펜더 중지 중...")
        windows_defender_off()
        self.update_progress(100, "윈도우 디펜더 중지 완료.")
        QMessageBox.information(None, "작업 완료", "윈도우 디펜더 중지되었습니다.")

    def office_update_off_clicked(self):
        print("Office Update Off Clicked")
        self.update_progress(0, "오피스 업데이트 중지 중...")
        office_update_off()
        self.update_progress(100, "오피스 업데이트 중지 완료.")
        QMessageBox.information(None, "작업 완료", "오피스 업데이트가 중지되었습니다.")

    def turbo_off_clicked(self):
        # 레지스트리 편집: Attributes 속성 변경
        try:
            self.update_progress(0, "새로운 전원옵션 추가 중...")
            reg_path = r"SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\be337238-0d82-4146-a960-4f3749d470c7"
            reg_key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, reg_path, 0, reg.KEY_SET_VALUE)
            reg.SetValueEx(reg_key, 'Attributes', 0, reg.REG_DWORD, 2)
            reg.CloseKey(reg_key)
            self.update_progress(30, "새로운 전원옵션 추가 완료")
            print("Registry key 'Attributes' set to 2 successfully.")
        except Exception as e:
            print(f"Failed to set registry key: {e}")

        # 새 전원 관리 프로필 생성
        self.update_progress(50, "터보부스트 해제 중...")
        result = subprocess.run(['powercfg', '/create', 'MyCustomScheme', 'Custom power scheme for specific needs'],
                                capture_output=True, text=True)
        if result.returncode == 0:
            print("Power scheme 'MyCustomScheme' created successfully.")
            # 생성된 전원 관리 프로필의 GUID를 추출하기 위해 전체 프로필 목록을 다시 가져옵니다.
            list_result = subprocess.run(['powercfg', '/l'], capture_output=True, text=True)
            schemes = list_result.stdout.split('\n')
            for scheme in schemes:
                if 'MyCustomScheme' in scheme:
                    # GUID 추출
                    scheme_guid = scheme.split()[3]
                    print(f"GUID for 'MyCustomScheme' is {scheme_guid}")
        else:
            print(f"Failed to create power scheme 'MyCustomScheme': {result.stderr}")
            scheme_guid = None

        if scheme_guid:

            self.update_progress(60, "전원옵션 최적화 중...")
            # 새 프로필 활성화
            subprocess.run(['powercfg', '/s', scheme_guid])
            print(f"Power scheme set to {scheme_guid}")

            # 디스플레이 끄기 시간 설정
            subprocess.run(['powercfg', '/change', scheme_guid, 'monitor-timeout-ac', '0'])
            subprocess.run(['powercfg', '/change', scheme_guid, 'monitor-timeout-dc', '5'])
            print(f"Display off time set: AC=off, DC=5 min for scheme {scheme_guid}")
            self.update_progress(70, "전원옵션 최적화 중...")

            # 시스템 절전 모드 설정
            subprocess.run(['powercfg', '/change', scheme_guid, 'standby-timeout-ac', '0'])
            subprocess.run(['powercfg', '/change', scheme_guid, 'standby-timeout-dc', '10'])
            print(f"Sleep time set: AC=off, DC=10 min for scheme {scheme_guid}")

            # 하드 디스크 끄기 시간 설정
            subprocess.run(['powercfg', '/change', scheme_guid, 'disk-timeout-ac', '0'])
            subprocess.run(['powercfg', '/change', scheme_guid, 'disk-timeout-dc', '0'])
            print(f"Disk off time set: AC=off, DC=off for scheme {scheme_guid}")

            self.update_progress(80, "전원옵션 최적화 중...")

            # 프로세서 성능 상태 최대 값 설정
            subprocess.run(['powercfg', '/setacvalueindex', scheme_guid, 'sub_processor', 'PROCTHROTTLEMAX', '90'])
            subprocess.run(['powercfg', '/setdcvalueindex', scheme_guid, 'sub_processor', 'PROCTHROTTLEMAX', '90'])
            print(f"Processor max state set: AC=90%, DC=90% for scheme {scheme_guid}")

            # 프로세서 성능 상태 최소 값 설정
            subprocess.run(['powercfg', '/setacvalueindex', scheme_guid, 'sub_processor', 'PROCTHROTTLEMIN', '5'])
            subprocess.run(['powercfg', '/setdcvalueindex', scheme_guid, 'sub_processor', 'PROCTHROTTLEMIN', '5'])
            print(f"Processor min state set: AC=5%, DC=5% for scheme {scheme_guid}")

            self.update_progress(90, "전원옵션 최적화 중...")

            # 시스템 냉각 정책 설정
            # 0: Passive, 1: Active
            subprocess.run(['powercfg', '/setacvalueindex', scheme_guid, 'sub_processor', 'sysfan', '0'])
            subprocess.run(
                ['powercfg', '/setdcvalueindex', scheme_guid, 'sub_processor', 'sysfan', '0'])
            print(f"System cooling policy set: AC=Active, DC=Passive for scheme {scheme_guid}")

            # 프로세서 성능 강화 모드 설정
            # 0: Disabled, 1: Enabled
            subprocess.run(
                ['powercfg', '/setacvalueindex', scheme_guid, 'sub_processor', 'perfboostmode', '0'])  # AC에서 Enabled
            subprocess.run(
                ['powercfg', '/setdcvalueindex', scheme_guid, 'sub_processor', 'perfboostmode', '0'])  # DC에서 Enabled
            print(f"Processor performance boost mode set: AC=Enabled, DC=Enabled for scheme {scheme_guid}")

            # 하이버네이션 설정
            subprocess.run(['powercfg', '/hibernate', 'on'])
            print(f"Hibernate set to on for scheme {scheme_guid}")

            self.update_progress(100, "전원옵션 최적화완료")
        else:
            print("No valid GUID found for the power scheme.")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_window = MainWindow()

    login_dialog = LoginDialog()
    if login_dialog.exec_() == QDialog.Accepted:
        main_window = MainWindow()
        main_window.show()
        sys.exit(app.exec_())
    else:
        sys.exit(0)