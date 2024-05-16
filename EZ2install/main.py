import sys
import os
import shutil
import logging
import tempfile
import subprocess
import resources
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QLineEdit, QVBoxLayout, QPushButton, QMessageBox
from windows import WindowsInstaller, WindowsActivator
from msoffice import MSOfficeInstaller
from util import (
    download_and_install_bandizip, download_and_install_everything, download_and_install_ccleaner,
    download_and_install_chrome, download_and_install_hnc2022, download_and_install_nespdf,
    download_and_install_pdanet, download_and_install_samsung_usb, download_and_install_adb_usb,
    download_and_install_samsung_switch, download_and_install_bandiview, download_and_install_mobaxterm,
    download_and_install_anydesk, download_and_install_kakaotalk, download_and_install_lg_gram, download_and_install_mop_printer, remove_onedrive, disable_windows_update, windows_defender_off, office_update_off
)

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

        self.setFixedSize(705, 540)


        pixmap = QPixmap(':/tan.png')
        self.tani.setPixmap(pixmap)
        self.tani.setScaledContents(True)

        self.add_exception_list()
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


    def add_exception_list(self):
        """
        여러 개의 PowerShell 명령어를 한 번에 실행하고 PowerShell 창을 표시하지 않습니다.
        """
        # 실행할 PowerShell 명령어들을 리스트로 정의
        powershell_commands = [
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147685180 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147735507 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147736914 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147743522 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147734094 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147743421 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147765679 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147783203 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 251873 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 213927 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ThreatIDDefaultAction_Ids 2147722906 -ThreatIDDefaultAction_Actions Allow -Force',
            'Add-MpPreference -ExclusionPath "$env:SystemRoot\\Temp" -Force'
        ]

        # PowerShell 명령어들을 줄바꿈으로 연결하여 하나의 명령어 문자열로 구성
        powershell_cmd = "; ".join(powershell_commands)

        # PowerShell 실행 명령어 설정
        powershell_process = subprocess.Popen(
            ["powershell.exe", "-Command", powershell_cmd],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            stdin=subprocess.PIPE,
            shell=True,
            text=True  # 텍스트 모드 사용 (Python 3.7+)
        )

        # PowerShell 프로세스의 결과 받기 (stdout, stderr)
        stdout, stderr = powershell_process.communicate()

        if stderr:
            print(f"PowerShell 명령어 실행 중 오류 발생: {stderr}")
        else:
            print("PowerShell 명령어 실행 완료.")


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

    def windows_install_clicked(self):
        print("Windows Install Clicked")
        selected_option = self.win_drop.currentText()
        try:
            installer = WindowsInstaller()
            installer.install_windows(selected_option)
            installer.setup_bootmenu()
            QMessageBox.information(None, "작업 완료", "Windows 설치가 완료되었습니다.")
        except Exception as e:
            print(f"윈도우 설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"윈도우 설치 중 오류가 발생했습니다: {str(e)}")

    def windows_activate_clicked(self):
        print("Windows Activate Clicked")
        try:
            activator = WindowsActivator()
            activator.activate_windows()
            QMessageBox.information(None, "작업 완료", "윈도우 인증이 완료되었습니다.")
        except Exception as e:
            print(f"윈도우 인증 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"윈도우 인증 중 오류가 발생했습니다: {str(e)}")

    def office_install_clicked(self):
        print("Office Install Clicked")
        version = self.office_drop.currentText()
        selected_apps = [
            "Word" if self.chk_word.isChecked() else None,
            "Teams" if self.chk_teams.isChecked() else None,
            "Excel" if self.chk_excel.isChecked() else None,
            "Outlook" if self.chk_outlook.isChecked() else None,
            "PowerPoint" if self.chk_powerpoint.isChecked() else None,
            "OneDrive" if self.chk_onedrive.isChecked() else None,
            "Access" if self.chk_access.isChecked() else None,
            "OneNote" if self.chk_onenote.isChecked() else None
        ]
        selected_apps = [app for app in selected_apps if app]

        try:
            office_installer = MSOfficeInstaller(version, selected_apps)
            office_installer.generate_config_xml()
            office_installer.execute_setup()
            QMessageBox.information(None, "작업 완료", "office 설치가 완료되었습니다.")
        except Exception as e:
            print(f"설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"설치 중 오류가 발생했습니다: {str(e)}")

    def office_activate_clicked(self):
        print("Office Activation Clicked")
        version = self.office_drop.currentText()

        try:
            office_installer = MSOfficeInstaller(version)
            office_installer.activate_office()
            QMessageBox.information(None, "작업 완료", "office 설치가 완료되었습니다.")
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
        try:
            if self.chk_bandizip.isChecked():
                download_and_install_bandizip()
            if self.chk_everything.isChecked():
                download_and_install_everything()
            if self.chk_ccleaner.isChecked():
                download_and_install_ccleaner()
            if self.chk_chrome.isChecked():
                download_and_install_chrome()
            if self.chk_hnc2022.isChecked():
                download_and_install_hnc2022()
            if self.chk_nespdf.isChecked():
                download_and_install_nespdf()
            if self.chk_pdanet.isChecked():
                download_and_install_pdanet()
            if self.chk_samusb.isChecked():
                download_and_install_samsung_usb()
            if self.chk_adbdriver.isChecked():
                download_and_install_adb_usb()
            if self.chk_samswitch.isChecked():
                download_and_install_samsung_switch()
            if self.chk_honeyview.isChecked():
                download_and_install_bandiview()
            if self.chk_mobaxterm.isChecked():
                download_and_install_mobaxterm()
            if self.chk_anydesk.isChecked():
                download_and_install_anydesk()
            if self.chk_kakao.isChecked():
                download_and_install_kakaotalk()
            if self.chk_printpdf.isChecked():
                download_and_install_mop_printer()
            if self.chk_lggram.isChecked():
                download_and_install_lg_gram()

            QMessageBox.information(None, "작업 완료", "util 설치가 완료되었습니다.")
        except Exception as e:
            print(f"유틸리티 설치 중 오류가 발생했습니다: {str(e)}")
            QMessageBox.warning(None, "오류", f"유틸리티 설치 중 오류가 발생했습니다: {str(e)}")

    def onedrive_delete_clicked(self):
        print("Onedrive Delete Clicked")
        remove_onedrive()
        QMessageBox.information(None, "작업 완료", "Onedrive 삭제가 완료되었습니다.")

    def windows_update_off_clicked(self):
        print("Windows Update Off Clicked")
        disable_windows_update()
        QMessageBox.information(None, "작업 완료", "윈도우 업데이트가 중지되었습니다.")

    def windows_defender_off_clicked(self):
        print("Windows Defender Off Clicked")
        windows_defender_off()
        QMessageBox.information(None, "작업 완료", "윈도우 디펜더 중지되었습니다.")

    def office_update_off_clicked(self):
        print("Office Update Off Clicked")
        office_update_off()
        QMessageBox.information(None, "작업 완료", "오피스 업데이트가 중지되었습니다.")

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
