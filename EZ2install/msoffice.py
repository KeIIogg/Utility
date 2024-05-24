import os
import sys
import xml.etree.ElementTree as ET
import tempfile
import subprocess
from PyQt5.QtWidgets import QMessageBox

class MSOfficeInstaller:

    def __init__(self, version, selected_apps):
        self.version = version
        self.selected_apps = selected_apps

    def generate_config_xml(self):
        if self.version == "MS Office 2019 Proplus":
            product_id = "ProPlus2019Volume"
            channel = "PerpetualVL2019"
        elif self.version == "MS Office 2021 Proplus":
            product_id = "ProPlus2021Volume"
            channel = "PerpetualVL2021"
        elif self.version == "MS Office 2024 Proplus":
            product_id = "ProPlus2024Volume"
            channel = "Production_LTSC2024"
        else:
            raise ValueError("올바른 버전을 선택하세요.")

        config = ET.Element("Configuration")
        add_element = ET.SubElement(config, "Add", OfficeClientEdition="64", Channel=channel)
        product_element = ET.SubElement(add_element, "Product", ID=product_id, PIDKEY="")
        ET.SubElement(product_element, "Language", ID="ko-kr")

        excluded_apps = ["Word", "Teams", "Excel", "Outlook", "PowerPoint", "OneDrive", "Access", "OneNote"]
        for app in excluded_apps:
            if app in self.selected_apps:
                ET.SubElement(product_element, "ExcludeApp", ID=app)

        ET.SubElement(config, "Display", Level="Full", AcceptEULA="TRUE")
        ET.SubElement(config, "Updates", Enabled="TRUE")
        ET.SubElement(config, "Property", Name="AUTOACTIVATE", Value="1")

        temp_folder = tempfile.gettempdir()

        # XML 파일 경로 설정
        xml_path = os.path.join(temp_folder, "config.xml")

        # config를 ElementTree로 변환하여 XML 파일로 저장
        tree = ET.ElementTree(config)
        tree.write(xml_path, encoding="utf-8", xml_declaration=True)




    def execute_setup(self):

        temp_folder = tempfile.gettempdir()
        xml_path = os.path.join(temp_folder, "config.xml")
        temp_folder = tempfile.gettempdir()
        if xml_path:
            setup_exe_path = os.path.join(temp_folder, "setup.exe")
            if os.path.exists(setup_exe_path):
                setup_cmd = f'cd "{os.path.dirname(xml_path)}" && "{setup_exe_path}" config config.xml'
                subprocess.Popen(setup_cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            else:
                raise FileNotFoundError("setup.exe 파일이 존재하지 않습니다.")


    def activate_office(self):

        if self.version == "MS Office 2019 Proplus":
            cmd = r"""
            cd /d "%ProgramFiles%\Microsoft Office\Office16"
            for /f %%x in ('dir /b "..\root\Licenses16\ProPlus2019VL*.xrm-ms"') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
            cscript ospp.vbs /setprt:1688
            cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
            cscript ospp.vbs /act
            """
        elif self.version == "MS Office 2021 Proplus":
            cmd = r"""
            cd /d "%ProgramFiles%\Microsoft Office\Office16"
            for /f %%x in ('dir /b "..\root\Licenses16\ProPlus2021VL*.xrm-ms"') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
            cscript ospp.vbs /setprt:1688
            cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
            cscript ospp.vbs /act
            """
        elif self.version == "MS Office 2024 Proplus":
            cmd = r"""
            cd /d "%ProgramFiles%\Microsoft Office\Office16"
            for /f %%x in ('dir /b "..\root\Licenses16\ProPlus2024PreviewVL*.xrm-ms"') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
            cscript ospp.vbs /setprt:1688
            cscript ospp.vbs /inpkey:2TDPW-NDQ7G-FMG99-DXQ7M-TX3T2
            cscript ospp.vbs /act
            """

        # PowerShell 실행 명령어 설정
        powershell_process = subprocess.Popen(
            ["powershell.exe", "-Command", cmd],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            stdin=subprocess.PIPE,
            shell=True,
            text=True  # 텍스트 모드 사용 (Python 3.7+)
        )

        # 실행 결과 받기 (stdout, stderr)
        stdout, stderr = powershell_process.communicate()

        if stderr:
            print(f"PowerShell 명령어 실행 중 오류 발생: {stderr}")
        else:
            print("Office 활성화가 완료되었습니다.")




