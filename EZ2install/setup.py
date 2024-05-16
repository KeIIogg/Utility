import sys
from cx_Freeze import setup, Executable

# 파일 포함 정보 설정
include_files = [
    ("d:/e2i/e2i_main.ui", "e2i_main.ui"),
    ("d:/e2i/ez2install.ico", "ez2install.ico"),
    ("d:/e2i/KeIIog/setup.exe", "setup.exe"),
    ("d:/e2i/KeIIog/uninstall.exe", "uninstall.exe")
]

# cx_Freeze 옵션 설정
options = {
    'build_exe': {
        'include_files': include_files,
        'includes': [
            'PyQt5.QtGui', 'PyQt5.QtCore', 'PyQt5.QtWidgets', 'xml.etree.ElementTree',
            'urllib.request', 'zipfile', 'mega'
        ],
        'packages': ['sys', 'os', 'shutil', 'logging', 'tempfile', 'subprocess'],
        'include_msvcr': True,
        'excludes': ['tkinter'],  # tkinter 라이브러리 제외
    }
}

# 실행 파일 생성 설정
executables = [
    Executable(
        "d:/e2i/main.py",
        target_name="e2i_app.exe",
        icon="d:/e2i/ez2install.ico",  # 실행 파일 아이콘 설정
        base="Win32GUI"  # GUI 실행 파일로 빌드
    )
]

# setup 호출
setup(
    name="e2i_app",
    version="1.0",
    description="EZ2install Application",
    options=options,
    executables=executables
)
