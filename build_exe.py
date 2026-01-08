"""
PyInstaller를 사용하여 exe 파일 빌드
"""
import subprocess
import sys

def build():
    # PyInstaller 명령어
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=DCR_Converter",
        "--onefile",           # 단일 exe 파일로 생성
        "--windowed",          # 콘솔 창 없이 실행 (GUI 앱)
        "--noconfirm",         # 기존 빌드 폴더 자동 덮어쓰기
        "--clean",             # 빌드 전 캐시 정리
        # 필요한 모듈들 포함
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtWidgets",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.styles",
        "--hidden-import=openpyxl.utils",
        "--hidden-import=openpyxl.formatting",
        "--hidden-import=openpyxl.formatting.rule",
        "--hidden-import=openpyxl.worksheet.filters",
        "--hidden-import=chardet",
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=xlrd",
        # files.json은 exe와 같은 위치에 동적으로 생성됨
        # 진입점
        "main.py"
    ]
    
    print("Building exe file...")
    print(" ".join(cmd))
    
    result = subprocess.run(cmd, cwd=".")
    
    if result.returncode == 0:
        print("\n" + "="*50)
        print("Build successful!")
        print("Exe file: dist/DCR_Converter.exe")
        print("="*50)
    else:
        print("\nBuild failed!")
    
    return result.returncode

if __name__ == "__main__":
    build()

