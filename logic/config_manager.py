"""
설정 관리 모듈
파일 경로 정보를 JSON으로 저장/로드
"""

import json
import os
import sys


def get_app_dir() -> str:
    """
    애플리케이션 디렉토리 반환
    - exe 실행 시: exe 파일이 있는 디렉토리
    - 개발 중: 프로젝트 루트 디렉토리
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller로 빌드된 exe 실행 시
        return os.path.dirname(sys.executable)
    else:
        # 개발 중 (python main.py)
        return os.path.dirname(os.path.dirname(__file__))


# 설정 파일 경로 (exe와 같은 위치에 files.json 생성)
APP_DIR = get_app_dir()
CONFIG_FILE = os.path.join(APP_DIR, "files.json")


def save_file_paths(net_file: str, vendorspec_file: str, 
                    partpin_file: str, outfile: str,
                    etching_dir: str = "", form_outfile: str = "",
                    dimension_file: str = "", dimension_sheet: str = "",
                    lslusl_file: str = "", merged_file: str = "",
                    operator_name: str = "",
                    item_name: str = "", item_code: str = "",
                    output_base_dir: str = "") -> bool:
    """
    파일 경로들을 JSON으로 저장
    
    Args:
        net_file: .NET 파일 경로
        vendorspec_file: vendorspec 파일 경로
        partpin_file: partpin 파일 경로
        outfile: 출력 파일 경로
        etching_dir: etching 디렉토리 경로 (Form Measurement용)
        form_outfile: Form Measurement 출력 파일 경로
        dimension_file: dimension 파일 경로
        dimension_sheet: dimension 시트 이름
        lslusl_file: LSLUSL 파일 경로
        merged_file: merged 파일 경로
        operator_name: 작업자 이름
        item_name: 아이템 이름
        item_code: 아이템 코드
        output_base_dir: 출력 기본 디렉토리
        
    Returns:
        저장 성공 여부
    """
    try:
        # 디렉토리가 없으면 생성 (일반적으로 exe와 같은 위치이므로 이미 존재함)
        config_dir = os.path.dirname(CONFIG_FILE)
        if config_dir and not os.path.exists(config_dir):
            os.makedirs(config_dir)
        
        config = {
            "net_file": net_file,
            "vendorspec_file": vendorspec_file,
            "partpin_file": partpin_file,
            "outfile": outfile,
            "etching_dir": etching_dir,
            "form_outfile": form_outfile,
            "dimension_file": dimension_file,
            "dimension_sheet": dimension_sheet,
            "lslusl_file": lslusl_file,
            "merged_file": merged_file,
            "operator_name": operator_name,
            "item_name": item_name,
            "item_code": item_code,
            "output_base_dir": output_base_dir
        }
        
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False


def load_file_paths() -> dict:
    """
    JSON에서 파일 경로들을 로드
    
    Returns:
        파일 경로 딕셔너리 (파일이 없으면 빈 값들)
    """
    default_config = {
        "net_file": "",
        "vendorspec_file": "",
        "partpin_file": "",
        "outfile": "DCR_format_yamaha.xlsx",
        "etching_dir": "",
        "form_outfile": "Form_measurement_result.xlsx",
        "dimension_file": "",
        "dimension_sheet": "",
        "lslusl_file": "",
        "merged_file": "",
        "operator_name": "",
        "item_name": "",
        "item_code": "",
        "output_base_dir": ""
    }
    
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 기본값과 병합 (누락된 키가 있을 경우 대비)
            for key in default_config:
                if key not in config:
                    config[key] = default_config[key]
            
            return config
        else:
            return default_config
    except Exception as e:
        print(f"Error loading config: {e}")
        return default_config

