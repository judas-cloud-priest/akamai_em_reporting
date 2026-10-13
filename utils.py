import os
import re

# Rule 8: 스크립트 절대 경로 확보
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()

# Rule 2: Data Cleaning & Normalization
def normalize_string(s, preserve_case=True):
    """문자열의 불필요한 공백/특수문자를 제거하고 숫자형 소수점을 정규화합니다."""
    if not isinstance(s, str):
        if isinstance(s, float) and s.is_integer():
            s = str(int(s))
        else:
            s = str(s)
    s = s.strip().replace('\xa0', ' ').replace('\r', '').replace('\n', '').replace('"', '').replace("'", "")
    return s if preserve_case else s.lower()

def get_auto_output_filename(base_name="output.pptx"):
    base_path = os.path.join(SCRIPT_DIR, base_name)
    if not os.path.exists(base_path):
        return base_path
    base, ext = os.path.splitext(base_path)
    counter = 1
    while True:
        new_name = f"{base}({counter}){ext}"
        if not os.path.exists(new_name):
            return new_name
        counter += 1
