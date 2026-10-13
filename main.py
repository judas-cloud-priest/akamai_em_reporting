import os
import csv
import argparse
import logging
from utils import get_auto_output_filename
from api_client import AkamaiAPIWrapper
from pptx_orchestrator import PPTXOrchestrator

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def parse_arguments():
    parser = argparse.ArgumentParser(description="Akamai Traffic Report Generator")
    parser.add_argument('--template', type=str, required=True, help="Path to the input PPTX template file")
    parser.add_argument('--output', type=str, help="Path to save the generated PPTX file (optional)")
    parser.add_argument('--profile', type=str, default='default', help="Akamai .edgerc profile name")
    parser.add_argument('--account-key', type=str, help="Account Switch Key (optional)")
    parser.add_argument('--service-name', type=str, help="Global Service Name")
    parser.add_argument('--start', type=str, help="Start Date (e.g., 2026-08-01T00:00:00Z)")
    parser.add_argument('--end', type=str, help="End Date (e.g., 2026-08-31T23:59:59Z)")
    parser.add_argument('--cpcodes', type=str, help="Comma-separated list of CPCodes or a path to a CSV/TXT file")
    parser.add_argument('--ignore-cache', action='store_true', help="Ignore all cache and force API calls")
    parser.add_argument('--ignore-cache-slides', type=str, help="Comma-separated slide numbers to ignore cache")
    return parser.parse_args()

def resolve_cpcodes(cpcodes_input):
    """
    사용자의 입력값이 직접 입력한 콤마 구분 문자열인지, 파일 경로인지 식별하여
    순수한 CP Code 문자열 목록(콤마 구분)으로 평탄화(Flatten)하여 반환합니다.
    """
    if not cpcodes_input:
        return None
        
    # '~/' 등의 사용자 홈 디렉토리 경로 확장
    file_path = os.path.expanduser(cpcodes_input.strip())
    
    # 1. 입력값이 실제로 존재하는 파일인 경우 (CSV, TXT, 확장자 없는 파일 등)
    if os.path.isfile(file_path):
        extracted_cpcodes = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                # 파일 내의 모든 텍스트를 읽어 CSV 형태로 파싱 (한 줄에 여러 개 있거나 여러 줄로 나뉘어 있어도 처리됨)
                for row in csv.reader(f):
                    for item in row:
                        val = item.strip()
                        # CSV 헤더(예: 'CPCODE', 'Name') 등을 무시하고 순수 숫자(CP Code)만 안전하게 추출
                        if val.isdigit():
                            extracted_cpcodes.append(val)
            
            logger.info(f"Successfully loaded {len(extracted_cpcodes)} CPCodes from file: {file_path}")
            return ",".join(extracted_cpcodes)
        except Exception as e:
            logger.error(f"Failed to read CPCode file {file_path}: {e}")
            return None
            
    # 2. 파일이 아니라 직접 콤마로 구분된 문자열로 입력된 경우
    else:
        # 띄어쓰기 등 불순물 제거 및 순수 숫자만 필터링
        extracted_cpcodes = [x.strip() for x in cpcodes_input.split(',') if x.strip().isdigit()]
        if extracted_cpcodes:
            logger.info(f"Parsed {len(extracted_cpcodes)} CPCodes from command-line argument.")
        return ",".join(extracted_cpcodes)

if __name__ == "__main__":
    args = parse_arguments()
    final_output = args.output if args.output else get_auto_output_filename()
    args.output = final_output
    
    # [핵심] cpcodes 인자를 가로채어 스마트 파싱 적용
    if args.cpcodes:
        args.cpcodes = resolve_cpcodes(args.cpcodes)
    
    logger.info(f"Starting script with template: {args.template}")
    logger.info(f"Output will be saved to: {args.output}")
    
    try:
        api_client = AkamaiAPIWrapper(
            profile=args.profile, 
            fallback_account_key=args.account_key
        )
        
        logger.info("Performing Fail-Fast validation...")
        api_client.request('GET', '/reporting-api/v2/reports')
        logger.info("Authentication successful. Commencing orchestration...")
        
        orchestrator = PPTXOrchestrator(
            template_path=args.template,
            output_path=args.output,
            api_client=api_client,
            global_args=args
        )
        orchestrator.process_presentation()
        
    except Exception as e:
        logger.error(f"Execution failed: {e}")
        exit(1)