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
    parser.add_argument('--template', type=str, help="Path to the input PPTX template file (required unless --test is used)")
    parser.add_argument('--slides', type=str, default=None, 
                        help='특정 슬라이드만 처리할 경우 슬라이드 번호를 콤마로 구분하여 입력 (예: "4,7-10,12")')
    parser.add_argument('--output', type=str, help="Path to save the generated PPTX file (optional)")
    parser.add_argument('--profile', type=str, default='default', help="Akamai .edgerc profile name")
    parser.add_argument('--account-key', type=str, help="Account Switch Key (optional)")
    parser.add_argument('--service-name', type=str, help="Global Service Name")
    parser.add_argument('--start', type=str, help="Start Date (e.g., 2026-08-01T00:00:00Z)")
    parser.add_argument('--end', type=str, help="End Date (e.g., 2026-08-31T23:59:59Z)")
    parser.add_argument('--cpcodes', type=str, help="Comma-separated list of CPCodes or a path to a CSV/TXT file")
    parser.add_argument('--ignore-cache', action='store_true', help="Ignore all cache and force API calls")
    parser.add_argument('--ignore-cache-slides', type=str, help="Ignore cache for specific slides")
    
    parser.add_argument('--max-cpcodes', type=int, default=20, help="Max number of CPCodes per API call (default: 20)")
    
    parser.add_argument('--test', type=str, help="Path to JSON file to test API directly to stdout")
    
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

def parse_slides_argument(slides_str):
    """'4,10-12,17' 형태의 문자열을 [4, 10, 11, 12, 17] 리스트로 파싱합니다."""
    if not slides_str:
        return []
        
    target_slides = []
    for part in slides_str.split(','):
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                target_slides.extend(range(start, end + 1))
            except ValueError:
                logger.warning(f"잘못된 슬라이드 범위 형식입니다: {part}")
        elif part.isdigit():
            target_slides.append(int(part))
            
    # 중복 제거 및 오름차순 정렬
    return sorted(list(set(target_slides)))


if __name__ == "__main__":
    args = parse_arguments()

    # Orchestrator로 넘기기 전에 main.py에서 깔끔한 List로 가공합니다.
    args.target_slides_list = parse_slides_argument(args.slides)

    # CP Code 스마트 파싱 적용
    if args.cpcodes:
        args.cpcodes = resolve_cpcodes(args.cpcodes)
        
    # ==========================================
    # [TEST MODE] --test 인자가 주어졌을 때의 동작
    # ==========================================
    if args.test:
        import sys
        import json
        
        logger.info(f"🔧 RUNNING IN TEST MODE: Reading directive from {args.test}")
        try:
            with open(args.test, 'r', encoding='utf-8') as f:
                directive = json.load(f)
                
            api_client = AkamaiAPIWrapper(
                profile=args.profile, 
                fallback_account_key=args.account_key,
                max_cpcodes=args.max_cpcodes
            )
            
            endpoint = directive.get("endpoint", "")
            if args.start: endpoint = endpoint.replace("{START_DATE}", args.start)
            if args.end: endpoint = endpoint.replace("{END_DATE}", args.end)
            
            payload = directive.get("payload", {})
            payload_str = json.dumps(payload)
            if args.start: payload_str = payload_str.replace("{START_DATE}", args.start)
            if args.end: payload_str = payload_str.replace("{END_DATE}", args.end)
            
            payload = json.loads(payload_str)
            cpcode_list = [c.strip() for c in args.cpcodes.split(',')] if args.cpcodes else []
            
            # 👇 수정한 부분
            def replace_cpcode_in_dict(d):
                for k, v in d.items():
                    if isinstance(v, dict):
                        replace_cpcode_in_dict(v)
                    elif isinstance(v, list):
                        new_list = []
                        for item in v:
                            if item == "{CPCODE_LIST}":
                                new_list.extend(cpcode_list)
                            elif isinstance(item, dict):
                                replace_cpcode_in_dict(item)
                                new_list.append(item)
                            else:
                                new_list.append(item)
                        d[k] = new_list
            replace_cpcode_in_dict(payload)
            # 👆 여기까지
            
            logger.info(f"Target Endpoint: {endpoint}")
            logger.info("Executing direct API Request (ignoring cache)...")
            
            response = api_client.request(
                method=directive.get("method", "POST"),
                endpoint=endpoint,
                payload=payload,
                ignore_cache=True
            )
            
            print("\n" + "="*60)
            print("✅ AKAMAI API RAW RESPONSE STDOUT".center(60))
            print("="*60)
            print(json.dumps(response, indent=2, ensure_ascii=False))
            print("="*60 + "\n")
            
            sys.exit(0)
            
        except Exception as e:
            logger.error(f"Test mode failed: {e}")
            sys.exit(1)
            
    # ==========================================
    # [NORMAL MODE] 기본 PPTX 생성 로직
    # ==========================================
    if not args.template:
        logger.error("the following arguments are required: --template (unless --test is used)")
        exit(1)

    final_output = args.output if args.output else get_auto_output_filename()
    args.output = final_output
    
    logger.info(f"Starting script with template: {args.template}")
    logger.info(f"Output will be saved to: {args.output}")
    
    try:
        # API Client 생성 시 max_cpcodes 전달
        api_client = AkamaiAPIWrapper(
            profile=args.profile, 
            fallback_account_key=args.account_key,
            max_cpcodes=args.max_cpcodes  # <-- 수정된 변수명 적용
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