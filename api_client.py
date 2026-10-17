import os
import sys
import time
import random
import copy
import logging
import configparser
import requests
from urllib.parse import urlparse
from akamai.edgegrid import EdgeGridAuth

from cache_manager import CacheManager
from utils import normalize_string

logger = logging.getLogger(__name__)

class AkamaiAPIWrapper:
    def __init__(self, profile='default', edgerc_path='~/.edgerc', fallback_account_key=None, max_cpcodes=20):
        self.edgerc_path = os.path.expanduser(edgerc_path)
        self.profile = profile
        self.max_cpcodes = max_cpcodes  # 💡 [수정됨] 변수명 변경
        self.session = requests.Session()
        
        try:
            # Rule 1: configparser를 사용한 명시적 Host 추출
            config = configparser.ConfigParser()
            config.read(self.edgerc_path)
            self.baseurl = f"https://{config.get(self.profile, 'host')}"
            
            self.session.auth = EdgeGridAuth.from_edgerc(self.edgerc_path, self.profile)
            
            self.account_switch_key = fallback_account_key
            if not self.account_switch_key:
                if config.has_option(self.profile, 'account_key'):
                    self.account_switch_key = config.get(self.profile, 'account_key')
                elif config.has_option(self.profile, 'accountSwitchKey'):
                    self.account_switch_key = config.get(self.profile, 'accountSwitchKey')
            
            # Rule 1: Key가 누락되었을 경우 대화형 프롬프트 활성화
            if not self.account_switch_key and sys.stdin.isatty():
                DEFAULT_ACCOUNT_ID = "1-166U8ZT"
                user_input = input(f"Enter Target Customer Account ID [Press Enter for default: {DEFAULT_ACCOUNT_ID}]: ").strip()
                self.account_switch_key = user_input if user_input else DEFAULT_ACCOUNT_ID
            
            # Rule 1: 대소문자 엄격 유지
            if self.account_switch_key:
                self.account_switch_key = normalize_string(self.account_switch_key, preserve_case=True)
            
            self.cache_manager = CacheManager()    
            
            logger.info(f"Initialized EdgeGrid auth with profile: {self.profile}")
            if self.account_switch_key:
                logger.info(f"Account Switch Key loaded: {self.account_switch_key}")
                
        except Exception as e:
            logger.error(f"Failed to initialize Akamai EdgeGrid authentication: {e}")
            raise

    # 💡 파라미터에 payload=None 이 추가되었습니다.
    def _merge_reporting_responses(self, responses, payload=None):
        if not responses: return {}
        if len(responses) == 1: return responses[0]

        base_resp = copy.deepcopy(responses[0])
        
        if "metadata" not in base_resp or "data" not in base_resp:
            logger.warning("⚠️ [Data Merge Warning] API 응답이 표준 구조를 따르지 않습니다.")
            return base_resp

        group_by_keys = base_resp.get("metadata", {}).get("groupBy", [])
        if not group_by_keys and payload and "dimensions" in payload:
            group_by_keys = payload.get("dimensions", [])

        # 💡 [핵심 수정] 비율 지표의 원본(Base) 지표 이름 유연한 매핑 (V1/V2 완벽 호환)
        original_keys = responses[0].get("data", [{}])[0].keys() if responses[0].get("data") else []
        pct_to_base = {}
        
        for k in original_keys:
            if "PERCENT" in k.upper() or "PCT" in k.upper():
                base_metric = None
                k_lower = k.lower()
                prefix = ""
                
                # 1. 접미사(Suffix)에 따라 유연하게 원본 키워드 추출
                if k_lower.endswith("percentage"): prefix = k[:-10]    # edgeHitsPercentage -> edgeHits
                elif k_lower.endswith("percent"): prefix = k[:-7]      # successfulHitsPercent -> successfulHits
                elif k_lower.endswith("pct"): prefix = k[:-3]          # xxxPct -> xxx
                
                if prefix:
                    # 2. V2 API (예: edgeHitsSum) 대응
                    if prefix + "Sum" in original_keys:
                        base_metric = prefix + "Sum"
                    # 3. V1 API (예: successfulHits) 대응
                    elif prefix in original_keys:
                        base_metric = prefix
                        
                if base_metric: 
                    pct_to_base[k] = base_metric

        true_global_totals = {}
        
        # 💡 [핵심 추가] 각 청크별로 (성공Hits + 실패Hits)를 모두 포함한 진짜 분모를 역산하여 누적
        for resp in responses:
            for pct_metric, base_metric in pct_to_base.items():
                chunk_true_total = 0.0
                max_val = -1
                best_row = None
                # 부동소수점 오차를 최소화하기 위해 해당 청크에서 값이 가장 큰 Row를 역산 기준으로 삼음
                for row in resp.get("data", []):
                    try:
                        v, p = float(row.get(base_metric, 0)), float(row.get(pct_metric, 0))
                        if v > max_val and p > 0:
                            max_val, best_row = v, row
                    except (ValueError, TypeError): pass
                
                if best_row:
                    v, p = float(best_row[base_metric]), float(best_row[pct_metric])
                    # 역산 공식: 진짜 총합 = 기준값 / (퍼센트 / 100)
                    chunk_true_total = v / (p / 100.0)
                
                true_global_totals[base_metric] = true_global_totals.get(base_metric, 0.0) + chunk_true_total

        merged_data = {}
        
        # 원시 데이터 그룹핑 및 합산
        for resp in responses:
            for row in resp.get("data", []):
                row_key = tuple(row.get(k, "") for k in group_by_keys)

                if row_key not in merged_data:
                    merged_data[row_key] = {k: row.get(k, "") for k in group_by_keys}

                for k, v in row.items():
                    if k in group_by_keys: continue
                    if "PERCENT" in k.upper() or "PCT" in k.upper() or "RATE" in k.upper(): continue

                    try:
                        merged_data[row_key][k] = merged_data[row_key].get(k, 0.0) + float(v)
                    except (ValueError, TypeError):
                        if k not in merged_data[row_key]:
                            merged_data[row_key][k] = v

        final_data_list = []

        # 퍼센트(비율) 재계산 시, 단순 합산값이 아닌 '진짜 총합(True Global Total)' 사용
        for row_key, row_dict in merged_data.items():
            for k in original_keys:
                # 1. 일반 분포 비율(%) 지표 처리 (진짜 분모 기반)
                if k in pct_to_base:
                    base_metric = pct_to_base[k]
                    global_denom = true_global_totals.get(base_metric, 0)
                    
                    if base_metric in row_dict and global_denom > 0:
                        row_dict[k] = (row_dict[base_metric] / global_denom) * 100
                    else:
                        row_dict[k] = 0.0

            # 💡 [핵심 추가] 오프로드율(Offload) 전용 재계산 로직 (V1/V2 호환 및 음수 방어)
            if "offloadedBytesPercentage" in original_keys:
                edge = float(row_dict.get("edgeBytesSum", row_dict.get("edgeBytes", 0)))
                origin = float(row_dict.get("originBytesSum", row_dict.get("originBytes", 0)))
                calc_pct = ((edge - origin) / edge) * 100 if edge > 0 else 0.0
                row_dict["offloadedBytesPercentage"] = max(0.0, calc_pct)
                
            if "offloadedHitsPercentage" in original_keys:
                edge_h = float(row_dict.get("edgeHitsSum", row_dict.get("edgeHits", 0)))
                origin_h = float(row_dict.get("originHitsSum", row_dict.get("originHits", 0)))
                calc_pct_h = ((edge_h - origin_h) / edge_h) * 100 if edge_h > 0 else 0.0
                row_dict["offloadedHitsPercentage"] = max(0.0, calc_pct_h)

            # 숫자 포맷팅 및 문자열 변환
            for k, v in row_dict.items():
                if isinstance(v, float):
                    row_dict[k] = str(int(v)) if v.is_integer() else str(v)

            final_data_list.append(row_dict)

        base_resp["data"] = final_data_list
        return base_resp

    def _extract_cpcodes_from_payload(self, payload):
        """Payload에서 CP Code 리스트와 삽입 경로를 찾습니다."""
        if not payload: return [], None
        
        # v1 형식: "objectIds" 배열
        if "objectIds" in payload and payload.get("objectType", "").lower() == "cpcode":
            return payload["objectIds"], "v1"
            
        # v2 형식: "filters" 배열 내부
        if "filters" in payload:
            for i, f in enumerate(payload["filters"]):
                if f.get("dimensionName", "").lower() == "cpcode":
                    return f.get("expressions", []), f"v2_{i}"
                    
        return [], None

    def request(self, method, endpoint, payload=None, ignore_cache=False, **kwargs):
        cp_codes, payload_type = self._extract_cpcodes_from_payload(payload)
        
        if cp_codes and len(cp_codes) > self.max_cpcodes:
            responses = []
            
            for i in range(0, len(cp_codes), self.max_cpcodes):
                chunk = cp_codes[i:i + self.max_cpcodes]
                chunk_payload = copy.deepcopy(payload)
                
                if payload_type == "v1":
                    chunk_payload["objectIds"] = chunk
                elif payload_type.startswith("v2_"):
                    filter_idx = int(payload_type.split("_")[1])
                    chunk_payload["filters"][filter_idx]["expressions"] = chunk
                    
                resp = self._make_single_request(method, endpoint, chunk_payload, ignore_cache, **kwargs)
                responses.append(resp)
                
            # 💡 [핵심 수정 3] 병합 함수에 payload를 넘겨주어 dimensions 키를 찾을 수 있게 합니다.
            return self._merge_reporting_responses(responses, payload=payload)
            
        return self._make_single_request(method, endpoint, payload, ignore_cache, **kwargs)

    def _make_single_request(self, method, endpoint, payload, ignore_cache, **kwargs):
        url = self.baseurl + endpoint if endpoint.startswith('/') else f"{self.baseurl}/{endpoint}"
        
        # Rule 1: params 딕셔너리를 활용한 안전한 쿼리 스트링 주입
        req_params = {}
        if self.account_switch_key:
            req_params['accountSwitchKey'] = self.account_switch_key
        
        cached_data = self.cache_manager.get_cached_response(method, url, payload, req_params, ignore_cache)
        if cached_data:
            return cached_data
        
        retries = 0
        max_retries = 5
        while retries < max_retries:
            try:
                response = self.session.request(
                    method=method.upper(),
                    url=url,
                    json=payload,
                    params=req_params
                )
                
                if response.status_code == 202:
                    final_data = self._handle_async_polling(response, url, req_params)
                    self.cache_manager.save_response(method, url, payload, req_params, final_data)
                    return final_data

                if response.status_code == 200:
                    final_data = response.json()
                    self.cache_manager.save_response(method, url, payload, req_params, final_data)
                    return final_data
                        
                if response.status_code == 429:
                    self._handle_rate_limit(response, retries)
                    retries += 1
                    continue
                
                response.raise_for_status()

            except requests.exceptions.RequestException as e:
                if 'response' in locals() and response is not None and response.status_code not in [429]:
                    # 1. 기본 에러 메시지 구성
                    error_msg = f"{response.status_code} Server Error for url: {url}\n{response.text}"
                    
                    # 2. 5xx 에러 발생 시 CP Code 개수를 검사하여 조치 가이드 추가
                    if response.status_code >= 500:
                        cp_codes, _ = self._extract_cpcodes_from_payload(payload)
                        num_cpcodes = len(cp_codes) if cp_codes else 0
                        
                        if num_cpcodes > 10:  # 한 번에 10개를 초과하여 요청한 경우 과부하로 간주
                            hint_msg = (
                                f"\n💡 [조치 권고] Akamai 백엔드 서버의 연산 한계 초과 (Timeout/Crash)가 발생했습니다. "
                                f"(현재 단일 호출에 포함된 CP Code: {num_cpcodes}개)\n"
                                f"👉 이 차트/표는 건너뜁니다. 명령어 실행 시 '--max-cpcodes 10' 이하로 값을 낮추어 다시 실행해 보세요."
                            )
                            error_msg += hint_msg
                    
                    # Orchestrator의 try...except가 잡아서 로그를 예쁘게 출력하도록 Exception 발생
                    raise Exception(error_msg)
                    
                elif 'response' not in locals() or response is None:
                    logger.error(f"Network error: {e}")
                    raise
                
        raise Exception("Max retries exceeded while calling Akamai API.")

    def _handle_rate_limit(self, response, retry_count):
        retry_after = response.headers.get('Retry-After')
        base_sleep = float(retry_after) if retry_after else 2 ** retry_count
        sleep_duration = base_sleep + random.uniform(0.5, 2.0)
        logger.warning(f"Rate limit (429) reached. Sleeping for {sleep_duration:.2f} seconds before retry...")
        time.sleep(sleep_duration)

    def _handle_async_polling(self, response, original_url, req_params):
        logger.info("Received 202 Accepted. Starting asynchronous polling...")
        poll_url = response.headers.get('Location')
        if not poll_url:
            resp_data = response.json()
            query_id = resp_data.get('queryId')
            if query_id:
                base_path = urlparse(original_url).path.replace('/data', '')
                poll_url = f"{base_path}/queries/{query_id}"
            else:
                raise Exception("Cannot determine polling URL from 202 response.")
        
        poll_url = self.baseurl + poll_url if poll_url.startswith('/') else f"{self.baseurl}/{poll_url}"
        
        max_polls = 20
        polls = 0
        while polls < max_polls:
            retry_after = float(response.headers.get('Retry-After', 5.0))
            time.sleep(retry_after)
            poll_resp = self.session.get(poll_url, params=req_params)
            
            if poll_resp.status_code == 200:
                return poll_resp.json()
            elif poll_resp.status_code == 202:
                response = poll_resp 
                polls += 1
            elif poll_resp.status_code == 429:
                self._handle_rate_limit(poll_resp, polls)
                polls += 1
            else:
                poll_resp.raise_for_status()
                
        raise Exception("Max polling attempts reached for asynchronous query.")