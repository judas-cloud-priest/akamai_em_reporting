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
    def __init__(self, profile='default', edgerc_path='~/.edgerc', fallback_account_key=None, chunk_size=20):
        self.edgerc_path = os.path.expanduser(edgerc_path)
        self.profile = profile
        self.chunk_size = chunk_size  # 💡 [새로 추가] 기본값 20 또는 CLI 지정값 저장
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

    def _merge_reporting_responses(self, responses):
        if not responses: return {}
        if len(responses) == 1: return responses[0]

        base_resp = copy.deepcopy(responses[0])
        
        # 💡 [신규] 응답 구조 검증 및 경고
        if "metadata" not in base_resp or "data" not in base_resp:
            logger.warning("⚠️ [Data Merge Warning] API 응답이 표준 Reporting API 구조(metadata/data)를 따르지 않습니다. 병합을 건너뛰고 첫 번째 응답만 반환합니다.")
            return base_resp

        group_by_keys = base_resp.get("metadata", {}).get("groupBy", [])
        if not group_by_keys:
            logger.warning("⚠️ [Data Merge Warning] 응답 메타데이터에 'groupBy' 키가 없어 데이터를 식별할 수 없습니다. 병합이 불완전할 수 있습니다.")

        merged_data = {}
        total_sums = {}

        for resp in responses:
            for row in resp.get("data", []):
                row_key = tuple(row.get(k, "") for k in group_by_keys)

                if row_key not in merged_data:
                    merged_data[row_key] = {k: row.get(k, "") for k in group_by_keys}

                for k, v in row.items():
                    if k in group_by_keys: continue

                    if "PERCENT" in k.upper() or "PCT" in k.upper() or "RATE" in k.upper():
                        if "CACHE" in k.upper() or "OFFLOAD" in k.upper():
                            logger.warning(f"⚠️ [Data Merge Warning] '{k}' 복합 비율(Rate) 지표가 감지되었습니다. 로컬 병합 시 데이터 왜곡이 발생할 수 있습니다.")
                        continue

                    try:
                        val = float(v)
                        merged_data[row_key][k] = merged_data[row_key].get(k, 0.0) + val
                        total_sums[k] = total_sums.get(k, 0.0) + val
                    except (ValueError, TypeError):
                        if k not in merged_data[row_key]:
                            merged_data[row_key][k] = v

        final_data_list = []
        original_keys = responses[0].get("data", [{}])[0].keys() if responses[0].get("data") else []

        for row_key, row_dict in merged_data.items():
            for k in original_keys:
                if "PERCENT" in k.upper() or "PCT" in k.upper():
                    base_metric = None
                    if k.lower().endswith("percent"): base_metric = k[:-7]
                    elif k.lower().endswith("pct"): base_metric = k[:-3]

                    if base_metric and base_metric in row_dict and total_sums.get(base_metric, 0) > 0:
                        row_dict[k] = (row_dict[base_metric] / total_sums[base_metric]) * 100
                    else:
                        row_dict[k] = 0.0

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
        # CP Code 추출 시도
        cp_codes, payload_type = self._extract_cpcodes_from_payload(payload)
        
        # CP Code 개수가 지정된 chunk_size를 초과할 경우 분할 호출
        if cp_codes and len(cp_codes) > self.chunk_size:
            responses = []
            
            for i in range(0, len(cp_codes), self.chunk_size):
                chunk = cp_codes[i:i + self.chunk_size]
                chunk_payload = copy.deepcopy(payload)
                
                # 추출된 경로에 맞게 청크 삽입
                if payload_type == "v1":
                    chunk_payload["objectIds"] = chunk
                elif payload_type.startswith("v2_"):
                    filter_idx = int(payload_type.split("_")[1])
                    chunk_payload["filters"][filter_idx]["expressions"] = chunk
                    
                resp = self._make_single_request(method, endpoint, chunk_payload, ignore_cache, **kwargs)
                responses.append(resp)
                
            return self._merge_reporting_responses(responses)
            
        # chunk_size 이하이거나 일반 요청인 경우 그대로 호출
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
                    logger.error(f"HTTP Error {response.status_code}: {response.text}")
                    raise
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