import os
import sys
import time
import random
import logging
import configparser
import requests
from urllib.parse import urlparse
from akamai.edgegrid import EdgeGridAuth

from cache_manager import CacheManager
from utils import normalize_string

logger = logging.getLogger(__name__)

class AkamaiAPIWrapper:
    def __init__(self, profile='default', edgerc_path='~/.edgerc', fallback_account_key=None):
        self.edgerc_path = os.path.expanduser(edgerc_path)
        self.profile = profile
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

    def request(self, method, endpoint, payload=None, ignore_cache=False):
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