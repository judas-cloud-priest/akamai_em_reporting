import os
import json
import time
import hashlib
from datetime import datetime, timezone, timedelta
from urllib.parse import urlparse, parse_qsl, urlencode

class CacheManager:
    def __init__(self, cache_dir='~/.akamai_report_cache'):
        self.cache_dir = os.path.expanduser(cache_dir)
        if not os.path.exists(self.cache_dir):
            os.makedirs(self.cache_dir)
            
    def _generate_cache_key(self, method, url, payload, params=None):
        parsed_url = urlparse(url)
        query_dict = dict(parse_qsl(parsed_url.query))
        if params:
            query_dict.update(params)
            
        sorted_query = sorted(query_dict.items())
        normalized_url = f"{parsed_url.path}?{urlencode(sorted_query)}"
        
        payload_str = ""
        if payload:
            payload_str = json.dumps(payload, sort_keys=True)
            
        raw_key = f"{method}|{normalized_url}|{payload_str}"
        return hashlib.sha256(raw_key.encode('utf-8')).hexdigest()

    def _extract_end_date(self, url):
        parsed_url = urlparse(url)
        query_params = dict(parse_qsl(parsed_url.query))
        end_str = query_params.get('end')
        
        if end_str:
            try:
                end_str = end_str.replace('Z', '+00:00')
                return datetime.fromisoformat(end_str)
            except ValueError:
                pass
        return None

    def get_cached_response(self, method, url, payload, params=None, ignore_cache=False):
        if ignore_cache:
            return None
            
        cache_key = self._generate_cache_key(method, url, payload, params)
        cache_path = os.path.join(self.cache_dir, f"{cache_key}.json")
        
        if not os.path.exists(cache_path):
            return None
            
        file_mtime = os.path.getmtime(cache_path)
        current_time = time.time()
        file_age_seconds = current_time - file_mtime
        
        if file_age_seconds < 900:  
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)
                
        end_date = self._extract_end_date(url)
        if end_date:
            now_utc = datetime.now(timezone.utc)
            data_age = now_utc - end_date
            if data_age < timedelta(days=7):
                return None
        else:
            return None
                
        with open(cache_path, 'r', encoding='utf-8') as f:
            return json.load(f)
            
    def save_response(self, method, url, payload, params, response_data):
        cache_key = self._generate_cache_key(method, url, payload, params)
        cache_path = os.path.join(self.cache_dir, f"{cache_key}.json")
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(response_data, f, ensure_ascii=False, indent=2)