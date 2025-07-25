import os
import json
import tempfile

class CacheUtils:
    @staticmethod
    def get_cache_file_path():
        try:
            temp_dir = tempfile.gettempdir()
            cache_file = os.path.join(temp_dir, 'Fusion360BatchParametricExport_cache.json')
            return cache_file
        except:
            return None

    @staticmethod
    def save_cached_export_path(path):
        try:
            if not path:
                return
            cache_file = CacheUtils.get_cache_file_path()
            cache_data = {}
            if cache_file and os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
            cache_data['export_path'] = path
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except:
            pass

    @staticmethod
    def load_cached_export_path():
        try:
            cache_file = CacheUtils.get_cache_file_path()
            if cache_file and os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
                    return cache_data.get('export_path', '')
        except:
            pass
        return ''

    @staticmethod
    def save_cached_excel_path(path):
        try:
            if not path:
                return
            cache_file = CacheUtils.get_cache_file_path()
            cache_data = {}
            if cache_file and os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
            cache_data['excel_path'] = path
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except:
            pass

    @staticmethod
    def load_cached_excel_path():
        try:
            cache_file = CacheUtils.get_cache_file_path()
            if cache_file and os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
                    return cache_data.get('excel_path', '')
        except:
            pass
        return '' 