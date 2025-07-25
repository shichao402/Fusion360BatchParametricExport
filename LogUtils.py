import datetime
import os

class LogUtils:
    LOG_FILE = os.path.join(os.path.dirname(__file__), 'Fusion360BatchExport.log')

    @staticmethod
    def log(msg, level='INFO'):
        log_str = f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{level}] {msg}"
        print(log_str, flush=True)
        try:
            with open(LogUtils.LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(log_str + '\n')
        except Exception as e:
            # 如果写文件失败，也输出到控制台
            print(f"[LogUtils ERROR] 写日志文件失败: {e}", flush=True)

    @staticmethod
    def error(msg):
        LogUtils.log(msg, level='ERROR')

    @staticmethod
    def warn(msg):
        LogUtils.log(msg, level='WARN')

    @staticmethod
    def info(msg):
        LogUtils.log(msg, level='INFO') 