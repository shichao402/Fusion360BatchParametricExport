import sys
import os
import adsk.core, adsk.fusion, json
from .LogUtils import LogUtils

plugin_dir = os.path.dirname(os.path.abspath(__file__))
if plugin_dir not in sys.path:
    sys.path.insert(0, plugin_dir)

class ConfigUtils:

    @staticmethod
    def write_configs_to_excel(file_path: str, configs: list, parameters: list):
        """
        将配置写入Excel文件
        :param file_path: Excel文件路径
        :param configs: 配置列表
        :param parameters: 参数列表
        """
        try:
            from openpyxl import Workbook
            from openpyxl.utils import get_column_letter
            
            wb = Workbook()
            ws = wb.active
            ws.title = "导出配置"
            
            # 写入表头
            headers = ['导出格式', '自定义名称']
            for param in parameters:
                headers.append(param['name'])
            
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header)
            
            # 写入配置数据
            for row, config in enumerate(configs, 2):
                ws.cell(row=row, column=1, value=config.get('format', 'step'))
                ws.cell(row=row, column=2, value=config.get('name', ''))
                
                for col, param in enumerate(parameters, 3):
                    param_value = config.get('parameters', {}).get(param['name'], param['expression'])
                    ws.cell(row=row, column=col, value=param_value)
            
            # 调整列宽
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 15
            
            wb.save(file_path)
            LogUtils.info(f'配置已保存到Excel文件: {file_path}')
            return True
        except Exception as e:
            LogUtils.error(f'保存Excel文件失败: {str(e)}')
            return False

    @staticmethod
    def read_configs_from_excel(file_path: str, parameters: list):
        """
        从Excel文件读取配置
        :param file_path: Excel文件路径
        :param parameters: 参数列表
        :return: 配置列表或None
        """
        try:
            from openpyxl import load_workbook
            
            if not os.path.exists(file_path):
                LogUtils.error(f'Excel文件不存在: {file_path}')
                return None
            
            wb = load_workbook(file_path)
            ws = wb.active
            
            configs = []
            
            # 从第二行开始读取数据（第一行是表头）
            for row in range(2, ws.max_row + 1):
                format_val = ws.cell(row=row, column=1).value
                name_val = ws.cell(row=row, column=2).value
                
                if not format_val and not name_val:
                    continue  # 跳过空行
                
                config = {
                    'format': format_val or 'step',
                    'name': name_val or '',
                    'parameters': {}
                }
                
                # 读取参数值
                for col, param in enumerate(parameters, 3):
                    param_value = ws.cell(row=row, column=col).value
                    if param_value is not None:
                        config['parameters'][param['name']] = str(param_value)
                
                configs.append(config)
            
            LogUtils.info(f'从Excel文件读取了 {len(configs)} 个配置: {file_path}')
            return configs
        except Exception as e:
            LogUtils.error(f'读取Excel文件失败: {str(e)}')
            return None

    @staticmethod
    def create_excel_template(file_path: str, parameters: list):
        """
        创建/补全Excel模板文件，保留原有数据，只补充缺失的参数列。
        :param file_path: Excel文件路径
        :param parameters: 参数列表
        """
        try:
            from openpyxl import Workbook, load_workbook
            from openpyxl.utils import get_column_letter
            from openpyxl.styles import Font, PatternFill, Alignment
            import os
            param_names = [p['name'] for p in parameters]
            param_exprs = {p['name']: p['expression'] for p in parameters}
            headers = ['导出格式', '自定义名称'] + param_names
            # 如果文件存在，读取原有数据
            if os.path.exists(file_path):
                wb = load_workbook(file_path)
                ws = wb.active
                # 读取原有表头
                old_headers = [cell.value for cell in ws[1]]
                # 计算缺失的参数列
                missing_params = [p for p in param_names if p not in old_headers]
                # 如果没有缺失，直接返回
                if not missing_params:
                    LogUtils.info(f'Excel模板已存在且无缺失列: {file_path}')
                    return True
                # 补充缺失的表头
                new_headers = old_headers + missing_params
                for i, header in enumerate(new_headers, 1):
                    cell = ws.cell(row=1, column=i, value=header)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                # 补充每行缺失的参数列
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    old_row_len = len(old_headers)
                    for idx, param in enumerate(missing_params, start=old_row_len+1):
                        row[0].parent.cell(row=row[0].row, column=idx, value=param_exprs.get(param, ''))
                # 调整列宽
                for col in range(1, len(new_headers) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 15
                wb.save(file_path)
                LogUtils.info(f'Excel模板已补全缺失列: {file_path}')
                return True
            # 文件不存在，创建新模板
            wb = Workbook()
            ws = wb.active
            ws.title = "导出配置"
            # 表头
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            # 示例行
            example_row = ['step', 'default'] + [param_exprs[name] for name in param_names]
            for col, value in enumerate(example_row, 1):
                ws.cell(row=2, column=col, value=value)
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 15
            wb.save(file_path)
            LogUtils.info(f'Excel模板已创建: {file_path}')
            return True
        except Exception as e:
            LogUtils.error(f'创建Excel模板失败: {str(e)}')
            return False 