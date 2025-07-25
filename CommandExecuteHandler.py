import adsk.core, adsk.fusion, traceback
import os
from .LogUtils import LogUtils
from .ConfigUtils import ConfigUtils
from .CacheUtils import CacheUtils

class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self, batch_exporter, handlers):
        super().__init__()
        self.batch_exporter = batch_exporter
        self.handlers = handlers

    def notify(self, args):
        try:
            inputs = args.command.commandInputs
            ui = adsk.core.Application.get().userInterface
            
            # 获取导出路径
            export_path = ''
            path_group = inputs.itemById('pathGroup')
            if path_group:
                export_path_input = path_group.children.itemById('exportPath')
                if export_path_input:
                    export_path = export_path_input.value
            else:
                export_path_input = inputs.itemById('exportPath')
                if export_path_input:
                    export_path = export_path_input.value
                    
            if not export_path or not os.path.exists(export_path):
                ui.messageBox('❌ 请选择有效的导出路径')
                return
                
            CacheUtils.save_cached_export_path(export_path)
            
            # 从Excel文件读取配置
            export_configs = self.collect_export_configs_from_excel(inputs)
            if not export_configs:
                ui.messageBox('❌ 请先创建Excel配置文件并添加至少一组导出配置')
                return
                
            # 执行批量导出
            self.batch_exporter.execute_batch_export(export_configs, export_path)
            
            ui.messageBox('✅ 导出完成！\n\n💡 提示：\n• 所有配置已成功导出\n• 每个零件已保存到对应子目录\n• 您可以继续编辑Excel文件进行新的导出')
            
        except Exception as e:
            LogUtils.error(f'执行导出时发生错误: {str(e)}')
            ui = adsk.core.Application.get().userInterface
            if ui:
                ui.messageBox(f'❌ 执行导出时发生错误:\n{str(e)}')

    def collect_export_configs_from_excel(self, inputs):
        """从Excel文件收集导出配置"""
        try:
            # 获取Excel文件路径
            excel_group = inputs.itemById('excelGroup')
            if excel_group:
                excel_path_input = excel_group.children.itemById('excelPath')
            else:
                excel_path_input = inputs.itemById('excelPath')
                
            if not excel_path_input or not excel_path_input.value.strip():
                LogUtils.error('未指定Excel配置文件路径')
                return []
                
            excel_path = excel_path_input.value.strip()
            
            if not os.path.exists(excel_path):
                LogUtils.error(f'Excel文件不存在: {excel_path}')
                return []
            
            # 从Excel文件读取配置
            configs = ConfigUtils.read_configs_from_excel(excel_path, self.batch_exporter.parameters)
            
            if not configs:
                LogUtils.error('Excel文件中没有找到有效配置')
                return []
            
            # 转换为导出格式
            export_configs = []
            for config in configs:
                export_config = {
                    'format': config.get('format', 'step').lower(),
                    'custom_name': config.get('name', ''),
                    'parameters': config.get('parameters', {})
                }
                
                # 验证必要字段
                if not export_config['custom_name']:
                    LogUtils.error('配置中自定义名称不能为空')
                    continue
                    
                export_configs.append(export_config)
            
            LogUtils.info(f'从Excel文件读取了 {len(export_configs)} 个有效配置')
            return export_configs
            
        except Exception as e:
            LogUtils.error(f'从Excel文件收集配置时发生错误: {str(e)}')
            return []

    def collect_export_configs(self, inputs):
        """从UI表格收集导出配置（备用方法）"""
        configs = []
        group = inputs.itemById('configGroup')
        if not group:
            LogUtils.error('无法找到配置分组')
            return []
        groupInputs = group.children
        row = 0
        while True:
            format_input = groupInputs.itemById(f'format_{row}')
            if not format_input:
                break
            config = {}
            if format_input.value:
                config['format'] = format_input.value.lower()
            else:
                row += 1
                continue
            name_input = groupInputs.itemById(f'name_{row}')
            if name_input and name_input.value:
                config['custom_name'] = name_input.value
            else:
                if row == 0:
                    LogUtils.error(f'第 {row + 1} 行的自定义名称不能为空')
                    return []
                row += 1
                continue
            config['parameters'] = {}
            for i in range(len(self.batch_exporter.parameters)):
                param_input = groupInputs.itemById(f'param_{row}_{i}')
                if param_input:
                    param_name = self.batch_exporter.parameters[i]['name']
                    config['parameters'][param_name] = param_input.value
            configs.append(config)
            row += 1
        return configs 