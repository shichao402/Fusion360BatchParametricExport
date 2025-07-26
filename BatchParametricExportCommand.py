import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import json
import csv
from typing import List, Dict, Any
import tempfile
from .ConfigUtils import ConfigUtils
from . import ExportUtils
from .LogUtils import LogUtils
from .CacheUtils import CacheUtils


class BatchParametricExportCommand:
    def __init__(self):
        self.parameters = []
        self.export_settings = []
        self.export_manager = ExportUtils.ExportManager()
        self.parameter_manager = ExportUtils.ParameterManager()
        self.config_group = 'BatchParametricExport'
        self.config_key = 'configs'

    def notify(self, args):
        try:
            global app, ui
            app = adsk.core.Application.get()
            ui = app.userInterface
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            if not design:
                LogUtils.error('❌ 请先打开一个 Fusion360 设计文件')
                return
            self.read_starred_parameters(design)
            self.show_dialog()
        except Exception as e:
            if ui:
                LogUtils.error(f'❌ 插件启动失败: {str(e)}\n详细错误: {traceback.format_exc()}')

    def read_starred_parameters(self, design):
        try:
            self.parameters = self.parameter_manager.get_starred_parameters(design)
            ui = adsk.core.Application.get().userInterface
            if not self.parameters:
                LogUtils.warn('⚠️ 未找到标星的参数，请先在参数面板中将需要的参数标星')
        except Exception as e:
            LogUtils.error(f'❌ 读取参数时发生错误: {str(e)}')

    # 移除 get_cache_file_path, save_cached_export_path, load_cached_export_path, save_cached_excel_path, load_cached_excel_path 方法，后续将从 CacheUtils 工具类导入

    def show_dialog(self):
        try:
            ui = adsk.core.Application.get().userInterface
            cmd_def = ui.commandDefinitions.itemById('BatchParametricExport')
            if cmd_def:
                cmd_def.execute()
            else:
                LogUtils.error('插件未正确初始化，请重新加载插件')
        except:
            ui = adsk.core.Application.get().userInterface
            if ui:
                LogUtils.error('创建对话框时发生错误: {}'.format(traceback.format_exc()))

    def execute_batch_export(self, export_configs, export_path, ignore_version=False):
        try:
            app = adsk.core.Application.get()
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            ui = app.userInterface
            if not design:
                LogUtils.error('无法获取当前设计')
                return
            # 获取当前文档名（用于目录）
            doc_name = app.activeDocument.name if app.activeDocument else 'Unnamed'
            
            # 使用传入的ignore_version参数
            
            # 如果勾选了忽略版本号，则去除文档名中的版本号
            if ignore_version and doc_name:
                import re
                # 去除常见的版本号格式，如 " xxx v13"、" xxx_v13"、" xxx-v13"、" [xxx v13]"
                # 支持多种分隔符和方括号格式
                original_name = doc_name
                # 先去除方括号
                doc_name = re.sub(r'^\[(.*)\]$', r'\1', doc_name)
                # 再去除版本号
                doc_name = re.sub(r'([ _\-]?v\d+)$', '', doc_name, flags=re.IGNORECASE).strip()
                # 如果处理后的名称为空，则使用原名称
                if not doc_name:
                    doc_name = original_name
                LogUtils.info(f'文档名处理: "{original_name}" -> "{doc_name}"')
            original_params = self.parameter_manager.backup_parameters(design)
            # 统计所有要导出的零件总数
            total_parts = 0
            for config in export_configs:
                root_component = design.rootComponent
                child_count = len(list(root_component.occurrences))
                if child_count == 0 and root_component.bRepBodies.count > 0:
                    total_parts += 1
                else:
                    total_parts += child_count
            progress_dialog = ui.createProgressDialog()
            progress_dialog.cancelButtonText = '取消'
            progress_dialog.isBackgroundTranslucent = False
            progress_dialog.isCancelButtonShown = True
            progress_dialog.show('批量导出 - Fusion360BatchParametricExport', '准备导出，请稍候...\n', 0, total_parts)
            adsk.doEvents()
            exported_count = 0
            part_progress = 0
            def update_progress(doc_name, part_name):
                nonlocal part_progress
                progress_dialog.progressValue = part_progress
                progress_dialog.message = f'正在导出文档: {doc_name}\n当前零件: {part_name}'
                adsk.doEvents()
                part_progress += 1
            try:
                for config in export_configs:
                    if progress_dialog.wasCancelled:
                        break
                    progress_dialog.message = f'正在导出文档: {config["custom_name"]}\n准备导出...'
                    param_applied = self.parameter_manager.apply_parameters(design, config['parameters'])
                    
                    # 验证参数应用结果
                    if param_applied:
                        # 记录当前参数状态用于调试
                        LogUtils.info(f'配置 {config["custom_name"]} 参数应用成功')
                        for param_name, param_value in config['parameters'].items():
                            try:
                                user_param = design.userParameters.itemByName(param_name)
                                if user_param:
                                    LogUtils.info(f'参数验证: {param_name} = {user_param.expression} (期望: {param_value})')
                            except:
                                pass
                    else:
                        LogUtils.error(f'配置 {config["custom_name"]} 参数应用失败')
                    
                    if param_applied:
                        # 创建目录结构：导出路径/文档名/配置名
                        doc_dir = os.path.join(export_path, doc_name)
                        sub_dir = os.path.join(doc_dir, config['custom_name'])
                        try:
                            if not os.path.exists(doc_dir):
                                os.makedirs(doc_dir)
                            if not os.path.exists(sub_dir):
                                os.makedirs(sub_dir)
                        except Exception as e:
                            LogUtils.error(f'创建目录失败: {sub_dir} {str(e)}')
                            continue
                        export_success = self.export_manager.export_design(
                            design, sub_dir, config['format'], config['custom_name'],
                            lambda part_name: update_progress(config['custom_name'], part_name)
                        )
                        if export_success:
                            exported_count += 1
                    adsk.doEvents()
            finally:
                progress_dialog.hide()
                self.parameter_manager.restore_parameters(design, original_params)
            failed_count = len(export_configs) - exported_count
            result_msg = f'批量导出完成！\n\n'
            result_msg += f'总配置数: {len(export_configs)}\n'
            result_msg += f'成功导出: {exported_count}\n'
            if failed_count > 0:
                result_msg += f'失败数量: {failed_count}\n\n'
                result_msg += '可能的失败原因:\n'
                result_msg += '- 文件名包含非法字符\n'
                result_msg += '- 参数值无效\n'
                result_msg += '- 导出路径权限不足\n'
                result_msg += '- 模型中没有可导出的实体\n\n'
            result_msg += f'导出路径: {export_path}\n'
            result_msg += f'文档目录: {doc_name}\n\n'
            if exported_count > 0:
                result_msg += '请检查导出目录中的文件。'
            else:
                result_msg += '没有文件被成功导出，请检查配置和模型。'
            LogUtils.info(result_msg)
        except Exception as e:
            LogUtils.error(f'批量导出时发生错误: {str(e)}')
