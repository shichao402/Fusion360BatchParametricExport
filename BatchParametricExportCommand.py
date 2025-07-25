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

    def execute_batch_export(self, export_configs, export_path):
        try:
            app = adsk.core.Application.get()
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            ui = app.userInterface
            if not design:
                LogUtils.error('无法获取当前设计')
                return
            original_params = self.parameter_manager.backup_parameters(design)
            progress_dialog = ui.createProgressDialog()
            progress_dialog.cancelButtonText = '取消'
            progress_dialog.isBackgroundTranslucent = False
            progress_dialog.isCancelButtonShown = True
            # 拉宽窗口：标题和消息都加长
            progress_dialog.show('批量导出 - Fusion360BatchParametricExport', '准备导出，请稍候...\n', 0, len(export_configs))
            adsk.doEvents()
            exported_count = 0
            def update_progress(doc_name, part_name, idx):
                progress_dialog.progressValue = idx
                progress_dialog.message = f'正在导出文档: {doc_name}\n当前零件: {part_name}'
                adsk.doEvents()
            try:
                for i, config in enumerate(export_configs):
                    if progress_dialog.wasCancelled:
                        break
                    progress_dialog.progressValue = i
                    progress_dialog.message = f'正在导出文档: {config["custom_name"]}\n准备导出...'
                    param_applied = self.parameter_manager.apply_parameters(design, config['parameters'])
                    if param_applied:
                        sub_dir = os.path.join(export_path, config['custom_name'])
                        try:
                            if not os.path.exists(sub_dir):
                                os.makedirs(sub_dir)
                        except Exception as e:
                            LogUtils.error(f'创建目录失败: {sub_dir} {str(e)}')
                            continue
                        export_success = self.export_manager.export_design(
                            design, sub_dir, config['format'], config['custom_name'],
                            lambda part_name: update_progress(config['custom_name'], part_name, i)
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
            result_msg += f'导出路径: {export_path}\n\n'
            if exported_count > 0:
                result_msg += '请检查导出目录中的文件。'
            else:
                result_msg += '没有文件被成功导出，请检查配置和模型。'
            LogUtils.info(result_msg)
        except Exception as e:
            LogUtils.error(f'批量导出时发生错误: {str(e)}')
