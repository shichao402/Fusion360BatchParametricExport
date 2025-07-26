import adsk.core, adsk.fusion, traceback
import json
import os
import random
import string
from .ConfigUtils import ConfigUtils
import datetime
from .LogUtils import LogUtils
from .CacheUtils import CacheUtils

class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self, batch_exporter, handlers):
        super().__init__()
        self.batch_exporter = batch_exporter
        self.handlers = handlers

    def notify(self, args):
        try:
            changedInput = args.input
            cmd_inputs = args.firingEvent.sender.commandInputs
            ui = adsk.core.Application.get().userInterface
            
            if changedInput.id == 'selectPath':
                if changedInput.value:
                    folderDialog = ui.createFolderDialog()
                    folderDialog.title = '选择导出文件夹'
                    current_path = CacheUtils.load_cached_export_path()
                    if current_path and os.path.exists(current_path):
                        folderDialog.initialDirectory = current_path
                    dialogResult = folderDialog.showDialog()
                    if dialogResult == adsk.core.DialogResults.DialogOK:
                        path_group = cmd_inputs.itemById('pathGroup')
                        if path_group:
                            export_path_input = path_group.children.itemById('exportPath')
                            if export_path_input:
                                export_path_input.value = folderDialog.folder
                                CacheUtils.save_cached_export_path(folderDialog.folder)
                        else:
                            export_path_input = cmd_inputs.itemById('exportPath')
                            if export_path_input:
                                export_path_input.value = folderDialog.folder
                                CacheUtils.save_cached_export_path(folderDialog.folder)
                    changedInput.value = False
                    
            elif changedInput.id == 'selectExcelPath':
                if changedInput.value:
                    fileDialog = ui.createFileDialog()
                    fileDialog.title = '选择Excel配置文件'
                    fileDialog.filter = 'Excel文件 (*.xlsx);;所有文件 (*.*)'
                    fileDialog.isMultiSelect = False
                    
                    # 设置初始目录
                    current_excel_path = cmd_inputs.itemById('excelPath').value
                    if current_excel_path and os.path.exists(os.path.dirname(current_excel_path)):
                        fileDialog.initialDirectory = os.path.dirname(current_excel_path)
                    elif current_excel_path:
                        fileDialog.initialDirectory = os.path.dirname(current_excel_path)
                    else:
                        export_path = CacheUtils.load_cached_export_path()
                        if export_path and os.path.exists(export_path):
                            fileDialog.initialDirectory = export_path
                    
                    dialogResult = fileDialog.showSave()
                    if dialogResult == adsk.core.DialogResults.DialogOK:
                        excel_group = cmd_inputs.itemById('excelGroup')
                        if excel_group:
                            excel_path_input = excel_group.children.itemById('excelPath')
                        else:
                            excel_path_input = cmd_inputs.itemById('excelPath')
                        if excel_path_input:
                            excel_path_input.value = fileDialog.filename
                            CacheUtils.save_cached_excel_path(fileDialog.filename)
                    changedInput.value = False
            elif changedInput.id == 'excelPath':
                # 只要值有变化就缓存
                try:
                    if changedInput.value:
                        CacheUtils.save_cached_excel_path(changedInput.value)
                except Exception:
                    pass
                    
            elif changedInput.id == 'exportTemplate':
                if changedInput.value:
                    # 导出模板前，先缓存当前路径
                    excel_group = cmd_inputs.itemById('excelGroup')
                    if excel_group:
                        excel_path_input = excel_group.children.itemById('excelPath')
                    else:
                        excel_path_input = cmd_inputs.itemById('excelPath')
                    if excel_path_input and excel_path_input.value:
                        CacheUtils.save_cached_excel_path(excel_path_input.value)
                    self.export_excel_template(cmd_inputs)
                    changedInput.value = False
                    
            elif changedInput.id == 'exportPath':
                try:
                    if changedInput.value:
                        CacheUtils.save_cached_export_path(changedInput.value)
                except Exception as e:
                    pass
            elif changedInput.id == 'batchExport':
                if changedInput.value:
                    try:
                        # 获取导出路径
                        export_path = ''
                        path_group = cmd_inputs.itemById('pathGroup')
                        if path_group:
                            export_path_input = path_group.children.itemById('exportPath')
                            if export_path_input:
                                export_path = export_path_input.value
                        else:
                            export_path_input = cmd_inputs.itemById('exportPath')
                            if export_path_input:
                                export_path = export_path_input.value
                        if not export_path or not os.path.exists(export_path):
                            ui.messageBox('❌ 请选择有效的导出路径')
                            changedInput.value = False
                            return
                        CacheUtils.save_cached_export_path(export_path)
                        # 获取导出配置
                        from .CommandExecuteHandler import CommandExecuteHandler
                        handler = CommandExecuteHandler(self.batch_exporter, self.handlers)
                        export_configs = handler.collect_export_configs_from_excel(cmd_inputs)
                        if not export_configs:
                            ui.messageBox('❌ 请先创建Excel配置文件并添加至少一组导出配置')
                            changedInput.value = False
                            return
                        self.batch_exporter.execute_batch_export(export_configs, export_path)
                        ui.messageBox('✅ 导出完成！\n\n💡 提示：\n• 所有配置已成功导出\n• 每个零件已保存到对应子目录\n• 您可以继续编辑Excel文件进行新的导出')
                    except Exception as e:
                        LogUtils.error(f'执行导出时发生错误: {str(e)}')
                        ui.messageBox(f'❌ 执行导出时发生错误:\n{str(e)}')
                    changedInput.value = False
            # 彻底移除 saveConfigs/loadConfigs/备用配置相关事件
        except:
            LogUtils.error('处理输入变化时发生错误:\n{}'.format(traceback.format_exc()))

    def export_excel_template(self, inputs):
        """导出Excel模板"""
        try:
            ui = adsk.core.Application.get().userInterface
            
            # 获取Excel文件路径
            excel_group = inputs.itemById('excelGroup')
            if excel_group:
                excel_path_input = excel_group.children.itemById('excelPath')
            else:
                excel_path_input = inputs.itemById('excelPath')
                
            if not excel_path_input:
                ui.messageBox('无法找到Excel路径输入')
                return
                
            excel_path = excel_path_input.value.strip()
            
            # 如果没有指定路径，弹出文件保存对话框
            if not excel_path:
                fileDialog = ui.createFileDialog()
                fileDialog.title = '保存Excel模板文件'
                fileDialog.filter = 'Excel文件 (*.xlsx)'
                fileDialog.isMultiSelect = False
                fileDialog.initialFilename = '导出配置模板.xlsx'
                
                # 设置初始目录
                export_path = CacheUtils.load_cached_export_path()
                if export_path and os.path.exists(export_path):
                    fileDialog.initialDirectory = export_path
                
                dialogResult = fileDialog.showSave()
                if dialogResult == adsk.core.DialogResults.DialogOK:
                    excel_path = fileDialog.filename
                    excel_path_input.value = excel_path
                else:
                    return
            
            # 检查参数
            if not self.batch_exporter.parameters:
                ui.messageBox('❌ 未找到任何标星参数\n请确保在参数面板中将参数标星（收藏）')
                return
            
            # 创建Excel模板
            success = ConfigUtils.create_excel_template(excel_path, self.batch_exporter.parameters)
            
            if success:
                CacheUtils.save_cached_excel_path(excel_path)
                ui.messageBox(f'✅ Excel模板已成功创建！\n\n文件路径：{excel_path}\n\n请在Excel中编辑配置后保存文件，然后点击"导出"按钮执行批量导出。')
            else:
                ui.messageBox('❌ 创建Excel模板失败，请检查文件路径和权限。')
                
        except Exception as e:
            LogUtils.error(f'导出Excel模板时发生错误: {str(e)}')
            ui.messageBox(f'❌ 导出Excel模板时发生错误:\n{str(e)}')

    # 彻底移除 save_current_configs_to_design 和 load_configs_to_ui 方法 