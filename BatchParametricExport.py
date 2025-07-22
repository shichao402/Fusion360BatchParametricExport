#Author-YourName
#Description-Fusion360 参数化批量导出插件

import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import json
import csv
from typing import List, Dict, Any
import tempfile

# 导入工具模块
try:
    from . import ExportUtils
except ImportError:
    import ExportUtils

# 全局变量
app = None
ui = None
handlers = []
loaded_row_count = 0  # 用于跟踪从项目加载的配置行数

class BatchParametricExportCommand:
    def __init__(self):
        self.parameters = []
        self.export_settings = []
        self.export_manager = ExportUtils.ExportManager()
        self.parameter_manager = ExportUtils.ParameterManager()
        
    def notify(self, args):
        try:
            # 获取应用程序和用户界面对象
            global app, ui
            app = adsk.core.Application.get()
            ui = app.userInterface
            
            # 获取当前设计
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            
            if not design:
                ui.messageBox('❌ 请先打开一个 Fusion360 设计文件')
                return
                
            # 读取标星参数
            self.read_starred_parameters(design)
            
            # 显示对话框
            self.show_dialog()
            
        except Exception as e:
            if ui:
                ui.messageBox(f'❌ 插件启动失败:\n{str(e)}\n\n详细错误:\n{traceback.format_exc()}')
    
    def read_starred_parameters(self, design):
        """读取设计中的标星参数"""
        try:
            self.parameters = self.parameter_manager.get_starred_parameters(design)
                    
            if not self.parameters:
                ui.messageBox('⚠️ 未找到标星的参数，请先在参数面板中将需要的参数标星')
        except Exception as e:
            ui.messageBox(f'❌ 读取参数时发生错误:\n{str(e)}')
    
    def get_cache_file_path(self):
        """获取缓存文件路径"""
        try:
            # 使用用户临时目录
            temp_dir = tempfile.gettempdir()
            cache_file = os.path.join(temp_dir, 'Fusion360BatchParametricExport_cache.json')
            return cache_file
        except:
            return None
    
    def load_cached_export_path(self):
        """加载缓存的导出路径"""
        try:
            cache_file = self.get_cache_file_path()
            if cache_file and os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
                    return cache_data.get('export_path', '')
        except:
            pass
        return ''
    
    def save_cached_export_path(self, path):
        """保存导出路径到缓存"""
        try:
            if not path:
                return
                
            cache_file = self.get_cache_file_path()
            if cache_file:
                cache_data = {'export_path': path}
                with open(cache_file, 'w', encoding='utf-8') as f:
                    json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def save_configs_to_design(self, design, configs):
        """将导出配置保存到设计文件的属性中"""
        try:
            # 将配置转换为JSON字符串
            config_json = json.dumps(configs, ensure_ascii=False, indent=2)
            
            # 获取或创建属性组
            attribGroup = design.findAttributes('BatchParametricExport', 'configs')
            if attribGroup and len(attribGroup) > 0:
                # 更新现有属性
                attribGroup[0].value = config_json
            else:
                # 创建新属性
                design.attributes.add('BatchParametricExport', 'configs', config_json)
            
            return True
        except Exception as e:
            global ui
            ui.messageBox(f'保存配置到项目失败: {str(e)}')
            return False
    
    def load_configs_from_design(self, design):
        """从设计文件的属性中加载导出配置"""
        try:
            # 查找属性
            attribGroup = design.findAttributes('BatchParametricExport', 'configs')
            if attribGroup and len(attribGroup) > 0:
                config_json = attribGroup[0].value
                configs = json.loads(config_json)
                return configs
            else:
                return []
        except Exception as e:
            # 加载失败时返回空列表，不显示错误
            return []
    
    def show_dialog(self):
        """显示主对话框"""
        try:
            # 获取已创建的命令定义
            cmd_def = ui.commandDefinitions.itemById('BatchParametricExport')
            if cmd_def:
                # 执行命令
                cmd_def.execute()
            else:
                ui.messageBox('插件未正确初始化，请重新加载插件')
            
        except:
            if ui:
                ui.messageBox('创建对话框时发生错误:\n{}'.format(traceback.format_exc()))
    
    def execute_batch_export(self, export_configs, export_path):
        """执行批量导出"""
        try:
            # 获取当前设计
            app = adsk.core.Application.get()
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            
            if not design:
                ui.messageBox('无法获取当前设计')
                return
            
            # 备份当前参数
            original_params = self.parameter_manager.backup_parameters(design)
            
            # 进度对话框
            progress_dialog = ui.createProgressDialog()
            progress_dialog.cancelButtonText = '取消'
            progress_dialog.isBackgroundTranslucent = False
            progress_dialog.isCancelButtonShown = True
            progress_dialog.show('批量导出', '准备导出...', 0, len(export_configs))
            
            exported_count = 0
            try:
                for i, config in enumerate(export_configs):
                    if progress_dialog.wasCancelled:
                        break
                        
                    progress_dialog.progressValue = i
                    progress_dialog.message = f'正在导出: {config["custom_name"]}'
                    
                    # 应用参数
                    param_applied = self.parameter_manager.apply_parameters(design, config['parameters'])
                    if param_applied:
                        # 创建子目录
                        sub_dir = os.path.join(export_path, config['custom_name'])
                        try:
                            if not os.path.exists(sub_dir):
                                os.makedirs(sub_dir)
                        except Exception as e:
                            ui.messageBox(f'创建目录失败: {sub_dir}\n{str(e)}')
                            continue
                        
                        # 导出文件
                        export_success = self.export_manager.export_design(design, sub_dir, config['format'], config['custom_name'])
                        if export_success:
                            exported_count += 1
                        # 失败的情况在最终消息中汇总
                    # 参数应用失败也在最终消息中汇总
                    
                    # 更新进度
                    adsk.doEvents()
                    
            finally:
                progress_dialog.hide()
                # 恢复原始参数
                self.parameter_manager.restore_parameters(design, original_params)
                
            # 构建详细的结果消息
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
            
            ui.messageBox(result_msg)
            
        except Exception as e:
            ui.messageBox(f'批量导出时发生错误:\n{str(e)}')

class CommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        try:
            cmd = args.command
            cmd.isRepeatable = False
            
            # 获取当前设计并读取参数
            app = adsk.core.Application.get()
            ui = app.userInterface
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            
            if not design:
                ui.messageBox('❌ 请先打开一个 Fusion360 设计文件')
                return
            
            # 重新读取标星参数
            batch_exporter.read_starred_parameters(design)
            
            # 尝试加载已保存的配置
            saved_configs = batch_exporter.load_configs_from_design(design)
            
            # 获取命令的输入
            inputs = cmd.commandInputs
            
            # 导出路径选择（放在最前面，常用功能）
            try:
                path_group = inputs.addGroupCommandInput('pathGroup', '📁 导出路径设置')
                path_group.isExpanded = True
                pathInputs = path_group.children
                
                # 加载缓存的路径
                cached_path = batch_exporter.load_cached_export_path()
                pathInputs.addStringValueInput('exportPath', '导出路径', cached_path)
                pathInputs.addBoolValueInput('selectPath', '🔍 选择路径...', False)
            except Exception as e:
                # 如果路径组创建失败，使用备用方案
                inputs.addStringValueInput('exportPath', '导出路径', batch_exporter.load_cached_export_path())
                inputs.addBoolValueInput('selectPath', '🔍 选择路径...', False)
            
            # 添加说明文本
            inputs.addTextBoxCommandInput('description', '', 
                '🎯 批量参数化导出插件 - 每个零件单独导出\n\n' +
                '✅ 支持导出格式: STEP, IGES, STL, OBJ, 3MF\n' +
                '✅ 每个零件组件单独导出到对应子目录\n' +
                '✅ 支持参数化批量导出\n\n' +
                '📋 配置说明:\n' +
                '• 导出格式: step, iges, stl, obj, 3mf\n' +
                '• 自定义名称: 必填，用于子目录和文件名\n' +
                '• 参数列: 您的标星参数值', 4, True)
            
            # 显示参数信息
            param_count = len(batch_exporter.parameters)
            param_info = f"当前标星参数 (共{param_count}个):\n"
            if param_count > 0:
                for param in batch_exporter.parameters:
                    param_info += f"- {param['name']}: {param['expression']}\n"
            else:
                param_info += "❌ 未找到任何标星参数\n请确保在参数面板中将参数标星（收藏）"
            
            inputs.addTextBoxCommandInput('paramInfo', '参数信息', param_info, 4, True)
            
            # 创建分组来组织输入
            group = inputs.addGroupCommandInput('configGroup', '📋 导出配置')
            group.isExpanded = True
            groupInputs = group.children
            
            # 根据已保存的配置创建行，如果没有则创建默认行
            if saved_configs and len(saved_configs) > 0:
                for row_idx, config in enumerate(saved_configs):
                    # 添加行标签
                    groupInputs.addTextBoxCommandInput(f'rowLabel{row_idx}', '', f'配置 {row_idx + 1}:', 1, True)
                    groupInputs.addStringValueInput(f'format_{row_idx}', '导出格式', config.get('format', 'step'))
                    groupInputs.addStringValueInput(f'name_{row_idx}', '自定义名称', config.get('name', ''))
                    
                    # 为每个参数添加输入
                    for i, param in enumerate(batch_exporter.parameters):
                        param_value = config.get('parameters', {}).get(param['name'], param['expression'])
                        groupInputs.addStringValueInput(f'param_{row_idx}_{i}', param['name'], param_value)
                
                # 更新全局row_count以匹配加载的配置
                global loaded_row_count
                loaded_row_count = len(saved_configs) - 1
            else:
                # 没有保存的配置时，添加第一行默认配置
                groupInputs.addTextBoxCommandInput('rowLabel0', '', '配置 1:', 1, True)
                groupInputs.addStringValueInput('format_0', '导出格式', 'step')
                groupInputs.addStringValueInput('name_0', '自定义名称', '')
                
                # 为每个参数添加输入
                for i, param in enumerate(batch_exporter.parameters):
                    groupInputs.addStringValueInput(f'param_0_{i}', param['name'], param['expression'])
                
                loaded_row_count = 0
            
            # 添加行按钮
            inputs.addBoolValueInput('addRow', '➕ 添加新行', False)
            inputs.addBoolValueInput('removeRow', '➖ 删除最后行', False)
            
            # 配置管理按钮
            config_group = inputs.addGroupCommandInput('configManagementGroup', '💾 配置管理')
            config_group.isExpanded = True
            configMgmtInputs = config_group.children
            configMgmtInputs.addBoolValueInput('saveConfigs', '💾 保存配置到项目', False)
            configMgmtInputs.addBoolValueInput('loadConfigs', '📂 加载项目配置', False)
            configMgmtInputs.addTextBoxCommandInput('configTip', '', 
                '💡 提示：配置保存到项目后会跟随设计文件一起移动', 2, True)
            
            # 添加事件处理器
            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
            
            onInputChanged = CommandInputChangedHandler()
            # 设置正确的行计数
            onInputChanged.row_count = loaded_row_count
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)
            
        except:
            if ui:
                ui.messageBox('创建命令时发生错误:\n{}'.format(traceback.format_exc()))

class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        try:
            # 执行批量导出逻辑
            inputs = args.command.commandInputs
            
            # 获取导出路径（从pathGroup中）
            export_path = ''
            path_group = inputs.itemById('pathGroup')
            if path_group:
                export_path_input = path_group.children.itemById('exportPath')
                if export_path_input:
                    export_path = export_path_input.value
            else:
                # 备用方案：直接查找exportPath
                export_path_input = inputs.itemById('exportPath')
                if export_path_input:
                    export_path = export_path_input.value
            
            if not export_path or not os.path.exists(export_path):
                ui.messageBox('请选择有效的导出路径')
                return
            
            # 保存路径到缓存
            batch_exporter.save_cached_export_path(export_path)
                
            # 收集参数配置
            export_configs = self.collect_export_configs(inputs)
            
            if not export_configs:
                ui.messageBox('请至少配置一组导出参数')
                return
                
            # 开始批量导出
            batch_exporter.execute_batch_export(export_configs, export_path)
            
            # 导出完成后显示提示信息
            ui.messageBox('✅ 导出完成！\n\n💡 提示：对话框将保持打开，您可以：\n• 修改参数值进行新的导出\n• 添加新行配置其他导出任务\n• 手动点击"取消"关闭对话框', '导出成功')
            
            # 不关闭对话框，让用户可以继续操作
            
        except:
            if ui:
                ui.messageBox('执行导出时发生错误:\n{}'.format(traceback.format_exc()))
    
    def collect_export_configs(self, inputs):
        """从输入中收集导出配置"""
        configs = []
        
        # 获取配置分组
        group = inputs.itemById('configGroup')
        if not group:
            ui.messageBox('无法找到配置分组')
            return []
            
        groupInputs = group.children
        
        # 收集配置数据 - 检查所有可能的行
        row = 0
        while True:
            # 尝试获取这一行的格式输入
            format_input = groupInputs.itemById(f'format_{row}')
            if not format_input:
                break
                
            config = {}
            
            # 获取导出格式
            if format_input.value:
                config['format'] = format_input.value.lower()
            else:
                row += 1
                continue
                
            # 获取自定义名称
            name_input = groupInputs.itemById(f'name_{row}')
            if name_input and name_input.value:
                config['custom_name'] = name_input.value
            else:
                if row == 0:  # 如果是第一行，必须有名称
                    ui.messageBox(f'第 {row + 1} 行的自定义名称不能为空')
                    return []
                row += 1
                continue
            
            # 获取参数值
            config['parameters'] = {}
            for i in range(len(batch_exporter.parameters)):
                param_input = groupInputs.itemById(f'param_{row}_{i}')
                if param_input:
                    param_name = batch_exporter.parameters[i]['name']
                    config['parameters'][param_name] = param_input.value
                    
            configs.append(config)
            row += 1
            
        return configs

class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
        self.row_count = 0  # 跟踪当前最大行索引（第一行是0）
        
    def notify(self, args):
        try:
            changedInput = args.input
            inputs = args.inputs
            
            if changedInput.id == 'selectPath':
                if changedInput.value:
                    # 打开文件夹选择对话框
                    folderDialog = ui.createFolderDialog()
                    folderDialog.title = '选择导出文件夹'
                    
                    # 设置默认路径为当前缓存的路径
                    current_path = batch_exporter.load_cached_export_path()
                    if current_path and os.path.exists(current_path):
                        folderDialog.initialDirectory = current_path
                    
                    dialogResult = folderDialog.showDialog()
                    if dialogResult == adsk.core.DialogResults.DialogOK:
                        # 获取路径输入控件（在pathGroup中）
                        path_group = inputs.itemById('pathGroup')
                        if path_group:
                            export_path_input = path_group.children.itemById('exportPath')
                            if export_path_input:
                                export_path_input.value = folderDialog.folder
                                # 立即保存到缓存
                                batch_exporter.save_cached_export_path(folderDialog.folder)
                        else:
                            # 备用方案：直接查找exportPath
                            export_path_input = inputs.itemById('exportPath')
                            if export_path_input:
                                export_path_input.value = folderDialog.folder
                                batch_exporter.save_cached_export_path(folderDialog.folder)
                    
                    changedInput.value = False
                    
            elif changedInput.id == 'addRow':
                if changedInput.value:
                    self.add_table_row(inputs)
                    changedInput.value = False
                    
            elif changedInput.id == 'removeRow':
                if changedInput.value:
                    self.remove_table_row(inputs)
                    changedInput.value = False
                    
            elif changedInput.id == 'exportPath':
                # 当用户手动输入路径时，也保存到缓存
                try:
                    if changedInput.value:
                        batch_exporter.save_cached_export_path(changedInput.value)
                except Exception as e:
                    # 忽略缓存保存错误，不影响主要功能
                    pass
                    
            elif changedInput.id == 'saveConfigs':
                if changedInput.value:
                    self.save_current_configs_to_design(inputs)
                    changedInput.value = False
                    
            elif changedInput.id == 'loadConfigs':
                if changedInput.value:
                    self.load_configs_to_ui(inputs)
                    changedInput.value = False
                    
        except:
            if ui:
                ui.messageBox('处理输入变化时发生错误:\n{}'.format(traceback.format_exc()))
    
    def add_table_row(self, inputs):
        """添加新的配置行"""
        try:
            group = inputs.itemById('configGroup')
            if not group:
                ui.messageBox('无法找到配置分组')
                return
                
            # 增加行计数，新行索引
            self.row_count += 1
            new_row = self.row_count
            groupInputs = group.children
            
            # 添加新行标签和输入
            groupInputs.addTextBoxCommandInput(f'rowLabel{new_row}', '', f'配置 {new_row + 1}:', 1, True)
            groupInputs.addStringValueInput(f'format_{new_row}', '导出格式', 'step')
            groupInputs.addStringValueInput(f'name_{new_row}', '自定义名称', '')
            
            # 为每个参数添加输入
            for i, param in enumerate(batch_exporter.parameters):
                groupInputs.addStringValueInput(f'param_{new_row}_{i}', param['name'], param['expression'])
                
        except Exception as e:
            ui.messageBox(f'添加行时发生错误:\n{str(e)}')
    
    def remove_table_row(self, inputs):
        """删除最后一行配置"""
        try:
            group = inputs.itemById('configGroup')
            if not group or self.row_count <= 0:
                return
                
            # 删除最后一行的所有控件
            row_to_remove = self.row_count
            groupInputs = group.children
            
            # 删除行标签
            label_input = groupInputs.itemById(f'rowLabel{row_to_remove}')
            if label_input:
                groupInputs.remove(label_input)
            
            # 删除格式输入
            format_input = groupInputs.itemById(f'format_{row_to_remove}')
            if format_input:
                groupInputs.remove(format_input)
                
            # 删除名称输入
            name_input = groupInputs.itemById(f'name_{row_to_remove}')
            if name_input:
                groupInputs.remove(name_input)
                
            # 删除参数输入
            for i in range(len(batch_exporter.parameters)):
                param_input = groupInputs.itemById(f'param_{row_to_remove}_{i}')
                if param_input:
                    groupInputs.remove(param_input)
            
            self.row_count -= 1
            
        except Exception as e:
            ui.messageBox(f'删除行时发生错误:\n{str(e)}')
    
    def save_current_configs_to_design(self, inputs):
        """保存当前配置到设计文件"""
        try:
            # 获取当前设计
            app = adsk.core.Application.get()
            ui = app.userInterface
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            
            if not design:
                ui.messageBox('❌ 无法获取当前设计文件')
                return
            
            # 收集当前UI中的所有配置
            configs = []
            
            # 获取配置分组
            group = inputs.itemById('configGroup')
            if not group:
                ui.messageBox('❌ 无法找到配置组')
                return
                
            groupInputs = group.children
            
            # 遍历所有行，收集配置
            for row_idx in range(self.row_count + 1):  # 包括第0行
                format_input = groupInputs.itemById(f'format_{row_idx}')
                name_input = groupInputs.itemById(f'name_{row_idx}')
                
                if format_input and name_input:
                    config = {
                        'format': format_input.value,
                        'name': name_input.value,
                        'parameters': {}
                    }
                    
                    # 收集参数值
                    for i, param in enumerate(batch_exporter.parameters):
                        param_input = groupInputs.itemById(f'param_{row_idx}_{i}')
                        if param_input:
                            config['parameters'][param['name']] = param_input.value
                    
                    configs.append(config)
            
            # 保存到设计文件
            if batch_exporter.save_configs_to_design(design, configs):
                ui.messageBox(f'✅ 成功保存 {len(configs)} 个配置到项目文件！\n\n💡 这些配置现在会跟随项目文件一起移动到其他电脑。')
            
        except Exception as e:
            ui.messageBox(f'保存配置时发生错误:\n{str(e)}')
    
    def load_configs_to_ui(self, inputs):
        """从设计文件加载配置到UI"""
        try:
            # 获取当前设计
            app = adsk.core.Application.get()
            ui = app.userInterface
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            
            if not design:
                ui.messageBox('❌ 无法获取当前设计文件')
                return
            
            # 从设计文件加载配置
            saved_configs = batch_exporter.load_configs_from_design(design)
            
            if not saved_configs or len(saved_configs) == 0:
                ui.messageBox('ℹ️ 项目中没有找到已保存的配置')
                return
            
            # 获取配置组
            group = inputs.itemById('configGroup')
            if not group:
                ui.messageBox('❌ 无法找到配置组')
                return
            
            # 删除现有的所有配置行
            for row_idx in range(self.row_count + 1):
                self._remove_config_row(group, row_idx)
            
            # 重置行计数
            self.row_count = len(saved_configs) - 1
            
            # 添加加载的配置
            groupInputs = group.children
            for row_idx, config in enumerate(saved_configs):
                # 添加行标签
                groupInputs.addTextBoxCommandInput(f'rowLabel{row_idx}', '', f'配置 {row_idx + 1}:', 1, True)
                groupInputs.addStringValueInput(f'format_{row_idx}', '导出格式', config.get('format', 'step'))
                groupInputs.addStringValueInput(f'name_{row_idx}', '自定义名称', config.get('name', ''))
                
                # 为每个参数添加输入
                for i, param in enumerate(batch_exporter.parameters):
                    param_value = config.get('parameters', {}).get(param['name'], param['expression'])
                    groupInputs.addStringValueInput(f'param_{row_idx}_{i}', param['name'], param_value)
            
            ui.messageBox(f'✅ 成功加载 {len(saved_configs)} 个配置！')
            
        except Exception as e:
            ui.messageBox(f'加载配置时发生错误:\n{str(e)}')
    
    def _remove_config_row(self, group, row_idx):
        """移除指定行的配置（辅助方法）"""
        try:
            groupInputs = group.children
            
            # 删除行标签
            label_input = groupInputs.itemById(f'rowLabel{row_idx}')
            if label_input:
                groupInputs.remove(label_input)
            
            # 删除格式输入
            format_input = groupInputs.itemById(f'format_{row_idx}')
            if format_input:
                groupInputs.remove(format_input)
                
            # 删除名称输入
            name_input = groupInputs.itemById(f'name_{row_idx}')
            if name_input:
                groupInputs.remove(name_input)
                
            # 删除参数输入
            for i in range(len(batch_exporter.parameters)):
                param_input = groupInputs.itemById(f'param_{row_idx}_{i}')
                if param_input:
                    groupInputs.remove(param_input)
                    
        except Exception:
            pass  # 忽略删除错误

 
  

# 全局实例
batch_exporter = BatchParametricExportCommand()

def run(context):
    try:
        global app, ui
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # 创建命令定义
        cmd_def = ui.commandDefinitions.itemById('BatchParametricExport')
        if not cmd_def:
            cmd_def = ui.commandDefinitions.addButtonDefinition(
                'BatchParametricExport', 
                '批量参数化导出', 
                '批量导出不同参数配置的模型文件，每个零件单独导出'
            )
        
        # 创建事件处理器
        command_created_handler = CommandCreatedEventHandler()
        cmd_def.commandCreated.add(command_created_handler)
        handlers.append(command_created_handler)
        
        # 获取DESIGN工作空间
        design_workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if not design_workspace:
            ui.messageBox('⚠️ 无法找到DESIGN工作空间')
            return
        
        # 获取ADD-INS面板（这是实用程序/外接程序面板的正确ID）
        addins_panel = design_workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
        if not addins_panel:
            ui.messageBox('⚠️ 无法找到ADD-INS面板')
            return
        
        # 检查按钮是否已存在
        existing_control = addins_panel.controls.itemById('BatchParametricExport')
        if not existing_control:
            # 添加按钮到ADD-INS面板
            button_control = addins_panel.controls.addCommand(cmd_def)
            # 确保按钮在面板中可见
            button_control.isPromotedByDefault = True
            button_control.isPromoted = True
            
            ui.messageBox('✅ 批量参数化导出插件已加载成功！\n\n📍 按钮已添加到: ADD-INS面板\n\n🎯 您可以在DESIGN工作空间的ADD-INS面板中找到"批量参数化导出"按钮。', '插件加载成功')
        else:
            ui.messageBox('✅ 批量参数化导出插件已存在！\n\n📍 按钮位置: ADD-INS面板', '插件已加载')
        
    except Exception as e:
        if ui:
            ui.messageBox(f'插件启动失败:\n{str(e)}')

def stop(context):
    try:
        global app, ui
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # 获取DESIGN工作空间并移除按钮
        try:
            design_workspace = ui.workspaces.itemById('FusionSolidEnvironment')
            if design_workspace:
                addins_panel = design_workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
                if addins_panel:
                    existing_control = addins_panel.controls.itemById('BatchParametricExport')
                    if existing_control:
                        existing_control.deleteMe()
        except:
            pass
        
        # 删除命令定义
        try:
            cmd_def = ui.commandDefinitions.itemById('BatchParametricExport')
            if cmd_def:
                cmd_def.deleteMe()
        except:
            pass
        
        # 清理事件处理器
        try:
            for handler in handlers:
                if hasattr(handler, '__del__'):
                    handler.__del__()
            handlers.clear()
        except:
            pass
            
    except Exception as e:
        if ui:
            ui.messageBox(f'插件停止时发生错误:\n{str(e)}') 