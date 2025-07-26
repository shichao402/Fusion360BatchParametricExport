import adsk.core, adsk.fusion, traceback
import json
from .LogUtils import LogUtils
from .CacheUtils import CacheUtils

class CommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self, batch_exporter, handlers):
        super().__init__()
        self.batch_exporter = batch_exporter
        self.handlers = handlers

    def notify(self, args):
        try:
            cmd = args.command
            cmd.isRepeatable = False
            cmd.isAutoTerminate = False
            cmd.isOKButtonVisible = False
            app = adsk.core.Application.get()
            ui = app.userInterface
            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            if not design:
                LogUtils.error('❌ 请先打开一个 Fusion360 设计文件')
                return
            self.batch_exporter.read_starred_parameters(design)
            # 获取命令的输入
            inputs = cmd.commandInputs
            
            # 导出路径选择（放在最前面，常用功能）
            try:
                path_group = inputs.addGroupCommandInput('pathGroup', '📁 导出路径设置')
                path_group.isExpanded = True
                pathInputs = path_group.children
                cached_path = CacheUtils.load_cached_export_path()
                pathInputs.addStringValueInput('exportPath', '导出路径', cached_path)
                pathInputs.addBoolValueInput('selectPath', '🔍 选择路径...', False)
                pathInputs.addBoolValueInput('ignoreVersionInDocName', '忽略文档版本号（仅用主名）', True)
            except Exception as e:
                inputs.addStringValueInput('exportPath', '导出路径', CacheUtils.load_cached_export_path())
                inputs.addBoolValueInput('selectPath', '🔍 选择路径...', False)
            
            # 移除关于插件说明和支持格式的说明文本（即description相关的addTextBoxCommandInput）
            
            # Excel配置管理组
            excel_group = inputs.addGroupCommandInput('excelGroup', '📊 Excel配置管理')
            excel_group.isExpanded = True
            excelInputs = excel_group.children
            # Excel文件路径输入
            cached_excel_path = CacheUtils.load_cached_excel_path()
            excelInputs.addStringValueInput('excelPath', 'Excel模板路径', cached_excel_path)
            excelInputs.addBoolValueInput('selectExcelPath', '🔍 选择Excel文件...', False)
            # Excel操作按钮
            excelInputs.addBoolValueInput('exportTemplate', '📤 导出模板', False)
            # 添加自定义批量导出按钮
            excelInputs.addBoolValueInput('batchExport', '🚀 批量导出', False)
            # 移除excelTip相关的addTextBoxCommandInput，不再添加Excel操作提示文本
            # 不再添加备用配置管理按钮和分组

            # 将参数信息移动到面板最末尾，并设置最大高度为300像素，超出时显示滚动条
            param_count = len(self.batch_exporter.parameters)
            param_info = f"当前标星参数 (共{param_count}个):\n"
            if param_count > 0:
                for param in self.batch_exporter.parameters:
                    comment = param.get('comment', '').strip()
                    if comment:
                        # 限制注释长度，避免显示过长
                        if len(comment) > 50:
                            comment = comment[:47] + '...'
                        param_info += f"- {param['name']} ({comment}): {param['expression']}\n"
                    else:
                        param_info += f"- {param['name']}: {param['expression']}\n"
            else:
                param_info += "❌ 未找到任何标星参数\n请确保在参数面板中将参数标星（收藏）"
            # 添加参数信息到最末尾，最大高度300，超出显示滚动条
            inputs.addTextBoxCommandInput('paramInfo', '参数信息', param_info, 10, True).tooltip = '参数信息（最多显示300像素高度，超出显示滚动条）'
            # 设置最大高度（Fusion360 API不直接支持像素高度，但可以通过行数近似控制，10行约300像素）
            
            # 注册 inputChanged 事件
            from .CommandInputChangedHandler import CommandInputChangedHandler
            onInputChanged = CommandInputChangedHandler(self.batch_exporter, self.handlers)
            cmd.inputChanged.add(onInputChanged)
            self.handlers.append(onInputChanged)
        except:
            app = adsk.core.Application.get()
            ui = app.userInterface
            if ui:
                LogUtils.error('创建命令时发生错误: {}'.format(traceback.format_exc())) 