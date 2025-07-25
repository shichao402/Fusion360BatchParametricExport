#Author-YourName
#Description-Fusion360 参数化批量导出插件

import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import json
import csv
from typing import List, Dict, Any
import tempfile
import random
import string
from .ConfigUtils import ConfigUtils
from . import ExportUtils
from .BatchParametricExportCommand import BatchParametricExportCommand
from .CommandCreatedEventHandler import CommandCreatedEventHandler
from .CommandExecuteHandler import CommandExecuteHandler
from .CommandInputChangedHandler import CommandInputChangedHandler
from .LogUtils import LogUtils

handlers = []
batch_exporter = BatchParametricExportCommand()

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        cmdDefs = ui.commandDefinitions
        cmdDef = cmdDefs.itemById('BatchParametricExport')
        if not cmdDef:
            cmdDef = cmdDefs.addButtonDefinition('BatchParametricExport', '批量参数化导出', '批量导出不同参数配置的模型文件，每个零件单独导出')

        # 注册commandCreated事件
        onCreated = CommandCreatedEventHandler(batch_exporter, handlers)
        cmdDef.commandCreated.add(onCreated)
        handlers.append(onCreated)

        # 添加按钮到ADD-INS面板
        addinsPanel = ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        if addinsPanel:
            if not addinsPanel.controls.itemById('BatchParametricExport'):
                addinsPanel.controls.addCommand(cmdDef)
    except:
        if ui:
            LogUtils.error('Failed: {}'.format(traceback.format_exc()))

def stop(context):
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        # 删除按钮
        addinsPanel = ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        if addinsPanel:
            cntrl = addinsPanel.controls.itemById('BatchParametricExport')
            if cntrl:
                cntrl.deleteMe()
        # 删除命令定义
        cmdDef = ui.commandDefinitions.itemById('BatchParametricExport')
        if cmdDef:
            cmdDef.deleteMe()
        # 清理事件处理器
        handlers.clear()
    except:
        if ui:
            LogUtils.error('Failed to stop add-in: {}'.format(traceback.format_exc())) 