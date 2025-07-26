"""
导出工具模块
提供各种格式的导出功能
"""

import adsk.core
import adsk.fusion
import os
from .LogUtils import LogUtils

class ExportManager:
    """导出管理器"""
    
    def __init__(self):
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
    
    def export_design(self, design, export_path, export_format, custom_name, progress_callback=None):
        """使用可见性控制导出设计中的所有子组件（每个零件单独导出）"""
        try:
            if not design:
                LogUtils.error('设计对象无效')
                return False
                
            if not export_path or not os.path.exists(export_path):
                LogUtils.error(f'导出路径无效: {export_path}')
                return False
                
            if not custom_name or not custom_name.strip():
                LogUtils.error('自定义名称不能为空')
                return False
            
            export_mgr = design.exportManager
            root_component = design.rootComponent
            
            if not export_mgr or not root_component:
                LogUtils.error('无法获取导出管理器或根组件')
                return False
            
            # 获取所有子组件（这是正确的方法）
            child_components = []
            for occurrence in root_component.occurrences:
                child_components.append({
                    'occurrence': occurrence,
                    'component': occurrence.component,
                    'name': occurrence.component.name
                })
            
            if not child_components:
                # 如果没有子组件，检查根组件是否有实体
                if root_component.bRepBodies.count > 0:
                    if progress_callback:
                        progress_callback(root_component.name)
                        adsk.doEvents()
                    # 直接导出根组件
                    return self._export_single_format(export_mgr, export_path, export_format, custom_name, root_component.name, None)
                else:
                    LogUtils.warn('设计中没有找到可导出的零件')
                    return False
            
            # 记录原始可见性（使用lightBulb状态）
            original_visibility = {}
            for occurrence in root_component.allOccurrences:
                original_visibility[occurrence.entityToken] = occurrence.isLightBulbOn
            
            export_success_count = 0
            
            # 为每个子组件单独导出
            for comp_info in child_components:
                try:
                    occurrence = comp_info['occurrence']
                    component = comp_info['component']
                    comp_name = comp_info['name']
                    
                    # 隐藏所有组件（使用lightBulb属性）
                    for occ in root_component.allOccurrences:
                        occ.isLightBulbOn = False
                    
                    # 只显示目标组件
                    occurrence.isLightBulbOn = True
                    
                    # 导出当前可见的组件
                    if progress_callback:
                        progress_callback(comp_name)
                        adsk.doEvents()
                    result = self._export_single_format(export_mgr, export_path, export_format, custom_name, comp_name, occurrence)
                    
                    if result:
                        export_success_count += 1
                        
                except Exception as comp_e:
                    # 继续处理下一个组件
                    continue
            
            # 恢复原始可见性
            try:
                for occurrence in root_component.allOccurrences:
                    if occurrence.entityToken in original_visibility:
                        occurrence.isLightBulbOn = original_visibility[occurrence.entityToken]
            except Exception as restore_e:
                pass
            
            # 返回是否至少成功导出了一个组件
            return export_success_count > 0
                
        except Exception as e:
            LogUtils.error(f'导出时发生错误: {str(e)}')
            return False
    
    def _export_single_format(self, export_mgr, export_path, export_format, custom_name, comp_name, occurrence):
        """导出单个格式的文件"""
        try:
            # 根据格式选择导出方法
            if export_format.lower() == 'step':
                return self._export_step_visibility(export_mgr, export_path, custom_name, comp_name)
            elif export_format.lower() == 'iges':
                return self._export_iges_visibility(export_mgr, export_path, custom_name, comp_name)
            elif export_format.lower() == 'stl':
                return self._export_stl_visibility(export_mgr, export_path, custom_name, comp_name, occurrence)
            elif export_format.lower() == 'obj':
                return self._export_obj_visibility(export_mgr, export_path, custom_name, comp_name, occurrence)
            elif export_format.lower() == '3mf':
                return self._export_3mf_visibility(export_mgr, export_path, custom_name, comp_name, occurrence)
            else:
                return False
                
        except Exception as e:
            return False
    
    def _sanitize_filename(self, filename):
        """清理文件名，移除非法字符"""
        if not filename:
            return 'Unnamed'
        
        # 移除或替换非法字符
        import re
        # 移除不允许的字符
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # 移除多余的空格和点
        sanitized = re.sub(r'\s+', '_', sanitized.strip())
        sanitized = sanitized.strip('.')
        
        # 确保不为空
        if not sanitized:
            sanitized = 'Unnamed'
        
        return sanitized
    
    def _export_step_visibility(self, export_mgr, export_path, custom_name, comp_name):
        """基于可见性导出STEP格式"""
        try:
            # 清理文件名
            safe_comp_name = self._sanitize_filename(comp_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_comp_name}-{safe_custom_name}.step'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 删除旧文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 导出当前可见的内容
            step_options = export_mgr.createSTEPExportOptions(filepath)
            step_options.sendToPrintUtility = False
            result = export_mgr.execute(step_options)
            
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_iges_visibility(self, export_mgr, export_path, custom_name, comp_name):
        """基于可见性导出IGES格式"""
        try:
            # 清理文件名
            safe_comp_name = self._sanitize_filename(comp_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_comp_name}-{safe_custom_name}.iges'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 删除旧文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 导出当前可见的内容
            iges_options = export_mgr.createIGESExportOptions('')
            iges_options.filename = filepath
            iges_options.sendToPrintUtility = False
            result = export_mgr.execute(iges_options)
            
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_stl_visibility(self, export_mgr, export_path, custom_name, comp_name, occurrence):
        """基于可见性导出STL格式"""
        try:
            if not occurrence:
                return False
                
            # 清理文件名
            safe_comp_name = self._sanitize_filename(comp_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_comp_name}-{safe_custom_name}.stl'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 删除旧文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # STL需要指定组件
            component = occurrence.component
            if component.bRepBodies.count == 0:
                return False
            
            stl_options = export_mgr.createSTLExportOptions(component)
            stl_options.filename = filepath
            stl_options.sendToPrintUtility = False
            stl_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            stl_options.isBinaryFormat = True
            
            result = export_mgr.execute(stl_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_obj_visibility(self, export_mgr, export_path, custom_name, comp_name, occurrence):
        """基于可见性导出OBJ格式"""
        try:
            if not occurrence:
                return False
                
            # 清理文件名
            safe_comp_name = self._sanitize_filename(comp_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_comp_name}-{safe_custom_name}.obj'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 删除旧文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # OBJ需要指定组件
            component = occurrence.component
            if component.bRepBodies.count == 0:
                return False
            
            obj_options = export_mgr.createOBJExportOptions(component)
            obj_options.filename = filepath
            obj_options.sendToPrintUtility = False
            obj_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            
            result = export_mgr.execute(obj_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_3mf_visibility(self, export_mgr, export_path, custom_name, comp_name, occurrence):
        """基于可见性导出3MF格式"""
        try:
            if not occurrence:
                return False
                
            # 清理文件名
            safe_comp_name = self._sanitize_filename(comp_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_comp_name}-{safe_custom_name}.3mf'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 删除旧文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 3MF需要指定组件
            component = occurrence.component
            if component.bRepBodies.count == 0:
                return False
            
            threemf_options = export_mgr.create3MFExportOptions(component)
            threemf_options.filename = filepath
            threemf_options.sendToPrintUtility = False
            threemf_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            
            result = export_mgr.execute(threemf_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_step(self, export_mgr, export_path, custom_name, component):
        """导出 STEP 格式 - 只导出指定组件"""
        try:
            # 获取组件名称，如果为空则使用默认名称
            component_name = component.name
            if not component_name or component_name.strip() == '':
                component_name = f'Component_{id(component)}'  # 使用唯一ID
            
            # 清理文件名中的非法字符
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            # 构建文件名
            filename = f'{safe_component_name}-{safe_custom_name}.step'
            
            # 确保路径格式正确
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 验证路径
            if not filepath or not os.path.exists(export_path):
                return False
            
            # 如果文件已存在，先删除
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 创建STEP导出选项 - 关键：指定要导出的组件
            try:
                # 创建一个临时的ObjectCollection，只包含要导出的组件的实体
                object_collection = adsk.core.ObjectCollection.create()
                
                # 添加该组件的所有实体
                for body in component.bRepBodies:
                    object_collection.add(body)
                
                # 如果没有实体，跳过这个组件
                if object_collection.count == 0:
                    return False
                
                # 使用createSTEPExportOptions的重载版本，传入要导出的对象集合
                step_options = export_mgr.createSTEPExportOptions(object_collection, filepath)
                step_options.sendToPrintUtility = False
                result = export_mgr.execute(step_options)
                
                if not result:
                    # 尝试备用方法：创建空选项然后设置文件名和对象
                    step_options = export_mgr.createSTEPExportOptions(object_collection, '')
                    step_options.filename = filepath
                    step_options.sendToPrintUtility = False
                    result = export_mgr.execute(step_options)
                    
            except Exception as e:
                # 如果上面的方法失败，尝试传统方法但临时隐藏其他组件
                return self._export_step_with_visibility(export_mgr, export_path, custom_name, component)
            
            # 验证文件是否真的被创建
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_step_with_visibility(self, export_mgr, export_path, custom_name, target_component):
        """通过控制可见性来导出单个组件的STEP"""
        try:
            # 获取设计
            design = self.app.activeProduct
            root_component = design.rootComponent
            
            # 记录原始可见性状态
            original_visibility = {}
            
            # 隐藏所有其他组件
            for occurrence in root_component.allOccurrences:
                original_visibility[occurrence.entityToken] = occurrence.isVisible
                if occurrence.component != target_component:
                    occurrence.isVisible = False
            
            # 确保目标组件可见
            for occurrence in root_component.allOccurrences:
                if occurrence.component == target_component:
                    occurrence.isVisible = True
            
            # 如果目标组件是根组件，隐藏其他根级实体
            if target_component == root_component:
                # 隐藏子组件，只保留根组件的实体
                for occurrence in root_component.occurrences:
                    original_visibility[occurrence.entityToken] = occurrence.isVisible
                    occurrence.isVisible = False
            
            # 构建文件路径
            component_name = target_component.name or f'Component_{id(target_component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            filename = f'{safe_component_name}-{safe_custom_name}.step'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 执行导出
            step_options = export_mgr.createSTEPExportOptions(filepath)
            step_options.sendToPrintUtility = False
            result = export_mgr.execute(step_options)
            
            # 恢复原始可见性
            for occurrence in root_component.allOccurrences:
                if occurrence.entityToken in original_visibility:
                    occurrence.isVisible = original_visibility[occurrence.entityToken]
            
            # 恢复根组件的子组件可见性
            if target_component == root_component:
                for occurrence in root_component.occurrences:
                    if occurrence.entityToken in original_visibility:
                        occurrence.isVisible = original_visibility[occurrence.entityToken]
            
            return result and os.path.exists(filepath)
            
        except Exception as e:
            # 确保恢复可见性
            try:
                design = self.app.activeProduct
                root_component = design.rootComponent
                for occurrence in root_component.allOccurrences:
                    if occurrence.entityToken in original_visibility:
                        occurrence.isVisible = original_visibility[occurrence.entityToken]
            except:
                pass
            return False
    
    def _export_iges(self, export_mgr, export_path, custom_name, component):
        """导出 IGES 格式 - 只导出指定组件"""
        try:
            # 获取和清理文件名
            component_name = component.name or f'Component_{id(component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_component_name}-{safe_custom_name}.iges'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 如果文件已存在，先删除
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            try:
                # 创建一个临时的ObjectCollection，只包含要导出的组件的实体
                object_collection = adsk.core.ObjectCollection.create()
                
                # 添加该组件的所有实体
                for body in component.bRepBodies:
                    object_collection.add(body)
                
                # 如果没有实体，跳过这个组件
                if object_collection.count == 0:
                    return False
                
                # 使用createIGESExportOptions的重载版本
                iges_options = export_mgr.createIGESExportOptions(object_collection, filepath)
                iges_options.sendToPrintUtility = False
                result = export_mgr.execute(iges_options)
                
                if not result:
                    iges_options = export_mgr.createIGESExportOptions(object_collection, '')
                    iges_options.filename = filepath
                    iges_options.sendToPrintUtility = False
                    result = export_mgr.execute(iges_options)
                    
            except Exception as e:
                # 如果失败，使用可见性方法
                return self._export_iges_with_visibility(export_mgr, export_path, custom_name, component)
            
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_iges_with_visibility(self, export_mgr, export_path, custom_name, target_component):
        """通过控制可见性来导出单个组件的IGES"""
        try:
            # 获取设计
            design = self.app.activeProduct
            root_component = design.rootComponent
            
            # 记录原始可见性状态
            original_visibility = {}
            
            # 隐藏所有其他组件
            for occurrence in root_component.allOccurrences:
                original_visibility[occurrence] = occurrence.isVisible
                if occurrence.component != target_component:
                    occurrence.isVisible = False
            
            # 确保目标组件可见
            for occurrence in root_component.allOccurrences:
                if occurrence.component == target_component:
                    occurrence.isVisible = True
            
            # 如果目标组件是根组件，隐藏其他子组件
            if target_component == root_component:
                for occurrence in root_component.occurrences:
                    occurrence.isVisible = False
            
            # 构建文件路径
            component_name = target_component.name or f'Component_{id(target_component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            filename = f'{safe_component_name}-{safe_custom_name}.iges'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 执行导出
            iges_options = export_mgr.createIGESExportOptions('')
            iges_options.filename = filepath
            iges_options.sendToPrintUtility = False
            result = export_mgr.execute(iges_options)
            
            # 恢复原始可见性
            for occurrence in root_component.allOccurrences:
                if occurrence.entityToken in original_visibility:
                    occurrence.isVisible = original_visibility[occurrence.entityToken]
            
            # 恢复根组件的子组件可见性
            if target_component == root_component:
                for occurrence in root_component.occurrences:
                    if occurrence.entityToken in original_visibility:
                        occurrence.isVisible = original_visibility[occurrence.entityToken]
            
            return result and os.path.exists(filepath)
            
        except Exception as e:
            # 确保恢复可见性
            try:
                design = self.app.activeProduct
                root_component = design.rootComponent
                for occurrence in root_component.allOccurrences:
                    if occurrence.entityToken in original_visibility:
                        occurrence.isVisible = original_visibility[occurrence.entityToken]
            except:
                pass
            return False
    
    def _export_stl(self, export_mgr, export_path, custom_name, component):
        """导出 STL 格式 - 只导出指定组件"""
        try:
            # 获取和清理文件名
            component_name = component.name or f'Component_{id(component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_component_name}-{safe_custom_name}.stl'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 如果文件已存在，先删除
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 检查组件是否有实体
            if component.bRepBodies.count == 0:
                return False
            
            # STL导出选项直接指定组件，这应该只导出该组件
            stl_options = export_mgr.createSTLExportOptions(component)
            stl_options.filename = filepath
            stl_options.sendToPrintUtility = False
            stl_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            stl_options.isBinaryFormat = True
            
            result = export_mgr.execute(stl_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_obj(self, export_mgr, export_path, custom_name, component):
        """导出 OBJ 格式 - 只导出指定组件"""
        try:
            # 获取和清理文件名
            component_name = component.name or f'Component_{id(component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_component_name}-{safe_custom_name}.obj'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 如果文件已存在，先删除
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 检查组件是否有实体
            if component.bRepBodies.count == 0:
                return False
            
            # OBJ导出选项直接指定组件，这应该只导出该组件
            obj_options = export_mgr.createOBJExportOptions(component)
            obj_options.filename = filepath
            obj_options.sendToPrintUtility = False
            obj_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            
            result = export_mgr.execute(obj_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False
    
    def _export_3mf(self, export_mgr, export_path, custom_name, component):
        """导出 3MF 格式 - 只导出指定组件"""
        try:
            # 获取和清理文件名
            component_name = component.name or f'Component_{id(component)}'
            safe_component_name = self._sanitize_filename(component_name)
            safe_custom_name = self._sanitize_filename(custom_name)
            
            filename = f'{safe_component_name}-{safe_custom_name}.3mf'
            filepath = os.path.normpath(os.path.join(export_path, filename))
            
            # 如果文件已存在，先删除
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    return False
            
            # 检查组件是否有实体
            if component.bRepBodies.count == 0:
                return False
            
            # 3MF导出选项直接指定组件，这应该只导出该组件
            threemf_options = export_mgr.create3MFExportOptions(component)
            threemf_options.filename = filepath
            threemf_options.sendToPrintUtility = False
            threemf_options.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            
            result = export_mgr.execute(threemf_options)
            return result and os.path.exists(filepath)
            
        except Exception as e:
            return False

class ParameterManager:
    """参数管理器"""
    
    def __init__(self):
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
    
    def get_starred_parameters(self, design):
        """获取标星参数"""
        parameters = []
        debug_info = []
        
        try:
            # 方法1: 检查用户参数
            user_params = design.userParameters
            user_count = user_params.count
            debug_info.append(f"用户参数总数: {user_count}")
            
            user_starred_count = 0
            for i in range(user_count):
                param = user_params.item(i)
                
                if param.isFavorite:
                    user_starred_count += 1
                    parameters.append({
                        'name': param.name,
                        'expression': param.expression,
                        'value': param.value,
                        'unit': param.unit if param.unit else '',
                        'comment': param.comment if param.comment else ''
                    })
            
            debug_info.append(f"用户参数中标星的: {user_starred_count}")
            
            # 方法2: 通过 AllParameters 获取所有参数（包括模型参数）
            try:
                all_params = design.allParameters
                debug_info.append(f"所有参数总数: {all_params.count}")
                
                all_starred_count = 0
                for param in all_params:
                    if hasattr(param, 'isFavorite') and param.isFavorite:
                        all_starred_count += 1
                        # 避免重复添加已经在用户参数中的参数
                        if not any(p['name'] == param.name for p in parameters):
                            parameters.append({
                                'name': param.name,
                                'expression': param.expression,
                                'value': param.value,
                                'unit': param.unit if param.unit else '',
                                'comment': param.comment if param.comment else ''
                            })
                
                debug_info.append(f"所有参数中标星的: {all_starred_count}")
                
            except Exception as e2:
                debug_info.append(f"无法访问 allParameters: {str(e2)}")
                # 如果 allParameters 不可用，使用递归方法
                self._check_component_parameters(design.rootComponent, parameters)
            
            # 显示调试信息（可选，已禁用）
            # debug_msg = '\n'.join(debug_info) + f'\n\n最终找到标星参数: {len(parameters)}'
            # for i, param in enumerate(parameters):
            #     debug_msg += f'\n{i+1}. {param["name"]} = {param["expression"]}'
            # 
            # self.ui.messageBox(f'调试信息:\n{debug_msg}')
            # LogUtils.info(f'调试信息: {debug_msg}')
                    
        except Exception as e:
            LogUtils.error(f'获取参数时发生错误: {str(e)}')
            
        return parameters
    
    def _check_component_parameters(self, component, parameters):
        """递归检查组件中的所有参数"""
        try:
            # 检查特征参数
            for feature in component.features:
                if hasattr(feature, 'parameters'):
                    for param in feature.parameters:
                        if hasattr(param, 'isFavorite') and param.isFavorite:
                            # 避免重复添加
                            if not any(p['name'] == param.name for p in parameters):
                                parameters.append({
                                    'name': param.name,
                                    'expression': param.expression,
                                    'value': param.value,
                                    'unit': param.unit if param.unit else '',
                                    'comment': param.comment if param.comment else ''
                                })
            
            # 检查草图参数
            for sketch in component.sketches:
                for dimension in sketch.sketchDimensions:
                    if hasattr(dimension, 'parameter') and dimension.parameter:
                        param = dimension.parameter
                        if hasattr(param, 'isFavorite') and param.isFavorite:
                            # 避免重复添加
                            if not any(p['name'] == param.name for p in parameters):
                                parameters.append({
                                    'name': param.name,
                                    'expression': param.expression,
                                    'value': param.value,
                                    'unit': param.unit if param.unit else '',
                                    'comment': param.comment if param.comment else ''
                                })
            
            # 递归检查子组件
            for occurrence in component.allOccurrences:
                if occurrence.component != component:  # 避免无限递归
                    self._check_component_parameters(occurrence.component, parameters)
                    
        except Exception as e:
            # 忽略单个组件的错误，继续处理其他组件
            pass
    
    def apply_parameters(self, design, parameters):
        """应用参数值"""
        try:
            success_count = 0
            total_count = len(parameters)
            
            for param_name, param_value in parameters.items():
                # 首先尝试用户参数
                user_param = design.userParameters.itemByName(param_name)
                if user_param:
                    user_param.expression = str(param_value)
                    success_count += 1
                    continue
                
                # 如果不是用户参数，尝试在所有参数中查找
                try:
                    all_params = design.allParameters
                    for param in all_params:
                        if param.name == param_name:
                            param.expression = str(param_value)
                            success_count += 1
                            break
                except:
                    pass
            
            # 触发ParametricText插件更新事件
            try:
                app = adsk.core.Application.get()
                app.fireCustomEvent('thomasa88_ParametricText_Ext_Update')
                LogUtils.info('已触发ParametricText更新事件')
            except Exception as e:
                LogUtils.warn(f'触发ParametricText更新事件失败: {str(e)}')
            
            # 等待ParametricText插件处理完成
            try:
                import time
                time.sleep(1)  # 等待1秒让ParametricText完成更新
                LogUtils.info('等待ParametricText更新完成')
            except Exception as e:
                LogUtils.warn(f'等待ParametricText更新时发生错误: {str(e)}')
            
            # 重新计算设计
            design.computeAll()
            
            if success_count < total_count:
                LogUtils.warn(f'警告: 只有 {success_count}/{total_count} 个参数被成功应用')
            
            return success_count > 0
            
        except Exception as e:
            LogUtils.error(f'应用参数时发生错误: {str(e)}')
            return False
    
    def backup_parameters(self, design):
        """备份当前参数值"""
        backup = {}
        
        try:
            user_params = design.userParameters
            
            for param in user_params:
                if param.isFavorite:
                    backup[param.name] = param.expression
                    
        except Exception as e:
            LogUtils.error(f'备份参数时发生错误: {str(e)}')
            
        return backup
    
    def restore_parameters(self, design, backup):
        """恢复参数值"""
        try:
            user_params = design.userParameters
            
            for param_name, param_value in backup.items():
                param = user_params.itemByName(param_name)
                if param:
                    param.expression = param_value
            
            # 触发ParametricText插件更新事件
            try:
                app = adsk.core.Application.get()
                app.fireCustomEvent('thomasa88_ParametricText_Ext_Update')
                LogUtils.info('已触发ParametricText更新事件（参数恢复）')
            except Exception as e:
                LogUtils.warn(f'触发ParametricText更新事件失败（参数恢复）: {str(e)}')
            
            # 等待ParametricText插件处理完成
            try:
                import time
                time.sleep(1)  # 等待1秒让ParametricText完成更新
                LogUtils.info('等待ParametricText更新完成（参数恢复）')
            except Exception as e:
                LogUtils.warn(f'等待ParametricText更新时发生错误（参数恢复）: {str(e)}')
            
            # 重新计算设计
            design.computeAll()
            return True
            
        except Exception as e:
            LogUtils.error(f'恢复参数时发生错误: {str(e)}')
            return False 