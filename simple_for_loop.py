# 简单For循环节点 - 专为Excel数据处理设计
import time
from typing import Dict, Any, Tuple

class SimpleForLoopStart:
    """
    简单For循环开始节点
    可以指定循环次数，每次循环自动将index+1直到循环终止
    """
    
    # 类变量用于存储循环状态
    _loop_states = {}
    
    def __init__(self):
        pass
    
    @classmethod
    def INPUT_TYPES(cls):
        """定义节点的输入参数"""
        return {
            "required": {
                "loop_count": ("INT", {
                    "default": 2,
                    "min": 1,
                    "max": 1000,
                    "step": 1,
                    "description": "循环次数"
                }),
                "start_index": ("INT", {
                    "default": 1,
                    "min": 0,
                    "max": 1000,
                    "step": 1,
                    "description": "起始索引值"
                }),
                "reset_trigger": ("BOOLEAN", {
                    "default": False,
                    "description": "重置循环（切换此开关重新开始循环）"
                })
            }
        }
    
    RETURN_TYPES = ("INT", "BOOLEAN", "STRING")
    RETURN_NAMES = ("current_index", "is_finished", "loop_info")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils/ForLoop"
    
    def execute(self, loop_count, start_index, reset_trigger, **kwargs):
        """
        执行For循环开始逻辑
        """
        # 生成稳定的循环ID（基于参数组合，不包含时间）
        loop_id = f"loop_{loop_count}_{start_index}"

        # 检查是否需要重置或初始化循环状态
        if reset_trigger or loop_id not in self._loop_states:
            self._loop_states[loop_id] = {
                "current_iteration": 0,
                "current_index": start_index,
                "total_count": loop_count,
                "start_index": start_index,
                "is_finished": False,
                "created_time": time.time()
            }
            print(f"🔄 [SimpleForLoopStart] 初始化循环: {loop_id}")
            print(f"  📊 循环次数: {loop_count}")
            print(f"  📊 起始索引: {start_index}")

        # 获取当前循环状态
        state = self._loop_states[loop_id]
        
        # 检查是否已完成
        if state["current_iteration"] >= state["total_count"]:
            state["is_finished"] = True
            loop_info = f"循环已完成 ({state['current_iteration']}/{state['total_count']})"
            print(f"✅ [SimpleForLoopStart] {loop_info}")
            return (state["current_index"], True, loop_info)
        
        # 计算当前索引
        current_index = state["start_index"] + state["current_iteration"]
        state["current_index"] = current_index
        
        # 准备循环信息
        loop_info = f"循环进行中: 第{state['current_iteration'] + 1}次/{state['total_count']}, 索引={current_index}"
        
        print(f"🔁 [SimpleForLoopStart] {loop_info}")
        
        # 更新迭代计数（为下次执行准备）
        state["current_iteration"] += 1
        
        return (current_index, False, loop_info)


class SimpleForLoopEnd:
    """
    简单For循环结束节点
    接收循环体的输出，并控制循环的继续或结束
    """
    
    def __init__(self):
        pass
    
    @classmethod
    def INPUT_TYPES(cls):
        """定义节点的输入参数"""
        return {
            "required": {
                "is_finished": ("BOOLEAN", {
                    "description": "来自ForLoopStart的完成状态"
                }),
                "loop_info": ("STRING", {
                    "description": "来自ForLoopStart的循环信息"
                })
            },
            "optional": {
                "loop_data": ("*", {
                    "description": "循环体产生的数据（可选）"
                })
            }
        }
    
    RETURN_TYPES = ("STRING", "BOOLEAN", "*")
    RETURN_NAMES = ("final_status", "completed", "collected_data")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils/ForLoop"
    
    def execute(self, is_finished, loop_info, loop_data=None, **kwargs):
        """
        执行For循环结束逻辑
        """
        print(f"🏁 [SimpleForLoopEnd] 接收状态: {loop_info}")
        
        if is_finished:
            final_status = f"循环已完成: {loop_info}"
            print(f"✅ [SimpleForLoopEnd] {final_status}")
            return (final_status, True, loop_data)
        else:
            status = f"循环继续: {loop_info}"
            print(f"🔄 [SimpleForLoopEnd] {status}")
            return (status, False, loop_data)


class SimpleExcelForLoop:
    """
    简单Excel For循环节点
    与SimpleForLoopStart配合使用，根据index读取Excel数据
    """
    
    def __init__(self):
        self._cached_file_path = None
        self._cached_file_mtime = None
        self._cached_headers = []
        self._cached_total_rows = 0
    
    @classmethod
    def INPUT_TYPES(cls):
        """定义节点的输入参数"""
        return {
            "required": {
                "excel_file_path": ("STRING", {
                    "multiline": False,
                    "default": "example.xlsx",
                    "description": "Excel 文件路径 (.xlsx 或 .xls)"
                }),
                "current_index": ("INT", {
                    "description": "来自SimpleForLoopStart的当前索引"
                }),
                "skip_header": ("BOOLEAN", {
                    "default": True,
                    "description": "是否跳过表头行"
                })
            },
            "optional": {
                "sheet_name": ("STRING", {
                    "multiline": False,
                    "default": "0",
                    "description": "要读取的工作表名称或0基索引"
                })
            }
        }
    
    RETURN_TYPES = tuple(["STRING"] * 10)
    RETURN_NAMES = tuple([f"column_{i+1}" for i in range(9)] + ["row_info"])
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils/ForLoop"
    
    def _should_refresh_cache(self, excel_file_path, auto_refresh=True):
        """检查是否需要刷新缓存"""
        import os
        
        if not os.path.exists(excel_file_path):
            return False
            
        if self._cached_file_path != excel_file_path:
            return True
            
        if auto_refresh:
            current_mtime = os.path.getmtime(excel_file_path)
            if self._cached_file_mtime != current_mtime:
                return True
        
        return False
    
    def _update_cache(self, excel_file_path, sheet_name):
        """更新缓存"""
        import pandas as pd
        import os
        
        # 确定工作表标识符
        sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
        
        # 读取Excel文件
        df = pd.read_excel(
            excel_file_path,
            header=0,  # 第一行作为表头
            engine='openpyxl',
            sheet_name=sheet_identifier
        )
        
        self._cached_file_path = excel_file_path
        self._cached_file_mtime = os.path.getmtime(excel_file_path)
        self._cached_headers = list(df.columns)
        self._cached_total_rows = len(df) + 1  # +1 因为表头
    
    def execute(self, excel_file_path, current_index, skip_header, sheet_name="0", **kwargs):
        """
        执行Excel读取逻辑
        """
        try:
            import pandas as pd
            import os
        except ImportError as e:
            error_msg = f"缺少必要库: {e}"
            return tuple([error_msg] + [""] * 9)
        
        try:
            # 检查文件是否存在
            if not os.path.exists(excel_file_path):
                error_msg = f"文件不存在: {excel_file_path}"
                return tuple([error_msg] + [""] * 9)
            
            # 更新缓存
            if self._should_refresh_cache(excel_file_path):
                self._update_cache(excel_file_path, sheet_name)
                print(f"📊 [SimpleExcelForLoop] 缓存已更新，发现 {len(self._cached_headers)} 列")
            
            # 读取数据
            sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
            df = pd.read_excel(
                excel_file_path,
                header=0,
                engine='openpyxl',
                sheet_name=sheet_identifier
            )
            
            # 计算实际的数据行索引
            if skip_header:
                # 如果跳过表头，current_index=1对应数据的第1行（DataFrame索引0）
                data_row_index = current_index - 1
                excel_row_number = current_index + 1  # Excel中的实际行号（+1因为表头）
            else:
                # 如果不跳过表头，current_index=1对应Excel的第1行（表头）
                if current_index == 1:
                    # 返回表头
                    column_results = [str(header) for header in self._cached_headers[:9]]
                    while len(column_results) < 9:
                        column_results.append("")
                    row_info = f"表头行 (索引{current_index})"
                    column_results.append(row_info)
                    print(f"📋 [SimpleExcelForLoop] 读取表头")
                    return tuple(column_results)
                else:
                    data_row_index = current_index - 2  # -2因为索引1是表头
                    excel_row_number = current_index
            
            print(f"🔍 [SimpleExcelForLoop] 读取数据:")
            print(f"  📊 当前索引: {current_index}")
            print(f"  📊 数据行索引: {data_row_index}")
            print(f"  📊 Excel行号: {excel_row_number}")
            print(f"  📊 总数据行数: {len(df)}")
            
            # 检查索引是否有效
            if data_row_index < 0 or data_row_index >= len(df):
                error_msg = f"索引超出范围: 数据行索引{data_row_index} (总行数: {len(df)})"
                result = [error_msg] + [""] * 9
                print(f"⚠️ [SimpleExcelForLoop] {error_msg}")
                return tuple(result)
            
            # 读取指定行的数据
            column_results = []
            
            # 处理实际存在的列（最多9列）
            for i, header in enumerate(self._cached_headers[:9]):
                if header in df.columns:
                    cell_value = df.iloc[data_row_index][header]
                    cell_str = str(cell_value) if pd.notna(cell_value) else ""
                    column_results.append(cell_str)
                else:
                    column_results.append("")
            
            # 填充剩余的输出槽位
            while len(column_results) < 9:
                column_results.append("")
            
            # 添加行信息
            row_info = f"Excel第{excel_row_number}行 (索引{current_index}, 数据行{data_row_index+1}/{len(df)})"
            column_results.append(row_info)
            
            print(f"✅ [SimpleExcelForLoop] 成功读取数据")
            
            return tuple(column_results)
            
        except Exception as e:
            error_msg = f"读取错误: {e}"
            print(f"❌ [SimpleExcelForLoop] {error_msg}")
            import traceback
            traceback.print_exc()
            return tuple([error_msg] + [""] * 9)


# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "SimpleForLoopStart": SimpleForLoopStart,
    "SimpleForLoopEnd": SimpleForLoopEnd,
    "SimpleExcelForLoop": SimpleExcelForLoop
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "SimpleForLoopStart": "Simple For Loop Start",
    "SimpleForLoopEnd": "Simple For Loop End", 
    "SimpleExcelForLoop": "Simple Excel For Loop"
}

print("✅ [DEBUG] Simple For Loop 节点注册成功")
