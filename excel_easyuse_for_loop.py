# Excel EasyUse For循环适配加载器节点
import os
import time
import traceback
from typing import Dict, List, Tuple, Any

# 尝试导入pandas
try:
    import pandas as pd
    print("✅ [DEBUG] pandas 导入成功 (ExcelEasyUseForLoop)")
except ImportError as e:
    print(f"❌ [ERROR] ExcelEasyUseForLoop 节点缺少 Pandas 库: {e}")
    print("请在ComfyUI根目录运行: python_embeded\\python.exe -m pip install pandas openpyxl")
    pd = None

# 尝试导入openpyxl
try:
    import openpyxl
    print("✅ [DEBUG] openpyxl 导入成功 (ExcelEasyUseForLoop)")
except ImportError as e:
    print(f"❌ [ERROR] openpyxl 导入失败 (ExcelEasyUseForLoop): {e}")

class ExcelEasyUseForLoop:
    """
    Excel EasyUse For循环适配加载器节点
    专门适配EasyUse的forLoopStart和forLoopEnd节点
    接收forLoopStart输出的index值作为行号
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
                "index": ("INT", {
                    "default": 2,
                    "description": "来自EasyUse forLoopStart的索引值（作为行号）"
                }),
                "refresh_trigger": ("BOOLEAN", {
                    "default": False,
                    "description": "手动刷新开关（切换此开关可强制重新解析Excel）"
                }),
                "auto_refresh": ("BOOLEAN", {
                    "default": True,
                    "description": "自动刷新（当Excel文件修改时间改变时自动重新解析）"
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
    CATEGORY = "ExcelUtils"
    
    def _parse_excel_structure(self, excel_file_path: str, sheet_name: str) -> Tuple[List[str], int]:
        """
        解析Excel结构，返回表头和总行数
        """
        if pd is None:
            raise ImportError("Pandas 库未安装")
        
        # 检查文件是否存在
        if not os.path.exists(excel_file_path):
            raise FileNotFoundError(f"文件未找到: {excel_file_path}")
        
        # 确定工作表标识符
        sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
        
        # 读取表头（第一行）
        header_df = pd.read_excel(
            excel_file_path, 
            header=0,  # 第一行作为表头
            nrows=0,   # 只读取表头
            engine='openpyxl',
            sheet_name=sheet_identifier
        )
        
        headers = list(header_df.columns)
        
        # 获取总行数（包括表头）
        full_df = pd.read_excel(
            excel_file_path,
            header=0,
            engine='openpyxl', 
            sheet_name=sheet_identifier
        )
        total_rows = len(full_df) + 1  # +1 因为表头
        
        return headers, total_rows
    
    def _should_refresh_cache(self, excel_file_path: str, refresh_trigger: bool, auto_refresh: bool) -> bool:
        """
        判断是否需要刷新缓存
        """
        # 手动刷新触发
        if refresh_trigger:
            return True
            
        # 文件路径改变
        if self._cached_file_path != excel_file_path:
            return True
            
        # 自动刷新：文件修改时间改变
        if auto_refresh and os.path.exists(excel_file_path):
            current_mtime = os.path.getmtime(excel_file_path)
            if self._cached_file_mtime != current_mtime:
                return True
        
        return False
    
    def _update_cache(self, excel_file_path: str, sheet_name: str):
        """
        更新缓存
        """
        headers, total_rows = self._parse_excel_structure(
            excel_file_path, sheet_name
        )
        
        self._cached_file_path = excel_file_path
        self._cached_file_mtime = os.path.getmtime(excel_file_path)
        self._cached_headers = headers
        self._cached_total_rows = total_rows
    
    def execute(self, excel_file_path, index, refresh_trigger, auto_refresh, sheet_name="0", **kwargs):
        """
        执行节点逻辑 - EasyUse For循环适配版本
        """
        if pd is None:
            return tuple(["错误：运行此节点需要 Pandas 库，请先安装"] + [""] * 9)
        
        try:
            # 检查是否需要刷新缓存
            if self._should_refresh_cache(excel_file_path, refresh_trigger, auto_refresh):
                self._update_cache(excel_file_path, sheet_name)
                print(f"[ExcelEasyUseForLoop] 缓存已刷新，发现 {len(self._cached_headers)} 列")
                for i, header in enumerate(self._cached_headers):
                    print(f"  列{i+1}: {header}")
            
            # 如果没有缓存数据，尝试初始化
            if not self._cached_headers:
                self._update_cache(excel_file_path, sheet_name)
            
            # 读取实际数据
            sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
            df = pd.read_excel(
                excel_file_path,
                header=0,  # 第一行作为表头
                engine='openpyxl',
                sheet_name=sheet_identifier
            )
            
            data_rows_count = len(df)
            
            # 使用index作为行号（index从forLoopStart传入）
            # index应该对应Excel的实际行号，需要转换为DataFrame的索引
            # Excel行号从1开始，DataFrame索引从0开始
            # 如果index=2（第2行，跳过表头），对应DataFrame索引0
            data_row_index = index - 2  # 减去表头行

            print(f"🔍 [ExcelEasyUseForLoop] 调试信息:")
        print(f"  📊 接收到的index: {index}")
        print(f"  📊 转换为数据行索引: {data_row_index}")
        print(f"  📊 总数据行数: {data_rows_count}")
        print(f"  📊 数据行范围: 2-{data_rows_count+1}")
            
            # 检查索引是否有效
            if data_row_index < 0 or data_row_index >= data_rows_count:
                # 超出范围，返回空值和状态信息
                result = [""] * 9 + [f"索引超出范围: {index} (数据行范围: 2-{data_rows_count+1})"]
                print(f"[ExcelEasyUseForLoop] 索引超出范围: {index}")
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
            
            # 填充剩余的输出槽位为空字符串（确保总共9个列输出）
            while len(column_results) < 9:
                column_results.append("")
            
            # 添加行信息
            row_info = f"行{index}: 数据行{data_row_index+1}/{data_rows_count}"
            column_results.append(row_info)
            
            print(f"[ExcelEasyUseForLoop] 成功读取第{index}行数据")
            
            return tuple(column_results)
            
        except Exception as e:
            print(f"[ExcelEasyUseForLoop] 错误: {e}")
            traceback.print_exc()
            # 返回固定10个输出，第一个包含错误信息，其余为空
            error_result = [f"错误: {e}"] + [""] * 9
            return tuple(error_result)

# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "ExcelEasyUseForLoop": ExcelEasyUseForLoop
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "ExcelEasyUseForLoop": "Excel EasyUse For Loop"
}

print("✅ [DEBUG] ExcelEasyUseForLoop 注册成功")
