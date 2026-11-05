# Excel自动循环加载器节点
import os
import time
import traceback
from typing import Dict, List, Tuple, Any

# 尝试导入pandas
try:
    import pandas as pd
    print("✅ [DEBUG] pandas 导入成功 (ExcelAutoLoopLoader)")
except ImportError as e:
    print(f"❌ [ERROR] ExcelAutoLoopLoader 节点缺少 Pandas 库: {e}")
    print("请在ComfyUI根目录运行: python_embeded\\python.exe -m pip install pandas openpyxl")
    pd = None

# 尝试导入openpyxl
try:
    import openpyxl
    print("✅ [DEBUG] openpyxl 导入成功 (ExcelAutoLoopLoader)")
except ImportError as e:
    print(f"❌ [ERROR] openpyxl 导入失败 (ExcelAutoLoopLoader): {e}")

class ExcelAutoLoopLoader:
    """
    Excel自动循环加载器节点
    自动循环读取Excel中的每一行，无需外部for循环节点
    """
    
    def __init__(self):
        self._cached_file_path = None
        self._cached_file_mtime = None
        self._cached_headers = []
        self._cached_total_rows = 0
        self._current_row = 2  # 从第2行开始（跳过表头）
    
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
                "loop_mode": (["single_row", "all_rows"], {
                    "default": "single_row",
                    "description": "循环模式：single_row=逐行循环，all_rows=一次读取所有行"
                }),
                "reset_loop": ("BOOLEAN", {
                    "default": False,
                    "description": "重置循环（回到第一行）"
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
    
    RETURN_TYPES = tuple(["STRING"] * 12)
    RETURN_NAMES = tuple([f"column_{i+1}" for i in range(9)] + ["current_row", "total_rows", "loop_status"])
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
    
    def execute(self, excel_file_path, loop_mode, reset_loop, 
                refresh_trigger, auto_refresh, sheet_name="0", **kwargs):
        """
        执行节点逻辑 - 自动循环版本
        """
        if pd is None:
            return tuple(["错误：运行此节点需要 Pandas 库，请先安装"] + [""] * 11)
        
        try:
            # 检查是否需要刷新缓存
            if self._should_refresh_cache(excel_file_path, refresh_trigger, auto_refresh):
                self._update_cache(excel_file_path, sheet_name)
                print(f"[ExcelAutoLoopLoader] 缓存已刷新，发现 {len(self._cached_headers)} 列")
                for i, header in enumerate(self._cached_headers):
                    print(f"  列{i+1}: {header} -> text")
            
            # 如果没有缓存数据，尝试初始化
            if not self._cached_headers:
                self._update_cache(excel_file_path, sheet_name)
            
            # 重置循环
            if reset_loop:
                self._current_row = 2
                print(f"[ExcelAutoLoopLoader] 循环已重置到第2行")
            
            # 读取实际数据
            sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
            df = pd.read_excel(
                excel_file_path,
                header=0,  # 第一行作为表头
                engine='openpyxl',
                sheet_name=sheet_identifier
            )
            
            data_rows_count = len(df)
            
            if loop_mode == "all_rows":
                # 一次性读取所有行
                all_data = []
                for row_idx in range(data_rows_count):
                    row_data = []
                    for i, header in enumerate(self._cached_headers[:9]):
                        if header in df.columns:
                            cell_value = str(df.iloc[row_idx][header]) if pd.notna(df.iloc[row_idx][header]) else ""
                            row_data.append(cell_value)
                        else:
                            row_data.append("")
                    # 填充到9列
                    while len(row_data) < 9:
                        row_data.append("")
                    all_data.append(row_data)
                
                # 将所有行数据合并，用换行符分隔
                combined_columns = []
                for col_idx in range(9):
                    column_data = [row[col_idx] for row in all_data if row[col_idx]]
                    combined_columns.append("\n".join(column_data))
                
                result = combined_columns + [f"1-{data_rows_count}", str(data_rows_count), "all_rows_complete"]
                return tuple(result)
            
            else:  # single_row mode
                # 检查是否超出数据范围
                if self._current_row - 2 >= data_rows_count:
                    # 循环结束，重置到开始
                    self._current_row = 2
                    loop_status = "loop_complete_reset"
                else:
                    loop_status = f"reading_row_{self._current_row}"
                
                # 读取当前行
                data_start_index = self._current_row - 2  # -2 因为_current_row是从2开始（跳过表头）
                
                if data_start_index < data_rows_count:
                    # 处理每一列的数据
                    column_results = []
                    
                    # 处理实际存在的列（最多9列）
                    for i, header in enumerate(self._cached_headers[:9]):
                        if header in df.columns:
                            cell_value = df.iloc[data_start_index][header]
                            cell_str = str(cell_value) if pd.notna(cell_value) else ""
                            column_results.append(cell_str)
                        else:
                            column_results.append("")
                    
                    # 填充剩余的输出槽位为空字符串（确保总共9个列输出）
                    while len(column_results) < 9:
                        column_results.append("")
                    
                    # 添加额外信息
                    column_results.append(str(self._current_row))  # 当前行号
                    column_results.append(str(data_rows_count))    # 总行数
                    column_results.append(loop_status)             # 循环状态
                    
                    # 准备下一次执行
                    self._current_row += 1
                    
                    print(f"[ExcelAutoLoopLoader] 读取第{self._current_row-1}行，下次将读取第{self._current_row}行")
                    
                    return tuple(column_results)
                else:
                    # 超出范围，返回空值
                    result = [""] * 9 + [str(self._current_row), str(data_rows_count), "out_of_range"]
                    return tuple(result)
            
        except Exception as e:
            print(f"[ExcelAutoLoopLoader] 错误: {e}")
            traceback.print_exc()
            # 返回固定12个输出，第一个包含错误信息，其余为空
            error_result = [f"错误: {e}"] + [""] * 11
            return tuple(error_result)

# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "ExcelAutoLoopLoader": ExcelAutoLoopLoader
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "ExcelAutoLoopLoader": "Excel Auto Loop Loader"
}

print("✅ [DEBUG] ExcelAutoLoopLoader 注册成功")
