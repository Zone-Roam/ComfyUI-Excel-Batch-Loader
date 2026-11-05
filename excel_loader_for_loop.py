# Excel加载器节点 - For循环专用版本
import os
import time
import traceback
from typing import Dict, List, Tuple, Any

# 尝试导入pandas
try:
    import pandas as pd
    print("✅ [DEBUG] pandas 导入成功 (ExcelLoaderForLoop)")
except ImportError as e:
    print(f"❌ [ERROR] ExcelLoaderForLoop 节点缺少 Pandas 库: {e}")
    print("请在ComfyUI根目录运行: python_embeded\\python.exe -m pip install pandas openpyxl")
    pd = None

# 尝试导入openpyxl
try:
    import openpyxl
    print("✅ [DEBUG] openpyxl 导入成功 (ExcelLoaderForLoop)")
except ImportError as e:
    print(f"❌ [ERROR] openpyxl 导入失败 (ExcelLoaderForLoop): {e}")

class ExcelLoaderForLoop:
    """
    Excel加载器节点 - For循环专用版本
    专门设计用于与for循环节点配合使用，移除了control_after_generate属性
    """
    
    def __init__(self):
        self._cached_file_path = None
        self._cached_file_mtime = None
        self._cached_headers = []
        self._cached_column_types = []
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

                "read_count": ("INT", {
                    "default": 1, 
                    "min": 1, 
                    "step": 1,
                    "description": "每次执行读取的行数"
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
                "index": ("INT", {
                    "default": 2,
                    "description": "来自EasyUse forLoopStart的索引值"
                }),
                "sheet_name": ("STRING", {
                    "multiline": False,
                    "default": "0",
                    "description": "要读取的工作表名称或0基索引"
                })
            }
        }
    
    RETURN_TYPES = tuple(["STRING"] * 10)
    RETURN_NAMES = tuple([f"column_{i+1}" for i in range(9)] + ["current_row"])
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
        self._cached_column_types = ['text'] * len(headers)  # 所有列都是文本类型
        self._cached_total_rows = total_rows
    
    def execute(self, excel_file_path, read_count,
                refresh_trigger, auto_refresh, index=None, sheet_name="0", **kwargs):
        """
        执行节点逻辑 - EasyUse For循环适配版本
        """
        # 确定实际的行号
        # 如果有index输入（来自forLoopStart），使用index作为行号
        # 否则使用默认的第2行（跳过表头）
        start_row_number = index if index is not None else 2

        # 添加详细的调试信息
        print(f"🔍 [ExcelLoaderForLoop] 执行开始:")
        print(f"  📁 文件路径: {excel_file_path}")
        print(f"  📊 索引值: {index}")
        print(f"  📊 实际行号: {start_row_number}")
        print(f"  📖 读取行数: {read_count}")
        print(f"  🔄 刷新触发: {refresh_trigger}")
        print(f"  ⚡ 自动刷新: {auto_refresh}")
        print(f"  📋 工作表: {sheet_name}")
        print(f"  🚨 关键检查: index是否为None? {index is None}")
        print(f"  🚨 forLoopStart是否正确传递index? 期望值应该是1或2")

        if pd is None:
            return tuple(["错误：运行此节点需要 Pandas 库，请先安装"] + [""] * 9)

        try:
            # 检查是否需要刷新缓存
            if self._should_refresh_cache(excel_file_path, refresh_trigger, auto_refresh):
                self._update_cache(excel_file_path, sheet_name)
                print(f"[ExcelLoaderForLoop] 缓存已刷新，发现 {len(self._cached_headers)} 列")
                for i, header in enumerate(self._cached_headers):
                    print(f"  列{i+1}: {header} -> text")
            
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
            
            # 计算实际读取范围
            data_start_index = start_row_number - 2  # -2 因为start_row_number是从2开始（跳过表头）
            data_end_index = min(data_start_index + read_count, len(df))

            print(f"  📊 数据范围计算:")
            print(f"    总数据行数: {len(df)}")
            print(f"    数据开始索引: {data_start_index}")
            print(f"    数据结束索引: {data_end_index}")

            if data_start_index >= len(df):
                # 超出数据范围，返回固定10个输出
                print(f"  ⚠️  超出数据范围，返回空值")
                result_values = [""] * 9 + [str(start_row_number)]
                return tuple(result_values)
            
            # 提取数据
            target_rows = df.iloc[data_start_index:data_end_index]
            
            # 处理每一列的数据
            column_results = []
            
            # 处理实际存在的列（最多9列）
            for i, header in enumerate(self._cached_headers[:9]):
                if header in df.columns:
                    # 获取该列的所有读取行数据
                    column_data = target_rows[header].fillna("").astype(str)
                    # 如果读取多行，用换行符连接
                    combined_data = "\n".join(column_data.tolist())
                    column_results.append(combined_data)
                else:
                    column_results.append("")
            
            # 填充剩余的输出槽位为空字符串（确保总共9个列输出）
            while len(column_results) < 9:
                column_results.append("")
            
            # 添加当前行号作为第10个输出
            column_results.append(str(start_row_number))

            print(f"  ✅ 成功读取 {len(target_rows)} 行数据")
            print(f"  📤 返回结果: {len(column_results)} 个输出")

            # 注意：不返回UI更新信息，避免与for循环冲突
            return tuple(column_results)
            
        except Exception as e:
            print(f"[ExcelLoaderForLoop] 错误: {e}")
            traceback.print_exc()
            # 返回固定10个输出，第一个包含错误信息，其余为空
            error_result = [f"错误: {e}"] + [""] * 9
            return tuple(error_result)

# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "ExcelLoaderForLoop": ExcelLoaderForLoop
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "ExcelLoaderForLoop": "Excel Loader (For Loop)"
}

print("✅ [DEBUG] ExcelLoaderForLoop 注册成功")
