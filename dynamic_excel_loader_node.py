# /ComfyUI/custom_nodes/ComfyUI_DynamicExcelLoader/dynamic_excel_loader_node.py

import os
import time
import traceback
from typing import Dict, List, Tuple, Any

# 尝试导入pandas
try:
    import pandas as pd
    print("✅ [DEBUG] pandas 导入成功")
except ImportError as e:
    print("\n##################################################################")
    print("❌ [ERROR] DynamicExcelLoader 节点缺少 Pandas 库。")
    print("请在ComfyUI根目录运行: python_embeded\\python.exe -m pip install pandas openpyxl")
    print("或者运行: pip install pandas openpyxl")
    print(f"详细错误: {e}")
    print("##################################################################\n")
    pd = None

# 尝试导入openpyxl
try:
    import openpyxl
    print("✅ [DEBUG] openpyxl 导入成功")
except ImportError as e:
    print(f"❌ [ERROR] openpyxl 导入失败: {e}")
    print("请安装: pip install openpyxl")

class DynamicExcelLoader:
    """
    Dynamic Excel Loader Node for ComfyUI
    根据Excel表头自动识别列类型，动态生成对应数量的输出组件
    支持文本和图像路径两种数据类型，具备刷新机制
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
                    "default": "input/example.xlsx", 
                    "description": "Excel 文件路径 (.xlsx 或 .xls)"
                }),
                "start_row_number": ("INT", {
                    "default": 2, 
                    "min": 2, 
                    "step": 1, 
                    "control_after_generate": True,
                    "description": "从哪一行开始读取数据（2表示跳过表头）"
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
                "sheet_name": ("STRING", {
                    "multiline": False, 
                    "default": "0", 
                    "description": "要读取的工作表名称或0基索引"
                })
            }
        }
    

    
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

    RETURN_TYPES = tuple(["STRING"] * 10)
    RETURN_NAMES = tuple([f"column_{i+1}" for i in range(9)] + ["current_row"])

    def _generate_dynamic_outputs(self) -> Tuple[Tuple[str, ...], Tuple[str, ...]]:
        """
        根据缓存的表头和列类型生成动态输出类型和名称
        """
        if not self._cached_headers:
            return ("STRING",), ("output",)

        output_types = []
        output_names = []

        for i, (header, col_type) in enumerate(zip(self._cached_headers, self._cached_column_types)):
            # 所有输出都是STRING类型，但名称会区分类型
            output_types.append("STRING")

            # 根据列类型生成输出名称
            if col_type == 'image_path':
                output_names.append(f"{header}_path")
            else:
                output_names.append(f"{header}_text")

        # 添加当前行号输出
        output_types.append("STRING")
        output_names.append("current_row")

        return tuple(output_types), tuple(output_names)

    def execute(self, excel_file_path, start_row_number, read_count,
                refresh_trigger, auto_refresh, sheet_name="0", **kwargs):
        """
        执行节点逻辑
        """
        if pd is None:
            return {"result": ("错误：运行此节点需要 Pandas 库，请先安装。",)}

        try:
            # 检查是否需要刷新缓存
            if self._should_refresh_cache(excel_file_path, refresh_trigger, auto_refresh):
                self._update_cache(excel_file_path, sheet_name)
                print(f"[DynamicExcelLoader] 缓存已刷新，发现 {len(self._cached_headers)} 列")
                for i, header in enumerate(self._cached_headers):
                    print(f"  列{i+1}: {header} -> text")

            # 如果没有缓存数据，尝试初始化
            if not self._cached_headers:
                self._update_cache(excel_file_path, sheet_name)

            # 动态设置返回类型和名称
            output_types, output_names = self._generate_dynamic_outputs()

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

            if data_start_index >= len(df):
                # 超出数据范围，返回固定10个输出
                result_values = [""] * 9 + [str(start_row_number)]
                return {"result": tuple(result_values)}

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

            # 计算下一个起始行号
            next_start_row = start_row_number + read_count

            # 生成列信息用于UI显示
            column_info = []
            for i, header in enumerate(self._cached_headers):
                column_info.append(f"列{i+1}: {header} (text)")

            return {
                "ui": {
                    "start_row_number": [next_start_row],
                    "column_info": column_info
                },
                "result": tuple(column_results)
            }

        except Exception as e:
            print(f"[DynamicExcelLoader] 错误: {e}")
            traceback.print_exc()
            # 返回固定10个输出，第一个包含错误信息，其余为空
            error_result = [f"错误: {e}"] + [""] * 9
            return {"result": tuple(error_result)}

    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"


# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "DynamicExcelLoader": DynamicExcelLoader
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "DynamicExcelLoader": "Dynamic Excel Loader"
}

print("--- 加载自定义节点: DynamicExcelLoader (ExcelUtils/Dynamic Excel Loader) v1.0 [Dynamic Column Detection & Multi-Type Output] ---")
