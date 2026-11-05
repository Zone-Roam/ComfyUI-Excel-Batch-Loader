# Excel行数统计节点
import os
import traceback

# 尝试导入pandas
try:
    import pandas as pd
    print("✅ [DEBUG] pandas 导入成功 (ExcelRowCounter)")
except ImportError as e:
    print(f"❌ [ERROR] ExcelRowCounter 节点缺少 Pandas 库: {e}")
    print("请在ComfyUI根目录运行: python_embeded\\python.exe -m pip install pandas openpyxl")
    pd = None

# 尝试导入openpyxl
try:
    import openpyxl
    print("✅ [DEBUG] openpyxl 导入成功 (ExcelRowCounter)")
except ImportError as e:
    print(f"❌ [ERROR] openpyxl 导入失败 (ExcelRowCounter): {e}")

class ExcelRowCounter:
    """
    Excel行数统计节点
    统计Excel表格中有内容的行数（不包括表头）
    """
    
    def __init__(self):
        self._cached_file_path = None
        self._cached_file_mtime = None
        self._cached_row_count = 0
        self._cached_include_header = None
        self._cached_sheet_name = None
    
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
                "include_header": ("BOOLEAN", {
                    "default": False,
                    "description": "是否包含表头行在计数中"
                }),
                "refresh_trigger": ("BOOLEAN", {
                    "default": False,
                    "description": "手动刷新开关（切换此开关可强制重新统计）"
                }),
                "auto_refresh": ("BOOLEAN", {
                    "default": True,
                    "description": "自动刷新（当Excel文件修改时间改变时自动重新统计）"
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
    
    RETURN_TYPES = ("INT", "STRING")
    RETURN_NAMES = ("row_count", "status_message")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"
    
    def _should_refresh_cache(self, excel_file_path: str, include_header: bool, sheet_name: str, refresh_trigger: bool, auto_refresh: bool) -> bool:
        """
        判断是否需要刷新缓存
        """
        # 手动刷新触发
        if refresh_trigger:
            return True

        # 文件路径改变
        if self._cached_file_path != excel_file_path:
            return True

        # 参数改变
        if self._cached_include_header != include_header:
            return True

        if self._cached_sheet_name != sheet_name:
            return True

        # 自动刷新：文件修改时间改变
        if auto_refresh and os.path.exists(excel_file_path):
            current_mtime = os.path.getmtime(excel_file_path)
            if self._cached_file_mtime != current_mtime:
                return True

        return False
    
    def _count_excel_rows(self, excel_file_path: str, sheet_name: str, include_header: bool) -> tuple:
        """
        统计Excel文件的行数
        返回 (行数, 状态信息)
        """
        if pd is None:
            return 0, "错误：Pandas 库未安装"
        
        # 检查文件是否存在
        if not os.path.exists(excel_file_path):
            return 0, f"错误：文件未找到 - {excel_file_path}"
        
        try:
            # 确定工作表标识符
            sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
            
            # 读取Excel文件
            df = pd.read_excel(
                excel_file_path,
                header=0,  # 第一行作为表头
                engine='openpyxl',
                sheet_name=sheet_identifier
            )
            
            # 统计有效行数
            # 去除所有列都为空的行
            non_empty_rows = df.dropna(how='all')
            data_row_count = len(non_empty_rows)
            
            # 根据是否包含表头决定最终计数
            if include_header:
                total_count = data_row_count + 1  # +1 for header
                status = f"成功：{data_row_count} 行数据 + 1 行表头 = {total_count} 行"
            else:
                total_count = data_row_count
                status = f"成功：{data_row_count} 行数据（不包含表头）"
            
            return total_count, status
            
        except FileNotFoundError:
            return 0, f"错误：文件未找到 - {excel_file_path}"
        except ValueError as e:
            return 0, f"错误：工作表 '{sheet_name}' 不存在 - {e}"
        except Exception as e:
            return 0, f"错误：读取文件失败 - {e}"
    
    def _update_cache(self, excel_file_path: str, sheet_name: str, include_header: bool):
        """
        更新缓存
        """
        row_count, status = self._count_excel_rows(excel_file_path, sheet_name, include_header)

        self._cached_file_path = excel_file_path
        self._cached_file_mtime = os.path.getmtime(excel_file_path) if os.path.exists(excel_file_path) else 0
        self._cached_row_count = row_count
        self._cached_include_header = include_header
        self._cached_sheet_name = sheet_name

        return row_count, status
    
    def execute(self, excel_file_path, include_header, refresh_trigger, auto_refresh, sheet_name="0", **kwargs):
        """
        执行节点逻辑
        """
        if pd is None:
            return (0, "错误：运行此节点需要 Pandas 库，请先安装")
        
        try:
            # 检查是否需要刷新缓存
            if self._should_refresh_cache(excel_file_path, include_header, sheet_name, refresh_trigger, auto_refresh):
                row_count, status = self._update_cache(excel_file_path, sheet_name, include_header)
                print(f"[ExcelRowCounter] 缓存已刷新：{status}")
            else:
                # 使用缓存的结果
                row_count = self._cached_row_count
                status = f"使用缓存：{row_count} 行"

            return (row_count, status)
            
        except Exception as e:
            print(f"[ExcelRowCounter] 错误: {e}")
            traceback.print_exc()
            return (0, f"错误: {e}")

# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "ExcelRowCounter": ExcelRowCounter
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "ExcelRowCounter": "Excel Row Counter"
}

print("✅ [DEBUG] ExcelRowCounter 注册成功")
