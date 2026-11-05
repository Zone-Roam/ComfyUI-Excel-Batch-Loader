# 简化版Excel加载器节点 - 用于测试ComfyUI注册
import os

class SimpleExcelLoader:
    """
    简化版Excel加载器 - 测试节点注册
    """
    
    def __init__(self):
        pass
    
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_file_path": ("STRING", {
                    "multiline": False, 
                    "default": "example.xlsx", 
                    "description": "Excel文件路径"
                }),
                "test_mode": ("BOOLEAN", {
                    "default": True,
                    "description": "测试模式"
                })
            }
        }
    
    RETURN_TYPES = ("STRING", "STRING")
    RETURN_NAMES = ("status", "message")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"
    
    def execute(self, excel_file_path, test_mode):
        """简单的执行函数"""
        try:
            if test_mode:
                return ("SUCCESS", f"测试模式：文件路径 = {excel_file_path}")
            else:
                if os.path.exists(excel_file_path):
                    return ("SUCCESS", f"文件存在：{excel_file_path}")
                else:
                    return ("ERROR", f"文件不存在：{excel_file_path}")
        except Exception as e:
            return ("ERROR", f"错误：{e}")

# 节点注册
NODE_CLASS_MAPPINGS = {
    "SimpleExcelLoader": SimpleExcelLoader
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "SimpleExcelLoader": "Simple Excel Loader (Test)"
}

print("✅ [DEBUG] SimpleExcelLoader 注册成功")
