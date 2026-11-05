#!/usr/bin/env python3
"""
ComfyUI Dynamic Excel Loader 安装脚本
"""

import os
import sys
import subprocess
import shutil

def find_comfyui_python():
    """查找ComfyUI的Python路径"""
    possible_paths = [
        # 当前目录的上级目录（如果在custom_nodes中）
        "../../python_embeded/python.exe",
        "../../../python_embeded/python.exe",
        # 常见的ComfyUI安装路径
        "python_embeded/python.exe",
        "../python_embeded/python.exe",
        # 如果在ComfyUI根目录
        "./python_embeded/python.exe",
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return os.path.abspath(path)

    return None

def install_dependencies():
    """安装依赖包"""
    print("📦 安装依赖包...")

    # 首先尝试找到ComfyUI的Python
    comfyui_python = find_comfyui_python()

    if comfyui_python:
        print(f"🔍 找到ComfyUI Python: {comfyui_python}")
        try:
            subprocess.check_call([comfyui_python, "-m", "pip", "install", "pandas", "openpyxl"])
            print("✅ 依赖包安装成功！")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ ComfyUI Python安装失败: {e}")
            print("🔄 尝试使用系统Python...")

    # 如果找不到ComfyUI Python，使用系统Python
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
        print("✅ 依赖包安装成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 依赖包安装失败: {e}")
        print("\n💡 手动安装提示:")
        print("请在ComfyUI根目录运行以下命令:")
        print("python_embeded\\python.exe -m pip install pandas openpyxl")
        return False

def test_installation():
    """测试安装"""
    print("🧪 测试安装...")
    try:
        from dynamic_excel_loader_node import DynamicExcelLoader
        loader = DynamicExcelLoader()
        print("✅ 节点导入成功！")
        return True
    except ImportError as e:
        print(f"❌ 节点导入失败: {e}")
        return False

def create_example():
    """创建示例文件"""
    print("📄 创建示例Excel文件...")
    try:
        import subprocess
        result = subprocess.run([sys.executable, "create_example_excel.py"], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print("✅ 示例文件创建成功！")
            return True
        else:
            print(f"❌ 示例文件创建失败: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ 示例文件创建失败: {e}")
        return False

def main():
    """主安装流程"""
    print("🚀 ComfyUI Dynamic Excel Loader 安装程序")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("❌ 需要Python 3.7或更高版本")
        sys.exit(1)
    
    print(f"✅ Python版本: {sys.version}")
    
    # 安装依赖
    if not install_dependencies():
        print("💥 安装失败！")
        sys.exit(1)
    
    # 测试安装
    if not test_installation():
        print("💥 测试失败！")
        sys.exit(1)
    
    # 创建示例
    create_example()
    
    print()
    print("🎉 安装完成！")
    print()
    print("📋 下一步:")
    print("1. 将此文件夹复制到ComfyUI的custom_nodes目录")
    print("2. 重启ComfyUI")
    print("3. 在节点菜单中找到 'Dynamic Excel Loader'")
    print("4. 使用example.xlsx作为测试文件")
    print()
    print("📖 详细使用说明请查看README.md")

if __name__ == "__main__":
    main()
