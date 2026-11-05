#!/usr/bin/env python3
"""
测试EasyUse For循环适配Excel加载器
模拟EasyUse forLoopStart的index输出
"""

import os
import sys
from excel_easyuse_for_loop import ExcelEasyUseForLoop

def test_easyuse_for_loop():
    """测试EasyUse For循环适配功能"""
    
    print("🔍 测试Excel EasyUse For循环适配加载器")
    print("=" * 60)
    
    # 测试文件
    excel_file_path = "example.xlsx"
    
    if not os.path.exists(excel_file_path):
        print(f"❌ 错误：找不到测试文件 {excel_file_path}")
        return False
    
    loader = ExcelEasyUseForLoop()
    
    print("🧪 测试1：模拟EasyUse forLoopStart的index输出")
    print("-" * 50)
    
    # 模拟forLoopStart输出的index值
    # 假设forLoopStart设置为: start_value=2, end_value=7, step=1
    # 这将产生index值: 2, 3, 4, 5, 6
    test_indices = [2, 3, 4, 5, 6, 7]  # 包含一个超出范围的值来测试边界
    
    for i, index in enumerate(test_indices):
        print(f"\n📋 第{i+1}次循环 (index={index}):")
        
        result = loader.execute(
            excel_file_path=excel_file_path,
            index=index,
            refresh_trigger=False,
            auto_refresh=True,
            sheet_name="0"
        )
        
        # 解析结果
        columns = result[:9]
        row_info = result[9]
        
        print(f"  行信息: {row_info}")
        
        # 显示有数据的列
        has_data = False
        for j, col_data in enumerate(columns):
            if col_data:
                print(f"  列{j+1}: {col_data}")
                has_data = True
        
        if not has_data:
            print("  (此行无数据或超出范围)")
    
    print("\n" + "=" * 60)
    print("🧪 测试2：测试缓存刷新功能")
    print("-" * 50)
    
    # 测试手动刷新
    result = loader.execute(
        excel_file_path=excel_file_path,
        index=2,
        refresh_trigger=True,  # 手动刷新
        auto_refresh=True,
        sheet_name="0"
    )
    
    columns = result[:9]
    row_info = result[9]
    
    print(f"手动刷新后 - {row_info}")
    for j, col_data in enumerate(columns):
        if col_data:
            print(f"  列{j+1}: {col_data}")
    
    print("\n" + "=" * 60)
    print("🧪 测试3：测试错误处理")
    print("-" * 50)
    
    # 测试不存在的文件
    result = loader.execute(
        excel_file_path="不存在的文件.xlsx",
        index=2,
        refresh_trigger=False,
        auto_refresh=True,
        sheet_name="0"
    )
    
    print(f"不存在文件的错误处理: {result[0]}")
    
    # 测试超出范围的索引
    result = loader.execute(
        excel_file_path=excel_file_path,
        index=100,  # 超出范围的索引
        refresh_trigger=False,
        auto_refresh=True,
        sheet_name="0"
    )
    
    print(f"超出范围索引的处理: {result[9]}")
    
    print("\n" + "=" * 60)
    print("🎯 EasyUse For循环使用说明:")
    print("-" * 50)
    print("1. 在ComfyUI中添加 'easy forLoopStart' 节点")
    print("2. 设置 start_value=2, end_value=6, step=1 (读取第2-5行)")
    print("3. 将 forLoopStart 的 'index' 输出连接到此节点的 'index' 输入")
    print("4. 在循环体内使用此节点读取Excel数据")
    print("5. 最后连接到 'easy forLoopEnd' 节点完成循环")
    print("\n✅ 测试完成！")
    
    return True

if __name__ == "__main__":
    test_easyuse_for_loop()
