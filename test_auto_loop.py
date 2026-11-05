#!/usr/bin/env python3
"""
测试自动循环Excel加载器
"""

import os
import sys
from excel_auto_loop_loader import ExcelAutoLoopLoader

def test_auto_loop():
    """测试自动循环功能"""
    
    print("🔍 测试Excel自动循环加载器")
    print("=" * 50)
    
    # 测试文件
    excel_file_path = "example.xlsx"
    
    if not os.path.exists(excel_file_path):
        print(f"❌ 错误：找不到测试文件 {excel_file_path}")
        return False
    
    loader = ExcelAutoLoopLoader()
    
    print("🧪 测试1：单行循环模式")
    print("-" * 30)
    
    # 模拟多次执行，看看循环效果
    for execution in range(8):  # 执行8次，看看循环效果
        print(f"\n📋 第{execution + 1}次执行:")
        
        result = loader.execute(
            excel_file_path=excel_file_path,
            loop_mode="single_row",
            reset_loop=False,
            refresh_trigger=False,
            auto_refresh=True,
            sheet_name="0"
        )
        
        # 解析结果
        columns = result[:9]
        current_row = result[9]
        total_rows = result[10]
        loop_status = result[11]
        
        print(f"  当前行: {current_row}")
        print(f"  总行数: {total_rows}")
        print(f"  循环状态: {loop_status}")
        
        # 显示有数据的列
        for i, col_data in enumerate(columns):
            if col_data:
                print(f"  列{i+1}: {col_data}")
    
    print("\n" + "=" * 50)
    print("🧪 测试2：重置循环")
    print("-" * 30)
    
    # 重置循环
    result = loader.execute(
        excel_file_path=excel_file_path,
        loop_mode="single_row",
        reset_loop=True,  # 重置
        refresh_trigger=False,
        auto_refresh=True,
        sheet_name="0"
    )
    
    columns = result[:9]
    current_row = result[9]
    total_rows = result[10]
    loop_status = result[11]
    
    print(f"重置后 - 当前行: {current_row}, 状态: {loop_status}")
    for i, col_data in enumerate(columns):
        if col_data:
            print(f"  列{i+1}: {col_data}")
    
    print("\n" + "=" * 50)
    print("🧪 测试3：一次性读取所有行")
    print("-" * 30)
    
    # 创建新实例测试all_rows模式
    loader2 = ExcelAutoLoopLoader()
    result = loader2.execute(
        excel_file_path=excel_file_path,
        loop_mode="all_rows",
        reset_loop=False,
        refresh_trigger=False,
        auto_refresh=True,
        sheet_name="0"
    )
    
    columns = result[:9]
    current_row = result[9]
    total_rows = result[10]
    loop_status = result[11]
    
    print(f"一次性读取 - 行范围: {current_row}, 总行数: {total_rows}, 状态: {loop_status}")
    for i, col_data in enumerate(columns):
        if col_data:
            print(f"  列{i+1}:")
            # 按行分割显示
            lines = col_data.split('\n')
            for j, line in enumerate(lines):
                print(f"    行{j+1}: {line}")
    
    return True

if __name__ == "__main__":
    test_auto_loop()
