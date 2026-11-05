#!/usr/bin/env python3
"""
测试DynamicExcelLoader节点的脚本
"""

import os
import sys
from dynamic_excel_loader_node import DynamicExcelLoader
from excel_row_counter import ExcelRowCounter

def test_dynamic_excel_loader():
    """测试DynamicExcelLoader节点"""
    
    print("=== DynamicExcelLoader 节点测试 ===\n")
    
    # 创建节点实例
    loader = DynamicExcelLoader()
    
    # 测试参数
    excel_file_path = "example.xlsx"
    start_row_number = 2
    read_count = 1

    refresh_trigger = False
    auto_refresh = True
    sheet_name = "0"
    
    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        print(f"❌ 错误：找不到测试文件 {excel_file_path}")
        print("请先运行 'python create_example_excel.py' 创建示例文件")
        return False
    
    print(f"📁 测试文件: {excel_file_path}")
    print(f"📊 开始行号: {start_row_number}")
    print(f"📖 读取行数: {read_count}")
    print()
    
    try:
        # 执行节点
        result = loader.execute(
            excel_file_path=excel_file_path,
            start_row_number=start_row_number,
            read_count=read_count,
            refresh_trigger=refresh_trigger,
            auto_refresh=auto_refresh,
            sheet_name=sheet_name
        )
        
        print("✅ 节点执行成功！")
        print()
        
        # 显示结果
        if "result" in result:
            outputs = result["result"]
            print("📤 输出结果:")
            for i, output in enumerate(outputs):
                if i < 9:  # 前9个是列数据
                    column_name = f"column_{i+1}"
                    if output:
                        print(f"  {column_name}: {output} (文本)")
                    else:
                        print(f"  {column_name}: (空)")
                else:  # 第10个是当前行号
                    print(f"  current_row: {output}")
        
        # 显示UI信息
        if "ui" in result:
            ui_info = result["ui"]
            print()
            print("🔧 UI信息:")
            if "start_row_number" in ui_info:
                print(f"  下次开始行号: {ui_info['start_row_number'][0]}")
            if "column_info" in ui_info:
                print("  列信息:")
                for info in ui_info["column_info"]:
                    print(f"    {info}")
        
        print()
        print("🎯 测试第二行数据...")
        
        # 测试读取第二行
        result2 = loader.execute(
            excel_file_path=excel_file_path,
            start_row_number=3,  # 第三行数据
            read_count=read_count,
            refresh_trigger=refresh_trigger,
            auto_refresh=auto_refresh,
            sheet_name=sheet_name
        )
        
        if "result" in result2:
            outputs2 = result2["result"]
            print("📤 第二行输出结果:")
            for i, output in enumerate(outputs2):
                if i < 9 and output:
                    column_name = f"column_{i+1}"
                    print(f"  {column_name}: {output} (文本)")
        
        print()
        print("✅ 动态Excel加载器测试通过！")
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_excel_row_counter():
    """测试Excel行数统计节点"""

    print("\n=== Excel行数统计测试 ===\n")

    # 创建节点实例
    counter = ExcelRowCounter()

    # 测试参数
    excel_file_path = "example.xlsx"
    include_header = False
    refresh_trigger = False
    auto_refresh = True
    sheet_name = "0"

    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        print(f"❌ 错误：找不到测试文件 {excel_file_path}")
        return False

    print(f"📁 测试文件: {excel_file_path}")
    print(f"📊 包含表头: {include_header}")
    print()

    try:
        # 测试不包含表头的计数
        print("🧪 测试1：不包含表头的行数统计")
        result1 = counter.execute(
            excel_file_path=excel_file_path,
            include_header=False,
            refresh_trigger=refresh_trigger,
            auto_refresh=auto_refresh,
            sheet_name=sheet_name
        )

        row_count1, status1 = result1
        print(f"  📊 行数: {row_count1}")
        print(f"  📝 状态: {status1}")
        print()

        # 测试包含表头的计数
        print("🧪 测试2：包含表头的行数统计")
        result2 = counter.execute(
            excel_file_path=excel_file_path,
            include_header=True,
            refresh_trigger=refresh_trigger,
            auto_refresh=auto_refresh,
            sheet_name=sheet_name
        )

        row_count2, status2 = result2
        print(f"  📊 行数: {row_count2}")
        print(f"  📝 状态: {status2}")
        print()

        # 验证逻辑
        if row_count2 == row_count1 + 1:
            print("✅ 包含表头的计数比不包含表头多1行，逻辑正确")
        else:
            print("❌ 计数逻辑可能有问题")
            return False

        # 测试手动刷新
        print("🧪 测试3：手动刷新功能")
        result3 = counter.execute(
            excel_file_path=excel_file_path,
            include_header=False,
            refresh_trigger=True,  # 手动刷新
            auto_refresh=auto_refresh,
            sheet_name=sheet_name
        )

        row_count3, status3 = result3
        print(f"  📊 行数: {row_count3}")
        print(f"  📝 状态: {status3}")

        if row_count3 == row_count1:
            print("✅ 手动刷新功能正常")
        else:
            print("❌ 手动刷新功能可能有问题")
            return False

        print()
        print("✅ Excel行数统计器测试通过！")
        return True

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🚀 开始测试 Excel 节点...\n")

    # 测试动态Excel加载器
    success1 = test_dynamic_excel_loader()

    # 测试Excel行数统计器
    success2 = test_excel_row_counter()

    if success1 and success2:
        print("\n🎉 所有测试完成！所有节点工作正常。")
    else:
        print("\n💥 部分测试失败，请检查错误信息。")
        if not success1:
            print("  - 动态Excel加载器测试失败")
        if not success2:
            print("  - Excel行数统计器测试失败")
        sys.exit(1)
