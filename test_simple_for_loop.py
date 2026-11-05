#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Simple For Loop节点
验证新的for循环实现是否正常工作
"""

import os
import sys

def test_simple_for_loop():
    """测试Simple For Loop节点的基本功能"""
    
    print("🧪 开始测试 Simple For Loop 节点...")
    
    # 导入节点
    try:
        from simple_for_loop import SimpleForLoopStart, SimpleForLoopEnd, SimpleExcelForLoop
        print("✅ 节点导入成功")
    except ImportError as e:
        print(f"❌ 节点导入失败: {e}")
        return False
    
    # 创建节点实例
    loop_start = SimpleForLoopStart()
    loop_end = SimpleForLoopEnd()
    excel_loop = SimpleExcelForLoop()
    
    print("\n🔧 测试 SimpleForLoopStart...")
    
    # 测试循环开始节点
    print("测试配置: loop_count=3, start_index=1")
    
    # 模拟3次循环
    for i in range(5):  # 多测试几次，看看是否正确停止
        current_index, is_finished, loop_info = loop_start.execute(
            loop_count=3,
            start_index=1,
            reset_trigger=False
        )
        
        print(f"  第{i+1}次调用: index={current_index}, finished={is_finished}, info='{loop_info}'")
        
        # 测试循环结束节点
        final_status, completed, collected_data = loop_end.execute(
            is_finished=is_finished,
            loop_info=loop_info,
            loop_data=f"数据_{current_index}"
        )
        
        print(f"    ForLoopEnd: status='{final_status}', completed={completed}")
        
        if is_finished:
            print("  ✅ 循环正确结束")
            break
    
    print("\n🔧 测试重置功能...")
    
    # 测试重置
    current_index, is_finished, loop_info = loop_start.execute(
        loop_count=2,
        start_index=5,
        reset_trigger=True  # 重置
    )
    print(f"  重置后第1次: index={current_index}, finished={is_finished}, info='{loop_info}'")
    
    current_index, is_finished, loop_info = loop_start.execute(
        loop_count=2,
        start_index=5,
        reset_trigger=False
    )
    print(f"  重置后第2次: index={current_index}, finished={is_finished}, info='{loop_info}'")
    
    current_index, is_finished, loop_info = loop_start.execute(
        loop_count=2,
        start_index=5,
        reset_trigger=False
    )
    print(f"  重置后第3次: index={current_index}, finished={is_finished}, info='{loop_info}'")
    
    print("\n🔧 测试 SimpleExcelForLoop...")
    
    # 检查是否有测试Excel文件
    test_excel_files = [
        "example.xlsx",
        "test.xlsx", 
        "C:\\Users\\admin\\Desktop\\example (3).xlsx"  # 从工作流中看到的路径
    ]
    
    test_file = None
    for file_path in test_excel_files:
        if os.path.exists(file_path):
            test_file = file_path
            break
    
    if test_file:
        print(f"  使用测试文件: {test_file}")
        
        # 测试读取不同索引的数据
        for index in [1, 2, 3]:
            print(f"\n  测试索引 {index}:")
            try:
                result = excel_loop.execute(
                    excel_file_path=test_file,
                    current_index=index,
                    skip_header=True,
                    sheet_name="0"
                )
                
                print(f"    结果: {result[:3]}...{result[-1]}")  # 显示前3列和行信息
                
            except Exception as e:
                print(f"    ❌ 错误: {e}")
    else:
        print("  ⚠️ 未找到测试Excel文件，跳过Excel读取测试")
        print(f"  尝试查找的文件: {test_excel_files}")
    
    print("\n🎯 完整工作流模拟...")
    
    # 模拟完整的工作流
    if test_file:
        print("模拟读取2行Excel数据的完整流程:")
        
        # 重置循环
        loop_start_new = SimpleForLoopStart()
        
        for step in range(4):  # 多测试一次，确保正确停止
            print(f"\n--- 步骤 {step + 1} ---")
            
            # 1. 获取循环状态
            current_index, is_finished, loop_info = loop_start_new.execute(
                loop_count=2,
                start_index=1,
                reset_trigger=(step == 0)  # 第一次重置
            )
            
            print(f"ForLoopStart: {loop_info}")
            
            if not is_finished:
                # 2. 读取Excel数据
                try:
                    excel_result = excel_loop.execute(
                        excel_file_path=test_file,
                        current_index=current_index,
                        skip_header=True,
                        sheet_name="0"
                    )
                    print(f"Excel数据: 第1列='{excel_result[0]}', 行信息='{excel_result[-1]}'")
                    
                    # 3. 处理循环结束
                    final_status, completed, collected_data = loop_end.execute(
                        is_finished=is_finished,
                        loop_info=loop_info,
                        loop_data=excel_result[0]  # 传递第一列数据
                    )
                    
                except Exception as e:
                    print(f"Excel读取错误: {e}")
            else:
                # 循环已结束
                final_status, completed, collected_data = loop_end.execute(
                    is_finished=is_finished,
                    loop_info=loop_info
                )
                print(f"循环结束: {final_status}")
                break
    
    print("\n✅ Simple For Loop 节点测试完成！")
    return True

if __name__ == "__main__":
    test_simple_for_loop()
