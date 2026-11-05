#!/usr/bin/env python3
"""
创建示例Excel文件的脚本
运行此脚本将生成一个示例Excel文件，用于测试DynamicExcelLoader节点
"""

import pandas as pd
import os

def create_example_excel():
    """创建示例Excel文件"""
    
    # 示例数据
    data = {
        '名称': ['苹果', '香蕉', '橙子', '葡萄', '草莓'],
        '描述': ['红色的水果', '黄色的水果', '橙色的水果', '紫色的水果', '红色的小水果'],
        '图像路径': [
            '/images/apple.jpg',
            '/images/banana.jpg', 
            '/images/orange.jpg',
            '/images/grape.jpg',
            '/images/strawberry.jpg'
        ],
        '标签': [
            '水果,红色,甜',
            '水果,黄色,甜',
            '水果,橙色,酸甜',
            '水果,紫色,甜',
            '水果,红色,酸甜'
        ],
        '图片文件': [
            '/photos/apple.png',
            '/photos/banana.png',
            '/photos/orange.png', 
            '/photos/grape.png',
            '/photos/strawberry.png'
        ],
        '价格': ['5.00', '3.50', '4.20', '8.00', '12.00'],
        '产地': ['山东', '海南', '江西', '新疆', '辽宁']
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    output_file = 'example.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"示例Excel文件已创建: {output_file}")
    print("\n文件结构:")
    print("- 列1: 名称 (文本)")
    print("- 列2: 描述 (文本)")  
    print("- 列3: 图像路径 (图像路径) - 包含'图像路径'关键词")
    print("- 列4: 标签 (文本)")
    print("- 列5: 图片文件 (图像路径) - 包含'图片'关键词")
    print("- 列6: 价格 (文本)")
    print("- 列7: 产地 (文本)")
    print(f"\n总共 {len(df)} 行数据（不包括表头）")
    
    return output_file

if __name__ == "__main__":
    create_example_excel()
