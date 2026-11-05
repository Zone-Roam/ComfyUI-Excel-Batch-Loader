# __init__.py
# This file makes the directory a Python package.
# ComfyUI looks for NODE_CLASS_MAPPINGS and NODE_DISPLAY_NAME_MAPPINGS here.

print("🔍 [DEBUG] 开始加载 ComfyUI_DynamicExcelLoader...")

# 初始化映射字典
NODE_CLASS_MAPPINGS = {}
NODE_DISPLAY_NAME_MAPPINGS = {}

# 首先加载简单测试节点（不依赖pandas）
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .simple_excel_loader import NODE_CLASS_MAPPINGS as simple_mappings, NODE_DISPLAY_NAME_MAPPINGS as simple_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import simple_excel_loader
        simple_mappings = simple_excel_loader.NODE_CLASS_MAPPINGS
        simple_display_mappings = simple_excel_loader.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(simple_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(simple_display_mappings)
    print("✅ [DEBUG] 简单Excel加载器注册成功")
except Exception as e:
    print(f"❌ [ERROR] 简单Excel加载器注册失败: {e}")

# 然后尝试加载完整功能节点（依赖pandas）
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .dynamic_excel_loader_node import NODE_CLASS_MAPPINGS as dynamic_mappings, NODE_DISPLAY_NAME_MAPPINGS as dynamic_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import dynamic_excel_loader_node
        dynamic_mappings = dynamic_excel_loader_node.NODE_CLASS_MAPPINGS
        dynamic_display_mappings = dynamic_excel_loader_node.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(dynamic_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(dynamic_display_mappings)
    print("✅ [DEBUG] 动态Excel加载器注册成功")

except ImportError as e:
    print(f"❌ [ERROR] 动态Excel加载器导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")
    print("💡 [HINT] 运行: python_embeded\\python.exe -m pip install pandas openpyxl")

except Exception as e:
    print(f"❌ [ERROR] 动态Excel加载器未知错误: {e}")
    import traceback
    traceback.print_exc()

# 加载Excel行数统计节点
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .excel_row_counter import NODE_CLASS_MAPPINGS as counter_mappings, NODE_DISPLAY_NAME_MAPPINGS as counter_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import excel_row_counter
        counter_mappings = excel_row_counter.NODE_CLASS_MAPPINGS
        counter_display_mappings = excel_row_counter.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(counter_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(counter_display_mappings)
    print("✅ [DEBUG] Excel行数统计器注册成功")

except ImportError as e:
    print(f"❌ [ERROR] Excel行数统计器导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")

except Exception as e:
    print(f"❌ [ERROR] Excel行数统计器未知错误: {e}")
    import traceback
    traceback.print_exc()

# 加载For循环专用Excel加载器
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .excel_loader_for_loop import NODE_CLASS_MAPPINGS as loop_mappings, NODE_DISPLAY_NAME_MAPPINGS as loop_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import excel_loader_for_loop
        loop_mappings = excel_loader_for_loop.NODE_CLASS_MAPPINGS
        loop_display_mappings = excel_loader_for_loop.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(loop_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(loop_display_mappings)
    print("✅ [DEBUG] For循环Excel加载器注册成功")

except ImportError as e:
    print(f"❌ [ERROR] For循环Excel加载器导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")

except Exception as e:
    print(f"❌ [ERROR] For循环Excel加载器未知错误: {e}")
    import traceback
    traceback.print_exc()

# 加载自动循环Excel加载器
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .excel_auto_loop_loader import NODE_CLASS_MAPPINGS as auto_loop_mappings, NODE_DISPLAY_NAME_MAPPINGS as auto_loop_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import excel_auto_loop_loader
        auto_loop_mappings = excel_auto_loop_loader.NODE_CLASS_MAPPINGS
        auto_loop_display_mappings = excel_auto_loop_loader.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(auto_loop_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(auto_loop_display_mappings)
    print("✅ [DEBUG] 自动循环Excel加载器注册成功")

except ImportError as e:
    print(f"❌ [ERROR] 自动循环Excel加载器导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")

except Exception as e:
    print(f"❌ [ERROR] 自动循环Excel加载器未知错误: {e}")
    import traceback
    traceback.print_exc()

# 加载EasyUse For循环适配Excel加载器
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .excel_easyuse_for_loop import NODE_CLASS_MAPPINGS as easyuse_loop_mappings, NODE_DISPLAY_NAME_MAPPINGS as easyuse_loop_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import excel_easyuse_for_loop
        easyuse_loop_mappings = excel_easyuse_for_loop.NODE_CLASS_MAPPINGS
        easyuse_loop_display_mappings = excel_easyuse_for_loop.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(easyuse_loop_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(easyuse_loop_display_mappings)
    print("✅ [DEBUG] EasyUse For循环适配Excel加载器注册成功")

except ImportError as e:
    print(f"❌ [ERROR] EasyUse For循环适配Excel加载器导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")

except Exception as e:
    print(f"❌ [ERROR] EasyUse For循环适配Excel加载器未知错误: {e}")
    import traceback
    traceback.print_exc()

# 加载Simple For循环节点（新的可靠for循环实现）
try:
    try:
        # 尝试相对导入（在ComfyUI中）
        from .simple_for_loop import NODE_CLASS_MAPPINGS as simple_loop_mappings, NODE_DISPLAY_NAME_MAPPINGS as simple_loop_display_mappings
    except ImportError:
        # 如果相对导入失败，尝试直接导入（在测试中）
        import simple_for_loop
        simple_loop_mappings = simple_for_loop.NODE_CLASS_MAPPINGS
        simple_loop_display_mappings = simple_for_loop.NODE_DISPLAY_NAME_MAPPINGS

    NODE_CLASS_MAPPINGS.update(simple_loop_mappings)
    NODE_DISPLAY_NAME_MAPPINGS.update(simple_loop_display_mappings)
    print("✅ [DEBUG] Simple For循环节点注册成功")

except ImportError as e:
    print(f"❌ [ERROR] Simple For循环节点导入失败: {e}")
    print("💡 [HINT] 请检查是否安装了pandas和openpyxl")

except Exception as e:
    print(f"❌ [ERROR] Simple For循环节点未知错误: {e}")
    import traceback
    traceback.print_exc()

print(f"📋 [DEBUG] 总共注册了 {len(NODE_CLASS_MAPPINGS)} 个节点")

# Export them so ComfyUI can find them
__all__ = ['NODE_CLASS_MAPPINGS', 'NODE_DISPLAY_NAME_MAPPINGS']

print(f"🎉 [DEBUG] 最终注册的节点: {list(NODE_CLASS_MAPPINGS.keys())}")
print("🎉 [DEBUG] ComfyUI_DynamicExcelLoader 加载完成！")
