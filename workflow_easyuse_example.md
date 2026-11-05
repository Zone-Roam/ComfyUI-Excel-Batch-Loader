# EasyUse For循环 + Excel加载器工作流示例

## 🎯 工作流说明

这个工作流展示了如何使用EasyUse的for循环节点配合Excel EasyUse For Loop节点来循环读取Excel数据。

## 📋 节点连接顺序

```
1. [easy forLoopStart] 
   ↓ (index输出)
2. [Excel EasyUse For Loop]
   ↓ (column_1, column_2, ... 输出)
3. [ShowText] 或其他处理节点
   ↓
4. [easy forLoopEnd]
```

## ⚙️ 节点配置

### 1. easy forLoopStart 节点配置
- **start_value**: `2` (从第2行开始，跳过表头)
- **end_value**: `6` (到第6行结束，实际读取2-5行)
- **step**: `1` (每次递增1)

### 2. Excel EasyUse For Loop 节点配置
- **excel_file_path**: `"example.xlsx"` (你的Excel文件路径)
- **index**: 连接到 forLoopStart 的 index 输出
- **refresh_trigger**: `False`
- **auto_refresh**: `True`
- **sheet_name**: `"0"` (第一个工作表)

### 3. easy forLoopEnd 节点配置
- **flow**: 连接到 forLoopStart 的 flow 输出
- 其他输入根据需要连接处理后的数据

## 🔄 执行流程

1. **forLoopStart** 开始循环，输出 index=2
2. **Excel EasyUse For Loop** 接收 index=2，读取Excel第2行数据
3. 处理节点（如ShowText）显示或处理数据
4. **forLoopEnd** 完成当前循环，index递增到3
5. 重复步骤2-4，直到index=6时循环结束

## 📊 数据流示例

假设Excel文件内容：
```
行1(表头): 名称    | 描述        | 价格
行2:      苹果    | 红色的水果   | 5.0
行3:      香蕉    | 黄色的水果   | 3.5
行4:      橙子    | 橙色的水果   | 4.2
行5:      葡萄    | 紫色的水果   | 8.0
```

循环执行结果：
- **第1次循环 (index=2)**: 读取"苹果, 红色的水果, 5.0"
- **第2次循环 (index=3)**: 读取"香蕉, 黄色的水果, 3.5"
- **第3次循环 (index=4)**: 读取"橙子, 橙色的水果, 4.2"
- **第4次循环 (index=5)**: 读取"葡萄, 紫色的水果, 8.0"
- **循环结束**

## 💡 使用技巧

1. **设置正确的范围**: end_value应该是最后一行+1，因为forLoopStart使用的是 `<` 比较
2. **处理空行**: 节点会自动处理空单元格，返回空字符串
3. **错误处理**: 如果index超出数据范围，会返回错误信息而不是崩溃
4. **性能优化**: 启用auto_refresh可以自动检测Excel文件变化

## 🚨 注意事项

- 确保已安装EasyUse节点包
- Excel文件路径要正确
- forLoopStart的start_value通常设为2（跳过表头）
- 如果Excel有多个工作表，可以通过sheet_name参数指定

## 🎉 优势

- **无缝集成**: 完美适配EasyUse的for循环机制
- **自动索引**: 无需手动计算行号，直接使用forLoopStart的index输出
- **错误处理**: 优雅处理超出范围和文件错误
- **高性能**: 智能缓存机制，避免重复解析Excel文件
