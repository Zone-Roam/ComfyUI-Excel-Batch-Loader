# ComfyUI Dynamic Excel Loader

**English summary.** A ComfyUI node pack that drives batch generation from a spreadsheet. The loader reads an Excel file, detects the header columns and exposes one output per column (up to 9), so each row becomes one generation run (a prompt, an image path, a filename and so on). It includes a row counter and for-loop helpers to iterate over every row, and refreshes automatically when the file changes. Built in 2025 for batch image and video production. Detailed docs below are in Chinese.

---

🚀 这是一个为ComfyUI设计的动态Excel加载器节点，能够根据Excel表头自动识别列类型，动态生成对应数量的输出组件，支持文本和图像路径两种数据类型。

## ✨ 主要特性

- **📊 动态输出生成**：根据Excel列数自动生成对应数量的输出
- **🔄 智能刷新机制**：支持手动和自动刷新，文件修改时自动重新解析
- **📋 多列同时处理**：最多支持9列数据同时输出
- **📏 行数统计**：专门的节点统计Excel表格行数
- **🎯 统一输出**：所有列输出都是STRING类型，简化使用

## 📦 安装

### 快速安装
1. 将此文件夹复制到ComfyUI的`custom_nodes`目录下
2. 运行自动安装脚本：
```bash
python install.py
```
3. 重启ComfyUI

### ComfyUI Python Embedded环境安装
如果ComfyUI使用Python embedded版本，请使用以下方法：

```bash
# 在ComfyUI根目录运行
python_embeded\python.exe -m pip install pandas openpyxl
```

详细安装说明请查看 [COMFYUI_INSTALL.md](COMFYUI_INSTALL.md)

## 🎮 节点说明

### 1. Dynamic Excel Loader（动态Excel加载器）

**输入参数：**

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `excel_file_path` | STRING | "example.xlsx" | Excel文件路径 |
| `start_row_number` | INT | 2 | 开始读取的行号（2表示跳过表头） |
| `read_count` | INT | 1 | 每次读取的行数 |
| `refresh_trigger` | BOOLEAN | False | 手动刷新开关 |
| `auto_refresh` | BOOLEAN | True | 自动刷新开关 |
| `sheet_name` | STRING | "0" | 工作表名称或索引 |

**输出说明：**
- **column_1 到 column_9**：对应Excel的前9列数据（STRING类型）
- **current_row**：当前读取的行号（STRING类型）

### 2. Excel Row Counter（Excel行数统计器）

**输入参数：**

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `excel_file_path` | STRING | "example.xlsx" | Excel文件路径 |
| `include_header` | BOOLEAN | False | 是否包含表头行在计数中 |
| `refresh_trigger` | BOOLEAN | False | 手动刷新开关 |
| `auto_refresh` | BOOLEAN | True | 自动刷新开关 |
| `sheet_name` | STRING | "0" | 工作表名称或索引 |

**输出说明：**
- **row_count**：表格行数（INT类型）
- **status_message**：状态信息（STRING类型）

### 3. Excel Loader (For Loop)（For循环专用Excel加载器）

**特殊说明：** 专门设计用于与for循环节点配合使用

**输入参数：**

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `excel_file_path` | STRING | "example.xlsx" | Excel文件路径 |
| `start_row_number` | INT | 2 | 开始读取的行号（**无control_after_generate**） |
| `read_count` | INT | 1 | 每次读取的行数 |
| `refresh_trigger` | BOOLEAN | False | 手动刷新开关 |
| `auto_refresh` | BOOLEAN | True | 自动刷新开关 |
| `sheet_name` | STRING | "0" | 工作表名称或索引 |

**输出说明：**
- **column_1 到 column_9**：对应Excel的前9列数据（STRING类型）
- **current_row**：当前读取的行号（STRING类型）

**与普通版本的区别：**
- 移除了`control_after_generate`属性，避免与for循环冲突
- 不返回UI更新信息，完全由for循环控制参数

### 4. Excel Auto Loop Loader（Excel自动循环加载器）

**特殊说明：** 内置循环功能，无需外部for循环节点

**输入参数：**

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `excel_file_path` | STRING | "example.xlsx" | Excel文件路径 |
| `loop_mode` | ENUM | "single_row" | 循环模式：single_row=逐行循环，all_rows=一次读取所有行 |
| `reset_loop` | BOOLEAN | False | 重置循环（回到第一行） |
| `refresh_trigger` | BOOLEAN | False | 手动刷新开关 |
| `auto_refresh` | BOOLEAN | True | 自动刷新开关 |
| `sheet_name` | STRING | "0" | 工作表名称或索引 |

**输出说明：**
- **column_1 到 column_9**：对应Excel的前9列数据（STRING类型）
- **current_row**：当前读取的行号或行范围（STRING类型）
- **total_rows**：总数据行数（STRING类型）
- **loop_status**：循环状态信息（STRING类型）

**循环模式说明：**
- **single_row模式**：每次执行读取一行，自动循环到下一行，读完后自动重置到第一行
- **all_rows模式**：一次性读取所有行，每列数据用换行符分隔

**循环状态说明：**
- `reading_row_X`：正在读取第X行
- `loop_complete_reset`：循环完成并重置到第一行
- `all_rows_complete`：一次性读取所有行完成
- `out_of_range`：超出数据范围

### 5. Excel EasyUse For Loop（Excel EasyUse For循环适配器）

**特殊说明：** 专门适配EasyUse的forLoopStart和forLoopEnd节点

**输入参数：**

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `excel_file_path` | STRING | "example.xlsx" | Excel文件路径 |
| `index` | INT | 2 | 来自EasyUse forLoopStart的索引值（作为行号） |
| `refresh_trigger` | BOOLEAN | False | 手动刷新开关 |
| `auto_refresh` | BOOLEAN | True | 自动刷新开关 |
| `sheet_name` | STRING | "0" | 工作表名称或索引 |

**输出说明：**
- **column_1 到 column_9**：对应Excel的前9列数据（STRING类型）
- **row_info**：行信息，格式为"行X: 数据行Y/总数"（STRING类型）

**使用方法：**
1. 添加 `easy forLoopStart` 节点，设置 `start_value=2`, `end_value=6`, `step=1`
2. 将 `forLoopStart` 的 `index` 输出连接到此节点的 `index` 输入
3. 在循环体内使用此节点读取Excel数据
4. 连接到 `easy forLoopEnd` 节点完成循环

**索引说明：**
- `index=2` 对应Excel第2行（跳过表头）
- `index=3` 对应Excel第3行，以此类推
- 超出数据范围时返回空值和错误信息

### Excel文件格式要求

#### 表头设计
表头（第一行）用于确定列类型：
- 包含关键词（image、img、path、picture、photo）的列 → 图像路径列
- 其他列 → 文本列

#### 示例Excel结构
```
| 名称 | 描述 | 图像路径 | 标签 | 图片文件 | 价格 | 产地 |
|------|------|----------|------|----------|------|------|
| 苹果 | 红色的水果 | /images/apple.jpg | 水果,红色 | /photos/apple.png | 5.00 | 山东 |
| 香蕉 | 黄色的水果 | /images/banana.jpg | 水果,黄色 | /photos/banana.png | 3.50 | 海南 |
```

## 🔧 使用特点

- **统一输出类型**：所有列都输出为STRING类型，简化连接和使用
- **智能缓存**：避免重复读取，提高性能
- **参数变化检测**：当参数改变时自动重新计算
- **多工作表支持**：可以指定不同的工作表

## 🔄 刷新机制

### 自动刷新
- 当`auto_refresh`为True时，节点会监控Excel文件的修改时间
- 文件更新后会自动重新解析表头和数据结构

### 手动刷新  
- 切换`refresh_trigger`开关可强制重新解析Excel文件
- 适用于需要立即刷新的场景

## 📝 使用示例

### Dynamic Excel Loader 使用示例

1. **准备Excel文件**：创建包含表头的Excel文件
2. **添加节点**：在ComfyUI中添加"Dynamic Excel Loader"节点
3. **配置路径**：设置`excel_file_path`为你的Excel文件路径
4. **运行测试**：执行节点查看输出结果
5. **连接输出**：将需要的输出连接到其他节点

**输出示例：**
假设Excel有以下列：名称、描述、图像路径、标签、图片文件

输出将是：
- `column_1`: "苹果"
- `column_2`: "红色的水果"
- `column_3`: "/images/apple.jpg"
- `column_4`: "水果,红色"
- `column_5`: "/photos/apple.png"
- `column_6`: "" (空，因为只有5列)
- `column_7`: "" (空)
- `column_8`: "" (空)
- `column_9`: "" (空)
- `current_row`: "2" (当前行号)

### Excel Row Counter 使用示例

1. **添加节点**：在ComfyUI中添加"Excel Row Counter"节点
2. **配置路径**：设置`excel_file_path`为你的Excel文件路径
3. **选择计数方式**：设置`include_header`决定是否包含表头
4. **获取结果**：
   - `row_count`: 5 (如果不包含表头)
   - `status_message`: "成功：5 行数据（不包含表头）"

## ⚠️ 注意事项

1. **列数限制**：最多支持9列数据输出
2. **文件格式**：支持.xlsx和.xls格式
3. **表头必需**：Excel第一行必须是表头
4. **路径格式**：使用正斜杠(/)或反斜杠(\)都可以
5. **编码支持**：支持中文等Unicode字符

## 🐛 故障排除

### 常见问题

**Q: 节点显示"错误：运行此节点需要 Pandas 库"**
A: 运行 `pip install pandas openpyxl` 安装依赖

**Q: 行数统计不正确**
A: 检查`include_header`设置，确认是否要包含表头行

**Q: 输出为空**
A: 检查`start_row_number`设置，确保不超过数据行数

**Q: 文件修改后没有自动刷新**
A: 确保`auto_refresh`为True，或手动切换`refresh_trigger`

**Q: 无法与for循环配合使用**
A: 使用"Excel Loader (For Loop)"节点，它专门为for循环设计，移除了会冲突的`control_after_generate`属性

**Q: 不想使用外部for循环节点**
A: 使用"Excel Auto Loop Loader"节点，它内置循环功能，可以自动逐行读取或一次性读取所有行

**Q: 想要配合EasyUse的for循环节点使用**
A: 使用"Excel EasyUse For Loop"节点，它专门适配EasyUse的forLoopStart和forLoopEnd节点，接收index值作为行号

## 📄 许可证

本项目基于MIT许可证开源。

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！
