# Simple For Loop 节点使用说明

## 🎯 设计目标

重新实现一个简单可靠的for循环系统，专门用于Excel数据处理，解决EasyUse for循环节点的问题。

## 📦 包含的节点

### 1. Simple For Loop Start
- **功能**: 控制循环的开始和索引递增
- **输入**:
  - `loop_count`: 循环次数 (默认: 2)
  - `start_index`: 起始索引值 (默认: 1)
  - `reset_trigger`: 重置循环开关 (默认: false)
- **输出**:
  - `current_index`: 当前循环索引
  - `is_finished`: 是否完成循环
  - `loop_info`: 循环状态信息

### 2. Simple For Loop End
- **功能**: 接收循环体输出，控制循环结束
- **输入**:
  - `is_finished`: 来自ForLoopStart的完成状态
  - `loop_info`: 来自ForLoopStart的循环信息
  - `loop_data`: 循环体产生的数据（可选）
- **输出**:
  - `final_status`: 最终状态信息
  - `completed`: 是否完成
  - `collected_data`: 收集的数据

### 3. Simple Excel For Loop
- **功能**: 根据索引读取Excel数据
- **输入**:
  - `excel_file_path`: Excel文件路径
  - `current_index`: 来自ForLoopStart的当前索引
  - `skip_header`: 是否跳过表头行 (默认: true)
  - `sheet_name`: 工作表名称或索引 (默认: "0")
- **输出**:
  - `column_1` 到 `column_9`: Excel的前9列数据
  - `row_info`: 行信息

## 🔗 连接方式

```
[String] → excel_file_path → [Simple Excel For Loop]
                                      ↑
[Simple For Loop Start] → current_index
         ↓
    is_finished → [Simple For Loop End]
    loop_info   →

[Simple Excel For Loop] → column_X → [处理节点]
```

## 🎮 使用步骤

### 1. 基本设置
1. 添加 `Simple For Loop Start` 节点
2. 设置 `loop_count` = 你的数据行数
3. 设置 `start_index` = 1 (如果要从第1行开始)

### 2. Excel读取设置
1. 添加 `Simple Excel For Loop` 节点
2. 连接 `ForLoopStart.current_index` → `ExcelForLoop.current_index`
3. 设置Excel文件路径
4. 设置 `skip_header` = true (如果要跳过表头)

### 3. 循环控制
1. 添加 `Simple For Loop End` 节点
2. 连接 `ForLoopStart.is_finished` → `ForLoopEnd.is_finished`
3. 连接 `ForLoopStart.loop_info` → `ForLoopEnd.loop_info`

### 4. 数据处理
1. 将 `ExcelForLoop.column_X` 连接到你的处理节点
2. 处理完成后可以连接到 `ForLoopEnd.loop_data`

## 📊 索引说明

### 如果 skip_header = true (推荐):
- `start_index = 1`: 读取数据第1行 (Excel第2行)
- `start_index = 2`: 读取数据第2行 (Excel第3行)
- `loop_count = 2`: 循环2次，读取2行数据

### 如果 skip_header = false:
- `start_index = 1`: 读取表头 (Excel第1行)
- `start_index = 2`: 读取数据第1行 (Excel第2行)
- `loop_count = 3`: 循环3次，读取表头+2行数据

## 🔧 示例配置

### 读取2行数据（跳过表头）:
```
Simple For Loop Start:
  - loop_count: 2
  - start_index: 1
  - reset_trigger: false

Simple Excel For Loop:
  - skip_header: true
  - sheet_name: "0"
```

### 读取包含表头的3行:
```
Simple For Loop Start:
  - loop_count: 3
  - start_index: 1
  - reset_trigger: false

Simple Excel For Loop:
  - skip_header: false
  - sheet_name: "0"
```

## 🐛 调试功能

- 所有节点都有详细的控制台输出
- `loop_info` 输出显示当前循环状态
- `row_info` 输出显示当前读取的行信息
- 自动检测索引超出范围的情况

## 🔄 重置循环

如果需要重新开始循环：
1. 切换 `Simple For Loop Start` 的 `reset_trigger` 开关
2. 重新运行工作流

## ⚡ 优势

1. **简单可靠**: 不依赖复杂的EasyUse框架
2. **自动递增**: 每次循环自动将index+1
3. **明确控制**: 可以精确指定循环次数
4. **详细调试**: 丰富的控制台输出
5. **灵活配置**: 支持跳过表头、指定工作表等
6. **错误处理**: 完善的异常处理和边界检查
