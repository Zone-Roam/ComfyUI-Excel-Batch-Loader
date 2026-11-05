# ComfyUI Dynamic Excel Loader 安装指南

## 🎯 针对ComfyUI Python Embedded环境的安装方法

### 📋 前提条件
- 已安装ComfyUI
- ComfyUI使用Python embedded版本

### 🚀 安装步骤

#### 步骤1：复制文件
将整个`ComfyUI_DynamicExcelLoader`文件夹复制到：
```
ComfyUI/custom_nodes/ComfyUI_DynamicExcelLoader/
```

#### 步骤2：安装依赖包

**方法A：自动安装（推荐）**
```bash
# 在ComfyUI_DynamicExcelLoader目录中运行
python install.py
```

**方法B：手动安装**

1. **打开命令行**，进入ComfyUI根目录

2. **运行安装命令**：
   ```bash
   # Windows
   python_embeded\python.exe -m pip install pandas openpyxl
   
   # 如果上面不行，尝试：
   python_embeded\Scripts\pip.exe install pandas openpyxl
   ```

3. **验证安装**：
   ```bash
   python_embeded\python.exe -c "import pandas, openpyxl; print('安装成功！')"
   ```

**方法C：使用requirements.txt**
```bash
# 在ComfyUI根目录运行
python_embeded\python.exe -m pip install -r custom_nodes\ComfyUI_DynamicExcelLoader\requirements.txt
```

#### 步骤3：重启ComfyUI
重启ComfyUI以加载新节点

#### 步骤4：验证安装
1. 在ComfyUI中查找"Dynamic Excel Loader"节点
2. 如果找到节点，说明安装成功

### 🔧 常见问题解决

#### Q1: 找不到python_embeded目录
**A:** 检查ComfyUI安装方式：
- **便携版**：通常在ComfyUI根目录有`python_embeded`文件夹
- **Anaconda版**：使用conda环境，运行 `conda install pandas openpyxl`
- **系统Python版**：使用 `pip install pandas openpyxl`

#### Q2: pip命令不存在
**A:** 尝试以下方法：
```bash
# 方法1：直接使用python安装
python_embeded\python.exe -c "import subprocess; subprocess.check_call(['python', '-m', 'pip', 'install', 'pandas', 'openpyxl'])"

# 方法2：下载wheel文件手动安装
# 从 https://pypi.org/ 下载pandas和openpyxl的wheel文件
python_embeded\python.exe -m pip install pandas-2.3.1-cp311-cp311-win_amd64.whl
python_embeded\python.exe -m pip install openpyxl-3.1.5-py2.py3-none-any.whl
```

#### Q3: 权限错误
**A:** 以管理员身份运行命令行

#### Q4: 网络连接问题
**A:** 使用国内镜像源：
```bash
python_embeded\python.exe -m pip install pandas openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple/
```

### 📁 目录结构检查

安装完成后，目录结构应该是：
```
ComfyUI/
├── python_embeded/          # Python环境
├── custom_nodes/
│   └── ComfyUI_DynamicExcelLoader/
│       ├── dynamic_excel_loader_node.py
│       ├── __init__.py
│       ├── requirements.txt
│       ├── README.md
│       ├── example.xlsx
│       └── ...
└── ...
```

### 🧪 测试安装

在ComfyUI_DynamicExcelLoader目录中运行：
```bash
# 使用ComfyUI的Python测试
../../python_embeded/python.exe test_node.py
```

### 💡 其他安装方式

#### 使用ComfyUI Manager（如果支持）
1. 打开ComfyUI Manager
2. 搜索"Dynamic Excel Loader"
3. 点击安装

#### 手动复制依赖
如果无法使用pip，可以手动复制pandas和openpyxl库：
1. 在有网络的环境安装依赖
2. 复制site-packages中的pandas和openpyxl文件夹
3. 粘贴到ComfyUI的python_embeded/Lib/site-packages/目录

### 🎉 安装完成

安装成功后，您就可以在ComfyUI中使用Dynamic Excel Loader节点了！

节点位置：`ExcelUtils > Dynamic Excel Loader`
