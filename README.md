# BOM生成工具

一个Windows桌面应用程序，用于从源Excel文件自动化生成BOM表。

## 项目概述

BOM生成工具是一个专业的Windows桌面应用程序，旨在帮助用户从源Excel文件自动化生成高质量的BOM（Bill of Materials）表。该工具优先考虑稳定性、准确性和用户友好性。

## 技术栈

- **语言**: Python 3.10+
- **核心库**:
  - `openpyxl`: 详细读取、写入和操作`.xlsx`文件，包括格式保留和行插入
  - `pandas`: 高效地从源文件中初步读取数据
- **GUI框架**: `Tkinter` (Python标准GUI库)
- **打包工具**: `PyInstaller` (创建最终的`.exe`文件)
- **测试框架**: `pytest`

## 项目结构

```
/bom/
├── .gitignore
├── README.md
├── src/
│   ├── __init__.py
│   ├── main.py             # 程序入口，GUI逻辑
│   ├── core/
│   │   ├── __init__.py
│   │   └── bom_generator.py  # 核心逻辑，处理并生成文件
│   └── resources/
│       ├── bom_template.xlsx # 内置的BOM模板文件
│       └── color_codes.json  # 内置的颜色代码
└── tests/
    ├── __init__.py
    └── test_bom_generator.py # 核心逻辑的测试文件
```

## 安装要求

- Python 3.10+
- 依赖库（将在requirements.txt中指定）

## 使用方法

### 运行程序
1. 双击 `dist/BOM_Generator_v1.0.exe` 启动应用程序
2. 点击"选择..."按钮选择源Excel文件（包含产品明细数据）
3. 点击"选择..."按钮选择输出文件夹
4. 点击"开始生成"按钮自动生成所有款式的BOM文件

### 输入文件要求
- Excel格式（.xlsx）
- 包含名为"明细表"的工作表
- 必须包含以下列：款式编码、波段、品类、开发颜色

### 输出结果
- 为每个款式编码生成独立的BOM Excel文件
- 文件名格式：{款式编码}.xlsx
- 自动填充品名、颜色信息、SKU等

## 构建说明

### 开发环境运行
```bash
# 直接运行Python脚本
python src/main.py
```

### 打包为独立exe文件

1. **安装PyInstaller**
   ```bash
   pip install pyinstaller
   ```

2. **执行打包命令**
   ```bash
   pyinstaller --name BOM_Generator_v1.0 --onefile --windowed --add-data "src/resources;resources" src/main.py
   ```

3. **命令参数说明**
   - `--name BOM_Generator_v1.0`: 指定生成的exe文件名
   - `--onefile`: 打包成单个可执行文件
   - `--windowed`: GUI应用，不显示控制台窗口
   - `--add-data "src/resources;resources"`: 包含资源文件（模板和颜色代码）

4. **输出文件**
   - 可执行文件：`dist/BOM_Generator_v1.0.exe`
   - 文件大小：约10.7 MB
   - 支持独立运行，无需安装Python环境

### 重新构建
如需重新构建，先清理之前的构建文件：
```bash
# PowerShell
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# 然后重新执行打包命令
pyinstaller --name BOM_Generator_v1.0 --onefile --windowed --add-data "src/resources;resources" src/main.py
```

## 开发规范

本项目严格遵循TDD（测试驱动开发）工作流和约定式提交规范。

## 更新日志

### 2025-09-16 项目初始化
- ✅ 创建完整的项目目录结构（src/, tests/, resources/等）
- ✅ 创建所有初始Python文件（__init__.py, main.py, bom_generator.py等）
- ✅ 创建标准Python .gitignore文件
- ✅ 创建资源文件占位符（color_codes.json, bom_template.xlsx.placeholder）
- ✅ 完成本地Git提交（提交哈希：8dc3a95）
- ⚠️ Git推送因网络问题暂时失败，需要后续手动推送

### 2025-09-16 资源文件准备
- ✅ 从Excel文件成功提取颜色代码数据（204个有效颜色代码）
- ✅ 生成格式化的color_codes.json文件（过滤掉222个停用数据）
- ✅ 复制BOM模板文件到resources目录（bom_template.xlsx）
- ✅ 创建测试fixtures目录并复制测试文件
  - 输入源：新品研发明细表-最终版.xlsx (52MB)
  - 预期结果：H5A413492.xlsx (10KB)
- ✅ 完成本地Git提交（提交哈希：9c88c80）
- ⚠️ Git推送因网络连接问题失败，需要稍后手动推送

项目资源文件已完整准备，开发环境搭建完成，准备进入TDD功能开发阶段。

### 2025-09-16 最终细节修正
- ✅ 修正颜色名称写入位置：将PRESET_COLOR_BLOCKS配置中的color_cell从D列修改为A列
  - 第1个颜色块：D8 → A8
  - 第2个颜色块：D11 → A11  
  - 第3个颜色块：D14 → A14
- ✅ 移除错误的单元格合并：删除所有与写入颜色名称相关的sheet.merge_cells(...)代码
- ✅ 验证修改效果：直接测试确认颜色名称正确显示在A列，且没有错误的合并单元格
- ✅ 完成Git提交（提交哈希：327b142）
- ⚠️ 网络连接问题导致Git推送失败，本地提交已完成

**最终验证结果：**
- 颜色名称（白色、黑色、杏色等）正确显示在A8、A11、A14单元格
- 颜色名称所在行没有任何错误的单元格合并
- 产品核心功能已完美实现

### 2025-09-16 产品打包与交付
- ✅ 安装PyInstaller打包工具
- ✅ 修复模块导入路径问题：添加sys.path调整和资源文件路径适配
- ✅ 修复资源文件访问问题：实现resource_path函数兼容开发和打包环境
- ✅ 解决开发环境兼容性：使用`getattr(sys, 'frozen', False)`标准方法判断运行环境
- ✅ 优化路径处理逻辑：开发环境使用`os.path.abspath("src")`，打包环境使用`sys._MEIPASS`
- ✅ 修复main.py模块导入：实现与bom_generator.py一致的路径处理策略
- ✅ 成功打包为独立exe文件：BOM_Generator_v1.0.exe (10.7 MB)
- ✅ 全面验证通过：开发环境正常，打包环境正常，构建质量良好
- ✅ 完善构建说明文档：详细记录最终版PyInstaller打包命令

**最终交付物验证：**
- ✅ 开发环境：`python src/main.py` 正常启动和运行
- ✅ 打包环境：独立exe文件完美启动，无任何错误
- ✅ 资源文件：模板和颜色代码正确打包并可访问
- ✅ 跨环境兼容：无需Python环境即可在Windows系统上使用
- ✅ 功能完整性：所有BOM生成功能验证通过

🎉 **项目完美交付！** BOM生成工具已成功开发、调试并完美打包交付。

---
最后更新时间: 2025-09-16
