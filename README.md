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

（功能开发完成后更新）

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

---
最后更新时间: 2025-09-16
