# ExportToXLSX ArcGIS工具箱

## 概述
这是一个ArcGIS工具箱，可以将要素类或表格数据导出到Excel(.xlsx)格式。该工具箱已经打包了所有必需的依赖库，可以在无网络环境下直接使用。

## 兼容性
- ArcGIS Desktop 10.2 及以上版本
- Python 2.7 和 Python 3.x
- 支持中文字段名和数据

## 功能特性
- ✅ 本地库依赖，无需网络连接
- ✅ Python 2/3 兼容性
- ✅ Unicode/中文字符支持
- ✅ 字段别名支持
- ✅ 域值和子类型描述转换
- ✅ 自定义字段选择
- ✅ 自定义工作表名称

## 安装和使用

### 安装步骤
1. 将整个 `ExportToXLSX` 文件夹复制到目标计算机的任意位置
2. 在ArcGIS Desktop中打开ArcToolbox
3. 右键点击ArcToolbox根节点，选择"添加工具箱"
4. 浏览到复制的文件夹，选择 `ExportToXLSX.tbx` 文件
5. 点击"打开"添加工具箱

### 使用说明
1. 展开ExportToXLSX工具箱
2. 双击"Export to XLSX"工具
3. 配置以下参数：
   - **输入要素类/表格**: 选择要导出的数据源
   - **输出Excel文件**: 指定输出的.xlsx文件路径
   - **使用字段别名**: 是否使用字段的别名作为Excel列标题
   - **使用域描述**: 是否将编码值域转换为描述值
   - **选择字段**: 选择要导出的字段（可选）
   - **工作表名称**: 指定Excel工作表的名称（可选）

## 目录结构
```
ExportToXLSX/
├── ExportToXLSX.py      # 主脚本文件
├── ExportToXLSX.tbx     # ArcGIS工具箱文件
├── README.md            # 说明文档
└── lib/                 # 依赖库目录
    ├── openpyxl_lib/    # Excel处理库
    ├── et_xmlfile_lib/  # XML处理库
    └── jdcal/           # 日期处理库
```

## 故障排除

### 导入错误
如果遇到"Cannot import openpyxl library"错误：
- 确保lib文件夹完整复制
- 检查lib/openpyxl_lib、lib/et_xmlfile_lib、lib/jdcal文件夹是否存在

### 中文字符问题
工具已自动处理中文字符编码，支持：
- 中文字段名
- 中文数据内容
- 中文工作表名称
- 中文输出路径

### 权限问题
- 确保ArcGIS有读取文件夹的权限
- 检查输出路径是否有效和可写
- 必要时以管理员权限运行ArcGIS

## 技术说明
该工具使用了以下主要库：
- **openpyxl**: 用于创建和操作Excel文件
- **et_xmlfile**: openpyxl的XML处理依赖
- **jdcal**: openpyxl的日期处理依赖

所有库都已包含在lib目录中，支持Python 2.7和Python 3.x，无需额外安装。
