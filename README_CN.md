<h4 align="center">**简体中文**|[English](./README.md)</h4>

### 本库使用说名

- 1.首先特别感谢原作者提供的开源库：[https://github.com/qifengzhang007/gooxml](https://github.com/qifengzhang007/gooxml)。   
  由于原始仓库不再维护，无法通过提交代码的方式修复原始仓库中存在的问题，因此我使用了原始仓库的代码，进行二次开发/修复一些可能存在的bug。
- 2.原始包中 `99%` 的命令在本包都是相同的,只有少量的命令参数有增加，请根据报错自行修正即可.

### 关于 gooxml 库

**gooxml** 是一个用于创建 Office Open XML 文档（.docx、.xlsx 和 .pptx）的库。它的目标是成为创建和编辑 docx/xlsx/pptx 文件的最兼容、最高性能的 Go 库。

### 包安装

```code  
    // 在安装此包之前，请检查最新的标签版本号，例如 v1.0.12-alpha
    go get  github.com/qifengzhang007/gooxml@v1.0.12-alpha
```

![gooxml](./_examples/preview.png "gooxml")

### 目前具备的主要功能

- 文档 (docx) [Word]
    - 读/写/编辑
    - 格式化
    - 图像
    - 表格
- 电子表格 (xlsx) [Excel]
    - 读/写/编辑
    - 单元格格式化，包括条件格式化
    - 单元格验证（下拉组合框、规则等）
    - 以 Excel 格式获取单元格值（例如，获取在 Excel 中显示的日期或数字）
    - 公式评估（目前支持 100 多个函数，根据需要会添加更多）
    - 嵌入图像
    - 所有图表类型
- 幻灯片 (pptx) [PowerPoint]
    - 从模板创建
    - 文本框/形状

### 性能

最近对电子表格创建/读取的性能数字很感兴趣，因此这里是 gooxml 在这个 [基准测试](./_examples/spreadsheet/lots-of-rows) 中的数字，该测试创建了一个包含 30k
行的表格，每行有 100 列。

    创建 30000 行 * 100 单元格耗时 3.92506863s
    保存耗时 89ns
    读取耗时 9.522383048s

创建速度相当快，由于不使用反射，保存速度非常快，而读取速度稍慢。缺点是二进制文件较大（33MB），因为它包含了所有 DOCX/XLSX/PPTX 的生成结构体、序列化和反序列化代码。

### word文档示例

- [简单文本格式化](./_examples/document/simple)    文本字体颜色、大小、突出显示等。
- [自动生成目录](./_examples/document/toc)    创建文档标题，并基于标题自动生成目录。
- [插入图像](./_examples/document/image)    在页面上放置图像，绝对定位，具有不同的文本环绕方式。
- [页眉 & 页脚](./_examples/document/header-footer)    创建页眉和页脚，包括页码。
- [多个页眉 & 页脚](./_examples/document/header-footer-multiple)    根据文档部分使用不同的页眉和页脚。
- [内嵌表格](./_examples/document/tables)    添加有边框和无边框的表格。
- [使用现有 Word 文档作为模板](./_examples/document/use-template)    打开文档作为模板，重用文档中创建的样式。
- [填写表单字段](./_examples/document/fill-out-form)    打开包含嵌入表单字段的文档，填写字段并将结果保存为新的已填写表单。
- [编辑现有文档](./_examples/document/edit-document)    打开现有文档并替换/删除文本，而不修改格式。

### 电子表格示例

- [简单](./_examples/spreadsheet/simple)    一个包含几个单元格的简单表格
- [命名单元格](./_examples/spreadsheet/named-cells)    引用行和单元格的不同方式
- [单元格数字/日期/时间格式](./_examples/spreadsheet/number-date-time-formats)    创建具有各种数字/日期/时间格式的单元格
- [折线图](./_examples/spreadsheet/line-chart)/[折线图 3D](./_examples/spreadsheet/line-chart-3d)
  折线图
- [柱状图](./_examples/spreadsheet/bar-chart)    柱状图
- [多个图表](./_examples/spreadsheet/multiple-charts)    在单个表格上多个图表
- [命名单元格范围](./_examples/spreadsheet/named-ranges)    命名单元格范围
- [合并单元格](./_examples/spreadsheet/merged)    合并和拆分单元格
- [条件格式化](./_examples/spreadsheet/conditional-formatting)    条件格式化单元格，样式，渐变，图标，数据条
- [复杂](./_examples/spreadsheet/complex)    多个图表，自动筛选和条件格式化
- [边框](./_examples/spreadsheet/borders)    单个单元格边框和围绕单元格范围的矩形边框。
- [验证](./_examples/spreadsheet/validation)    数据验证，包括组合框下拉列表。
- [冻结行/列](./_examples/spreadsheet/freeze-rows-cols)    具有冻结标题列和行的表格

### 幻灯片示例

- [简单文本框](./_examples/presentation/simple)    简单文本框和形状
- [图像](https://github.com/qifengzhang007/gooxml)
