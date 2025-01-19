# Excel Title Batch Insertion Tool

This tool is designed to extract titles from a specific column in an Excel file and batch insert them into the first row of multiple sheets in another Excel file. It supports unmerging existing cells, re-merging specified ranges, and inserting titles in a formatted manner.

## Features

- Extract titles from a specific column in a source Excel file.
- Insert these titles into the first row of multiple sheets in a target Excel file in sequence.
- Automatically unmerge any merged cells and re-merge the specified range.
- Insert titles with center alignment.
- Supports flexible configuration options such as file paths, title columns, starting rows, and sheets.

## Dependencies

- `openpyxl`: Used to read and modify Excel files. You can install it with the following command:
  ```bash
  pip install openpyxl


# Excel标题批量插入工具

该工具用于从一个 Excel 文件中的特定列提取标题，并批量将这些标题按顺序插入到另一个 Excel 文件的指定工作表的第一行。支持取消合并单元格、重新合并指定范围，并按格式插入标题。

## 功能

- 从源文件中提取指定列的标题。
- 将这些标题按顺序插入目标 Excel 文件的多个工作表的第一行。
- 自动取消已有的合并单元格，并重新合并指定范围的单元格。
- 以居中格式插入标题。
- 支持灵活配置的参数，如源文件路径、目标文件路径、标题列、开始行和工作表等。

## 依赖

- `openpyxl`：用于读取和修改 Excel 文件。可以使用以下命令安装：
  ```bash
  pip install openpyxl
