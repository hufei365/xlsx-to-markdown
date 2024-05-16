# xlsx2md

将 XLSX 文件转成 Markdown 表格。

## 安装

```bash
npm install -g xlsx2md
```
## 使用
```bash
xlsx2md --file <path-to-xlsx> --sheet <sheet-name> --columns <comma-separated-columns> [--output <output-file>]
```
例子🌰：
```bash
xlsx2md --file data.xlsx --sheet Sheet1 --columns Name,Age,Location --output table.md
```
