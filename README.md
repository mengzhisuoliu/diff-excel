# Diff Excel
简单高效的 Excel 差异对比工具

Tools for comparing two Excel files

## 描述：
- 支持选择两个 Excel 文件及指定 Sheet 进行内容对比
- 自动生成对比差异的 Excel 文件，高亮显示不同单元格
- 输出差异日志TXT，详细记录差异内容和位置
- 简洁直观的图形界面，易用且跨平台（支持 `Windows`、`macOS`、`Linux`）
- 使用 `Go` + `fyne` + `excelize` 实现，轻量高效

<img width="500" alt="diffExcel_v1" src="https://github.com/user-attachments/assets/3df203a1-a948-4e3c-9f19-a07d830e1f78" />

## 下载
- MacOS: [diffExcel](https://github.com/zbuzhi/diff-excel/releases/download/v1.0.0/diffExcel)
- Windows: [diffExcel.exe](https://github.com/zbuzhi/diff-excel/releases/download/v1.0.0/diffExcel.exe)

## 运行/编译
```go
# 下载依赖
go mod tidy

# 运行
go run main.go

# 编译（linux/macOS）
go build -o diffExcel

# 编译（Windows）
go build -o diffExcel.exe
```
