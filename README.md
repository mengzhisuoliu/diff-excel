# Diff Excel
简单高效的 Excel 差异对比工具

Tools for comparing two Excel files

## 描述：
- 支持选择两个 Excel 文件及指定 Sheet 进行内容对比
- 自动生成对比差异的 Excel 文件，高亮显示不同单元格
- 输出差异日志TXT，详细记录差异内容和位置
- 简洁直观的图形界面，易用且跨平台（支持 `Windows`、`macOS`、`Linux`）
- 使用 `Go` + `fyne` + `excelize` 实现，轻量高效

<img width="500" alt="diffExcel_v1" src="https://github.com/user-attachments/assets/09253ebb-c056-4058-a20c-82cd9c024b49" />


## 下载
- MacOS: [diffExcel](https://github.com/zbuzhi/diff-excel/releases/download/v1.0.1/diffExcel)
- Windows: [diffExcel.exe](https://github.com/zbuzhi/diff-excel/releases/download/v1.0.1/diffExcel.exe)

## 运行
```sh
# 下载依赖
go mod tidy

# 运行
go run main.go
```
## 编译
```sh
# linux/macOS
go build -o DiffExcel

# Windows
go build -o DiffExcel.exe
```

### 在 macOS 上编译 Windows 可执行文件
需要安装`mingw-w64`，配置`CC`
```sh
# 1. Homebrew 安装 mingw-w64
brew install mingw-w64

# 2. CC是否安装成功
x86_64-w64-mingw32-gcc --version

# 3. 编译 Windows 可执行文件
export CC=x86_64-w64-mingw32-gcc
GOOS=windows GOARCH=amd64 CGO_ENABLED=1 go build -o DiffExcel.exe
```
