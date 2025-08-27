# Diff Excel
简单高效的 Excel 差异对比工具

Tools for comparing two Excel files

## 下载
- MacOS: [diffExcel](https://github.com/zbuzhi/diff-excel/releases/download/v1.1.0/diffExcel)
- Windows: [diffExcel.exe](https://github.com/zbuzhi/diff-excel/releases/download/v1.1.0/diffExcel.exe)


## 🚀 功能特性

### 核心功能
- **Excel文件对比**：支持对比两个Excel文件的差异
- **多Sheet对比**：支持选择多对Sheet同时对比，一对一映射，任意搭配
- **格式保持**：可选择保持原始文件Sheet的单元格格式（单元格大小、合并单元格、字体样式等）
- **差异高亮**：使用颜色高亮显示不同的单元格
- **备注功能**：可在差异单元格中添加备注显示原始内容
- **自定义颜色**：支持自定义差异高亮颜色


### 技术特点
- **模块化架构**：清晰的文件结构，易于维护和扩展
- **GUI界面**：基于Fyne框架的直观用户界面
- **跨平台**：支持Windows、macOS、Linux
- **高性能**：支持大型Excel文件处理

## 📋 快速开始

### 安装和运行

```bash
# 1. 克隆项目
git clone https://github.com/zbuzhi/diff-excel
cd diff_excel

# 2. 下载依赖
go mod tidy

# 3. 运行应用程序（推荐方式）
go run .

# 或者编译后运行
# Linux/macOS
go build -o DiffExcel
./DiffExcel

# Windows
go build -o DiffExcel
./DiffExcel.exe


# 在 macOS 上编译 Windows 可执行文件，需要安装 mingw-w64，配置 CC ⬇️

# 1. Homebrew 安装 mingw-w64
brew install mingw-w64
# 2. CC是否安装成功
x86_64-w64-mingw32-gcc --version
# 3. 编译 Windows 可执行文件
export CC=x86_64-w64-mingw32-gcc
GOOS=windows GOARCH=amd64 CGO_ENABLED=1 go build -o DiffExcel.exe
```

## 📄 许可证

本项目采用MIT许可证 - 详见 [LICENSE](LICENSE) 文件

## ⭐ 致谢

感谢以下开源项目：
- [Fyne](https://fyne.io/) - Go语言GUI框架
- [excelize](https://github.com/qax-os/excelize) - Go语言Excel处理库

---

如果这个项目对您有帮助，请给个⭐️支持一下！
