package main

// 程序入口
// 应用程序结构体和核心逻辑 app.go
// UI相关逻辑 ui.go
// Excel对比逻辑 compare.go
// 工具函数 utils
func main() {
	app := NewExcelCompareApp(750, 900)
	app.Run()
}
