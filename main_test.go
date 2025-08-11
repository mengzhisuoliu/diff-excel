package main

import (
	"os"
	"testing"
)

// 这里测试 CompareExcelFiles 函数，跳过UI，直接调用纯逻辑
func TestCompareExcelFiles(t *testing.T) {
	srcFile := "testdata/A.xlsx"
	cmpFile := "testdata/B.xlsx"
	srcSheet := "1.总表"
	cmpSheet := "1.总表"
	outExcel := "testdata/out_diff.xlsx"
	outLog := "testdata/out_diff.txt"
	color := "#FF0000"

	// srcFile = "testdata/AA.xlsx"
	// cmpFile = "testdata/BB.xlsx"
	// srcSheet = "Sheet1"
	// cmpSheet = "Sheet1"

	// 运行对比逻辑
	err := CompareExcelFiles(srcFile, srcSheet, cmpFile, cmpSheet, color, outExcel, outLog)
	if err != nil {
		t.Fatalf("CompareExcelFiles failed: %v", err)
	}

	// 简单断言：检查输出文件是否生成
	if _, err := os.Stat(outExcel); os.IsNotExist(err) {
		t.Errorf("输出 Excel 文件未生成")
	}
	if _, err := os.Stat(outLog); os.IsNotExist(err) {
		t.Errorf("输出日志文件未生成")
	}

	// 测试完可以删除生成文件，保持环境干净
	// os.Remove(outExcel)
	// os.Remove(outLog)
}
