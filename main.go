package main

import (
	"fmt"
	"image/color"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

type ExcelCompareApp struct {
	myApp        fyne.App
	myWindow     fyne.Window
	logRich      *widget.Entry
	srcFile      string
	cmpFile      string
	srcSheet     string
	cmpSheet     string
	outDir       string
	outExcelFile string
	outLogFile   string
	highlightClr string

	srcEntry       *widget.Entry
	srcEntryBox    fyne.CanvasObject
	cmpEntry       *widget.Entry
	cmpEntryBox    fyne.CanvasObject
	outDirEntry    *widget.Entry
	outDirBox      fyne.CanvasObject
	srcSheetSelect *widget.Select
	cmpSheetSelect *widget.Select
	colorEntry     *widget.Entry

	srcUploadFunc func()
	cmpUploadFunc func()
}

func NewExcelCompareApp(w, h float32) *ExcelCompareApp {
	a := &ExcelCompareApp{
		myApp:        app.NewWithID("com.zbuzhi.diffexcel"),
		highlightClr: "#FF0000",
	}
	a.myWindow = a.myApp.NewWindow("Excel 对比")
	a.myWindow.Resize(fyne.NewSize(w+40, h))

	// 初始化只读日志框
	a.logRich = widget.NewMultiLineEntry()
	a.logRich.Wrapping = fyne.TextWrapWord
	a.logRich.Disable() // 只读模式，禁止手动输入

	a.initUI()
	return a
}

// 追加日志：同步终端输出和日志框内容
func (a *ExcelCompareApp) appendLog(text string) {
	fmt.Print(text) // 保留终端输出
	a.logRich.SetText(a.logRich.Text + text)
}

// 宽 Entry 制作函数，禁用编辑
func (a *ExcelCompareApp) makeWideMultiLineEntry(width, height float32, defaultInput string) (*widget.Entry, fyne.CanvasObject) {
	e := widget.NewMultiLineEntry()
	e.Wrapping = fyne.TextWrapWord
	e.SetMinRowsVisible(2) // 让输入框高度能显示两行
	e.Disable()            // 只读
	e.SetText(defaultInput)
	e.SetPlaceHolder("请输入文件地址")

	box := container.NewGridWrap(fyne.NewSize(width, height), e)
	return e, box
}

func (a *ExcelCompareApp) initUI() {
	//execPath, _ := getExeDir()
	w := a.myWindow.Canvas().Size().Width
	a.srcEntry, a.srcEntryBox = a.makeWideMultiLineEntry(w/2-10, 60, "")
	a.cmpEntry, a.cmpEntryBox = a.makeWideMultiLineEntry(w/2-10, 60, "")
	a.outDirEntry, a.outDirBox = a.makeWideMultiLineEntry(w/2-10, 60, "")

	a.srcSheetSelect = widget.NewSelect([]string{}, func(selected string) {
		a.srcSheet = selected
	})
	a.srcSheetSelect.PlaceHolder = "上传文件后选择Sheet"

	a.cmpSheetSelect = widget.NewSelect([]string{}, func(selected string) {
		a.cmpSheet = selected
	})
	a.cmpSheetSelect.PlaceHolder = "上传文件后选择Sheet"

	makeClearBtn := func(entry *widget.Entry, sheetSelect *widget.Select, clearVars ...*string) *widget.Button {
		return widget.NewButton("清空", func() {
			entry.SetText("")
			for _, v := range clearVars {
				*v = ""
			}
			if sheetSelect != nil {
				sheetSelect.Options = []string{}
				sheetSelect.SetSelected("")
			}
		})
	}
	// 解析 hex 颜色
	parseHexColor := func(s string) color.Color {
		s = strings.TrimPrefix(s, "#")
		if len(s) != 6 {
			return color.Black
		}
		r, _ := strconv.ParseUint(s[0:2], 16, 8)
		g, _ := strconv.ParseUint(s[2:4], 16, 8)
		b, _ := strconv.ParseUint(s[4:6], 16, 8)
		return color.RGBA{uint8(r), uint8(g), uint8(b), 255}
	}
	// 创建带彩色方块选择器的颜色输入框
	makeColorSelector := func(onChange func(string)) fyne.CanvasObject {
		// 常用颜色列表
		colors := []string{
			"#FF0000", "#00FF00", "#0000FF",
			"#800080", "#00FFFF", "#A52A2A",
			"#008000", "#808000", "#800000",
			"#008080", "#FFD700", "#FF6347",
		}

		var colorButtons []fyne.CanvasObject
		for _, c := range colors {
			colorCode := c
			// 彩色背景方块
			rect := canvas.NewRectangle(parseHexColor(c))
			rect.SetMinSize(fyne.NewSize(28, 20))

			// 透明按钮覆盖在方块上
			btn := widget.NewButton("", func() {
				if onChange != nil {
					onChange(colorCode)
				}
			})
			btn.Importance = widget.LowImportance // 去掉高亮样式

			// 把按钮和颜色块叠加
			colorButtons = append(colorButtons, container.NewMax(rect, btn))
		}

		return container.NewHBox(
			container.NewHBox(colorButtons...),
		)
	}

	// 上传按钮封装
	a.srcUploadFunc = func() {
		fd := dialog.NewFileOpen(func(r fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, a.myWindow)
				return
			}
			if r == nil {
				return
			}
			a.srcFile = r.URI().Path()
			a.srcEntry.SetText(a.srcFile)
			r.Close()

			sheets, err := getSheets(a.srcFile)
			if err != nil {
				dialog.ShowError(err, a.myWindow)
				return
			}
			a.srcSheetSelect.Options = sheets
			if len(sheets) > 0 {
				a.srcSheetSelect.SetSelected(sheets[0])
				a.srcSheet = sheets[0]
			} else {
				a.srcSheetSelect.SetSelected("")
				a.srcSheet = ""
			}
		}, a.myWindow)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	}
	a.cmpUploadFunc = func() {
		fd := dialog.NewFileOpen(func(r fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, a.myWindow)
				return
			}
			if r == nil {
				return
			}
			a.cmpFile = r.URI().Path()
			a.cmpEntry.SetText(a.cmpFile)
			r.Close()

			sheets, err := getSheets(a.cmpFile)
			if err != nil {
				dialog.ShowError(err, a.myWindow)
				return
			}
			a.cmpSheetSelect.Options = sheets
			if len(sheets) > 0 {
				a.cmpSheetSelect.SetSelected(sheets[0])
				a.cmpSheet = sheets[0]
			} else {
				a.cmpSheetSelect.SetSelected("")
				a.cmpSheet = ""
			}
		}, a.myWindow)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	}

	srcBtn := widget.NewButton("上传", a.srcUploadFunc)
	cmpBtn := widget.NewButton("上传", a.cmpUploadFunc)
	setBlue(srcBtn)
	setBlue(cmpBtn)

	srcClearBtn := makeClearBtn(a.srcEntry, a.srcSheetSelect, &a.srcFile, &a.srcSheet)
	cmpClearBtn := makeClearBtn(a.cmpEntry, a.cmpSheetSelect, &a.cmpFile, &a.cmpSheet)

	outDirBtn := widget.NewButton("选择输出目录", func() {
		fd := dialog.NewFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, a.myWindow)
				return
			}
			if uri == nil {
				return
			}
			a.outDir = uri.Path()
			a.outDirEntry.SetText(a.outDir)
		}, a.myWindow)
		fd.Show()
	})
	outDirClearBtn := widget.NewButton("清空", func() {
		a.outDir = ""
		a.outDirEntry.SetText("")
	})

	a.colorEntry = widget.NewEntry()
	a.colorEntry.SetText(a.highlightClr)
	a.colorEntry.SetPlaceHolder("请输入颜色代码，如 #FF0000")

	compareBtn := widget.NewButton("开始对比", func() {
		a.compareFunc()
	})
	setBlue(compareBtn)

	// 日志框滚动容器，固定高度
	logScroll := container.NewScroll(a.logRich)
	logScroll.SetMinSize(fyne.NewSize(750, 150))

	// 清空日志按钮，带确认对话框
	clearLogBtn := widget.NewButton("清空日志", func() {
		dialog.ShowConfirm("确认清空", "确定要清空日志吗？", func(ok bool) {
			if ok {
				a.logRich.SetText("")
				a.appendLog("")
			}
		}, a.myWindow)
	})

	// 日志标题栏，带清空按钮
	logHeader := container.NewBorder(nil, nil, nil, clearLogBtn, widget.NewLabel("日志输出："))

	setTitle := func(str string) *canvas.Text {
		text := canvas.NewText(str, &color.NRGBA{R: 255, G: 255, B: 255, A: 255}) // 纯黑色
		text.TextStyle = fyne.TextStyle{Bold: true}
		text.TextSize = 18
		return text
	}
	spacing := canvas.NewRectangle(nil)     // 空矩形
	spacing.SetMinSize(fyne.NewSize(0, 10)) // 高度 10 像素，宽度 0
	leftBox := container.NewVBox(
		setTitle("原始文件"),
		spacing,
		container.NewHBox(srcBtn, a.srcSheetSelect, srcClearBtn),
		a.srcEntryBox,
	)
	rightBox := container.NewVBox(
		setTitle("对比文件"),
		spacing,
		container.NewHBox(cmpBtn, a.cmpSheetSelect, cmpClearBtn),
		a.cmpEntryBox,
	)

	leftBox2 := container.NewVBox(
		setTitle("输出目录"),
		spacing,
		container.NewHBox(outDirBtn, outDirClearBtn),
		a.outDirBox,
	)

	rightBox2 := container.NewVBox(
		setTitle("高亮颜色 (#RRGGBB)"),
		spacing,
		makeColorSelector(func(color string) {
			a.colorEntry.SetText(color)
		}),
		a.colorEntry,
	)

	content := container.NewVBox(
		container.New(layout.NewGridLayout(2), leftBox, rightBox),
		spacing,
		container.New(layout.NewGridLayout(2), leftBox2, rightBox2),

		container.NewVBox(
			widget.NewLabel(""),
			widget.NewSeparator(), // 分割线
		),

		compareBtn,
		logHeader,
		logScroll,
	)

	paddedContent := container.NewVBox(
		spacing, // 上间距
		container.NewHBox(
			spacing, // 左间距
			content,
			spacing, // 右间距
		),
		spacing, // 下间距
	)

	a.myWindow.SetContent(container.NewScroll(paddedContent))
	//a.myWindow.Resize(fyne.NewSize(800, 850))
}

func (a *ExcelCompareApp) compareFunc() {
	if a.srcFile == "" || a.cmpFile == "" {
		dialog.ShowError(fmt.Errorf("请先选择源 Excel 和对比 Excel 文件"), a.myWindow)
		return
	}
	if a.srcSheet == "" {
		dialog.ShowError(fmt.Errorf("请选择源 Excel 的 Sheet"), a.myWindow)
		return
	}
	if a.cmpSheet == "" {
		dialog.ShowError(fmt.Errorf("请选择对比 Excel 的 Sheet"), a.myWindow)
		return
	}
	if a.outDir == "" {
		dialog.ShowError(fmt.Errorf("请选择输出目录"), a.myWindow)
		return
	}

	inputColor := strings.TrimSpace(a.colorEntry.Text)
	if !isValidColorCode(inputColor) {
		dialog.ShowError(fmt.Errorf("颜色格式错误，需形如 #RRGGBB 或 #RGB"), a.myWindow)
		return
	}
	// a.highlightClr = inputColor
	a.outExcelFile = filepath.Join(a.outDir, fmt.Sprintf("diff_excel_%s.xlsx", time.Now().Format("2006.01.02_15_04_05")))
	a.outLogFile = filepath.Join(a.outDir, fmt.Sprintf("diff_log_%s.txt", time.Now().Format("2006.01.02_15_04_05")))

	a.appendLog("\n====== [" + time.Now().Format("2006.01.02 15:04:05") + "] 开始======\n")
	err := CompareExcelFiles(a.srcFile, a.srcSheet, a.cmpFile, a.cmpSheet, a.highlightClr, a.outExcelFile, a.outLogFile, func(s string) {
		a.appendLog(s)
	})
	if err != nil {
		dialog.ShowError(err, a.myWindow)
		a.appendLog(fmt.Sprintf("错误：%v\n", err))
	} else {
		a.appendLog(fmt.Sprintf("Excel文件: %s\n日志文件: %s\n", a.outExcelFile, a.outLogFile))
		a.appendLog("====== [" + time.Now().Format("2006.01.02 15:04:05") + "] 结束======\n\n")
	}
}

func (a *ExcelCompareApp) Run() {
	a.myWindow.ShowAndRun()
}

// 纯逻辑函数，不依赖 UI
func CompareExcelFiles(srcFile, srcSheet, cmpFile, cmpSheet, highlightColor, outExcel, outLog string, controlTxt func(string)) error {
	src, err := excelize.OpenFile(srcFile)
	if err != nil {
		controlTxt(fmt.Sprintf("打开源 Excel 出错: %v", err))
		return fmt.Errorf("打开源 Excel 出错: %v", err)
	}
	defer src.Close()

	cmp, err := excelize.OpenFile(cmpFile)
	if err != nil {
		controlTxt(fmt.Sprintf("打开对比 Excel 出错: %v", err))
		return fmt.Errorf("打开对比 Excel 出错: %v", err)
	}
	defer cmp.Close()

	srcRows, err := src.GetRows(srcSheet)
	if err != nil {
		controlTxt(fmt.Sprintf("读取源 Excel 行失败: %v", err))
		return fmt.Errorf("读取源 Excel 行失败: %v", err)
	}
	cmpRows, err := cmp.GetRows(cmpSheet)
	if err != nil {
		controlTxt(fmt.Sprintf("读取对比 Excel 行失败: %v", err))
		return fmt.Errorf("读取对比 Excel 行失败: %v", err)
	}

	diffF := excelize.NewFile()
	diffSheet := diffF.GetSheetName(0)

	style := &excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{highlightColor},
		},
	}
	styleID, err := diffF.NewStyle(style)
	if err != nil {
		controlTxt(fmt.Sprintf("创建样式失败: %v", err))
		return fmt.Errorf("创建样式失败: %v", err)
	}

	maxRow := len(srcRows)
	diffMaxRow := len(cmpRows)
	if diffMaxRow > maxRow {
		maxRow = diffMaxRow
	}

	controlTxt(fmt.Sprintf("【原始文件】%s 行数据\n", strconv.Itoa(maxRow)))
	controlTxt(fmt.Sprintf("【对比文件】%s 行数据\n", strconv.Itoa(diffMaxRow)))

	var logBuilder strings.Builder

	controlTxt("\n\n --------- 差异单元格 --------- \n")
	diffCount := 0
	for r := 0; r < maxRow; r++ {
		maxCol := 0
		if r < len(srcRows) && len(srcRows[r]) > maxCol {
			maxCol = len(srcRows[r])
		}
		if r < len(cmpRows) && len(cmpRows[r]) > maxCol {
			maxCol = len(cmpRows[r])
		}

		for c := 0; c < maxCol; c++ {
			var oldVal, newVal string
			if r < len(srcRows) && c < len(srcRows[r]) {
				oldVal = srcRows[r][c]
			}
			if r < len(cmpRows) && c < len(cmpRows[r]) {
				newVal = cmpRows[r][c]
			}

			cell, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if oldVal != newVal {
				diffCount++
				controlTxt(fmt.Sprintf(" %s |", cell))
				logLine := fmt.Sprintf("差异单元格: %s 旧数据: %s 新数据: %s\n", cell, oldVal, newVal)
				logBuilder.WriteString(logLine)
				diffF.SetCellValue(diffSheet, cell, newVal)
				diffF.SetCellStyle(diffSheet, cell, cell, styleID)
			} else {
				diffF.SetCellValue(diffSheet, cell, newVal)
			}
		}
	}
	controlTxt(fmt.Sprintf("\n\n --------- 差异数：%v -------- \n", diffCount))

	if err := diffF.SaveAs(outExcel); err != nil {
		controlTxt(fmt.Sprintf("保存差异 Excel 文件失败: %v", err))
		return fmt.Errorf("保存差异 Excel 文件失败: %v", err)
	}

	err = os.WriteFile(outLog, []byte(logBuilder.String()), 0644)
	if err != nil {
		controlTxt(fmt.Sprintf("写入日志 TXT 文件失败: %v", err))
		return fmt.Errorf("写入日志 TXT 文件失败: %v", err)
	}

	return nil
}

func getSheets(path string) ([]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return f.GetSheetList(), nil
}

func isValidColorCode(s string) bool {
	if len(s) != 7 && len(s) != 4 {
		return false
	}
	if s[0] != '#' {
		return false
	}
	for _, c := range s[1:] {
		if !((c >= '0' && c <= '9') ||
			(c >= 'a' && c <= 'f') ||
			(c >= 'A' && c <= 'F')) {
			return false
		}
	}
	return true
}

func getExeDir() (string, error) {
	exePath, err := os.Executable()
	if err != nil {
		return "", err
	}
	dir := filepath.Dir(exePath)
	return dir, nil
}

func setBlue(b *widget.Button) {
	b.Importance = widget.HighImportance
}

func main() {
	app := NewExcelCompareApp(750, 650)
	app.Run()
}
