package main

// report_worksheet.go
// Worksheet Report — กระดาษทำการ (ถอด Logic จาก FoxPro WKS1, WKS2)

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// Data Structures
// ─────────────────────────────────────────────────────────────────
type WorksheetRow struct {
	AcCode  string
	AcName  string
	Opening float64     // ยอดยกมา (เฉพาะ Type 1)
	Periods [12]float64 // P01 - P12
	Total   float64     // ยอดยกไป (Type 1) หรือ ยอดรวม (Type 2)
}

// ─────────────────────────────────────────────────────────────────
// Core Logic: ดึงข้อมูลจาก Ledger_Master โดยตรง
//
// ใช้ Ledger_Master เป็น single source of truth แทนการคำนวณซ้ำจาก Book_items
// เพื่อให้สอดคล้องกับ Trial Balance และ Ledger UI ทุกรายงาน
//
// Layout ของ Ledger_Master (0-based index):
//
//	0=Comcode  1=Ac_code  2=Ac_name  3=Gcode  4=Gname
//	5=BBAL     6=CBAL     7=Debit    8=Credit
//	9=Bthisyear (ยอดยกมาต้นปี)
//	10-21 = Thisper01-12 (ยอดสุทธิ dr-cr แต่ละงวด คำนวณโดย RecalculateLedgerMaster)
//
// ข้อแม้: ต้องเรียก RecalculateLedgerMaster() ก่อนออกรายงาน
// ซึ่งระบบทำอยู่แล้วหลัง save voucher ทุกครั้ง
// และผู้ใช้สามารถกดปุ่ม "ประมวลผลยอด" ใน Ledger UI ได้ตลอดเวลา
// ─────────────────────────────────────────────────────────────────
func buildWorksheetData(xlOptions excelize.Options, repType int, glFrom, glTo string) ([]WorksheetRow, error) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)

	// currentPeriod = งวดปัจจุบัน (อ่านจาก Company_Profile.G2)
	// ใช้จำกัดการแสดงผล Thisper — งวดที่เกินยังไม่มีข้อมูล ไม่ควรแสดง
	currentPeriod := LoadCurrentPeriod(xlOptions)

	rows, _ := f.GetRows("Ledger_Master")

	var result []WorksheetRow

	for i, r := range rows {
		if i == 0 || len(r) < 4 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := safeGet(r, 1) // col B
		acName := safeGet(r, 2) // col C
		gCode := safeGet(r, 3)  // col D (Gcode 3 หลัก เช่น "120", "360")

		// ── Filter ช่วง Gcode ที่ user เลือก ──
		if gCode < glFrom || gCode > glTo {
			continue
		}

		// ── Filter ตาม Type ──
		// Type 1 = หมวด 1xx-3xx (สินทรัพย์ หนี้สิน ทุน)
		// Type 2 = หมวด 4xx ขึ้นไป (รายได้ ค่าใช้จ่าย)
		firstChar := ""
		if len(gCode) > 0 {
			firstChar = gCode[:1]
		}
		if repType == 1 && firstChar > "3" {
			continue
		}
		if repType == 2 && firstChar <= "3" {
			continue
		}

		rowObj := WorksheetRow{
			AcCode: acCode,
			AcName: acName,
		}

		// ── อ่านยอดจาก Ledger_Master โดยตรง ──
		// Bthisyear (index 9) = ยอดยกมาต้นปี (เฉพาะ Type 1 / หมวด 1-3)
		// Thisper01-12 (index 10-21) = ยอดสุทธิ (dr-cr) แต่ละงวด
		// ค่าเหล่านี้ถูกคำนวณและอัพเดทโดย RecalculateLedgerMaster() อัตโนมัติ
		// ครอบคลุม special codes (360PLA, 524RND, 450RND ฯลฯ) ทุกตัวโดยไม่ต้อง hardcode

		openingBal := parseFloat(safeGet(r, 9)) // Bthisyear

		// Thisper01-12 เก็บเป็น cumulative balance (ยอดสะสม)
		// ต้องแปลงกลับเป็นยอดเคลื่อนไหวรายงวด (net movement) ก่อนนำมาแสดง
		// สูตร: movement[p] = cumulative[p] - cumulative[p-1]
		//        movement[0] = cumulative[0] - Bthisyear
		// งวดที่เกิน currentPeriod ยังไม่มีข้อมูล → แสดง 0
		sumPeriods := 0.0
		prevCumulative := openingBal
		for p := 0; p < 12; p++ {
			if p >= currentPeriod {
				// งวดที่ยังไม่ถึง → ไม่มีข้อมูล
				rowObj.Periods[p] = 0
				continue
			}
			cumVal := parseFloat(safeGet(r, 10+p))
			movement := cumVal - prevCumulative
			rowObj.Periods[p] = movement
			sumPeriods += movement
			prevCumulative = cumVal
		}

		if repType == 1 {
			rowObj.Opening = openingBal
			rowObj.Total = openingBal + sumPeriods // ยอดยกไป
		} else {
			rowObj.Total = sumPeriods // ยอดรวม
		}

		// ── ซ่อนบรรทัดที่ไม่มียอดใดๆ เลย ──
		hasValue := rowObj.Opening != 0 || rowObj.Total != 0
		if !hasValue {
			for _, v := range rowObj.Periods {
				if v != 0 {
					hasValue = true
					break
				}
			}
		}

		if hasValue {
			result = append(result, rowObj)
		}
	}

	sortWorksheetRows(result)
	return result, nil
}

// sortWorksheetRows — เรียงตามหมวดบัญชีสากล: 1=สินทรัพย์, 2=หนี้สิน, 3=ทุน, 4=รายได้, 5=ค่าใช้จ่าย, 6+=อื่นๆ
// ภายในแต่ละหมวดเรียงตาม AcCode ตามลำดับ
func sortWorksheetRows(rows []WorksheetRow) {
	acGroup := func(acCode string) int {
		if len(acCode) == 0 {
			return 9
		}
		switch acCode[0] {
		case '1':
			return 1
		case '2':
			return 2
		case '3':
			return 3
		case '4':
			return 4
		case '5':
			return 5
		default:
			return 6
		}
	}
	for i := 1; i < len(rows); i++ {
		key := rows[i]
		gi := acGroup(key.AcCode)
		j := i - 1
		for j >= 0 {
			gj := acGroup(rows[j].AcCode)
			if gj > gi || (gj == gi && rows[j].AcCode > key.AcCode) {
				rows[j+1] = rows[j]
				j--
			} else {
				break
			}
		}
		rows[j+1] = key
	}
}

// ─────────────────────────────────────────────────────────────────
// Export Excel
// ─────────────────────────────────────────────────────────────────
func exportWorksheetExcel(xlOptions excelize.Options, repType int, glFrom, glTo, savePath string) (string, error) {
	data, err := buildWorksheetData(xlOptions, repType, glFrom, glTo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	f.Close()

	wb := excelize.NewFile()
	sh := "Worksheet"
	wb.SetSheetName("Sheet1", sh)

	// Styles
	boldC, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true, Size: 12, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "center"}})
	hdrStyle, _ := wb.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "center"},
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 1}},
	})
	txtStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}})
	numStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}, CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;""`)})
	totNumStyle, _ := wb.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"}, CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 6}}, // 6 = double line
	})

	// Header
	wb.MergeCell(sh, "A1", "O1")
	wb.SetCellValue(sh, "A1", comName)
	wb.SetCellStyle(sh, "A1", "A1", boldC)

	title := "กระดาษทำการ - สินทรัพย์ หนี้สิน ทุน"
	if repType == 2 {
		title = "กระดาษทำการ - รายได้ ค่าใช้จ่าย"
	}
	wb.MergeCell(sh, "A2", "O2")
	wb.SetCellValue(sh, "A2", title)
	wb.SetCellStyle(sh, "A2", "A2", boldC)

	// Table Headers
	headers := []string{"รหัสบัญชี", "ชื่อบัญชี"}
	if repType == 1 {
		headers = append(headers, "ยอดยกมา")
	}
	for i := 1; i <= 12; i++ {
		headers = append(headers, fmt.Sprintf("งวด %d", i))
	}
	if repType == 1 {
		headers = append(headers, "ยอดยกไป")
	} else {
		headers = append(headers, "รวม")
	}

	for i, h := range headers {
		col, _ := excelize.ColumnNumberToName(i + 1)
		wb.SetCellValue(sh, col+"4", h)
		wb.SetCellStyle(sh, col+"4", col+"4", hdrStyle)

		switch i {
		case 0:
			wb.SetColWidth(sh, col, col, 12)
		case 1:
			wb.SetColWidth(sh, col, col, 30)
		default:
			wb.SetColWidth(sh, col, col, 14)
		}
	}

	// Data Rows
	rowIdx := 5
	var sumOpen, sumTotal float64
	var sumP [12]float64

	for _, r := range data {
		rs := strconv.Itoa(rowIdx)
		wb.SetCellValue(sh, "A"+rs, r.AcCode)
		wb.SetCellValue(sh, "B"+rs, r.AcName)
		wb.SetCellStyle(sh, "A"+rs, "B"+rs, txtStyle)

		colOffset := 3
		if repType == 1 {
			colName, _ := excelize.ColumnNumberToName(colOffset)
			wb.SetCellFloat(sh, colName+rs, roundF(r.Opening), 2, 64)
			wb.SetCellStyle(sh, colName+rs, colName+rs, numStyle)
			sumOpen += r.Opening
			colOffset++
		}

		for p := 0; p < 12; p++ {
			colName, _ := excelize.ColumnNumberToName(colOffset + p)
			wb.SetCellFloat(sh, colName+rs, roundF(r.Periods[p]), 2, 64)
			wb.SetCellStyle(sh, colName+rs, colName+rs, numStyle)
			sumP[p] += r.Periods[p]
		}
		colOffset += 12

		colName, _ := excelize.ColumnNumberToName(colOffset)
		wb.SetCellFloat(sh, colName+rs, roundF(r.Total), 2, 64)
		wb.SetCellStyle(sh, colName+rs, colName+rs, numStyle)
		sumTotal += r.Total

		rowIdx++
	}

	// Footer Totals
	rs := strconv.Itoa(rowIdx)
	wb.SetCellValue(sh, "B"+rs, "รวม")
	wb.SetCellStyle(sh, "B"+rs, "B"+rs, hdrStyle)

	colOffset := 3
	if repType == 1 {
		colName, _ := excelize.ColumnNumberToName(colOffset)
		wb.SetCellFloat(sh, colName+rs, roundF(sumOpen), 2, 64)
		wb.SetCellStyle(sh, colName+rs, colName+rs, totNumStyle)
		colOffset++
	}
	for p := 0; p < 12; p++ {
		colName, _ := excelize.ColumnNumberToName(colOffset + p)
		wb.SetCellFloat(sh, colName+rs, roundF(sumP[p]), 2, 64)
		wb.SetCellStyle(sh, colName+rs, colName+rs, totNumStyle)
	}
	colOffset += 12
	colName, _ := excelize.ColumnNumberToName(colOffset)
	wb.SetCellFloat(sh, colName+rs, roundF(sumTotal), 2, 64)
	wb.SetCellStyle(sh, colName+rs, colName+rs, totNumStyle)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return wb.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// Export PDF (A4 Landscape)
// ─────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────
// Export PDF (รองรับ A4 และ B3 Landscape)
// ─────────────────────────────────────────────────────────────────
func exportWorksheetPDF(xlOptions excelize.Options, repType int, glFrom, glTo, savePath, paperSize string) (string, error) {
	data, err := buildWorksheetData(xlOptions, repType, glFrom, glTo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	f.Close()

	// Font Setup
	userFontsDir := filepath.Join(os.Getenv("LOCALAPPDATA"), "Microsoft", "Windows", "Fonts")
	sysFontsDir := filepath.Join(os.Getenv("WINDIR"), "Fonts")
	if sysFontsDir == "Fonts" || sysFontsDir == "" {
		sysFontsDir = filepath.Join("C:", "Windows", "Fonts")
	}
	fontsDir := userFontsDir
	if _, err := os.Stat(filepath.Join(userFontsDir, "Sarabun-Regular.ttf")); os.IsNotExist(err) {
		fontsDir = sysFontsDir
	}
	fontPath := ""
	for _, name := range []string{"Sarabun-Regular.ttf", "THSarabunNew.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "THSarabunNew Bold.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	pdf := gopdf.GoPdf{}

	// กำหนดขนาดกระดาษ
	var pageW, pageH float64
	if paperSize == "B3" {
		// B3 Landscape (1417.32 x 1000.63 pt)
		pageW = 1417.32
		pageH = 1000.63
		pdf.Start(gopdf.Config{PageSize: gopdf.Rect{W: pageW, H: pageH}, Unit: gopdf.UnitPT})
	} else {
		// A4 Landscape (841.89 x 595.28 pt)
		pageW = 841.89
		pageH = 595.28
		pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4Landscape, Unit: gopdf.UnitPT})
	}

	pdf.AddTTFFont("normal", fontPath)
	pdf.AddTTFFont("bold", boldPath)

	// ตั้งค่า Layout ตามขนาดกระดาษ
	var mL, mR, mT, mB, rowH, wCode, wName, fontSize float64
	if paperSize == "B3" {
		mL = 40.0
		mR = 40.0
		mT = 50.0
		mB = 50.0
		rowH = 18.0
		wCode = 60.0
		wName = 180.0
		fontSize = 9.0 // B3 กระดาษใหญ่ ใช้ฟอนต์ใหญ่ได้
	} else {
		mL = 20.0
		mR = 20.0
		mT = 30.0
		mB = 30.0
		rowH = 12.0
		wCode = 45.0
		wName = 100.0
		fontSize = 7.0 // A4 กระดาษเล็ก ต้องบีบฟอนต์
	}

	// คำนวณความกว้างคอลัมน์ตัวเลข (มี 14 คอลัมน์สำหรับ Type 1, 13 คอลัมน์สำหรับ Type 2)
	numCols := 14
	if repType == 2 {
		numCols = 13
	}
	wNum := (pageW - mL - mR - wCode - wName) / float64(numCols)

	pageNo := 0
	y := pageH

	title := "รายงานกระดาษทำการ - สินทรัพย์ หนี้สิน ทุน"
	if repType == 2 {
		title = "รายงานกระดาษทำการ - รายได้ ค่าใช้จ่าย"
	}

	printHeader := func() {
		pageNo++
		pdf.AddPage()
		y = mT

		pdf.SetFont("bold", "", fontSize+4)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - mL - mR, H: 15}, comName, gopdf.CellOption{Align: gopdf.Center})
		pdf.SetFont("normal", "", fontSize+1)
		pdf.SetXY(pageW-mR-50, y)
		pdf.CellWithOption(&gopdf.Rect{W: 50, H: 15}, fmt.Sprintf("หน้า %d", pageNo), gopdf.CellOption{Align: gopdf.Right})
		y += 15

		pdf.SetFont("normal", "", fontSize+3)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - mL - mR, H: 15}, title, gopdf.CellOption{Align: gopdf.Center})
		y += 20

		// Table Header
		pdf.SetLineWidth(0.5)
		pdf.Line(mL, y, pageW-mR, y)
		y += 3
		pdf.SetFont("bold", "", fontSize+1)

		curX := mL
		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wCode, H: rowH}, "รหัสบัญชี", gopdf.CellOption{})
		curX += wCode
		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wName, H: rowH}, "ชื่อบัญชี", gopdf.CellOption{})
		curX += wName

		if repType == 1 {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, "ยอดยกมา", gopdf.CellOption{Align: gopdf.Right})
			curX += wNum
		}
		for i := 1; i <= 12; i++ {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, fmt.Sprintf("งวด %d", i), gopdf.CellOption{Align: gopdf.Right})
			curX += wNum
		}
		if repType == 1 {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, "ยอดยกไป", gopdf.CellOption{Align: gopdf.Right})
		} else {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, "รวม", gopdf.CellOption{Align: gopdf.Right})
		}
		y += rowH
		pdf.Line(mL, y, pageW-mR, y)
		y += 3
		pdf.SetFont("normal", "", fontSize) // reset font หลัง header เสมอ ป้องกัน bold ล้นไปหน้าถัดไป
	}

	checkPage := func(need float64) {
		if y+need > pageH-mB {
			printHeader()
		}
	}

	printHeader()
	pdf.SetFont("normal", "", fontSize)

	var sumOpen, sumTotal float64
	var sumP [12]float64

	for _, r := range data {
		checkPage(rowH)
		curX := mL

		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wCode, H: rowH}, r.AcCode, gopdf.CellOption{})
		curX += wCode

		name := r.AcName
		maxLen := 25
		if paperSize == "B3" {
			maxLen = 45
		} // B3 แสดงชื่อได้ยาวขึ้น
		if len([]rune(name)) > maxLen {
			name = string([]rune(name)[:maxLen]) + ".."
		}
		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wName, H: rowH}, name, gopdf.CellOption{})
		curX += wName

		if repType == 1 {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(r.Opening), gopdf.CellOption{Align: gopdf.Right})
			curX += wNum
			sumOpen += r.Opening
		}

		for p := 0; p < 12; p++ {
			pdf.SetXY(curX, y)
			pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(r.Periods[p]), gopdf.CellOption{Align: gopdf.Right})
			curX += wNum
			sumP[p] += r.Periods[p]
		}

		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(r.Total), gopdf.CellOption{Align: gopdf.Right})
		sumTotal += r.Total

		y += rowH
	}

	// Footer
	checkPage(rowH + 10)
	pdf.Line(mL, y, pageW-mR, y)
	y += 3
	pdf.SetFont("bold", "", fontSize)

	curX := mL + wCode
	pdf.SetXY(curX, y)
	pdf.CellWithOption(&gopdf.Rect{W: wName, H: rowH}, "รวม", gopdf.CellOption{})
	curX += wName

	if repType == 1 {
		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(sumOpen), gopdf.CellOption{Align: gopdf.Right})
		curX += wNum
	}
	for p := 0; p < 12; p++ {
		pdf.SetXY(curX, y)
		pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(sumP[p]), gopdf.CellOption{Align: gopdf.Right})
		curX += wNum
	}
	pdf.SetXY(curX, y)
	pdf.CellWithOption(&gopdf.Rect{W: wNum, H: rowH}, formatNum(sumTotal), gopdf.CellOption{Align: gopdf.Right})

	y += rowH
	pdf.Line(mL, y, pageW-mR, y)
	y += 2
	pdf.Line(mL, y, pageW-mR, y)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// UI Dialog
// ─────────────────────────────────────────────────────────────────

func showWorksheetDialog(w fyne.Window, onGoSetup func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}

	reportDir := getReportDir(xlOptions)
	if strings.HasSuffix(filepath.ToSlash(reportDir), "/Desktop") || reportDir == filepath.ToSlash(filepath.Dir(currentDBPath)) {
		var warn dialog.Dialog
		btnGo := newEnterButton("ไปตั้งค่า (Enter)", func() {
			warn.Hide()
			if onGoSetup != nil {
				onGoSetup()
			}
		})
		btnCancel2 := newEscButton("ยกเลิก (Esc)", func() { warn.Hide() })
		warn = dialog.NewCustomWithoutButtons("⚠️ แจ้งเตือน", container.NewVBox(
			widget.NewLabel("กรุณาตั้งค่า Report Path ที่ Setup > Company Profile"),
			container.NewCenter(container.NewHBox(btnGo, btnCancel2)),
		), w)
		warn.Show()
		return
	}

	// ─── เตรียมตัวเลือกรหัสบัญชีจาก Acct_Group ───
	comCode := getComCodeFromExcel(xlOptions)
	var acOptsType1, acOptsType2 []string
	acMap := make(map[string]string)

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err == nil {
		rows, _ := f.GetRows("Acct_Group") // ดึงจาก Acct_Group ตามที่ผู้ใช้บอก
		for i, row := range rows {
			if i == 0 {
				continue
			}
			if len(row) >= 3 && row[0] == comCode {
				gCode := row[1]
				gName := row[2]
				if gCode != "" {
					displayTxt := fmt.Sprintf("%s | %s", gCode, gName)
					acMap[displayTxt] = gCode

					// แยกเก็บตามหมวด (Type 1 = 1-3, Type 2 = 4-6)
					firstChar := ""
					if len(gCode) > 0 {
						firstChar = gCode[:1]
					}

					if firstChar <= "3" {
						acOptsType1 = append(acOptsType1, displayTxt)
					} else {
						acOptsType2 = append(acOptsType2, displayTxt)
					}
				}
			}
		}
		f.Close()
	}

	if len(acOptsType1) == 0 {
		acOptsType1 = []string{"ไม่มีข้อมูลบัญชี"}
	}
	if len(acOptsType2) == 0 {
		acOptsType2 = []string{"ไม่มีข้อมูลบัญชี"}
	}

	cboAcFrom := widget.NewSelect(acOptsType1, nil)
	cboAcTo := widget.NewSelect(acOptsType1, nil)
	cboAcFrom.SetSelected(acOptsType1[0])
	cboAcTo.SetSelected(acOptsType1[len(acOptsType1)-1])

	// ─── Radio Report Type ───
	radioType := widget.NewRadioGroup([]string{
		"1 พิมพ์ในส่วน สินทรัพย์ หนี้สิน ทุน",
		"2 พิมพ์ในส่วน รายได้ ค่าใช้จ่าย",
	}, nil)
	radioType.Horizontal = false

	// เมื่อเปลี่ยน Type ให้เปลี่ยนตัวเลือกใน Dropdown ด้วย
	radioType.OnChanged = func(sel string) {
		if sel == "1 พิมพ์ในส่วน สินทรัพย์ หนี้สิน ทุน" {
			cboAcFrom.Options = acOptsType1
			cboAcTo.Options = acOptsType1
			cboAcFrom.SetSelected(acOptsType1[0])
			cboAcTo.SetSelected(acOptsType1[len(acOptsType1)-1])
		} else {
			cboAcFrom.Options = acOptsType2
			cboAcTo.Options = acOptsType2
			cboAcFrom.SetSelected(acOptsType2[0])
			cboAcTo.SetSelected(acOptsType2[len(acOptsType2)-1])
		}
	}
	radioType.SetSelected("1 พิมพ์ในส่วน สินทรัพย์ หนี้สิน ทุน")

	// ─── Radio Paper Size ───
	radioPaper := widget.NewRadioGroup([]string{"A4 Landscape", "B3 Landscape"}, nil)
	radioPaper.Horizontal = true
	radioPaper.SetSelected("A4 Landscape")

	radioLang := widget.NewRadioGroup([]string{"Thai", "English"}, nil)
	radioLang.Horizontal = true
	radioLang.SetSelected("Thai")

	var pop *widget.PopUp
	closePopup := func() {
		if pop != nil {
			pop.Hide()
		}
	}

	runExport := func(isPDF bool) {
		repType := 1
		if radioType.Selected == "2 พิมพ์ในส่วน รายได้ ค่าใช้จ่าย" {
			repType = 2
		}

		glFrom := acMap[cboAcFrom.Selected]
		glTo := acMap[cboAcTo.Selected]

		// ป้องกันกรณีผู้ใช้เลือกสลับกัน
		if glFrom > glTo {
			dialog.ShowError(fmt.Errorf("รหัสบัญชีเริ่มต้น ต้องน้อยกว่ารหัสบัญชีสิ้นสุด"), w)
			return
		}

		paperSize := "A4"
		if radioPaper.Selected == "B3 Landscape" {
			paperSize = "B3"
		}

		ext := ".xlsx"
		if isPDF {
			ext = ".pdf"
		}
		savePath := filepath.Join(reportDir, fmt.Sprintf("Worksheet_Type%d_%s%s", repType, paperSize, ext))
		closePopup()

		var pathToOpen string
		var err error
		if isPDF {
			pathToOpen, err = exportWorksheetPDF(xlOptions, repType, glFrom, glTo, savePath, paperSize)
		} else {
			pathToOpen, err = exportWorksheetExcel(xlOptions, repType, glFrom, glTo, savePath)
		}

		if err != nil {
			showErrDialog(w, "สร้างรายงานไม่ได้: "+err.Error())
		} else {
			isTmp := strings.HasSuffix(pathToOpen, ".tmp")
			if isTmp {
				showErrDialog(w, "เปิดรายงานชั่วคราว\nปิดไฟล์เดิมก่อน แล้วกด Export ใหม่เพื่อบันทึกถาวร")
			}
			openFile(pathToOpen)
		}
	}

	btnExcel := widget.NewButton("📊 Excel", func() { runExport(false) })
	btnExcel.Importance = widget.HighImportance
	btnPDF := widget.NewButton("📄 PDF", func() { runExport(true) })
	btnPDF.Importance = widget.HighImportance
	btnCancel := widget.NewButton("Cancel", closePopup)

	// ─── Layout ───
	form := container.NewVBox(
		container.NewPadded(radioType),
		widget.NewSeparator(),
		container.NewHBox(
			widget.NewLabelWithStyle("Code", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
			cboAcFrom,
			widget.NewLabelWithStyle("To", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			cboAcTo,
		),
		widget.NewSeparator(),
		container.NewHBox(
			layout.NewSpacer(),
			widget.NewLabelWithStyle("Paper Size (PDF)", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
			radioPaper,
			layout.NewSpacer(),
		),
		widget.NewSeparator(),
		container.NewHBox(
			layout.NewSpacer(),
			widget.NewLabelWithStyle("Report Language", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
			radioLang,
			layout.NewSpacer(),
		),
		widget.NewSeparator(),
		container.NewHBox(layout.NewSpacer(), btnExcel, btnPDF, btnCancel, layout.NewSpacer()),
	)

	pop = widget.NewModalPopUp(
		container.NewVBox(
			container.NewHBox(widget.NewLabelWithStyle("Select Period", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}), layout.NewSpacer(), widget.NewButton("X", closePopup)),
			widget.NewSeparator(),
			container.NewPadded(form),
		), w.Canvas(),
	)
	pop.Resize(fyne.NewSize(550, 350))
	pop.Show()
}
