package main

// report_trial_v2.go
// Trial Balance V2 — รหัสบัญชี | ชื่อบัญชี | ยอดยกต้นงวด | ยอดยกมางวดนี้ | ยอดคงเหลือ DR | ยอดคงเหลือ CR

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// TrialRowV2 — 1 แถวใน Trial Balance V2
// รหัสบัญชี | ชื่อบัญชี | ยอดยกต้นงวด | ยอดยกมางวดนี้ | ยอดคงเหลือ DR | ยอดคงเหลือ CR
// ─────────────────────────────────────────────────────────────────
type TrialRowV2 struct {
	AcCode     string
	AcName     string
	Opening    float64 // ยอดยกมาต้นงวด = Bthisyear + P01..P(n-1)  ← สะสมก่อนงวดนี้
	BroughtFwd float64 // ยอดเคลื่อนไหวงวดนี้ = Thisper(n)
	BalDR      float64 // ยอดคงเหลือฝั่ง DR (balance > 0)
	BalCR      float64 // ยอดคงเหลือฝั่ง CR (balance < 0, เก็บเป็น abs)
}

// buildTrialBalanceV2 — คำนวณ Trial Balance V2
//
// Logic (เทียบกับ ledger_ui.go):
//
//	ยอดยกมาต้นงวด  = bbal  (Period Beginning = Bthisyear + Σ(DR-CR)[งวด 1..N-1])
//	ยอดยกมางวดนี้   = dr - cr  (สุทธิงวดปัจจุบัน จาก Book_items)
//	ยอดคงเหลือ DR/CR = cbal  (Period Closing = bbal + dr - cr)
//	                  > 0 → DR,  < 0 → CR (abs)
//
// Performance: เปิดไฟล์ครั้งเดียว โหลด Book_items + Ledger_Master ทั้งหมดเข้า
// memory แล้วคำนวณทุก account ใน loop เดียว — ไม่เปิดไฟล์ซ้ำต่อ account
func buildTrialBalanceV2(xlOptions excelize.Options, periodNo int) ([]TrialRowV2, CompanyPeriodConfig, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, cfg, err
	}
	if periodNo < 1 || periodNo > cfg.TotalPeriods {
		return nil, cfg, fmt.Errorf("period %d ต้องอยู่ระหว่าง 1-%d", periodNo, cfg.TotalPeriods)
	}
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)

	// ── เปิดไฟล์ครั้งเดียว ──────────────────────────────────────────
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, cfg, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)

	// ── Step 1: โหลด Book_items → map[acCode][periodIdx]{dr,cr} ──────
	type periodSum struct{ dr, cr float64 }
	acPeriods := make(map[string][]periodSum) // acCode → slice[TotalPeriods]

	bookRows, _ := f.GetRows("Book_items")
	for i, row := range bookRows {
		if i == 0 || len(row) < 11 || safeGet(row, 0) != comCode {
			continue
		}
		acCode := safeGet(row, 5)
		if acCode == "" {
			continue
		}
		dateStr := strings.TrimSpace(safeGet(row, 1))
		t, parseErr := parseSubbookDate(dateStr)
		if parseErr != nil {
			continue
		}
		dr := parseFloat(safeGet(row, 9))
		cr := parseFloat(safeGet(row, 10))

		// หา period index
		pIdx := -1
		for pi, p := range periods {
			if !t.Before(p.PStart) && !t.After(p.PEnd) {
				pIdx = pi
				break
			}
		}
		if pIdx < 0 {
			continue
		}

		// สะสมยอดของ account นี้
		if acPeriods[acCode] == nil {
			acPeriods[acCode] = make([]periodSum, cfg.TotalPeriods)
		}
		acPeriods[acCode][pIdx].dr += dr
		acPeriods[acCode][pIdx].cr += cr

		// 360PLA — สะสมยอดจากหมวด 4 และ 5 ทั้งหมด
		if len(acCode) > 0 && (acCode[0] == '4' || acCode[0] == '5') {
			if acPeriods["360PLA"] == nil {
				acPeriods["360PLA"] = make([]periodSum, cfg.TotalPeriods)
			}
			acPeriods["360PLA"][pIdx].dr += dr
			acPeriods["360PLA"][pIdx].cr += cr
		}
	}

	// ── Step 2: helper คำนวณ bbal/cbal/dr/cr จาก in-memory data ─────
	// ไม่เปิดไฟล์อีกเลย — ทุกอย่างอยู่ใน acPeriods แล้ว
	calcAccount := func(acCode string, bthisyear float64) (bbal, cbal, drN, crN float64) {
		sums := acPeriods[acCode]
		prev := bthisyear
		for p := 0; p < periodNo-1; p++ {
			if sums != nil {
				prev += sums[p].dr - sums[p].cr
			}
		}
		bbal = prev
		if sums != nil {
			drN = sums[periodNo-1].dr
			crN = sums[periodNo-1].cr
		}
		cbal = bbal + drN - crN
		return
	}

	// ── Step 3: วน Ledger_Master และ build result ────────────────────
	ledgerRows, _ := f.GetRows("Ledger_Master")
	var result []TrialRowV2

	for _, row := range ledgerRows {
		if len(row) < 3 || row[0] != comCode {
			continue
		}
		acCode := safeGet(row, 1)
		acName := safeGet(row, 2)
		if acCode == "" {
			continue
		}

		bthisyear := parseFloat(safeGet(row, 9)) // col J = Bthisyear

		// 360PLA — แสดง net P&L ของงวดนี้เท่านั้น (ไม่มี Opening/Closing)
		if acCode == "360PLA" {
			_, _, dr, cr := calcAccount(acCode, 0)
			net := dr - cr
			if net == 0 {
				continue
			}
			result = append(result, TrialRowV2{
				AcCode: acCode, AcName: acName,
				Opening: 0, BroughtFwd: net,
				BalDR: 0, BalCR: 0,
			})
			continue
		}

		bbal, cbal, dr, cr := calcAccount(acCode, bthisyear)
		opening := bbal
		thisPeriod := dr - cr
		balance := cbal

		if opening == 0 && thisPeriod == 0 {
			continue
		}

		var balDR, balCR float64
		if balance > 0 {
			balDR = balance
		} else if balance < 0 {
			balCR = -balance
		}

		result = append(result, TrialRowV2{
			AcCode:     acCode,
			AcName:     acName,
			Opening:    opening,
			BroughtFwd: thisPeriod,
			BalDR:      balDR,
			BalCR:      balCR,
		})
	}
	sortTrialRowsV2(result)
	return result, cfg, nil
}

// sortTrialRowsV2 — เรียงตามหมวดบัญชีสากล เหมือน sortTrialRows แต่สำหรับ TrialRowV2
func sortTrialRowsV2(rows []TrialRowV2) {
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
// exportTrialBalanceV2 — Excel V2
// ─────────────────────────────────────────────────────────────────
func exportTrialBalanceV2(xlOptions excelize.Options, periodNo int, savePath string) (string, error) {
	rows, cfg, err := buildTrialBalanceV2(xlOptions, periodNo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comAddr, _ := f.GetCellValue("Company_Profile", "C2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	pEnd := periods[periodNo-1].PEnd

	wb := excelize.NewFile()
	sh := "Trial Balance V2"
	wb.SetSheetName("Sheet1", sh)

	// ── Styles ──
	boldCenter, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 12, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	boldCenterSm, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	addrStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 9, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
	})
	hdrCenter, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"D9E1F2"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	hdrRight, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"D9E1F2"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	normalStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	numStyle, _ := wb.NewStyle(&excelize.Style{
		Font:         &excelize.Font{Size: 10, Family: "TH Sarabun New"},
		Alignment:    &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
	})
	totalStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 6},
		},
		CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
	})
	totalLabelStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 6},
		},
	})

	// ── Column widths: A=รหัส B=ชื่อ C=ยอดยกต้นงวด D=ยอดยกมางวดนี้ E=DR F=CR ──
	wb.SetColWidth(sh, "A", "A", 14)
	wb.SetColWidth(sh, "B", "B", 38)
	wb.SetColWidth(sh, "C", "F", 18)

	// ── Header rows ──
	wb.MergeCell(sh, "A1", "F1")
	wb.SetCellValue(sh, "A1", comName)
	wb.SetCellStyle(sh, "A1", "F1", boldCenter)
	wb.SetRowHeight(sh, 1, 22)

	wb.MergeCell(sh, "A2", "F2")
	wb.SetCellValue(sh, "A2", comAddr+"  เลขประจำตัวผู้เสียภาษีอากร "+comTax)
	wb.SetCellStyle(sh, "A2", "F2", addrStyle)
	wb.SetRowHeight(sh, 2, 30)
	wb.SetRowHeight(sh, 3, 6)

	wb.MergeCell(sh, "A4", "F4")
	wb.SetCellValue(sh, "A4", "งบทดลอง")
	wb.SetCellStyle(sh, "A4", "F4", boldCenter)
	wb.SetRowHeight(sh, 4, 20)

	wb.MergeCell(sh, "A5", "F5")
	wb.SetCellValue(sh, "A5", "ณ วันที่ "+pEnd.Format("02/01/06"))
	wb.SetCellStyle(sh, "A5", "F5", boldCenterSm)
	wb.SetRowHeight(sh, 5, 18)
	wb.SetRowHeight(sh, 6, 6)

	// ── Column headers ──
	headers := []struct {
		col   string
		label string
		style int
	}{
		{"A", "รหัสบัญชี", hdrCenter},
		{"B", "ชื่อบัญชี", hdrCenter},
		{"C", "ยอดยกมาต้นงวด", hdrRight},
		{"D", "ยอดงวดปัจจุบัน", hdrRight},
		{"E", "ยอดคงเหลือ/DR", hdrRight},
		{"F", "ยอดคงเหลือ/CR", hdrRight},
	}
	for _, h := range headers {
		cell := h.col + "7"
		wb.SetCellValue(sh, cell, h.label)
		wb.SetCellStyle(sh, cell, cell, h.style)
	}
	wb.SetRowHeight(sh, 7, 20)

	// ── Data rows ──
	startRow := 8
	var sumOpening, sumBF, sumBalDR, sumBalCR float64

	for i, r := range rows {
		rowNum := startRow + i
		rs := strconv.Itoa(rowNum)

		wb.SetCellValue(sh, "A"+rs, r.AcCode)
		wb.SetCellValue(sh, "B"+rs, r.AcName)
		wb.SetCellFloat(sh, "C"+rs, roundF(r.Opening), 2, 64)
		wb.SetCellFloat(sh, "D"+rs, roundF(r.BroughtFwd), 2, 64)
		wb.SetCellFloat(sh, "E"+rs, roundF(r.BalDR), 2, 64)
		wb.SetCellFloat(sh, "F"+rs, roundF(r.BalCR), 2, 64)

		wb.SetCellStyle(sh, "A"+rs, "A"+rs, normalStyle)
		wb.SetCellStyle(sh, "B"+rs, "B"+rs, normalStyle)
		wb.SetCellStyle(sh, "C"+rs, "F"+rs, numStyle)
		wb.SetRowHeight(sh, rowNum, 18)

		// sum ทุก account ยกเว้น 360PLA (double count กับ 4xx/5xx)
		if r.AcCode != "360PLA" {
			sumOpening += r.Opening
			sumBF += r.BroughtFwd
			sumBalDR += r.BalDR
			sumBalCR += r.BalCR
		}
	}

	// ── Total row ──
	totalRow := startRow + len(rows)
	ts := strconv.Itoa(totalRow)
	wb.MergeCell(sh, "A"+ts, "B"+ts)
	wb.SetCellValue(sh, "A"+ts, "รวม (Total)")
	wb.SetCellStyle(sh, "A"+ts, "B"+ts, totalLabelStyle)
	wb.SetCellFloat(sh, "C"+ts, roundF(sumOpening), 2, 64)
	wb.SetCellFloat(sh, "D"+ts, roundF(sumBF), 2, 64)
	wb.SetCellFloat(sh, "E"+ts, roundF(sumBalDR), 2, 64)
	wb.SetCellFloat(sh, "F"+ts, roundF(sumBalCR), 2, 64)
	wb.SetCellStyle(sh, "C"+ts, "F"+ts, totalStyle)
	wb.SetRowHeight(sh, totalRow, 22)

	// ── Page setup ──
	wb.SetPageLayout(sh, &excelize.PageLayoutOptions{
		Orientation: strPtr("portrait"),
		Size:        intPtr(9),
	})
	wb.SetHeaderFooter(sh, &excelize.HeaderFooterOptions{
		OddHeader: `&C&"TH Sarabun New,Bold"&10` + comName,
		OddFooter: `&Lพิมพ์วันที่: &D &T&R&"TH Sarabun New"&10หน้า &P / &N`,
	})
	wb.SetSheetProps(sh, &excelize.SheetPropsOptions{FitToPage: boolPtr(true)})

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return wb.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// exportTrialBalanceV2PDF — PDF V2
// ─────────────────────────────────────────────────────────────────
func exportTrialBalanceV2PDF(xlOptions excelize.Options, periodNo int, savePath string) (string, error) {
	rows, cfg, err := buildTrialBalanceV2(xlOptions, periodNo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comAddr, _ := f.GetCellValue("Company_Profile", "C2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	pEnd := periods[periodNo-1].PEnd

	// หา font (reuse logic เดิม)
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
	for _, name := range []string{"Sarabun-Regular.ttf", "Sarabun-Medium.ttf", "THSarabunNew.ttf", "arial.ttf", "Arial.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	if fontPath == "" {
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s", fontsDir)
	}
	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf", "arialbd.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	pdf := gopdf.GoPdf{}
	pdfPageSize := gopdf.Rect{W: 595.28, H: 841.89} // A4 Portrait
	pdf.Start(gopdf.Config{PageSize: pdfPageSize, Unit: gopdf.UnitPT})
	if err = pdf.AddTTFFont("normal", fontPath); err != nil {
		return "", fmt.Errorf("โหลด font ไม่ได้: %v", err)
	}
	pdf.AddTTFFont("bold", boldPath)

	const (
		marginL  = 36.0
		marginR  = 36.0
		marginT  = 36.0
		pageW    = 595.28
		pageH    = 841.89
		contentW = pageW - marginL - marginR // 523.28
		rowH     = 15.0
	)
	var (
		codeW   = 55.0
		nameW   = 148.0
		numW    = (contentW - codeW - nameW) / 4.0 // ~80 each
		colCode = marginL
		colName = colCode + codeW
		colOpen = colName + nameW
		colBF   = colOpen + numW
		colDR   = colBF + numW
		colCR   = colDR + numW
	)
	colWidths := []float64{codeW, nameW, numW, numW, numW, numW}

	pageNo := 0
	newPage := func() {
		pageNo++
		pdf.AddPageWithOption(gopdf.PageOption{PageSize: &gopdf.Rect{W: 595.28, H: 841.89}})
		y := marginT

		// page number — มุมขวาบน
		pdf.SetFont("normal", "", 9)
		pageStr := fmt.Sprintf("Page: %d", pageNo)
		pw, _ := pdf.MeasureTextWidth(pageStr)
		pdf.SetXY(marginL+contentW-pw, y)
		pdf.Cell(nil, pageStr)

		pdf.SetFont("bold", "", 13)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: contentW, H: 18}, comName, gopdf.CellOption{Align: gopdf.Center})
		y += 18

		pdf.SetFont("normal", "", 9)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: contentW, H: 12},
			comAddr+"  เลขประจำตัวผู้เสียภาษีอากร "+comTax, gopdf.CellOption{Align: gopdf.Center})
		y += 14

		pdf.SetFont("bold", "", 12)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: contentW, H: 16}, "งบทดลอง", gopdf.CellOption{Align: gopdf.Center})
		y += 16

		pdf.SetFont("normal", "", 10)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: contentW, H: 14},
			"ณ วันที่ "+pEnd.Format("02/01/2006"), gopdf.CellOption{Align: gopdf.Center})
		y += 18

		// column headers
		pdf.SetFont("bold", "", 9)
		pdf.SetLineWidth(0.5)
		pdf.SetStrokeColor(0, 0, 0)
		pdf.Line(marginL, y, marginL+contentW, y)
		y += 2
		hdrLabels := []string{"รหัสบัญชี", "ชื่อบัญชี", "ยอดยกมาต้นงวด", "ยอดงวดปัจจุบัน", "ยอดคงเหลือ/DR", "ยอดคงเหลือ/CR"}
		xCols := []float64{colCode, colName, colOpen, colBF, colDR, colCR}
		hdrAligns := []int{gopdf.Center, gopdf.Left, gopdf.Right, gopdf.Right, gopdf.Right, gopdf.Right}
		for i, h := range hdrLabels {
			pdf.SetXY(xCols[i], y)
			pdf.CellWithOption(&gopdf.Rect{W: colWidths[i], H: rowH}, h, gopdf.CellOption{Align: hdrAligns[i]})
		}
		y += rowH
		pdf.Line(marginL, y, marginL+contentW, y)
		// reset font กลับ normal ก่อนออกจาก newPage เสมอ
		pdf.SetFont("normal", "", 9)
	}

	headerH := marginT + 18.0 + 14.0 + 16.0 + 18.0 + 2.0 + rowH + 2.0
	currentY := headerH
	newPage()
	var sumOpen, sumBFTotal, sumBalDR, sumBalCR float64

	for _, r := range rows {
		if currentY+rowH > pageH-marginT {
			newPage()
			currentY = headerH
		}
		y := currentY

		pdf.SetXY(colCode, y)
		pdf.CellWithOption(&gopdf.Rect{W: codeW, H: rowH}, r.AcCode, gopdf.CellOption{})
		pdf.SetXY(colName, y)
		pdf.CellWithOption(&gopdf.Rect{W: nameW, H: rowH}, r.AcName, gopdf.CellOption{})

		for _, nc := range []struct{ x, v float64 }{
			{colOpen, r.Opening},
			{colBF, r.BroughtFwd},
			{colDR, r.BalDR},
			{colCR, r.BalCR},
		} {
			pdf.SetXY(nc.x, y)
			pdf.CellWithOption(&gopdf.Rect{W: numW, H: rowH}, formatNum(nc.v), gopdf.CellOption{Align: gopdf.Right})
		}
		currentY += rowH

		// ✅ Logic ใหม่: รวมเฉพาะงบดุล (Balance Sheet: หมวด 1, 2, 3)
		// แต่ต้องยกเว้น 360PLA (เพราะเป็นบัญชีกำไรขาดทุน)

		if r.AcCode != "360PLA" {
			sumOpen += r.Opening
			sumBFTotal += r.BroughtFwd
			sumBalDR += r.BalDR
			sumBalCR += r.BalCR
		}
	}

	// Total row
	y := currentY + 4
	pdf.SetLineWidth(0.5)
	pdf.Line(marginL, y, marginL+contentW, y)
	y += 4
	pdf.SetFont("bold", "", 9)
	pdf.SetXY(colCode, y)
	pdf.CellWithOption(&gopdf.Rect{W: codeW + nameW, H: rowH}, "รวม (Total)", gopdf.CellOption{Align: gopdf.Center})
	for _, nc := range []struct{ x, v float64 }{
		{colOpen, sumOpen}, {colBF, sumBFTotal}, {colDR, sumBalDR}, {colCR, sumBalCR},
	} {
		pdf.SetXY(nc.x, y)
		pdf.CellWithOption(&gopdf.Rect{W: numW, H: rowH}, formatNum(nc.v), gopdf.CellOption{Align: gopdf.Right})
	}
	y += rowH
	pdf.Line(marginL, y, marginL+contentW, y)
	y += 2
	pdf.Line(marginL, y, marginL+contentW, y)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// showTrialBalanceV2Dialog — dialog V2
// ─────────────────────────────────────────────────────────────────
func showTrialBalanceV2Dialog(w fyne.Window, onGoSetup func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}

	var opts []string
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	// แสดงเฉพาะ period 1 ถึง NowPeriod (ที่มีข้อมูลแล้ว)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	for _, p := range periods[:showUpTo] {
		opts = append(opts, fmt.Sprintf("งวด %d  (%s - %s)",
			p.PNo, p.PStart.Format("02/01/06"), p.PEnd.Format("02/01/06")))
	}
	cbo := widget.NewSelect(opts, nil)
	cbo.SetSelected(opts[showUpTo-1])

	getPeriod := func() int {
		for i, o := range opts {
			if o == cbo.Selected {
				return i + 1
			}
		}
		return cfg.NowPeriod
	}

	// guard: ต้องตั้ง Report Path ก่อน
	reportDir := getReportDir(xlOptions)
	if strings.HasSuffix(filepath.ToSlash(reportDir), "/Desktop") ||
		reportDir == filepath.ToSlash(filepath.Dir(currentDBPath)) {
		var warn dialog.Dialog
		btnGo := newEnterButton("ไปตั้งค่า (Enter)", func() {
			warn.Hide()
			if onGoSetup != nil {
				onGoSetup()
			}
		})
		btnGo.Importance = widget.HighImportance
		btnCancel2 := newEscButton("ยกเลิก (Esc)", func() { warn.Hide() })
		warn = dialog.NewCustomWithoutButtons(
			"⚠️  ยังไม่ได้ตั้งค่า Report Path",
			container.NewVBox(
				widget.NewLabel("กรุณาตั้งค่า Report Path ที่ Setup > Company Profile"),
				widget.NewLabel("เพื่อให้ไฟล์รายงานเก็บในที่เดียวกันทุกครั้ง"),
				widget.NewSeparator(),
				container.NewCenter(container.NewHBox(btnGo, btnCancel2)),
			), w)
		warn.Show()
		w.Canvas().Focus(btnGo)
		return
	}

	// showDone — แสดง success แล้วเปิดไฟล์เมื่อกด OK
	// ถ้า pathToOpen ลงท้าย .tmp แสดงว่าไฟล์เดิมถูกล็อกอยู่
	showDone := func(pathToOpen string) {
		ext2 := filepath.Ext(pathToOpen)
		isTmp := strings.HasSuffix(strings.TrimSuffix(pathToOpen, ext2), "_tmp")
		title := "✅ บันทึกรายงานแล้ว"
		note := ""
		if isTmp {
			title = "⚠️ เปิดรายงานชั่วคราว"
			note = "ปิดไฟล์ PDF เดิมก่อน แล้วกด Export ใหม่เพื่อบันทึกถาวร"
		}
		var done dialog.Dialog
		ok2 := newEnterButton("OK — เปิดไฟล์", func() {
			done.Hide()
			openFile(pathToOpen)
		})
		btnClose := newEscButton("ปิด", func() { done.Hide() })
		body := container.NewVBox(
			widget.NewLabel(title),
			widget.NewLabel(filepath.Base(pathToOpen)),
		)
		if note != "" {
			body.Add(widget.NewLabel(note))
		}
		body.Add(widget.NewSeparator())
		body.Add(container.NewCenter(container.NewHBox(ok2, btnClose)))
		done = dialog.NewCustomWithoutButtons("รายงาน", body, w)
		done.Show()
		w.Canvas().Focus(ok2)
	}

	var pop *widget.PopUp
	var prevKey func(*fyne.KeyEvent)
	closePopup := func() {
		if pop != nil {
			pop.Hide()
		}
		w.Canvas().SetOnTypedKey(prevKey)
	}

	btnExcel := widget.NewButton("📊 Excel (.xlsx)", nil)
	btnExcel.Importance = widget.HighImportance

	btnPDF := widget.NewButton("📄 PDF", nil)

	btnCancel := widget.NewButton("Cancel (Esc)", func() { closePopup() })

	content := container.NewVBox(
		widget.NewLabelWithStyle("งบทดลอง V2 — Trial Balance", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		widget.NewLabelWithStyle("เลือกงวดที่ต้องการ", fyne.TextAlignCenter, fyne.TextStyle{}),
		cbo,
		widget.NewSeparator(),
		container.NewCenter(container.NewHBox(btnExcel, btnPDF, btnCancel)),
	)
	pop = widget.NewModalPopUp(content, w.Canvas())
	pop.Resize(fyne.NewSize(420, 200))
	pop.Show()
	w.Canvas().Focus(nil)

	prevKey = w.Canvas().OnTypedKey()
	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		if key.Name == fyne.KeyEscape {
			closePopup()
			w.Canvas().SetOnTypedKey(prevKey)
			return
		}
		if prevKey != nil {
			prevKey(key)
		}
	})
	btnExcel.OnTapped = func() {
		periodNo := getPeriod()
		savePath := filepath.Join(reportDir, fmt.Sprintf("TrialBalance_V2_P%02d.xlsx", periodNo))
		closePopup()
		w.Canvas().SetOnTypedKey(prevKey)
		pathToOpen, err := exportTrialBalanceV2(xlOptions, periodNo, savePath)
		if err != nil {
			showErrDialog(w, "สร้าง Excel ไม่ได้: "+err.Error())
			return
		}
		showDone(pathToOpen)
	}
	btnPDF.OnTapped = func() {
		periodNo := getPeriod()
		savePath := filepath.Join(reportDir, fmt.Sprintf("TrialBalance_V2_P%02d.pdf", periodNo))
		closePopup()
		w.Canvas().SetOnTypedKey(prevKey)
		pathToOpen, err := exportTrialBalanceV2PDF(xlOptions, periodNo, savePath)
		if err != nil {
			showErrDialog(w, "สร้าง PDF ไม่ได้: "+err.Error())
			return
		}
		showDone(pathToOpen)
	}
}

// (computeRealTimeLedgerForPeriod ถูกแทนที่ด้วย in-memory calcAccount ใน buildTrialBalanceV2)
