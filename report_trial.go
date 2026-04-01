package main

// report_trial.go
// Trial Balance Report — งบทดลอง

import (
	"fmt"
	"io"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

type TrialRow struct {
	AcCode     string
	AcName     string
	BroughtFwd float64
	Debit      float64
	Credit     float64
	Balance    float64
}

func buildTrialBalance(xlOptions excelize.Options, periodNo int) ([]TrialRow, CompanyPeriodConfig, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, cfg, err
	}
	if periodNo < 1 || periodNo > cfg.TotalPeriods {
		return nil, cfg, fmt.Errorf("period %d ต้องอยู่ระหว่าง 1-%d", periodNo, cfg.TotalPeriods)
	}

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, cfg, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	rows, _ := f.GetRows("Ledger_Master")

	// Ledger_Master column layout (0-based):
	//   0=Comcode  1=Ac_code  2=Ac_name  3=Gcode  4=Gname
	//   5=BBAL     6=CBAL     7=Debit    8=Credit  9=Bthisyear
	//   10=Thisper01 .. 21=Thisper12
	//   22=Blastyear  23=Lastper01 .. 34=Lastper12
	//
	// FoxPro frm12/frm13 ตรรกะ:
	//   &I_PEMBEF = ยอดยกมา = Bthisyear + Thisper01..Thisper(n-1)
	//   DR        = Thisper01..Thisper(n)  รวมสะสมตั้งแต่ต้นปีถึงงวด n
	//               (แต่ FoxPro เก็บ Debit/Credit แยกสะสมรายงวดใน GLAC.DR/GLAC.CR
	//                ซึ่งใน Ledger_Master เราไม่มี field นี้โดยตรง)
	//   วิธีที่ถูก: DR_cumulative = sum(Thisper01..Thisper(n) ที่เป็นบวก)
	//              CR_cumulative = sum(Thisper01..Thisper(n) ที่เป็นลบ — abs)
	//   แต่จริงๆ แล้ว FoxPro เก็บ DR/CR แยกกันในฐานข้อมูล
	//   เราไม่มีข้อมูลนั้น → ใช้ทางเลือกที่ดีที่สุด:
	//     BroughtFwd = Bthisyear + sum(Thisper01..Thisper(n-1))  ← ยอดยกมางวดนี้
	//     Debit      = sum(Thisper(1..n) ที่ > 0)               ← ยอดเดบิตสะสมงวดนี้
	//     Credit     = abs(sum(Thisper(1..n) ที่ < 0))          ← ยอดเครดิตสะสมงวดนี้
	//     Balance    = BroughtFwd + Debit - Credit              ← CLOS ใน FoxPro
	//
	// หมายเหตุ: งบทดลองที่ถูกต้อง sumBAL ต้อง = 0 เสมอ
	//   เพราะ sumBAL = sum(BroughtFwd ทั้งหมด) + sumDR - sumCR
	//              = 0 + 0 = 0  (ถ้าบัญชีสมดุล)

	// FoxPro frm12/frm13:
	//   &I_PEMBEF = Bthisyear + Thisper01..Thisper(n-1)  ← ยอดยกมา
	//   DR        = GLAC.DR  ← gross debit สะสมทั้งปี
	//   CR        = GLAC.CR  ← gross credit สะสมทั้งปี
	//   CLOS      = BroughtFwd + DR - CR
	//
	// Ledger_Master col 7=Debit(YTD), col 8=Credit(YTD)
	// ตรงกับ GLAC.DR/GLAC.CR ที่ RecalculateLedgerMaster เขียนไว้
	//
	// 360PLA: FoxPro P_CLOSE.PRG ใส่ net P&L เข้า CLOS โดยตรง (ไม่ผ่าน DR/CR)
	// ดังนั้น 360PLA.DR=0, 360PLA.CR=0, CLOS = BroughtFwd + Thisper(n)
	// r_msum4 (sumBAL) จึงไม่เป็น 0 — FoxPro ยอมรับแบบนี้

	var result []TrialRow
	for _, row := range rows {
		if len(row) < 3 || row[0] != comCode {
			continue
		}
		acCode := safeGet(row, 1)
		acName := safeGet(row, 2)
		if acCode == "" {
			continue
		}

		// ยอดยกมา = Bthisyear + Thisper01..Thisper(n-1)
		bthis := parseFloat(safeGet(row, 9))
		var cumPrev float64
		for i := 0; i < periodNo-1; i++ {
			cumPrev += parseFloat(safeGet(row, 10+i))
		}
		broughtFwd := bthis + cumPrev

		// ── 360PLA: DR=0, CR=0, Balance = net P&L สะสม YTD ────────
		// BUG FIX: เดิมใช้แค่ Thisper(n) (งวดเดียว) ทำให้ sumBAL ≠ 0
		//   ต้องใช้ Σ Thisper(1..n) เพื่อให้ตรงกับยอดหมวด 4,5 ที่ใช้ DR/CR แบบ YTD
		//   FoxPro P_CLOSE.PRG ใส่ net P&L เข้า CLOS โดยตรง ซึ่งคือยอดสะสมทั้งปี
		if acCode == "360PLA" {
			var cumThisPer float64
			for i := 0; i < periodNo; i++ {
				cumThisPer += parseFloat(safeGet(row, 10+i))
			}
			balance := broughtFwd + cumThisPer
			if broughtFwd == 0 && cumThisPer == 0 {
				continue
			}
			result = append(result, TrialRow{
				AcCode:     acCode,
				AcName:     acName,
				BroughtFwd: broughtFwd,
				Debit:      0,
				Credit:     0,
				Balance:    balance,
			})
			continue
		}

		// gross DR/CR จาก col 7, 8 (YTD accumulated)
		dr := parseFloat(safeGet(row, 7))
		cr := parseFloat(safeGet(row, 8))
		balance := broughtFwd + dr - cr

		if broughtFwd == 0 && dr == 0 && cr == 0 {
			continue
		}
		result = append(result, TrialRow{
			AcCode:     acCode,
			AcName:     acName,
			BroughtFwd: broughtFwd,
			Debit:      dr,
			Credit:     cr,
			Balance:    balance,
		})
	}
	sortTrialRows(result)
	return result, cfg, nil
}

func getBookDRCR(f *excelize.File, comCode, acCode string, cfg CompanyPeriodConfig, periodNo int) (dr, cr float64) {
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if periodNo < 1 || periodNo > len(periods) {
		return
	}
	p := periods[periodNo-1]
	rows, _ := f.GetRows("Book_items")
	for i, row := range rows {
		if i == 0 || len(row) < 11 || row[0] != comCode {
			continue
		}
		if safeGet(row, 5) != acCode {
			continue
		}
		bdate := strings.TrimSpace(safeGet(row, 1))
		t, err := time.Parse("02/01/06", bdate)
		if err != nil {
			continue
		}
		if t.Before(p.PStart) || t.After(p.PEnd) {
			continue
		}
		dr += parseFloat(safeGet(row, 9))
		cr += parseFloat(safeGet(row, 10))
	}
	return
}

func roundF(v float64) float64 {
	return math.Round(v*100) / 100
}

// sortTrialRows — เรียงตามหมวดบัญชีสากล: 1=สินทรัพย์, 2=หนี้สิน, 3=ทุน, 4=รายได้, 5=ค่าใช้จ่าย, 6+=อื่นๆ
// ภายในแต่ละหมวดเรียงตาม AcCode ตามลำดับ lexicographic
func sortTrialRows(rows []TrialRow) {
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
	// insertion sort — stable, เหมาะกับข้อมูลไม่มาก (ledger accounts)
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
// exportTrialBalance — Excel
// ─────────────────────────────────────────────────────────────────
func exportTrialBalance(xlOptions excelize.Options, periodNo int, savePath string) (string, error) {
	rows, cfg, err := buildTrialBalance(xlOptions, periodNo)
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
	sh := "Trial Balance"
	wb.SetSheetName("Sheet1", sh)

	boldCenter, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 12, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	boldCenterSm, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	headerStyle, _ := wb.NewStyle(&excelize.Style{
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
	headerNumStyle, _ := wb.NewStyle(&excelize.Style{
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
		Font:      &excelize.Font{Size: 9, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	numStyleZero, _ := wb.NewStyle(&excelize.Style{
		Font:         &excelize.Font{Size: 9, Family: "TH Sarabun New"},
		Alignment:    &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
	})
	totalStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 9, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 6},
		},
		CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
	})

	wb.SetColWidth(sh, "A", "A", 14)
	wb.SetColWidth(sh, "B", "B", 45)
	wb.SetColWidth(sh, "C", "F", 18)

	wb.MergeCell(sh, "A1", "F1")
	wb.SetCellValue(sh, "A1", comName)
	wb.SetCellStyle(sh, "A1", "F1", boldCenter)
	wb.SetRowHeight(sh, 1, 22)

	wb.MergeCell(sh, "A2", "F2")
	wb.SetCellValue(sh, "A2", comAddr+"  เลขประจำตัวผู้เสียภาษีอากร "+comTax)
	addrStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 9, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
	})
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

	headers := []string{"รหัสบัญชี", "ชื่อบัญชี", "ยอดยกมา", "เดบิต", "เครดิต", "ยอดคงเหลือ"}
	cols := []string{"A", "B", "C", "D", "E", "F"}
	for i, h := range headers {
		cell := cols[i] + "7"
		wb.SetCellValue(sh, cell, h)
		if i >= 2 { // C-F: ตัวเลข → ชิดขวา
			wb.SetCellStyle(sh, cell, cell, headerNumStyle)
		} else {
			wb.SetCellStyle(sh, cell, cell, headerStyle)
		}
	}
	wb.SetRowHeight(sh, 7, 20)

	startRow := 8
	var sumBF, sumDR, sumCR, sumBAL float64
	for i, r := range rows {
		rowNum := startRow + i
		rowStr := strconv.Itoa(rowNum)
		wb.SetCellValue(sh, "A"+rowStr, r.AcCode)
		wb.SetCellValue(sh, "B"+rowStr, r.AcName)
		wb.SetCellFloat(sh, "C"+rowStr, roundF(r.BroughtFwd), 2, 64)
		wb.SetCellFloat(sh, "D"+rowStr, roundF(r.Debit), 2, 64)
		wb.SetCellFloat(sh, "E"+rowStr, roundF(r.Credit), 2, 64)
		wb.SetCellFloat(sh, "F"+rowStr, roundF(r.Balance), 2, 64)
		wb.SetCellStyle(sh, "A"+rowStr, "A"+rowStr, normalStyle)
		wb.SetCellStyle(sh, "B"+rowStr, "B"+rowStr, normalStyle)
		wb.SetCellStyle(sh, "C"+rowStr, "F"+rowStr, numStyleZero)
		wb.SetRowHeight(sh, rowNum, 18)
		// FoxPro frm12:
		//   r_msum1 += &I_PEMBEF  เฉพาะ code < '4' (สินทรัพย์/หนี้สิน/ทุน)
		//   r_msum2 += DR   ทุก account
		//   r_msum3 += CR   ทุก account
		//   r_msum4 += CLOS ทุก account → ต้อง = 0 เสมอ (งบดุล)

		if len(r.AcCode) > 0 && r.AcCode[0] < '4' {
			sumBF += r.BroughtFwd
		}
		sumDR += r.Debit
		sumCR += r.Credit
		// sumBAL += r.Balance // ถ้า account balance ทุก row ถูกต้อง sumBAL จะ = 0
		if r.AcCode != "360PLA" {
			sumBAL += r.Balance
		}
	}

	totalRow := startRow + len(rows)
	totalStr := strconv.Itoa(totalRow)
	wb.MergeCell(sh, "A"+totalStr, "B"+totalStr)
	wb.SetCellValue(sh, "A"+totalStr, "รวม (Total)")
	totalLabelStyle, _ := wb.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 6},
		},
	})
	wb.SetCellStyle(sh, "A"+totalStr, "B"+totalStr, totalLabelStyle)
	wb.SetCellFloat(sh, "C"+totalStr, roundF(sumBF), 2, 64)
	wb.SetCellFloat(sh, "D"+totalStr, roundF(sumDR), 2, 64)
	wb.SetCellFloat(sh, "E"+totalStr, roundF(sumCR), 2, 64)
	wb.SetCellFloat(sh, "F"+totalStr, roundF(sumBAL), 2, 64)
	wb.SetCellStyle(sh, "C"+totalStr, "F"+totalStr, totalStyle)
	wb.SetRowHeight(sh, totalRow, 22)

	wb.SetPageLayout(sh, &excelize.PageLayoutOptions{
		Orientation: strPtr("portrait"),
		Size:        intPtr(9), // 9 = A4
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

func strPtr(s string) *string { return &s }
func intPtr(i int) *int       { return &i }
func boolPtr(b bool) *bool    { return &b }

// ─────────────────────────────────────────────────────────────────
// exportTrialBalancePDF — PDF
// ─────────────────────────────────────────────────────────────────
func exportTrialBalancePDF(xlOptions excelize.Options, periodNo int, savePath string) (string, error) {
	rows, cfg, err := buildTrialBalance(xlOptions, periodNo)
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

	// หา font — per-user ก่อน แล้ว system
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
		entries, _ := os.ReadDir(fontsDir)
		var names []string
		for _, e := range entries {
			if strings.HasSuffix(strings.ToLower(e.Name()), ".ttf") {
				names = append(names, e.Name())
			}
		}
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s\nไฟล์ .ttf ที่มี: %s", fontsDir, strings.Join(names, ", "))
	}
	fmt.Printf("[PDF] font: %s\n", fontPath)

	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf", "arialbd.ttf", "ArialBD.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	pdf := gopdf.GoPdf{}
	// A4 Portrait: W=595.28, H=841.89
	pdfPageSize := gopdf.Rect{W: 595.28, H: 841.89}
	pdf.Start(gopdf.Config{PageSize: pdfPageSize, Unit: gopdf.UnitPT})
	if err = pdf.AddTTFFont("normal", fontPath); err != nil {
		return "", fmt.Errorf("โหลด font ไม่ได้: %v", err)
	}
	pdf.AddTTFFont("bold", boldPath)

	const (
		marginL  = 36.0
		marginR  = 36.0
		marginT  = 36.0
		pageW    = 595.28                    // A4 Portrait width
		pageH    = 841.89                    // A4 Portrait height
		contentW = pageW - marginL - marginR // 523.28
		rowH     = 15.0
	)
	var (
		codeW   = 55.0
		nameW   = 148.0
		numW    = (contentW - codeW - nameW) / 4.0 // ~75 each
		colCode = marginL
		colName = colCode + codeW
		colBF   = colName + nameW
		colDR   = colBF + numW
		colCR   = colDR + numW
		colBAL  = colCR + numW
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
		// ชื่อบริษัท
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
		pdf.SetFont("bold", "", 9)
		pdf.SetLineWidth(0.5)
		pdf.SetStrokeColor(0, 0, 0)
		pdf.Line(marginL, y, marginL+contentW, y)
		y += 2
		headerAligns := []int{gopdf.Center, gopdf.Left, gopdf.Right, gopdf.Right, gopdf.Right, gopdf.Right}
		for i, h := range []string{"รหัสบัญชี", "ชื่อบัญชี", "ยอดยกมา", "เดบิต", "เครดิต", "ยอดคงเหลือ"} {
			xCols := []float64{colCode, colName, colBF, colDR, colCR, colBAL}
			pdf.SetXY(xCols[i], y)
			pdf.CellWithOption(&gopdf.Rect{W: colWidths[i], H: rowH}, h, gopdf.CellOption{Align: headerAligns[i]})
		}
		y += rowH
		pdf.Line(marginL, y, marginL+contentW, y)
		// reset font กลับ normal ก่อนออกจาก newPage เสมอ
		pdf.SetFont("normal", "", 9)
	}

	headerH := marginT + 18.0 + 14.0 + 16.0 + 18.0 + 2.0 + rowH + 2.0
	currentY := headerH
	newPage()
	var sumBF, sumDR, sumCR, sumBAL float64
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
		for _, nc := range []struct{ x, v float64 }{{colBF, r.BroughtFwd}, {colDR, r.Debit}, {colCR, r.Credit}, {colBAL, r.Balance}} {
			pdf.SetXY(nc.x, y)
			pdf.CellWithOption(&gopdf.Rect{W: numW, H: rowH}, formatNum(nc.v), gopdf.CellOption{Align: gopdf.Right})
		}
		currentY += rowH
		// FoxPro frm12: r_msum1 += &I_PEMBEF เฉพาะ code < '4'
		if len(r.AcCode) > 0 && r.AcCode[0] < '4' {
			sumBF += r.BroughtFwd
		}
		sumDR += r.Debit
		sumCR += r.Credit
		// sumBAL += r.Balance
		if r.AcCode != "360PLA" {
			sumBAL += r.Balance
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
	for _, nc := range []struct{ x, v float64 }{{colBF, sumBF}, {colDR, sumDR}, {colCR, sumCR}, {colBAL, sumBAL}} {
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

// fileExists — ตรวจไฟล์มีอยู่จริงมั้ย
func fileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}

// formatNum — format ตัวเลข comma + 2 decimal, 0 = "-"
func formatNum(v float64) string {
	if v == 0 {
		return "0.00"
	}
	abs := v
	neg := ""
	if v < 0 {
		abs = -v
		neg = "-"
	}
	int64Part := int64(abs)
	decPart := int64(math.Round((abs - float64(int64Part)) * 100))
	if decPart == 100 {
		int64Part++
		decPart = 0
	}
	s := fmt.Sprintf("%d", int64Part)
	if len(s) > 3 {
		var out []byte
		for i, ch := range s {
			if i > 0 && (len(s)-i)%3 == 0 {
				out = append(out, ',')
			}
			out = append(out, byte(ch))
		}
		s = string(out)
	}
	return fmt.Sprintf("%s%s.%02d", neg, s, decPart)
}

// getReportDir — อ่าน Report Path จาก I2 หรือ fallback Desktop
func getReportDir(xlOptions excelize.Options) string {
	dir := ""
	if rf, err := excelize.OpenFile(currentDBPath, xlOptions); err == nil {
		dir, _ = rf.GetCellValue("Company_Profile", "I2")
		rf.Close()
	}
	dir = strings.TrimSpace(dir)
	if dir == "" {
		dir = filepath.Join(os.Getenv("USERPROFILE"), "Desktop")
	}
	if dir == "" {
		dir = filepath.Dir(currentDBPath)
	}
	return dir
}

// safeWriteFile — เขียนไฟล์รายงานแบบปลอดภัย
//
// กรณีปกติ (ไฟล์ไม่ถูกล็อก):
//
//	เขียน .tmp → ลบเก่า → rename เป็นไฟล์จริง → เปิดไฟล์จริง
//
// กรณีไฟล์ถูกเปิดอยู่ใน PDF reader:
//
//	เขียน .tmp → เปิด .tmp ให้ดูก่อนเลย (ไม่บล็อก user)
//	→ แจ้งว่ารายงานล่าสุดอยู่ใน .tmp
//	→ รอบหน้าที่ export ใหม่และไม่มีการล็อก จะ replace ไฟล์จริงอัตโนมัติ
//
// คืน (pathToOpen, error)
//
//	pathToOpen = path ที่ควรเปิดให้ user ดู (อาจเป็น .tmp หรือไฟล์จริง)
func safeWriteFile(savePath string, writeFn func(tmpPath string) error) (string, error) {
	// ต้องคง extension เดิม (.xlsx / .pdf) ไว้ใน tmpPath
	// เพราะ excelize และ gopdf ตรวจ extension ก่อน save
	// เช่น savePath = "TrialBalance_P01.xlsx" → tmpPath = "TrialBalance_P01_tmp.xlsx"
	ext := filepath.Ext(savePath)
	base := strings.TrimSuffix(savePath, ext)
	tmpPath := base + "_tmp" + ext

	// ลบ tmp เก่าถ้ามี (ไม่สน error)
	os.Remove(tmpPath)

	// เขียนรายงานลง tmp ก่อนเสมอ
	if err := writeFn(tmpPath); err != nil {
		os.Remove(tmpPath)
		return "", err
	}

	// ลองลบไฟล์เก่า — ถ้าถูกล็อกจะ error
	if err := os.Remove(savePath); err != nil && !os.IsNotExist(err) {
		// ไฟล์ถูกเปิดอยู่ → เปิด tmp ให้ดูก่อน ไม่บล็อก
		return tmpPath, nil
	}

	// rename tmp → ไฟล์จริง
	if err := os.Rename(tmpPath, savePath); err != nil {
		os.Remove(tmpPath)
		return "", fmt.Errorf("ไม่สามารถสร้างไฟล์รายงานได้: %v", err)
	}
	return savePath, nil
}

// openFile — เปิดไฟล์ด้วย default app ของ OS
// ถ้าเป็น .tmp ให้ copy เป็น _preview.<ext> ก่อนเปิด
// เพื่อให้ Windows/Mac รู้ว่าเปิดด้วยโปรแกรมอะไร
func openFile(path string) {
	cleanPath := filepath.Clean(filepath.FromSlash(path))

	// ถ้าเป็น _tmp.xlsx หรือ _tmp.pdf → copy เป็น _preview.<ext> ก่อนเปิด
	openPath := cleanPath
	ext := filepath.Ext(cleanPath)
	isTmp := strings.HasSuffix(strings.TrimSuffix(cleanPath, ext), "_tmp")
	if isTmp && ext != "" {
		base := strings.TrimSuffix(cleanPath, ext) // ลงท้าย _tmp
		baseNoTmp := strings.TrimSuffix(base, "_tmp")
		if ext != "" {
			previewPath := baseNoTmp + "_preview" + ext
			if err := func() error {
				src, err := os.Open(cleanPath)
				if err != nil {
					return err
				}
				defer src.Close()
				dst, err := os.Create(previewPath)
				if err != nil {
					return err
				}
				defer dst.Close()
				_, err = io.Copy(dst, src)
				return err
			}(); err == nil {
				openPath = previewPath
			}
		}
	}

	var cmd string
	var args []string
	switch runtime.GOOS {
	case "windows":
		cmd = "cmd"
		args = []string{"/c", "start", "", openPath}
	case "darwin":
		cmd = "open"
		args = []string{openPath}
	default:
		cmd = "xdg-open"
		args = []string{openPath}
	}
	if err := exec.Command(cmd, args...).Start(); err != nil {
		fmt.Printf("⚠️ เปิดไฟล์ไม่ได้: %v\n", err)
	}
}

// ─────────────────────────────────────────────────────────────────
// showTrialBalanceDialog — dialog เลือก period + Excel / PDF
// ─────────────────────────────────────────────────────────────────
func showTrialBalanceDialog(w fyne.Window, onGoSetup func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}

	// ── guard: ต้องตั้ง Report Path ก่อน ──
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
		widget.NewLabelWithStyle("งบทดลอง — Trial Balance", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
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

	// ESC จาก canvas level
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
		savePath := filepath.Join(reportDir, fmt.Sprintf("TrialBalance_P%02d.xlsx", periodNo))
		closePopup()
		w.Canvas().SetOnTypedKey(prevKey)
		pathToOpen, err := exportTrialBalance(xlOptions, periodNo, savePath)
		if err != nil {
			showErrDialog(w, "สร้าง Excel ไม่ได้: "+err.Error())
			return
		}
		showDone(pathToOpen)
	}
	btnPDF.OnTapped = func() {
		periodNo := getPeriod()
		savePath := filepath.Join(reportDir, fmt.Sprintf("TrialBalance_P%02d.pdf", periodNo))
		closePopup()
		w.Canvas().SetOnTypedKey(prevKey)
		pathToOpen, err := exportTrialBalancePDF(xlOptions, periodNo, savePath)
		if err != nil {
			showErrDialog(w, "สร้าง PDF ไม่ได้: "+err.Error())
			return
		}
		showDone(pathToOpen)
	}
}
