package main

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// โครงสร้างข้อมูลสำหรับงบกำไรขาดทุน
// ─────────────────────────────────────────────────────────────────
type PnlData struct {
	ComName    string
	PeriodFrom int
	PeriodTo   int
	DateFrom   time.Time
	DateTo     time.Time

	// ปีปัจจุบัน (Current Year)
	RevSales      float64 // 400-410
	RevOther      float64 // 450
	TotalRevenues float64

	ExpCostOfSales float64 // 500-516
	ExpSelling     float64 // 520-524 ค่าใช้จ่ายในการขาย+บริหาร+อื่น
	ExpFinance     float64 // 530 ต้นทุนทางการเงิน
	ExpTax         float64 // 540 ภาษีเงินได้นิติบุคคล
	ExpSpecial     float64 // 550 รายการพิเศษ
	TotalExpenses  float64

	NetProfit float64 // กำไร(ขาดทุน)สุทธิ

	RetainedEarningsBegin float64 // กำไร(ขาดทุน)สะสมต้นปียกมา (350-399)
	RetainedEarningsEnd   float64 // กำไร(ขาดทุน)สะสมปลายปียกไป (Begin + NetProfit)

	EarningsPerShare float64 // กำไร(ขาดทุน)ต่อหุ้น

	// ปีที่แล้ว (Last Year) - ดึงจากคอลัมน์ Lastper
	LY_RevSales      float64
	LY_RevOther      float64
	LY_TotalRevenues float64

	LY_ExpCostOfSales float64
	LY_ExpSelling     float64
	LY_ExpFinance     float64
	LY_ExpTax         float64
	LY_ExpSpecial     float64
	LY_TotalExpenses  float64

	LY_NetProfit float64

	LY_RetainedEarningsBegin float64
	LY_RetainedEarningsEnd   float64

	LY_EarningsPerShare float64
}

type pnlLabels struct {
	Title                 string
	ForPeriod             string
	CurrentYear           string
	LastYear              string
	Revenues              string
	RevSales              string
	RevOther              string
	TotalRevenues         string
	Expenses              string
	ExpCostOfSales        string
	ExpSelling            string
	ExpFinance            string
	ExpTax                string
	ExpSpecial            string
	TotalExpenses         string
	NetProfit             string
	RetainedEarningsBegin string
	RetainedEarningsEnd   string
	EarningsPerShare      string
}

var pnlLabelsTH = pnlLabels{
	Title:                 "งบกำไรขาดทุน",
	ForPeriod:             "ตั้งแต่ %s ถึง %s",
	CurrentYear:           "ปีปัจจุบัน",
	LastYear:              "ปีที่แล้ว",
	Revenues:              "รายได้",
	RevSales:              "รายได้จากการขาย",
	RevOther:              "รายได้อื่นๆ",
	TotalRevenues:         "รวมรายได้",
	Expenses:              "ค่าใช้จ่าย",
	ExpCostOfSales:        "ต้นทุนสินค้าที่ขาย",
	ExpSelling:            "ค่าใช้จ่ายในการขายและบริหาร",
	ExpFinance:            "ต้นทุนทางการเงิน",
	ExpTax:                "ภาษีเงินได้นิติบุคคล",
	ExpSpecial:            "รายการพิเศษ",
	TotalExpenses:         "รวมค่าใช้จ่าย",
	NetProfit:             "กำไร(ขาดทุน)สุทธิ",
	RetainedEarningsBegin: "กำไร(ขาดทุน)สะสมต้นปียกมา",
	RetainedEarningsEnd:   "กำไร(ขาดทุน)สะสมปลายปียกไป",
	EarningsPerShare:      "กำไร(ขาดทุน)ต่อหุ้น",
}

var pnlLabelsEN = pnlLabels{
	Title:                 "Statement of Earnings",
	ForPeriod:             "From %s to %s",
	CurrentYear:           "Current Year",
	LastYear:              "Prior Year",
	Revenues:              "Revenues",
	RevSales:              "Sales Revenue",
	RevOther:              "Other Income",
	TotalRevenues:         "Total Revenues",
	Expenses:              "Expenses",
	ExpCostOfSales:        "Cost of Goods Sold",
	ExpSelling:            "Selling and Admin Expenses",
	ExpFinance:            "Finance Costs",
	ExpTax:                "Corporate Income Tax",
	ExpSpecial:            "Special Items",
	TotalExpenses:         "Total Expenses",
	NetProfit:             "Net Profit (Loss)",
	RetainedEarningsBegin: "Retained Earnings (Beginning)",
	RetainedEarningsEnd:   "Retained Earnings (Ending)",
	EarningsPerShare:      "Earnings Per Share",
}

// ─────────────────────────────────────────────────────────────────
// Core Logic: คำนวณงบกำไรขาดทุน (อิงตาม VB.NET)
// ─────────────────────────────────────────────────────────────────
func buildProfitAndLoss(xlOptions excelize.Options, prdFrom, prdTo int) (*PnlData, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, err
	}

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if prdFrom < 1 || prdFrom > len(periods) || prdTo < 1 || prdTo > len(periods) || prdFrom > prdTo {
		return nil, fmt.Errorf("ช่วง Period ไม่ถูกต้อง")
	}

	dateFrom := periods[prdFrom-1].PStart
	dateTo := periods[prdTo-1].PEnd

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")

	pnl := &PnlData{
		ComName:    comName,
		PeriodFrom: prdFrom,
		PeriodTo:   prdTo,
		DateFrom:   dateFrom,
		DateTo:     dateTo,
	}

	// 1. โหลด Transaction จาก Book_items
	type bItem struct {
		acCode string
		date   time.Time
		dr     float64
		cr     float64
	}
	var allTx []bItem
	bRows, _ := f.GetRows("Book_items")
	for i, r := range bRows {
		if i == 0 || len(r) < 11 || safeGet(r, 0) != comCode {
			continue
		}
		dtStr := safeGet(r, 1)
		dt, err := parseSubbookDate(dtStr)
		if err != nil {
			continue
		}
		allTx = append(allTx, bItem{
			acCode: safeGet(r, 5),
			date:   dt,
			dr:     parseFloat(safeGet(r, 9)),
			cr:     parseFloat(safeGet(r, 10)),
		})
	}

	// 2. โหลด Ledger_Master และคำนวณยอด
	lRows, _ := f.GetRows("Ledger_Master")
	for i, r := range lRows {
		if i == 0 || len(r) < 4 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := safeGet(r, 1)

		// ดึงหมวดบัญชี (3 ตัวแรก)
		prefix := acCode
		if len(acCode) >= 3 {
			prefix = acCode[:3]
		}

		pNum := 0
		fmt.Sscanf(prefix, "%d", &pNum)

		// --- คำนวณกำไรสะสมปลายปี (Retained Earnings End) ---
		// รวมหมวด 350-599 จนถึง dateTo เพื่อให้ได้กำไรสะสมปลายงวดที่ถูกต้องเสมอ
		if pNum >= 350 && pNum <= 599 {
			bThisYear := parseFloat(safeGet(r, 9))

			// ยอดยกมาต้นปี
			if pNum >= 350 && pNum <= 499 {
				pnl.RetainedEarningsEnd += -bThisYear // เครดิตเป็นบวก
			} else {
				pnl.RetainedEarningsEnd -= bThisYear // เดบิตเป็นลบ
			}

			// ยอดเคลื่อนไหวตั้งแต่ต้นปีจนถึง dateTo
			for _, tx := range allTx {
				if tx.acCode != acCode {
					continue
				}
				if !tx.date.After(dateTo) {
					if pNum >= 350 && pNum <= 499 {
						pnl.RetainedEarningsEnd += (tx.cr - tx.dr)
					} else {
						pnl.RetainedEarningsEnd -= (tx.dr - tx.cr)
					}
				}
			}
		}

		// งบกำไรขาดทุน สนใจเฉพาะหมวด 4 และ 5
		if pNum >= 400 && pNum <= 599 {
			// คำนวณยอดเคลื่อนไหวในช่วง Period ที่เลือก
			netAmount := 0.0
			if prdFrom == 1 {
				// ถ้างวดเริ่มเป็นงวด 1 ให้นำยอดยกมา (BTHISYEAR) มาบวกด้วย
				bThisYear := parseFloat(safeGet(r, 9))
				if prefix >= "400" && prefix <= "499" {
					netAmount += -bThisYear // รายได้: เครดิตเป็นบวก (ในระบบอาจเก็บเป็นลบ)
				} else {
					netAmount += bThisYear // ค่าใช้จ่าย: เดบิตเป็นบวก
				}
			}
			for _, tx := range allTx {
				if tx.acCode != acCode {
					continue
				}
				if !tx.date.Before(dateFrom) && !tx.date.After(dateTo) {
					if prefix >= "400" && prefix <= "499" {
						netAmount += (tx.cr - tx.dr) // รายได้: เครดิต - เดบิต
					} else {
						netAmount += (tx.dr - tx.cr) // ค่าใช้จ่าย: เดบิต - เครดิต
					}
				}
			}

			switch {
			case pNum >= 400 && pNum <= 410:
				pnl.RevSales += netAmount
			case pNum == 450:
				pnl.RevOther += netAmount
			case pNum >= 500 && pNum <= 516:
				pnl.ExpCostOfSales += netAmount
			case pNum >= 520 && pNum <= 524: // ค่าใช้จ่ายขาย+บริหาร+อื่น
				pnl.ExpSelling += netAmount
			case pNum == 530: // ต้นทุนทางการเงิน
				pnl.ExpFinance += netAmount
			case pNum == 540: // ภาษีเงินได้นิติบุคคล
				pnl.ExpTax += netAmount
			case pNum == 550: // รายการพิเศษ
				pnl.ExpSpecial += netAmount
			}
		}

		// --- คำนวณปีที่แล้ว (Last Year) ---
		// คำนวณกำไรสะสมปลายปีที่แล้ว (LY_RetainedEarningsEnd) โดยรวมหมวด 350-599 จาก BTHISYEAR
		if pNum >= 350 && pNum <= 599 {
			bThisYear := parseFloat(safeGet(r, 9))
			if pNum >= 350 && pNum <= 499 {
				pnl.LY_RetainedEarningsEnd += -bThisYear
			} else {
				pnl.LY_RetainedEarningsEnd -= bThisYear
			}
		}

		// ดึงจาก BTHISYEAR (ยอดยกมาต้นปี = ยอดสุทธิของปีที่แล้ว)
		lyAmount := 0.0
		if pNum >= 400 && pNum <= 599 {
			bThisYear := parseFloat(safeGet(r, 9))
			if prefix >= "400" && prefix <= "499" {
				lyAmount = -bThisYear // รายได้: เครดิตเป็นบวก
			} else {
				lyAmount = bThisYear // ค่าใช้จ่าย: เดบิตเป็นบวก
			}
		}

		if pNum >= 400 && pNum <= 599 {
			switch {
			case pNum >= 400 && pNum <= 410:
				pnl.LY_RevSales += lyAmount
			case pNum == 450:
				pnl.LY_RevOther += lyAmount
			case pNum >= 500 && pNum <= 516:
				pnl.LY_ExpCostOfSales += lyAmount
			case pNum >= 520 && pNum <= 524:
				pnl.LY_ExpSelling += lyAmount
			case pNum == 530:
				pnl.LY_ExpFinance += lyAmount
			case pNum == 540:
				pnl.LY_ExpTax += lyAmount
			case pNum == 550:
				pnl.LY_ExpSpecial += lyAmount
			}
		}
	}

	// 3. รวมยอดปีปัจจุบัน
	pnl.TotalRevenues = pnl.RevSales + pnl.RevOther
	pnl.TotalExpenses = pnl.ExpCostOfSales + pnl.ExpSelling + pnl.ExpFinance + pnl.ExpTax + pnl.ExpSpecial
	pnl.NetProfit = pnl.TotalRevenues - pnl.TotalExpenses
	pnl.RetainedEarningsBegin = pnl.RetainedEarningsEnd - pnl.NetProfit

	// ดึงจำนวนหุ้นจาก Capital (schema แนวนอน: A=Comcode B=ThisYearQty C=ThisYearValue D=LastYearQty E=LastYearValue)
	shares := 0.0
	capRows, _ := f.GetRows("Capital")
	for i, r := range capRows {
		if i == 0 || len(r) < 2 || safeGet(r, 0) != comCode {
			continue
		}
		shares = parseFloat(safeGet(r, 1))
		break
	}

	if shares > 0 {
		pnl.EarningsPerShare = pnl.NetProfit / shares
	}

	// 4. รวมยอดปีที่แล้ว
	pnl.LY_TotalRevenues = pnl.LY_RevSales + pnl.LY_RevOther
	pnl.LY_TotalExpenses = pnl.LY_ExpCostOfSales + pnl.LY_ExpSelling + pnl.LY_ExpFinance + pnl.LY_ExpTax + pnl.LY_ExpSpecial
	pnl.LY_NetProfit = pnl.LY_TotalRevenues - pnl.LY_TotalExpenses
	pnl.LY_RetainedEarningsBegin = pnl.LY_RetainedEarningsEnd - pnl.LY_NetProfit

	if shares > 0 {
		pnl.LY_EarningsPerShare = pnl.LY_NetProfit / shares
	}

	return pnl, nil
}

// ─────────────────────────────────────────────────────────────────
// Export to Excel
// ─────────────────────────────────────────────────────────────────
func exportPNLExcel(pnl *PnlData, lbl pnlLabels, savePath string) (string, error) {
	fx := excelize.NewFile()
	sn := "Statement of Earnings"
	fx.SetSheetName("Sheet1", sn)

	// ปิด grid lines ให้ดูสะอาดเหมือนต้นแบบ
	fx.SetSheetView(sn, 0, &excelize.ViewOptions{ShowGridLines: func() *bool { b := false; return &b }()})

	// column widths — เลียนแบบ Balance Sheet
	fx.SetColWidth(sn, "A", "A", 4)
	fx.SetColWidth(sn, "B", "S", 3)
	fx.SetColWidth(sn, "T", "T", 18) // ปีปัจจุบัน
	fx.SetColWidth(sn, "U", "U", 2)
	fx.SetColWidth(sn, "V", "V", 18) // ปีที่แล้ว
	fx.SetColWidth(sn, "W", "W", 2)

	// ── styles ──
	stCtrBold, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 12, Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	stCtr, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	stColHdr, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Border:    []excelize.Border{{Type: "bottom", Color: "000000", Style: 1}},
	})
	stSec, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "left"},
	})
	stLbl, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "left"},
	})
	stLblI, _ := fx.NewStyle(&excelize.Style{ // indent 1
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "left", Indent: 1},
	})
	stLblI2, _ := fx.NewStyle(&excelize.Style{ // indent 2
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "left", Indent: 2},
	})
	stNum, _ := fx.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})
	stNumU, _ := fx.NewStyle(&excelize.Style{ // single underline
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "right"},
		Border:    []excelize.Border{{Type: "bottom", Color: "000000", Style: 1}},
	})
	stNumD, _ := fx.NewStyle(&excelize.Style{ // double underline
		Font:      &excelize.Font{Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "right"},
		Border:    []excelize.Border{{Type: "bottom", Color: "000000", Style: 6}},
	})

	r := 1

	// helpers
	mc := func(c1, c2 string) {
		fx.MergeCell(sn, fmt.Sprintf("%s%d", c1, r), fmt.Sprintf("%s%d", c2, r))
	}
	set := func(col string, val interface{}, st int) {
		cell := fmt.Sprintf("%s%d", col, r)
		fx.SetCellValue(sn, cell, val)
		fx.SetCellStyle(sn, cell, cell, st)
	}
	num := func(col string, v float64, st int) {
		// Format ตัวเลขแบบมีวงเล็บถ้าติดลบ
		if math.Abs(v) < 0.005 {
			set(col, "0.00", st)
			return
		}
		neg := v < 0
		if neg {
			v = -v
		}
		intPart := int64(v)
		dec := int64(math.Round((v - float64(intPart)) * 100))
		if dec >= 100 {
			intPart++
			dec -= 100
		}
		s := fmt.Sprintf("%d", intPart)
		result := ""
		for i, ch := range s {
			if i > 0 && (len(s)-i)%3 == 0 {
				result += ","
			}
			result += string(ch)
		}
		out := fmt.Sprintf("%s.%02d", result, dec)
		if neg {
			set(col, "("+out+")", st)
		} else {
			set(col, out, st)
		}
	}
	br := func() { r++ } // blank row

	// ── Header ──
	br() // row 1 blank
	mc("E", "V")
	set("E", pnl.ComName, stCtrBold)
	br()
	mc("E", "V")
	set("E", lbl.Title, stCtrBold)
	br()
	mc("E", "V")
	set("E", fmt.Sprintf(lbl.ForPeriod, pnl.DateFrom.Format("02/01/06"), pnl.DateTo.Format("02/01/06")), stCtr)
	br()
	br() // blank

	// ── Column headers row ──
	set("T", lbl.CurrentYear, stColHdr)
	mc("V", "W")
	set("V", lbl.LastYear, stColHdr)
	br()

	// helpers ที่ใช้บ่อย
	secRow := func(col1, col2, label string) {
		mc(col1, col2)
		set(col1, label, stSec)
		br()
	}
	dataRow := func(labelCol1, labelCol2, label string, lSt int, cur, prev float64) {
		mc(labelCol1, labelCol2)
		set(labelCol1, label, lSt)
		num("T", cur, stNum)
		mc("V", "W")
		num("V", prev, stNum)
		br()
	}
	totRow := func(labelCol1, labelCol2, label string, cur, prev float64) {
		mc(labelCol1, labelCol2)
		set(labelCol1, label, stLblI2)
		num("T", cur, stNumU)
		mc("V", "W")
		num("V", prev, stNumU)
		br()
	}
	grandRow := func(labelCol1, labelCol2, label string, cur, prev float64) {
		mc(labelCol1, labelCol2)
		set(labelCol1, label, stLbl)
		num("T", cur, stNumD)
		mc("V", "W")
		num("V", prev, stNumD)
		br()
	}

	// ── รายได้ ──
	secRow("B", "H", lbl.Revenues)
	dataRow("C", "J", lbl.RevSales, stLblI, pnl.RevSales, pnl.LY_RevSales)
	dataRow("C", "J", lbl.RevOther, stLblI, pnl.RevOther, pnl.LY_RevOther)
	totRow("D", "M", lbl.TotalRevenues, pnl.TotalRevenues, pnl.LY_TotalRevenues)
	br()

	// ── ค่าใช้จ่าย ──
	secRow("B", "H", lbl.Expenses)
	dataRow("C", "J", lbl.ExpCostOfSales, stLblI, pnl.ExpCostOfSales, pnl.LY_ExpCostOfSales)
	dataRow("C", "J", lbl.ExpSelling, stLblI, pnl.ExpSelling, pnl.LY_ExpSelling)
	dataRow("C", "J", lbl.ExpFinance, stLblI, pnl.ExpFinance, pnl.LY_ExpFinance)
	dataRow("C", "J", lbl.ExpTax, stLblI, pnl.ExpTax, pnl.LY_ExpTax)
	if pnl.ExpSpecial != 0 || pnl.LY_ExpSpecial != 0 {
		dataRow("C", "J", lbl.ExpSpecial, stLblI, pnl.ExpSpecial, pnl.LY_ExpSpecial)
	}
	totRow("D", "M", lbl.TotalExpenses, pnl.TotalExpenses, pnl.LY_TotalExpenses)
	br()

	// ── กำไรสุทธิ ──
	grandRow("B", "H", lbl.NetProfit, pnl.NetProfit, pnl.LY_NetProfit)
	br()

	// ── กำไรสะสม ──
	dataRow("B", "J", lbl.RetainedEarningsBegin, stLbl, pnl.RetainedEarningsBegin, pnl.LY_RetainedEarningsBegin)
	dataRow("B", "J", lbl.NetProfit, stLbl, pnl.NetProfit, pnl.LY_NetProfit)
	grandRow("B", "J", lbl.RetainedEarningsEnd, pnl.RetainedEarningsEnd, pnl.LY_RetainedEarningsEnd)
	br()

	// ── กำไรต่อหุ้น ──
	grandRow("B", "H", lbl.EarningsPerShare, pnl.EarningsPerShare, pnl.LY_EarningsPerShare)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return fx.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// Export to PDF (ใช้ gopdf เหมือน Balance Sheet)
// ─────────────────────────────────────────────────────────────────
func exportPNLPDF(pnl *PnlData, lbl pnlLabels, savePath string) (string, error) {
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
	for _, name := range []string{"Sarabun-Regular.ttf", "Sarabun-Medium.ttf", "THSarabunNew.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	if fontPath == "" {
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s", fontsDir)
	}

	boldPath := fontPath // fallback = font เดิมถ้าไม่พบ bold
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}
	// ────────────────

	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})
	pdf.AddPage()
	if err := pdf.AddTTFFont("thai", fontPath); err != nil {
		return "", err
	}
	pdf.AddTTFFont("thai-bold", boldPath)

	const (
		lm    = 45.0  // left margin
		rm    = 550.0 // right margin
		colT  = 430.0 // ปีปัจจุบัน right edge
		colV  = 550.0 // ปีที่แล้ว right edge
		fs    = 9
		pageH = 810.0
	)
	y := 35.0

	nl := func(h float64) {
		y += h
		if y > pageH {
			pdf.AddPage()
			y = 35
		}
	}
	sf := func() { pdf.SetFont("thai", "", fs) }
	hln := func(x1, x2, yy, lw float64) {
		pdf.SetLineWidth(lw)
		pdf.Line(x1, yy, x2, yy)
	}

	printCtr := func(text string, size float64) {
		pdf.SetFont("thai", "", size)
		w, _ := pdf.MeasureTextWidth(text)
		pdf.SetXY((595-w)/2, y)
		pdf.Cell(nil, text)
		nl(size + 4)
	}

	// Format ตัวเลขแบบมีวงเล็บถ้าติดลบ
	bsNum := func(v float64) string {
		if math.Abs(v) < 0.005 {
			return "0.00"
		}
		neg := v < 0
		if neg {
			v = -v
		}
		intPart := int64(v)
		dec := int64(math.Round((v - float64(intPart)) * 100))
		if dec >= 100 {
			intPart++
			dec -= 100
		}
		s := fmt.Sprintf("%d", intPart)
		result := ""
		for i, ch := range s {
			if i > 0 && (len(s)-i)%3 == 0 {
				result += ","
			}
			result += string(ch)
		}
		out := fmt.Sprintf("%s.%02d", result, dec)
		if neg {
			return "(" + out + ")"
		}
		return out
	}

	// วาง 2 ตัวเลขชิดขวา col T และ col V
	putNums := func(cur, prev float64) {
		cs, ps := bsNum(cur), bsNum(prev)
		cw, _ := pdf.MeasureTextWidth(cs)
		pw, _ := pdf.MeasureTextWidth(ps)
		pdf.SetXY(colT-cw, y)
		pdf.Cell(nil, cs)
		pdf.SetXY(colV-pw, y)
		pdf.Cell(nil, ps)
	}

	printRow := func(label string, indent float64, cur, prev float64) {
		sf()
		pdf.SetXY(lm+indent, y)
		pdf.Cell(nil, label)
		putNums(cur, prev)
		nl(fs + 4)
	}

	printSec := func(label string) {
		sf()
		pdf.SetXY(lm, y)
		pdf.Cell(nil, label)
		nl(fs + 4)
	}

	printTot := func(label string, indent float64, cur, prev float64) {
		// single underline ก่อนตัวเลข
		hln(colT-120, colT+3, y, 0.3)
		hln(colV-120, colV+3, y, 0.3)
		nl(1)
		sf()
		pdf.SetXY(lm+indent, y)
		pdf.Cell(nil, label)
		putNums(cur, prev)
		nl(fs + 4)
	}

	printGrand := func(label string, cur, prev float64) {
		// single underline เหนือตัวเลข
		hln(colT-120, colT+3, y, 0.5)
		hln(colV-120, colV+3, y, 0.5)
		nl(1)
		sf()
		pdf.SetXY(lm, y)
		pdf.Cell(nil, label)
		putNums(cur, prev)
		nl(fs + 4)
		// double underline
		y1 := y
		hln(colT-120, colT+3, y1, 0.5)
		hln(colV-120, colV+3, y1, 0.5)
		y += 2
		y2 := y
		hln(colT-120, colT+3, y2, 0.5)
		hln(colV-120, colV+3, y2, 0.5)
		nl(6)
	}

	// ── Header ──
	// printCtr(pnl.ComName, 13)
	pdf.SetFont("thai-bold", "", 13)
	w, _ := pdf.MeasureTextWidth(pnl.ComName)
	pdf.SetXY((595-w)/2, y)
	pdf.Cell(nil, pnl.ComName)
	nl(13 + 4)
	printCtr(lbl.Title, 11)
	printCtr(fmt.Sprintf(lbl.ForPeriod, pnl.DateFrom.Format("02/01/06"), pnl.DateTo.Format("02/01/06")), 9)
	nl(6)

	// column header line
	sf()
	ch, _ := pdf.MeasureTextWidth(lbl.CurrentYear)
	ph, _ := pdf.MeasureTextWidth(lbl.LastYear)
	pdf.SetXY(colT-ch/2, y)
	pdf.Cell(nil, lbl.CurrentYear)
	pdf.SetXY(colV-ph, y)
	pdf.Cell(nil, lbl.LastYear)
	nl(fs + 3)

	// ── รายได้ ──
	printSec(lbl.Revenues)
	printRow(lbl.RevSales, 20, pnl.RevSales, pnl.LY_RevSales)
	printRow(lbl.RevOther, 20, pnl.RevOther, pnl.LY_RevOther)
	printTot(lbl.TotalRevenues, 40, pnl.TotalRevenues, pnl.LY_TotalRevenues)
	nl(4)

	// ── ค่าใช้จ่าย ──
	printSec(lbl.Expenses)
	printRow(lbl.ExpCostOfSales, 20, pnl.ExpCostOfSales, pnl.LY_ExpCostOfSales)
	printRow(lbl.ExpSelling, 20, pnl.ExpSelling, pnl.LY_ExpSelling)
	printRow(lbl.ExpFinance, 20, pnl.ExpFinance, pnl.LY_ExpFinance)
	printRow(lbl.ExpTax, 20, pnl.ExpTax, pnl.LY_ExpTax)
	if pnl.ExpSpecial != 0 || pnl.LY_ExpSpecial != 0 {
		printRow(lbl.ExpSpecial, 20, pnl.ExpSpecial, pnl.LY_ExpSpecial)
	}
	printTot(lbl.TotalExpenses, 40, pnl.TotalExpenses, pnl.LY_TotalExpenses)
	nl(4)

	// ── กำไรสุทธิ ──
	printGrand(lbl.NetProfit, pnl.NetProfit, pnl.LY_NetProfit)
	nl(4)

	// ── กำไรสะสม ──
	printRow(lbl.RetainedEarningsBegin, 0, pnl.RetainedEarningsBegin, pnl.LY_RetainedEarningsBegin)
	printRow(lbl.NetProfit, 0, pnl.NetProfit, pnl.LY_NetProfit)
	printGrand(lbl.RetainedEarningsEnd, pnl.RetainedEarningsEnd, pnl.LY_RetainedEarningsEnd)
	nl(4)

	// ── กำไรต่อหุ้น ──
	printGrand(lbl.EarningsPerShare, pnl.EarningsPerShare, pnl.LY_EarningsPerShare)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// UI Dialog
// ─────────────────────────────────────────────────────────────────
func showProfitAndLossDialog(w fyne.Window, onGoSetup func()) {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOpts)
	if err != nil {
		dialog.ShowError(err, w)
		return
	}

	reportDir := getReportDir(xlOpts)
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

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	options := make([]string, showUpTo)
	for i, p := range periods[:showUpTo] {
		options[i] = fmt.Sprintf("Period %d (%s)", i+1, p.PEnd.Format("02/01/06"))
	}

	selPeriodFrom := widget.NewSelect(options, nil)
	selPeriodFrom.SetSelectedIndex(0) // Default: Period 1

	selPeriodTo := widget.NewSelect(options, nil)
	selPeriodTo.SetSelectedIndex(showUpTo - 1) // Default: NowPeriod

	btnExcelTH := widget.NewButton("📊 Excel TH", nil)
	btnExcelEN := widget.NewButton("📊 Excel EN", nil)
	btnPDFTH := widget.NewButton("📄 PDF TH", nil)
	btnPDFEN := widget.NewButton("📄 PDF EN", nil)
	btnCancel := widget.NewButton("❌ ปิด", nil)

	var pop *widget.PopUp
	prevKey := w.Canvas().OnTypedKey()
	closePopup := func() {
		if pop != nil {
			pop.Hide()
		}
		w.Canvas().SetOnTypedKey(prevKey)
	}

	pop = widget.NewModalPopUp(
		container.NewVBox(
			widget.NewLabelWithStyle("งบกำไรขาดทุน / Statement of Earnings",
				fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			container.NewHBox(
				selPeriodFrom,
				widget.NewLabel(" To "),
				selPeriodTo,
			),
			widget.NewSeparator(),
			container.NewHBox(btnExcelTH, btnExcelEN, btnPDFTH, btnPDFEN, btnCancel),
		),
		w.Canvas(),
	)
	pop.Resize(fyne.NewSize(550, 180))

	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		if key.Name == fyne.KeyEscape {
			closePopup()
		}
	})
	btnCancel.OnTapped = closePopup

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

	run := func(isPDF bool, lbl pnlLabels, lang string) {
		prdFrom := selPeriodFrom.SelectedIndex() + 1
		prdTo := selPeriodTo.SelectedIndex() + 1

		if prdFrom < 1 || prdTo < 1 {
			dialog.ShowInformation("แจ้งเตือน", "กรุณาเลือก Period ให้ครบถ้วน", w)
			return
		}
		if prdFrom > prdTo {
			dialog.ShowInformation("แจ้งเตือน", "Period เริ่มต้นต้องไม่มากกว่า Period สิ้นสุด", w)
			return
		}

		closePopup()

		pnl, err := buildProfitAndLoss(xlOpts, prdFrom, prdTo)
		if err != nil {
			dialog.ShowError(err, w)
			return
		}

		var savePath string
		var pathToOpen string
		var exportErr error
		if isPDF {
			savePath = filepath.Join(reportDir, fmt.Sprintf("ProfitAndLoss_P%02d-P%02d_%s.pdf", prdFrom, prdTo, lang))
			pathToOpen, exportErr = exportPNLPDF(pnl, lbl, savePath)
		} else {
			savePath = filepath.Join(reportDir, fmt.Sprintf("ProfitAndLoss_P%02d-P%02d_%s.xlsx", prdFrom, prdTo, lang))
			pathToOpen, exportErr = exportPNLExcel(pnl, lbl, savePath)
		}
		if exportErr != nil {
			dialog.ShowError(exportErr, w)
			return
		}
		showDone(pathToOpen)
	}

	btnExcelTH.OnTapped = func() { run(false, pnlLabelsTH, "TH") }
	btnExcelEN.OnTapped = func() { run(false, pnlLabelsEN, "EN") }
	btnPDFTH.OnTapped = func() { run(true, pnlLabelsTH, "TH") }
	btnPDFEN.OnTapped = func() { run(true, pnlLabelsEN, "EN") }

	pop.Show()
	w.Canvas().Focus(nil)
}
