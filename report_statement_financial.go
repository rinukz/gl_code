package main

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
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

// ─────────────────────────────────────────────────────────────────
// Data structure
// ─────────────────────────────────────────────────────────────────

type BalanceSheet struct {
	ComName     string
	PeriodNo    int
	PeriodStart string
	PeriodEnd   string

	// Capital
	ThisYearQty   float64
	ThisYearValue float64
	LastYearQty   float64
	LastYearValue float64

	// สินทรัพย์หมุนเวียน
	Cash        float64 // 100-111
	ShortInvest float64 // 112
	Receivable  float64 // 113-115
	LoanST      float64 // 116
	Inventory   float64 // 117-119
	OtherCA     float64 // 120
	TotalCA     float64

	// สินทรัพย์ไม่หมุนเวียน
	DirLoan    float64 // 150
	LongInvest float64 // 160
	PPE        float64 // 170-180
	OtherNCA   float64 // 190
	TotalNCA   float64

	TotalAssets float64

	// หนี้สินหมุนเวียน
	BankOD      float64 // 200
	Payable     float64 // 210-215
	Dividend    float64 // 220
	CurrentLTL  float64 // 225
	CurrentLoan float64 // 230
	OtherCL     float64 // 235
	TotalCL     float64

	// หนี้สินไม่หมุนเวียน
	DirLiab  float64 // 250
	LTLiab   float64 // 260-265 เงินกู้ระยะยาวจากบริษัทในเครือ
	LTLoan   float64 // 270     เงินกู้ระยะยาว
	Pension  float64 // 275     เงินทุนเลี้ยงชีพ
	OtherNCL float64 // 280
	TotalNCL float64

	TotalLiab float64

	// ส่วนของผู้ถือหุ้น
	Capital      float64 // 300
	SharePremium float64 // 302
	RetainedEarn float64 // 350-360 → -1* (credit)
	TotalEquity  float64

	TotalLiabEquity float64

	// ปีก่อน (BTHISYEAR)
	Prev struct {
		Cash, ShortInvest, Receivable, LoanST, Inventory, OtherCA, TotalCA   float64
		DirLoan, LongInvest, PPE, OtherNCA, TotalNCA                         float64
		TotalAssets                                                          float64
		BankOD, Payable, Dividend, CurrentLTL, CurrentLoan, OtherCL, TotalCL float64
		DirLiab, LTLiab, LTLoan, Pension, OtherNCL, TotalNCL                 float64
		TotalLiab                                                            float64
		Capital, SharePremium, RetainedEarn, TotalEquity                     float64
		TotalLiabEquity                                                      float64
	}
}

// ─────────────────────────────────────────────────────────────────
// buildBalanceSheet
// ─────────────────────────────────────────────────────────────────
func buildBalanceSheet(xlOpts excelize.Options, periodNo int) (BalanceSheet, error) {
	bs := BalanceSheet{}
	cfg, err := loadCompanyPeriod(xlOpts)
	if err != nil {
		return bs, err
	}
	if periodNo < 1 || periodNo > cfg.TotalPeriods {
		return bs, fmt.Errorf("period %d ต้องอยู่ระหว่าง 1-%d", periodNo, cfg.TotalPeriods)
	}

	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err != nil {
		return bs, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOpts)
	bs.ComName, _ = f.GetCellValue("Company_Profile", "B2")
	bs.PeriodNo = periodNo

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if periodNo <= len(periods) {
		bs.PeriodStart = periods[0].PStart.Format("02/01/06")
		bs.PeriodEnd = periods[periodNo-1].PEnd.Format("02/01/06")
	}

	// อ่าน Capital sheet (schema แนวนอน: A=Comcode B=ThisYearQty C=ThisYearValue D=LastYearQty E=LastYearValue)
	capRows, _ := f.GetRows("Capital")
	for i, row := range capRows {
		if i == 0 || len(row) < 5 || row[0] != comCode {
			continue
		}
		bs.ThisYearQty = parseFloat(row[1])
		bs.ThisYearValue = parseFloat(row[2])
		bs.LastYearQty = parseFloat(row[3])
		bs.LastYearValue = parseFloat(row[4])
		break
	}

	// อ่าน Ledger_Master สร้าง acMap: ac3 → (cur, prev)
	type acVal struct{ cur, prev float64 }
	acMap := map[int]acVal{}
	rows, _ := f.GetRows("Ledger_Master")
	for _, row := range rows {
		if len(row) < 3 || safeGet(row, 0) != comCode {
			continue
		}
		acCode := strings.TrimSpace(safeGet(row, 1))
		if len(acCode) < 3 {
			continue
		}
		ac3, err2 := strconv.Atoi(acCode[:3])
		if err2 != nil {
			continue
		}
		bthis := parseFloat(safeGet(row, 9))
		dr, cr := getBSBookDRCR(f, comCode, acCode, cfg, periodNo)
		v := acMap[ac3]
		v.cur += bthis + dr - cr
		v.prev += bthis
		acMap[ac3] = v
	}

	sum := func(from, to int) (cur, prev float64) {
		for ac3, v := range acMap {
			if ac3 >= from && ac3 <= to {
				cur += v.cur
				prev += v.prev
			}
		}
		return
	}
	absS := func(from, to int) (cur, prev float64) {
		c, p := sum(from, to)
		return math.Abs(c), math.Abs(p)
	}

	// สินทรัพย์
	bs.Cash, bs.Prev.Cash = sum(100, 111)
	bs.ShortInvest, bs.Prev.ShortInvest = sum(112, 112)
	bs.Receivable, bs.Prev.Receivable = sum(113, 115)
	bs.LoanST, bs.Prev.LoanST = sum(116, 116)
	bs.Inventory, bs.Prev.Inventory = sum(117, 119)
	bs.OtherCA, bs.Prev.OtherCA = sum(120, 120)
	bs.TotalCA = bs.Cash + bs.ShortInvest + bs.Receivable + bs.LoanST + bs.Inventory + bs.OtherCA
	bs.Prev.TotalCA = bs.Prev.Cash + bs.Prev.ShortInvest + bs.Prev.Receivable + bs.Prev.LoanST + bs.Prev.Inventory + bs.Prev.OtherCA

	bs.DirLoan, bs.Prev.DirLoan = sum(150, 150)
	bs.LongInvest, bs.Prev.LongInvest = sum(160, 160)
	bs.PPE, bs.Prev.PPE = sum(170, 180)
	bs.OtherNCA, bs.Prev.OtherNCA = sum(190, 190)
	bs.TotalNCA = bs.DirLoan + bs.LongInvest + bs.PPE + bs.OtherNCA
	bs.Prev.TotalNCA = bs.Prev.DirLoan + bs.Prev.LongInvest + bs.Prev.PPE + bs.Prev.OtherNCA
	bs.TotalAssets = bs.TotalCA + bs.TotalNCA
	bs.Prev.TotalAssets = bs.Prev.TotalCA + bs.Prev.TotalNCA

	// หนี้สิน → ABS
	bs.BankOD, bs.Prev.BankOD = absS(200, 200)
	bs.Payable, bs.Prev.Payable = absS(210, 215)
	bs.Dividend, bs.Prev.Dividend = absS(220, 220)
	bs.CurrentLTL, bs.Prev.CurrentLTL = absS(225, 225)
	bs.CurrentLoan, bs.Prev.CurrentLoan = absS(230, 230)
	bs.OtherCL, bs.Prev.OtherCL = absS(235, 235)
	bs.TotalCL = bs.BankOD + bs.Payable + bs.Dividend + bs.CurrentLTL + bs.CurrentLoan + bs.OtherCL
	bs.Prev.TotalCL = bs.Prev.BankOD + bs.Prev.Payable + bs.Prev.Dividend + bs.Prev.CurrentLTL + bs.Prev.CurrentLoan + bs.Prev.OtherCL

	bs.DirLiab, bs.Prev.DirLiab = absS(250, 250)
	bs.LTLiab, bs.Prev.LTLiab = absS(260, 265)
	bs.LTLoan, bs.Prev.LTLoan = absS(275, 275) // เงินกู้ระยะยาว (VB: 2_275)
	bs.Pension, bs.Prev.Pension = absS(270, 270)
	bs.OtherNCL, bs.Prev.OtherNCL = absS(276, 280) // หนี้สินไม่หมุนเวียนอื่น (VB: 2_280)
	bs.TotalNCL = bs.DirLiab + bs.LTLiab + bs.LTLoan + bs.Pension + bs.OtherNCL
	bs.Prev.TotalNCL = bs.Prev.DirLiab + bs.Prev.LTLiab + bs.Prev.LTLoan + bs.Prev.Pension + bs.Prev.OtherNCL
	bs.TotalLiab = bs.TotalCL + bs.TotalNCL
	bs.Prev.TotalLiab = bs.Prev.TotalCL + bs.Prev.TotalNCL

	// ทุน
	c300, p300 := sum(300, 300)
	bs.Capital, bs.Prev.Capital = math.Abs(c300), math.Abs(p300)
	c302, p302 := sum(302, 302)
	bs.SharePremium, bs.Prev.SharePremium = math.Abs(c302), math.Abs(p302)
	// RetainedEarn = กำไร(ขาดทุน)สะสม (350-360) + กำไร(ขาดทุน)สุทธิปีนี้ (400-599)
	// สูตรตรงกับ VB: -1*(bthis+dr-cr ของ 350-360) + (-1)*(dr-cr ของ 400-599)
	// cPL = dr-cr ของ 400-599: บวก=ขาดทุน, ลบ=กำไร
	// RetainedEarn = -c350 - cPL
	c350, p350 := sum(350, 360)
	cPL, _ := sum(400, 599)
	bs.RetainedEarn = -c350 - cPL
	bs.Prev.RetainedEarn = -p350 // ปีก่อน = Bthisyear ของ 350-360 อย่างเดียว
	bs.TotalEquity = bs.Capital + bs.SharePremium + bs.RetainedEarn
	bs.Prev.TotalEquity = bs.Prev.Capital + bs.Prev.SharePremium + bs.Prev.RetainedEarn
	bs.TotalLiabEquity = bs.TotalLiab + bs.TotalEquity
	bs.Prev.TotalLiabEquity = bs.Prev.TotalLiab + bs.Prev.TotalEquity

	return bs, nil
}

// getBSBookDRCR — sum Debit/Credit ตั้งแต่ period 1 ถึง periodNo (cumulative)
func getBSBookDRCR(f *excelize.File, comCode, acCode string, cfg CompanyPeriodConfig, periodNo int) (dr, cr float64) {
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if periodNo < 1 || len(periods) == 0 {
		return
	}
	pStart := periods[0].PStart
	pEnd := periods[periodNo-1].PEnd
	bookRows, _ := f.GetRows("Book_items")
	for i, row := range bookRows {
		if i == 0 || len(row) < 11 || row[0] != comCode {
			continue
		}
		if safeGet(row, 5) != acCode {
			continue
		}
		t, err := time.Parse("02/01/06", strings.TrimSpace(safeGet(row, 1)))
		if err != nil || t.Before(pStart) || t.After(pEnd) {
			continue
		}
		dr += parseFloat(safeGet(row, 9))
		cr += parseFloat(safeGet(row, 10))
	}
	return
}

// bsNum — format ตัวเลข: วงเล็บถ้าลบ, 2 decimal, comma separator
func bsNum(v float64) string {
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

// ─────────────────────────────────────────────────────────────────
// bsLabels — label text สำหรับแต่ละภาษา
// ─────────────────────────────────────────────────────────────────
type bsLabels struct {
	// header
	Title     string // งบแสดงฐานะการเงิน
	DateFmt   string // ตั้งแต่ %s ถึง %s
	AssetHdr  string // สินทรัพย์
	LiabEqHdr string // หนี้สินและส่วนของผู้ถือหุ้น
	ColCur    string // ปีปัจจุบัน
	ColPrev   string // ปีที่แล้ว
	// CA
	SecCA       string
	Cash        string
	ShortInvest string
	Receivable  string
	LoanST      string
	Inventory   string
	OtherCA     string
	TotalCA     string
	// NCA
	SecNCA      string
	DirLoan     string
	LongInvest  string
	PPE         string
	OtherNCA    string
	TotalNCA    string
	TotalAssets string
	// CL
	SecCL       string
	BankOD      string
	Payable     string
	Dividend    string
	CurrentLTL  string
	CurrentLoan string
	OtherCL     string
	TotalCL     string
	// NCL
	SecNCL    string
	DirLiab   string
	LTLiab    string
	Pension   string
	LTLoan    string
	OtherNCL  string
	TotalNCL  string
	TotalLiab string
	// Equity
	SecEquity    string
	SecCapital   string
	ShareQtyLbl  string
	ShareQtyUnit string
	ShareValLbl  string
	ShareValUnit string
	PaidCapital  string
	SharePremium string
	RetainedEarn string
	TotalEquity  string
	TotalLiabEq  string
}

var bsLabelsTH = bsLabels{
	Title: "งบแสดงฐานะการเงิน", DateFmt: "ตั้งแต่  %s  ถึง  %s",
	AssetHdr: "สินทรัพย์", LiabEqHdr: "หนี้สินและส่วนของผู้ถือหุ้น",
	ColCur: "ปีปัจจุบัน", ColPrev: "ปีที่แล้ว",
	SecCA: "สินทรัพย์หมุนเวียน",
	Cash:  "เงินสดและรายการเทียบเท่าเงินสด", ShortInvest: "เงินลงทุนระยะสั้น",
	Receivable: "ลูกหนี้การค้าและตั๋วเงินรับ - สุทธิ", LoanST: "เงินให้กู้ยืมระยะสั้นจากบริษัทในเครือ",
	Inventory: "สินค้าคงเหลือ", OtherCA: "สินทรัพย์หมุนเวียนอื่น", TotalCA: "รวมสินทรัพย์หมุนเวียน",
	SecNCA:  "สินทรัพย์ไม่หมุนเวียน",
	DirLoan: "ลูกหนี้และเงินให้กู้ยืมแก่กรรมการและลูกจ้าง", LongInvest: "เงินลงทุนระยะยาว",
	PPE: "ที่ดิน อาคารและอุปกรณ์ สุทธิ", OtherNCA: "สินทรัพย์ไม่หมุนเวียนอื่น",
	TotalNCA: "รวมสินทรัพย์ไม่หมุนเวียน", TotalAssets: "รวมสินทรัพย์",
	SecCL:  "หนี้สินหมุนเวียน",
	BankOD: "เงินเบิกเกินบัญชีและเงินกู้ยืมจากสถาบันการเงิน", Payable: "เจ้าหนี้การค้าและตั๋วเงินจ่าย",
	Dividend: "เงินปันผลค้างจ่าย", CurrentLTL: "หนี้สินระยะยาวถึงกำหนดชำระ",
	CurrentLoan: "เงินกู้ยืมระยะสั้นจากบริษัทในเครือ", OtherCL: "หนี้สินหมุนเวียนอื่น",
	TotalCL: "รวมหนี้สินหมุนเวียน",
	SecNCL:  "หนี้สินไม่หมุนเวียน",
	DirLiab: "เจ้าหนี้และเงินกู้ยืมจากกรรมการ", LTLiab: "เงินกู้ระยะยาวจากบริษัทในเครือ",
	Pension: "เงินทุนเลี้ยงชีพและบำเหน็จ", LTLoan: "เงินกู้ระยะยาว",
	OtherNCL: "หนี้สินไม่หมุนเวียนอื่น", TotalNCL: "รวมหนี้สินไม่หมุนเวียน",
	TotalLiab: "รวมหนี้สิน",
	SecEquity: "ส่วนของผู้ถือหุ้น", SecCapital: "ทุนเรือนหุ้น",
	ShareQtyLbl: "หุ้นสามัญจำนวน", ShareQtyUnit: "หุ้น",
	ShareValLbl: "มูลค่าหุ้นละ", ShareValUnit: "บาท",
	PaidCapital: "ทุนที่ออกและเรียกชำระแล้ว", SharePremium: "ส่วนเกินมูลค่าหุ้น",
	RetainedEarn: "กำไร(ขาดทุน)สะสม", TotalEquity: "รวมส่วนของผู้ถือหุ้น",
	TotalLiabEq: "รวมหนี้สินและส่วนของผู้ถือหุ้น",
}

var bsLabelsEN = bsLabels{
	Title: "Statement of Financial Position", DateFmt: "From  %s  To  %s",
	AssetHdr: "Assets", LiabEqHdr: "Liabilities and Shareholders' Equity",
	ColCur: "Current Year", ColPrev: "Prior Year",
	SecCA: "Current Assets",
	Cash:  "Cash and Cash Equivalents", ShortInvest: "Short-term Investments",
	Receivable: "Trade and Notes Receivable - Net", LoanST: "Short-term Loans to Related Companies",
	Inventory: "Inventories", OtherCA: "Other Current Assets", TotalCA: "Total Current Assets",
	SecNCA:  "Non-Current Assets",
	DirLoan: "Loans to Directors and Employees", LongInvest: "Long-term Investments",
	PPE: "Property, Plant and Equipment - Net", OtherNCA: "Other Non-Current Assets",
	TotalNCA: "Total Non-Current Assets", TotalAssets: "Total Assets",
	SecCL:  "Current Liabilities",
	BankOD: "Bank Overdrafts and Short-term Borrowings", Payable: "Trade and Notes Payable",
	Dividend: "Dividends Payable", CurrentLTL: "Current Portion of Long-term Liabilities",
	CurrentLoan: "Short-term Loans from Related Companies", OtherCL: "Other Current Liabilities",
	TotalCL: "Total Current Liabilities",
	SecNCL:  "Non-Current Liabilities",
	DirLiab: "Loans from Directors", LTLiab: "Long-term Loans from Related Companies",
	Pension: "Pension and Retirement Fund", LTLoan: "Long-term Loans",
	OtherNCL: "Other Non-Current Liabilities", TotalNCL: "Total Non-Current Liabilities",
	TotalLiab: "Total Liabilities",
	SecEquity: "Shareholders' Equity", SecCapital: "Share Capital",
	ShareQtyLbl: "Ordinary Shares", ShareQtyUnit: "shares",
	ShareValLbl: "Par value", ShareValUnit: "Baht",
	PaidCapital: "Issued and Fully Paid-up Capital", SharePremium: "Share Premium",
	RetainedEarn: "Retained Earnings (Deficit)", TotalEquity: "Total Shareholders' Equity",
	TotalLiabEq: "Total Liabilities and Shareholders' Equity",
}

// ─────────────────────────────────────────────────────────────────
// Excel export — structure ตรงกับ balance_sheet.xlsx ต้นแบบ
// cols: A(narrow) B-S(labels) T(ปีปัจจุบัน) U(narrow) V-W(ปีที่แล้ว)
// ─────────────────────────────────────────────────────────────────
func exportBalanceSheetExcel(bs BalanceSheet, lbl bsLabels, savePath string) (string, error) {
	fx := excelize.NewFile()
	sn := "Balance Sheet"
	fx.SetSheetName("Sheet1", sn)

	// column widths — เลียนแบบต้นแบบ
	// ปิด grid lines ให้ดูสะอาดเหมือนต้นแบบ
	fx.SetSheetView(sn, 0, &excelize.ViewOptions{ShowGridLines: func() *bool { b := false; return &b }()})

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
	stSec, _ := fx.NewStyle(&excelize.Style{ // section header underline
		Font:      &excelize.Font{Size: 10, Underline: "single"},
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
		Font:      &excelize.Font{Size: 10, Bold: true},
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
	num := func(col string, v float64, st int) { set(col, bsNum(v), st) }
	br := func() { r++ } // blank row

	// ── Header ──
	br() // row 1 blank
	mc("E", "V")
	set("E", bs.ComName, stCtrBold)
	br()
	mc("E", "V")
	set("E", lbl.Title, stCtrBold)
	br()
	mc("E", "V")
	set("E", fmt.Sprintf(lbl.DateFmt, bs.PeriodStart, bs.PeriodEnd), stCtr)
	br()
	br() // blank

	// ── Column headers row ──
	mc("B", "H")
	set("B", lbl.AssetHdr, stSec)
	set("T", lbl.ColCur, stColHdr)
	mc("V", "W")
	set("V", lbl.ColPrev, stColHdr)
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
		set(labelCol1, label, stLbl)
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

	// ── สินทรัพย์หมุนเวียน ──
	secRow("B", "H", lbl.SecCA)
	dataRow("C", "J", lbl.Cash, stLblI, bs.Cash, bs.Prev.Cash)
	dataRow("C", "J", lbl.ShortInvest, stLblI, bs.ShortInvest, bs.Prev.ShortInvest)
	dataRow("C", "J", lbl.Receivable, stLblI, bs.Receivable, bs.Prev.Receivable)
	dataRow("C", "O", lbl.LoanST, stLblI, bs.LoanST, bs.Prev.LoanST)
	dataRow("C", "J", lbl.Inventory, stLblI, bs.Inventory, bs.Prev.Inventory)
	dataRow("C", "J", lbl.OtherCA, stLblI, bs.OtherCA, bs.Prev.OtherCA)
	totRow("D", "M", lbl.TotalCA, bs.TotalCA, bs.Prev.TotalCA)
	br()

	// ── สินทรัพย์ไม่หมุนเวียน ──
	secRow("B", "H", lbl.SecNCA)
	dataRow("C", "R", lbl.DirLoan, stLblI, bs.DirLoan, bs.Prev.DirLoan)
	dataRow("C", "J", lbl.LongInvest, stLblI, bs.LongInvest, bs.Prev.LongInvest)
	dataRow("C", "J", lbl.PPE, stLblI, bs.PPE, bs.Prev.PPE)
	dataRow("C", "J", lbl.OtherNCA, stLblI, bs.OtherNCA, bs.Prev.OtherNCA)
	totRow("D", "M", lbl.TotalNCA, bs.TotalNCA, bs.Prev.TotalNCA)
	br()
	grandRow("B", "H", lbl.TotalAssets, bs.TotalAssets, bs.Prev.TotalAssets)
	br()

	// ── หนี้สินและส่วนของผู้ถือหุ้น header ──
	mc("B", "H")
	set("B", lbl.LiabEqHdr, stSec)
	br()

	// ── หนี้สินหมุนเวียน ──
	secRow("B", "H", lbl.SecCL)
	dataRow("C", "R", lbl.BankOD, stLblI, bs.BankOD, bs.Prev.BankOD)
	dataRow("C", "J", lbl.Payable, stLblI, bs.Payable, bs.Prev.Payable)
	dataRow("C", "J", lbl.Dividend, stLblI, bs.Dividend, bs.Prev.Dividend)
	dataRow("C", "J", lbl.CurrentLTL, stLblI, bs.CurrentLTL, bs.Prev.CurrentLTL)
	dataRow("C", "J", lbl.CurrentLoan, stLblI, bs.CurrentLoan, bs.Prev.CurrentLoan)
	dataRow("C", "J", lbl.OtherCL, stLblI, bs.OtherCL, bs.Prev.OtherCL)
	totRow("D", "M", lbl.TotalCL, bs.TotalCL, bs.Prev.TotalCL)
	br()

	// ── หนี้สินไม่หมุนเวียน ──
	secRow("B", "H", lbl.SecNCL)
	dataRow("C", "N", lbl.DirLiab, stLblI, bs.DirLiab, bs.Prev.DirLiab)
	dataRow("C", "N", lbl.LTLiab, stLblI, bs.LTLiab, bs.Prev.LTLiab)
	dataRow("C", "J", lbl.Pension, stLblI, bs.Pension, bs.Prev.Pension)
	dataRow("C", "J", lbl.LTLoan, stLblI, bs.LTLoan, bs.Prev.LTLoan)
	dataRow("C", "J", lbl.OtherNCL, stLblI, bs.OtherNCL, bs.Prev.OtherNCL)
	totRow("D", "M", lbl.TotalNCL, bs.TotalNCL, bs.Prev.TotalNCL)
	br()
	totRow("D", "M", lbl.TotalLiab, bs.TotalLiab, bs.Prev.TotalLiab)
	br()

	// ── ส่วนของผู้ถือหุ้น ──
	mc("B", "N")
	set("B", lbl.SecEquity, stSec)
	br()

	mc("C", "I")
	set("C", lbl.SecCapital, stLblI)
	br()

	// หุ้นสามัญจำนวน G=qty M=หุ้น
	mc("C", "F")
	set("C", lbl.ShareQtyLbl, stLblI2)
	mc("G", "K")
	set("G", bsNum(bs.ThisYearQty), stNum)
	mc("M", "P")
	set("M", lbl.ShareQtyUnit, stLbl)
	br()

	// มูลค่าหุ้นละ H=val M=บาท T=ยอดทุน V=ยอดทุนปีก่อน
	mc("D", "G")
	set("D", lbl.ShareValLbl, stLblI2)
	mc("H", "K")
	set("H", bsNum(bs.ThisYearValue), stNum)
	mc("M", "P")
	set("M", lbl.ShareValUnit, stLbl)
	num("T", bs.Capital, stNum)
	mc("V", "W")
	num("V", bs.Prev.Capital, stNum)
	br()
	br()

	dataRow("C", "I", lbl.PaidCapital, stLblI, bs.Capital, bs.Prev.Capital)
	dataRow("C", "I", lbl.SharePremium, stLblI, bs.SharePremium, bs.Prev.SharePremium)
	dataRow("C", "I", lbl.RetainedEarn, stLblI, bs.RetainedEarn, bs.Prev.RetainedEarn)
	totRow("D", "M", lbl.TotalEquity, bs.TotalEquity, bs.Prev.TotalEquity)
	br()
	grandRow("B", "J", lbl.TotalLiabEq, bs.TotalLiabEquity, bs.Prev.TotalLiabEquity)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return fx.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// PDF export — A4, 2-column numbers layout
// ─────────────────────────────────────────────────────────────────
func exportBalanceSheetPDF(bs BalanceSheet, lbl bsLabels, savePath string) (string, error) {
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

	// หา bold font — fallback เป็น fontPath เดิมถ้าไม่พบ
	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

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
	sfBold := func() { pdf.SetFont("thai-bold", "", fs) }
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
		lw, _ := pdf.MeasureTextWidth(label)
		nl(fs + 3)
		hln(lm, lm+lw, y, 0.3)
		nl(2)
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
		// single underline เหนือตัวเลข — วาด colT และ colV ใน y เดียวกัน
		hln(colT-120, colT+3, y, 0.5)
		hln(colV-120, colV+3, y, 0.5)
		nl(1)
		sf()
		pdf.SetXY(lm, y)
		pdf.Cell(nil, label)
		putNums(cur, prev)
		nl(fs + 4)
		// double underline — วาดทั้ง 2 col ก่อน nl
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
	// ── Header ──
	// ComName — Bold
	pdf.SetFont("thai-bold", "", 13)
	w, _ := pdf.MeasureTextWidth(bs.ComName)
	pdf.SetXY((595-w)/2, y)
	pdf.Cell(nil, bs.ComName)
	nl(13 + 4)

	printCtr(lbl.Title, 11)
	printCtr(fmt.Sprintf(lbl.DateFmt, bs.PeriodStart, bs.PeriodEnd), 9)
	nl(6)

	// column header line
	sfBold()
	pdf.SetXY(lm, y)
	pdf.Cell(nil, lbl.AssetHdr)
	ch, _ := pdf.MeasureTextWidth(lbl.ColCur)
	ph, _ := pdf.MeasureTextWidth(lbl.ColPrev)
	pdf.SetXY(colT-ch/2, y)
	pdf.Cell(nil, lbl.ColCur)
	pdf.SetXY(colV-ph, y)
	pdf.Cell(nil, lbl.ColPrev)
	sf()
	nl(fs + 3)

	// ── สินทรัพย์หมุนเวียน ──
	printSec(lbl.SecCA)
	printRow(lbl.Cash, 10, bs.Cash, bs.Prev.Cash)
	printRow(lbl.ShortInvest, 10, bs.ShortInvest, bs.Prev.ShortInvest)
	printRow(lbl.Receivable, 10, bs.Receivable, bs.Prev.Receivable)
	printRow(lbl.LoanST, 10, bs.LoanST, bs.Prev.LoanST)
	printRow(lbl.Inventory, 10, bs.Inventory, bs.Prev.Inventory)
	printRow(lbl.OtherCA, 10, bs.OtherCA, bs.Prev.OtherCA)
	printTot(lbl.TotalCA, 5, bs.TotalCA, bs.Prev.TotalCA)
	nl(4)

	// ── สินทรัพย์ไม่หมุนเวียน ──
	printSec(lbl.SecNCA)
	printRow(lbl.DirLoan, 10, bs.DirLoan, bs.Prev.DirLoan)
	printRow(lbl.LongInvest, 10, bs.LongInvest, bs.Prev.LongInvest)
	printRow(lbl.PPE, 10, bs.PPE, bs.Prev.PPE)
	printRow(lbl.OtherNCA, 10, bs.OtherNCA, bs.Prev.OtherNCA)
	printTot(lbl.TotalNCA, 5, bs.TotalNCA, bs.Prev.TotalNCA)
	nl(4)
	printGrand(lbl.TotalAssets, bs.TotalAssets, bs.Prev.TotalAssets)

	// ── หนี้สินและส่วนของผู้ถือหุ้น (center, ไม่มี underline) ──
	sfBold()
	pdf.SetXY(lm, y)
	pdf.Cell(nil, lbl.LiabEqHdr)
	sf()
	nl(fs + 3)

	// ── หนี้สินหมุนเวียน ──
	printSec(lbl.SecCL)
	printRow(lbl.BankOD, 10, bs.BankOD, bs.Prev.BankOD)
	printRow(lbl.Payable, 10, bs.Payable, bs.Prev.Payable)
	printRow(lbl.Dividend, 10, bs.Dividend, bs.Prev.Dividend)
	printRow(lbl.CurrentLTL, 10, bs.CurrentLTL, bs.Prev.CurrentLTL)
	printRow(lbl.CurrentLoan, 10, bs.CurrentLoan, bs.Prev.CurrentLoan)
	printRow(lbl.OtherCL, 10, bs.OtherCL, bs.Prev.OtherCL)
	printTot(lbl.TotalCL, 5, bs.TotalCL, bs.Prev.TotalCL)
	nl(4)

	// ── หนี้สินไม่หมุนเวียน ──
	printSec(lbl.SecNCL)
	printRow(lbl.DirLiab, 10, bs.DirLiab, bs.Prev.DirLiab)
	printRow(lbl.LTLiab, 10, bs.LTLiab, bs.Prev.LTLiab)
	printRow(lbl.Pension, 10, bs.Pension, bs.Prev.Pension)
	printRow(lbl.LTLoan, 10, bs.LTLoan, bs.Prev.LTLoan)
	printRow(lbl.OtherNCL, 10, bs.OtherNCL, bs.Prev.OtherNCL)
	printTot(lbl.TotalNCL, 5, bs.TotalNCL, bs.Prev.TotalNCL)
	nl(2)
	printTot(lbl.TotalLiab, 5, bs.TotalLiab, bs.Prev.TotalLiab)
	nl(4)

	// ── ส่วนของผู้ถือหุ้น ──
	printSec(lbl.SecEquity)
	sf()
	pdf.SetXY(lm+10, y)
	pdf.Cell(nil, lbl.SecCapital)
	nl(fs + 2)

	// หุ้นสามัญจำนวน xxx หุ้น
	sf()
	pdf.SetXY(lm+20, y)
	pdf.Cell(nil, lbl.ShareQtyLbl)
	qs := bsNum(bs.ThisYearQty) + "  " + lbl.ShareQtyUnit
	qw, _ := pdf.MeasureTextWidth(qs)
	pdf.SetXY(colT-200-qw, y)
	pdf.Cell(nil, qs)
	nl(fs + 2)

	// มูลค่าหุ้นละ xxx บาท  [ยอดทุนปีนี้]  [ยอดทุนปีก่อน]
	sf()
	pdf.SetXY(lm+20, y)
	pdf.Cell(nil, lbl.ShareValLbl)
	vs := bsNum(bs.ThisYearValue) + "  " + lbl.ShareValUnit
	vw, _ := pdf.MeasureTextWidth(vs)
	pdf.SetXY(colT-200-vw, y)
	pdf.Cell(nil, vs)
	putNums(bs.Capital, bs.Prev.Capital)
	nl(fs + 4)

	printRow(lbl.PaidCapital, 10, bs.Capital, bs.Prev.Capital)
	printRow(lbl.SharePremium, 10, bs.SharePremium, bs.Prev.SharePremium)
	printRow(lbl.RetainedEarn, 10, bs.RetainedEarn, bs.Prev.RetainedEarn)
	printTot(lbl.TotalEquity, 5, bs.TotalEquity, bs.Prev.TotalEquity)
	nl(4)
	printGrand(lbl.TotalLiabEq, bs.TotalLiabEquity, bs.Prev.TotalLiabEquity)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// UI Dialog
// ─────────────────────────────────────────────────────────────────
func showBalanceSheetDialog(w fyne.Window, onGoSetup func()) {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOpts)
	if err != nil {
		dialog.ShowError(err, w)
		return
	}

	// ── guard: ต้องตั้ง Report Path ก่อน ──
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
	// แสดงเฉพาะ period 1 ถึง NowPeriod (ที่มีข้อมูลแล้ว)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	options := make([]string, showUpTo)
	for i, p := range periods[:showUpTo] {
		options[i] = fmt.Sprintf("Period %d (%s)", i+1, p.PEnd.Format("02/01/06"))
	}
	selPeriod := widget.NewSelect(options, nil)
	selPeriod.SetSelectedIndex(showUpTo - 1)

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
			widget.NewLabelWithStyle("งบแสดงฐานะการเงิน / Statement of Financial Position",
				fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			widget.NewLabel("เลือก Period:"),
			selPeriod,
			container.NewHBox(btnExcelTH, btnExcelEN, btnPDFTH, btnPDFEN, btnCancel),
		),
		w.Canvas(),
	)
	pop.Resize(fyne.NewSize(520, 180))

	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		if key.Name == fyne.KeyEscape {
			closePopup()
		}
	})
	btnCancel.OnTapped = closePopup

	run := func(isPDF bool, lbl bsLabels, lang string) {
		periodNo := selPeriod.SelectedIndex() + 1
		if periodNo < 1 {
			dialog.ShowInformation("แจ้งเตือน", "กรุณาเลือก Period", w)
			return
		}
		closePopup()

		bs, err := buildBalanceSheet(xlOpts, periodNo)
		if err != nil {
			dialog.ShowError(err, w)
			return
		}

		var savePath string
		var pathToOpen string
		var exportErr error
		if isPDF {
			savePath = filepath.Join(reportDir, fmt.Sprintf("BalanceSheet_P%02d_%s.pdf", periodNo, lang))
			pathToOpen, exportErr = exportBalanceSheetPDF(bs, lbl, savePath)
		} else {
			savePath = filepath.Join(reportDir, fmt.Sprintf("BalanceSheet_P%02d_%s.xlsx", periodNo, lang))
			pathToOpen, exportErr = exportBalanceSheetExcel(bs, lbl, savePath)
		}
		if exportErr != nil {
			dialog.ShowError(exportErr, w)
			return
		}
		showDoneDialog(w, pathToOpen)
	}

	btnExcelTH.OnTapped = func() { run(false, bsLabelsTH, "TH") }
	btnExcelEN.OnTapped = func() { run(false, bsLabelsEN, "EN") }
	btnPDFTH.OnTapped = func() { run(true, bsLabelsTH, "TH") }
	btnPDFEN.OnTapped = func() { run(true, bsLabelsEN, "EN") }
	pop.Show()
	w.Canvas().Focus(nil)
}
