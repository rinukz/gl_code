package main

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// ledgerNum — format ตัวเลขสำหรับ Ledger UI
//
//	บวก  →  1,234,567.89
//	ลบ   → -1,234,567.89   (ใช้เครื่องหมายลบ ไม่ใช้วงเล็บ)
//	ศูนย์→  0.00
//
// ใช้ math.Round เพื่อตัด floating-point error
// ─────────────────────────────────────────────────────────────────
func ledgerNum(v float64) string {
	// ป้องกัน -0.00
	v = math.Round(v*100) / 100
	if v == 0 {
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
		return "-" + out
	}
	return out
}

// ledgerParseNum — parse ตัวเลขที่ format ด้วย ledgerNum กลับเป็น float64
// รองรับทั้ง "-1,234.56" และ "(1,234.56)" (legacy) และ "1234.56"
func ledgerParseNum(s string) float64 {
	s = strings.TrimSpace(s)
	if s == "" || s == "-" {
		return 0
	}
	neg := false
	// รองรับ legacy format วงเล็บ
	if strings.HasPrefix(s, "(") && strings.HasSuffix(s, ")") {
		s = s[1 : len(s)-1]
		neg = true
	}
	s = strings.ReplaceAll(s, ",", "")
	v, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0
	}
	if neg {
		return -v
	}
	return v
}

// ─────────────────────────────────────────────────────────────────
// LedgerRecord เก็บข้อมูล 1 record ใน Ledger_Master
// ─────────────────────────────────────────────────────────────────
type LedgerRecord struct {
	Comcode   string
	AcCode    string
	AcName    string
	Gcode     string
	Gname     string
	BBal      string
	CBal      string
	Debit     string
	Credit    string
	Bthisyear string
	Thisper   [12]string // [0]=per01 ... [11]=per12
	Blastyear string
	Lastper   [12]string
}

func emptyLedger() LedgerRecord {
	r := LedgerRecord{}
	r.BBal = "0.00"
	r.CBal = "0.00"
	r.Debit = "0.00"
	r.Credit = "0.00"
	r.Bthisyear = "0.00"
	r.Blastyear = "0.00"
	for i := 0; i < 12; i++ {
		r.Thisper[i] = "0.00"
		r.Lastper[i] = "0.00"
	}
	return r
}

func loadLedgerRecord(xlOptions excelize.Options, acCode string) (LedgerRecord, bool) {
	comCode := getComCodeFromExcel(xlOptions)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return emptyLedger(), false
	}
	defer f.Close()

	rows, _ := f.GetRows("Ledger_Master")
	for _, row := range rows {
		if len(row) >= 3 && row[0] == comCode && row[1] == acCode {
			r := emptyLedger()
			r.Comcode = row[0]
			r.AcCode = row[1]
			r.AcName = row[2]
			if len(row) > 3 {
				r.Gcode = row[3]
			}
			if len(row) > 4 {
				r.Gname = row[4]
			}
			if len(row) > 5 {
				r.BBal = row[5]
			}
			if len(row) > 6 {
				r.CBal = row[6]
			}
			if len(row) > 7 {
				r.Debit = row[7]
			}
			if len(row) > 8 {
				r.Credit = row[8]
			}
			if len(row) > 9 {
				r.Bthisyear = row[9]
			}
			for i := 0; i < 12; i++ {
				if len(row) > 10+i {
					r.Thisper[i] = row[10+i]
				}
			}
			if len(row) > 22 {
				r.Blastyear = row[22]
			}
			for i := 0; i < 12; i++ {
				if len(row) > 23+i {
					r.Lastper[i] = row[23+i]
				}
			}
			return r, true
		}
	}
	return emptyLedger(), false
}

func loadAllAcCodes(xlOptions excelize.Options) []string {
	comCode := getComCodeFromExcel(xlOptions)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	var codes []string
	rows, _ := f.GetRows("Ledger_Master")
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 2 && row[0] == comCode {
			codes = append(codes, row[1])
		}
	}
	sort.Strings(codes)
	return codes
}

func lookupGname(xlOptions excelize.Options, gcode string) string {
	comCode := getComCodeFromExcel(xlOptions)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return ""
	}
	defer f.Close()

	rows, _ := f.GetRows("Acct_Group")
	for _, row := range rows {
		if len(row) >= 3 && row[0] == comCode && row[1] == gcode {
			return row[2]
		}
	}
	return ""
}

func isSpecialCode(acCode string) bool {
	if len(acCode) < 3 {
		return false
	}
	prefix := acCode[:3]
	return acCode == "120VAT" || acCode == "235VAT" || acCode == "235TVAT" || acCode == "120WHT" ||
		acCode == "235WHT" || acCode == "524RND" || acCode == "450RND" ||
		prefix == "350" || prefix == "360"
}

// ─────────────────────────────────────────────────────────────────
// saveLedgerRecord — บันทึก LedgerRecord ลง Excel
//
// BUG FIX: BBal/CBal/Debit/Credit เป็น computed fields ไม่ควรบันทึกจาก form
// ดังนั้นจะ strip comma ออกก่อนบันทึกเสมอ
// ─────────────────────────────────────────────────────────────────
func saveLedgerRecord(xlOptions excelize.Options, r LedgerRecord, isNew bool) error {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return err
	}
	defer f.Close()

	sheet := "Ledger_Master"
	rows, _ := f.GetRows(sheet)

	targetRow := -1
	if !isNew {
		for i, row := range rows {
			if len(row) >= 2 && row[0] == r.Comcode && row[1] == r.AcCode {
				targetRow = i + 1
				break
			}
		}
	}
	if targetRow == -1 {
		if len(rows) == 0 {
			headers := []string{
				"Comcode", "Ac_code", "Ac_name", "Gcode", "Gname",
				"BBAL", "CBAL", "Debit", "Credit", "Bthisyear",
				"Thisper01", "Thisper02", "Thisper03", "Thisper04",
				"Thisper05", "Thisper06", "Thisper07", "Thisper08",
				"Thisper09", "Thisper10", "Thisper11", "Thisper12",
				"Blastyear",
				"Lastper01", "Lastper02", "Lastper03", "Lastper04",
				"Lastper05", "Lastper06", "Lastper07", "Lastper08",
				"Lastper09", "Lastper10", "Lastper11", "Lastper12",
			}
			for i, h := range headers {
				col, _ := excelize.ColumnNumberToName(i + 1)
				f.SetCellValue(sheet, fmt.Sprintf("%s1", col), h)
			}
			rows = append(rows, headers)
		}
		targetRow = len(rows) + 1
	}

	// helper: แปลง string → float64 (รองรับทั้ง comma-format และ raw)
	toFloat := func(s string) float64 {
		return ledgerParseNum(s)
	}

	// ค่าที่บันทึก: col 0-4 เป็น string, col 5+ เป็น float
	stringVals := []string{r.Comcode, r.AcCode, r.AcName, r.Gcode, r.Gname}
	for i, v := range stringVals {
		col, _ := excelize.ColumnNumberToName(i + 1)
		f.SetCellValue(sheet, fmt.Sprintf("%s%d", col, targetRow), v)
	}

	// numeric fields: BBAL(6) CBal(7) Debit(8) Credit(9) Bthisyear(10)
	// Thisper01-12(11-22) Blastyear(23) Lastper01-12(24-35)
	numVals := []string{
		r.BBal, r.CBal, r.Debit, r.Credit, r.Bthisyear,
		r.Thisper[0], r.Thisper[1], r.Thisper[2], r.Thisper[3],
		r.Thisper[4], r.Thisper[5], r.Thisper[6], r.Thisper[7],
		r.Thisper[8], r.Thisper[9], r.Thisper[10], r.Thisper[11],
		r.Blastyear,
		r.Lastper[0], r.Lastper[1], r.Lastper[2], r.Lastper[3],
		r.Lastper[4], r.Lastper[5], r.Lastper[6], r.Lastper[7],
		r.Lastper[8], r.Lastper[9], r.Lastper[10], r.Lastper[11],
	}
	for i, v := range numVals {
		col, _ := excelize.ColumnNumberToName(i + 6) // เริ่มที่ col 6
		f.SetCellValue(sheet, fmt.Sprintf("%s%d", col, targetRow), toFloat(v))
	}

	return f.Save()
}

// isAcCodeUsedInBook — ตรวจว่า acCode ถูกใช้ใน Book_items ไหม
// คืนค่า count (จำนวน voucher ที่ใช้) และ examples (ตัวอย่าง Bitem สูงสุด 3 รายการ)
func isAcCodeUsedInBook(xlOptions excelize.Options, acCode string) (count int, examples []string) {
	comCode := getComCodeFromExcel(xlOptions)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return 0, nil
	}
	defer f.Close()

	rows, _ := f.GetRows("Book_items")
	seen := map[string]bool{} // นับแต่ละ Bitem ครั้งเดียว
	for i, row := range rows {
		if i == 0 || len(row) < 6 || safeGet(row, 0) != comCode {
			continue
		}
		if safeGet(row, 5) != acCode {
			continue
		}
		bitem := safeGet(row, 3)
		if seen[bitem] {
			continue
		}
		seen[bitem] = true
		count++
		if len(examples) < 3 {
			examples = append(examples, bitem)
		}
	}
	return count, examples
}

func deleteLedgerRecord(xlOptions excelize.Options, acCode string) error {
	comCode := getComCodeFromExcel(xlOptions)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return err
	}
	defer f.Close()

	sheet := "Ledger_Master"
	rows, _ := f.GetRows(sheet)
	for i, row := range rows {
		if len(row) >= 2 && row[0] == comCode && row[1] == acCode {
			f.RemoveRow(sheet, i+1)
			break
		}
	}
	return f.Save()
}

// ═══════════════════════════════════════════════════════════
// RecalculateLedgerMaster — คำนวณยอด Thisper01-12, Debit, Credit
// จาก Book_items แล้วอัปเดตลง Ledger_Master
// ═══════════════════════════════════════════════════════════
func RecalculateLedgerMaster(xlOptions excelize.Options) error {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return err
	}

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)

	// 1. โหลดและรวมยอดจาก Book_items
	type acctSum struct {
		periods [12]float64
		dr      float64
		cr      float64
	}
	sumMap := make(map[string]*acctSum)

	bRows, _ := f.GetRows("Book_items")
	for i, r := range bRows {
		if i == 0 || len(r) < 11 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := safeGet(r, 5)
		if acCode == "" {
			continue
		}
		dtStr := safeGet(r, 1)
		dt, err := parseSubbookDate(dtStr)
		if err != nil {
			continue
		}
		dr := parseFloat(safeGet(r, 9))
		cr := parseFloat(safeGet(r, 10))

		if sumMap[acCode] == nil {
			sumMap[acCode] = &acctSum{}
		}
		sumMap[acCode].dr += dr
		sumMap[acCode].cr += cr

		for pIdx, p := range periods {
			if !dt.Before(p.PStart) && !dt.After(p.PEnd) {
				sumMap[acCode].periods[pIdx] += (dr - cr)
				break
			}
		}
	}

	// คำนวณ 360PLA (Profit/Loss Account) โดยรวมจากหมวด 4 และ 5
	plaSum := &acctSum{}
	hasPLA := false
	for code, data := range sumMap {
		if len(code) > 0 && (code[0] == '4' || code[0] == '5') {
			hasPLA = true
			plaSum.dr += data.dr
			plaSum.cr += data.cr
			for pIdx := 0; pIdx < 12; pIdx++ {
				plaSum.periods[pIdx] += data.periods[pIdx]
			}
		}
	}
	if hasPLA {
		sumMap["360PLA"] = plaSum
	}

	// 2. อัปเดตลง Ledger_Master
	lRows, _ := f.GetRows("Ledger_Master")
	for i, r := range lRows {
		if i == 0 || len(r) < 4 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := safeGet(r, 1)
		rowNum := strconv.Itoa(i + 1)

		// 2.1 Backfill Bthisyear (col J=10) จาก Lastper12 (col AI=35) ถ้า Bthisyear ยังเป็น 0
		// เฉพาะบัญชีงบดุล (code < "4") — P&L ไม่มียอดยกมา
		// ยกเว้น 360PLA (กำไรสุทธิประจำปี) ซึ่งจะถูกปิดเข้ากำไรสะสม (350REA) ตอนสิ้นปี จึงไม่มียอดยกมา
		bthis := parseFloat(safeGet(r, 9)) // Bthisyear col J
		if acCode < "4" && acCode != "360PLA" {
			lastper12 := parseFloat(safeGet(r, 34)) // Lastper12 col AI
			if bthis == 0 && lastper12 != 0 {
				bthis = lastper12
				f.SetCellValue("Ledger_Master", "J"+rowNum, bthis)
			}
		} else {
			bthis = 0 // P&L ไม่มี Bthisyear มาทบ
		}

		// 2.2 อัปเดต Debit, Credit และยอดสะสมรายเดือน (Thisper01-12)
		sumData, exists := sumMap[acCode]
		if !exists {
			sumData = &acctSum{}
		}

		// อัปเดต Debit (col H=8), Credit (col I=9)
		f.SetCellValue("Ledger_Master", "H"+rowNum, sumData.dr)
		f.SetCellValue("Ledger_Master", "I"+rowNum, sumData.cr)

		// อัปเดต Thisper01-12 (col K=11 ถึง V=22) ให้เป็นยอดสะสม (Cumulative Balance)
		cumulative := bthis
		for pIdx := 0; pIdx < 12; pIdx++ {
			cumulative += sumData.periods[pIdx]
			colName, _ := excelize.ColumnNumberToName(11 + pIdx)
			f.SetCellValue("Ledger_Master", colName+rowNum, cumulative)
		}
	}

	return f.Save()
}

// ─────────────────────────────────────────────────────────────────
// computeRealTimeLedger — คำนวณ BBal/CBal/Debit/Credit แบบ real-time
//
// เลียนแบบ VB.NET FrmLedger.calculatePeriod() chain:
//
//	BBal[1]   = Bthisyear
//	BBal[N]   = BBal[N-1] + SumDR[N-1] - SumCR[N-1]
//	Debit/CR  = SUM(Book_items) WHERE period = NowPeriod
//	CBal      = BBal[Now] + Debit - Credit
//
// BUG FIX: เดิมเปิด Excel ซ้ำทุกงวด → ช้ามาก
// แก้ไข: โหลด Book_items ครั้งเดียว แล้ว filter ใน memory
// ─────────────────────────────────────────────────────────────────
func computeRealTimeLedger(xlOptions excelize.Options, acCode string) (bbal, cbal, dr, cr float64) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return
	}
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if len(periods) == 0 || cfg.NowPeriod < 1 || cfg.NowPeriod > len(periods) {
		return
	}

	// อ่าน Bthisyear จาก Ledger_Master
	rec, found := loadLedgerRecord(xlOptions, acCode)
	if !found {
		return
	}
	bthisyear := parseFloat(rec.Bthisyear)

	// โหลด Book_items ครั้งเดียว
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	bookRows, _ := f.GetRows("Book_items")

	// pre-parse transactions ของ acCode นี้ทั้งหมด (หรือ 4xx, 5xx ถ้าเป็น 360PLA)
	type txn struct {
		date time.Time
		dr   float64
		cr   float64
	}
	var txns []txn
	for i, row := range bookRows {
		if i == 0 || len(row) < 11 || safeGet(row, 0) != comCode {
			continue
		}
		rowAcCode := safeGet(row, 5)

		// ถ้าเป็น 360PLA ให้รวมยอดจากหมวด 4 และ 5 ทั้งหมด
		if acCode == "360PLA" {
			if len(rowAcCode) < 1 || (rowAcCode[0] != '4' && rowAcCode[0] != '5') {
				continue
			}
		} else {
			if rowAcCode != acCode {
				continue
			}
		}

		t, err := time.Parse("02/01/06", strings.TrimSpace(safeGet(row, 1)))
		if err != nil {
			continue
		}
		txns = append(txns, txn{
			date: t,
			dr:   parseFloat(safeGet(row, 9)),
			cr:   parseFloat(safeGet(row, 10)),
		})
	}

	// sumForPeriod — รวมยอดจาก txns ที่อยู่ใน period
	sumForPeriod := func(p PeriodInfo) (pDR, pCR float64) {
		for _, tx := range txns {
			if !tx.date.Before(p.PStart) && !tx.date.After(p.PEnd) {
				pDR += tx.dr
				pCR += tx.cr
			}
		}
		return
	}

	// Chain: Bthisyear → BBal ของงวดปัจจุบัน
	prev := bthisyear
	for p := 1; p < cfg.NowPeriod; p++ {
		pDR, pCR := sumForPeriod(periods[p-1])
		prev = prev + pDR - pCR
	}
	bbal = prev

	// Debit/Credit ของงวดปัจจุบัน
	dr, cr = sumForPeriod(periods[cfg.NowPeriod-1])
	cbal = bbal + dr - cr
	return
}

// sumBookForPeriod — helper สำหรับ RecalculateLedgerMaster
// (ยังคงไว้ เพราะอาจถูกเรียกจากไฟล์อื่น)
func sumBookForPeriod(xlOptions excelize.Options, acCode string, period PeriodInfo) (dr, cr float64) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()
	comCode := getComCodeFromExcel(xlOptions)
	rows, _ := f.GetRows("Book_items")
	for i, row := range rows {
		if i == 0 || len(row) < 11 || row[0] != comCode || safeGet(row, 5) != acCode {
			continue
		}
		t, err := time.Parse("02/01/06", safeGet(row, 1))
		if err != nil || t.Before(period.PStart) || t.After(period.PEnd) {
			continue
		}
		dr += parseFloat(safeGet(row, 9))
		cr += parseFloat(safeGet(row, 10))
	}
	return
}

// ═══════════════════════════════════════════════════════════
// getLedgerGUI — หน้าจอ Ledger Master
// ═══════════════════════════════════════════════════════════
func getLedgerGUI(w fyne.Window) (fyne.CanvasObject, func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	comCode := getComCodeFromExcel(xlOptions)

	allCodes := loadAllAcCodes(xlOptions)
	currentIdx := 0

	actionFlag := "VIEW"
	var saveAction func()

	// ── Fields ──
	enAcCode := newSmartEntry(func() { saveAction() })
	lblGname := widget.NewLabel("")
	enAcName := newSmartEntry(func() { saveAction() })

	// BBal/CBal/Debit/Credit เป็น read-only computed fields
	enBBal := newSmartEntry(nil)
	enBBal.Disable()
	enCBal := newSmartEntry(nil)
	enCBal.Disable()
	enDebit := newSmartEntry(nil)
	enDebit.Disable()
	enCredit := newSmartEntry(nil)
	enCredit.Disable()

	enBthisyear := newSmartEntry(func() { saveAction() })
	var enThisper [12]*smartEntry
	var enLastper [12]*smartEntry
	enBlastyear := newSmartEntry(func() { saveAction() })

	for i := 0; i < 12; i++ {
		enThisper[i] = newSmartEntry(nil)
		enThisper[i].Disable() // This Year readonly — คำนวณจาก Book_items
		enLastper[i] = newSmartEntry(func() { saveAction() })
	}

	var btnSave, btnAdd, btnEdit, btnDel *widget.Button
	var btnPrev, btnNext *widget.Button

	// ── loadForm — โหลดข้อมูลเข้า form ──────────────────────────
	// ทุก numeric field แสดงผลด้วย ledgerNum() เหมือนกัน
	loadForm := func(r LedgerRecord) {
		enAcCode.SetText(r.AcCode)
		enAcName.SetText(r.AcName)
		lblGname.SetText(r.Gname)

		// BBal/CBal/Debit/Credit คำนวณ real-time จาก Book_items
		bbal, cbal, dr, cr := computeRealTimeLedger(xlOptions, r.AcCode)
		enBBal.SetText(ledgerNum(bbal))
		enCBal.SetText(ledgerNum(cbal))
		enDebit.SetText(ledgerNum(dr))
		enCredit.SetText(ledgerNum(cr))

		// Bthisyear (ยอดยกมาต้นปี — user กรอก)
		enBthisyear.SetText(ledgerNum(parseFloat(r.Bthisyear)))

		// Blastyear
		enBlastyear.SetText(ledgerNum(parseFloat(r.Blastyear)))

		// โหลดค่า Current Period จากไฟล์ helpers.go
		currentPeriod := LoadCurrentPeriod(xlOptions)

		// นำมาเช็คเงื่อนไขตอน Loop Thisper01-12
		for i := 0; i < 12; i++ {
			if i < currentPeriod {
				// งวดที่ถึงแล้ว (i เริ่ม 0, ดังนั้น i < 1 คือ งวด 1) แสดงผลปกติ
				enThisper[i].SetText(ledgerNum(parseFloat(r.Thisper[i])))
			} else {
				// งวดที่ยังไม่ถึง ให้แสดง "-"
				enThisper[i].SetText("0.00")
			}
		}

		// Lastper01-12 (user กรอก)
		// หมายเหตุ: อันนี้เป็นข้อมูลปีที่แล้ว ปกติควรจะโชว์ครบ 12 งวด เลยไม่ต้องใส่เงื่อนไขดัก
		for i := 0; i < 12; i++ {
			enLastper[i].SetText(ledgerNum(parseFloat(r.Lastper[i])))
		}

	}

	modeLabel := widget.NewLabelWithStyle("● VIEW MODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})

	setViewMode := func() {
		actionFlag = "VIEW"
		modeLabel.SetText("● VIEW MODE")
		enAcCode.Disable()
		enAcName.Disable()
		enBthisyear.Disable()
		enBlastyear.Disable()
		for i := 0; i < 12; i++ {
			enLastper[i].Disable()
		}
		btnSave.Disable()
		btnAdd.Enable()
		btnEdit.Enable()
		btnDel.Enable()
		btnPrev.Enable()
		btnNext.Enable()
	}

	setEditMode := func(isAdd bool) {
		if isAdd {
			actionFlag = "ADD"
			modeLabel.SetText("● ADD MODE")
		} else {
			actionFlag = "EDIT"
			modeLabel.SetText("● EDIT MODE")
			enAcCode.Disable()
		}
		enAcName.Enable()
		enBthisyear.Enable()
		enBlastyear.Enable()
		for i := 0; i < 12; i++ {
			enLastper[i].Enable()
		}
		btnSave.Enable()
		btnAdd.Disable()
		btnEdit.Disable()
		btnDel.Disable()
		btnPrev.Disable()
		btnNext.Disable()
	}

	nextAction := func() {
		if actionFlag == "VIEW" && currentIdx < len(allCodes)-1 {
			currentIdx++
			r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
			loadForm(r)
		}
	}
	prevAction := func() {
		if actionFlag == "VIEW" && currentIdx > 0 {
			currentIdx--
			r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
			loadForm(r)
		}
	}

	ledgerNextFunc = nextAction
	ledgerPrevFunc = prevAction

	cancelAction := func() {
		if actionFlag == "ADD" || actionFlag == "EDIT" {
			if len(allCodes) > 0 {
				r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
				loadForm(r)
			} else {
				loadForm(emptyLedger())
			}
			setViewMode()
		}
	}

	openSearch := func() {
		if actionFlag != "VIEW" {
			return
		}
		showLedgerSearch(w, xlOptions, allCodes, func(r LedgerRecord, idx int) {
			if idx >= 0 {
				currentIdx = idx
			}
			full, found := loadLedgerRecord(xlOptions, r.AcCode)
			if found {
				loadForm(full)
			}
			setViewMode()
		})
	}

	ledgerSearchFunc = openSearch

	// ── ผูก keyboard shortcuts ทุก field ──
	allSmartFields := []*smartEntry{
		enAcCode, enAcName, enBBal, enCBal, enDebit, enCredit,
		enBthisyear, enBlastyear,
	}
	for _, e := range allSmartFields {
		e.onEsc = cancelAction
		e.onF3 = openSearch
		e.onPageDown = nextAction
		e.onPageUp = prevAction
	}
	for i := 0; i < 12; i++ {
		enThisper[i].onEsc = cancelAction
		enThisper[i].onF3 = openSearch
		enThisper[i].onPageDown = nextAction
		enThisper[i].onPageUp = prevAction

		enLastper[i].onEsc = cancelAction
		enLastper[i].onF3 = openSearch
		enLastper[i].onPageDown = nextAction
		enLastper[i].onPageUp = prevAction
	}

	// ── showErr ──
	showErr := func(msg string, refocus fyne.Focusable) {
		var d dialog.Dialog
		okBtn := newEnterButton("OK", func() {
			d.Hide()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() {
					if refocus != nil {
						w.Canvas().Focus(refocus)
					} else {
						w.Canvas().Focus(enAcCode)
					}
				})
			}()
		})
		d = dialog.NewCustomWithoutButtons("Error",
			container.NewVBox(
				widget.NewLabel(msg),
				container.NewCenter(okBtn),
			), w)
		d.Show()
		go func() {
			time.Sleep(50 * time.Millisecond)
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
		}()
	}

	// ── saveAction ──────────────────────────────────────────────
	// BUG FIX: ไม่บันทึก BBal/CBal/Debit/Credit (computed) ลง Excel
	// แต่บันทึกเฉพาะ Bthisyear และ Lastper ที่ user กรอก
	saveAction = func() {
		if actionFlag == "VIEW" {
			return
		}
		acCode := strings.TrimSpace(enAcCode.Text)
		acName := strings.TrimSpace(enAcName.Text)
		if acCode == "" || acName == "" {
			showErr("กรุณาระบุ Account Code และ Account Name", enAcCode)
			return
		}

		if actionFlag == "ADD" && isSpecialCode(acCode) {
			enAcCode.SetText("")
			showErr("Special Account Code ไม่อนุญาตให้เพิ่มที่นี่", enAcCode)
			return
		}

		if actionFlag == "EDIT" && isSpecialCode(acCode) {
			// ดึง AcName เดิมกลับมา (ไม่ให้ user เปลี่ยน)
			if orig, found := loadLedgerRecord(xlOptions, acCode); found {
				acName = orig.AcName
			}
		}

		if actionFlag == "ADD" {
			for _, code := range allCodes {
				if code == acCode {
					showErr(fmt.Sprintf("รหัส %s มีอยู่แล้วในระบบ", acCode), enAcCode)
					return
				}
			}
		}

		// อ่านค่า Lastper01-11 จาก form (user กรอก)
		// Lastper12 → sync เป็น Bthisyear อัตโนมัติ
		var lastperVals [12]string
		for i := 0; i < 12; i++ {
			// strip comma เพื่อให้ saveLedgerRecord แปลงเป็น float ได้
			lastperVals[i] = strings.ReplaceAll(enLastper[i].Text, ",", "")
		}

		// Bthisyear = Lastper12 (ตาม VB.NET logic)
		bthisyearVal := strings.ReplaceAll(enLastper[11].Text, ",", "")

		r := LedgerRecord{
			Comcode: comCode,
			AcCode:  acCode,
			AcName:  acName,
			Gcode:   acCode[:min3(len(acCode))],
			Gname:   lblGname.Text,
			// ดึงค่าเดิมกลับไปบันทึก เพื่อไม่ให้ค่าหาย
			BBal:      strings.ReplaceAll(enBBal.Text, ",", ""),
			CBal:      strings.ReplaceAll(enCBal.Text, ",", ""),
			Debit:     strings.ReplaceAll(enDebit.Text, ",", ""),
			Credit:    strings.ReplaceAll(enCredit.Text, ",", ""),
			Bthisyear: bthisyearVal,
			Blastyear: strings.ReplaceAll(enBlastyear.Text, ",", ""),
		}
		// Thisper01-12 ไม่บันทึกจาก form ตรงๆ แต่จะดึงค่าเดิมที่แสดงอยู่บนหน้าจอกลับไปบันทึก
		// เพื่อไม่ให้ค่าหายตอนกด Save (ค่าจะถูกคำนวณใหม่ตอนกด Recalculate)
		for i := 0; i < 12; i++ {
			r.Thisper[i] = strings.ReplaceAll(enThisper[i].Text, ",", "")
			r.Lastper[i] = lastperVals[i]
		}

		isNew := actionFlag == "ADD"
		if err := saveLedgerRecord(xlOptions, r, isNew); err != nil {
			showErr(fmt.Sprintf("บันทึกไม่สำเร็จ: %v", err), enAcCode)
			return
		}

		// sync Special Code name
		if isSpecialCode(acCode) {
			ff, err := excelize.OpenFile(currentDBPath, xlOptions)
			if err == nil {
				syncSpecialCodeName(ff, comCode, acCode, acName, "ledger→special")
				ff.Save()
				ff.Close()
			}
		}

		// รีโหลด allCodes
		allCodes = loadAllAcCodes(xlOptions)
		found := false
		for i, c := range allCodes {
			if c == acCode {
				currentIdx = i
				found = true
				break
			}
		}
		if !found || currentIdx < 0 || currentIdx >= len(allCodes) {
			currentIdx = 0
		}

		var d dialog.Dialog
		okBtn := newEnterButton("OK", func() {
			d.Hide()
			w.Canvas().Focus(nil)
			if len(allCodes) > 0 {
				rec, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
				loadForm(rec)
			}
			setViewMode()
		})
		content := container.NewVBox(
			widget.NewLabel("บันทึกเรียบร้อยแล้ว"),
			container.NewCenter(okBtn),
		)
		d = dialog.NewCustomWithoutButtons("สำเร็จ", content, w)
		d.Show()
		go func() {
			time.Sleep(50 * time.Millisecond)
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
		}()
	}

	// ── Buttons ──
	btnSave = widget.NewButtonWithIcon("", theme.DocumentSaveIcon(), saveAction)

	btnAdd = widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() {
		loadForm(emptyLedger())
		enAcCode.SetText("")
		lblGname.SetText("")
		setEditMode(true)
		go func() {
			fyne.Do(func() {
				enAcCode.Enable()
				w.Canvas().Focus(enAcCode)
			})
		}()
	})

	btnEdit = widget.NewButtonWithIcon("", theme.DocumentCreateIcon(), func() {
		if enAcCode.Text == "" {
			return
		}
		if isSpecialCode(enAcCode.Text) {
			// Special code: อนุญาตให้แก้ Last Year เท่านั้น
			// AcCode และ AcName ห้ามแก้ไข
			setEditMode(false)
			enAcName.Disable() // lock AcName ด้วย (special code ห้ามเปลี่ยนชื่อ)
			w.Canvas().Focus(enBlastyear)
			return
		}

		// ── Guard: ตรวจว่า AcCode ถูกใช้ใน Book_items ไหม ──
		// ถ้ามีการใช้งานอยู่ → แจ้งเตือน แต่ยังให้แก้ได้เฉพาะ AcName (ไม่ใช่ AcCode)
		// เหตุผล: AcName เปลี่ยนได้ แต่ AcCode เปลี่ยนไม่ได้อยู่แล้ว (Disable ใน EDIT mode)
		// ดังนั้น guard นี้ทำหน้าที่แจ้งให้ผู้ใช้รู้ว่ากำลังแก้ account ที่มีประวัติข้อมูล
		if cnt, ex := isAcCodeUsedInBook(xlOptions, enAcCode.Text); cnt > 0 {
			exStr := strings.Join(ex, ", ")
			if cnt > 3 {
				exStr += fmt.Sprintf(" ... และอีก %d voucher", cnt-3)
			}
			var warnD dialog.Dialog
			proceedBtn := newEnterEscButton(
				"แก้ไข AcName ต่อ (Enter)",
				func() { // onEnter
					warnD.Hide()
					setEditMode(false)
					w.Canvas().Focus(enAcName)
				},
				func() { // onEsc
					warnD.Hide()
					go func() {
						time.Sleep(50 * time.Millisecond)
						fyne.Do(func() { w.Canvas().Focus(enAcCode) })
					}()
				},
			)
			proceedBtn.Importance = widget.WarningImportance
			cancelBtn := newEscButton("ยกเลิก (Esc)", func() {
				warnD.Hide()
				go func() {
					time.Sleep(50 * time.Millisecond)
					fyne.Do(func() { w.Canvas().Focus(enAcCode) })
				}()
			})
			msg := fmt.Sprintf(
				"⚠️  Account Code \"%s\" มีการใช้งานอยู่ใน %d voucher\n(เช่น Item: %s)\n\n"+
					"สามารถแก้ไข Account Name ได้\n"+
					"แต่ Account Code จะไม่เปลี่ยน (ข้อมูลใน Book ยังคงสมบูรณ์)",
				enAcCode.Text, cnt, exStr,
			)
			warnD = dialog.NewCustomWithoutButtons("แจ้งเตือน — Account มีข้อมูลใน Book",
				container.NewVBox(
					widget.NewLabel(msg),
					widget.NewSeparator(),
					container.NewCenter(container.NewHBox(proceedBtn, cancelBtn)),
				), w)
			warnD.Show()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() { w.Canvas().Focus(proceedBtn) })
			}()
			return
		}
		setEditMode(false)
		w.Canvas().Focus(enAcName)
	})

	btnDel = widget.NewButtonWithIcon("", theme.DeleteIcon(), func() {
		acCode := enAcCode.Text
		if acCode == "" {
			return
		}
		if isSpecialCode(acCode) {
			var d dialog.Dialog
			okBtn := newEnterButton("OK", func() {
				d.Hide()
				go func() {
					time.Sleep(50 * time.Millisecond)
					fyne.Do(func() { w.Canvas().Focus(enAcCode) })
				}()
			})
			d = dialog.NewCustomWithoutButtons("Error",
				container.NewVBox(
					widget.NewLabel("ไม่อนุญาตให้ลบ Special Account Code"),
					container.NewCenter(okBtn),
				), w)
			d.Show()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() { w.Canvas().Focus(okBtn) })
			}()
			return
		}
		// ── Guard: ตรวจว่า AcCode ถูกใช้ใน Book_items ไหม ──
		// ถ้ามีการใช้งานอยู่ → ห้ามลบเด็ดขาด เพราะ Book_items จะกลายเป็น orphan
		if cnt, ex := isAcCodeUsedInBook(xlOptions, acCode); cnt > 0 {
			exStr := strings.Join(ex, ", ")
			if cnt > 3 {
				exStr += fmt.Sprintf(" ... และอีก %d voucher", cnt-3)
			}
			var blockD dialog.Dialog
			okBtn := newEnterButton("รับทราบ (Enter)", func() {
				blockD.Hide()
				go func() {
					time.Sleep(50 * time.Millisecond)
					fyne.Do(func() { w.Canvas().Focus(enAcCode) })
				}()
			})
			msg := fmt.Sprintf(
				"🚫  ไม่สามารถลบ Account Code \"%s\" ได้\n\n"+
					"เนื่องจากมีการใช้งานอยู่ใน %d voucher\n(เช่น Item: %s)\n\n"+
					"หากต้องการลบ ให้แก้ไขหรือลบ voucher ที่ใช้ Account นี้ก่อน\nจากนั้นค่อยลบ Account Code ออก",
				acCode, cnt, exStr,
			)
			blockD = dialog.NewCustomWithoutButtons("ไม่สามารถลบได้",
				container.NewVBox(
					widget.NewLabel(msg),
					widget.NewSeparator(),
					container.NewCenter(okBtn),
				), w)
			blockD.Show()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() { w.Canvas().Focus(okBtn) })
			}()
			return
		}
		var confirmD dialog.Dialog
		doCancel := func() {
			confirmD.Hide()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() { w.Canvas().Focus(enAcCode) })
			}()
		}
		yesBtn := newEnterEscButton("ลบ (Enter)", func() {
			confirmD.Hide()
			if err := deleteLedgerRecord(xlOptions, acCode); err != nil {
				var errD dialog.Dialog
				okBtn := newEnterButton("OK", func() {
					errD.Hide()
					go func() {
						time.Sleep(50 * time.Millisecond)
						fyne.Do(func() { w.Canvas().Focus(enAcCode) })
					}()
				})
				errD = dialog.NewCustomWithoutButtons("Error",
					container.NewVBox(
						widget.NewLabel(fmt.Sprintf("ลบไม่สำเร็จ: %v", err)),
						container.NewCenter(okBtn),
					), w)
				errD.Show()
				go func() {
					time.Sleep(50 * time.Millisecond)
					fyne.Do(func() { w.Canvas().Focus(okBtn) })
				}()
				return
			}
			allCodes = loadAllAcCodes(xlOptions)
			if currentIdx < 0 || currentIdx >= len(allCodes) {
				currentIdx = len(allCodes) - 1
			}
			if currentIdx < 0 {
				currentIdx = 0
			}
			if len(allCodes) > 0 {
				r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
				loadForm(r)
			} else {
				loadForm(emptyLedger())
			}
			setViewMode()
			go func() {
				time.Sleep(50 * time.Millisecond)
				fyne.Do(func() { w.Canvas().Focus(enAcCode) })
			}()
		}, doCancel)
		noBtn := newEscButton("ยกเลิก (Esc)", doCancel)
		yesBtn.Importance = widget.DangerImportance
		content := container.NewVBox(
			widget.NewLabel("ยืนยันการลบ: "+acCode+"?"),
			container.NewCenter(container.NewHBox(yesBtn, noBtn)),
		)
		confirmD = dialog.NewCustomWithoutButtons("ยืนยัน", content, w)
		confirmD.Show()
		go func() {
			time.Sleep(50 * time.Millisecond)
			fyne.Do(func() { w.Canvas().Focus(yesBtn) })
		}()
	})

	btnPrev = widget.NewButtonWithIcon("", theme.NavigateBackIcon(), func() {
		if currentIdx > 0 {
			currentIdx--
			r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
			loadForm(r)
		}
	})
	btnNext = widget.NewButtonWithIcon("", theme.NavigateNextIcon(), func() {
		if currentIdx < len(allCodes)-1 {
			currentIdx++
			r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
			loadForm(r)
		}
	})

	// ── ปุ่ม Recalculate ──
	btnRecalc := widget.NewButtonWithIcon("ประมวลผลยอด", theme.ViewRefreshIcon(), func() {
		var d dialog.Dialog
		yesBtn := newEnterEscButton("ยืนยัน (Enter)", func() {
			d.Hide()
			prog := dialog.NewProgressInfinite("กำลังประมวลผล", "ระบบกำลังคำนวณยอดบัญชี กรุณารอสักครู่...", w)
			prog.Show()

			go func() {
				err := RecalculateLedgerMaster(xlOptions)
				fyne.Do(func() {
					prog.Hide()
					if err != nil {
						// [แก้ไข 1] เปลี่ยน ShowError เป็น Custom Dialog ที่รองรับ Enter
						var errD dialog.Dialog
						okBtn := newEnterButton("OK (Enter)", func() {
							errD.Hide()
							// คืน Focus กลับไปที่ปุ่ม (หรือช่องกรอก)
							w.Canvas().Focus(enAcCode)
						})
						errD = dialog.NewCustomWithoutButtons("Error",
							container.NewVBox(
								widget.NewLabel(fmt.Sprintf("ประมวลผลไม่สำเร็จ: %v", err)),
								container.NewCenter(okBtn),
							), w)
						errD.Show()
						go func() {
							time.Sleep(50 * time.Millisecond)
							fyne.Do(func() { w.Canvas().Focus(okBtn) })
						}()
					} else {
						if len(allCodes) > 0 {
							r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
							loadForm(r)
						}
						// [แก้ไข 2] เปลี่ยน ShowInformation เป็น Custom Dialog ที่รองรับ Enter
						var infoD dialog.Dialog
						okBtn := newEnterButton("OK (Enter)", func() {
							infoD.Hide()
							w.Canvas().Focus(enAcCode)
						})
						infoD = dialog.NewCustomWithoutButtons("สำเร็จ",
							container.NewVBox(
								widget.NewLabel("ประมวลผลยอดบัญชีเสร็จสมบูรณ์"),
								container.NewCenter(okBtn),
							), w)
						infoD.Show()
						go func() {
							time.Sleep(50 * time.Millisecond)
							fyne.Do(func() { w.Canvas().Focus(okBtn) })
						}()
					}
				})
			}()
		}, func() {
			// กรณีผู้ใช้กด Esc ที่ปุ่มยืนยัน
			d.Hide()
			w.Canvas().Focus(enAcCode)
		})

		noBtn := newEscButton("ยกเลิก (Esc)", func() {
			d.Hide()
			w.Canvas().Focus(enAcCode)
		})

		d = dialog.NewCustomWithoutButtons("ยืนยันการประมวลผล",
			container.NewVBox(
				widget.NewLabel("ระบบจะทำการคำนวณยอดบัญชีรายเดือน (Thisper01-12) ใหม่ทั้งหมด\nโดยดึงข้อมูลจากสมุดรายวัน (Book_items)\n\nคุณต้องการดำเนินการต่อหรือไม่?"),
				container.NewCenter(container.NewHBox(yesBtn, noBtn)),
			), w)
		d.Show()

		go func() {
			time.Sleep(50 * time.Millisecond)
			fyne.Do(func() { w.Canvas().Focus(yesBtn) })
		}()
	})

	// ── Ac_code OnSubmitted → lookup Gname ──
	enAcCode.onSave = saveAction
	enAcCode.OnSubmitted = func(s string) {
		s = strings.TrimSpace(s)
		if len(s) < 3 {
			return
		}
		gcode := s[:3]
		gname := lookupGname(xlOptions, gcode)
		if gname == "" {
			enAcCode.SetText("")
			showErr(fmt.Sprintf("ไม่พบ Account Group: %s", gcode), enAcCode)
			return
		}
		lblGname.SetText(gname)

		if actionFlag == "ADD" {
			_, found := loadLedgerRecord(xlOptions, s)
			if found {
				enAcCode.SetText("")
				showErr(fmt.Sprintf("Ac_code '%s' มีอยู่แล้ว", s), enAcCode)
				return
			}
		}
		w.Canvas().Focus(enAcName)
	}

	// ── Lastper12 live format + sync → Bthisyear ──────────────────
	// ใช้ flag กัน recursive loop
	lastper12Updating := false
	enLastper[11].OnChanged = func(s string) {
		if lastper12Updating {
			return
		}
		// กรอง: รับเฉพาะตัวเลข 0-9, จุดทศนิยม, และ "-" นำหน้า
		raw := ""
		hasDot := false
		hasMinus := false
		for idx, r := range s {
			if r == '-' && idx == 0 {
				hasMinus = true
				raw += string(r)
			} else if r >= '0' && r <= '9' {
				raw += string(r)
			} else if r == '.' && !hasDot {
				raw += string(r)
				hasDot = true
			}
		}
		// แยก integer / decimal
		intPart, decPart := raw, ""
		sign := ""
		if strings.HasPrefix(raw, "-") {
			sign = "-"
			intPart = raw[1:]
		}
		if dotIdx := strings.Index(intPart, "."); dotIdx >= 0 {
			decPart = intPart[dotIdx:]
			intPart = intPart[:dotIdx]
		}
		// format comma บน integer part
		formatted := intPart
		if v, err := strconv.ParseInt(intPart, 10, 64); err == nil && intPart != "" {
			tmp := fmt.Sprintf("%d", v)
			res := ""
			for i, ch := range tmp {
				if i > 0 && (len(tmp)-i)%3 == 0 {
					res += ","
				}
				res += string(ch)
			}
			formatted = res
		}
		formatted = sign + formatted + decPart
		if formatted != s && formatted != "-" {
			lastper12Updating = true
			enLastper[11].SetText(formatted)
			lastper12Updating = false
		}
		// sync Bthisyear = Lastper12 (ใส่ comma ให้สวยงาม)
		rawVal := strings.ReplaceAll(formatted, ",", "")
		if !hasMinus || rawVal == "-" {
			rawVal = strings.TrimPrefix(rawVal, "-")
		}
		if rawVal == "" {
			rawVal = "0"
		}

		lastper12Updating = true
		if v, err := strconv.ParseFloat(rawVal, 64); err == nil {
			enBthisyear.SetText(ledgerNum(v))
		} else {
			enBthisyear.SetText(rawVal)
		}
		lastper12Updating = false
	}
	enLastper[11].onFocusLost = func() {
		raw := strings.ReplaceAll(enLastper[11].Text, ",", "")
		if raw == "" {
			return
		}
		if v, err := strconv.ParseFloat(raw, 64); err == nil {
			lastper12Updating = true
			enLastper[11].SetText(ledgerNum(v))
			enBthisyear.SetText(ledgerNum(v))
			lastper12Updating = false
		}
	}

	// ── Register Ctrl+S ──
	registerCtrlS(w, saveAction, func() {
		if actionFlag == "ADD" || actionFlag == "EDIT" {
			if len(allCodes) > 0 {
				r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
				loadForm(r)
			} else {
				loadForm(emptyLedger())
			}
			setViewMode()
		}
	})

	// ── Layout ──
	comName := ""
	ff, errr := excelize.OpenFile(currentDBPath, xlOptions)
	if errr == nil {
		comName, _ = ff.GetCellValue("Company_Profile", "B2")
		ff.Close()
	}

	toolbar := container.NewHBox(
		btnSave, btnAdd, btnEdit, btnDel,
		widget.NewButtonWithIcon("", theme.SearchIcon(), func() {
			showLedgerSearch(w, xlOptions, allCodes, func(r LedgerRecord, idx int) {
				currentIdx = idx
				full, found := loadLedgerRecord(xlOptions, r.AcCode)
				if found {
					loadForm(full)
				}
				setViewMode()
			})
		}),
		btnPrev, btnNext,
		widget.NewSeparator(),
		btnRecalc,
		layout.NewSpacer(),
		modeLabel,
		widget.NewLabelWithStyle(comName, fyne.TextAlignTrailing, fyne.TextStyle{Bold: true}),
	)

	headerForm := container.New(layout.NewFormLayout(),
		widget.NewLabel("Account Code"), container.NewHBox(
			container.NewGridWrap(fyne.NewSize(150, 32), enAcCode),
			widget.NewLabel("Account Group:"),
			lblGname,
		),
		widget.NewLabel("Account Name"), enAcName,
		widget.NewLabel("Period Beginning"), container.NewHBox(
			container.NewGridWrap(fyne.NewSize(150, 32), enBBal),
			widget.NewLabel("Debit :"),
			container.NewGridWrap(fyne.NewSize(150, 32), enDebit),
		),
		widget.NewLabel("Period Closing"), container.NewHBox(
			container.NewGridWrap(fyne.NewSize(150, 32), enCBal),
			widget.NewLabel("Credit :"),
			container.NewGridWrap(fyne.NewSize(150, 32), enCredit),
		),
	)

	periodHeader := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("")),
		container.NewGridWrap(fyne.NewSize(200, 30), widget.NewLabelWithStyle("This Year", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(200, 30), widget.NewLabelWithStyle("Last Year", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})),
	)

	periodRows := container.NewVBox(periodHeader)
	labels := []string{
		"Beginning:", "Period 1:", "Period 2:", "Period 3:", "Period 4:",
		"Period 5:", "Period 6:", "Period 7:", "Period 8:", "Period 9:",
		"Period 10:", "Period 11:", "Period 12:",
	}

	thisYearFields := []fyne.CanvasObject{enBthisyear}
	for i := 0; i < 12; i++ {
		thisYearFields = append(thisYearFields, enThisper[i])
	}
	lastYearFields := []fyne.CanvasObject{enBlastyear}
	for i := 0; i < 12; i++ {
		lastYearFields = append(lastYearFields, enLastper[i])
	}

	for i := 0; i < 13; i++ {
		lbl := widget.NewLabelWithStyle(labels[i], fyne.TextAlignTrailing, fyne.TextStyle{})
		row := container.NewHBox(
			container.NewGridWrap(fyne.NewSize(90, 30), lbl),
			container.NewGridWrap(fyne.NewSize(200, 30), thisYearFields[i]),
			container.NewGridWrap(fyne.NewSize(200, 30), lastYearFields[i]),
		)
		periodRows.Add(row)
	}

	topSection := container.NewVBox(
		toolbar,
		widget.NewLabelWithStyle("Account Code", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		headerForm,
		widget.NewSeparator(),
		widget.NewLabelWithStyle("Period", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
	)

	// โหลด record แรก
	if len(allCodes) > 0 {
		r, _ := loadLedgerRecord(xlOptions, allCodes[0])
		loadForm(r)
	}
	setViewMode()

	// ── Refresh callback (เรียกจาก book_ui หลัง POST) ──
	refreshCallback := createRefreshCallback(xlOptions,
		func() string { return enAcCode.Text },
		enBBal, enCBal, enDebit, enCredit)
	refreshLedgerFunc = refreshCallback

	reset := func() {
		allCodes = loadAllAcCodes(xlOptions) // reload เผื่อมีเพิ่ม/ลบ
		if len(allCodes) > 0 {
			// clamp
			if currentIdx >= len(allCodes) {
				currentIdx = len(allCodes) - 1
			}
			r, _ := loadLedgerRecord(xlOptions, allCodes[currentIdx])
			loadForm(r)
		} else {
			loadForm(emptyLedger())
		}
		setViewMode()
		fyne.Do(func() { w.Canvas().Focus(enAcCode) })
	}

	return container.NewBorder(topSection, nil, nil, nil,
		container.NewVScroll(periodRows),
	), reset
}

// ─────────────────────────────────────────────────────────────────
// refreshLedgerRecord — อัปเดต BBal/CBal/Debit/Credit ใน form
// ─────────────────────────────────────────────────────────────────
func refreshLedgerRecord(xlOptions excelize.Options, acCode string,
	enBBal, enCBal, enDebit, enCredit *smartEntry) error {

	if acCode == "" {
		return fmt.Errorf("acCode ว่าง")
	}
	_, found := loadLedgerRecord(xlOptions, acCode)
	if !found {
		return fmt.Errorf("ไม่พบ Account Code: %s", acCode)
	}

	bbal, cbal, dr, cr := computeRealTimeLedger(xlOptions, acCode)
	enBBal.SetText(ledgerNum(bbal))
	enCBal.SetText(ledgerNum(cbal))
	enDebit.SetText(ledgerNum(dr))
	enCredit.SetText(ledgerNum(cr))
	return nil
}

// ─────────────────────────────────────────────────────────────────
// createRefreshCallback — สร้าง callback สำหรับ book_ui เรียก
// ─────────────────────────────────────────────────────────────────
func createRefreshCallback(xlOptions excelize.Options, currentAcCode func() string,
	enBBal, enCBal, enDebit, enCredit *smartEntry) func() {

	return func() {
		acCode := currentAcCode()
		if acCode == "" {
			return
		}
		if err := refreshLedgerRecord(xlOptions, acCode, enBBal, enCBal, enDebit, enCredit); err != nil {
			fmt.Printf("Refresh failed: %v\n", err)
			return
		}
		fmt.Printf("✅ Refreshed Ledger: %s\n", acCode)
	}
}

// ─────────────────────────────────────────────────────────────────
// min3 — helper: min(n, 3)
// ─────────────────────────────────────────────────────────────────
func min3(n int) int {
	if n < 3 {
		return n
	}
	return 3
}
