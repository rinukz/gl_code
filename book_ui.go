package main

import (
	"fmt"
	"image/color"
	"math"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

var bperiodNum int

// ===book_ui.go===update@03-08-2026 เวอร์ชั่นล่าสุด=======

// ─────────────────────────────────────────────────────────────────
// BookLine — เพิ่ม IsVATLine, ParentBline, Bperiod สำหรับ VAT tracking
// และแยก item ต่าง period ไม่ให้ทับกัน
//
//	IsVATLine   = true  → line นี้คือ VAT auto-gen line
//	ParentBline = Bline ของ parent line ที่สร้าง VAT นี้
//	             (ใช้ค้นหาตอน EDIT เพื่ออัพเดท VAT ตาม)
//	Bperiod     = period number ที่ item นี้สังกัด (1-based)
//	             KEY สำหรับ DELETE: comCode + bitem + bperiod
//
// ─────────────────────────────────────────────────────────────────
type BookLine struct {
	Comcode     string
	Bdate       string
	Bvoucher    string
	Bitem       string
	Bline       int
	AcCode      string
	AcName      string
	Scode       string
	Sname       string
	Bdebit      string
	Bcredit     string
	Bref        string
	Boff        string
	Bcomtaxid   string
	Bnote       string
	Bnote2      string
	Bchqno      string
	Bchqdate    string
	IsVATLine   bool // col R: "1" = VAT line, "" = ปกติ
	ParentBline int  // col S: Bline ของ parent (เฉพาะ VAT line)
	Bperiod     int  // col V (21): period number ที่ item นี้สังกัด
}

// ─────────────────────────────────────────────────────────────────
// Excel column mapping (0-based index ใน row slice):
//  0=Comcode 1=Bdate 2=Bvoucher 3=Bitem 4=Bline
//  5=Ac_code 6=Ac_name 7=Scode 8=Sname
//  9=Bdebit 10=Bcredit 11=Bref 12=Boff
//  13=Bcomtaxid 14=Bnote 15=Bchqno 16=Bchqdate
//  17=IsVATLine 18=ParentBline 19=Posted 20=Bnote2
//  21=Bperiod  ← col ใหม่ ใช้แยก item ต่าง period ไม่ให้ทับกัน
// ─────────────────────────────────────────────────────────────────

// deletableList — intercept F2/F4/Esc
type deletableList struct {
	widget.List
	onF2Key  func()
	onF4Key  func()
	onEscKey func()
}

func newDeletableList(
	length func() int,
	create func() fyne.CanvasObject,
	update func(widget.ListItemID, fyne.CanvasObject),
	onF2Key func(),
	onF4Key func(),
	onEscKey func(),
) *deletableList {
	dl := &deletableList{
		onF2Key:  onF2Key,
		onF4Key:  onF4Key,
		onEscKey: onEscKey,
	}
	dl.List.Length = length
	dl.List.CreateItem = create
	dl.List.UpdateItem = update
	dl.ExtendBaseWidget(dl)
	return dl
}

func (dl *deletableList) TypedKey(ev *fyne.KeyEvent) {
	switch ev.Name {
	case fyne.KeyF2:
		if dl.onF2Key != nil {
			dl.onF2Key()
		}
	case fyne.KeyF4:
		if dl.onF4Key != nil {
			dl.onF4Key()
		}
	case fyne.KeyEscape:
		if dl.onEscKey != nil {
			dl.onEscKey()
		}
	default:
		dl.List.TypedKey(ev)
	}
}

// ─────────────────────────────────────────────────────────────────

// upsertCustomer — บันทึก TaxID → CustName ลง Customer_Log (ใช้ *excelize.File ที่เปิดอยู่แล้ว)
// [DONE] upsertCustomer — save/update TaxID+CustName to Customer_Log sheet, working correctly
func upsertCustomer(f *excelize.File, comCode, taxID, custName string) {
	if taxID == "" || custName == "" {
		return
	}
	const sh = "Customer_Log"
	if idx, _ := f.GetSheetIndex(sh); idx == -1 {
		f.NewSheet(sh)
		f.SetCellValue(sh, "A1", "Comcode")
		f.SetCellValue(sh, "B1", "TaxID")
		f.SetCellValue(sh, "C1", "CustName")
	}
	rows, _ := f.GetRows(sh)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 2 && row[1] == taxID {
			if safeGet(row, 2) != custName {
				f.SetCellValue(sh, fmt.Sprintf("C%d", i+1), custName)
			}
			return
		}
	}
	newRow := len(rows) + 1
	f.SetCellValue(sh, fmt.Sprintf("A%d", newRow), comCode)
	f.SetCellValue(sh, fmt.Sprintf("B%d", newRow), taxID)
	f.SetCellValue(sh, fmt.Sprintf("C%d", newRow), custName)
}

// lookupCustomerName — ค้นหา CustName จาก TaxID
// [DONE] lookupCustomerName — lookup customer name from Customer_Log by TaxID, working correctly
func lookupCustomerName(xlOptions excelize.Options, taxID string) string {
	if taxID == "" {
		return ""
	}
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return ""
	}
	defer f.Close()
	rows, _ := f.GetRows("Customer_Log")
	for i, row := range rows {
		if i == 0 {
			continue
		}
		// match แค่ TaxID — universal
		if len(row) >= 3 && row[1] == taxID {
			return row[2]
		}
	}
	return ""
}

func getBookGUI(w fyne.Window) (fyne.CanvasObject, func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	comCode := getComCodeFromExcel(xlOptions)

	// ─── Period config (อ่านครั้งเดียวตอน init) ───────────────────
	// เทียบกับ Fox Pro: GOTO I_PEMNOW → DR1=PSAT, DR2=PEND
	periodCfg, periodErr := loadCompanyPeriod(xlOptions)
	var dr1, dr2 time.Time
	if periodErr == nil {
		dr1, dr2, _ = getCurrentPeriodRange(periodCfg.YearEnd, periodCfg.TotalPeriods, periodCfg.NowPeriod)
	}

	var lines []BookLine
	warningConfirmed := false
	actionFlag := "VIEW"
	currentBitem := ""
	currentBperiod := 0  // period number ของ item ที่กำลัง view/edit (ใช้เป็น DELETE key)
	prevBitem := ""      // จำ bitem ก่อน ADD เพื่อใช้ตอน Esc cancel
	selectedLineID := -1 // -1=add new, >=0=edit index นั้น

	var saveAction func()

	var checkAndSave func()

	modeLabel := widget.NewLabelWithStyle("● VIEW MODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})
	// แสดงงวดปัจจุบันใน toolbar เช่น "งวด 3/12  (01/03/26 - 31/03/26)"
	periodLabel := widget.NewLabel("")
	if periodErr == nil {
		periodLabel.SetText(getPeriodSummary(periodCfg))
	} else {
		periodLabel.SetText("⚠ ไม่พบข้อมูล Period")
	}

	// --- Header Fields ---
	enBitem := widget.NewLabel("")
	enBdate := newDateEntry()
	enBvoucher := newSmartEntry(func() { checkAndSave() })
	enBref := newSmartEntry(func() { checkAndSave() })
	enBchqno := newSmartEntry(func() { checkAndSave() })
	enBchqdate := newDateEntry()
	enBoff := newSmartEntry(func() { checkAndSave() })
	enBcomtaxid := newSmartEntry(func() { checkAndSave() })
	taxLabel := widget.NewLabel("")
	enBnote := newSmartEntry(func() { checkAndSave() })
	enBnote2 := newSmartEntry(func() { checkAndSave() })

	// --- Line Item Fields ---
	enAcCode := newSmartEntry(func() { checkAndSave() })
	enAcName := newSmartEntry(func() { checkAndSave() })
	enScode := newSmartEntry(func() { checkAndSave() })
	enSname := newSmartEntry(func() { checkAndSave() })
	enBdebit := newSmartEntry(func() { checkAndSave() })
	enBcredit := newSmartEntry(func() { checkAndSave() })

	vatSelect := widget.NewSelect([]string{"No-VAT", "Debit VAT", "Credit VAT", "Both VAT"}, nil)
	vatSelect.SetSelectedIndex(0)
	enVatPct := newSmartEntry(nil)
	enVatPct.SetText("7")
	enVatPct.Disable()

	enBvoucher.OnChanged = func(s string) {
		corrected := forceEnglishVoucher(s)
		if corrected != s {
			enBvoucher.SetText(corrected)
		}
	}
	vatPctText := "7" // เก็บค่าแยก เพราะ Fyne clear Entry.Text ตอน Disable()
	vatSelect.OnChanged = func(s string) {
		if vatSelect.SelectedIndex() != 0 {
			enVatPct.Enable()
			enVatPct.SetText(vatPctText) // restore ค่าที่เก็บไว้
		} else {
			if strings.TrimSpace(enVatPct.Text) != "" {
				vatPctText = enVatPct.Text // บันทึกก่อน Disable
			}
			enVatPct.Disable()
		}
	}
	enVatPct.OnSubmitted = func(s string) { w.Canvas().Focus(enBdebit) }
	enVatPct.OnChanged = func(s string) {
		if strings.TrimSpace(s) != "" {
			vatPctText = s
		}
	}

	// ── WHT: No-WHT | 120WHT (ถูกหัก-สินทรัพย์) | 235WHT (หัก-หนี้สิน) ──
	whtSelect := widget.NewSelect([]string{"No-WHT", "120WHT (ถูกหัก)", "235WHT (หัก ณ ที่จ่าย)"}, nil)
	whtSelect.SetSelectedIndex(0)
	enWhtPct := newSmartEntry(nil)
	enWhtPct.SetText("3")
	enWhtPct.Disable()
	whtPctText := "3"
	whtSelect.OnChanged = func(s string) {
		if whtSelect.SelectedIndex() != 0 {
			enWhtPct.Enable()
			enWhtPct.SetText(whtPctText)
		} else {
			if strings.TrimSpace(enWhtPct.Text) != "" {
				whtPctText = enWhtPct.Text
			}
			enWhtPct.Disable()
		}
	}
	enWhtPct.OnSubmitted = func(s string) { w.Canvas().Focus(enBdebit) }
	enWhtPct.OnChanged = func(s string) {
		if strings.TrimSpace(s) != "" {
			whtPctText = s
		}
	}

	lblSumDebit := widget.NewLabel("0.00")
	lblSumCredit := widget.NewLabel("0.00")
	lblDiff := canvas.NewText("ส่วนต่าง: 0.00", theme.ForegroundColor())
	lblDiff.TextStyle = fyne.TextStyle{Bold: true}
	lblDiff.TextSize = 14

	var lineList *deletableList

	calcSum := func() {
		var sumD, sumC float64
		for _, l := range lines {
			sumD += parseFloat(l.Bdebit)
			sumC += parseFloat(l.Bcredit)
		}
		lblSumDebit.SetText(bsNum(sumD))
		lblSumCredit.SetText(bsNum(sumC))

		diff := sumD - sumC
		if diff < 0 {
			diff = -diff
		}
		if diff < 0.005 {
			lblDiff.Text = "✓ Balance"
			lblDiff.Color = color.RGBA{R: 0x0D, G: 0x54, B: 0x2B, A: 0xFF}
		} else {
			lblDiff.Text = fmt.Sprintf("ส่วนต่าง: %s", bsNum(diff))
			lblDiff.Color = color.RGBA{R: 0xC0, G: 0x20, B: 0x20, A: 0xFF}
		}
		lblDiff.Refresh()
	}

	clearLineInput := func() {
		selectedLineID = -1
		enAcCode.SetText("")
		enAcName.SetText("")
		enScode.SetText("")
		enSname.SetText("")
		enBdebit.SetText("0.00")
		enBcredit.SetText("0.00")
		vatSelect.SetSelectedIndex(0)
		whtSelect.SetSelectedIndex(0)
		lineList.UnselectAll()
	}

	clearHeader := func() {
		enBdate.SetDate("")
		enBvoucher.SetText("")
		enBref.SetText("")
		enBchqno.SetText("")
		enBchqdate.SetDate("")
		enBoff.SetText("")
		enBcomtaxid.SetText("")
		enBnote.SetText("")
		enBnote2.SetText("")
		lines = nil
		calcSum()
	}

	// ─────────────────────────────────────────────────
	// loadBitem — อ่าน col 17=IsVATLine, 18=ParentBline, 21=Bperiod
	// ─────────────────────────────────────────────────
	loadBitem := func(bitem string, targetPeriod int) { // ✅ รับ targetPeriod
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()

		rows, _ := f.GetRows("Book_items")
		lines = nil
		headerLoaded := false
		for i, row := range rows {
			if i == 0 || len(row) < 5 || row[0] != comCode || row[3] != bitem {
				continue
			}

			var bperiodNum int
			if len(row) > 21 {
				fmt.Sscanf(safeGet(row, 21), "%d", &bperiodNum)
			}

			// ✅ กรองให้ตรงกับ Period ที่ต้องการโหลด
			if bperiodNum != 0 && bperiodNum != targetPeriod {
				continue
			}

			bl := BookLine{
				Comcode:   row[0],
				Bdate:     safeGet(row, 1),
				Bvoucher:  safeGet(row, 2),
				Bitem:     row[3],
				AcCode:    safeGet(row, 5),
				AcName:    safeGet(row, 6),
				Scode:     safeGet(row, 7),
				Sname:     safeGet(row, 8),
				Bdebit:    safeGet(row, 9),
				Bcredit:   safeGet(row, 10),
				Bref:      safeGet(row, 11),
				Boff:      safeGet(row, 12),
				Bcomtaxid: safeGet(row, 13),
				Bnote:     safeGet(row, 14),
				Bchqno:    safeGet(row, 15),
				Bchqdate:  safeGet(row, 16),
				Bnote2:    safeGet(row, 20),
				IsVATLine: safeGet(row, 17) == "1",
			}
			var parentBline int
			fmt.Sscanf(safeGet(row, 18), "%d", &parentBline)
			bl.ParentBline = parentBline

			var blineNum int
			fmt.Sscanf(safeGet(row, 4), "%d", &blineNum)
			bl.Bline = blineNum
			bl.Bperiod = bperiodNum

			if !headerLoaded {
				enBdate.SetDate(bl.Bdate)
				enBvoucher.SetText(bl.Bvoucher)
				enBref.SetText(bl.Bref)
				enBchqno.SetText(bl.Bchqno)
				enBchqdate.SetDate(bl.Bchqdate)
				enBoff.SetText(bl.Boff)
				enBcomtaxid.SetText(bl.Bcomtaxid)
				enBnote.SetText(bl.Bnote)
				enBnote2.SetText(bl.Bnote2)
				headerLoaded = true
			}
			lines = append(lines, bl)
		}
		currentBitem = bitem
		currentBperiod = targetPeriod // ✅ อัปเดต currentBperiod เสมอ
		enBitem.SetText(bitem)
		selectedLineID = -1
		calcSum()
		lineList.Refresh()
	}

	autoNextBitem := func(targetPeriod int) string { // ✅ รับ targetPeriod
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return "001"
		}
		defer f.Close()
		rows, _ := f.GetRows("Book_items")
		maxItem := 0
		for i, row := range rows {
			if i == 0 || len(row) < 4 || row[0] != comCode {
				continue
			}
			var bperiod int
			if len(row) > 21 {
				fmt.Sscanf(safeGet(row, 21), "%d", &bperiod)
			}
			// ✅ นับเฉพาะใน Period ที่กำลังจะ Add
			if bperiod == targetPeriod {
				n := 0
				fmt.Sscanf(row[3], "%d", &n)
				if n > maxItem {
					maxItem = n
				}
			}
		}
		return fmt.Sprintf("%03d", maxItem+1)
	}

	var btnSave, btnAdd, btnEdit, btnDel, btnCheck *widget.Button
	var btnPrev, btnNext, btnRepair, btnVoid *widget.Button

	// ─────────────────────────────────────────────────
	// applyAutoRounding — คำนวณและสร้าง rounding line real-time
	// ตรรกะ:
	//   - คำนวณส่วนต่าง rawDiff = sumD - sumC (ก่อนปัดเศษ)
	//   - ดูทศนิยมตำแหน่งที่ 3 (เช่น 0.125 → ตำแหน่งที่ 3 = 5)
	//   - ถ้า >= 5 → ปัดขึ้นปกติ (math.Round)
	//   - ถ้า < 5  → ตัดทิ้ง (math.Floor / floorVAT logic)
	// ─────────────────────────────────────────────────
	applyAutoRounding := func() {
		// 1. ตรวจสอบว่าบิลนี้มีการใช้ Auto VAT หรือ Auto WHT หรือไม่
		hasAutoCalc := false
		for _, l := range lines {
			if l.IsVATLine && l.AcCode != "524RND" && l.AcCode != "450RND" {
				hasAutoCalc = true
				break
			}
		}

		// 2. ลบ rounding line เก่าออกก่อน (ลบเฉพาะที่ระบบสร้างให้ คือ IsVATLine = true)
		for i := len(lines) - 1; i >= 0; i-- {
			if lines[i].IsVATLine && (lines[i].AcCode == "524RND" || lines[i].AcCode == "450RND") {
				lines = append(lines[:i], lines[i+1:]...)
			}
		}

		// 3. ถ้าไม่มี Auto VAT/WHT เลย ให้หยุดทำงาน
		if !hasAutoCalc {
			return
		}

		// 4. คำนวณยอดรวมใหม่ทั้งหมด (รวมถึงบรรทัดที่ user อาจจะ manual คีย์ 524/450 เองด้วย)
		var sumD, sumC float64
		for _, l := range lines {
			sumD += parseFloat(l.Bdebit)
			sumC += parseFloat(l.Bcredit)
		}

		diff := math.Round((sumD-sumC)*100) / 100

		if diff != 0 && math.Abs(diff) <= 0.99 {
			var scode, sname string
			for _, l := range lines {
				if !l.IsVATLine {
					scode = l.Scode
					sname = l.Sname
					break
				}
			}
			roundLine := BookLine{
				Comcode:   comCode,
				Bdate:     enBdate.GetDate(),
				Bvoucher:  enBvoucher.Text,
				Bitem:     currentBitem,
				Bline:     len(lines) + 1,
				Scode:     scode,
				Sname:     sname,
				IsVATLine: true, // ✅ ให้ระบบรู้ว่านี่คือบรรทัดที่ Auto-gen ขึ้นมา
			}
			if diff > 0 {
				roundLine.AcCode = "450RND"
				roundLine.AcName = "กำไรจากการปัดเศษ"
				roundLine.Bcredit = fmt.Sprintf("%.2f", diff)
				roundLine.Bdebit = "0.00"
			} else {
				roundLine.AcCode = "524RND"
				roundLine.AcName = "ขาดทุนจากการปัดเศษ"
				roundLine.Bdebit = fmt.Sprintf("%.2f", -diff)
				roundLine.Bcredit = "0.00"
			}
			lines = append(lines, roundLine)
		}
	}

	// // ─────────────────────────────────────────────────
	// // applyAutoRounding — คำนวณและสร้าง rounding line real-time
	// // ตรรกะ:
	// //   - คำนวณส่วนต่าง rawDiff = sumD - sumC (ก่อนปัดเศษ)
	// //   - ดูทศนิยมตำแหน่งที่ 3 (เช่น 0.125 → ตำแหน่งที่ 3 = 5)
	// //   - ถ้า >= 5 → ปัดขึ้นปกติ (math.Round)
	// //   - ถ้า < 5  → ตัดทิ้ง (math.Floor / floorVAT logic)
	// // ─────────────────────────────────────────────────
	// applyAutoRounding := func() {
	// 	var sumD, sumC float64
	// 	for _, l := range lines {
	// 		if l.AcCode != "524RND" && l.AcCode != "450RND" {
	// 			sumD += parseFloat(l.Bdebit)
	// 			sumC += parseFloat(l.Bcredit)
	// 		}
	// 	}

	// 	// คำนวณส่วนต่าง raw (ยังไม่ปัดเศษ)
	// 	rawDiff := sumD - sumC

	// 	// ดึงเลขทศนิยมตำแหน่งที่ 3 เพื่อตัดสินใจวิธีปัดเศษ
	// 	absDiff := math.Abs(rawDiff)
	// 	// เดิม
	// 	// thirdDecimal := int(absDiff*1000) % 10

	// 	// ใหม่ (แก้แล้ว)
	// 	thirdDecimal := int(math.Round(absDiff*1000)) % 10

	// 	var diff float64
	// 	if thirdDecimal >= 5 {
	// 		// ทศนิยมตำแหน่งที่ 3 >= 5 → ปัดขึ้นปกติ
	// 		diff = math.Round(rawDiff*100) / 100
	// 	} else {
	// 		// ทศนิยมตำแหน่งที่ 3 < 5 → ตัดทิ้ง (floor)
	// 		if rawDiff >= 0 {
	// 			diff = math.Floor(rawDiff*100) / 100
	// 		} else {
	// 			diff = -math.Floor(-rawDiff*100) / 100
	// 		}
	// 	}

	// 	// ลบ rounding line เก่าออกก่อน
	// 	for i := len(lines) - 1; i >= 0; i-- {
	// 		if lines[i].AcCode == "524RND" || lines[i].AcCode == "450RND" {
	// 			lines = append(lines[:i], lines[i+1:]...)
	// 		}
	// 	}

	// 	if diff != 0 && math.Abs(diff) <= 0.99 {
	// 		// ดึง Scode/Sname จาก line แรกที่ไม่ใช่ rounding
	// 		var scode, sname string
	// 		for _, l := range lines {
	// 			if l.AcCode != "524RND" && l.AcCode != "450RND" {
	// 				scode = l.Scode
	// 				sname = l.Sname
	// 				break
	// 			}
	// 		}
	// 		roundLine := BookLine{
	// 			Comcode:   comCode,
	// 			Bdate:     enBdate.GetDate(),
	// 			Bvoucher:  enBvoucher.Text,
	// 			Bitem:     currentBitem,
	// 			Bline:     len(lines) + 1,
	// 			Scode:     scode,
	// 			Sname:     sname,
	// 			IsVATLine: true,
	// 		}
	// 		if diff > 0 {
	// 			roundLine.AcCode = "450RND"
	// 			roundLine.AcName = "กำไรจากการปัดเศษ"
	// 			roundLine.Bcredit = fmt.Sprintf("%.2f", diff)
	// 			roundLine.Bdebit = "0.00"
	// 		} else {
	// 			roundLine.AcCode = "524RND"
	// 			roundLine.AcName = "ขาดทุนจากการปัดเศษ"
	// 			roundLine.Bdebit = fmt.Sprintf("%.2f", -diff)
	// 			roundLine.Bcredit = "0.00"
	// 		}
	// 		lines = append(lines, roundLine)
	// 	}
	// }

	// ─────────────────────────────────────────────────
	// execDeleteLine — ลบ line + VAT line ที่ผูกอยู่ (ถ้ามี)
	// ─────────────────────────────────────────────────
	execDeleteLine := func() {
		if actionFlag == "VIEW" {
			return
		}
		if selectedLineID < 0 || selectedLineID >= len(lines) {
			return
		}
		sel := lines[selectedLineID]

		// ✅ แทนที่ตรงนี้ครับ: ถ้าเลือก VAT line → กระโดดไป parent แทน (ยกเว้นบรรทัดปัดเศษ)
		if sel.IsVATLine && sel.AcCode != "524RND" && sel.AcCode != "450RND" {
			dialog.ShowInformation("แจ้งเตือน",
				"บรรทัดนี้คือ VAT auto-gen\nหากต้องการลบ ให้ลบ parent line แทน", w)
			return
		}

		// นับว่ามี VAT line ผูกอยู่ไหม (ส่วนนี้คงไว้เหมือนเดิม)
		vatMsg := ""
		for _, l := range lines {
			if l.IsVATLine && l.ParentBline == sel.Bline {
				vatMsg = "\n(จะลบ VAT line ที่ผูกอยู่ด้วย)"
				break
			}
		}

		targetCode := sel.AcCode
		var confirmD dialog.Dialog
		yesBtn := newEnterButton("ลบ", func() {
			confirmD.Hide()
			// ลบ VAT line ที่ผูกอยู่ก่อน (scan จาก tail เพื่อ index ไม่เพี้ยน)
			for i := len(lines) - 1; i >= 0; i-- {
				if lines[i].IsVATLine && lines[i].ParentBline == sel.Bline {
					lines = append(lines[:i], lines[i+1:]...)
				}
			}
			// ลบ parent line
			lines = append(lines[:selectedLineID], lines[selectedLineID+1:]...)

			applyAutoRounding() // ระบบจะคำนวณใหม่หลังลบเสร็จ
			calcSum()
			lineList.Refresh()
			clearLineInput()
			w.Canvas().Focus(enAcCode)
		})

		noBtn := newEnterButton("ยกเลิก", func() {
			confirmD.Hide()
			w.Canvas().Focus(enAcCode)
		})
		yesBtn.Importance = widget.DangerImportance
		content := container.NewVBox(
			widget.NewLabel(fmt.Sprintf("ลบบรรทัด [%s] ?%s", targetCode, vatMsg)),
			widget.NewLabel("(ยังไม่ได้บันทึกลง Excel)"),
			container.NewCenter(container.NewHBox(yesBtn, noBtn)),
		)
		confirmD = dialog.NewCustomWithoutButtons("ยืนยัน", content, w)
		confirmD.Show()
	}

	var setViewMode func()

	cancelEdit := func() {
		if actionFlag == "VIEW" {
			return
		}
		if actionFlag == "ADD" {
			if prevBitem != "" {
				currentBitem = prevBitem
				loadBitem(prevBitem, currentBperiod) // ✅ เติม , currentBperiod
			} else {
				currentBitem = ""
				enBitem.SetText("")
				clearHeader()
				lines = nil
				lineList.Refresh()
			}
		} else if currentBitem != "" {
			loadBitem(currentBitem, currentBperiod) // ✅ เติม , currentBperiod
		}
		setViewMode()
	}

	enterEditMode := func() {
		if actionFlag != "VIEW" || currentBitem == "" {
			return
		}
		btnEdit.OnTapped()
	}

	setViewMode = func() {
		actionFlag = "VIEW"
		modeLabel.SetText("● VIEW MODE")
		enBdate.Disable()
		enBvoucher.Disable()
		enBref.Disable()
		enBchqno.Disable()
		enBchqdate.Disable()
		enBoff.Disable()
		enBcomtaxid.Disable()
		enBnote.Disable()
		enBnote2.Disable()
		enAcCode.Disable()
		enAcName.Disable()
		enScode.Disable()
		enBdebit.Disable()
		enBcredit.Disable()
		btnSave.Disable()
		btnCheck.Disable()

		// ✅ ตรวจสอบว่า Item ปัจจุบันอยู่ในงวดที่อนุญาตให้แก้ไขได้ (dr1..dr2)
		canEdit := true
		if currentBitem != "" {
			t, err := time.Parse("02/01/06", enBdate.GetDate())
			if err == nil {
				if t.Before(dr1) || t.After(dr2) {
					canEdit = false
				}
			}
		}

		btnAdd.Enable() // Add ทำได้เสมอ (วันที่จะถูก default เป็น dr1)
		if canEdit {
			btnEdit.Enable()
			btnVoid.Enable() // ✅ Void ทำได้ใน VIEW MODE เท่านั้น
		} else {
			btnEdit.Disable()
			btnVoid.Disable()
		}
		btnDel.Disable() // ✅ Disable ใน VIEW MODE
		btnPrev.Enable()
		btnNext.Enable()
		clearLineInput()
	}

	setEditMode := func(isAdd bool) {
		if isAdd {
			actionFlag = "ADD"
			modeLabel.SetText("● ADD MODE")
		} else {
			actionFlag = "EDIT"
			modeLabel.SetText("● EDIT MODE")
		}
		enBdate.Enable()
		enBvoucher.Enable()
		enBref.Enable()
		enBchqno.Enable()
		enBchqdate.Enable()
		enBoff.Enable()
		enBcomtaxid.Enable()
		enBnote.Enable()
		enBnote2.Enable()
		enAcCode.Enable()
		enAcName.Enable()
		enScode.Enable()
		enBdebit.Enable()
		enBcredit.Enable()
		btnSave.Enable()
		btnCheck.Enable()
		btnAdd.Disable()
		btnEdit.Disable()
		btnDel.Enable()   // ✅ Enable ใน EDIT/ADD MODE
		btnVoid.Disable() // ✅ Void ทำไม่ได้ขณะ Edit
		btnPrev.Disable()
		btnNext.Disable()
	}

	// ─────────────────────────────────────────────────
	// saveAction
	//
	// ADD mode:
	//   - append bl ใหม่
	//   - ถ้า vatIdx != 0 → append vatLine พร้อมเซ็ต
	//     IsVATLine=true, ParentBline=bl.Bline
	//
	// EDIT mode (selectedLineID >= 0):
	//   - update lines[selectedLineID] in-place
	//   - scan หา VAT line ที่ ParentBline == bl.Bline เดิม
	//     → ถ้าพบ และ vatSelect != No-VAT → คำนวณ VAT ใหม่
	//     → ถ้า vatSelect == No-VAT → ลบ VAT line ออก
	//     → ถ้าไม่พบ และ vatSelect != No-VAT → append vatLine ใหม่
	// ─────────────────────────────────────────────────
	saveAction = func() {
		if actionFlag == "VIEW" {
			return
		}
		acCode := strings.TrimSpace(enAcCode.Text)
		if acCode == "" {
			return
		}

		// ถ้า user คลิก VAT line มา → ไม่อนุญาต edit โดยตรง (ยกเว้นบรรทัดปัดเศษ)
		if selectedLineID >= 0 && lines[selectedLineID].IsVATLine && lines[selectedLineID].AcCode != "524RND" && lines[selectedLineID].AcCode != "450RND" {
			dialog.ShowInformation("แจ้งเตือน",
				"บรรทัดนี้คือ VAT auto-gen\nให้แก้ไขที่ parent line แทน", w)
			clearLineInput()
			return
		}

		debit := strings.TrimSpace(enBdebit.Text)
		credit := strings.TrimSpace(enBcredit.Text)
		vatIdx := vatSelect.SelectedIndex()
		vatPct := parseFloat(vatPctText) / 100

		bl := BookLine{
			Comcode:   comCode,
			Bdate:     enBdate.GetDate(),
			Bvoucher:  enBvoucher.Text,
			Bitem:     currentBitem,
			AcCode:    acCode,
			AcName:    enAcName.Text,
			Scode:     enScode.Text,
			Sname:     enSname.Text,
			Bdebit:    debit,
			Bcredit:   credit,
			Bref:      enBref.Text,
			Boff:      enBoff.Text,
			Bcomtaxid: enBcomtaxid.Text,
			Bnote:     enBnote.Text,
			Bnote2:    enBnote2.Text,
			Bchqno:    enBchqno.Text,
			Bchqdate:  enBchqdate.GetDate(),
			IsVATLine: false,
		}

		if selectedLineID >= 0 {
			// ══════════════════════════════════
			// EDIT line in-place
			// ══════════════════════════════════
			oldBline := lines[selectedLineID].Bline
			bl.Bline = oldBline
			lines[selectedLineID] = bl

			// หา VAT line ที่ผูกอยู่กับ oldBline
			vatLineIdx := -1
			for i, l := range lines {
				if l.IsVATLine && l.ParentBline == oldBline {
					vatLineIdx = i
					break
				}
			}

			if vatIdx == 0 {
				if vatLineIdx >= 0 {
					lines = append(lines[:vatLineIdx], lines[vatLineIdx+1:]...)
				}
			} else {
				vatLine := BookLine{
					Comcode:     comCode,
					Bdate:       bl.Bdate,
					Bvoucher:    bl.Bvoucher,
					Bitem:       currentBitem,
					AcCode:      "235TVAT",
					AcName:      "ภาษีมูลค่าเพิ่มตั้งพัก",
					Scode:       bl.Scode,
					Sname:       bl.Sname,
					Bref:        bl.Bref,
					Boff:        bl.Boff,
					Bcomtaxid:   bl.Bcomtaxid,
					Bnote:       bl.Bnote,
					Bnote2:      bl.Bnote2,
					Bchqno:      bl.Bchqno,
					Bchqdate:    bl.Bchqdate,
					IsVATLine:   true,
					ParentBline: oldBline,
				}
				switch vatIdx {
				case 1:
					vatLine.Bdebit = floorVAT(parseFloat(debit) * vatPct)
					vatLine.Bcredit = "0.00"
				case 2:
					vatLine.Bdebit = "0.00"
					vatLine.Bcredit = floorVAT(parseFloat(credit) * vatPct)
				case 3:
					vatLine.Bdebit = floorVAT(parseFloat(debit) * vatPct)
					vatLine.Bcredit = floorVAT(parseFloat(credit) * vatPct)
				}
				if vatLineIdx >= 0 {
					vatLine.Bline = lines[vatLineIdx].Bline
					lines[vatLineIdx] = vatLine
				} else {
					vatLine.Bline = len(lines) + 1
					lines = append(lines, vatLine)
				}
			}

			// ── Auto WHT EDIT mode ──
			whtIdx := whtSelect.SelectedIndex()
			whtPct := parseFloat(whtPctText) / 100
			// หา WHT line เดิมที่ผูกกับ oldBline (AcCode=120WHT หรือ 235WHT)
			whtLineIdx := -1
			for i, l := range lines {
				if l.IsVATLine && l.ParentBline == oldBline &&
					(l.AcCode == "120WHT" || l.AcCode == "235WHT") {
					whtLineIdx = i
					break
				}
			}
			if whtIdx == 0 {
				if whtLineIdx >= 0 {
					lines = append(lines[:whtLineIdx], lines[whtLineIdx+1:]...)
				}
			} else {
				rawBase := parseFloat(debit)
				if rawBase == 0 {
					rawBase = parseFloat(credit)
				}
				// เช็คว่ารายการนี้มีการคิด VAT ด้วยหรือไม่
				actualVatPct := 0.0
				if vatIdx != 0 {
					actualVatPct = vatPct // ถ้ามี VAT ให้ใช้ vatPct (เช่น 0.07)
				}
				// ถอดหา Base จากยอดเงินสด (Net)
				// สูตรหา Base: Net = Base + VAT - WHT
				// ดังนั้น Base = Net / (1 + VAT - WHT)
				divisor := 1.0 + actualVatPct - whtPct

				if whtIdx == 1 && parseFloat(debit) > 0 {
					rawBase = rawBase / divisor
				} else if whtIdx == 2 && parseFloat(credit) > 0 {
					rawBase = rawBase / divisor
				}

				whtAmt := floorVAT(rawBase * whtPct)
				whtLine := BookLine{
					Comcode:     comCode,
					Bdate:       bl.Bdate,
					Bvoucher:    bl.Bvoucher,
					Bitem:       currentBitem,
					Scode:       bl.Scode,
					Sname:       bl.Sname,
					Bref:        bl.Bref,
					Boff:        bl.Boff,
					Bcomtaxid:   bl.Bcomtaxid,
					Bnote:       bl.Bnote,
					Bnote2:      bl.Bnote2,
					Bchqno:      bl.Bchqno,
					Bchqdate:    bl.Bchqdate,
					IsVATLine:   true,
					ParentBline: oldBline,
				}
				switch whtIdx {
				case 1:
					whtLine.AcCode = "120WHT"
					whtLine.AcName = "ภาษีหัก ณ ที่จ่ายรอรับคืน"
					whtLine.Bdebit = whtAmt
					whtLine.Bcredit = "0.00"
				case 2:
					whtLine.AcCode = "235WHT"
					whtLine.AcName = "ภาษีหัก ณ ที่จ่ายค้างจ่าย"
					whtLine.Bdebit = "0.00"
					whtLine.Bcredit = whtAmt
				}
				if whtLineIdx >= 0 {
					whtLine.Bline = lines[whtLineIdx].Bline
					lines[whtLineIdx] = whtLine
				} else {
					whtLine.Bline = len(lines) + 1
					lines = append(lines, whtLine)
				}
			}

		} else {
			// ══════════════════════════════════
			// ADD line ใหม่
			// ══════════════════════════════════
			bl.Bline = len(lines) + 1
			lines = append(lines, bl)

			if vatIdx != 0 {
				vatLine := BookLine{
					Comcode:     comCode,
					Bdate:       bl.Bdate,
					Bvoucher:    bl.Bvoucher,
					Bitem:       currentBitem,
					Bline:       len(lines) + 1,
					AcCode:      "235TVAT",
					AcName:      "ภาษีมูลค่าเพิ่มตั้งพัก",
					Scode:       bl.Scode,
					Sname:       bl.Sname,
					Bref:        bl.Bref,
					Boff:        bl.Boff,
					Bcomtaxid:   bl.Bcomtaxid,
					Bnote:       bl.Bnote,
					Bnote2:      bl.Bnote2,
					Bchqno:      bl.Bchqno,
					Bchqdate:    bl.Bchqdate,
					IsVATLine:   true,
					ParentBline: bl.Bline,
				}
				switch vatIdx {
				case 1:
					vatLine.Bdebit = floorVAT(parseFloat(debit) * vatPct)
					vatLine.Bcredit = "0.00"
				case 2:
					vatLine.Bdebit = "0.00"
					vatLine.Bcredit = floorVAT(parseFloat(credit) * vatPct)
				case 3:
					vatLine.Bdebit = floorVAT(parseFloat(debit) * vatPct)
					vatLine.Bcredit = floorVAT(parseFloat(credit) * vatPct)
				}
				lines = append(lines, vatLine)
			}

			// ── Auto WHT line ──
			// 120WHT: ฝั่งถูกหัก → Debit 120WHT (สินทรัพย์รอรับคืน)
			//         ยอด = baseAmt * whtPct (คำนวณจาก Debit ของ parent)
			// 235WHT: ฝั่งหักให้คนอื่น → Credit 235WHT (หนี้สินต้องนำส่ง)
			//         ยอด = baseAmt * whtPct (คำนวณจาก Debit ของ parent)
			// ── Auto WHT line ──
			whtIdx := whtSelect.SelectedIndex()
			whtPct := parseFloat(whtPctText) / 100

			if whtIdx != 0 {
				// 1. ดึงยอด Base จากบรรทัดปัจจุบัน (Expense หรือ Revenue)
				rawBase := parseFloat(debit)
				if rawBase == 0 {
					rawBase = parseFloat(credit)
				}

				// ❌ ลบโค้ดส่วน actualVatPct และ divisor ทิ้งทั้งหมด ❌
				// เพราะ rawBase คือฐานที่ถูกต้องอยู่แล้ว ไม่ต้องถอดฐานอีก

				// 2. คำนวณ WHT จากฐานตรงๆ (Base * %)
				whtAmt := floorVAT(rawBase * whtPct) // เช่น 100 * 0.03 = 3.00

				// 3. สร้างบรรทัด WHT
				whtLine := BookLine{
					Comcode:     comCode,
					Bdate:       bl.Bdate,
					Bvoucher:    bl.Bvoucher,
					Bitem:       currentBitem,
					Bline:       len(lines) + 1,
					Scode:       bl.Scode,
					Sname:       bl.Sname,
					Bref:        bl.Bref,
					Boff:        bl.Boff,
					Bcomtaxid:   bl.Bcomtaxid,
					Bnote:       bl.Bnote,
					Bnote2:      bl.Bnote2,
					Bchqno:      bl.Bchqno,
					Bchqdate:    bl.Bchqdate,
					IsVATLine:   true, // ใช้ flag เดียวกันเพื่อ lock ไม่ให้ edit โดยตรง
					ParentBline: bl.Bline,
				}

				// 4. กำหนดฝั่ง Dr / Cr ตามประเภท WHT
				switch whtIdx {
				case 1: // 120WHT — ถูกหัก ณ ที่จ่าย (เรามีรายได้ -> สินทรัพย์รอรับคืน -> Dr.)
					whtLine.AcCode = "120WHT"
					whtLine.AcName = "ภาษีหัก ณ ที่จ่ายรอรับคืน"
					whtLine.Bdebit = whtAmt
					whtLine.Bcredit = "0.00"
				case 2: // 235WHT — หัก ณ ที่จ่าย (เรามีค่าใช้จ่าย -> หนี้สินต้องนำส่ง -> Cr.)
					whtLine.AcCode = "235WHT"
					whtLine.AcName = "ภาษีหัก ณ ที่จ่ายค้างจ่าย"
					whtLine.Bdebit = "0.00"
					whtLine.Bcredit = whtAmt
				}

				lines = append(lines, whtLine)
			}
		}

		applyAutoRounding()
		calcSum()
		lineList.Refresh()
		clearLineInput()
		w.Canvas().Focus(enAcCode)
	}

	// ─────────────────────────────────────────────────
	// checkAndSave — commit ลง Excel พร้อม col ใหม่
	// col 17 = IsVATLine (1/0), col 18 = ParentBline
	// ─────────────────────────────────────────────────
	// checkRefDuplicate — ตรวจ Bref ซ้ำใน Bitem อื่น (global)
	checkRefDuplicate := func(ref string) (string, bool) {
		if strings.TrimSpace(ref) == "" {
			return "", false
		}
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return "", false
		}
		defer f.Close()
		rows, _ := f.GetRows("Book_items")
		for i, row := range rows {
			if i == 0 || len(row) < 12 || row[0] != comCode {
				continue
			}
			if safeGet(row, 11) == strings.TrimSpace(ref) && safeGet(row, 3) != currentBitem {
				return safeGet(row, 3), true
			}
		}
		return "", false
	}

	// checkVoucherDuplicate — ตรวจ Bvoucher ซ้ำใน Bitem อื่น
	// ADD mode: ตรวจทุก row | EDIT mode: ยกเว้น currentBitem ของตัวเอง
	checkVoucherDuplicate := func(voucher string) (string, bool) {
		voucher = strings.TrimSpace(voucher)
		if voucher == "" {
			return "", false
		}
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return "", false
		}
		defer f.Close()
		rows, _ := f.GetRows("Book_items")
		seen := map[string]bool{}
		for i, row := range rows {
			if i == 0 || len(row) < 4 || row[0] != comCode {
				continue
			}
			bitem := safeGet(row, 3)
			if bitem == currentBitem { // ยกเว้น voucher ของตัวเอง (EDIT mode)
				continue
			}
			if seen[bitem] {
				continue
			}
			seen[bitem] = true
			if strings.TrimSpace(safeGet(row, 2)) == voucher {
				return bitem, true
			}
		}
		return "", false
	}

	checkAndSave = func() {
		if actionFlag == "VIEW" {
			return
		}

		// sync header fields (Bref, Bcomtaxid, Bnote, Bnote2 ฯลฯ) จาก entry → lines[]
		for i := range lines {
			lines[i].Bdate = enBdate.GetDate()
			lines[i].Bvoucher = enBvoucher.Text
			lines[i].Bref = enBref.Text
			lines[i].Boff = enBoff.Text
			lines[i].Bcomtaxid = enBcomtaxid.Text
			lines[i].Bnote = enBnote.Text
			lines[i].Bnote2 = enBnote2.Text
			lines[i].Bchqno = enBchqno.Text
			lines[i].Bchqdate = enBchqdate.GetDate()
		}

		var sumD, sumC float64
		for _, l := range lines {
			sumD += parseFloat(l.Bdebit)
			sumC += parseFloat(l.Bcredit)
		}
		if strings.TrimSpace(enBvoucher.Text) == "" {
			showErrorDialog("แจ้งเตือน", "กรุณากรอก Voucher", w, func() { w.Canvas().Focus(enBvoucher) })
			return
		}
		if dupBitem, found := checkVoucherDuplicate(enBvoucher.Text); found {
			showErrorDialog("แจ้งเตือน", fmt.Sprintf("Voucher \"%s\" ซ้ำกับ Item %s\nไม่สามารถบันทึกได้", strings.TrimSpace(enBvoucher.Text), dupBitem), w, func() { w.Canvas().Focus(enBvoucher) })
			return
		}
		if len(lines) == 0 {
			showErrorDialog("แจ้งเตือน", "กรุณาเพิ่มรายการอย่างน้อย 1 รายการ", w, nil)
			return
		}

		if fmt.Sprintf("%.2f", sumD) != fmt.Sprintf("%.2f", sumC) {
			showErrorDialog("แจ้งเตือน", fmt.Sprintf("Debit (%s) ≠ Credit (%s)\nกรุณาตรวจสอบ", bsNum(sumD), bsNum(sumC)), w, nil)
			return
		}

		// ตรวจ Bref ซ้ำก่อน save (ครอบคลุมทั้ง ADD และ EDIT)
		if ref := strings.TrimSpace(enBref.Text); ref != "" {
			if dupBitem, found := checkRefDuplicate(ref); found {
				showErrorDialog("แจ้งเตือน", fmt.Sprintf("Reference \"%s\" ซ้ำกับ Item %s\nไม่สามารถบันทึกได้", ref, dupBitem), w, nil)
				w.Canvas().Focus(enBref)
				return
			}
		}

		{
			hook := NewJournalHook(nil)
			var entries []BookEntry
			for _, bl := range lines {
				entries = append(entries, BookEntry{
					AccountCode: bl.AcCode,
					AccountName: bl.AcName,
					Debit:       parseFloat(bl.Bdebit),
					Credit:      parseFloat(bl.Bcredit),
				})
			}
			summary, valErr := hook.BeforeSave(entries)
			if valErr != nil && IsBlocked(valErr) {
				showErrorDialog("ผิดหลักบัญชี", valErr.Error(), w, nil)
				return
			}
			if summary.HighLevel == LevelWarning && !warningConfirmed {
				// ตรวจว่ามี negative_amount pattern ไหม → แสดงปุ่ม "แก้ไขอัตโนมัติ" เฉพาะกรณีนี้
				hasNegative := false
				for _, r := range summary.Results {
					if r.PatternName == "negative_amount" {
						hasNegative = true
						break
					}
				}

				var warnDlg dialog.Dialog
				btnYes := newEnterButton("ยืนยันบันทึก", func() {
					warnDlg.Hide()
					warningConfirmed = true
					checkAndSave()
					warningConfirmed = false
				})
				btnYes.Importance = widget.WarningImportance

				// ปุ่มแก้ไขอัตโนมัติ — แสดงเฉพาะกรณีตัวเลขติดลบ (negative_amount)
				btnAutoFix := widget.NewButton("แก้ไขอัตโนมัติ (สลับ Dr/Cr)", func() {
					warnDlg.Hide()
					for i, bl := range lines {
						d := parseFloat(bl.Bdebit)
						c := parseFloat(bl.Bcredit)
						if d < 0 {
							lines[i].Bcredit = fmt.Sprintf("%.2f", -d)
							lines[i].Bdebit = "0.00"
						} else if c < 0 {
							lines[i].Bdebit = fmt.Sprintf("%.2f", -c)
							lines[i].Bcredit = "0.00"
						}
					}
					warningConfirmed = true
					checkAndSave()
					warningConfirmed = false
				})
				btnAutoFix.Importance = widget.HighImportance
				if !hasNegative {
					btnAutoFix.Hide() // ซ่อนถ้าไม่ใช่ negative case
				}

				btnNo := newEscButton("ยกเลิก (Esc)", func() {
					warnDlg.Hide()
					go func() { fyne.Do(func() { w.Canvas().Focus(enAcCode) }) }()
				})

				body := container.NewVBox()
				for _, r := range summary.Results {
					if r.Level != LevelWarning {
						continue
					}
					// แปลง message ให้เป็นภาษาไทยที่นักบัญชีเข้าใจง่าย
					displayMsg := thaiWarningMessage(r.PatternName, r.Intent, r.Message)
					intentLbl := widget.NewLabelWithStyle(
						r.Intent,
						fyne.TextAlignLeading, fyne.TextStyle{Bold: true},
					)
					msgLbl := widget.NewLabel(displayMsg)
					msgLbl.Wrapping = fyne.TextWrapWord
					body.Add(intentLbl)
					body.Add(msgLbl)
					body.Add(widget.NewSeparator())
				}
				warnDlg = dialog.NewCustomWithoutButtons(
					"⚠️  ตรวจพบรายการที่ควรตรวจสอบก่อนบันทึก",
					container.NewVBox(
						body,
						container.NewCenter(container.NewHBox(btnYes, btnAutoFix, btnNo)),
					), w)
				warnDlg.Show()
				go func() { fyne.Do(func() { w.Canvas().Focus(btnYes) }) }()
				return
			}
		}

		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			dialog.ShowError(fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err), w)
			return
		}
		// ✅ ไม่ใช้ defer — safePostVoucher เรียก f.Save() หลายรอบ
		//    ต้องปิด file เองหลัง POST เสร็จเท่านั้น

		// ✅ EDIT mode: UNPOST old first
		// ✅ EDIT mode: UNPOST old first
		if actionFlag == "EDIT" {
			if err := unpostVoucherFromLedger(currentBitem, comCode, currentBperiod, f, xlOptions); err != nil {
				f.Close()
				dialog.ShowError(fmt.Errorf("UNPOST failed: %v", err), w)
				return
			}
		}

		sheet := "Book_items"
		rows, _ := f.GetRows(sheet)

		// ลบ rows เดิมของ Bitem + Bperiod นี้ (ใช้ทั้งคู่เป็น key เพื่อไม่ทับ item อื่น period)
		for i := len(rows); i >= 1; i-- {
			row := rows[i-1]
			if len(row) >= 4 && row[0] == comCode && row[3] == currentBitem {
				var rowPeriod int
				if len(row) > 21 {
					fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
				}

				// ✅ BUG 2 FIX: ใช้ Date Range เป็น Fallback สำหรับข้อมูลเก่า
				rowDate := safeGet(row, 1)
				t, dateErr := time.Parse("02/01/06", rowDate)

				isInPeriod := (rowPeriod != 0 && rowPeriod == currentBperiod) ||
					(rowPeriod == 0 && dateErr == nil && !t.Before(dr1) && !t.After(dr2))

				if isInPeriod {
					f.RemoveRow(sheet, i)
				}
			}
		}

		// เพิ่ม header ถ้ายังไม่มี (รวม col ใหม่ IsVATLine, ParentBline, Posted, Bperiod)
		rows, _ = f.GetRows(sheet)
		if len(rows) == 0 {
			headers := []string{
				"Comcode", "Bdate", "Bvoucher", "Bitem", "Bline",
				"Ac_code", "Ac_name", "Scode", "Sname",
				"Bdebit", "Bcredit", "Bref", "Boff",
				"Bcomtaxid", "Bnote", "Bchqno", "Bchqdate",
				"IsVATLine", "ParentBline", "Posted", "Bnote2",
				"Bperiod",
			}
			for i, h := range headers {
				col, _ := excelize.ColumnNumberToName(i + 1)
				f.SetCellValue(sheet, fmt.Sprintf("%s1", col), h)
			}
			rows, _ = f.GetRows(sheet)
		}

		startRow := len(rows) + 1
		for i, bl := range lines {
			row := startRow + i
			isVAT := 0
			if bl.IsVATLine {
				isVAT = 1
			}
			vals := []interface{}{
				bl.Comcode, bl.Bdate, bl.Bvoucher, bl.Bitem, i + 1,
				bl.AcCode, bl.AcName, bl.Scode, bl.Sname,
				parseFloat(bl.Bdebit), parseFloat(bl.Bcredit),
				bl.Bref, bl.Boff, bl.Bcomtaxid, bl.Bnote, bl.Bchqno, bl.Bchqdate,
				isVAT,          // col R (17)
				bl.ParentBline, // col S (18)
				0,              // col T (19=Posted)
				bl.Bnote2,      // col U (20=Bnote2)
				currentBperiod, // col V (21=Bperiod) ← KEY สำหรับแยก period
			}
			for j, v := range vals {
				col, _ := excelize.ColumnNumberToName(j + 1)
				f.SetCellValue(sheet, fmt.Sprintf("%s%d", col, row), v)
			}
		}

		if err := f.Save(); err != nil {
			f.Close()
			dialog.ShowError(fmt.Errorf("บันทึกไม่สำเร็จ: %v", err), w)
			return
		}

		// auto-save Customer: TaxID → CustName (Note)
		// ต้อง Save อีกรอบหลัง upsert เพราะ upsertCustomer เขียนลง f แล้ว
		if tid := enBcomtaxid.Text; len(tid) == 13 && enBnote.Text != "" {
			upsertCustomer(f, comCode, tid, enBnote.Text)
			_ = f.Save() // flush Customer_Log ลง Excel
		}
		// // 🌟 เติมโค้ดตรงนี้ครับ: ทำการ POST ลง Ledger ทันทีที่ Save เสร็จ (ตอนนี้ใช้  RecalculateLedgerMaster)
		// if err := safePostVoucher(currentBitem, comCode, currentBperiod, f, xlOptions); err != nil {
		// 	fmt.Printf("Auto-Post failed: %v\n", err)
		// 	// ไม่ต้อง return error ให้ user เห็นหน้าจอค้าง แค่ log ไว้
		// 	// เพราะข้อมูลลงสมุดรายวันสำเร็จแล้ว ค่อยให้ user กดปุ่ม Repair ทีหลังได้
		// }

		// f.Close() // close หลัง save และ post สำเร็จ

		// ──────────────────────────────────────────────────────────
		// ✅ NEW: อัปเดตยอด Thisper01-12 แบบ Real-time ทันทีที่บันทึกเสร็จ
		// ──────────────────────────────────────────────────────────
		go func() {
			// เรียกใช้ฟังก์ชันคำนวณยอดที่เพิ่งสร้างใน ledger_ui.go
			err := RecalculateLedgerMaster(xlOptions)
			if err != nil {
				fmt.Printf("Auto-recalculate failed: %v\n", err)
			}

			// ✅ Refresh Ledger display (callback from ledger_ui)
			fyne.Do(func() {
				if refreshLedgerFunc != nil {
					refreshLedgerFunc()
				}
			})
		}()

		var d dialog.Dialog
		savedBitem := currentBitem
		savedBperiod := currentBperiod // ✅ จำ period ไว้ด้วย
		okBtn := newEnterButton("OK", func() {
			d.Hide()
			setViewMode()
			loadBitem(savedBitem, savedBperiod) // ✅ เติม , savedBperiod
			w.Canvas().Focus(nil)
		})

		content := container.NewVBox(
			widget.NewLabel("บันทึกเรียบร้อยแล้ว"),
			container.NewCenter(okBtn),
		)
		d = dialog.NewCustomWithoutButtons("สำเร็จ", content, w)
		d.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(okBtn) }) }()
	}
	// ─────────────────────────────────────────────────
	// lineList
	// ─────────────────────────────────────────────────
	lineList = newDeletableList(
		func() int { return len(lines) },
		func() fyne.CanvasObject {
			return container.NewHBox(
				container.NewGridWrap(fyne.NewSize(100, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(350, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(60, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(150, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(100, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(100, 28), widget.NewLabel("")),
			)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			bl := lines[id]
			row := o.(*fyne.Container)
			acCodeLabel := row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label)
			acCodeLabel.SetText(bl.AcCode)
			// VAT line แสดงตัวเอียงให้รู้ว่าเป็น auto-gen
			if bl.IsVATLine {
				acCodeLabel.TextStyle = fyne.TextStyle{Italic: true}
			} else {
				acCodeLabel.TextStyle = fyne.TextStyle{}
			}
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(bl.AcName)
			row.Objects[2].(*fyne.Container).Objects[0].(*widget.Label).SetText(bl.Scode)
			row.Objects[3].(*fyne.Container).Objects[0].(*widget.Label).SetText(bl.Sname)
			row.Objects[4].(*fyne.Container).Objects[0].(*widget.Label).SetText(bsNum(parseFloat(bl.Bdebit)))
			row.Objects[5].(*fyne.Container).Objects[0].(*widget.Label).SetText(bsNum(parseFloat(bl.Bcredit)))
		},
		enterEditMode,
		execDeleteLine,
		cancelEdit,
	)

	// ─────────────────────────────────────────────────
	// OnSelected — ถ้าเลือก VAT line:
	//   pull ค่ามาใส่ form เหมือนปกติ แต่ set vatSelect
	//   ตาม parent line เพื่อให้ user เห็น context
	//   (saveAction จะ block การ edit VAT line โดยตรงอยู่แล้ว)
	// ─────────────────────────────────────────────────
	lineList.OnSelected = func(id widget.ListItemID) {
		if actionFlag == "VIEW" {
			lineList.UnselectAll()
			return
		}
		bl := lines[id]
		selectedLineID = id

		enAcCode.SetText(bl.AcCode)
		enAcName.SetText(bl.AcName)
		enScode.SetText(bl.Scode)
		enSname.SetText(bl.Sname)
		enBdebit.SetText(bl.Bdebit)
		enBcredit.SetText(bl.Bcredit)

		if bl.IsVATLine {
			// VAT line — หา vatIdx จาก parent line
			for _, parent := range lines {
				if !parent.IsVATLine && parent.Bline == bl.ParentBline {
					// ดูว่า parent มี debit หรือ credit
					d := parseFloat(parent.Bdebit)
					c := parseFloat(parent.Bcredit)
					switch {
					case d > 0 && c == 0:
						vatSelect.SetSelectedIndex(1) // Debit VAT
					case d == 0 && c > 0:
						vatSelect.SetSelectedIndex(2) // Credit VAT
					default:
						vatSelect.SetSelectedIndex(3) // Both VAT
					}
					break
				}
			}
			// restore WHT select จาก AcCode ของ WHT line ที่ผูกกับ parent
			whtSelect.SetSelectedIndex(0)
			for _, l := range lines {
				if l.IsVATLine && l.ParentBline == bl.Bline {
					switch l.AcCode {
					case "120WHT":
						whtSelect.SetSelectedIndex(1)
					case "235WHT":
						whtSelect.SetSelectedIndex(2)
					}
				}
			}
		} else {
			vatSelect.SetSelectedIndex(0)
			whtSelect.SetSelectedIndex(0)
		}

		w.Canvas().Focus(enAcCode)
	}

	// --- Buttons ---
	btnSave = widget.NewButtonWithIcon("", theme.DocumentSaveIcon(), func() { checkAndSave() })

	btnAdd = widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() {
		if actionFlag == "ADD" || actionFlag == "EDIT" {
			return
		}
		prevBitem = currentBitem

		targetPeriod := periodCfg.NowPeriod
		if workingPeriod != 0 {
			targetPeriod = workingPeriod
		}

		newBitem := autoNextBitem(targetPeriod) // ✅ เติม targetPeriod
		currentBitem = newBitem
		currentBperiod = targetPeriod

		enBitem.SetText(newBitem)
		clearHeader()
		// ✅ Default วันที่เป็นวันแรกของงวดที่ทำงานอยู่ ป้องกันการหลงงวด
		enBdate.SetDate(dr1.Format("02/01/06"))
		clearLineInput()
		lines = nil
		lineList.Refresh()
		setEditMode(true)
		w.Canvas().Focus(enBdate)
	})

	btnEdit = widget.NewButtonWithIcon("", theme.DocumentCreateIcon(), func() {
		if currentBitem == "" {
			return
		}
		setEditMode(false)
		clearLineInput()
		w.Canvas().Focus(enBdate)
	})

	btnDel = widget.NewButtonWithIcon("", theme.DeleteIcon(), func() {
		execDeleteLine() // ← Simple! เรียก execDeleteLine เท่านั้น
	})
	btnDel.Disable()

	btnCheck = widget.NewButtonWithIcon("Check", theme.ConfirmIcon(), func() { checkAndSave() })

	// ✅ REPAIR Button
	btnRepair = widget.NewButtonWithIcon("🔧 Repair", theme.WarningIcon(), func() {
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			dialog.ShowError(fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err), w)
			return
		}
		// ✅ ไม่ใช้ defer — ปิด file เองหลัง recoverPosting เสร็จ
		//    เพราะ recoverPosting เรียก f.Save() หลายรอบ ถ้า defer close
		//    ก่อน save เสร็จจะ corrupt

		if err := recoverPosting(f, xlOptions); err != nil {
			f.Close()
			dialog.ShowError(fmt.Errorf("Recovery failed: %v", err), w)
			return
		}
		f.Close()

		dialog.ShowInformation("สำเร็จ", "ซ่อมแซม Posting เรียบร้อยแล้ว", w)

		// ❌ โค้ดเดิม:
		// allAfterRepair := loadAllBitemsInRange(xlOptions, comCode, dr1, dr2)
		// if len(allAfterRepair) > 0 {
		// 	loadBitem(allAfterRepair[0])
		// }

		// ✅ เปลี่ยนเป็น:
		targetP := periodCfg.NowPeriod
		if workingPeriod != 0 {
			targetP = workingPeriod
		}
		allAfterRepair := loadAllBitemsInPeriod(xlOptions, comCode, targetP)
		if len(allAfterRepair) > 0 {
			loadBitem(allAfterRepair[0], targetP)
		}
		setViewMode()
	})
	btnRepair.Disable() // Enable only if recovery needed

	// ── btnVoid — ยกเลิกรายการ (Void) คลิกเดียว ──
	btnVoid = widget.NewButtonWithIcon("Void", theme.CancelIcon(), func() {
		if currentBitem == "" {
			return
		}
		// ตรวจว่า void ไปแล้วหรือยัง
		for _, l := range lines {
			if strings.HasPrefix(l.Bvoucher, "VOID-") {
				showErrorDialog("แจ้งเตือน", fmt.Sprintf("Item %s ถูก Void ไปแล้ว", currentBitem), w, nil)
				return
			}
		}
		// ── อ่าน Void Password จาก Company_Profile col J ──
		voidPwd := ""
		{
			pf, err := excelize.OpenFile(currentDBPath, xlOptions)
			if err == nil {
				voidPwd, _ = pf.GetCellValue("Company_Profile", "J2")
				pf.Close()
			}
		}
		if strings.TrimSpace(voidPwd) == "" {
			showErrorDialog("ไม่สามารถ Void ได้",
				"ยังไม่ได้ตั้ง Void Password\nกรุณาไปตั้งค่าใน Setup → Company ก่อน", w, nil)
			return
		}
		// ── Password dialog ──
		snapshotBitem := currentBitem
		var pwdDlg dialog.Dialog
		enPwd := widget.NewPasswordEntry()
		enPwd.SetPlaceHolder("กรอก Void Password")
		pwdErrLbl := widget.NewLabel("")

		execVoid := func() {
			if enPwd.Text != voidPwd {
				pwdErrLbl.SetText("❌ Password ไม่ถูกต้อง")
				enPwd.SetText("")
				go func() { fyne.Do(func() { w.Canvas().Focus(enPwd) }) }()
				return
			}
			pwdDlg.Hide()
			// ── ดำเนินการ Void ──
			f, err := excelize.OpenFile(currentDBPath, xlOptions)
			if err != nil {
				dialog.ShowError(fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err), w)
				return
			}
			// 1. Unpost ถ้า posted อยู่
			if err := unpostVoucherFromLedger(snapshotBitem, comCode, currentBperiod, f, xlOptions); err != nil {
				f.Close()
				dialog.ShowError(fmt.Errorf("Unpost failed: %v", err), w)
				return
			}
			// 2. ลบ rows เดิมทั้งหมด

			sheet := "Book_items"
			rows, _ := f.GetRows(sheet)
			for i := len(rows); i >= 1; i-- {
				row := rows[i-1]
				if len(row) >= 4 && row[0] == comCode && row[3] == snapshotBitem {
					var rowPeriod int
					if len(row) > 21 {
						fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
					}

					// ✅ BUG 2 FIX: ใช้ Date Range เป็น Fallback สำหรับข้อมูลเก่า
					rowDate := safeGet(row, 1)
					t, dateErr := time.Parse("02/01/06", rowDate)

					isInPeriod := (rowPeriod != 0 && rowPeriod == currentBperiod) ||
						(rowPeriod == 0 && dateErr == nil && !t.Before(dr1) && !t.After(dr2))

					if isInPeriod {
						f.RemoveRow(sheet, i)
					}
				}
			}

			// 3. เขียน 1 line Void แทน
			rows, _ = f.GetRows(sheet)
			voidVoucher := "VOID-" + snapshotBitem
			voidDate := ""
			if len(lines) > 0 {
				voidDate = lines[0].Bdate
			}
			if voidDate == "" {
				voidDate = dr1.Format("02/01/06")
			}
			newRow := len(rows) + 1
			vals := []interface{}{
				comCode, voidDate, voidVoucher, snapshotBitem, 1,
				"", "ยกเลิกรายการ", "", "",
				0.0, 0.0,
				"ยกเลิกรายการ", "", "", "ยกเลิกรายการ", "", "",
				0, 0, 0, "",
				currentBperiod, // col V (21=Bperiod)
			}
			for j, v := range vals {
				col, _ := excelize.ColumnNumberToName(j + 1)
				f.SetCellValue(sheet, fmt.Sprintf("%s%d", col, newRow), v)
			}

			// ✅ BUG 1 FIX: จัดการ f.Close() ให้ดู Clean ขึ้น
			errSave := f.Save()
			f.Close() // ปิดไฟล์ทันทีหลัง Save เสร็จ ไม่ว่าจะ Error หรือไม่

			if errSave != nil {
				dialog.ShowError(fmt.Errorf("บันทึกไม่สำเร็จ: %v", errSave), w)
				return
			}

			loadBitem(snapshotBitem, currentBperiod)
			setViewMode()
			var okD dialog.Dialog
			okBtn := newEnterButton("OK (Enter)", func() { okD.Hide() })
			okD = dialog.NewCustomWithoutButtons("✓ Void สำเร็จ", container.NewVBox(
				widget.NewLabel(fmt.Sprintf("Item %s ถูก Void เรียบร้อย\nVoucher: %s", snapshotBitem, voidVoucher)),
				container.NewCenter(okBtn),
			), w)
			okD.Show()
			go func() { fyne.Do(func() { w.Canvas().Focus(okBtn) }) }()
		}

		enPwd.OnSubmitted = func(_ string) { execVoid() }
		btnPwdOK := newEnterButton("ยืนยัน (Enter)", func() { execVoid() })
		btnPwdOK.Importance = widget.DangerImportance
		btnPwdCancel := newEscButton("ยกเลิก (Esc)", func() {
			pwdDlg.Hide()
			go func() { fyne.Do(func() { w.Canvas().Focus(nil) }) }()
		})
		pwdDlg = dialog.NewCustomWithoutButtons(
			fmt.Sprintf("🔒 Void Item %s — ยืนยัน Password", snapshotBitem),
			container.NewVBox(
				widget.NewLabel("กรอก Void Password เพื่อดำเนินการ"),
				enPwd,
				pwdErrLbl,
				container.NewHBox(layout.NewSpacer(), btnPwdOK, btnPwdCancel, layout.NewSpacer()),
			), w)
		pwdDlg.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(enPwd) }) }()
	})
	btnVoid.Importance = widget.DangerImportance
	btnVoid.Disable() // เริ่มต้น disable จนกว่าจะโหลด item

	// Check if repair needed on init
	checkRepairNeeded := func() {
		f, _ := excelize.OpenFile(currentDBPath, xlOptions)
		if f != nil {
			rows, _ := f.GetRows("Posting_Log")
			for i, row := range rows {
				if i > 0 && len(row) >= 4 {
					if safeGet(row, 3) == "IN_PROGRESS" {
						btnRepair.Enable()
						break
					}
				}
			}
			f.Close()
		}
	}
	checkRepairNeeded()

	// ── +/- nav actions ──────────────────────────────────────────────
	nextBitem := func() {
		if actionFlag != "VIEW" {
			return
		}
		// ✅ โหลดเฉพาะ Item ใน Period ปัจจุบัน
		all := loadAllBitemsInPeriod(xlOptions, comCode, currentBperiod)
		for i, b := range all {
			if b == currentBitem && i < len(all)-1 {
				loadBitem(all[i+1], currentBperiod)
				return
			}
		}
	}
	prevBitemNav := func() {
		if actionFlag != "VIEW" {
			return
		}
		// ✅ โหลดเฉพาะ Item ใน Period ปัจจุบัน
		all := loadAllBitemsInPeriod(xlOptions, comCode, currentBperiod)
		for i, b := range all {
			if b == currentBitem && i > 0 {
				loadBitem(all[i-1], currentBperiod)
				return
			}
		}
	}

	bookNextFunc = nextBitem
	bookPrevFunc = prevBitemNav

	// ── setWorkingPeriodFunc — callback จาก Select Period dialog ──────────
	setWorkingPeriodFunc = func(period int) {
		periods := calcAllPeriods(periodCfg.YearEnd, periodCfg.TotalPeriods)
		if period < 1 || period > len(periods) {
			return
		}
		p := periods[period-1]
		dr1 = p.PStart
		dr2 = p.PEnd
		fyne.Do(func() {
			if workingPeriod == 0 {
				periodLabel.SetText(getPeriodSummary(periodCfg))
			} else {
				periodLabel.SetText(fmt.Sprintf("⚡ งวด %d/%d  (%s - %s)",
					period, periodCfg.TotalPeriods,
					dr1.Format("02/01/06"), dr2.Format("02/01/06")))
			}
			all := loadAllBitemsInPeriod(xlOptions, comCode, period) // ✅ ใช้ loadAllBitemsInPeriod
			if len(all) > 0 {
				loadBitem(all[0], period) // ✅ เติม , period
			} else {
				currentBitem = ""
				enBitem.SetText("")
				clearHeader()
				lines = nil
				calcSum()
				lineList.Refresh()
			}
			setViewMode()
		})
	}

	bookAddFunc = func() {
		if actionFlag == "ADD" || actionFlag == "EDIT" {
			return
		}
		prevBitem = currentBitem

		targetPeriod := periodCfg.NowPeriod
		if workingPeriod != 0 {
			targetPeriod = workingPeriod
		}

		newBitem := autoNextBitem(targetPeriod) // ✅ ส่ง targetPeriod
		currentBitem = newBitem
		currentBperiod = targetPeriod

		enBitem.SetText(newBitem)
		clearHeader()
		enBdate.SetDate(dr1.Format("02/01/06"))
		clearLineInput()
		lines = nil
		lineList.Refresh()
		setEditMode(true)
		w.Canvas().Focus(enBdate)
	}
	// // FocusGained → popup search
	// enBcomtaxid.onFocusGained = func() {
	// 	if actionFlag != "EDIT" && actionFlag != "ADD" {
	// 		return
	// 	}
	// 	showTaxIDSearch(w, xlOptions, func(taxID, custName string) {
	// 		enBcomtaxid.SetText(taxID)
	// 		if custName != "" && enBnote.Text == "" {
	// 			enBnote.SetText(custName)
	// 		}
	// 		w.Canvas().Focus(enBnote)
	// 	}, func() {
	// 		w.Canvas().Focus(enBcomtaxid)
	// 	})
	// }

	// wire dateEntry PageDown/PageUp → nav (เพราะ dateEntry ไม่ใช่ smartEntry)
	enBdate.onPageDown = func() { nextBitem() }
	enBdate.onPageUp = func() { prevBitemNav() }
	enBchqdate.onPageDown = func() { nextBitem() }
	enBchqdate.onPageUp = func() { prevBitemNav() }
	btnPrev = widget.NewButtonWithIcon("", theme.NavigateBackIcon(), func() { prevBitemNav() })
	btnNext = widget.NewButtonWithIcon("", theme.NavigateNextIcon(), func() { nextBitem() })

	// ── F3 search action ──────────────────────────────────────────────
	openBookSearch := func() {
		if actionFlag != "VIEW" {
			return
		}
		showBookSearch(w, xlOptions, comCode, func(bitem string, bperiod int) {
			loadBitem(bitem, bperiod) // ✅ รับและส่ง bperiod
			setViewMode()
		})
	}

	bookSearchFunc = openBookSearch
	// register bookGotoFunc → inventory Auto-JV ใช้ jump ตรงไป Bitem
	bookGotoFunc = func(bitem string, bperiod int) {
		loadBitem(bitem, bperiod)
		setViewMode()
	}

	// register Ctrl+S → checkAndSave, Esc → cancelEdit (Window level)
	registerCtrlS(w, func() {
		// ✅ ปิด PopUp ก่อน (ถ้าเปิดอยู่)
		if closeAcSearchPopup != nil {
			closeAcSearchPopup()
		}
		// ✅ ค่อย Save
		if actionFlag == "ADD" || actionFlag == "EDIT" {
			checkAndSave()
		}
	}, func() {
		cancelEdit()
	})

	allEntries := []*smartEntry{
		enBvoucher, enBref, enBchqno,
		enBoff, enBcomtaxid, enBnote, enBnote2,
		enAcCode, enAcName, enScode, enBdebit, enBcredit,
	}
	// for _, en := range allEntries {
	// 	en.onEsc = cancelEdit
	// 	en.onF2 = enterEditMode
	// 	en.onF4 = execDeleteLine
	// 	en.onF3 = openBookSearch
	// 	en.onPageDown = nextBitem
	// 	en.onPageUp = prevBitemNav
	// }
	for _, en := range allEntries {
		en.onEsc = cancelEdit
		en.onF2 = enterEditMode
		en.onF4 = execDeleteLine
		en.onF3 = openBookSearch
		en.onPageDown = nextBitem
		en.onPageUp = prevBitemNav
	}

	// ✅ เพิ่มโค้ดนี้ต่อท้ายลูป เพื่อให้ช่อง Tax ID กด F3 แล้วเปิดหน้าค้นหาลูกค้า
	enBcomtaxid.onF3 = func() {
		if actionFlag != "EDIT" && actionFlag != "ADD" {
			return
		}
		showTaxIDSearch(w, xlOptions, func(taxID, custName string) {
			enBcomtaxid.SetText(taxID)
			if custName != "" && enBnote.Text == "" {
				enBnote.SetText(custName)
			}
			w.Canvas().Focus(enBnote)
		}, func() {
			w.Canvas().Focus(enBcomtaxid)
		})
	}
	// dateEntry fields — ผูก callbacks แยก
	for _, de := range []*dateEntry{enBdate, enBchqdate} {
		de.onEsc = cancelEdit
		de.onEnter = func() { w.Canvas().Focus(enBvoucher) }
		de.onF3 = openBookSearch
		de.onSave = func() {
			if actionFlag == "ADD" || actionFlag == "EDIT" {
				checkAndSave()
			}
		}
	}

	// --- OnSubmitted ---
	enAcCode.onFocusGained = func() {
		if actionFlag != "EDIT" && actionFlag != "ADD" {
			return
		}
		go func() {
			fyne.Do(func() {
				showBookLedgerSearch(w, xlOptions, func(acCode, acName string) {
					enAcCode.SetText(acCode)
					enAcName.SetText(acName)
					w.Canvas().Focus(enScode)
				})
			})
		}()
	}
	enAcCode.onEnter = func() {
		s := strings.TrimSpace(enAcCode.Text)
		openAcSearch := func() {
			go func() {
				fyne.Do(func() {
					showBookLedgerSearch(w, xlOptions, func(acCode, acName string) {
						enAcCode.SetText(acCode)
						enAcName.SetText(acName)
						w.Canvas().Focus(enScode)
					})
				})
			}()
		}
		if s == "" {
			openAcSearch()
			return
		}
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()
		found := false
		rows, _ := f.GetRows("Ledger_Master")
		for _, row := range rows {
			if len(row) >= 3 && row[0] == comCode && row[1] == s {
				enAcName.SetText(row[2])
				w.Canvas().Focus(enScode)
				found = true
				break
			}
		}
		if !found {
			openAcSearch()
		}
	}
	enScode.onFocusGained = func() {
		if actionFlag != "EDIT" && actionFlag != "ADD" {
			return
		}
		go func() {
			fyne.Do(func() {
				showBookSubbookSearch(w, xlOptions, comCode, func(scode, sname string) {
					enScode.SetText(scode)
					enSname.SetText(sname)
					w.Canvas().Focus(vatSelect)
				})
			})
		}()
	}
	enScode.onEnter = func() {
		s := strings.TrimSpace(enScode.Text)
		openScSearch := func() {
			go func() {
				fyne.Do(func() {
					showBookSubbookSearch(w, xlOptions, comCode, func(scode, sname string) {
						enScode.SetText(scode)
						enSname.SetText(sname)
						w.Canvas().Focus(vatSelect)
					})
				})
			}()
		}
		if s == "" {
			openScSearch()
			return
		}
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()
		found := false
		rows, _ := f.GetRows("Subsidiary_Books")
		for _, row := range rows {
			if len(row) >= 3 && row[1] == s {
				enSname.SetText(row[2])
				w.Canvas().Focus(vatSelect)
				found = true
				break
			}
		}
		if !found {
			openScSearch()
		}
	}
	// enBdate — validate แล้ว focus ต่อ (เรียกจาก onEnter ที่ set ไว้แล้ว และพิมพ์ครบ 6 หลักอัตโนมัติ)
	enBdate.onEnter = func() {
		s := enBdate.GetDate()
		if s == "" {
			w.Canvas().Focus(enBvoucher)
			return
		}
		_, err := time.Parse("02/01/06", s)
		if err != nil {
			showErrorDialog("แจ้งเตือน", "รูปแบบวันที่ต้องเป็น dd/mm/yy", w, nil)
			enBdate.SetDate(time.Now().Format("02/01/06"))
			return
		}
		if periodErr != nil {
			dialog.ShowError(fmt.Errorf("ไม่สามารถตรวจสอบ Period ได้\n%v", periodErr), w)
			return
		}
		if verr := validateVoucherDate(s, dr1, dr2); verr != nil {
			dialog.ShowError(verr, w)
			enBdate.SetDate(dr1.Format("02/01/06"))
			return
		}
		w.Canvas().Focus(enBvoucher)
	}

	// ── Inline Autocomplete Dropdown ──────────────────────────────────────
	// digit-only filter + taxLabel + auto-fill Note จาก Customer_Log
	enBcomtaxid.OnChanged = func(s string) {
		digits := ""
		for _, r := range s {
			if r >= '0' && r <= '9' {
				digits += string(r)
			}
		}
		if len(digits) > 13 {
			digits = digits[:13]
		}
		if digits != s {
			enBcomtaxid.SetText(digits)
			return
		}
		if len(digits) == 13 {
			taxLabel.SetText("✅")
			if enBnote.Text == "" {
				if name := lookupCustomerName(xlOptions, digits); name != "" {
					enBnote.SetText(name)
				}
			}
		} else {
			taxLabel.SetText("")
		}
	}

	enBvoucher.OnSubmitted = func(s string) {
		s = strings.TrimSpace(s)
		if s != "" {
			if dupBitem, found := checkVoucherDuplicate(s); found {
				showErrorDialog("แจ้งเตือน", fmt.Sprintf("Voucher \"%s\" ซ้ำกับ Item %s", s, dupBitem), w, nil)
				enBvoucher.SetText("")
				w.Canvas().Focus(enBvoucher)
				return
			}
		}
		w.Canvas().Focus(enBref)
	}
	enBref.OnSubmitted = func(s string) {
		s = strings.TrimSpace(s)
		if s != "" {
			if dupBitem, found := checkRefDuplicate(s); found {
				showErrorDialog("แจ้งเตือน", fmt.Sprintf("Reference \"%s\" ซ้ำกับ Item %s", s, dupBitem), w, nil)
				enBref.SetText("")
				w.Canvas().Focus(enBref)
				return
			}
		}
		w.Canvas().Focus(enBcomtaxid)
	}
	enBcomtaxid.OnSubmitted = func(s string) {
		if s != "" && len(s) != 13 {
			showErrorDialog("แจ้งเตือน", "Tax ID ต้องมี 13 หลัก", w, nil)
			enBcomtaxid.SetText("")
			return
		}
		w.Canvas().Focus(enBoff)
	}
	enBoff.OnSubmitted = func(s string) { w.Canvas().Focus(enBnote) }
	enBnote.OnSubmitted = func(s string) { w.Canvas().Focus(enBnote2) }
	enBnote2.OnSubmitted = func(s string) { w.Canvas().Focus(enBchqno) }
	enBnote.OnSubmitted = func(s string) { w.Canvas().Focus(enBchqno) }
	enBchqno.OnSubmitted = func(s string) { w.Canvas().Focus(enBchqdate) }
	enBchqdate.onEnter = func() {
		s := enBchqdate.GetDate()
		if s != "" {
			_, err := time.Parse("02/01/06", s)
			if err != nil {
				showErrorDialog("แจ้งเตือน", "รูปแบบวันที่ต้องเป็น dd/mm/yy", w, nil)
				enBchqdate.SetDate("")
				return
			}
		}
		w.Canvas().Focus(enAcCode)
	}
	enBdebit.OnSubmitted = func(s string) { w.Canvas().Focus(enBcredit) }

	// --- Bdebit/Bcredit: live number format ---
	bookNumUpdating := false
	applyBookNumFormat := func(en *smartEntry) {
		if bookNumUpdating {
			return
		}
		s := en.Text
		// กรอง: รับเฉพาะตัวเลข 0-9 และจุดทศนิยม 1 จุด
		raw := ""
		hasDot := false
		for _, r := range s {
			if r >= '0' && r <= '9' {
				raw += string(r)
			} else if r == '.' && !hasDot {
				raw += string(r)
				hasDot = true
			}
		}
		// แยก integer / decimal
		intPart, decPart := raw, ""
		if idx := strings.Index(raw, "."); idx >= 0 {
			intPart = raw[:idx]
			decPart = raw[idx:]
		}
		// format comma เฉพาะ integer part
		formatted := intPart
		if v, err := strconv.ParseInt(intPart, 10, 64); err == nil && intPart != "" {
			formatted = formatComma(float64(v))
			if idx := strings.Index(formatted, "."); idx >= 0 {
				formatted = formatted[:idx]
			}
		}
		formatted += decPart
		if formatted == s {
			return
		}
		// ปรับ cursor
		oldCursor := en.CursorColumn
		commasBefore := 0
		for i, r := range formatted {
			if i >= oldCursor {
				break
			}
			if r == ',' {
				commasBefore++
			}
		}
		newCursor := oldCursor + commasBefore
		bookNumUpdating = true
		en.SetText(formatted)
		if newCursor > len([]rune(formatted)) {
			newCursor = len([]rune(formatted))
		}
		en.CursorColumn = newCursor
		en.Refresh()
		bookNumUpdating = false
	}
	enBdebit.OnChanged = func(s string) { applyBookNumFormat(enBdebit) }
	enBcredit.OnChanged = func(s string) { applyBookNumFormat(enBcredit) }
	// FocusLost: format เป็น 1,234,567.89 สมบูรณ์
	enBdebit.onFocusLost = func() {
		raw := strings.ReplaceAll(enBdebit.Text, ",", "")
		if v, err := strconv.ParseFloat(raw, 64); err == nil {
			bookNumUpdating = true
			enBdebit.SetText(formatComma(v))
			bookNumUpdating = false
		}
	}
	enBcredit.onFocusLost = func() {
		raw := strings.ReplaceAll(enBcredit.Text, ",", "")
		if v, err := strconv.ParseFloat(raw, 64); err == nil {
			bookNumUpdating = true
			enBcredit.SetText(formatComma(v))
			bookNumUpdating = false
		}
	}
	// enBcredit.OnSubmitted = func(s string) { saveAction() }
	enBcredit.OnSubmitted = func(s string) {
		// ส่งตัวแปรที่มีอยู่ใน Scope นี้เข้าไปใน validateForm
		if err := validateForm(enBdate, enBvoucher, enBcomtaxid, dr1, dr2); err != nil {
			var focusFn func()
			if strings.TrimSpace(enBvoucher.Text) == "" {
				focusFn = func() { w.Canvas().Focus(enBvoucher) }
			} else {
				focusFn = func() { w.Canvas().Focus(enBdate) }
			}
			showErrorDialog("แจ้งเตือน", err.Error(), w, focusFn)
			return
		}

		saveAction()
	}

	// --- Layout ---
	comName := ""
	ff, errr := excelize.OpenFile(currentDBPath, xlOptions)
	if errr == nil {
		comName, _ = ff.GetCellValue("Company_Profile", "B2")
		ff.Close()
	}

	toolbar := container.NewHBox(
		btnSave, btnAdd, btnEdit, btnDel, btnVoid,
		widget.NewButtonWithIcon("", theme.SearchIcon(), func() {
			showBookSearch(w, xlOptions, comCode, func(bitem string, bperiod int) { // ✅ รับ bperiod เพิ่ม
				loadBitem(bitem, bperiod) // ✅ ส่ง bperiod
				setViewMode()
			})
		}),

		btnPrev, btnNext, btnCheck,
		btnRepair,
		layout.NewSpacer(),
		periodLabel,
		widget.NewLabel("  "),
		modeLabel,
		widget.NewLabelWithStyle(comName, fyne.TextAlignTrailing, fyne.TextStyle{Bold: true}),
	)

	row1col1 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("Item")),
		container.NewGridWrap(fyne.NewSize(50, 30), enBitem),
		container.NewGridWrap(fyne.NewSize(50, 30), widget.NewLabel("Date")),
		container.NewGridWrap(fyne.NewSize(145, 30), enBdate),
	)
	row1col2 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(60, 30), widget.NewLabel("Voucher")),
		container.NewGridWrap(fyne.NewSize(150, 30), enBvoucher),
	)
	row1col3 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(50, 30), widget.NewLabel("Cheque")),
		container.NewGridWrap(fyne.NewSize(150, 30), enBchqno),
	)
	row2col1 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("Reference")),
		container.NewGridWrap(fyne.NewSize(250, 30), enBref),
	)
	row2col2 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(60, 30), widget.NewLabel("Branch")),
		container.NewGridWrap(fyne.NewSize(150, 30), enBoff),
	)
	row2col3 := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(80, 30), widget.NewLabel("Cheque Date")),
		container.NewGridWrap(fyne.NewSize(150, 30), enBchqdate),
	)

	headerForm := container.NewVBox(
		container.NewHBox(row1col1, widget.NewLabel("  "), row1col2, widget.NewLabel("  "), row1col3, layout.NewSpacer()),
		container.NewHBox(row2col1, widget.NewLabel("  "), row2col2, widget.NewLabel("  "), row2col3, layout.NewSpacer()),
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("Tax ID")),
			container.NewGridWrap(fyne.NewSize(250, 30), enBcomtaxid),
			widget.NewButtonWithIcon("", theme.SearchIcon(), func() {
				showTaxIDSearch(w, xlOptions, func(taxID, custName string) {
					enBcomtaxid.SetText(taxID)
					if custName != "" && enBnote.Text == "" {
						enBnote.SetText(custName)
					}
					w.Canvas().Focus(enBnote)
				}, func() {
					w.Canvas().Focus(enBcomtaxid)
				})
			}),
			taxLabel,
			layout.NewSpacer(),
		),
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("Note")),
			container.NewGridWrap(fyne.NewSize(800, 30), enBnote),
			layout.NewSpacer(),
		),
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(90, 30), widget.NewLabel("Note2")),
			container.NewGridWrap(fyne.NewSize(800, 30), enBnote2),
			layout.NewSpacer(),
		),
	)

	lineForm := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(90, 30), enAcCode),
		container.NewGridWrap(fyne.NewSize(200, 30), enAcName),
		newEnterButtonWithIcon(theme.SearchIcon(), func() {
			showBookLedgerSearch(w, xlOptions, func(acCode, acName string) {
				enAcCode.SetText(acCode)
				enAcName.SetText(acName)
				w.Canvas().Focus(enScode)
			})
		}),
		container.NewGridWrap(fyne.NewSize(70, 30), enScode),
		newEnterButtonWithIcon(theme.SearchIcon(), func() {
			showBookSubbookSearch(w, xlOptions, comCode, func(scode, sname string) {
				enScode.SetText(scode)
				enSname.SetText(sname)
				w.Canvas().Focus(vatSelect)
			})
		}),
		vatSelect,
		container.NewGridWrap(fyne.NewSize(40, 30), enVatPct),
		whtSelect,
		container.NewGridWrap(fyne.NewSize(40, 30), enWhtPct),
		container.NewGridWrap(fyne.NewSize(100, 30), enBdebit),
		container.NewGridWrap(fyne.NewSize(100, 30), enBcredit),
		widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() { saveAction() }),
	)

	gridHeader := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(100, 25), widget.NewLabelWithStyle("AC CODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(350, 25), widget.NewLabelWithStyle("AC NAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(60, 25), widget.NewLabelWithStyle("BOOK", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(150, 25), widget.NewLabelWithStyle("BOOK NAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(100, 25), widget.NewLabelWithStyle("DEBIT", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(100, 25), widget.NewLabelWithStyle("CREDIT", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
	)

	sumRow := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(200, 28), lblDiff),
		layout.NewSpacer(),
		widget.NewLabelWithStyle("รวม Debit:", fyne.TextAlignTrailing, fyne.TextStyle{Bold: true}),
		container.NewGridWrap(fyne.NewSize(100, 28), lblSumDebit),
		widget.NewLabelWithStyle("Credit:", fyne.TextAlignTrailing, fyne.TextStyle{Bold: true}),
		container.NewGridWrap(fyne.NewSize(100, 28), lblSumCredit),
	)

	topSection := container.NewVBox(
		toolbar,
		widget.NewLabelWithStyle("Voucher", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		widget.NewLabelWithStyle("Voucher Heading", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		headerForm,
		widget.NewSeparator(),
		lineForm,
		widget.NewSeparator(),
		gridHeader,
	)

	// init: เริ่มต้นที่ item แรกของ current period
	targetP := periodCfg.NowPeriod
	if workingPeriod != 0 {
		targetP = workingPeriod
	}
	initBitems := loadAllBitemsInPeriod(xlOptions, comCode, targetP) // ✅ ใช้ loadAllBitemsInPeriod
	if len(initBitems) > 0 {
		loadBitem(initBitems[0], targetP) // ✅ เติม , targetP
	}
	setViewMode()

	// reset := func() {
	// 	if newCfg, err := loadCompanyPeriod(xlOptions); err == nil {
	// 		periodCfg = newCfg
	// 		periodErr = nil
	// 		workingPeriod = 0
	// 		dr1, dr2, _ = getCurrentPeriodRange(periodCfg.YearEnd, periodCfg.TotalPeriods, periodCfg.NowPeriod)
	// 		periodLabel.SetText(getPeriodSummary(periodCfg))
	// 	}
	// 	resetBitems := loadAllBitemsInPeriod(xlOptions, comCode, periodCfg.NowPeriod) // ✅ ใช้ loadAllBitemsInPeriod
	// 	if len(resetBitems) > 0 {
	// 		loadBitem(resetBitems[0], periodCfg.NowPeriod) // ✅ เติม , periodCfg.NowPeriod
	// 	} else {
	// 		currentBitem = ""
	// 		enBitem.SetText("")
	// 		clearHeader()
	// 		lines = nil
	// 		calcSum()
	// 		lineList.Refresh()
	// 	}
	// 	setViewMode()
	// 	checkRepairNeeded()
	// 	fyne.Do(func() { w.Canvas().Focus(nil) })
	// }
	reset := func() {
		// re-read period config ใหม่ทุกครั้ง
		if newCfg, err := loadCompanyPeriod(xlOptions); err == nil {
			periodCfg = newCfg
			periodErr = nil

			// ❌ ลบบรรทัดนี้ทิ้งไปเลยครับ: workingPeriod = 0

			// ✅ หา target period ปัจจุบันที่กำลังทำงานอยู่
			targetP := periodCfg.NowPeriod
			if workingPeriod != 0 {
				targetP = workingPeriod
			}

			// ✅ โหลดวันที่ dr1, dr2 ตาม targetP
			dr1, dr2, _ = getCurrentPeriodRange(periodCfg.YearEnd, periodCfg.TotalPeriods, targetP)

			// ✅ อัปเดต Label ให้ถูกต้อง
			if workingPeriod == 0 {
				periodLabel.SetText(getPeriodSummary(periodCfg))
			} else {
				periodLabel.SetText(fmt.Sprintf("⚡ งวด %d/%d  (%s - %s)",
					targetP, periodCfg.TotalPeriods,
					dr1.Format("02/01/06"), dr2.Format("02/01/06")))
			}
		}

		targetP := periodCfg.NowPeriod
		if workingPeriod != 0 {
			targetP = workingPeriod
		}

		// ✅ โหลด Item ตาม targetP
		resetBitems := loadAllBitemsInPeriod(xlOptions, comCode, targetP)
		if len(resetBitems) > 0 {
			loadBitem(resetBitems[0], targetP)
		} else {
			currentBitem = ""
			enBitem.SetText("")
			clearHeader()
			lines = nil
			calcSum()
			lineList.Refresh()
		}
		setViewMode()
		checkRepairNeeded()
		fyne.Do(func() { w.Canvas().Focus(nil) })
	}

	return container.NewBorder(
		topSection, sumRow, nil, nil,
		container.NewVScroll(lineList),
	), reset
}

// ฟังก์ชันกลางสำหรับเช็คความถูกต้องของทั้งฟอร์ม
// เพิ่ม Parameter เข้าไปในวงเล็บ (ระบุ Type ให้ถูกต้องตามที่ประกาศไว้ตอนต้น)
func validateForm(enBdate *dateEntry, enBvoucher *smartEntry, enBcomtaxid *smartEntry, dr1 time.Time, dr2 time.Time) error {
	// 1. เช็ควันที่
	dateStr := enBdate.GetDate()
	if verr := validateVoucherDate(dateStr, dr1, dr2); verr != nil {
		return verr
	}

	// 2. เช็ค Voucher ห้ามว่าง
	if strings.TrimSpace(enBvoucher.Text) == "" {
		return fmt.Errorf("กรุณากรอก Voucher")
	}

	// 3. เช็ค Tax ID
	taxID := strings.TrimSpace(enBcomtaxid.Text)
	if taxID != "" && len(taxID) != 13 {
		return fmt.Errorf("Tax ID ต้องกรอกตัวเลข 13 หลัก")
	}

	return nil
}
