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
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ==helpers.go

// --- enterButton: ปุ่มที่รับ Enter ได้ ---
// [DONE] enterButton — button that responds to Enter key, working correctly
type enterButton struct {
	widget.Button
}

func newEnterButton(label string, onTapped func()) *enterButton {
	b := &enterButton{}
	b.Text = label
	b.OnTapped = onTapped
	b.ExtendBaseWidget(b)
	return b
}

func newEnterButtonWithIcon(icon fyne.Resource, onTapped func()) *enterButton {
	b := &enterButton{}
	b.Icon = icon
	b.OnTapped = onTapped
	b.ExtendBaseWidget(b)
	return b
}

func (b *enterButton) TypedKey(key *fyne.KeyEvent) {
	if key.Name == fyne.KeyReturn || key.Name == fyne.KeyEnter {
		if b.OnTapped != nil {
			b.OnTapped()
		}
		return
	}
	b.Button.TypedKey(key)
}

// --- escButton: ปุ่มที่รับ Esc ได้ ---
// [DONE] escButton — button that responds to Esc key, working correctly
type escButton struct {
	widget.Button
}

func newEscButton(label string, onTapped func()) *escButton {
	b := &escButton{}
	b.Text = label
	b.OnTapped = onTapped
	b.ExtendBaseWidget(b)
	return b
}

func (b *escButton) TypedKey(key *fyne.KeyEvent) {
	if key.Name == fyne.KeyEscape {
		if b.OnTapped != nil {
			b.OnTapped()
		}
		return
	}
	b.Button.TypedKey(key)
}

// --- popupButton: ปุ่มใน ModalPopUp ที่ ESC ปิด popup เสมอ ---
// [DONE] popupButton — button in ModalPopUp, Esc always closes popup, working correctly
type popupButton struct {
	widget.Button
	onEsc func()
}

func newPopupButton(label string, onTapped func(), onEsc func()) *popupButton {
	b := &popupButton{onEsc: onEsc}
	b.Text = label
	b.OnTapped = onTapped
	b.ExtendBaseWidget(b)
	return b
}

func (b *popupButton) TypedKey(key *fyne.KeyEvent) {
	if key.Name == fyne.KeyEscape {
		if b.onEsc != nil {
			b.onEsc()
		}
		return
	}
	b.Button.TypedKey(key)
}

// --- enterEscButton: ปุ่มที่รับทั้ง Enter (ยืนยัน) และ Esc (ยกเลิก) ---
// [DONE] enterEscButton — button that responds to both Enter and Esc, working correctly
type enterEscButton struct {
	widget.Button
	onEnter func()
	onEsc   func()
}

func newEnterEscButton(label string, onEnter func(), onEsc func()) *enterEscButton {
	b := &enterEscButton{onEnter: onEnter, onEsc: onEsc}
	b.Text = label
	b.OnTapped = onEnter
	b.ExtendBaseWidget(b)
	return b
}

func (b *enterEscButton) TypedKey(key *fyne.KeyEvent) {
	switch key.Name {
	case fyne.KeyReturn, fyne.KeyEnter:
		if b.onEnter != nil {
			b.onEnter()
		}
	case fyne.KeyEscape:
		if b.onEsc != nil {
			b.onEsc()
		}
	default:
		b.Button.TypedKey(key)
	}
}

// ─────────────────────────────────────────────────────────────
// smartEntry — Entry ที่รับ keyboard shortcuts ทั้งหมด
//
//	Esc        → onEsc   : ยกเลิก / กลับ View mode
//	F2         → onF2    : เข้า Edit mode
//	F4         → onF4    : ลบ line ที่ selected
//	Ctrl+S     → onSave  : Save / checkAndSave
//	PageDown   → onPlus  : Next record
//	PageUp     → onMinus : Prev record
//	↓ Arrow    → onDown  : เลื่อน highlight ลง (ใช้ใน search popup)
//	↑ Arrow    → onUp    : เลื่อน highlight ขึ้น (ใช้ใน search popup)
//	Enter      → onEnter : ยืนยัน (ใช้ใน search popup)
//	            ถ้าไม่ได้ set onEnter → ส่งต่อให้ Entry จัดการ (OnSubmitted)
//
// ─────────────────────────────────────────────────────────────
// [DONE] smartEntry — custom entry with onEnter/onEsc/onFocusGained/onFocusLost/nav callbacks, working correctly
type smartEntry struct {
	widget.Entry
	onSave        func()
	onEsc         func()
	onF2          func()
	onF3          func()
	onF4          func()
	onPageDown    func()
	onPageUp      func()
	onDown        func()
	onUp          func()
	onEnter       func()
	onFocusLost   func()
	onFocusGained func()
}

func newSmartEntry(onSave func()) *smartEntry {
	e := &smartEntry{onSave: onSave}
	e.ExtendBaseWidget(e)
	return e
}

// FocusGained — เมื่อได้รับ focus (Tab / คลิก) ให้ Select All ทันที
func (e *smartEntry) FocusGained() {
	e.Entry.FocusGained()
	if e.Text != "" {
		// มี text: SelectAll เพื่อให้พิมพ์ทับได้ทันที (EDIT mode)
		e.Entry.TypedShortcut(&fyne.ShortcutSelectAll{})
	} else {
		// field ว่าง: Entry.FocusGained() อาจ set internal selection
		// กด End key เพื่อ deselect และ move cursor ไปท้าย
		e.Entry.TypedKey(&fyne.KeyEvent{Name: fyne.KeyEnd})
	}
	if e.onFocusGained != nil {
		fn := e.onFocusGained
		go func() { fyne.Do(fn) }() // ต้องเรียก UI จาก main thread
	}
}

func (e *smartEntry) FocusLost() {
	if e.onFocusLost != nil {
		e.onFocusLost()
	}
	e.Entry.FocusLost()
}

func (e *smartEntry) TypedKey(key *fyne.KeyEvent) {
	switch key.Name {
	case fyne.KeyEscape:
		if e.onEsc != nil {
			e.onEsc()
		}
	case fyne.KeyF2:
		if e.onF2 != nil {
			e.onF2()
		}
	case fyne.KeyF3:
		if e.onF3 != nil {
			e.onF3()
		}
	case fyne.KeyF4:
		if e.onF4 != nil {
			e.onF4()
		}
	case fyne.KeyPageDown:
		if e.onPageDown != nil {
			e.onPageDown()
		}
	case fyne.KeyPageUp:
		if e.onPageUp != nil {
			e.onPageUp()
		}
	case fyne.KeyDown:
		if e.onDown != nil {
			e.onDown()
		}
	case fyne.KeyUp:
		if e.onUp != nil {
			e.onUp()
		}
	case fyne.KeyReturn, fyne.KeyEnter:
		if e.onEnter != nil {
			e.onEnter()
		} else {
			e.Entry.TypedKey(key) // → OnSubmitted ทำงานปกติ
		}
	default:
		e.Entry.TypedKey(key)
	}
}

//	func (e *smartEntry) TypedShortcut(s fyne.Shortcut) {
//		if cs, ok := s.(*desktop.CustomShortcut); ok {
//			if cs.Modifier == fyne.KeyModifierControl && cs.KeyName == fyne.KeyS {
//				if e.onSave != nil {
//					e.onSave()
//				}
//				return
//			}
//		}
//		e.Entry.TypedShortcut(s)
//	}
func (e *smartEntry) TypedShortcut(s fyne.Shortcut) {
	if cs, ok := s.(*desktop.CustomShortcut); ok {
		if cs.Modifier == fyne.KeyModifierControl && cs.KeyName == fyne.KeyS {
			if e.onSave != nil {
				e.onSave()
				return // <--- ย้ายเข้ามาตรงนี้
			}
			// ถ้า e.onSave เป็น nil ให้ปล่อยผ่าน ไม่ต้อง return
			// เพื่อให้ Ctrl+S ทะลุไปถึง Window (registerCtrlS)
		}
	}
	e.Entry.TypedShortcut(s)
}

// --- smartMultiEntry: ป้องกัน Tab เพี้ยน + รับ Ctrl+S ---
type smartMultiEntry struct {
	widget.Entry
	onSave       func()
	nextFocusRef fyne.Focusable
}

func newSmartMultiEntry(onSave func()) *smartMultiEntry {
	e := &smartMultiEntry{onSave: onSave}
	e.MultiLine = true
	e.Wrapping = fyne.TextWrapWord
	e.ExtendBaseWidget(e)
	return e
}

func (e *smartMultiEntry) TypedKey(key *fyne.KeyEvent) {
	if key.Name == fyne.KeyTab {
		if e.nextFocusRef != nil {
			canvas := fyne.CurrentApp().Driver().CanvasForObject(e)
			if canvas != nil {
				if obj, ok := e.nextFocusRef.(fyne.CanvasObject); ok {
					if obj.Visible() {
						canvas.Focus(e.nextFocusRef)
					}
				}
			}
		}
		return
	}
	e.Entry.TypedKey(key)
}

func (e *smartMultiEntry) TypedShortcut(s fyne.Shortcut) {
	cs, ok := s.(*desktop.CustomShortcut)
	if ok && cs.Modifier == fyne.KeyModifierControl {
		switch cs.KeyName {
		case fyne.KeyS:
			if e.onSave != nil {
				e.onSave()
			}
			return
		}
	}
	e.Entry.TypedShortcut(s)
}

// registerCtrlS ลงทะเบียน Ctrl+S และ Esc ที่ระดับ Window Canvas
func registerCtrlS(w fyne.Window, saveFunc func(), escFunc func()) {
	ctrlS := &desktop.CustomShortcut{
		KeyName:  fyne.KeyS,
		Modifier: fyne.KeyModifierControl,
	}
	w.Canvas().RemoveShortcut(ctrlS)
	w.Canvas().AddShortcut(ctrlS, func(shortcut fyne.Shortcut) {
		saveFunc()
	})

	escKey := &desktop.CustomShortcut{
		KeyName:  fyne.KeyEscape,
		Modifier: fyne.KeyModifier(0),
	}
	w.Canvas().RemoveShortcut(escKey)
	w.Canvas().AddShortcut(escKey, func(shortcut fyne.Shortcut) {
		if escFunc != nil {
			escFunc()
		}
	})
}

// ─────────────────────────────────────────────────────────────
// autocompleteEntry — ใช้เฉพาะ Login screen (showConnectScreen)
// ยังคงไว้เพราะ main_ui.go ใช้ newAutocompleteEntry() อยู่
// ─────────────────────────────────────────────────────────────
type autocompleteEntry struct {
	widget.Entry
	onDown        func()
	onUp          func()
	onEnter       func()
	onFocusLost   func()
	onFocusGained func()
}

func newAutocompleteEntry() *autocompleteEntry {
	e := &autocompleteEntry{}
	e.ExtendBaseWidget(e)
	return e
}

func (e *autocompleteEntry) TypedKey(key *fyne.KeyEvent) {
	switch key.Name {
	case fyne.KeyDown:
		if e.onDown != nil {
			e.onDown()
		}
		return
	case fyne.KeyUp:
		if e.onUp != nil {
			e.onUp()
		}
		return
	case fyne.KeyReturn, fyne.KeyEnter:
		if e.onEnter != nil {
			e.onEnter()
			return
		}
	}
	e.Entry.TypedKey(key)
}

// --- Utility ---
func getComCodeFromExcel(xlOptions excelize.Options) string {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return ""
	}
	defer f.Close()
	val, _ := f.GetCellValue("Company_Profile", "A2")
	return val
}

func parseFloat(s string) float64 {
	s = strings.ReplaceAll(s, ",", "") // รองรับ comma-formatted เช่น "1,000.00"
	var f float64
	fmt.Sscanf(s, "%f", &f)
	return f
}

// floorVAT — ตัดทศนิยมลงเสมอ 2 ตำแหน่ง (กฎสรรพากร: ห้ามปัดเกิน 0.99)

// [DONE] floorVAT — floor VAT to 2 decimal places per Thai Revenue Dept rules, working correctly
func floorVAT(v float64) string {
	// ดูทศนิยมตำแหน่งที่ 3
	thirdDecimal := int(math.Round(math.Abs(v)*1000)) % 10
	var result float64
	if thirdDecimal >= 5 {
		result = math.Round(v*100) / 100 // ปัดขึ้น
	} else {
		result = math.Floor(v*100+0.0000001) / 100 // ตัดทิ้ง
	}
	return fmt.Sprintf("%.2f", result)
}

func safeGet(row []string, idx int) string {
	if idx < len(row) {
		return row[idx]
	}
	return ""
}

func loadAllBitems(xlOptions excelize.Options, comCode string) []string {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	rows, _ := f.GetRows("Book_items")
	seen := map[string]bool{}

	type bitemInfo struct {
		bitem   string
		bitemNo int
	}
	var items []bitemInfo

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 4 && row[0] == comCode {
			bitem := row[3]
			if !seen[bitem] {
				seen[bitem] = true
				n, _ := strconv.Atoi(bitem)
				items = append(items, bitemInfo{bitem: bitem, bitemNo: n})
			}
		}
	}

	// เรียงตาม item number (numeric) เป็นหลัก
	sort.Slice(items, func(i, j int) bool {
		if items[i].bitemNo != items[j].bitemNo {
			return items[i].bitemNo < items[j].bitemNo
		}
		return items[i].bitem < items[j].bitem
	})

	var result []string
	for _, item := range items {
		result = append(result, item.bitem)
	}
	return result
}

// loadAllBitemsInRange — โหลด Bitem เฉพาะที่ date อยู่ใน [rangeStart, rangeEnd]
// ใช้ตอน Select Period mode เพื่อจำกัด navigation และ ADD/EDIT ในงวดนั้นเท่านั้น
func loadAllBitemsInRange(xlOptions excelize.Options, comCode string, rangeStart, rangeEnd time.Time) []string {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	rows, _ := f.GetRows("Book_items")
	seen := map[string]bool{}

	var items []string

	for i, row := range rows {
		if i == 0 || len(row) < 4 || row[0] != comCode {
			continue
		}
		bitem := row[3]
		if seen[bitem] {
			continue
		}
		t, err := time.Parse("02/01/06", strings.TrimSpace(safeGet(row, 1)))
		if err != nil {
			continue
		}
		if !t.Before(rangeStart) && !t.After(rangeEnd) {
			seen[bitem] = true
			items = append(items, bitem)
		}
	}

	// เรียงตาม item number (numeric) เป็นหลัก
	sort.Slice(items, func(i, j int) bool {
		ni, erri := strconv.Atoi(items[i])
		nj, errj := strconv.Atoi(items[j])
		if erri == nil && errj == nil {
			return ni < nj
		}
		return items[i] < items[j]
	})

	return items
}

func VerifyCompanyCode(xlOptions excelize.Options) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	sheetName := "Company_Profile"

	comCode, _ := f.GetCellValue(sheetName, "A2")
	if comCode != "" {
		return
	}

	fmt.Println("⚠️ ไม่พบ ComCode กำลังตั้งค่าเริ่มต้น C01...")

	h, _ := f.GetCellValue(sheetName, "A1")
	if h == "" {
		f.SetCellValue(sheetName, "A1", "ComCode")
		f.SetCellValue(sheetName, "B1", "ComName")
		f.SetCellValue(sheetName, "C1", "ComAddr")
		f.SetCellValue(sheetName, "D1", "ComTaxID")
		f.SetCellValue(sheetName, "E1", "ComYEnd")
		f.SetCellValue(sheetName, "F1", "ComPeriod")
		f.SetCellValue(sheetName, "G1", "ComNPeriod")
	}

	f.SetCellValue(sheetName, "A2", "C01")

	if err := f.Save(); err != nil {
		fmt.Println("❌ บันทึกไม่สำเร็จ:", err)
		return
	}

	fmt.Println("✅ VerifyCompanyCode: ตั้งค่า ComCode = C01 เรียบร้อยแล้ว")
}

func VerifySpecialAccounts(xlOptions excelize.Options) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	sheetName := "Special_code"
	targetComCode := getComCodeFromExcel(xlOptions)

	// required: CODE → default NAME (ใช้เฉพาะตอน insert ใหม่ ไม่ทับของเดิม)
	requiredData := [][]string{
		{"120VAT", "Purchase VAT"},
		{"120WHT", "120 ถูกหัก ณ ที่จ่าย"},
		{"235WHT", "235 หัก ณ ที่จ่าย"},
		{"235TVAT", "Value Added Tax"},
		{"235VAT", "เจ้าหนี้-กรมสรรพากร"},
		{"350RTE", "Retained Earning"},
		{"360PLA", "Profit (Loss) Account"},
		{"524RND", "ขาดทุนจากการปัดเศษ"},
		{"450RND", "กำไรจากการปัดเศษ"},
	}

	rows, _ := f.GetRows(sheetName)

	// สร้าง header + สร้าง sheet ใหม่ถ้าว่างเปล่า
	if len(rows) == 0 {
		f.SetCellValue(sheetName, "A1", "ComCode")
		f.SetCellValue(sheetName, "B1", "CODE")
		f.SetCellValue(sheetName, "C1", "NAME")
		rows = [][]string{{"ComCode", "CODE", "NAME"}}
	}

	// สร้าง requiredCodes set สำหรับเช็ค valid CODE
	requiredSet := make(map[string]bool)
	for _, item := range requiredData {
		requiredSet[item[0]] = true
	}

	// รวบรวม CODE ที่ valid และ NAME ที่ user แก้ไว้ (เก็บไว้ใช้ต่อ)
	validRows := map[string]string{} // CODE → NAME (เก็บ NAME ที่ user แก้ไว้)
	hasInvalid := false
	for _, row := range rows {
		if len(row) < 2 || row[0] != targetComCode {
			continue
		}
		code := row[1]
		name := ""
		if len(row) >= 3 {
			name = row[2]
		}
		if requiredSet[code] {
			validRows[code] = name // CODE ถูก → เก็บ NAME ไว้
		} else {
			hasInvalid = true // มี CODE ที่ไม่ถูกต้อง
		}
	}

	// ตรวจว่าครบและไม่มี invalid
	allValid := !hasInvalid && len(validRows) == len(requiredData)
	if allValid {
		// ✅ CODE ครบและถูกต้องทั้งหมด → ไม่ทำอะไร
	} else {
		// ❌ มี CODE ผิด หรือขาด → ลบ rows ของ comCode นี้ทิ้ง แล้ว insert ใหม่
		// เก็บ NAME ที่ user แก้ไว้ (จาก validRows) ถ้ามี ใช้ต่อ
		fmt.Println("⚠️ Special_code: พบ CODE ผิดหรือขาด กำลังแก้ไข...")

		// ลบเฉพาะ rows ของ comCode นี้ออก (ลบจากท้ายขึ้นมาเพื่อไม่ให้ index เลื่อน)
		for i := len(rows); i >= 1; i-- {
			if len(rows[i-1]) >= 1 && rows[i-1][0] == targetComCode {
				f.RemoveRow(sheetName, i)
			}
		}

		// อ่าน rows ใหม่หลังลบ เพื่อหา rowNum ถัดไป
		rows2, _ := f.GetRows(sheetName)
		nextRow := len(rows2) + 1

		// insert CODE ที่ถูกต้องพร้อม NAME (ใช้ NAME ที่ user แก้ไว้ถ้ามี ไม่งั้นใช้ default)
		for _, item := range requiredData {
			name := validRows[item[0]] // NAME ที่ user แก้ไว้
			if name == "" {
				name = item[1] // default name
			}
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", nextRow), targetComCode)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", nextRow), item[0])
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", nextRow), name)
			nextRow++
		}

		if err := f.Save(); err != nil {
			fmt.Println("❌ บันทึกไม่สำเร็จ:", err)
			return
		}
		fmt.Println("✅ VerifySpecialAccounts: แก้ไข Special_code เรียบร้อย")
	}

	// ── ยัด Special Codes เข้า Ledger_Master ด้วย (ถ้ายังไม่มี) ──
	verifySpecialCodesInLedger(f, targetComCode)
}

// verifySpecialCodesInLedger — ตรวจและเพิ่ม special codes เข้า Ledger_Master อัตโนมัติ
// เรียกหลัง VerifySpecialAccounts เสมอ — ไม่ทับยอดที่มีอยู่แล้ว
func verifySpecialCodesInLedger(f *excelize.File, comCode string) {
	const ledger = "Ledger_Master"
	const zeroCols = 29 // BBAL,CBAL,Debit,Credit,Bthisyear,Thisper01-12,Blastyear,Lastper01-12

	// Gcode/Gname สำหรับ insert ใหม่ (ไม่เปลี่ยนแปลง)
	type scEntry struct{ Gcode, Gname string }
	specialMeta := map[string]scEntry{
		"120VAT":  {"120", "เงินลงทุนระยะสั้น"},
		"120WHT":  {"120", "เงินลงทุนระยะสั้น"},
		"235WHT":  {"235", "หนี้สินหมุนเวียนอื่น"},
		"235TVAT": {"235", "หนี้สินหมุนเวียนอื่น"},
		"235VAT":  {"235", "หนี้สินหมุนเวียนอื่น"},
		"350RTE":  {"350", "กำไรสะสม"},
		"360PLA":  {"360", "กำไร(ขาดทุน)สุทธิ"},
		"524RND":  {"524", "ค่าใช้จ่ายอื่น"},
		"450RND":  {"450", "รายได้อื่น"},
	}

	// ดึง NAME จาก Special_code sheet → source of truth
	specialNames := map[string]string{}
	scRows, _ := f.GetRows("Special_code")
	for _, row := range scRows {
		if len(row) >= 3 && row[0] == comCode {
			specialNames[row[1]] = row[2]
		}
	}

	ledgerRows, _ := f.GetRows(ledger)

	// สร้าง header ถ้ายังไม่มี
	if len(ledgerRows) == 0 {
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
			f.SetCellValue(ledger, fmt.Sprintf("%s1", col), h)
		}
		ledgerRows = append(ledgerRows, headers)
	}

	// รวบรวม AcCode ที่มีอยู่แล้ว
	existing := map[string]bool{}
	for i, row := range ledgerRows {
		if i == 0 || len(row) < 2 {
			continue
		}
		if row[0] == comCode {
			existing[row[1]] = true
		}
	}

	// เพิ่มเฉพาะที่ยังไม่มี
	added := 0
	for _, acCode := range []string{"120VAT", "120WHT", "235TVAT", "235WHT", "235VAT", "350RTE", "360PLA", "524RND", "450RND"} {
		if existing[acCode] {
			continue
		}
		meta := specialMeta[acCode]
		acName := specialNames[acCode] // ดึง NAME จาก Special_code sheet
		if acName == "" {
			acName = acCode // fallback ถ้ายังไม่มีใน Special_code
		}
		rowNum := len(ledgerRows) + 1
		f.SetCellValue(ledger, fmt.Sprintf("A%d", rowNum), comCode)
		f.SetCellValue(ledger, fmt.Sprintf("B%d", rowNum), acCode)
		f.SetCellValue(ledger, fmt.Sprintf("C%d", rowNum), acName)
		f.SetCellValue(ledger, fmt.Sprintf("D%d", rowNum), meta.Gcode)
		f.SetCellValue(ledger, fmt.Sprintf("E%d", rowNum), meta.Gname)
		for j := 0; j < zeroCols; j++ {
			col, _ := excelize.ColumnNumberToName(6 + j)
			f.SetCellValue(ledger, fmt.Sprintf("%s%d", col, rowNum), 0)
		}
		ledgerRows = append(ledgerRows, []string{comCode, acCode}) // อัพเดท slice ป้องกัน dup
		added++
	}

	if added > 0 {
		f.Save()
		fmt.Printf("✅ verifySpecialCodesInLedger: เพิ่ม %d special codes เข้า Ledger_Master\n", added)
	}
}

func VerifyAcctGroup(xlOptions excelize.Options) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	sheetName := "Acct_Group"
	targetComCode := getComCodeFromExcel(xlOptions)
	if targetComCode == "" {
		return
	}

	presetData := [][]string{
		// Assets
		{"100", "เงินสดและรายการเทียบเท่าเงินสด", "Cash and Cash Equivalents"},
		{"111", "เงินฝากธนาคาร", "Bank Deposits"},
		{"112", "เงินลงทุนระยะสั้น", "Short-Term Investments"},
		{"113", "ลูกหนี้การค้าและตั๋วเงินรับ", "Accounts and Notes Receivable"},
		{"114", "ตั๋วรับเงิน", "Notes Receivable"},
		{"115", "สำรองหนี้สูญ", "Allowance for Doubtful Accounts"},
		{"116", "เงินให้กู้ยืมระยะสั้นจากบริษัทในเครือ", "Short-Term Loans to Subsidiaries"},
		{"117", "สินค้าสำเร็จรูป", "Finished Goods"},
		{"118", "งานระหว่างทาง", "Work in Process"},
		{"119", "วัตถุดิบ", "Raw Materials"},
		{"120", "สินทรัพย์หมุนเวียนอื่น", "Other Current Assets"},
		{"150", "ลูกหนี้และเงินให้กู้ยืมแก่กรรมการ", "Receivables and Loans to Directors"},
		{"160", "เงินลงทุนระยะยาว", "Long-Term Investments"},
		{"170", "ที่ดิน อาคารและอุปกรณ์", "Property, Plant and Equipment"},
		{"180", "ค่าเสื่อมราคาสะสม", "Accumulated Depreciation"},
		{"190", "สินทรัพย์ไม่หมุนเวียนอื่น", "Other Non-Current Assets"},
		// Liabilities
		{"200", "เงินเบิกเกินและเงินกู้ธนาคาร", "Bank Overdrafts and Loans"},
		{"210", "เจ้าหนี้การค้าและตั๋วเงินจ่าย", "Accounts and Notes Payable"},
		{"215", "ตั๋วเงินนำจ่าย", "Notes Payable"},
		{"220", "เงินปันผลนำจ่าย", "Dividends Payable"},
		{"225", "หนี้ระยะยาวถึงกำหนดชำระ", "Long-Term Debt Due Within One Year"},
		{"230", "เงินกู้ยืมระยะยาวที่ครบกำหนดในหนึ่งปี", "Current Portion of Long-Term Loans"},
		{"235", "หนี้สินหมุนเวียนอื่น", "Other Current Liabilities"},
		{"250", "เงินกู้ยืมบริษัทในเครือและบริษัทร่วม", "Loans from Subsidiaries and Associates"},
		{"260", "เงินกู้ยืมระยะยาว", "Long-Term Loans"},
		{"265", "เงินกู้ระยะยาวจากบริษัทในเครือ", "Long-Term Loans from Subsidiaries"},
		{"270", "เงินทุนเลี้ยงชีพและบำเหน็จ", "Provident Fund and Severance Pay"},
		{"275", "เงินกู้ระยะยาว — อื่น", "Other Long-Term Borrowings"},
		{"280", "หนี้สินอื่น", "Other Liabilities"},
		// Equity
		{"300", "ทุนที่ออกและเรียกชำระแล้ว", "Issued and Paid-Up Share Capital"},
		{"302", "ส่วนเกินมูลค่าหุ้น", "Share Premium"},
		{"350", "กำไรสะสม", "Retained Earnings"},
		{"351", "กำไรสะสมสำรองตามกฎหมาย", "Legal Reserve"},
		{"352", "กำไรสะสมสำรองอื่น", "Other Reserves"},
		{"360", "กำไรสุทธิประจำปี", "Net Profit for the Year"},
		// Revenue
		{"400", "รายได้", "Revenue"},
		{"402", "รับคืน", "Sales Returns"},
		{"410", "ส่วนลดจ่าย", "Sales Discounts"},
		{"450", "รายได้อื่น", "Other Income"},
		// Cost of Goods Sold
		{"500", "ซื้อวัตถุดิบ", "Raw Materials Purchases"},
		{"502", "ค่าใช้จ่ายในการซื้อวัตถุดิบ", "Purchasing Expenses - Raw Materials"},
		{"504", "ตัวหักในการซื้อวัตถุดิบ", "Purchase Returns - Raw Materials"},
		{"506", "ค่าแรงงานทางตรง", "Direct Labor"},
		{"508", "โสหุ้ยอุปกรณ์โรงงาน", "Factory Overhead"},
		{"510", "ซื้อสินค้าสำเร็จรูป", "Finished Goods Purchases"},
		{"512", "ส่วนลดรับในการซื้อสินค้า", "Purchase Discounts"},
		{"514", "ค่าใช้จ่ายในการซื้อสินค้า", "Purchasing Expenses - Goods"},
		{"516", "ต้นทุนขายอื่น", "Other Cost of Sales"},
		// Expenses
		{"520", "ค่าใช้จ่ายในการขาย", "Selling Expenses"},
		{"522", "ค่าใช้จ่ายในการบริหาร", "Administrative Expenses"},
		{"524", "ค่าใช้จ่ายอื่น", "Other Expenses"},
		{"530", "ต้นทุนทางการเงิน", "Finance Costs"},
		{"540", "ภาษีเงินได้นิติบุคคล", "Corporate Income Tax"},
		{"550", "รายการพิเศษ", "Extraordinary Items"},
		// Adjustments
		{"600", "ปรับปรุงสินค้า", "Inventories Adjustment"},
	}

	rows, _ := f.GetRows(sheetName)
	existingCodes := make(map[string]bool)
	for _, row := range rows {
		if len(row) >= 2 && row[0] == targetComCode {
			existingCodes[row[1]] = true
		}
	}

	isMissing := false
	for _, item := range presetData {
		if !existingCodes[item[0]] {
			isMissing = true
			break
		}
	}

	if isMissing {
		fmt.Println("⚠️ Acct_Group ไม่ครบ กำลังล้างและยัดใหม่...")

		for i := len(rows); i >= 1; i-- {
			f.RemoveRow(sheetName, i)
		}

		f.SetCellValue(sheetName, "A1", "Comcode")
		f.SetCellValue(sheetName, "B1", "Gcode")
		f.SetCellValue(sheetName, "C1", "Gname")
		f.SetCellValue(sheetName, "D1", "GnameEN")

		for i, item := range presetData {
			rowNum := i + 2
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), targetComCode)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), item[0])
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), item[1])
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), item[2])
		}

		if err := f.Save(); err != nil {
			fmt.Println("❌ บันทึกไม่สำเร็จ:", err)
			return
		}
		fmt.Println("✅ VerifyAcctGroup: ยัดข้อมูล preset เรียบร้อยแล้ว")
	}
}

func VerifyCapitalInfo(xlOptions excelize.Options) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	sheetName := "Capital"
	targetComCode := getComCodeFromExcel(xlOptions)

	// เช็ค Header คอลัมน์ B ว่าเป็นโครงสร้างใหม่ (แนวนอน) หรือยัง
	headerB, _ := f.GetCellValue(sheetName, "B1")

	// ถ้ายังไม่ใช่โครงสร้างใหม่ (ThisYearQty) ให้ทำการ Reset/Migrate ทันที
	if headerB != "ThisYearQty" {
		fmt.Println("⚠️ โครงสร้าง Capital เป็นแบบเก่า หรือยังไม่มีข้อมูล กำลัง Reset เป็นแนวนอน...")

		rows, _ := f.GetRows(sheetName)
		// ลบข้อมูลเก่าทิ้งทั้งหมด
		for i := len(rows); i >= 1; i-- {
			f.RemoveRow(sheetName, i)
		}

		// สร้าง Header แนวนอน
		headers := []string{"Comcode", "ThisYearQty", "ThisYearValue", "LastYearQty", "LastYearValue"}
		for i, h := range headers {
			col, _ := excelize.ColumnNumberToName(i + 1)
			f.SetCellValue(sheetName, fmt.Sprintf("%s1", col), h)
		}

		// เพิ่มข้อมูลเริ่มต้นเป็น 0 สำหรับบริษัทนี้
		f.SetCellValue(sheetName, "A2", targetComCode)
		f.SetCellValue(sheetName, "B2", "0") // ThisYearQty
		f.SetCellValue(sheetName, "C2", "0") // ThisYearValue
		f.SetCellValue(sheetName, "D2", "0") // LastYearQty
		f.SetCellValue(sheetName, "E2", "0") // LastYearValue

		f.SaveAs(currentDBPath, xlOptions)
	} else {
		// ถ้าเป็นโครงสร้างใหม่แล้ว เช็คว่ามี Comcode นี้หรือยัง
		rows, _ := f.GetRows(sheetName)
		hasData := false
		for i, row := range rows {
			if i > 0 && len(row) > 0 && row[0] == targetComCode {
				hasData = true
				break
			}
		}

		// ถ้ามี Header แล้ว แต่ยังไม่มีบรรทัดของบริษัทนี้ ให้เพิ่มเข้าไป
		if !hasData {
			nextRow := len(rows) + 1
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", nextRow), targetComCode)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", nextRow), "0")
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", nextRow), "0")
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", nextRow), "0")
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", nextRow), "0")
			f.SaveAs(currentDBPath, xlOptions)
		}
	}
}

// ─────────────────────────────────────────────────────────────
// dateEntry — Masked date input แบบ dd/mm/yy
//
//   - "/" ค้างอยู่เสมอ ผู้ใช้พิมพ์แค่ตัวเลข
//   - cursor กระโดดข้าม "/" อัตโนมัติ
//   - FocusGained → Select All (เหมือน smartEntry)
//   - onTab, onEsc, onEnter รองรับ keyboard navigation
//
// ─────────────────────────────────────────────────────────────
type dateEntry struct {
	widget.Entry
	onTab      func()
	onEsc      func()
	onEnter    func()
	onF3       func()
	onPageDown func()
	onPageUp   func()
	onSave     func() // Ctrl+S
}

func newDateEntry() *dateEntry {
	e := &dateEntry{}
	e.ExtendBaseWidget(e)
	e.SetText("__/__/__")
	return e
}

// SetDate — set ค่าจาก string "dd/mm/yy"
func (e *dateEntry) SetDate(s string) {
	if s == "" {
		e.Entry.SetText("__/__/__")
		return
	}
	e.Entry.SetText(s)
}

// GetDate — คืนค่า "dd/mm/yy" หรือ "" ถ้ายังไม่ได้พิมพ์
func (e *dateEntry) GetDate() string {
	t := e.Text
	if t == "__/__/__" || t == "" {
		return ""
	}
	return t
}

func (e *dateEntry) FocusGained() {
	e.Entry.FocusGained()
	e.Entry.TypedShortcut(&fyne.ShortcutSelectAll{})
}

func (e *dateEntry) TypedShortcut(s fyne.Shortcut) {
	if cs, ok := s.(*desktop.CustomShortcut); ok {
		if cs.Modifier == fyne.KeyModifierControl && cs.KeyName == fyne.KeyS {
			if e.onSave != nil {
				e.onSave()
			}
			return
		}
	}
	e.Entry.TypedShortcut(s)
}

func (e *dateEntry) TypedKey(key *fyne.KeyEvent) {
	switch key.Name {
	case fyne.KeyEscape:
		if e.onEsc != nil {
			e.onEsc()
		}
	case fyne.KeyReturn, fyne.KeyEnter:
		if e.onEnter != nil {
			e.onEnter()
		}
	case fyne.KeyTab:
		if e.onTab != nil {
			e.onTab()
		}
	case fyne.KeyF3:
		if e.onF3 != nil {
			e.onF3()
		}
	case fyne.KeyPageDown:
		if e.onPageDown != nil {
			e.onPageDown()
		}
	case fyne.KeyPageUp:
		if e.onPageUp != nil {
			e.onPageUp()
		}
	case fyne.KeyBackspace:
		e.handleBackspace()
	default:
		e.Entry.TypedKey(key)
	}
}

func (e *dateEntry) TypedRune(r rune) {
	if r < '0' || r > '9' {
		return
	}

	cur := []rune(e.Text)
	if len(cur) != 8 {
		cur = []rune("__/__/__")
	}

	// ── เพิ่ม: ถ้าไม่มี _ เลย (field เต็ม) → reset แล้วเริ่มพิมพ์ใหม่ ──
	hasBlank := false
	for _, c := range cur {
		if c == '_' {
			hasBlank = true
			break
		}
	}
	if !hasBlank {
		cur = []rune("__/__/__")
	}

	pos := -1
	for i, c := range cur {
		if c == '_' {
			pos = i
			break
		}
	}
	if pos == -1 {
		return
	}

	cur[pos] = r
	e.Entry.SetText(string(cur))

	full := true
	for _, c := range cur {
		if c == '_' {
			full = false
			break
		}
	}
	if full && e.onEnter != nil {
		e.onEnter()
	}
}

func (e *dateEntry) handleBackspace() {
	cur := []rune(e.Text)
	if len(cur) != 8 {
		return
	}
	// หา position สุดท้ายที่ไม่ใช่ '_' และไม่ใช่ '/'
	for i := len(cur) - 1; i >= 0; i-- {
		if cur[i] != '_' && cur[i] != '/' {
			cur[i] = '_'
			e.Entry.SetText(string(cur))
			return
		}
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// showInlineDropdown — dropdown ใต้ field ขณะพิมพ์ (autocomplete)
// pattern: suggWrapper VBox ใน layout (ไม่ใช้ PopUp/Overlay → ไม่ขโมย focus)
// dropWrapper: container.NewVBox() ที่วางใต้ entry ใน layout ของ caller
// ─────────────────────────────────────────────────────────────────────────────
func showInlineDropdown(
	entry *smartEntry,
	dropWrapper *fyne.Container,
	allItems []string,
	isActive func() bool,
	onSelect func(val string),
) {
	var suggList *widget.List
	var filtered []string
	selectedIdx := -1
	isNavigating := false
	suppressDrop := false // ป้องกัน OnChanged re-open หลัง SetText จาก onSelect

	hide := func() {
		selectedIdx = -1
		dropWrapper.Objects = nil
		dropWrapper.Refresh()
	}

	doSelect := func(id int) {
		if id < 0 || id >= len(filtered) {
			return
		}
		val := filtered[id]
		suppressDrop = true
		hide()
		onSelect(val)
		suppressDrop = false
	}

	suggList = widget.NewList(
		func() int { return len(filtered) },
		func() fyne.CanvasObject { return widget.NewLabel("") },
		func(id widget.ListItemID, o fyne.CanvasObject) {
			if int(id) < len(filtered) {
				o.(*widget.Label).SetText(filtered[id])
			}
		},
	)
	suggList.OnSelected = func(id widget.ListItemID) {
		if isNavigating {
			return
		}
		doSelect(int(id))
	}

	entry.onDown = func() {
		if len(filtered) == 0 || len(dropWrapper.Objects) == 0 {
			return
		}
		isNavigating = true
		if selectedIdx < len(filtered)-1 {
			selectedIdx++
		}
		suggList.Select(widget.ListItemID(selectedIdx))
		suggList.ScrollTo(widget.ListItemID(selectedIdx))
		isNavigating = false
	}

	entry.onUp = func() {
		if len(filtered) == 0 || selectedIdx <= 0 {
			return
		}
		isNavigating = true
		selectedIdx--
		suggList.Select(widget.ListItemID(selectedIdx))
		suggList.ScrollTo(widget.ListItemID(selectedIdx))
		isNavigating = false
	}

	entry.onEnter = func() {
		if selectedIdx >= 0 && selectedIdx < len(filtered) {
			doSelect(selectedIdx)
			return
		}
		hide()
		if entry.OnSubmitted != nil {
			entry.OnSubmitted(entry.Text)
		}
	}

	origEsc := entry.onEsc
	entry.onEsc = func() {
		if len(dropWrapper.Objects) > 0 {
			hide()
			return
		}
		if origEsc != nil {
			origEsc()
		}
	}

	entry.OnChanged = func(kw string) {
		if suppressDrop {
			return
		}
		if isActive != nil && !isActive() {
			hide()
			return
		}
		kwLow := strings.ToLower(strings.TrimSpace(kw))
		filtered = nil
		selectedIdx = -1
		for _, v := range allItems {
			if kwLow != "" && strings.Contains(strings.ToLower(v), kwLow) {
				filtered = append(filtered, v)
			}
		}
		suggList.Refresh()
		if len(filtered) == 0 {
			hide()
			return
		}
		maxShow := 6
		if len(filtered) < maxShow {
			maxShow = len(filtered)
		}
		if len(dropWrapper.Objects) == 0 {
			dropWrapper.Add(container.NewGridWrap(fyne.NewSize(240, float32(28)*float32(maxShow)), suggList))
			dropWrapper.Refresh()
		} else {
			dropWrapper.Refresh()
		}
	}
}

// formatComma — format float64 เป็น string พร้อม comma เช่น 1,234,567.89
func formatComma(v float64) string {
	neg := v < 0
	if neg {
		v = -v
	}
	intPart := int64(v)
	dec := int64((v-float64(intPart))*100 + 0.5)
	s := fmt.Sprintf("%d", intPart)
	// ใส่ comma ทุก 3 หลัก
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
// syncSpecialCodeName — sync NAME ระหว่าง Special_code ↔ Ledger_Master
//
// direction:
//
//	"special→ledger" : แก้ Special_code.NAME → อัพเดท Ledger_Master.Ac_name
//	"ledger→special" : แก้ Ledger_Master.Ac_name → อัพเดท Special_code.NAME
//
// f ต้องเปิดอยู่แล้ว (caller เป็นคนปิด)
// ─────────────────────────────────────────────────────────────────
func syncSpecialCodeName(f *excelize.File, comCode, acCode, newName, direction string) {
	switch direction {
	case "special→ledger":
		// อัพเดท Ledger_Master.Ac_name ที่ตรงกับ acCode
		rows, _ := f.GetRows("Ledger_Master")
		for i, row := range rows {
			if len(row) >= 3 && row[0] == comCode && row[1] == acCode {
				f.SetCellValue("Ledger_Master", fmt.Sprintf("C%d", i+1), newName)
				break
			}
		}

	case "ledger→special":
		// อัพเดท Special_code.NAME ที่ตรงกับ acCode
		rows, _ := f.GetRows("Special_code")
		for i, row := range rows {
			if len(row) >= 3 && row[0] == comCode && row[1] == acCode {
				f.SetCellValue("Special_code", fmt.Sprintf("C%d", i+1), newName)
				break
			}
		}
	}
}

// showErrorDialog — แสดง error dialog ที่กด Enter/ESC ปิดได้
// ใช้แทน dialog.ShowError ทั่วระบบ
func showErrorDialog(title, msg string, w fyne.Window, afterClose func()) {
	var d dialog.Dialog
	okBtn := newEnterButton("OK", func() {
		d.Hide()
		if afterClose != nil {
			afterClose()
		}
	})
	d = dialog.NewCustomWithoutButtons(title,
		container.NewVBox(
			widget.NewLabel(msg),
			container.NewCenter(okBtn),
		), w)
	d.Show()
	fyne.Do(func() { w.Canvas().Focus(okBtn) })
}

// showInfoDialog — แสดง info dialog ที่กด Enter/ESC ปิดได้
func showInfoDialog(title, msg string, w fyne.Window, afterClose func()) {
	showErrorDialog(title, msg, w, afterClose)
}

func LoadCurrentPeriod(xlOptions excelize.Options) int {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return 1 // Safety fallback: ถ้าเปิดไฟล์ไม่ได้ให้ถือว่าเป็นงวด 1
	}
	defer f.Close()

	// ใน company_setup_ui บันทึก Period ปัจจุบันไว้ที่ G2
	val, err := f.GetCellValue("Company_Profile", "G2")
	if err != nil || val == "" {
		return 1
	}

	period, err := strconv.Atoi(val)
	if err != nil {
		return 1
	}

	return period
}

// loadAllBitemsInPeriod — โหลด Bitem เฉพาะใน Period ที่ระบุ
func loadAllBitemsInPeriod(xlOptions excelize.Options, comCode string, targetPeriod int) []string {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	rows, _ := f.GetRows("Book_items")
	seen := map[string]bool{}
	var items []string

	for i, row := range rows {
		if i == 0 || len(row) < 4 || row[0] != comCode {
			continue
		}
		var bp int
		if len(row) > 21 {
			fmt.Sscanf(safeGet(row, 21), "%d", &bp)
		}
		if bp == targetPeriod {
			bitem := row[3]
			if !seen[bitem] {
				seen[bitem] = true
				items = append(items, bitem)
			}
		}
	}

	sort.Slice(items, func(i, j int) bool {
		ni, _ := strconv.Atoi(items[i])
		nj, _ := strconv.Atoi(items[j])
		return ni < nj
	})
	return items
}

// ─────────────────────────────────────────────────────────────────
// getAccountNameMap — โหลด map[acCode]acName จาก Ledger_Master
//
// version 1: เปิดไฟล์เอง (ใช้เมื่อยังไม่มี *excelize.File เปิดอยู่)
// version 2: ใช้ไฟล์ที่เปิดอยู่แล้ว (ใช้ใน report ที่เปิด f ไว้แล้ว — ไม่เปิดซ้ำ)
// core:      buildAcNameMapFromFile — private, รับ *excelize.File โดยตรง
//
// Fallback Pattern (ใช้ใน caller):
//
//	acName := acctMap[acCode]
//	if acName == "" {
//	    acName = safeGet(row, 6) // snapshot fallback จาก Book_items col Ac_name
//	    if acName != "" {
//	        acName = "[" + acName + "]" // บ่งบอกว่าชื่อนี้มาจาก snapshot (account อาจถูกลบแล้ว)
//	    }
//	}
//
// ─────────────────────────────────────────────────────────────────

// getAccountNameMap — version 1: เปิดไฟล์เอง
func getAccountNameMap(xlOptions excelize.Options) map[string]string {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return make(map[string]string)
	}
	defer f.Close()
	return buildAcNameMapFromFile(f, getComCodeFromExcel(xlOptions))
}

// getAccountNameMapFromFile — version 2: ใช้ไฟล์ที่เปิดอยู่แล้ว (ไม่เปิดซ้ำ)
func getAccountNameMapFromFile(f *excelize.File, comCode string) map[string]string {
	return buildAcNameMapFromFile(f, comCode)
}

// buildAcNameMapFromFile — core logic (private)
func buildAcNameMapFromFile(f *excelize.File, comCode string) map[string]string {
	acctMap := make(map[string]string)
	rows, _ := f.GetRows("Ledger_Master")
	for i, row := range rows {
		if i == 0 || len(row) < 3 || row[0] != comCode {
			continue
		}
		acctMap[strings.TrimSpace(row[1])] = strings.TrimSpace(row[2])
	}
	return acctMap
}
