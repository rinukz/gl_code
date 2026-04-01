package main

import (
	"fmt"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ---------------------------------------------------------------------------
// โครงสร้างข้อมูลสินค้าคงเหลือปลายงวด
// ---------------------------------------------------------------------------
type InventoryItem struct {
	RowIndex  int // row index ใน sheet (1-based) สำหรับ Edit/Delete
	ProdCode  string
	ProdName  string
	Qty       float64
	UnitCost  float64
	TotalAmt  float64
	AcctCode  string
	CountDate string
}

// ---------------------------------------------------------------------------
// invNumEntry — รับเฉพาะตัวเลขและจุดทศนิยม (ใช้เฉพาะใน inventory_ui)
// ใช้ TypedRune แทน OnChanged เพื่อ block ตัวอักษรทันทีที่กด
// ---------------------------------------------------------------------------
type invNumEntry struct {
	widget.Entry
	onSave func()
}

func newInvNumEntry(onSave func()) *invNumEntry {
	e := &invNumEntry{onSave: onSave}
	e.ExtendBaseWidget(e)
	return e
}

func (e *invNumEntry) TypedRune(r rune) {
	if (r >= '0' && r <= '9') || r == '.' {
		e.Entry.TypedRune(r)
	}
}

func (e *invNumEntry) TypedShortcut(s fyne.Shortcut) {
	cs, ok := s.(*desktop.CustomShortcut)
	if ok && cs.Modifier == fyne.KeyModifierControl && cs.KeyName == fyne.KeyS {
		if e.onSave != nil {
			e.onSave()
		}
		return
	}
	e.Entry.TypedShortcut(s)
}

// ---------------------------------------------------------------------------
// ตรวจสอบและสร้าง Sheet สำหรับเก็บข้อมูลสินค้าคงเหลือ
// ---------------------------------------------------------------------------
func ensureInventorySheet(xlOpts excelize.Options) error {
	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err != nil {
		return err
	}
	defer f.Close()
	idx, _ := f.GetSheetIndex("Ending_Inventory")
	if idx == -1 {
		f.NewSheet("Ending_Inventory")
		f.SetSheetRow("Ending_Inventory", "A1", &[]interface{}{
			"ComCode", "Period", "CountDate", "ProdCode", "ProdName", "Qty", "UnitCost", "TotalAmt", "AcctCode",
		})
		return f.Save()
	}
	return nil
}

// ---------------------------------------------------------------------------
// showInventorySearch — F3 Popup ค้นหารายการสินค้า (pattern เดียวกับ showBookSearch)
// ---------------------------------------------------------------------------
func showInventorySearch(w fyne.Window, xlOpts excelize.Options, periodNo string,
	onSelect func(item InventoryItem)) {

	// โหลดรายการทั้งหมดใน period นี้
	var allItems []InventoryItem
	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err == nil {
		rows, _ := f.GetRows("Ending_Inventory")
		for i, row := range rows {
			if i == 0 || len(row) < 9 {
				continue
			}
			if row[0] != currentCompanyCode || row[1] != periodNo {
				continue
			}
			allItems = append(allItems, InventoryItem{
				RowIndex:  i + 1,
				ProdCode:  row[3],
				ProdName:  row[4],
				Qty:       parseFloat(row[5]),
				UnitCost:  parseFloat(row[6]),
				TotalAmt:  parseFloat(row[7]),
				AcctCode:  row[8],
				CountDate: row[2],
			})
		}
		f.Close()
	}

	filtered := make([]InventoryItem, len(allItems))
	copy(filtered, allItems)
	selectedIdx := 0

	var list *widget.List
	var pop *widget.PopUp

	doSelect := func(id int) {
		if id < 0 || id >= len(filtered) {
			return
		}
		it := filtered[id]
		pop.Hide()
		onSelect(it)
	}

	highlightOnly := func(id int) {
		if id < 0 || id >= len(filtered) {
			return
		}
		selectedIdx = id
		list.Select(widget.ListItemID(selectedIdx))
		list.ScrollTo(widget.ListItemID(selectedIdx))
	}

	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("ค้นหารหัสสินค้า / ชื่อสินค้า... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")
	searchEntry.onEsc = func() { pop.Hide() }
	searchEntry.onDown = func() {
		if len(filtered) == 0 {
			return
		}
		next := selectedIdx + 1
		if next >= len(filtered) {
			next = len(filtered) - 1
		}
		highlightOnly(next)
	}
	searchEntry.onUp = func() {
		if len(filtered) == 0 {
			return
		}
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}
	searchEntry.onEnter = func() { doSelect(selectedIdx) }

	list = widget.NewList(
		func() int { return len(filtered) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			row := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(80, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(200, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(70, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(90, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(40, 28), widget.NewLabel("")),
			)
			return container.NewMax(btn, row)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			it := filtered[id]
			c := o.(*fyne.Container)
			c.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
			row := c.Objects[1].(*fyne.Container)
			row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(it.ProdCode)
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(it.ProdName)
			row.Objects[2].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("%.2f", it.Qty))
			row.Objects[3].(*fyne.Container).Objects[0].(*widget.Label).SetText(bsNum(it.TotalAmt))
			row.Objects[4].(*fyne.Container).Objects[0].(*widget.Label).SetText(it.AcctCode)
		},
	)

	searchEntry.OnChanged = func(kw string) {
		kw = strings.ToLower(strings.TrimSpace(kw))
		filtered = nil
		selectedIdx = 0
		for _, it := range allItems {
			if kw == "" ||
				strings.Contains(strings.ToLower(it.ProdCode), kw) ||
				strings.Contains(strings.ToLower(it.ProdName), kw) {
				filtered = append(filtered, it)
			}
		}
		list.UnselectAll()
		list.Refresh()
		if len(filtered) > 0 {
			list.Select(0)
			list.ScrollTo(0)
		}
	}

	listHeader := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(80, 25), widget.NewLabelWithStyle("รหัสสินค้า", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(200, 25), widget.NewLabelWithStyle("ชื่อสินค้า", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(70, 25), widget.NewLabelWithStyle("จำนวน", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(90, 25), widget.NewLabelWithStyle("มูลค่า", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(40, 25), widget.NewLabelWithStyle("GL", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
	)

	content := container.NewBorder(
		container.NewVBox(searchEntry, listHeader),
		nil, nil, nil,
		list,
	)

	closeBtn := widget.NewButton("✕ ปิด", func() { pop.Hide() })

	pop = widget.NewModalPopUp(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหารายการสินค้า (F3)", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			container.NewGridWrap(fyne.NewSize(500, 400), content),
		),
		w.Canvas(),
	)

	pop.Show()
	if len(filtered) > 0 {
		list.Select(0)
		list.ScrollTo(0)
	}

	// Focus searchEntry หลัง popup render เสร็จ (เหมือน book_search_ui)
	go func() {
		time.Sleep(350 * time.Millisecond)
		fyne.Do(func() {
			if w.Canvas() != nil {
				w.Canvas().Focus(nil)
				w.Canvas().Focus(searchEntry)
			}
		})
	}()
}

// ---------------------------------------------------------------------------
// UI หลัก
// ---------------------------------------------------------------------------
// buildInventoryAcctOptions อ่าน ac_code จาก Ledger_Master
// filter เฉพาะที่ prefix 3 หลักเป็น 117, 118, 119
// ถ้าไม่พบ code ใดเลย fallback เป็น hardcode default
// ensureInventoryLedgerCodes ตรวจสอบว่า Ledger_Master มี 117/118/119/600 ครบไหม
// ถ้าไม่มี code ใด → auto-add เข้าไปทันที ค่าตัวเลขทั้งหมด = 0
// คืนค่า list ชื่อ code ที่เพิ่งเพิ่มเข้าไป (ไว้แจ้งผู้ใช้)
func ensureInventoryLedgerCodes(xlOpts excelize.Options, f *excelize.File) []string {
	comCode := getComCodeFromExcel(xlOpts)

	// required: {prefix3, acCode(suffix), ac_name, gcode, gname}
	// acCode ใช้ suffix เหมือน pattern ใน seed_ledger เพื่อให้รายงานดึงได้ถูกต้อง
	// prefix3 ใช้ตรวจว่ามี ac_code ที่ขึ้นต้นด้วย prefix นี้แล้วหรือยัง
	required := []struct {
		prefix, acCode, acName, gcode, gname string
	}{
		{"117", "117001", "สินค้าสำเร็จรูป", "117", "สินค้าสำเร็จรูป"},
		{"118", "118001", "งานระหว่างทำ", "118", "งานระหว่างทำ"},
		{"119", "119001", "วัตถุดิบ", "119", "วัตถุดิบ"},
		{"600", "600111", "ปรับปรุงสินค้า", "600", "ปรับปรุงสินค้า"},
	}

	// อ่าน ac_code ที่มีอยู่แล้ว — เก็บทั้ง exact และ prefix 3 หลัก
	existingPrefix := map[string]bool{} // key = prefix 3 หลัก
	rows, _ := f.GetRows("Ledger_Master")
	for i, r := range rows {
		if i == 0 || len(r) < 2 || safeGet(r, 0) != comCode {
			continue
		}
		ac := strings.TrimSpace(safeGet(r, 1))
		if len(ac) >= 3 {
			existingPrefix[ac[:3]] = true
		}
	}

	// หา next row
	nextRow := len(rows) + 1
	zeroCols := 29 // BBAL,CBAL,Debit,Credit,Bthisyear,Thisper01-12,Blastyear,Lastper01-12

	var added []string
	for _, ac := range required {
		if existingPrefix[ac.prefix] {
			continue // มี ac_code ที่ขึ้นต้นด้วย prefix นี้แล้ว ข้ามไป
		}
		// auto-add ด้วย acCode ที่มี suffix (ไม่ใช่ 3 หลักล้วน)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("A%d", nextRow), comCode)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("B%d", nextRow), ac.acCode)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("C%d", nextRow), ac.acName)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("D%d", nextRow), ac.gcode)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("E%d", nextRow), ac.gname)
		for j := 0; j < zeroCols; j++ {
			col, _ := excelize.ColumnNumberToName(6 + j)
			f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", col, nextRow), 0)
		}
		added = append(added, ac.acCode+"-"+ac.acName)
		nextRow++
	}
	return added
}

func buildInventoryAcctOptions(xlOpts excelize.Options) []string {
	fallback := []string{
		"117-สินค้าสำเร็จรูป",
		"118-งานระหว่างทำ",
		"119-วัตถุดิบ",
	}

	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err != nil {
		return fallback
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOpts)
	rows, _ := f.GetRows("Ledger_Master")

	// เก็บ found แยกตาม group 117/118/119
	type acEntry struct{ code, name string }
	groups := map[string][]acEntry{
		"117": {},
		"118": {},
		"119": {},
	}

	for i, r := range rows {
		if i == 0 || len(r) < 3 || safeGet(r, 0) != comCode {
			continue
		}
		ac := strings.TrimSpace(safeGet(r, 1))
		name := strings.TrimSpace(safeGet(r, 2))
		if len(ac) < 3 {
			continue
		}
		prefix := ac[:3]
		if _, ok := groups[prefix]; ok {
			groups[prefix] = append(groups[prefix], acEntry{ac, name})
		}
	}

	// สร้าง options: ถ้า group นั้นมีใน Ledger ให้แสดงทุก ac_code
	// ถ้าไม่มีเลย ให้ fallback เฉพาะ group นั้น
	var opts []string
	fallbackMap := map[string]string{
		"117": "117-สินค้าสำเร็จรูป",
		"118": "118-งานระหว่างทำ",
		"119": "119-วัตถุดิบ",
	}
	for _, prefix := range []string{"117", "118", "119"} {
		entries := groups[prefix]
		if len(entries) == 0 {
			// ไม่มีใน Ledger → fallback
			opts = append(opts, fallbackMap[prefix])
		} else {
			for _, e := range entries {
				opts = append(opts, e.code+"-"+e.name)
			}
		}
	}
	return opts
}

func getEndingInventoryGUI(w fyne.Window) fyne.CanvasObject {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	if err := ensureInventorySheet(xlOpts); err != nil {
		return widget.NewLabel("Error: " + err.Error())
	}

	// ── Period selector: แสดงแค่ Current Period (NowPeriod) เท่านั้น (แบบ A) ──
	// เหตุผล: Ending Inventory ทำได้แค่งวดปัจจุบัน ถ้าต้องการงวดอื่นต้องไป Select Period [F11] ก่อน
	cfg, _ := loadCompanyPeriod(xlOpts)
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	nowIdx := cfg.NowPeriod - 1
	if nowIdx < 0 {
		nowIdx = 0
	}
	if nowIdx >= len(periods) {
		nowIdx = len(periods) - 1
	}
	nowPeriod := periods[nowIdx]
	periodOptions := []string{
		fmt.Sprintf("Period %d (%s – %s)",
			cfg.NowPeriod,
			nowPeriod.PStart.Format("02/01/06"),
			nowPeriod.PEnd.Format("02/01/06")),
	}
	selPeriod := widget.NewSelect(periodOptions, nil)
	selPeriod.SetSelectedIndex(0)

	// ── Date entry — default = PEnd ของ current period ──
	enCountDate := newDateEntry()
	enCountDate.SetDate(nowPeriod.PEnd.Format("02/01/06"))

	// ── Form fields ──
	enProdCode := widget.NewEntry()
	enProdCode.SetPlaceHolder("รหัสสินค้า *")

	enProdName := widget.NewEntry()
	enProdName.SetPlaceHolder("ชื่อสินค้า *")

	var saveItem func()

	enQty := newInvNumEntry(func() {
		if saveItem != nil {
			saveItem()
		}
	})
	enQty.SetPlaceHolder("0.00")

	enUnitCost := newInvNumEntry(func() {
		if saveItem != nil {
			saveItem()
		}
	})
	enUnitCost.SetPlaceHolder("0.00")

	// ── acctOptions: อ่านจาก Ledger_Master filter prefix 117/118/119 ──
	// ถ้าไม่มีใน Ledger ให้ fallback เป็น hardcode เพื่อป้องกัน list ว่าง
	acctOptions := buildInventoryAcctOptions(xlOpts)
	selAcct := widget.NewSelect(acctOptions, nil)
	if len(acctOptions) > 0 {
		selAcct.SetSelectedIndex(0)
	}

	// ── State ──
	var items []InventoryItem
	var totalValue float64
	var editRowIndex int = -1 // -1=add, >=1=edit (sheet row)

	lblTotal := widget.NewLabelWithStyle("รวมมูลค่าสินค้าคงเหลือทั้งสิ้น: 0.00 บาท", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})
	lblMode := widget.NewLabelWithStyle("โหมด: เพิ่มรายการใหม่", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})

	// ── helpers ──
	showInfoDlg := func(title, msg string) {
		var d dialog.Dialog
		btn := newEnterEscButton("OK (Enter/Esc)", func() { d.Hide() }, func() { d.Hide() })
		d = dialog.NewCustomWithoutButtons(title, container.NewVBox(
			widget.NewLabel(msg),
			container.NewCenter(btn),
		), w)
		d.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(btn) }) }()
	}

	showConfirmDlg := func(title, msg string, onYes func()) {
		var d dialog.Dialog
		btnYes := newEnterEscButton("ยืนยัน (Enter)", func() { d.Hide(); onYes() }, func() { d.Hide() })
		btnNo := widget.NewButton("ยกเลิก (Esc)", func() { d.Hide() })
		d = dialog.NewCustomWithoutButtons(title, container.NewVBox(
			widget.NewLabel(msg),
			container.NewHBox(layout.NewSpacer(), btnYes, btnNo, layout.NewSpacer()),
		), w)
		d.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(btnYes) }) }()
	}

	// ── loadItems ──
	var reloadItems func()
	reloadItems = func() {
		items = nil
		totalValue = 0
		if selPeriod.SelectedIndex() == -1 {
			return
		}
		target := fmt.Sprintf("%d", cfg.NowPeriod)
		f, err := excelize.OpenFile(currentDBPath, xlOpts)
		if err != nil {
			return
		}
		defer f.Close()
		rows, _ := f.GetRows("Ending_Inventory")
		for i, row := range rows {
			if i == 0 || len(row) < 9 {
				continue
			}
			if row[0] != currentCompanyCode || row[1] != target {
				continue
			}
			amt := parseFloat(row[7])
			items = append(items, InventoryItem{
				RowIndex:  i + 1,
				ProdCode:  row[3],
				ProdName:  row[4],
				Qty:       parseFloat(row[5]),
				UnitCost:  parseFloat(row[6]),
				TotalAmt:  amt,
				AcctCode:  row[8],
				CountDate: row[2],
			})
			totalValue += amt
		}
		lblTotal.SetText(fmt.Sprintf("รวมมูลค่าสินค้าคงเหลือทั้งสิ้น: %s บาท", bsNum(totalValue)))
	}

	// ── resetForm ──
	var resetForm func()
	resetForm = func() {
		editRowIndex = -1
		enProdCode.SetText("")
		enProdName.SetText("")
		enQty.SetText("")
		enUnitCost.SetText("")
		selAcct.SetSelectedIndex(0)
		lblMode.SetText("โหมด: เพิ่มรายการใหม่")
		w.Canvas().Focus(enProdCode)
	}

	// ── validation ──
	validateForm := func() string {
		if strings.TrimSpace(enProdCode.Text) == "" {
			return "กรุณากรอกรหัสสินค้า"
		}
		if strings.TrimSpace(enProdName.Text) == "" {
			return "กรุณากรอกชื่อสินค้า"
		}
		if strings.TrimSpace(enQty.Text) == "" || parseFloat(enQty.Text) <= 0 {
			return "กรุณากรอกจำนวน (ต้องมากกว่า 0)"
		}
		if strings.TrimSpace(enUnitCost.Text) == "" || parseFloat(enUnitCost.Text) <= 0 {
			return "กรุณากรอกต้นทุน/หน่วย (ต้องมากกว่า 0)"
		}
		if enCountDate.GetDate() == "" {
			return "กรุณากรอกวันที่ตรวจนับ (DD/MM/YY)"
		}
		// validate date ต้องอยู่ใน range ของ current period
		if err := validateVoucherDate(enCountDate.GetDate(), nowPeriod.PStart, nowPeriod.PEnd); err != nil {
			return fmt.Sprintf("วันที่ตรวจนับไม่อยู่ใน Period %d\n(%s)",
				cfg.NowPeriod, err.Error())
		}
		return ""
	}

	// ── saveItem ──
	saveItem = func() {
		if msg := validateForm(); msg != "" {
			showInfoDlg("แจ้งเตือน", msg)
			return
		}
		qty := parseFloat(enQty.Text)
		cost := parseFloat(enUnitCost.Text)
		amt := qty * cost
		acctCode := strings.Split(selAcct.Selected, "-")[0]
		periodNo := fmt.Sprintf("%d", cfg.NowPeriod)
		countDate := enCountDate.GetDate()

		f, err := excelize.OpenFile(currentDBPath, xlOpts)
		if err != nil {
			showInfoDlg("Error", err.Error())
			return
		}
		defer f.Close()

		// ── ตรวจสอบและ auto-add 117/118/119/600 เข้า Ledger_Master ──
		if added := ensureInventoryLedgerCodes(xlOpts, f); len(added) > 0 {
			if err := f.Save(); err == nil {
				showInfoDlg("เพิ่ม Ledger อัตโนมัติ",
					"ระบบเพิ่ม account code ต่อไปนี้เข้า Ledger_Master อัตโนมัติ:\n"+
						strings.Join(added, "\n"))
			}
		}

		// ── resolve acctCode จริงจาก Ledger_Master ──
		// ถ้า dropdown เลือก fallback "117-xxx" → acctCode = "117" (3 หลัก)
		// ให้ค้นหา ac_code จริงที่ขึ้นต้นด้วย prefix นั้นใน Ledger_Master แทน
		if len(acctCode) == 3 {
			lmRows, _ := f.GetRows("Ledger_Master")
			comC := getComCodeFromExcel(xlOpts)
			for i, lr := range lmRows {
				if i == 0 || len(lr) < 2 || safeGet(lr, 0) != comC {
					continue
				}
				ac := strings.TrimSpace(safeGet(lr, 1))
				if len(ac) >= 3 && ac[:3] == acctCode {
					acctCode = ac // ใช้ ac_code จริง
					break
				}
			}
		}

		row := &[]interface{}{
			currentCompanyCode, periodNo, countDate,
			strings.TrimSpace(enProdCode.Text),
			strings.TrimSpace(enProdName.Text),
			qty, cost, amt, acctCode,
		}

		if editRowIndex == -1 {
			existingRows, _ := f.GetRows("Ending_Inventory")
			f.SetSheetRow("Ending_Inventory", fmt.Sprintf("A%d", len(existingRows)+1), row)
		} else {
			f.SetSheetRow("Ending_Inventory", fmt.Sprintf("A%d", editRowIndex), row)
		}

		if err := f.Save(); err != nil {
			showInfoDlg("Error", err.Error())
			return
		}
		resetForm()
		reloadItems()
	}

	// ── openSearch: F3 เปิด Popup ──
	openSearch := func() {
		periodNo := fmt.Sprintf("%d", cfg.NowPeriod)
		showInventorySearch(w, xlOpts, periodNo, func(it InventoryItem) {
			// load ข้อมูลเข้า form → Edit mode
			editRowIndex = it.RowIndex
			enProdCode.SetText(it.ProdCode)
			enProdName.SetText(it.ProdName)
			enQty.SetText(fmt.Sprintf("%.2f", it.Qty))
			enUnitCost.SetText(fmt.Sprintf("%.2f", it.UnitCost))
			enCountDate.SetDate(it.CountDate)
			for i, opt := range acctOptions {
				if strings.HasPrefix(opt, it.AcctCode) {
					selAcct.SetSelectedIndex(i)
					break
				}
			}
			lblMode.SetText(fmt.Sprintf("โหมด: แก้ไข — [%s] %s", it.ProdCode, it.ProdName))
			w.Canvas().Focus(enProdName)
		})
	}

	// ── register globals ──
	inventorySubmitFunc = func() { saveItem() }
	inventorySearchFunc = openSearch

	// ── ปุ่ม Delete ──
	btnDelete := widget.NewButton("ลบรายการที่เลือก [F3→เลือก→ลบ]", func() {
		if editRowIndex == -1 {
			showInfoDlg("แจ้งเตือน", "กรุณากด F3 เพื่อเลือกรายการที่ต้องการลบก่อน")
			return
		}
		showConfirmDlg("ยืนยันการลบ",
			fmt.Sprintf("ลบรายการ: [%s] %s\nออกจากระบบ?", enProdCode.Text, enProdName.Text),
			func() {
				f, err := excelize.OpenFile(currentDBPath, xlOpts)
				if err != nil {
					showInfoDlg("Error", err.Error())
					return
				}
				defer f.Close()
				rows, _ := f.GetRows("Ending_Inventory")
				f.DeleteSheet("Ending_Inventory")
				f.NewSheet("Ending_Inventory")
				f.SetSheetRow("Ending_Inventory", "A1", &[]interface{}{
					"ComCode", "Period", "CountDate", "ProdCode", "ProdName", "Qty", "UnitCost", "TotalAmt", "AcctCode",
				})
				newRow := 2
				for i, row := range rows {
					if i == 0 || len(row) < 9 || i+1 == editRowIndex {
						continue
					}
					vals := make([]interface{}, len(row))
					for j, v := range row {
						vals[j] = v
					}
					f.SetSheetRow("Ending_Inventory", fmt.Sprintf("A%d", newRow), &vals)
					newRow++
				}
				if err := f.Save(); err != nil {
					showInfoDlg("Error", err.Error())
					return
				}
				resetForm()
				reloadItems()
			})
	})
	btnDelete.Importance = widget.DangerImportance

	btnCancel := widget.NewButton("ยกเลิกแก้ไข (Esc)", func() { resetForm() })

	// ── Period change ──
	selPeriod.OnChanged = func(s string) {
		resetForm()
		reloadItems()
	}
	reloadItems()

	// ── doAutoJV ──
	doAutoJV := func(mode string) {
		if totalValue <= 0 {
			showInfoDlg("แจ้งเตือน", "ไม่มียอดสินค้าคงเหลือให้ปรับปรุง")
			return
		}
		// ── validate วันที่ตรวจนับต้องอยู่ใน current period ──
		if err := validateVoucherDate(enCountDate.GetDate(), nowPeriod.PStart, nowPeriod.PEnd); err != nil {
			showInfoDlg("วันที่ตรวจนับไม่ถูกต้อง",
				fmt.Sprintf("วันที่ตรวจนับต้องอยู่ใน Period %d\n%s\n\nกรุณาแก้ไขวันที่ก่อนกด Auto-JV",
					cfg.NowPeriod, err.Error()))
			return
		}
		f, err := excelize.OpenFile(currentDBPath, xlOpts)
		if err != nil {
			showInfoDlg("Error", err.Error())
			return
		}
		defer f.Close()

		// ══════════════════════════════════════════════════════════════
		// ตรวจสอบ INV voucher ของ period นี้ใน Book_items
		//   case 1: ไม่มี → สร้างใหม่
		//   case 2: มีแต่ NOT posted → UPDATE (ลบ lines เก่า + เขียนใหม่)
		//   case 3: มีและ POSTED → block
		// ══════════════════════════════════════════════════════════════
		periodNo := cfg.NowPeriod
		ref := fmt.Sprintf("ปรับปรุงสินค้าคงเหลือปลายงวด %d", periodNo)
		bdate := enCountDate.GetDate()

		type invVoucherInfo struct {
			bitem    string
			bvoucher string
			posted   bool
			rowNos   []int // row index ใน Book_items (1-based) ที่เป็น line ของ voucher นี้
		}

		allRows, _ := f.GetRows("Book_items")

		// ค้นหา INV voucher ที่มี ref ตรงกับ period นี้
		var existing *invVoucherInfo
		{
			seen := map[string]*invVoucherInfo{}
			for i, row := range allRows {
				if i == 0 || len(row) < 4 || row[0] != currentCompanyCode {
					continue
				}
				vch := strings.TrimSpace(safeGet(row, 2))
				if !strings.HasPrefix(vch, "INV-") {
					continue
				}
				// ตรวจ ref ตรงกับ period นี้ (col 11 = Bref)
				rowRef := strings.TrimSpace(safeGet(row, 11))
				if rowRef != ref {
					continue
				}
				bi := strings.TrimSpace(safeGet(row, 3))
				if _, ok := seen[vch]; !ok {
					isPosted := strings.TrimSpace(safeGet(row, 19)) == "1"
					seen[vch] = &invVoucherInfo{bitem: bi, bvoucher: vch, posted: isPosted}
				}
				seen[vch].rowNos = append(seen[vch].rowNos, i+1)
			}
			for _, v := range seen {
				existing = v
				break // มีได้แค่ 1 voucher ต่อ period
			}
		}

		var bitem, bvoucher string
		var newRowNo int
		bline := 1

		if existing == nil {
			// ── Case 1: ไม่มี INV voucher ใน period นี้ → สร้างใหม่ ──
			maxItem := 0
			for i, row := range allRows {
				if i == 0 || len(row) < 4 || row[0] != currentCompanyCode {
					continue
				}
				// ✅ เพิ่มการเช็ค Period ตรงนี้
				var rowPeriod int
				if len(row) > 21 {
					fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
				}
				if rowPeriod == periodNo {
					n := 0
					fmt.Sscanf(safeGet(row, 3), "%d", &n)
					if n > maxItem {
						maxItem = n
					}
				}
			}
			bitem = fmt.Sprintf("%03d", maxItem+1)
			bvoucher = "INV-" + bitem
			newRowNo = len(allRows) + 1

		} else if existing.posted {
			// ── Case 3: POSTED แล้ว → block + แนะนำขั้นตอน Unpost ──
			var d3 dialog.Dialog
			btn3 := newEnterEscButton("รับทราบ (Enter/Esc)", func() { d3.Hide() }, func() { d3.Hide() })
			msg3 := fmt.Sprintf(
				"Voucher %s (Period %d) ถูก Post ลง Ledger แล้ว\n"+
					"ไม่สามารถแก้ไขโดยตรงได้\n\n"+
					"วิธีแก้ไข (4 ขั้นตอน):\n"+
					"  1. ไปหน้า Book [F9]\n"+
					"  2. เปิด Voucher %s\n"+
					"  3. กดปุ่ม Edit (✏️)\n"+
					"     → ระบบจะ Unpost อัตโนมัติ\n"+
					"  4. กด Esc เพื่อออกจาก Edit\n"+
					"     → Voucher กลับเป็น 'ยังไม่ได้ Post'\n\n"+
					"จากนั้นกลับมากด Auto-JV ใหม่\n"+
					"ระบบจะอัพเดทยอดให้อัตโนมัติ",
				existing.bvoucher, periodNo, existing.bvoucher,
			)
			lbl3 := widget.NewLabel(msg3)
			lbl3.Wrapping = fyne.TextWrapWord
			d3 = dialog.NewCustomWithoutButtons(
				"⚠️  Voucher ถูก Post แล้ว — ต้อง Unpost ก่อน",
				container.NewVBox(lbl3, container.NewCenter(btn3)),
				w,
			)
			d3.Show()
			go func() { fyne.Do(func() { w.Canvas().Focus(btn3) }) }()
			return

		} else {
			// ── Case 2: มีแต่ NOT posted → UPDATE ──
			// ลบ lines เก่าทั้งหมดของ voucher นี้ก่อน แล้ว rewrite
			delSet := map[int]bool{}
			for _, rn := range existing.rowNos {
				delSet[rn] = true
			}
			// สร้าง sheet ใหม่โดยข้าม rows ที่ต้องลบ
			f.DeleteSheet("Book_items")
			f.NewSheet("Book_items")
			nextWrite := 2
			for i, row := range allRows {
				if i == 0 {
					// เขียน header
					vals := make([]interface{}, len(row))
					for j, v := range row {
						vals[j] = v
					}
					f.SetSheetRow("Book_items", "A1", &vals)
					continue
				}
				if delSet[i+1] {
					continue // ข้าม line เก่าของ INV voucher นี้
				}
				vals := make([]interface{}, len(row))
				for j, v := range row {
					vals[j] = v
				}
				f.SetSheetRow("Book_items", fmt.Sprintf("A%d", nextWrite), &vals)
				nextWrite++
			}
			bitem = existing.bitem
			bvoucher = existing.bvoucher
			newRowNo = nextWrite
		}

		// acctNames: อ่านจาก Ledger_Master (prefix 3 หลัก) + fallback
		acctNames := map[string]string{
			"117": "สินค้าสำเร็จรูป",
			"118": "งานระหว่างทำ",
			"119": "วัตถุดิบ",
		}
		{
			lmRows, _ := f.GetRows("Ledger_Master")
			comC := getComCodeFromExcel(xlOpts)
			for i, lr := range lmRows {
				if i == 0 || len(lr) < 3 || safeGet(lr, 0) != comC {
					continue
				}
				ac := strings.TrimSpace(safeGet(lr, 1))
				if len(ac) < 3 {
					continue
				}
				pfx := ac[:3]
				// ดึงชื่อจาก ac_code แรกที่ prefix ตรง (รองรับทั้ง "117", "117001", "117001" ฯลฯ)
				if _, ok := acctNames[pfx]; ok {
					acctNames[pfx] = strings.TrimSpace(safeGet(lr, 2))
				}
			}
		}

		if mode == "summary" {
			// group ตาม prefix 3 หลัก รองรับทั้ง "117", "117001", "117001" ฯลฯ
			// acctTotals[prefix] = ยอดรวม, acctCode[prefix] = ac_code จริงตัวแรกที่พบ
			acctTotals := map[string]float64{}
			acctCode := map[string]string{} // prefix → ac_code จริง (ใช้ post ไป Book)
			for _, it := range items {
				pfx := it.AcctCode
				if len(pfx) > 3 {
					pfx = pfx[:3]
				}
				acctTotals[pfx] += it.TotalAmt
				if _, seen := acctCode[pfx]; !seen {
					acctCode[pfx] = it.AcctCode // เก็บ ac_code จริงตัวแรก
				}
			}
			for _, pfx := range []string{"117", "118", "119"} {
				amt, ok := acctTotals[pfx]
				if !ok || amt == 0 {
					continue
				}
				// ใช้ ac_code จริงจาก items (ไม่ใช่ prefix 3 หลักล้วน)
				realAc := acctCode[pfx]
				f.SetSheetRow("Book_items", fmt.Sprintf("A%d", newRowNo), &[]interface{}{
					currentCompanyCode, bdate, bvoucher, bitem, bline,
					realAc, acctNames[pfx], "รวท", "สมุดรายวันทั่วไป", amt, 0.0,
					ref, "", "", ref, "", "", 0, 0, 0, "", periodNo,
				})
				newRowNo++
				bline++
			}
		} else {
			for _, it := range items {
				f.SetSheetRow("Book_items", fmt.Sprintf("A%d", newRowNo), &[]interface{}{
					currentCompanyCode, bdate, bvoucher, bitem, bline,
					it.AcctCode, it.ProdName, "รวท", "สมุดรายวันทั่วไป", it.TotalAmt, 0.0,
					ref, "", "", ref, "", "", 0, 0, 0, "", periodNo,
				})
				newRowNo++
				bline++
			}
		}
		// หา ac_code จริงสำหรับ 600 จาก Ledger_Master (prefix "600")
		ac600 := "600111" // fallback
		{
			lmRows, _ := f.GetRows("Ledger_Master")
			comC := getComCodeFromExcel(xlOpts)
			for i, lr := range lmRows {
				if i == 0 || len(lr) < 2 || safeGet(lr, 0) != comC {
					continue
				}
				ac := strings.TrimSpace(safeGet(lr, 1))
				if len(ac) >= 3 && ac[:3] == "600" {
					ac600 = ac
					break
				}
			}
		}
		f.SetSheetRow("Book_items", fmt.Sprintf("A%d", newRowNo), &[]interface{}{
			currentCompanyCode, bdate, bvoucher, bitem, bline,
			ac600, "ปรับปรุงสินค้า", "รวท", "สมุดรายวันทั่วไป", 0.0, totalValue,
			ref, "", "", ref, "", "", 0, 0, 0, "", periodNo,
		})

		if err := f.Save(); err != nil {
			showInfoDlg("Error", err.Error())
			return
		}

		gotoBitem := bitem
		modeLabel := "ยอดรวมทั้งก้อน"
		if mode == "detail" {
			modeLabel = "แยกรายการละเอียด"
		}
		var sd dialog.Dialog
		btnYes := newEnterEscButton("ไปหน้า Book (Enter)", func() {
			sd.Hide()
			if changePage != nil {
				changePage(bookPage)
			}
			if bookGotoFunc != nil {
				// ❌ โค้ดเดิม:
				// go func() { fyne.Do(func() { bookGotoFunc(gotoBitem) }) }()

				// ✅ เปลี่ยนเป็น: ส่ง periodNo (งวดปัจจุบัน) ไปด้วย
				go func() { fyne.Do(func() { bookGotoFunc(gotoBitem, periodNo) }) }()
			}
		}, func() { sd.Hide() })
		btnNo := widget.NewButton("ปิด (Esc)", func() { sd.Hide() })
		sd = dialog.NewCustomWithoutButtons("Auto-JV สำเร็จ", container.NewVBox(
			widget.NewLabel(fmt.Sprintf("Auto-JV สำเร็จ (%s)\nBitem: %s | Voucher: %s\nมูลค่ารวม: %s บาท\n\nไปหน้า Book เพื่อตรวจสอบและ Post?",
				modeLabel, bitem, bvoucher, bsNum(totalValue))),
			container.NewHBox(layout.NewSpacer(), btnYes, btnNo, layout.NewSpacer()),
		), w)
		sd.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(btnYes) }) }()
	}

	// validateBeforeJV — ตรวจก่อน confirm ทั้ง 2 ปุ่ม
	validateBeforeJV := func() bool {
		if totalValue <= 0 {
			showInfoDlg("แจ้งเตือน", "ไม่มียอดสินค้าคงเหลือให้ปรับปรุง")
			return false
		}
		if err := validateVoucherDate(enCountDate.GetDate(), nowPeriod.PStart, nowPeriod.PEnd); err != nil {
			showInfoDlg("วันที่ตรวจนับไม่ถูกต้อง",
				fmt.Sprintf("วันที่ตรวจนับต้องอยู่ใน Period %d\n%s\n\nกรุณาแก้ไขวันที่ก่อนกด Auto-JV",
					cfg.NowPeriod, err.Error()))
			return false
		}
		return true
	}

	btnJV1 := widget.NewButton("Auto-JV แบบ 1 (รวมก้อน)", func() {
		if !validateBeforeJV() {
			return
		}
		showConfirmDlg("ยืนยัน Auto-JV แบบที่ 1",
			fmt.Sprintf("บันทึกแบบยอดรวม (Summary)\nDr.117/118/119 รวมตาม GL / Cr.600\nมูลค่ารวม: %s บาท\n\nยืนยัน?", bsNum(totalValue)),
			func() { doAutoJV("summary") })
	})
	btnJV1.Importance = widget.MediumImportance

	btnJV2 := widget.NewButton("Auto-JV แบบ 2 (แยกรายการ)", func() {
		if !validateBeforeJV() {
			return
		}
		showConfirmDlg("ยืนยัน Auto-JV แบบที่ 2",
			fmt.Sprintf("บันทึกแบบแยกรายการ (Detail)\nDr. แต่ละสินค้า / Cr.600\nรวม %d รายการ มูลค่า: %s บาท\n\nยืนยัน?", len(items), bsNum(totalValue)),
			func() { doAutoJV("detail") })
	})
	btnJV2.Importance = widget.HighImportance

	// ── Form ──
	itemForm := &widget.Form{
		Items: []*widget.FormItem{
			{Text: "รหัสสินค้า *", Widget: enProdCode},
			{Text: "ชื่อสินค้า *", Widget: enProdName},
			{Text: "จำนวน (Qty) *", Widget: enQty},
			{Text: "ต้นทุน/หน่วย *", Widget: enUnitCost},
			{Text: "จัดเข้าบัญชี (GL)", Widget: selAcct},
		},
		OnSubmit: func() { saveItem() },
	}

	headerForm := &widget.Form{
		Items: []*widget.FormItem{
			{Text: "เลือก Period ที่ตรวจนับ", Widget: selPeriod},
			{Text: "วันที่ตรวจนับจริง", Widget: enCountDate},
		},
	}

	// ── Layout ──
	leftAligned := func(content fyne.CanvasObject) fyne.CanvasObject {
		return container.NewGridWithColumns(2, content, layout.NewSpacer())
	}

	return container.NewVBox(
		widget.NewLabelWithStyle("บันทึกรายละเอียดสินค้าคงเหลือปลายงวด (Ending Inventory)", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		leftAligned(headerForm),
		widget.NewSeparator(),
		widget.NewLabelWithStyle("เพิ่มรายการสินค้าที่ตรวจนับได้  [F3=ค้นหา/แก้ไข]", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		lblMode,
		leftAligned(itemForm),
		leftAligned(container.NewHBox(btnCancel, btnDelete)),
		widget.NewSeparator(),
		leftAligned(lblTotal),
		widget.NewSeparator(),
		widget.NewLabelWithStyle("สร้างใบสำคัญปรับปรุงสต๊อก (Auto-JV)", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		leftAligned(container.NewVBox(btnJV1, widget.NewLabel(""), btnJV2)),
		widget.NewLabel("แบบที่ 1: Dr.117/118/119 รวมตาม GL | แบบที่ 2: Dr. แยกทีละสินค้า"),
	)
}
