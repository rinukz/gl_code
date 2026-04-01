package main

import (
	"fmt"
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

type BookHeader struct {
	Bitem     string
	Bdate     string
	Bvoucher  string
	Bref      string
	Bnote     string
	SumDebit  float64
	SumCredit float64
	Bperiod   int // ✅ เพิ่ม Bperiod
}

// ─────────────────────────────────────────────────────────────
// showBookSearch — ค้นหาเอกสาร Book
// ─────────────────────────────────────────────────────────────
func showBookSearch(w fyne.Window, xlOptions excelize.Options, comCode string,
	onSelect func(bitem string, bperiod int)) {

	allHeaders := loadAllBookHeaders(xlOptions, comCode)
	filteredHeaders := make([]BookHeader, len(allHeaders))
	copy(filteredHeaders, allHeaders)

	selectedIdx := 0
	selectedPeriod := 0 // 0 = All, 1..N = period ที่เลือก

	// โหลด period range สำหรับ filter
	var periods []PeriodInfo
	if cfg, err := loadCompanyPeriod(xlOptions); err == nil {
		periods = calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	}

	var list *widget.List
	var pop *widget.PopUp

	// applyFilter — รวม keyword + period filter ไว้ที่เดียว
	var applyFilter func(keyword string)
	applyFilter = func(keyword string) {
		keyword = strings.ToLower(strings.TrimSpace(keyword))
		filteredHeaders = nil
		selectedIdx = 0
		for _, h := range allHeaders {
			// ── filter period ──
			if selectedPeriod != 0 && selectedPeriod <= len(periods) {
				p := periods[selectedPeriod-1]
				t, err := time.Parse("02/01/06", h.Bdate)
				if err != nil || t.Before(p.PStart) || t.After(p.PEnd) {
					continue
				}
			}
			// ── filter keyword ──
			if keyword != "" &&
				!strings.Contains(strings.ToLower(h.Bitem), keyword) &&
				!strings.Contains(strings.ToLower(h.Bvoucher), keyword) &&
				!strings.Contains(strings.ToLower(h.Bref), keyword) {
				continue
			}
			filteredHeaders = append(filteredHeaders, h)
		}
		list.UnselectAll()
		list.Refresh()
		if len(filteredHeaders) > 0 {
			list.Select(0)
			list.ScrollTo(0)
		}
	}

	doSelect := func(id int) {
		if id < 0 || id >= len(filteredHeaders) {
			return
		}
		selected := filteredHeaders[id]
		pop.Hide()
		onSelect(selected.Bitem, selected.Bperiod) // ✅ ส่ง bperiod กลับไปด้วย
	}

	highlightOnly := func(id int) {
		if id < 0 || id >= len(filteredHeaders) {
			return
		}
		selectedIdx = id
		list.Select(widget.ListItemID(selectedIdx))
		list.ScrollTo(widget.ListItemID(selectedIdx))
	}

	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("ค้นหา Bitem / Voucher / Reference... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")
	searchEntry.onEsc = func() { pop.Hide() }

	searchEntry.onDown = func() {
		if len(filteredHeaders) == 0 {
			return
		}
		next := selectedIdx + 1
		if next >= len(filteredHeaders) {
			next = len(filteredHeaders) - 1
		}
		highlightOnly(next)
	}

	searchEntry.onUp = func() {
		if len(filteredHeaders) == 0 {
			return
		}
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}

	searchEntry.onEnter = func() {
		doSelect(selectedIdx)
	}

	list = widget.NewList(
		func() int { return len(filteredHeaders) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			contentRow := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(60, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(80, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(150, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(200, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(80, 28), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(80, 28), widget.NewLabel("")),
			)
			return container.NewMax(btn, contentRow)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			h := filteredHeaders[id]
			maxCtr := o.(*fyne.Container)
			maxCtr.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
			row := maxCtr.Objects[1].(*fyne.Container)
			row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(h.Bitem)
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(h.Bdate)
			row.Objects[2].(*fyne.Container).Objects[0].(*widget.Label).SetText(h.Bvoucher)
			row.Objects[3].(*fyne.Container).Objects[0].(*widget.Label).SetText(h.Bref)
			row.Objects[4].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("%.2f", h.SumDebit))
			row.Objects[5].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("%.2f", h.SumCredit))
		},
	)

	searchEntry.OnChanged = func(keyword string) {
		applyFilter(keyword)
	}

	// ── Period dropdown: All, P01, P02, ... ──
	periodOpts := []string{"All"}
	for i, p := range periods {
		periodOpts = append(periodOpts,
			fmt.Sprintf("P%02d (%s)", i+1, p.PEnd.Format("02/01/06")))
	}
	periodSelect := widget.NewSelect(periodOpts, func(s string) {
		if s == "All" {
			selectedPeriod = 0
		} else {
			fmt.Sscanf(s, "P%02d", &selectedPeriod)
		}
		applyFilter(searchEntry.Text)
	})
	periodSelect.SetSelected("All")

	listHeader := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(60, 25), widget.NewLabelWithStyle("ITEM", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(80, 25), widget.NewLabelWithStyle("DATE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(150, 25), widget.NewLabelWithStyle("VOUCHER", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(200, 25), widget.NewLabelWithStyle("REFERENCE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(80, 25), widget.NewLabelWithStyle("DEBIT", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(80, 25), widget.NewLabelWithStyle("CREDIT", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
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
				widget.NewLabelWithStyle("ค้นหาเอกสาร", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				widget.NewLabel("Period"),
				container.NewGridWrap(fyne.NewSize(120, 28), periodSelect),
				closeBtn,
			),
			container.NewGridWrap(fyne.NewSize(680, 400), content),
		),
		w.Canvas(),
	)

	pop.Show()
	if len(filteredHeaders) > 0 {
		list.Select(0)
		list.ScrollTo(0)
	}

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

// ─────────────────────────────────────────────────────────────
// [DONE- Dont-Delete and Ask before edit]  showBookLedgerSearch — ค้นหา Account Code ใน Book
// ─────────────────────────────────────────────────────────────
func showBookLedgerSearch(w fyne.Window, xlOptions excelize.Options,
	onSelect func(acCode string, acName string)) {

	allRecords := loadAllLedgerRecords(xlOptions)
	filteredRecords := make([]LedgerRecord, len(allRecords))
	copy(filteredRecords, allRecords)

	selectedIdx := 0

	var list *widget.List
	var pop *widget.PopUp

	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("ค้นหา AcCode / AcName... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")

	searchEntry.onSave = func() {
		if pop != nil {
			pop.Hide()
		}
		closeAcSearchPopup = nil
		if bookSaveFunc != nil {
			bookSaveFunc()
		}
	}

	doSelect := func(id int) {
		if id < 0 || id >= len(filteredRecords) {
			return
		}
		r := filteredRecords[id]
		pop.Hide()
		closeAcSearchPopup = nil // ← เพิ่ม
		onSelect(r.AcCode, r.AcName)
	}

	// 🟢 1. สร้างฟังก์ชัน Quick Add สำหรับ Ledger
	quickAddLedger := func(kw string) {
		enNewCode := newSmartEntry(nil) // onSave จะ patch หลัง saveFunc ประกาศ
		enNewCode.SetText(strings.ToUpper(kw))
		enNewName := newSmartEntry(nil)

		var d dialog.Dialog
		saveFunc := func() {
			code := strings.TrimSpace(enNewCode.Text)
			name := strings.TrimSpace(enNewName.Text)
			if len(code) < 3 || name == "" {
				dialog.ShowInformation("Error", "ข้อมูลไม่ครบถ้วน หรือ Code สั้นกว่า 3 ตัวอักษร", w)
				return
			}
			gcode := code[:3]
			gname := lookupGname(xlOptions, gcode)
			if gname == "" {
				dialog.ShowInformation("Error", "ไม่พบ Account Group: "+gcode+"\nกรุณาตั้งค่าใน Setup ก่อน", w)
				return
			}
			_, found := loadLedgerRecord(xlOptions, code)
			if found {
				dialog.ShowInformation("Error", "A/C Code นี้มีอยู่แล้ว", w)
				return
			}

			// บันทึกลง Excel
			r := emptyLedger()
			r.Comcode = getComCodeFromExcel(xlOptions)
			r.AcCode = code
			r.AcName = name
			r.Gcode = gcode
			r.Gname = gname
			saveLedgerRecord(xlOptions, r, true)

			d.Hide()
			pop.Hide()           // ปิด popup ค้นหา
			onSelect(code, name) // ส่งค่ากลับไปที่ฟอร์มหลัก
		}

		cancelFunc := func() {
			d.Hide()
			w.Canvas().Focus(searchEntry)
		}
		enNewCode.onSave = saveFunc // Ctrl+S จาก enNewCode
		enNewName.onSave = saveFunc // Ctrl+S จาก enNewName
		enNewCode.onEnter = func() { w.Canvas().Focus(enNewName) }
		enNewCode.onEsc = cancelFunc
		enNewName.onEnter = saveFunc // Enter จาก enNewName → save
		enNewName.onEsc = cancelFunc
		enNewCode.OnSubmitted = func(s string) { w.Canvas().Focus(enNewName) }
		enNewName.OnSubmitted = func(s string) { saveFunc() }

		btnSave := newEnterButton("บันทึก", saveFunc)
		btnSave.Importance = widget.HighImportance
		btnCancel := newEnterButton("ยกเลิก", func() {
			d.Hide()
			w.Canvas().Focus(searchEntry)
		})

		form := widget.NewForm(
			widget.NewFormItem("A/C Code", enNewCode),
			widget.NewFormItem("A/C Name", enNewName),
		)

		content := container.NewVBox(
			form,
			container.NewCenter(container.NewHBox(btnSave, btnCancel)),
		)

		d = dialog.NewCustomWithoutButtons("เพิ่ม Account Code ใหม่", content, w)
		d.Resize(fyne.NewSize(300, 200))
		d.Show()
		w.Canvas().Focus(enNewCode)
	}

	// 🟢 2. สร้าง UI สำหรับกรณีค้นหาไม่พบ
	var notFoundLbl *widget.Label
	var btnAddNew *widget.Button
	var notFoundContainer *fyne.Container

	notFoundLbl = widget.NewLabelWithStyle("", fyne.TextAlignCenter, fyne.TextStyle{Italic: true})
	btnAddNew = widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() {
		quickAddLedger(searchEntry.Text)
	})
	btnAddNew.Importance = widget.HighImportance

	notFoundContainer = container.NewVBox(
		layout.NewSpacer(),
		notFoundLbl,
		container.NewCenter(btnAddNew),
		layout.NewSpacer(),
	)
	notFoundContainer.Hide() // ซ่อนไว้ก่อน

	highlightOnly := func(id int) {
		if id < 0 || id >= len(filteredRecords) {
			return
		}
		selectedIdx = id
		list.Select(widget.ListItemID(selectedIdx))
		list.ScrollTo(widget.ListItemID(selectedIdx))
	}

	searchEntry.onEsc = func() {
		pop.Hide()
		closeAcSearchPopup = nil // ← เพิ่ม
	}

	searchEntry.onDown = func() {
		if len(filteredRecords) == 0 {
			return
		}
		next := selectedIdx + 1
		if next >= len(filteredRecords) {
			next = len(filteredRecords) - 1
		}
		highlightOnly(next)
	}

	searchEntry.onUp = func() {
		if len(filteredRecords) == 0 {
			return
		}
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}

	searchEntry.onEnter = func() {
		if len(filteredRecords) > 0 {
			doSelect(selectedIdx)
		} else {
			quickAddLedger(searchEntry.Text) // กด Enter ตอนไม่เจอ = Add New
		}
	}

	list = widget.NewList(
		func() int { return len(filteredRecords) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			contentRow := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(250, 30), widget.NewLabel("")),
			)
			return container.NewMax(btn, contentRow)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			r := filteredRecords[id]
			maxCtr := o.(*fyne.Container)
			maxCtr.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
			row := maxCtr.Objects[1].(*fyne.Container)
			row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.AcCode)
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.AcName)
		},
	)

	// 🟢 3. อัปเดต OnChanged ให้สลับการแสดงผล
	searchEntry.OnChanged = func(kw string) {
		keyword := strings.ToLower(strings.TrimSpace(kw))
		filteredRecords = nil
		selectedIdx = 0
		for _, r := range allRecords {
			if keyword == "" ||
				strings.Contains(strings.ToLower(r.AcCode), keyword) ||
				strings.Contains(strings.ToLower(r.AcName), keyword) {
				filteredRecords = append(filteredRecords, r)
			}
		}
		list.UnselectAll()
		list.Refresh()

		if len(filteredRecords) > 0 {
			list.Show()
			notFoundContainer.Hide()
			list.Select(0)
			list.ScrollTo(0)
		} else {
			list.Hide()
			notFoundLbl.SetText(fmt.Sprintf("ไม่พบข้อมูล \"%s\"", kw))
			btnAddNew.SetText(fmt.Sprintf("+ เพิ่ม AC Code ใหม่: %s", kw))
			notFoundContainer.Show()
		}
	}

	header := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabelWithStyle("A/C CODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(250, 30), widget.NewLabelWithStyle("A/C NAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
	)

	closeBtn := widget.NewButton("✕ ปิด", func() { pop.Hide() })

	// 🟢 4. ใช้ container.NewStack ซ้อน list กับ notFoundContainer
	content := container.NewBorder(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหา Account Code", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			searchEntry, header,
		),
		nil, nil, nil,
		container.NewStack(list, notFoundContainer),
	)

	go func() {
		// จังหวะที่ 1: หน่วงให้ Mouse Up ทำงานเสร็จก่อน ป้องกัน PopUp ปิดตัวเอง
		time.Sleep(150 * time.Millisecond)
		fyne.Do(func() {
			const popW, popH float32 = 380, 300
			popSize := fyne.NewSize(popW, popH)

			pop = widget.NewPopUp(content, w.Canvas())
			pop.Resize(popSize)

			// anchor ใต้ field ที่ focus
			var popX, popY float32
			if focused := w.Canvas().Focused(); focused != nil {
				if obj, ok := focused.(fyne.CanvasObject); ok {
					pos := fyne.CurrentApp().Driver().AbsolutePositionForObject(obj)
					popX = pos.X
					popY = pos.Y + obj.Size().Height
				}
			}
			// fallback: มุมขวาล่าง
			if popX == 0 && popY == 0 {
				cs := w.Canvas().Size()
				popX = cs.Width - popW - 30
				popY = cs.Height - popH - 30
			}
			// clamp กันล้นขอบ
			cs := w.Canvas().Size()
			if popY+popH > cs.Height {
				popY = cs.Height - popH - 10
			}
			if popX+popW > cs.Width {
				popX = cs.Width - popW - 10
			}

			pop.Move(fyne.NewPos(popX, popY))
			pop.Show()

			// ✅ register global closer ให้ Ctrl+S ปิด popup นี้ก่อน save
			closeAcSearchPopup = func() {
				if pop != nil {
					pop.Hide()
				}
				closeAcSearchPopup = nil
			}

			if len(filteredRecords) > 0 {
				list.Select(0)
				list.ScrollTo(0)
			}
		})

		// จังหวะที่ 2: หน่วงให้ Fyne วาด UI เสร็จก่อน Focus
		time.Sleep(150 * time.Millisecond)
		fyne.Do(func() {
			if w.Canvas() != nil {
				w.Canvas().Focus(searchEntry)
			}
		})
	}()
}

// ─────────────────────────────────────────────────────────────
// showBookSubbookSearch — ค้นหา Subbook
// ─────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
// showBookSubbookSearch — ค้นหา Subbook
// ─────────────────────────────────────────────────────────────
// [DONE- Dont-Delete and Ask before edit] showBookSubbookSearch — anchor popup ใต้ field, clamp กันล้นขอบ, keyboard nav
func showBookSubbookSearch(w fyne.Window, xlOptions excelize.Options, comCode string,
	onSelect func(scode string, sname string)) {

	type SubItem struct {
		Scode string
		Sname string
	}

	allSubs := func() []SubItem {
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return nil
		}
		defer f.Close()
		rows, _ := f.GetRows("Subsidiary_Books")
		var result []SubItem
		for i, row := range rows {
			if i == 0 || len(row) < 3 {
				continue
			}
			if row[0] != comCode {
				continue
			}
			result = append(result, SubItem{Scode: row[1], Sname: row[2]})
		}
		return result
	}()

	filteredSubs := make([]SubItem, len(allSubs))
	copy(filteredSubs, allSubs)

	selectedIdx := 0

	var list *widget.List
	var pop *widget.PopUp
	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("ค้นหา Scode / Sname... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")

	closePop := func() {
		if pop != nil {
			pop.Hide()
		}
	}

	doSelect := func(id int) {
		if id < 0 || id >= len(filteredSubs) {
			return
		}
		s := filteredSubs[id]
		closePop()
		onSelect(s.Scode, s.Sname)
	}

	// 🟢 1. สร้างฟังก์ชัน Quick Add สำหรับ Subbook
	quickAddSubbook := func(kw string) {
		enNewCode := newSmartEntry(nil)
		enNewCode.SetText(strings.ToUpper(kw))
		enNewName := newSmartEntry(nil)

		var d dialog.Dialog
		saveFunc := func() {
			code := strings.TrimSpace(enNewCode.Text)
			name := strings.TrimSpace(enNewName.Text)
			if code == "" || name == "" {
				dialog.ShowInformation("Error", "ข้อมูลไม่ครบถ้วน", w)
				return
			}

			// บันทึกลง Excel ทันที
			f, err := excelize.OpenFile(currentDBPath, xlOptions)
			if err == nil {
				sheetName := "Subsidiary_Books"
				rows, _ := f.GetRows(sheetName)

				// เช็คซ้ำ
				dup := false
				for _, row := range rows {
					if len(row) >= 2 && row[0] == comCode && row[1] == code {
						dup = true
						break
					}
				}
				if dup {
					f.Close()
					dialog.ShowInformation("Error", "SCODE นี้มีอยู่แล้ว", w)
					return
				}

				targetRow := len(rows) + 1
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", targetRow), comCode)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", targetRow), code)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", targetRow), name)
				f.Save()
				f.Close()
			}

			d.Hide()
			pop.Hide()           // ปิด popup ค้นหา
			onSelect(code, name) // ส่งค่ากลับไปที่ฟอร์มหลัก
		}

		cancelFunc := func() {
			d.Hide()
			w.Canvas().Focus(searchEntry)
		}
		enNewCode.onSave = saveFunc // Ctrl+S จาก enNewCode
		enNewName.onSave = saveFunc // Ctrl+S จาก enNewName
		enNewCode.onEnter = func() { w.Canvas().Focus(enNewName) }
		enNewCode.onEsc = cancelFunc
		enNewName.onEnter = saveFunc // Enter จาก enNewName → save
		enNewName.onEsc = cancelFunc
		enNewCode.OnSubmitted = func(s string) { w.Canvas().Focus(enNewName) }
		enNewName.OnSubmitted = func(s string) { saveFunc() }

		btnSave := newEnterButton("บันทึก", saveFunc)
		btnSave.Importance = widget.HighImportance
		btnCancel := newEnterButton("ยกเลิก", func() {
			d.Hide()
			w.Canvas().Focus(searchEntry)
		})

		form := widget.NewForm(
			widget.NewFormItem("SCODE", enNewCode),
			widget.NewFormItem("SNAME", enNewName),
		)

		content := container.NewVBox(
			form,
			container.NewCenter(container.NewHBox(btnSave, btnCancel)),
		)

		d = dialog.NewCustomWithoutButtons("เพิ่ม Subbook ใหม่", content, w)
		d.Resize(fyne.NewSize(300, 200))
		d.Show()
		w.Canvas().Focus(enNewCode)
	}

	// 🟢 2. สร้าง UI สำหรับกรณีค้นหาไม่พบ
	var notFoundLbl *widget.Label
	var btnAddNew *widget.Button
	var notFoundContainer *fyne.Container

	notFoundLbl = widget.NewLabelWithStyle("", fyne.TextAlignCenter, fyne.TextStyle{Italic: true})
	btnAddNew = widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() {
		quickAddSubbook(searchEntry.Text)
	})
	btnAddNew.Importance = widget.HighImportance

	notFoundContainer = container.NewVBox(
		layout.NewSpacer(),
		notFoundLbl,
		container.NewCenter(btnAddNew),
		layout.NewSpacer(),
	)
	notFoundContainer.Hide()

	highlightOnly := func(id int) {
		if id < 0 || id >= len(filteredSubs) {
			return
		}
		selectedIdx = id
		list.Select(widget.ListItemID(selectedIdx))
		list.ScrollTo(widget.ListItemID(selectedIdx))
	}

	searchEntry.onEsc = closePop
	searchEntry.onEnter = func() {
		if len(filteredSubs) > 0 {
			doSelect(selectedIdx)
		} else {
			quickAddSubbook(searchEntry.Text)
		}
	}
	searchEntry.onDown = func() {
		if len(filteredSubs) == 0 {
			return
		}
		next := selectedIdx + 1
		if next >= len(filteredSubs) {
			next = len(filteredSubs) - 1
		}
		highlightOnly(next)
	}
	searchEntry.onUp = func() {
		if len(filteredSubs) == 0 {
			return
		}
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}

	list = widget.NewList(
		func() int { return len(filteredSubs) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			contentRow := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabel("")),
				container.NewGridWrap(fyne.NewSize(250, 30), widget.NewLabel("")),
			)
			return container.NewMax(btn, contentRow)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			s := filteredSubs[id]
			maxCtr := o.(*fyne.Container)
			maxCtr.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
			row := maxCtr.Objects[1].(*fyne.Container)
			row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(s.Scode)
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(s.Sname)
		},
	)

	// 🟢 3. อัปเดต OnChanged ให้สลับการแสดงผล
	searchEntry.OnChanged = func(kw string) {
		keyword := strings.ToLower(strings.TrimSpace(kw))
		filteredSubs = nil
		selectedIdx = 0
		for _, s := range allSubs {
			if keyword == "" ||
				strings.Contains(strings.ToLower(s.Scode), keyword) ||
				strings.Contains(strings.ToLower(s.Sname), keyword) {
				filteredSubs = append(filteredSubs, s)
			}
		}
		list.UnselectAll()
		list.Refresh()

		if len(filteredSubs) > 0 {
			list.Show()
			notFoundContainer.Hide()
			list.Select(0)
			list.ScrollTo(0)
		} else {
			list.Hide()
			notFoundLbl.SetText(fmt.Sprintf("ไม่พบข้อมูล \"%s\"", kw))
			btnAddNew.SetText(fmt.Sprintf("+ เพิ่ม SCODE ใหม่: %s", kw))
			notFoundContainer.Show()
		}
	}

	header := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabelWithStyle("SCODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(250, 30), widget.NewLabelWithStyle("SNAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
	)

	closeBtn := widget.NewButton("✕ ปิด", closePop)

	// 🟢 4. ใช้ container.NewStack ซ้อน list กับ notFoundContainer
	content := container.NewBorder(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหา Subbook", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			searchEntry, header,
		),
		nil, nil, nil,
		container.NewStack(list, notFoundContainer),
	)

	go func() {
		// จังหวะที่ 1: หน่วงให้ Mouse Up ทำงานเสร็จก่อน ป้องกัน PopUp ปิดตัวเอง
		time.Sleep(150 * time.Millisecond)
		fyne.Do(func() {
			const popW, popH float32 = 380, 300
			popSize := fyne.NewSize(popW, popH)

			pop = widget.NewPopUp(content, w.Canvas())
			pop.Resize(popSize)

			// anchor ใต้ field ที่ focus
			var popX, popY float32
			if focused := w.Canvas().Focused(); focused != nil {
				if obj, ok := focused.(fyne.CanvasObject); ok {
					pos := fyne.CurrentApp().Driver().AbsolutePositionForObject(obj)
					popX = pos.X
					popY = pos.Y + obj.Size().Height
				}
			}
			// fallback: มุมขวาล่าง
			if popX == 0 && popY == 0 {
				cs := w.Canvas().Size()
				popX = cs.Width - popW - 30
				popY = cs.Height - popH - 30
			}
			// clamp กันล้นขอบ
			cs := w.Canvas().Size()
			if popY+popH > cs.Height {
				popY = cs.Height - popH - 10
			}
			if popX+popW > cs.Width {
				popX = cs.Width - popW - 10
			}

			pop.Move(fyne.NewPos(popX, popY))
			pop.Show()

			if len(filteredSubs) > 0 {
				list.Select(0)
				list.ScrollTo(0)
			}
		})

		// จังหวะที่ 2: หน่วงให้ Fyne วาด UI เสร็จก่อน Focus
		time.Sleep(150 * time.Millisecond)
		fyne.Do(func() {
			if w.Canvas() != nil {
				w.Canvas().Focus(searchEntry)
			}
		})
	}()
}

// ─────────────────────────────────────────────────────────────
// loadAllBookHeaders
// ─────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
// loadAllBookHeaders
// ─────────────────────────────────────────────────────────────
func loadAllBookHeaders(xlOptions excelize.Options, comCode string) []BookHeader {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	rows, _ := f.GetRows("Book_items")

	type tempHeader struct {
		BookHeader
		seen bool
	}
	headerMap := make(map[string]*tempHeader)
	var order []string

	for i, row := range rows {
		if i == 0 || len(row) < 4 {
			continue
		}
		if row[0] != comCode {
			continue
		}

		bitem := row[3]
		var bperiod int
		if len(row) > 21 {
			fmt.Sscanf(safeGet(row, 21), "%d", &bperiod)
		}

		mapKey := fmt.Sprintf("%s_%d", bitem, bperiod)

		if _, exists := headerMap[mapKey]; !exists {
			headerMap[mapKey] = &tempHeader{
				BookHeader: BookHeader{
					Bitem:    bitem,
					Bdate:    safeGet(row, 1),
					Bvoucher: safeGet(row, 2),
					Bref:     safeGet(row, 11),
					Bnote:    safeGet(row, 14),
					Bperiod:  bperiod, // ✅ เก็บ Bperiod
				},
			}
			order = append(order, mapKey)
		}
		headerMap[mapKey].SumDebit += parseFloat(safeGet(row, 9))
		headerMap[mapKey].SumCredit += parseFloat(safeGet(row, 10))
	}

	var result []BookHeader
	for _, key := range order {
		result = append(result, headerMap[key].BookHeader)
	}

	// ✅ เพิ่มโค้ดส่วนนี้เพื่อทำการ Sort (เรียงลำดับ) ก่อนส่งไปแสดงผล
	sort.Slice(result, func(i, j int) bool {
		// 1. เรียงตาม Period ก่อน (จากน้อยไปมาก)
		if result[i].Bperiod != result[j].Bperiod {
			return result[i].Bperiod < result[j].Bperiod
		}

		// 2. ถ้า Period เท่ากัน ให้เรียงตาม Bitem (แปลงเป็นตัวเลขเพื่อป้องกันปัญหา 001, 002)
		ni, _ := strconv.Atoi(result[i].Bitem)
		nj, _ := strconv.Atoi(result[j].Bitem)
		if ni != nj {
			return ni < nj
		}

		// 3. ถ้า Bitem เท่ากัน (ไม่ควรเกิด) ให้เรียงตามวันที่
		return result[i].Bdate < result[j].Bdate
	})

	return result
}

// ─────────────────────────────────────────────────────────────────
// showBookHistorySearch — ค้นหา TaxID / Note / Note2 จาก Book_items ที่บันทึกไว้
// field: "taxid" | "note" | "note2"
// ─────────────────────────────────────────────────────────────────
// [DONE] showBookHistorySearch — search history by column, working correctly
func showBookHistorySearch(w fyne.Window, xlOptions excelize.Options, comCode string,
	field string, onSelect func(val string)) {

	loadValues := func() []string {
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return nil
		}
		defer f.Close()
		rows, _ := f.GetRows("Book_items")
		seen := map[string]bool{}
		var result []string
		colIdx := 13
		switch field {
		case "note":
			colIdx = 14
		case "note2":
			colIdx = 20
		}
		for i, row := range rows {
			if i == 0 {
				continue
			}
			if len(row) <= colIdx || row[0] != comCode {
				continue
			}
			v := strings.TrimSpace(row[colIdx])
			if v == "" || seen[v] {
				continue
			}
			seen[v] = true
			result = append(result, v)
		}
		return result
	}

	allVals := loadValues()
	filtered := make([]string, len(allVals))
	copy(filtered, allVals)

	selectedIdx := 0

	var list *widget.List
	var pop *widget.PopUp

	doSelect := func(id int) {
		if id < 0 || id >= len(filtered) {
			return
		}
		pop.Hide()
		onSelect(filtered[id])
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
	title := map[string]string{"taxid": "Tax ID", "note": "Note", "note2": "Note2"}[field]
	searchEntry.SetPlaceHolder("ค้นหา " + title + "... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")
	searchEntry.onEsc = func() { pop.Hide() }
	searchEntry.onEnter = func() { doSelect(selectedIdx) }

	searchEntry.onDown = func() {
		next := selectedIdx + 1
		if next >= len(filtered) {
			next = len(filtered) - 1
		}
		highlightOnly(next)
	}
	searchEntry.onUp = func() {
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}

	list = widget.NewList(
		func() int { return len(filtered) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			contentRow := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(440, 28), widget.NewLabel("")),
			)
			return container.NewMax(btn, contentRow)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			maxCtr := o.(*fyne.Container)
			maxCtr.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
			maxCtr.Objects[1].(*fyne.Container).Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(filtered[id])
		},
	)

	searchEntry.OnChanged = func(kw string) {
		kw = strings.ToLower(strings.TrimSpace(kw))
		filtered = nil
		for _, v := range allVals {
			if kw == "" || strings.Contains(strings.ToLower(v), kw) {
				filtered = append(filtered, v)
			}
		}
		list.UnselectAll()
		list.Refresh()
		if len(filtered) > 0 {
			highlightOnly(0)
		}
	}

	closeBtn := widget.NewButton("✕ ปิด", func() { pop.Hide() })
	content := container.NewBorder(
		searchEntry, nil, nil, nil,
		container.NewGridWrap(fyne.NewSize(460, 250), list),
	)
	pop = widget.NewModalPopUp(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหา "+title, fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			content,
		),
		w.Canvas(),
	)

	list.UnselectAll()
	pop.Show()
	if len(filtered) > 0 {
		highlightOnly(0)
	}
	go func() {
		time.Sleep(150 * time.Millisecond)
		fyne.Do(func() { w.Canvas().Focus(searchEntry) })
	}()
}

// ─────────────────────────────────────────────────────────────────────────────
// showTaxIDSearch — popup ค้นหา TaxID จาก Customer_Log
// ─────────────────────────────────────────────────────────────────────────────
// [DONE] showTaxIDSearch — search TaxID from Customer_Log, onFocusGained trigger, auto-fill Note, working correctly
func showTaxIDSearch(w fyne.Window, xlOptions excelize.Options, onSelect func(taxID, custName string), onClose func()) {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return
	}
	defer f.Close()

	rows, _ := f.GetRows("Customer_Log")
	type item struct{ taxID, custName string }
	var items []item
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 3 {
			items = append(items, item{row[1], row[2]})
		}
	}
	if len(items) == 0 {
		var d dialog.Dialog
		okBtn := newEnterButton("OK", func() { d.Hide() })
		d = dialog.NewCustomWithoutButtons("Customer_Log",
			container.NewVBox(
				widget.NewLabel("ยังไม่มีข้อมูล TaxID ที่บันทึกไว้"),
				container.NewCenter(okBtn),
			), w)
		d.Show()
		fyne.Do(func() { w.Canvas().Focus(okBtn) })
		return
	}

	var filtered []item
	copyItems := func() { filtered = make([]item, len(items)); copy(filtered, items) }
	copyItems()

	enSearch := newSmartEntry(nil)
	enSearch.SetPlaceHolder("พิมพ์ TaxID หรือ ชื่อบริษัท...")

	list := widget.NewList(
		func() int { return len(filtered) },
		func() fyne.CanvasObject {
			return container.NewHBox(
				container.NewGridWrap(fyne.NewSize(150, 28), widget.NewLabel("")),
				widget.NewLabel(""),
			)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			box := o.(*fyne.Container)
			box.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(filtered[id].taxID)
			box.Objects[1].(*widget.Label).SetText(filtered[id].custName)
		},
	)

	var pop *widget.PopUp
	closePop := func() {
		if pop != nil {
			pop.Hide()
			pop = nil
		}
	}

	// isNavigating — กัน OnSelected fire ตอนแค่กด ↑↓ เลื่อนดู
	isNavigating := false
	cursorIdx := -1

	doConfirm := func() {
		if cursorIdx >= 0 && cursorIdx < len(filtered) {
			closePop()
			onSelect(filtered[cursorIdx].taxID, filtered[cursorIdx].custName)
		}
	}

	list.OnSelected = func(id widget.ListItemID) {
		if isNavigating {
			cursorIdx = int(id)
			return // แค่ highlight ไม่ confirm
		}
		// คลิก mouse โดยตรง → confirm ทันที
		cursorIdx = int(id)
		doConfirm()
	}

	enSearch.OnChanged = func(kw string) {
		kwLow := strings.ToLower(strings.TrimSpace(kw))
		if kwLow == "" {
			copyItems()
		} else {
			filtered = nil
			for _, it := range items {
				if strings.Contains(strings.ToLower(it.taxID), kwLow) ||
					strings.Contains(strings.ToLower(it.custName), kwLow) {
					filtered = append(filtered, it)
				}
			}
		}
		cursorIdx = -1
		list.UnselectAll()
		list.Refresh()
	}

	// ↓ จาก search → เลื่อน highlight ลงไป (ไม่ confirm)
	enSearch.onDown = func() {
		if len(filtered) == 0 {
			return
		}
		isNavigating = true
		next := cursorIdx + 1
		if next >= len(filtered) {
			next = 0
		}
		cursorIdx = next
		list.Select(widget.ListItemID(cursorIdx))
		list.ScrollTo(widget.ListItemID(cursorIdx))
		isNavigating = false
	}

	// Enter จาก search → confirm รายการที่ highlight อยู่
	enSearch.onEnter = func() {
		if cursorIdx >= 0 {
			doConfirm()
		}
	}

	// ESC → ปิด popup และ focus กลับ
	enSearch.onEsc = func() {
		closePop()
		if onClose != nil {
			fyne.Do(onClose)
		}
	}

	closeBtn := widget.NewButton("✕ ปิด", closePop)
	content := container.NewBorder(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหา TaxID / ชื่อบริษัท", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			enSearch,
		),
		nil, nil, nil,
		container.NewGridWrap(fyne.NewSize(500, 300), list),
	)
	pop = widget.NewModalPopUp(content, w.Canvas())
	pop.Resize(fyne.NewSize(520, 400))
	pop.Show()
	w.Canvas().Focus(enSearch)
}
