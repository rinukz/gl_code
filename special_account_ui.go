package main

import (
	"fmt"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

func getSpecialAccountGUI(w fyne.Window) fyne.CanvasObject {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	// ดึงค่า ComCode มาเตรียมไว้ (แต่เราจะเช็คให้ยืดหยุ่นขึ้น)
	targetComCode := getComCodeFromExcel(xlOptions)

	var listData [][]string
	var filteredData [][]string
	var saveAction func()

	// 1. โหลดข้อมูล (เปลี่ยนชื่อ Sheet ให้ตรงกับ Verify)
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err == nil {
		sheetName := "Special_code"
		rows, _ := f.GetRows(sheetName)
		for i, row := range rows {
			if i == 0 {
				continue
			} // ข้าม Header

			// ปรับให้ดึงข้อมูลขึ้นมาก่อน โดยเช็คแค่ว่ามีคอลัมน์ CODE และ NAME
			if len(row) >= 2 {
				// ถ้ามี ComCode (row[0]) ให้เช็ค ถ้าไม่มีหรือไม่ตรง ก็ยังดึงขึ้นมาให้เห็นก่อน
				// หรือถ้าป๋าอยากล็อคเฉพาะบริษัท ก็ใส่เงื่อนไข row[0] == targetComCode กลับไปได้ครับ
				listData = append(listData, []string{row[1], row[2]})
			}
		}
		f.Close()
	}
	filteredData = listData

	// 2. UI Elements
	enCode := widget.NewEntry()
	enCode.Disable() // 🔒 ล็อคตายตัว ห้ามพิมพ์เพิ่มเอง

	enName := newSmartEntry(func() { saveAction() })
	enName.SetPlaceHolder("แก้ไขชื่อบัญชีที่นี่...")

	searchEntry := widget.NewEntry()
	searchEntry.SetPlaceHolder("🔍 ค้นหา...")

	// 3. List (ไม่มีปุ่มลบ)
	theList := widget.NewList(
		func() int { return len(filteredData) },
		func() fyne.CanvasObject {
			return container.NewGridWithColumns(2,
				widget.NewLabel("CODE"),
				widget.NewLabel("NAME"),
			)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			if id >= len(filteredData) {
				return
			}
			row := filteredData[id]
			grid := o.(*fyne.Container)
			grid.Objects[0].(*widget.Label).SetText(row[0])
			grid.Objects[1].(*widget.Label).SetText(row[1])
		},
	)

	// คลิกเพื่อเลือกมาแก้ชื่อเท่านั้น
	theList.OnSelected = func(id widget.ListItemID) {
		enCode.SetText(filteredData[id][0])
		enName.SetText(filteredData[id][1])
		theList.Unselect(id)
	}

	// 4. บันทึก (Update เฉพาะที่มีอยู่แล้วเท่านั้น)
	saveAction = func() {
		code, name := enCode.Text, enName.Text
		if code == "" || name == "" {
			dialog.ShowError(fmt.Errorf("กรุณาเลือกรายการจากตารางก่อนแก้ไข"), w)
			return
		}

		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()

		sheetName := "Special_code"
		rows, _ := f.GetRows(sheetName)

		found := false
		for i, row := range rows {
			if len(row) >= 2 && row[0] == targetComCode && row[1] == code {
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", i+1), name)
				found = true
				break
			}
		}

		if !found {
			dialog.ShowError(fmt.Errorf("ไม่พบรหัสนี้ในระบบ (ห้ามเพิ่มรหัสใหม่)"), w)
			return
		}

		// ── Sync NAME → Ledger_Master.Ac_name ──
		syncSpecialCodeName(f, targetComCode, code, name, "special→ledger")

		if err := f.Save(); err == nil {
			// อัพเดท listData ใน memory แทนการ SetContent ใหม่ (ป้องกัน navbar หาย)
			for i, row := range listData {
				if row[0] == code {
					listData[i][1] = name
					break
				}
			}
			filteredData = listData
			theList.Refresh()

			var d dialog.Dialog
			okBtn := newEnterButton("OK", func() {
				d.Hide()
			})
			d = dialog.NewCustomWithoutButtons("สำเร็จ", container.NewVBox(
				widget.NewLabel("แก้ไขชื่อบัญชีสำเร็จ"),
				container.NewCenter(okBtn),
			), w)
			d.Show()
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
		}
	}

	// 5. Live Filter & Layout
	searchEntry.OnChanged = func(s string) {
		s = strings.ToLower(s)
		filteredData = nil
		for _, row := range listData {
			if s == "" || strings.Contains(strings.ToLower(row[0]), s) || strings.Contains(strings.ToLower(row[1]), s) {
				filteredData = append(filteredData, row)
			}
		}
		theList.Refresh()
	}

	btnSave := widget.NewButtonWithIcon("", theme.DocumentSaveIcon(), func() { saveAction() })
	// ทุกหน้าที่มี AddShortcut ให้เพิ่มบรรทัดนี้ก่อนเสมอ
	ctrlS := &desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: fyne.KeyModifierControl}
	w.Canvas().RemoveShortcut(ctrlS)
	w.Canvas().AddShortcut(ctrlS, func(shortcut fyne.Shortcut) { saveAction() })

	// w.Canvas().AddShortcut(&desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: fyne.KeyModifierControl}, func(shortcut fyne.Shortcut) { saveAction() })

	topSection := container.NewVBox(
		container.NewHBox(btnSave, layout.NewSpacer()),
		widget.NewLabelWithStyle("Special Account Code", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewGridWithColumns(2,
			container.NewVBox(widget.NewLabel("CODE"), enCode),
			container.NewVBox(widget.NewLabel("NAME"), enName),
		),
		widget.NewSeparator(),
		container.NewPadded(searchEntry),
		container.NewGridWithColumns(2,
			widget.NewLabelWithStyle("  CODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
			widget.NewLabelWithStyle("NAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		),
	)

	return container.NewBorder(topSection, nil, nil, nil, theList)
}
