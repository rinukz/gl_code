package main

import (
	"fmt"
	"strings" // ✅ เพิ่มบรรทัดนี้

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// func getComCodeFromExcel(xlOptions excelize.Options) string {
// 	f, err := excelize.OpenFile(currentDBPath, xlOptions)
// 	if err != nil {
// 		return ""
// 	}
// 	defer f.Close()
// 	val, _ := f.GetCellValue("Company_Profile", "A2")
// 	return val
// }

func getSubsidiaryBooksGUI(w fyne.Window) fyne.CanvasObject {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	targetComCode := getComCodeFromExcel(xlOptions)

	// ข้อมูลทั้งหมด (ไม่เปลี่ยน)
	var listData [][]string
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err == nil {
		rows, _ := f.GetRows("Subsidiary_Books")
		for i, row := range rows {
			if i == 0 {
				continue
			}
			if len(row) >= 3 && row[0] == targetComCode {
				listData = append(listData, []string{row[1], row[2]})
			}
		}
		f.Close()
	}

	// ✅ filteredData คือที่ List จะอ่าน (เริ่มต้น = ทั้งหมด)
	filteredData := make([][]string, len(listData))
	copy(filteredData, listData)

	var saveAction func()

	enCode := newSmartEntry(func() { saveAction() })
	enCode.SetPlaceHolder("CODE")
	enName := newSmartEntry(func() { saveAction() })
	enName.SetPlaceHolder("NAME")

	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("ค้นหา...")

	// List อ่านจาก filteredData
	var theList *widget.List
	theList = widget.NewList(
		func() int { return len(filteredData) },
		func() fyne.CanvasObject {
			return container.NewBorder(nil, nil,
				widget.NewLabel("CODE_TMP"),
				widget.NewButtonWithIcon("", theme.DeleteIcon(), nil),
				widget.NewLabel("NAME_TMP"),
			)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			row := filteredData[id] // ✅ เปลี่ยนจาก listData → filteredData
			ctr := o.(*fyne.Container)
			ctr.Objects[1].(*widget.Label).SetText(row[0])
			ctr.Objects[0].(*widget.Label).SetText(row[1])

			btnDel := ctr.Objects[2].(*widget.Button)
			btnDel.OnTapped = func() {
				var confirmDialog dialog.Dialog

				yesBtn := newEnterButton("ลบ", func() {
					confirmDialog.Hide()
					deleteSubBook(row[0], xlOptions)
					// reload listData จาก Excel แทน SetContent (ป้องกัน navbar หาย)
					listData = nil
					if ff, e2 := excelize.OpenFile(currentDBPath, xlOptions); e2 == nil {
						if allRows, _ := ff.GetRows("Subsidiary_Books"); len(allRows) > 1 {
							for _, r := range allRows[1:] {
								if len(r) >= 3 && r[0] == targetComCode {
									listData = append(listData, []string{r[1], r[2]})
								}
							}
						}
						ff.Close()
					}
					filteredData = make([][]string, len(listData))
					copy(filteredData, listData)
					theList.Refresh()
				})
				noBtn := newEnterButton("ยกเลิก", func() {
					confirmDialog.Hide()
				})
				yesBtn.Importance = widget.DangerImportance

				content := container.NewVBox(
					widget.NewLabel("ยืนยันการลบ: "+row[0]+"?"),
					container.NewCenter(container.NewHBox(yesBtn, noBtn)),
				)

				confirmDialog = dialog.NewCustomWithoutButtons("ยืนยัน", content, w)
				confirmDialog.Show()
				fyne.Do(func() { w.Canvas().Focus(noBtn) })
			}
		},
	)

	// ✅ Live Filter
	searchEntry.OnChanged = func(keyword string) {
		keyword = strings.ToLower(keyword)
		filteredData = nil
		for _, row := range listData {
			if keyword == "" ||
				strings.Contains(strings.ToLower(row[0]), keyword) ||
				strings.Contains(strings.ToLower(row[1]), keyword) {
				filteredData = append(filteredData, row)
			}
		}
		theList.Refresh()
	}

	saveAction = func() {
		code := strings.TrimSpace(enCode.Text)
		name := strings.TrimSpace(enName.Text)
		// ถ้าตรงกับ placeholder → ถือว่าว่าง
		if code == "CODE" {
			code = ""
		}
		if name == "NAME" {
			name = ""
		}

		if code == "" || name == "" || targetComCode == "" {
			dialog.ShowError(fmt.Errorf("กรุณากรอก CODE และ NAME"), w)
			return
		}

		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()

		sheetName := "Subsidiary_Books"
		index, _ := f.GetSheetIndex(sheetName)
		if index == -1 {
			f.NewSheet(sheetName)
			f.SetCellValue(sheetName, "A1", "ComCode")
			f.SetCellValue(sheetName, "B1", "CODE")
			f.SetCellValue(sheetName, "C1", "NAME")
		}

		rows, _ := f.GetRows(sheetName)
		targetRow := -1
		isExactMatch := false

		for i, row := range rows {
			if len(row) < 2 || row[0] != targetComCode {
				continue
			}
			existingCode := row[1]
			existingName := ""
			if len(row) >= 3 {
				existingName = row[2]
			}
			if strings.EqualFold(existingCode, code) {
				targetRow = i + 1
				if strings.EqualFold(existingName, name) {
					isExactMatch = true
				}
			} else {
				if strings.EqualFold(existingName, name) {
					dialog.ShowError(fmt.Errorf("ชื่อ '%s' มีอยู่แล้ว (CODE: %s)", name, existingCode), w)
					return
				}
			}
		}

		if isExactMatch {
			dialog.ShowInformation("แจ้งเตือน", "ข้อมูลนี้มีอยู่แล้วในระบบ ไม่มีการเปลี่ยนแปลง", w)
			return
		}

		if targetRow == -1 {
			targetRow = len(rows) + 1
		}

		f.SetCellValue(sheetName, fmt.Sprintf("A%d", targetRow), targetComCode)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", targetRow), code)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", targetRow), name)

		if err := f.Save(); err == nil {
			// โหลด listData ใหม่จาก Excel เลย (แน่ใจว่าข้อมูลตรง)
			listData = nil
			if ff, e2 := excelize.OpenFile(currentDBPath, xlOptions); e2 == nil {
				if allRows, _ := ff.GetRows(sheetName); len(allRows) > 1 {
					for _, r := range allRows[1:] {
						if len(r) >= 3 && r[0] == targetComCode {
							listData = append(listData, []string{r[1], r[2]})
						}
					}
				}
				ff.Close()
			}
			filteredData = make([][]string, len(listData))
			copy(filteredData, listData)
			theList.Refresh()
			enCode.SetText("")
			enName.SetText("")

			var d dialog.Dialog
			okBtn := newEnterButton("OK", func() {
				d.Hide()
			})
			content := container.NewVBox(
				widget.NewLabel("บันทึกในนามบริษัท: "+targetComCode),
				container.NewCenter(okBtn),
			)
			d = dialog.NewCustomWithoutButtons("สำเร็จ", content, w)
			d.Show()
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
		} else {
			dialog.ShowError(fmt.Errorf("บันทึกไม่สำเร็จ: %v", err), w)
		}
	}

	// ทุกหน้าที่มี AddShortcut ให้เพิ่มบรรทัดนี้ก่อนเสมอ
	ctrlS := &desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: fyne.KeyModifierControl}
	w.Canvas().RemoveShortcut(ctrlS)
	w.Canvas().AddShortcut(ctrlS, func(shortcut fyne.Shortcut) { saveAction() })
	// w.Canvas().AddShortcut(
	// 	&desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: fyne.KeyModifierControl},
	// 	func(shortcut fyne.Shortcut) { saveAction() },
	// )

	btnSave := widget.NewButtonWithIcon("", theme.DocumentSaveIcon(), func() { saveAction() })

	inputRow := container.NewGridWithColumns(2,
		container.NewVBox(widget.NewLabel("CODE"), enCode),
		container.NewVBox(widget.NewLabel("NAME"), enName),
	)

	topSection := container.NewVBox(
		container.NewHBox(btnSave, layout.NewSpacer()),
		widget.NewLabelWithStyle("Subsidiary Books Setup", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		inputRow,
		widget.NewSeparator(),
		container.NewPadded(searchEntry),
		container.NewHBox(widget.NewLabel("  CODE"), layout.NewSpacer(), widget.NewLabel("NAME"), layout.NewSpacer(), widget.NewLabel("ACTION  ")),
	)

	return container.NewBorder(topSection, nil, nil, nil, theList)
}

func deleteSubBook(code string, opts excelize.Options) {
	targetComCode := getComCodeFromExcel(opts)
	f, err := excelize.OpenFile(currentDBPath, opts)
	if err != nil {
		return
	}
	defer f.Close()

	sheet := "Subsidiary_Books"
	rows, _ := f.GetRows(sheet)
	for i, row := range rows {
		if len(row) >= 2 && row[0] == targetComCode && row[1] == code {
			f.RemoveRow(sheet, i+1)
			break
		}
	}
	f.Save()
}
