package main

import (
	"fmt"
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ---------company_setup_ui.go------------------------------------------------

func getCompanySetupGUI(w fyne.Window) fyne.CanvasObject {
	xlOptions := excelize.Options{Password: "@A123456789a"}

	var saveAction func()

	enComCode := newSmartEntry(func() { saveAction() })
	enComName := newSmartEntry(func() { saveAction() })
	enComAddr := newSmartMultiEntry(func() { saveAction() })
	enComTaxID := newSmartEntry(func() { saveAction() })
	enComPeriod := newSmartEntry(func() { saveAction() })
	enComNPeriod := newSmartEntry(func() { saveAction() })
	enComYearEnd := newDateEntry()
	enBackupPath := newSmartEntry(func() { saveAction() })
	enReportPath := newSmartEntry(func() { saveAction() })
	enVoidPassword := newSmartEntry(func() { saveAction() })
	enVoidPassword.Password = true

	enComAddr.nextFocusRef = enComTaxID

	// โหลดข้อมูลเดิม
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err == nil {
		sheet := "Company_Profile"
		vCode, _ := f.GetCellValue(sheet, "A2")
		if vCode != "" {
			enComCode.SetText(vCode)
			enComCode.Disable()
			currentCompanyCode = vCode
		}
		vName, _ := f.GetCellValue(sheet, "B2")
		enComName.SetText(vName)
		vAddr, _ := f.GetCellValue(sheet, "C2")
		enComAddr.SetText(vAddr)
		vTax, _ := f.GetCellValue(sheet, "D2")
		enComTaxID.SetText(vTax)
		vYEnd, _ := f.GetCellValue(sheet, "E2")
		enComYearEnd.SetDate(vYEnd)
		vPer, _ := f.GetCellValue(sheet, "F2")
		enComPeriod.SetText(vPer)
		vNPer, _ := f.GetCellValue(sheet, "G2")
		enComNPeriod.SetText(vNPer)
		vBackup, _ := f.GetCellValue(sheet, "H2")
		enBackupPath.SetText(vBackup)
		vReport, _ := f.GetCellValue(sheet, "I2")
		enReportPath.SetText(vReport)
		vVoidPwd, _ := f.GetCellValue(sheet, "J2")
		enVoidPassword.SetText(vVoidPwd)
		f.Close()
	}

	// ── once-only fields: ถ้ามีค่าแล้ว disable ──
	if enComTaxID.Text != "" {
		enComTaxID.Disable()
	}
	if enComYearEnd.GetDate() != "" {
		enComYearEnd.Disable()
	}
	if enComPeriod.Text != "" {
		enComPeriod.Disable()
	}
	if enComNPeriod.Text != "" {
		enComNPeriod.Disable()
	}

	// ฟังก์ชันบันทึก
	saveAction = func() {
		if enComCode.Text == "" {
			dialog.ShowError(fmt.Errorf("กรุณาระบุรหัสบริษัทก่อน"), w)
			return
		}

		saveFile, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			dialog.ShowError(fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err), w)
			return
		}
		defer saveFile.Close()

		s := "Company_Profile"
		saveFile.SetCellValue(s, "A2", enComCode.Text)
		saveFile.SetCellValue(s, "B2", enComName.Text)
		saveFile.SetCellValue(s, "C2", enComAddr.Text)
		saveFile.SetCellValue(s, "D2", enComTaxID.Text)
		saveFile.SetCellValue(s, "E2", enComYearEnd.GetDate())

		p, _ := strconv.Atoi(enComPeriod.Text)
		np, _ := strconv.Atoi(enComNPeriod.Text)
		saveFile.SetCellValue(s, "F2", p)
		saveFile.SetCellValue(s, "G2", np)
		saveFile.SetCellValue(s, "H2", enBackupPath.Text)
		saveFile.SetCellValue(s, "I2", enReportPath.Text)
		saveFile.SetCellValue(s, "J2", enVoidPassword.Text)

		if err := saveFile.Save(); err == nil {
			currentCompanyCode = enComCode.Text
			enComCode.Disable()
			// once-only: disable หลัง save ครั้งแรก
			enComTaxID.Disable()
			enComYearEnd.Disable()
			enComPeriod.Disable()
			enComNPeriod.Disable()

			// sync book_ui ให้ re-read period config ใหม่ทันที (กัน case ตั้งค่าครั้งแรก)
			if resetBook != nil {
				go func() { fyne.Do(func() { resetBook() }) }()
			}

			// ✅ ใช้ enterButton แทน widget.Button ปกติ
			var d dialog.Dialog

			okBtn := newEnterButton("OK", func() {
				d.Hide()
				w.Canvas().Focus(nil)
			})

			content := container.NewVBox(
				widget.NewLabel("บันทึกข้อมูลเรียบร้อยแล้ว"),
				container.NewCenter(okBtn),
			)

			d = dialog.NewCustomWithoutButtons("สำเร็จ", content, w)
			d.Show()

			fyne.Do(func() {
				w.Canvas().Focus(okBtn)
			})

		} else {
			dialog.ShowError(fmt.Errorf("บันทึกไม่สำเร็จ: %v", err), w)
		}
	}

	// ── Validation: TaxID — ตัวเลขเท่านั้น + จำกัด 13 หลัก ──
	enComTaxID.OnChanged = func(s string) {
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
			enComTaxID.SetText(digits)
		}
	}

	// ── enComYearEnd ใช้ dateEntry — masked dd/mm/yy (จัดการ filter ใน TypedRune แล้ว) ──

	// ── Validation: จำนวน Period — ตัวเลขเท่านั้น ──
	enComPeriod.OnChanged = func(s string) {
		digits := ""
		for _, r := range s {
			if r >= '0' && r <= '9' {
				digits += string(r)
			}
		}
		if digits != s {
			enComPeriod.SetText(digits)
		}
	}

	// ── Validation: Period ปัจจุบัน — ตัวเลขเท่านั้น ──
	enComNPeriod.OnChanged = func(s string) {
		digits := ""
		for _, r := range s {
			if r >= '0' && r <= '9' {
				digits += string(r)
			}
		}
		if digits != s {
			enComNPeriod.SetText(digits)
		}
	}

	ctrlS := &desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: fyne.KeyModifierControl}
	w.Canvas().RemoveShortcut(ctrlS)
	w.Canvas().AddShortcut(ctrlS, func(shortcut fyne.Shortcut) { saveAction() })

	// สร้างปุ่ม Browse สำหรับ Backup Path
	btnBrowseBackup := widget.NewButton("Browse...", func() {
		dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if uri == nil {
				return
			}
			enBackupPath.SetText(uri.Path())
		}, w)
	})

	// สร้างปุ่ม Browse สำหรับ Report Path
	btnBrowseReport := widget.NewButton("Browse...", func() {
		dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if uri == nil {
				return
			}
			enReportPath.SetText(uri.Path())
		}, w)
	})

	form := &widget.Form{
		Items: []*widget.FormItem{
			{Text: "รหัสบริษัท", Widget: enComCode},
			{Text: "ชื่อบริษัท", Widget: enComName},
			{Text: "ที่อยู่", Widget: enComAddr},
			// 👇 เอา container.NewBorder ออก ใส่ enComTaxID เข้าไปตรงๆ เลย 👇
			{Text: "เลขประจำตัวผู้เสียภาษี", Widget: enComTaxID},
			{Text: "วันปิดรอบงบ (dd/mm/yy)", Widget: enComYearEnd},
			{Text: "จำนวน Period", Widget: enComPeriod},
			{Text: "Period ปัจจุบัน", Widget: enComNPeriod},
			{Text: "Database Backup Path", Widget: container.NewBorder(nil, nil, nil, btnBrowseBackup, enBackupPath)},
			{Text: "Report Save Path", Widget: container.NewBorder(nil, nil, nil, btnBrowseReport, enReportPath)},
			{Text: "Password", Widget: enVoidPassword},
		},
		OnSubmit: saveAction,
	}

	// สร้าง Grid 2 คอลัมน์ (แบ่งครึ่งหน้าจอ 50/50)
	// คอลัมน์ 1 (ซ้าย) ใส่ฟอร์ม
	// คอลัมน์ 2 (ขวา) ปล่อยว่างไว้
	leftAlignedForm := container.NewGridWithColumns(2,
		form,                // ฟอร์มอยู่ฝั่งซ้าย (กว้าง 50%)
		container.NewVBox(), // พื้นที่ว่างฝั่งขวา
	)

	// refreshCompanySetupFunc — เรียกหลัง Close Period เพื่อ reload enComNPeriod
	refreshCompanySetupFunc = func() {
		rf, rerr := excelize.OpenFile(currentDBPath, xlOptions)
		if rerr == nil {
			vNPer, _ := rf.GetCellValue("Company_Profile", "G2")
			rf.Close()
			fyne.Do(func() {
				enComNPeriod.Enable()
				enComNPeriod.SetText(vNPer)
				enComNPeriod.Disable()
			})
		}
	}

	return container.NewVBox(
		widget.NewLabelWithStyle("ตั้งค่าข้อมูลบริษัท", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		leftAlignedForm, // ใช้ Layout ที่จัดชิดซ้ายแล้ว
	)
}
