package main

import (
	"fmt"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

// numericEntry: Live format ระหว่างพิมพ์ ไม่ cursor กระโดด
type numericEntry struct {
	widget.Entry
	isDecimal bool
	printer   *message.Printer
	updating  bool
	onSave    func()
}

func newNumericEntry(isDecimal bool, onSave func()) *numericEntry {
	e := &numericEntry{
		isDecimal: isDecimal,
		printer:   message.NewPrinter(language.English),
		onSave:    onSave,
	}
	e.ExtendBaseWidget(e)

	e.OnChanged = func(s string) {
		if e.updating {
			return
		}

		raw := ""
		hasDot := false
		for _, r := range s {
			if r >= '0' && r <= '9' {
				raw += string(r)
			} else if r == '.' && isDecimal && !hasDot {
				raw += string(r)
				hasDot = true
			}
		}

		if raw == "" {
			return
		}

		intPart := raw
		decPart := ""
		if idx := strings.Index(raw, "."); idx >= 0 {
			intPart = raw[:idx]
			decPart = raw[idx:]
		}

		formatted := ""
		if intPart == "" {
			formatted = "0"
		} else {
			val, err := strconv.ParseInt(intPart, 10, 64)
			if err == nil {
				formatted = e.printer.Sprintf("%d", val)
			} else {
				formatted = intPart
			}
		}
		formatted += decPart

		if formatted == s {
			return
		}

		oldCursor := e.CursorColumn
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

		e.updating = true
		e.SetText(formatted)
		if newCursor > len([]rune(formatted)) {
			newCursor = len([]rune(formatted))
		}
		e.CursorColumn = newCursor
		e.Refresh()
		e.updating = false
	}

	return e
}

func (e *numericEntry) TypedShortcut(s fyne.Shortcut) {
	cs, ok := s.(*desktop.CustomShortcut)
	if ok && cs.Modifier == fyne.KeyModifierControl && cs.KeyName == fyne.KeyS {
		if e.onSave != nil {
			e.onSave()
		}
		return
	}
	e.Entry.TypedShortcut(s)
}

func (e *numericEntry) FocusLost() {
	e.Entry.FocusLost()
	raw := strings.ReplaceAll(e.Text, ",", "")
	if raw == "" {
		return
	}
	val, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return
	}
	var formatted string
	if e.isDecimal {
		formatted = e.printer.Sprintf("%.2f", val)
	} else {
		formatted = e.printer.Sprintf("%.0f", val)
	}
	e.updating = true
	e.SetText(formatted)
	e.updating = false
}

func (e *numericEntry) getRawValue() string {
	return strings.ReplaceAll(e.Text, ",", "")
}

func (e *numericEntry) setFormattedText(s string) {
	raw := strings.ReplaceAll(s, ",", "")
	if raw == "" {
		e.SetText("")
		return
	}
	val, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		e.SetText(s)
		return
	}
	if e.isDecimal {
		e.SetText(e.printer.Sprintf("%.2f", val))
	} else {
		e.SetText(e.printer.Sprintf("%.0f", val))
	}
}

// ---------------------------------------------------------

func getCapitalGUI(w fyne.Window) fyne.CanvasObject {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	sheetName := "Capital"
	targetComCode := getComCodeFromExcel(xlOptions)

	var saveAction func()

	enThisQty := newNumericEntry(false, func() { saveAction() })
	enThisValue := newNumericEntry(true, func() { saveAction() })
	enLastQty := newNumericEntry(false, func() { saveAction() })
	enLastValue := newNumericEntry(true, func() { saveAction() })

	loadData := func() {
		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()
		rows, _ := f.GetRows(sheetName)
		// schema แนวนอน: A=Comcode B=ThisYearQty C=ThisYearValue D=LastYearQty E=LastYearValue
		for i, row := range rows {
			if i == 0 || len(row) < 5 || row[0] != targetComCode {
				continue
			}
			enThisQty.setFormattedText(row[1])
			enThisValue.setFormattedText(row[2])
			enLastQty.setFormattedText(row[3])
			enLastValue.setFormattedText(row[4])
			break
		}
	}
	loadData()

	saveAction = func() {
		// อ่านค่าก่อน focus nil
		thisQty := enThisQty.getRawValue()
		thisValue := enThisValue.getRawValue()
		lastQty := enLastQty.getRawValue()
		lastValue := enLastValue.getRawValue()

		w.Canvas().Focus(nil)

		f, err := excelize.OpenFile(currentDBPath, xlOptions)
		if err != nil {
			return
		}
		defer f.Close()

		rows, _ := f.GetRows(sheetName)

		// schema แนวนอน: A=Comcode B=ThisYearQty C=ThisYearValue D=LastYearQty E=LastYearValue
		found := false
		for i, row := range rows {
			if i == 0 || len(row) < 1 || row[0] != targetComCode {
				continue
			}
			rowNum := i + 1
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), thisQty)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), thisValue)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), lastQty)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), lastValue)
			found = true
			break
		}
		if !found {
			// ไม่พบ row ของ comcode นี้ → เพิ่มใหม่
			nextRow := len(rows) + 1
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", nextRow), targetComCode)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", nextRow), thisQty)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", nextRow), thisValue)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", nextRow), lastQty)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", nextRow), lastValue)
		}

		if err := f.Save(); err == nil {
			var d dialog.Dialog
			okBtn := newEnterButton("OK", func() {
				d.Hide()
				w.Canvas().Focus(nil)
			})
			d = dialog.NewCustomWithoutButtons("สำเร็จ", container.NewVBox(
				widget.NewLabel("บันทึกเรียบร้อย"),
				container.NewCenter(okBtn),
			), w)
			d.Show()
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
			loadData()
		} else {
			var d dialog.Dialog
			okBtn := newEnterButton("OK", func() {
				d.Hide()
				w.Canvas().Focus(nil)
			})
			d = dialog.NewCustomWithoutButtons("ผิดพลาด", container.NewVBox(
				widget.NewLabel(fmt.Sprintf("บันทึกไม่สำเร็จ: %v", err)),
				container.NewCenter(okBtn),
			), w)
			d.Show()
			fyne.Do(func() { w.Canvas().Focus(okBtn) })
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

	mainContent := container.NewVBox(
		widget.NewLabelWithStyle("Capital Setup - "+targetComCode, fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		widget.NewLabelWithStyle("📅 ปีปัจจุบัน (This Year)", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.New(layout.NewFormLayout(),
			widget.NewLabel("จำนวนหุ้น:"), enThisQty,
			widget.NewLabel("มูลค่าต่อหุ้น:"), enThisValue,
		),
		widget.NewSeparator(),
		widget.NewLabelWithStyle("⏳ ปีที่แล้ว (Last Year)", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.New(layout.NewFormLayout(),
			widget.NewLabel("จำนวนหุ้น:"), enLastQty,
			widget.NewLabel("มูลค่าต่อหุ้น:"), enLastValue,
		),
		container.NewPadded(
			container.NewHBox(
				layout.NewSpacer(),
				widget.NewButtonWithIcon("บันทึกข้อมูล", theme.DocumentSaveIcon(), saveAction),
				layout.NewSpacer(),
			),
		),
	)

	return container.NewPadded(mainContent)
}
