package main

// select_period_ui.go
// Select Period — เทียบเท่า FrmSpecial ใน VB.NET
//
// CONCEPT:
//   NowPeriod (Company_Profile G2) = งวดบัญชีจริง ที่ Close Period มาแล้ว
//                                     → ห้ามแก้ที่นี่เด็ดขาด
//   workingPeriod (global, in-memory) = งวดที่กำลังทำงานใน session นี้
//                                       → Select Period แก้ได้ ไม่เขียน DB
//
// UI:
//   ● กลับสู่งวดบัญชีปัจจุบัน → workingPeriod = 0 (ใช้ NowPeriod จริง)
//   ○ เลือกงวด 1..NowPeriod   → workingPeriod = N
//   [OK] [Cancel]

import (
	"fmt"
	"image/color"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

func showSelectPeriodDialog(w fyne.Window) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}

	// Auto mode = กลับสู่ NowPeriod จริง (workingPeriod = 0)
	// Manual mode = เลือกงวด 1..NowPeriod
	useAuto := true
	selectedPeriod := cfg.NowPeriod // ค่า preview ใน dialog

	// แสดง working period ปัจจุบัน
	currentWorking := workingPeriod
	if currentWorking == 0 {
		currentWorking = cfg.NowPeriod
	}

	lblMode := canvas.NewText("● กลับสู่งวดบัญชีปัจจุบัน", color.RGBA{R: 0x00, G: 0x7A, B: 0xFF, A: 0xFF})
	lblMode.TextStyle = fyne.TextStyle{Bold: true}
	lblMode.TextSize = 13
	lblMode.Alignment = fyne.TextAlignCenter

	lblPeriod := canvas.NewText(
		fmt.Sprintf("จะเปลี่ยนไป: งวด %d (งวดบัญชีจริง)", cfg.NowPeriod),
		color.RGBA{R: 0x00, G: 0x7A, B: 0xFF, A: 0xFF},
	)
	lblPeriod.TextStyle = fyne.TextStyle{Bold: true}
	lblPeriod.TextSize = 13
	lblPeriod.Alignment = fyne.TextAlignCenter

	refreshLabels := func() {
		if useAuto {
			lblMode.Text = "● กลับสู่งวดบัญชีปัจจุบัน"
			lblMode.Color = color.RGBA{R: 0x00, G: 0x7A, B: 0xFF, A: 0xFF}
			lblPeriod.Text = fmt.Sprintf("จะเปลี่ยนไป: งวด %d (งวดบัญชีจริง)", cfg.NowPeriod)
			lblPeriod.Color = color.RGBA{R: 0x00, G: 0x7A, B: 0xFF, A: 0xFF}
		} else {
			lblMode.Text = fmt.Sprintf("● เลือกงวด: %d", selectedPeriod)
			lblMode.Color = color.RGBA{R: 0xC0, G: 0x20, B: 0x20, A: 0xFF}
			lblPeriod.Text = fmt.Sprintf("จะเปลี่ยนไป: งวด %d / %d", selectedPeriod, cfg.NowPeriod)
			lblPeriod.Color = color.RGBA{R: 0xC0, G: 0x20, B: 0x20, A: 0xFF}
		}
		lblMode.Refresh()
		lblPeriod.Refresh()
	}
	refreshLabels()

	var d dialog.Dialog

	doConfirm := func() {
		var target int
		if useAuto {
			// กลับสู่ NowPeriod จริง → reset workingPeriod (in-memory only)
			target = cfg.NowPeriod
			workingPeriod = 0
		} else {
			target = selectedPeriod
			workingPeriod = target // in-memory only — ไม่แตะ Company_Profile เลย
		}
		d.Hide()

		// แจ้ง book_ui ให้ reload period range ใหม่
		if setWorkingPeriodFunc != nil {
			setWorkingPeriodFunc(target)
		}

		msg := fmt.Sprintf("เปลี่ยนมาทำงานในงวด %d แล้ว", target)
		if useAuto {
			msg = fmt.Sprintf("กลับสู่งวดบัญชีปัจจุบัน (งวด %d) แล้ว", target)
		}
		var done dialog.Dialog
		ok2 := newEnterButton("OK", func() { done.Hide() })
		done = dialog.NewCustomWithoutButtons("สำเร็จ", container.NewVBox(
			widget.NewLabel(msg),
			container.NewCenter(ok2),
		), w)
		done.Show()
		go func() {
			time.Sleep(50 * time.Millisecond)
			fyne.Do(func() { w.Canvas().Focus(ok2) })
		}()
	}

	btn := &selectPeriodButton{
		cfg:            cfg,
		useAuto:        &useAuto,
		selectedPeriod: &selectedPeriod,
		refreshLabels:  refreshLabels,
		onEnter:        doConfirm,
		onEsc:          func() { d.Hide() },
	}
	btn.Text = "OK (Enter)"
	btn.Importance = widget.HighImportance
	btn.OnTapped = doConfirm
	btn.ExtendBaseWidget(btn)

	btnCancel := newEscButton("Cancel (Esc)", func() { d.Hide() })

	d = dialog.NewCustomWithoutButtons("เลือกงวดบัญชี", container.NewVBox(
		widget.NewSeparator(),
		widget.NewLabelWithStyle(
			fmt.Sprintf("งวดบัญชีจริง: %d / %d   |   กำลังทำงาน: งวด %d",
				cfg.NowPeriod, cfg.TotalPeriods, currentWorking),
			fyne.TextAlignCenter, fyne.TextStyle{},
		),
		widget.NewSeparator(),
		container.NewCenter(lblMode),
		container.NewCenter(lblPeriod),
		widget.NewLabelWithStyle(
			fmt.Sprintf("  ↓ = สลับโหมด/เพิ่มงวด (1-%d)   ↑ = ลดงวด/กลับ Auto  ", cfg.NowPeriod),
			fyne.TextAlignCenter, fyne.TextStyle{Italic: true},
		),
		widget.NewSeparator(),
		container.NewCenter(container.NewHBox(btn, btnCancel)),
	), w)
	d.Show()
	w.Canvas().Focus(btn)
}

// selectPeriodButton — button ที่รับ key ทั้งหมดสำหรับ Select Period dialog
type selectPeriodButton struct {
	widget.Button
	cfg            CompanyPeriodConfig
	useAuto        *bool
	selectedPeriod *int
	refreshLabels  func()
	onEnter        func()
	onEsc          func()
}

func (b *selectPeriodButton) TypedKey(key *fyne.KeyEvent) {
	switch key.Name {
	case fyne.KeyReturn, fyne.KeyEnter:
		if b.onEnter != nil {
			b.onEnter()
		}
	case fyne.KeyEscape:
		if b.onEsc != nil {
			b.onEsc()
		}
	case fyne.KeyDown:
		if *b.useAuto {
			// สลับจาก Auto → Manual เริ่มที่งวด 1
			*b.useAuto = false
			*b.selectedPeriod = 1
		} else {
			// จำกัดไม่เกิน NowPeriod (งวดที่เปิดแล้วเท่านั้น)
			if *b.selectedPeriod < b.cfg.NowPeriod {
				*b.selectedPeriod++
			}
		}
		b.refreshLabels()
	case fyne.KeyUp:
		if !*b.useAuto {
			if *b.selectedPeriod > 1 {
				*b.selectedPeriod--
			} else {
				// กลับไป Auto mode
				*b.useAuto = true
				*b.selectedPeriod = b.cfg.NowPeriod
			}
		}
		b.refreshLabels()
	}
}

// getMaxPeriodFromBook — หา period ล่าสุดที่มีข้อมูลใน Book_items (ยังคงไว้เผื่อใช้ที่อื่น)
func getMaxPeriodFromBook(xlOptions excelize.Options, cfg CompanyPeriodConfig) int {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return cfg.NowPeriod
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	rows, _ := f.GetRows("Book_items")
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)

	maxPeriod := 1
	for i, row := range rows {
		if i == 0 || len(row) < 2 || row[0] != comCode {
			continue
		}
		t, err := time.Parse("02/01/06", strings.TrimSpace(safeGet(row, 1)))
		if err != nil {
			continue
		}
		for _, p := range periods {
			if !t.Before(p.PStart) && !t.After(p.PEnd) {
				if p.PNo > maxPeriod {
					maxPeriod = p.PNo
				}
				break
			}
		}
	}
	return maxPeriod
}

// updateNowPeriod — set ComNPeriod ใน Company_Profile G2
// ใช้เฉพาะ Close Period เท่านั้น — Select Period ห้ามเรียก
func updateNowPeriod(xlOptions excelize.Options, newPeriod int) error {
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return err
	}
	f.SetCellValue("Company_Profile", "G2", newPeriod)
	if err := f.Save(); err != nil {
		f.Close()
		return err
	}
	return f.Close()
}

// showErrDialog — error dialog helper
func showErrDialog(w fyne.Window, msg string) {
	var d dialog.Dialog
	okBtn := newEnterButton("OK", func() { d.Hide() })
	d = dialog.NewCustomWithoutButtons("Error", container.NewVBox(
		widget.NewLabel(msg),
		container.NewCenter(okBtn),
	), w)
	d.Show()
	go func() {
		time.Sleep(50 * time.Millisecond)
		fyne.Do(func() { w.Canvas().Focus(okBtn) })
	}()
}
