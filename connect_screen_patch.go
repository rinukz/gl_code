// connect_screen_patch.go
// วางไว้ใน package main (แทนที่ showConnectScreen เดิมใน main.go)
// ─────────────────────────────────────────────────────────────────
// ฟีเจอร์ที่เพิ่ม:
//   1. Create New Database  — Copy raw_database.xlsx → dont_edit/<name>.xlsx
//   2. Database Transfer    — Copy From DB → To DB (Update GL / migrate)
// ─────────────────────────────────────────────────────────────────

package main

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// createNewDatabase — Copy raw_database.xlsx → dont_edit/<newName>.xlsx
//
//	templatePath : path ของ raw_database.xlsx (ไฟล์ต้นแบบ)
//	newDbPath    : path ปลายทางที่จะสร้าง
//
// ─────────────────────────────────────────────────────────────────
func createNewDatabase(templatePath string, newDbPath string) error {
	srcFile, err := os.Open(templatePath)
	if err != nil {
		return fmt.Errorf("เปิดไฟล์ต้นแบบไม่ได้: %v", err)
	}
	defer srcFile.Close()

	dstFile, err := os.Create(newDbPath)
	if err != nil {
		return fmt.Errorf("สร้างไฟล์ใหม่ไม่ได้: %v", err)
	}
	defer dstFile.Close()

	if _, err := io.Copy(dstFile, srcFile); err != nil {
		return fmt.Errorf("คัดลอกข้อมูลไม่สำเร็จ: %v", err)
	}
	return dstFile.Sync()
}

// ─────────────────────────────────────────────────────────────────
// transferDatabase — Copy Book_items + Ledger_Master จาก fromDB → toDB
//
//	ใช้สำหรับ "Update GL" / migrate ข้ามปี
//	โดย toDB ต้องมีอยู่แล้ว (สร้างด้วย createNewDatabase ก่อน)
//
// ─────────────────────────────────────────────────────────────────
func transferDatabase(fromPath, toPath string) error {
	xlOptions := excelize.Options{Password: "@A123456789a"}

	src, err := excelize.OpenFile(fromPath, xlOptions)
	if err != nil {
		return fmt.Errorf("เปิดฐานข้อมูลต้นทางไม่ได้: %v", err)
	}
	defer src.Close()

	dst, err := excelize.OpenFile(toPath, xlOptions)
	if err != nil {
		return fmt.Errorf("เปิดฐานข้อมูลปลายทางไม่ได้: %v", err)
	}

	// sheets ที่จะ transfer
	sheets := []string{
		"Company_Profile",
		"Ledger_Master",
		"Acct_Group",
		"Subsidiary_Books",
		"Special_code",
		"Customer_Log",
	}

	for _, sheet := range sheets {
		rows, _ := src.GetRows(sheet)
		if len(rows) == 0 {
			continue
		}
		// ลบข้อมูลเก่าใน dst (เว้นแต่ header row 1)
		dstRows, _ := dst.GetRows(sheet)
		for i := len(dstRows); i >= 2; i-- {
			dst.RemoveRow(sheet, i)
		}
		// เขียน header จาก src ถ้า dst ว่าง
		dstRows2, _ := dst.GetRows(sheet)
		if len(dstRows2) == 0 {
			for i, h := range rows[0] {
				col, _ := excelize.ColumnNumberToName(i + 1)
				dst.SetCellValue(sheet, fmt.Sprintf("%s1", col), h)
			}
		}
		// เขียน data rows
		for ri, row := range rows {
			if ri == 0 {
				continue // ข้าม header
			}
			for ci, val := range row {
				col, _ := excelize.ColumnNumberToName(ci + 1)
				dst.SetCellValue(sheet, fmt.Sprintf("%s%d", col, ri+1), val)
			}
		}
	}

	if err := dst.Save(); err != nil {
		dst.Close()
		return fmt.Errorf("บันทึกฐานข้อมูลปลายทางไม่สำเร็จ: %v", err)
	}
	dst.Close()
	return nil
}

// ─────────────────────────────────────────────────────────────────
// loadDBNames — scan ไฟล์ .xlsx ใน ./dont_edit/
// ─────────────────────────────────────────────────────────────────
func loadDBNames() []string {
	var names []string
	if entries, err := os.ReadDir("./dont_edit/"); err == nil {
		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(entry.Name(), ".xlsx") {
				name := strings.TrimSuffix(entry.Name(), ".xlsx")
				// ข้าม raw_database
				if name == "raw_database" {
					continue
				}
				names = append(names, name)
			}
		}
	}
	return names
}

// ─────────────────────────────────────────────────────────────────
// showConnectScreen — หน้า Login พร้อม Create New Database
// ─────────────────────────────────────────────────────────────────
func showConnectScreen(w fyne.Window, a fyne.App) {
	w.SetMainMenu(nil)

	// ── Header ──────────────────────────────────────────────────
	logoImg := canvas.NewImageFromFile("logo.png")
	logoImg.FillMode = canvas.ImageFillContain
	logoImg.SetMinSize(fyne.NewSize(80, 80))

	appTitle := canvas.NewText("General Ledger", color.Black)
	appTitle.TextStyle = fyne.TextStyle{Bold: true}
	appTitle.TextSize = 36

	appSubtitle := canvas.NewText("Accounting System", color.RGBA{R: 80, G: 80, B: 80, A: 255})
	appSubtitle.TextSize = 16

	header := container.NewHBox(
		layout.NewSpacer(),
		logoImg,
		container.NewVBox(
			layout.NewSpacer(),
			appTitle,
			appSubtitle,
			layout.NewSpacer(),
		),
		layout.NewSpacer(),
	)

	// ── ฟังก์ชัน reload dropdown ──────────────────────────────────
	var allNames []string
	var filteredNames []string
	var suggList *widget.List
	var suggWrapper *fyne.Container

	reloadNames := func() {
		allNames = loadDBNames()
		filteredNames = make([]string, len(allNames))
		copy(filteredNames, allNames)
		if suggList != nil {
			suggList.Refresh()
		}
	}
	reloadNames()

	// ── Section 1: Log In ─────────────────────────────────────────
	selectedIdx := -1
	isNavigating := false

	enDBName := newAutocompleteEntry()
	enDBName.SetPlaceHolder("พิมพ์ชื่อฐานข้อมูล...")

	doLogin := func(input string) {
		input = strings.TrimSuffix(strings.TrimSpace(input), ".xlsx")
		if input == "" {
			return
		}
		fullPath := "./dont_edit/" + input + ".xlsx"
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			dialog.ShowError(fmt.Errorf("ไม่พบฐานข้อมูล: %s", input), w)
			return
		}
		currentDBPath = fullPath

		loadingLabel := widget.NewLabelWithStyle(
			"กำลังเชื่อมต่อ... "+input,
			fyne.TextAlignCenter, fyne.TextStyle{Italic: true},
		)
		w.SetContent(container.NewCenter(loadingLabel))

		safeGo(func() {
			xlOptions := excelize.Options{Password: "@A123456789a"}
			VerifyCompanyCode(xlOptions)
			VerifySpecialAccounts(xlOptions)
			VerifyAcctGroup(xlOptions)
			fyne.Do(func() {
				showMainInterface(w, a)
			})
		})
	}

	hideSugg := func() {
		selectedIdx = -1
		if suggWrapper != nil {
			suggWrapper.Objects = nil
			suggWrapper.Refresh()
		}
	}

	showSugg := func() {
		if suggWrapper != nil && len(suggWrapper.Objects) == 0 && len(filteredNames) > 0 {
			maxH := float32(28) * float32(min4(len(filteredNames), 5))
			suggWrapper.Add(container.NewGridWrap(fyne.NewSize(340, maxH), suggList))
			suggWrapper.Refresh()
		}
	}

	suggList = widget.NewList(
		func() int { return len(filteredNames) },
		func() fyne.CanvasObject { return widget.NewLabel("") },
		func(id widget.ListItemID, o fyne.CanvasObject) {
			if int(id) < len(filteredNames) {
				o.(*widget.Label).SetText(filteredNames[id])
			}
		},
	)
	suggList.OnSelected = func(id widget.ListItemID) {
		if isNavigating {
			return
		}
		if int(id) < len(filteredNames) {
			selected := filteredNames[id]
			hideSugg()
			enDBName.SetText(selected)
			doLogin(selected)
		}
	}

	enDBName.onDown = func() {
		if len(filteredNames) == 0 {
			return
		}
		showSugg()
		isNavigating = true
		if selectedIdx < len(filteredNames)-1 {
			selectedIdx++
		}
		suggList.Select(widget.ListItemID(selectedIdx))
		suggList.ScrollTo(widget.ListItemID(selectedIdx))
		isNavigating = false
	}
	enDBName.onUp = func() {
		if len(filteredNames) == 0 || selectedIdx <= 0 {
			return
		}
		isNavigating = true
		selectedIdx--
		suggList.Select(widget.ListItemID(selectedIdx))
		suggList.ScrollTo(widget.ListItemID(selectedIdx))
		isNavigating = false
	}
	enDBName.onEnter = func() {
		if selectedIdx >= 0 && selectedIdx < len(filteredNames) {
			selected := filteredNames[selectedIdx]
			hideSugg()
			enDBName.SetText(selected)
			doLogin(selected)
		} else {
			hideSugg()
			doLogin(enDBName.Text)
		}
	}
	enDBName.OnChanged = func(s string) {
		keyword := strings.ToLower(strings.TrimSuffix(s, ".xlsx"))
		filteredNames = nil
		selectedIdx = -1
		for _, name := range allNames {
			if keyword == "" || strings.Contains(strings.ToLower(name), keyword) {
				filteredNames = append(filteredNames, name)
			}
		}
		if suggList != nil {
			suggList.Refresh()
		}
		if s == "" || len(filteredNames) == 0 {
			hideSugg()
		} else {
			showSugg()
		}
	}
	enDBName.OnSubmitted = func(s string) {
		hideSugg()
		doLogin(s)
	}

	suggWrapper = container.NewVBox()

	btnLogin := widget.NewButtonWithIcon("Log In", theme.LoginIcon(), func() {
		hideSugg()
		doLogin(enDBName.Text)
	})
	btnLogin.Importance = widget.HighImportance

	loginSection := container.NewVBox(
		widget.NewLabelWithStyle("เลือกฐานข้อมูล", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil,
			widget.NewLabel("Database :"),
			container.NewGridWrap(fyne.NewSize(90, 32), btnLogin),
			container.NewGridWrap(fyne.NewSize(340, 32), enDBName),
		),
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(72, 1), widget.NewLabel("")), // indent
			suggWrapper,
		),
	)

	// ── Section 2: Create New Database ───────────────────────────
	enNewDB := newSmartEntry(nil)
	enNewDB.SetPlaceHolder("ปี_เลขtaxid13หลัก_สาขา(ถ้ามี) (เช่น 2025_0100000000000)")

	statusCreate := widget.NewLabel("")
	statusCreate.Wrapping = fyne.TextWrapWord

	btnCreate := widget.NewButtonWithIcon("Create", theme.ContentAddIcon(), func() {
		name := strings.TrimSpace(enNewDB.Text)
		if name == "" {
			statusCreate.SetText("⚠️  กรุณากรอกชื่อฐานข้อมูล")
			return
		}
		// ตรวจชื่อ: ห้ามมี / \ : * ? " < > |
		invalidChars := `/\:*?"<>|`
		for _, ch := range invalidChars {
			if strings.ContainsRune(name, ch) {
				statusCreate.SetText(fmt.Sprintf("⚠️  ชื่อห้ามมีอักขระพิเศษ: %s", invalidChars))
				return
			}
		}

		templatePath := "./dont_edit/raw_database.xlsx"
		if _, err := os.Stat(templatePath); os.IsNotExist(err) {
			statusCreate.SetText("❌  ไม่พบไฟล์ต้นแบบ: raw_database.xlsx")
			return
		}

		newPath := fmt.Sprintf("./dont_edit/%s.xlsx", name)
		if _, err := os.Stat(newPath); !os.IsNotExist(err) {
			statusCreate.SetText(fmt.Sprintf("⚠️  มีฐานข้อมูลชื่อ \"%s\" อยู่แล้ว", name))
			return
		}

		statusCreate.SetText("กำลังสร้าง...")
		safeGo(func() {
			err := createNewDatabase(templatePath, newPath)
			fyne.Do(func() {
				if err != nil {
					statusCreate.SetText("❌  " + err.Error())
					return
				}
				statusCreate.SetText(fmt.Sprintf("✅  สร้าง \"%s\" สำเร็จแล้ว", name))
				enNewDB.SetText("")
				reloadNames()

				var d dialog.Dialog
				okBtn := newEnterButton("เปิดใช้งานทันที", func() {
					d.Hide()
					doLogin(name)
				})
				okBtn.Importance = widget.HighImportance
				laterBtn := widget.NewButton("ปิด", func() { d.Hide() })

				d = dialog.NewCustomWithoutButtons(
					"✅ สร้างฐานข้อมูลสำเร็จ",
					container.NewVBox(
						widget.NewLabel(fmt.Sprintf("สร้างฐานข้อมูล \"%s\" เรียบร้อยแล้ว", name)),
						widget.NewLabel("ต้องการเปิดใช้งานทันทีหรือไม่?"),
						container.NewCenter(container.NewHBox(okBtn, laterBtn)),
					), w)
				d.Show()
				fyne.Do(func() { w.Canvas().Focus(okBtn) })
			})
		})
	})
	btnCreate.Importance = widget.HighImportance

	createSection := container.NewVBox(
		widget.NewLabelWithStyle("สร้างฐานข้อมูลใหม่", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil,
			widget.NewLabel("ชื่อฐานข้อมูล :"),
			container.NewGridWrap(fyne.NewSize(90, 32), btnCreate),
			container.NewGridWrap(fyne.NewSize(340, 32), enNewDB),
		),
		statusCreate,
	)

	// ── Section 3: Database Transfer (Update GL) ──────────────────
	enFromDB := newSmartEntry(nil)
	enFromDB.SetPlaceHolder("ฐานข้อมูลต้นทาง (From)")
	enToDB := newSmartEntry(nil)
	enToDB.SetPlaceHolder("ฐานข้อมูลปลายทาง (To)")

	statusTransfer := widget.NewLabel("")
	statusTransfer.Wrapping = fyne.TextWrapWord

	// Browse buttons
	makeBrowseBtn := func(targetEntry *smartEntry) *widget.Button {
		return widget.NewButtonWithIcon("", theme.FolderOpenIcon(), func() {
			names := loadDBNames()
			if len(names) == 0 {
				dialog.ShowInformation("แจ้งเตือน", "ไม่พบฐานข้อมูลใน ./dont_edit/", w)
				return
			}
			var d dialog.Dialog
			list := widget.NewList(
				func() int { return len(names) },
				func() fyne.CanvasObject { return widget.NewLabel("") },
				func(id widget.ListItemID, o fyne.CanvasObject) {
					o.(*widget.Label).SetText(names[id])
				},
			)
			list.OnSelected = func(id widget.ListItemID) {
				targetEntry.SetText(names[id])
				d.Hide()
			}
			d = dialog.NewCustomWithoutButtons("เลือกฐานข้อมูล",
				container.NewGridWrap(fyne.NewSize(320, 250), list), w)
			d.Show()
		})
	}

	btnTransfer := widget.NewButtonWithIcon("Update GL", theme.ViewRefreshIcon(), func() {
		fromName := strings.TrimSpace(enFromDB.Text)
		toName := strings.TrimSpace(enToDB.Text)

		if fromName == "" || toName == "" {
			statusTransfer.SetText("⚠️  กรุณาเลือกทั้งฐานข้อมูลต้นทางและปลายทาง")
			return
		}
		if fromName == toName {
			statusTransfer.SetText("⚠️  ต้นทางและปลายทางต้องไม่ใช่ฐานข้อมูลเดียวกัน")
			return
		}

		fromPath := fmt.Sprintf("./dont_edit/%s.xlsx", fromName)
		toPath := fmt.Sprintf("./dont_edit/%s.xlsx", toName)

		if _, err := os.Stat(fromPath); os.IsNotExist(err) {
			statusTransfer.SetText(fmt.Sprintf("❌  ไม่พบฐานข้อมูลต้นทาง: %s", fromName))
			return
		}
		if _, err := os.Stat(toPath); os.IsNotExist(err) {
			statusTransfer.SetText(fmt.Sprintf("❌  ไม่พบฐานข้อมูลปลายทาง: %s", toName))
			return
		}

		var confirmD dialog.Dialog
		yesBtn := newEnterButton("ยืนยัน Transfer", func() {
			confirmD.Hide()
			statusTransfer.SetText("กำลัง Transfer ข้อมูล...")
			safeGo(func() {
				err := transferDatabase(fromPath, toPath)
				fyne.Do(func() {
					if err != nil {
						statusTransfer.SetText("❌  " + err.Error())
					} else {
						statusTransfer.SetText(fmt.Sprintf(
							"✅  Transfer เสร็จสมบูรณ์\n%s → %s", fromName, toName))
					}
				})
			})
		})
		yesBtn.Importance = widget.DangerImportance
		noBtn := widget.NewButton("ยกเลิก", func() { confirmD.Hide() })
		confirmD = dialog.NewCustomWithoutButtons("ยืนยัน Transfer",
			container.NewVBox(
				widget.NewLabel(fmt.Sprintf("จะโอนข้อมูล Master จาก\n\"%s\"  →  \"%s\"", fromName, toName)),
				widget.NewLabel("⚠️  ข้อมูลเดิมในปลายทางจะถูกแทนที่"),
				container.NewCenter(container.NewHBox(yesBtn, noBtn)),
			), w)
		confirmD.Show()
		fyne.Do(func() { w.Canvas().Focus(noBtn) })
	})

	transferSection := container.NewVBox(
		widget.NewLabelWithStyle("Database Transfer", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil,
			widget.NewLabel("From :"),
			makeBrowseBtn(enFromDB),
			container.NewGridWrap(fyne.NewSize(340, 32), enFromDB),
		),
		container.NewBorder(nil, nil,
			widget.NewLabel("To      :"),
			makeBrowseBtn(enToDB),
			container.NewGridWrap(fyne.NewSize(340, 32), enToDB),
		),
		container.NewHBox(
			layout.NewSpacer(),
			btnTransfer,
		),
		statusTransfer,
	)

	// ── Main Card Layout ──────────────────────────────────────────
	makeDivider := func(label string) fyne.CanvasObject {
		line := canvas.NewLine(color.RGBA{R: 200, G: 200, B: 200, A: 255})
		line.StrokeWidth = 1
		lbl := widget.NewLabelWithStyle(label, fyne.TextAlignCenter,
			fyne.TextStyle{Bold: true, Italic: true})
		return container.NewVBox(
			container.NewHBox(line, lbl, line),
		)
	}
	_ = makeDivider // ถ้าไม่ใช้ให้ลบบรรทัดนี้

	formCard := container.NewVBox(
		loginSection,
		widget.NewSeparator(),
		createSection,
		widget.NewSeparator(),
		transferSection,
	)

	// ── Footer ───────────────────────────────────────────────────
	btnExit := widget.NewButtonWithIcon("EXIT", theme.CancelIcon(), func() { a.Quit() })
	footer := container.NewHBox(
		layout.NewSpacer(),
		container.NewGridWrap(fyne.NewSize(90, 36), btnExit),
	)

	// ── Full Layout ───────────────────────────────────────────────
	// ✅ วาง formCard ชิดบน-กลาง โดยใช้ HBox + Spacer ซ้าย/ขวา
	// ไม่ใช้ NewCenter เพราะมันดัน content ลงกลางแนวตั้งเสมอ
	cardWidth := fyne.NewSize(560, 0)
	centeredRow := container.NewHBox(
		layout.NewSpacer(),
		container.NewGridWrap(cardWidth, formCard),
		layout.NewSpacer(),
	)

	// padding ด้านบน form นิดหน่อย ให้ดูมีระยะห่างจาก header
	bodyWithPad := container.NewVBox(
		container.NewGridWrap(fyne.NewSize(1, 12), widget.NewLabel("")), // spacer 12px
		centeredRow,
	)

	mainLayout := container.NewBorder(
		// Top: header + separator
		container.NewVBox(
			container.NewPadded(header),
			widget.NewSeparator(),
		),
		// Bottom: separator + footer
		container.NewVBox(
			widget.NewSeparator(),
			container.NewPadded(footer),
		),
		nil, nil,
		// Center: form ชิดบน ไม่ scroll ถ้าหน้าจอสูงพอ
		container.NewVScroll(bodyWithPad),
	)

	w.SetContent(mainLayout)

	// Focus login field
	fyne.Do(func() { w.Canvas().Focus(enDBName) })
}

// ─────────────────────────────────────────────────────────────────
// min4 — helper ใช้ใน suggList height calc
// ─────────────────────────────────────────────────────────────────
func min4(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// ─────────────────────────────────────────────────────────────────
// openFilePathInExplorer — helper เปิด folder ใน file manager (optional)
// ─────────────────────────────────────────────────────────────────
func dbFolderPath() string {
	abs, _ := filepath.Abs("./dont_edit")
	return abs
}
