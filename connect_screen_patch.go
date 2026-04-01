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
	defer dst.Close()

	// 1. Company_Profile
	if err := copySheetDataTransfer(src, dst, "Company_Profile"); err != nil {
		return fmt.Errorf("Company_Profile: %v", err)
	}
	// 2. Subsidiary_Books
	if err := copySheetDataTransfer(src, dst, "Subsidiary_Books"); err != nil {
		return fmt.Errorf("Subsidiary_Books: %v", err)
	}
	// 3. Ledger_Master (logic หลัก)
	if err := transferLedgerMaster(src, dst); err != nil {
		return fmt.Errorf("Ledger_Master: %v", err)
	}

	if err := dst.Save(); err != nil {
		return fmt.Errorf("บันทึกฐานข้อมูลปลายทางไม่สำเร็จ: %v", err)
	}
	return nil
}

// copySheetDataTransfer — ลบ data rows เก่าใน dst แล้ว copy จาก src (ข้าม header)
func copySheetDataTransfer(src, dst *excelize.File, sheet string) error {
	srcRows, _ := src.GetRows(sheet)
	if len(srcRows) == 0 {
		return nil
	}
	// ลบ data rows เก่าออก (จากท้ายขึ้นมา)
	dstRows, _ := dst.GetRows(sheet)
	for i := len(dstRows); i >= 2; i-- {
		dst.RemoveRow(sheet, i)
	}
	// ถ้า dst ไม่มี header ให้ copy จาก src
	dstRows2, _ := dst.GetRows(sheet)
	if len(dstRows2) == 0 {
		for ci, v := range srcRows[0] {
			col, _ := excelize.ColumnNumberToName(ci + 1)
			dst.SetCellValue(sheet, fmt.Sprintf("%s1", col), v)
		}
	}
	// copy data rows
	for ri := 1; ri < len(srcRows); ri++ {
		for ci, v := range srcRows[ri] {
			col, _ := excelize.ColumnNumberToName(ci + 1)
			dst.SetCellValue(sheet, fmt.Sprintf("%s%d", col, ri+1), v)
		}
	}
	return nil
}

// transferLedgerMaster — ย้ายยอดปีเก่า → ปีใหม่
//
// Ledger_Master column layout (1-based Excel col):
//
//	A=Comcode  B=Ac_code  C=Ac_name  D=Gcode  E=Gname
//	F=BBAL  G=CBAL  H=Debit  I=Credit  J=Bthisyear
//	K=Thisper01 ... V=Thisper12   (col 11-22)
//	W=Blastyear                   (col 23)
//	X=Lastper01 ... AI=Lastper12  (col 24-35)
//
// Rules:
//   - Update Gcode/Gname ใน To ตาม AcCode ที่มีอยู่แล้ว
//   - Bthisyear(J) → Blastyear(W)
//   - Thisper01-12(K-V) → Lastper01-12(X-AI)
//   - หมวด 1 (สินทรัพย์), 2 (หนี้สิน), 3 (ทุน): โปรแกรมจะคำนวณยอดยกไปของปีเก่า มาใส่เป็นยอดยกมา (Bthisyear) ของปีใหม่ให้อัตโนมัติ
//   - หมวด 4 (รายได้), 5 (ค่าใช้จ่าย): โปรแกรมจะล้างยอดยกมาเป็น 0.00 ให้ถูกต้องตามหลักการปิดบัญชี
//   - กำไรสุทธิ (360PLA): จะถูกคำนวณและโอนไปทบเป็นยอดยกมาของกำไรสะสม (350RTE) ให้อัตโนมัติ และล้างยอด 360PLA ของปีใหม่เป็น 0
//
// transferLedgerMaster — คัดลอกผังบัญชี (A-E) จากปีเก่ามาปีใหม่ทั้งหมด และคำนวณยอดยกมา
func transferLedgerMaster(src, dst *excelize.File) error {
	const sheet = "Ledger_Master"

	srcRows, _ := src.GetRows(sheet)
	if len(srcRows) <= 1 {
		return nil // ไม่มีข้อมูลให้ Transfer
	}

	// 1. ลบข้อมูลเก่าใน To DB ทิ้งทั้งหมด (เหลือทิ้งไว้แค่ Header แถวที่ 1)
	dstRows, _ := dst.GetRows(sheet)
	for i := len(dstRows); i >= 2; i-- {
		dst.RemoveRow(sheet, i)
	}

	// 2. หาค่ากำไรสุทธิ (360PLA) จาก From DB ก่อน (เพื่อเตรียมไปบวกเข้า 350RTE)
	plaThisper12 := 0.0
	hasPLA := false

	for ri, row := range srcRows {
		if ri == 0 || len(row) < 5 {
			continue
		}
		acCode := safeGet(row, 1) // คอลัมน์ B
		if acCode == "360PLA" {
			hasPLA = true
			// 🚨 ดึงยอด Thisper12 ของ 360PLA มาเลย (เพราะเป็นยอดสะสม YTD อยู่แล้ว)
			plaThisper12 = parseFloat(safeGet(row, 21)) // คอลัมน์ V (index 21)
			break                                       // เจอ 360PLA แล้วหยุดหาได้เลย
		}
	}

	// 3. วนลูปอ่านข้อมูลจาก From DB และเขียนลง To DB
	excelRow := 2 // เริ่มเขียนที่แถว 2 (ต่อจาก Header)
	for ri, row := range srcRows {
		if ri == 0 || len(row) < 5 {
			continue // ข้าม Header หรือแถวที่ข้อมูลไม่ครบ
		}

		// ดึงข้อมูลพื้นฐาน (A-E) จาก From DB
		comCode := safeGet(row, 0)
		acCode := safeGet(row, 1)
		acName := safeGet(row, 2)
		gCode := safeGet(row, 3)
		gName := safeGet(row, 4)

		if acCode == "" {
			continue
		}

		// ดึงยอดยกมาและยอดเคลื่อนไหวของปีเก่า
		bThisYear := parseFloat(safeGet(row, 9))
		var thisPer [12]float64
		for i := 0; i < 12; i++ {
			thisPer[i] = parseFloat(safeGet(row, 10+i))
		}

		// ── 🚨 แก้ไขตรงนี้: ยอดยกไปของปีเก่า คือยอดสะสมเดือน 12 (Thisper12) ──
		// ไม่ใช่การเอาทุกเดือนมาบวกกัน เพราะข้อมูลเป็นแบบ YTD (Year-To-Date)
		oldEndingBalance := thisPer[11]

		// ── กำหนดค่า Bthisyear (ยอดยกมาต้นปี) สำหรับปีใหม่ ──
		newBthisyear := 0.0
		// เช็คหมวดบัญชีจากตัวเลขตัวแรก (1=สินทรัพย์, 2=หนี้สิน, 3=ทุน)
		if strings.HasPrefix(acCode, "1") || strings.HasPrefix(acCode, "2") || strings.HasPrefix(acCode, "3") {
			newBthisyear = oldEndingBalance
		}

		// ── กรณีพิเศษ 350RTE (กำไรสะสม) ──
		if acCode == "350RTE" && hasPLA {
			newBthisyear += plaThisper12
		}

		// ── กรณีพิเศษ 360PLA (กำไรปีปัจจุบัน) ──
		if acCode == "360PLA" {
			newBthisyear = 0.0
		}

		// ── เขียนข้อมูลลง To DB (สร้างแถวใหม่) ──
		dst.SetCellValue(sheet, fmt.Sprintf("A%d", excelRow), comCode)
		dst.SetCellValue(sheet, fmt.Sprintf("B%d", excelRow), acCode)
		dst.SetCellValue(sheet, fmt.Sprintf("C%d", excelRow), acName)
		dst.SetCellValue(sheet, fmt.Sprintf("D%d", excelRow), gCode)
		dst.SetCellValue(sheet, fmt.Sprintf("E%d", excelRow), gName)

		// J = Bthisyear (ยอดยกมาต้นปี)
		dst.SetCellValue(sheet, fmt.Sprintf("J%d", excelRow), newBthisyear)

		// F = BBAL (ยอดยกมาต้นงวด), G = CBAL (ยอดยกไปปลายงวด)
		// เริ่มปีใหม่ งวดที่ 1 ยอดยกมาต้นงวด ต้องเท่ากับ ยอดยกมาต้นปี
		dst.SetCellValue(sheet, fmt.Sprintf("F%d", excelRow), newBthisyear)
		dst.SetCellValue(sheet, fmt.Sprintf("G%d", excelRow), newBthisyear)

		// H = Debit, I = Credit (ยังไม่มีรายการเดินบัญชี ให้เป็น 0)
		dst.SetCellValue(sheet, fmt.Sprintf("H%d", excelRow), 0.0)
		dst.SetCellValue(sheet, fmt.Sprintf("I%d", excelRow), 0.0)

		// K ถึง V (Thisper01-12) = 0
		for i := 0; i < 12; i++ {
			col, _ := excelize.ColumnNumberToName(11 + i)
			dst.SetCellValue(sheet, fmt.Sprintf("%s%d", col, excelRow), 0.0)
		}

		// ── จัดการข้อมูลเปรียบเทียบปีที่แล้ว (Last Year) ──
		if acCode == "360PLA" {
			// 360PLA ไม่ต้องเก็บประวัติ Last Year
			dst.SetCellValue(sheet, fmt.Sprintf("W%d", excelRow), 0.0)
			for i := 0; i < 12; i++ {
				col, _ := excelize.ColumnNumberToName(24 + i)
				dst.SetCellValue(sheet, fmt.Sprintf("%s%d", col, excelRow), 0.0)
			}
		} else {
			// บัญชีปกติ: ย้าย Bthisyear เดิม -> Blastyear(W), Thisper01-12 เดิม -> Lastper01-12(X-AI)
			dst.SetCellValue(sheet, fmt.Sprintf("W%d", excelRow), bThisYear)
			for i := 0; i < 12; i++ {
				col, _ := excelize.ColumnNumberToName(24 + i)
				dst.SetCellValue(sheet, fmt.Sprintf("%s%d", col, excelRow), thisPer[i])
			}

			// 350RTE: เอายอดเคลื่อนไหวเดือน 12 ของ 360PLA มาโชว์ใน Lastper12 ด้วย
			if acCode == "350RTE" && hasPLA {
				newLastper12 := thisPer[11] + plaThisper12
				colAI, _ := excelize.ColumnNumberToName(35)
				dst.SetCellValue(sheet, fmt.Sprintf("%s%d", colAI, excelRow), newLastper12)
			}
		}

		excelRow++ // ขยับไปเขียนแถวถัดไป
	}

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
// showConnectScreen — หน้า Login พร้อม Create New Database และ Transfer
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

	// ── ฟังก์ชัน reload รายชื่อฐานข้อมูล ───────────────────────────
	var allNames []string
	reloadNames := func() {
		allNames = loadDBNames()
	}
	reloadNames()

	// ── 🚨 ตัวแปรควบคุมไม่ให้ Dropdown ซ้อนกัน ─────────────────────
	var activeDropdown *autocompleteEntry
	var currentHideSugg func()

	hideAllDropdowns := func() {
		if currentHideSugg != nil {
			currentHideSugg()
			currentHideSugg = nil
			activeDropdown = nil
		}
	}

	// ── Helper: สร้างช่องกรอกแบบ Auto-Complete ────────────────────
	makeAutoEntry := func(placeholder string, onSubmit func(string)) (*autocompleteEntry, *fyne.Container) {
		entry := newAutocompleteEntry()
		entry.SetPlaceHolder(placeholder)

		var filteredNames []string
		selectedIdx := -1
		isNavigating := false
		suppressDrop := false // 🚨 ป้องกัน OnChanged ทำงานซ้ำตอนคลิกเลือก
		suggWrapper := container.NewVBox()

		suggList := widget.NewList(
			func() int { return len(filteredNames) },
			func() fyne.CanvasObject { return widget.NewLabel("") },
			func(id widget.ListItemID, o fyne.CanvasObject) {
				if int(id) < len(filteredNames) {
					o.(*widget.Label).SetText(filteredNames[id])
				}
			},
		)

		hideSugg := func() {
			selectedIdx = -1
			suggList.UnselectAll() // 🚨 ล้างค่าการเลือกทิ้ง
			suggWrapper.Objects = nil
			suggWrapper.Refresh()
		}

		showSugg := func() {
			// 🚨 ถ้ามี Dropdown อื่นเปิดอยู่ (ที่ไม่ใช่ของตัวเอง) ให้ปิดตัวนั้นก่อน
			if activeDropdown != nil && activeDropdown != entry {
				hideAllDropdowns()
			}

			activeDropdown = entry
			currentHideSugg = hideSugg

			if len(suggWrapper.Objects) == 0 && len(filteredNames) > 0 {
				maxH := float32(28) * float32(min4(len(filteredNames), 5))
				suggWrapper.Add(container.NewGridWrap(fyne.NewSize(340, maxH), suggList))
				suggWrapper.Refresh()
			}
		}

		suggList.OnSelected = func(id widget.ListItemID) {
			if isNavigating {
				return
			}
			if int(id) < len(filteredNames) {
				selected := filteredNames[id]

				suppressDrop = true // 🚨 ปิดการทำงานของ OnChanged ชั่วคราว
				entry.SetText(selected)
				suppressDrop = false

				hideSugg()
				if activeDropdown == entry {
					activeDropdown = nil
					currentHideSugg = nil
				}

				if onSubmit != nil {
					onSubmit(selected)
				}
			}
		}

		entry.onDown = func() {
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

		entry.onUp = func() {
			if len(filteredNames) == 0 || selectedIdx <= 0 {
				return
			}
			isNavigating = true
			selectedIdx--
			suggList.Select(widget.ListItemID(selectedIdx))
			suggList.ScrollTo(widget.ListItemID(selectedIdx))
			isNavigating = false
		}

		entry.onEnter = func() {
			if selectedIdx >= 0 && selectedIdx < len(filteredNames) {
				selected := filteredNames[selectedIdx]

				suppressDrop = true
				entry.SetText(selected)
				suppressDrop = false

				hideSugg()
				if activeDropdown == entry {
					activeDropdown = nil
					currentHideSugg = nil
				}

				if onSubmit != nil {
					onSubmit(selected)
				}
			} else {
				hideSugg()
				if activeDropdown == entry {
					activeDropdown = nil
					currentHideSugg = nil
				}
				if onSubmit != nil {
					onSubmit(entry.Text)
				}
			}
		}

		entry.OnChanged = func(s string) {
			if suppressDrop {
				return // 🚨 ถ้าโปรแกรมกำลัง SetText เอง ห้ามทำ OnChanged
			}

			keyword := strings.ToLower(strings.TrimSuffix(s, ".xlsx"))
			filteredNames = nil
			selectedIdx = -1
			for _, name := range allNames {
				if keyword == "" || strings.Contains(strings.ToLower(name), keyword) {
					filteredNames = append(filteredNames, name)
				}
			}
			suggList.UnselectAll()
			suggList.Refresh()

			if s == "" || len(filteredNames) == 0 {
				hideSugg()
			} else {
				showSugg()
			}
		}

		entry.OnSubmitted = func(s string) {
			hideSugg()
			if activeDropdown == entry {
				activeDropdown = nil
				currentHideSugg = nil
			}
			if onSubmit != nil {
				onSubmit(s)
			}
		}

		return entry, suggWrapper
	}

	// ── Section 1: Log In ─────────────────────────────────────────
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

	enDBName, suggLogin := makeAutoEntry("พิมพ์ชื่อฐานข้อมูล...", doLogin)

	btnLogin := widget.NewButtonWithIcon("Log In", theme.LoginIcon(), func() {
		hideAllDropdowns()
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
			suggLogin,
		),
	)

	// ── Section 2: Create New Database ───────────────────────────
	enNewDB := newSmartEntry(nil)
	enNewDB.SetPlaceHolder("ปี_เลขtaxid13หลัก_สาขา(ถ้ามี) (เช่น 2025_0100000000000)")

	statusCreate := widget.NewLabel("")
	statusCreate.Wrapping = fyne.TextWrapWord

	btnCreate := widget.NewButtonWithIcon("Create", theme.ContentAddIcon(), func() {
		hideAllDropdowns()

		name := strings.TrimSpace(enNewDB.Text)
		if name == "" {
			statusCreate.SetText("⚠️  กรุณากรอกชื่อฐานข้อมูล")
			return
		}
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
	enFromDB, suggFrom := makeAutoEntry("ฐานข้อมูลต้นทาง (From)", nil)
	enToDB, suggTo := makeAutoEntry("ฐานข้อมูลปลายทาง (To)", nil)

	statusTransfer := widget.NewLabel("")
	statusTransfer.Wrapping = fyne.TextWrapWord

	makeBrowseBtn := func(targetEntry *autocompleteEntry) *widget.Button {
		return widget.NewButtonWithIcon("", theme.FolderOpenIcon(), func() {
			hideAllDropdowns()

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
				hideAllDropdowns()
				d.Hide()
			}
			d = dialog.NewCustomWithoutButtons("เลือกฐานข้อมูล",
				container.NewGridWrap(fyne.NewSize(320, 250), list), w)
			d.Show()
		})
	}

	btnTransfer := widget.NewButtonWithIcon("Update GL", theme.ViewRefreshIcon(), func() {
		hideAllDropdowns()

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
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(45, 1), widget.NewLabel("")), // indent
			suggFrom,
		),
		container.NewBorder(nil, nil,
			widget.NewLabel("To      :"),
			makeBrowseBtn(enToDB),
			container.NewGridWrap(fyne.NewSize(340, 32), enToDB),
		),
		container.NewHBox(
			container.NewGridWrap(fyne.NewSize(45, 1), widget.NewLabel("")), // indent
			suggTo,
		),
		container.NewHBox(
			layout.NewSpacer(),
			btnTransfer,
		),
		statusTransfer,
	)

	// ── Main Card Layout ──────────────────────────────────────────
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
	cardWidth := fyne.NewSize(560, 0)
	centeredRow := container.NewHBox(
		layout.NewSpacer(),
		container.NewGridWrap(cardWidth, formCard),
		layout.NewSpacer(),
	)

	bodyWithPad := container.NewVBox(
		container.NewGridWrap(fyne.NewSize(1, 12), widget.NewLabel("")), // spacer 12px
		centeredRow,
	)

	mainLayout := container.NewBorder(
		container.NewVBox(
			container.NewPadded(header),
			widget.NewSeparator(),
		),
		container.NewVBox(
			widget.NewSeparator(),
			container.NewPadded(footer),
		),
		nil, nil,
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
