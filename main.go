package main

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime/debug"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"

	"fyne.io/fyne/v2/container"

	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

// var currentDBPath string // เก็บ Path เต็ม เช่น ./dont_edit/2026_xxx.xlsx
var isConnected bool = false
var backupStopCh chan struct{} // cancel channel สำหรับ background goroutines

// ─────────────────────────────────────────────────────────
// safeGo — ห่อ goroutine ให้จับ panic แล้วเขียน log
// ใช้แทน go func(){}() ทุกที่ที่สำคัญ
// overhead: defer 1 ตัว (~100 bytes/goroutine), 0% CPU ขณะรันปกติ
// ─────────────────────────────────────────────────────────
func safeGo(fn func()) {
	go func() {
		defer func() {
			if r := recover(); r != nil {
				writeLog(fmt.Sprintf("[%s] GOROUTINE PANIC: %v\n--- stack trace ---\n%s\n",
					time.Now().Format("2006-01-02 15:04:05"), r, debug.Stack()))
			}
		}()
		fn()
	}()
}

// crashLog — global log writer (nil-safe)
// ตั้งค่าใน main() ก่อน ShowAndRun
var crashLog io.Writer = os.Stderr // fallback = stderr เสมอ

// writeLog — เขียน log แบบ nil-safe ผ่าน crashLog
func writeLog(msg string) {
	if crashLog != nil {
		fmt.Fprint(crashLog, msg)
	}
}

// --- [ 1. หน้า Login ฉบับ Autocomplete & Clean UI ] ---
// ---------------------------------------------------------

// ─────────────────────────────────────────────────────────
// backupDatabase — copy currentDBPath → backup_database/
// ─────────────────────────────────────────────────────────
func backupDatabase() {
	if currentDBPath == "" {
		return
	}
	backupDir := filepath.Join(filepath.Dir(currentDBPath), "..", "backup_database")
	if err := os.MkdirAll(backupDir, 0755); err != nil {
		writeLog(fmt.Sprintf("[Backup] สร้าง folder ไม่ได้: %v\n", err))
		return
	}

	base := strings.TrimSuffix(filepath.Base(currentDBPath), ".xlsx")
	timestamp := time.Now().Format("20060102_150405")
	dstPath := filepath.Join(backupDir, fmt.Sprintf("%s_%s.xlsx", base, timestamp))

	src, err := os.Open(currentDBPath)
	if err != nil {
		writeLog(fmt.Sprintf("[Backup] เปิดไฟล์ต้นฉบับไม่ได้: %v\n", err))
		return
	}
	defer src.Close()

	dst, err := os.Create(dstPath)
	if err != nil {
		writeLog(fmt.Sprintf("[Backup] สร้างไฟล์ backup ไม่ได้: %v\n", err))
		return
	}
	defer dst.Close()

	if _, err := io.Copy(dst, src); err != nil {
		writeLog(fmt.Sprintf("[Backup] copy ไม่สำเร็จ: %v\n", err))
		return
	}

	writeLog(fmt.Sprintf("[Backup] ✅ %s\n", filepath.Base(dstPath)))

	pattern := filepath.Join(backupDir, base+"_*.xlsx")
	files, _ := filepath.Glob(pattern)
	if len(files) > 168 {
		for _, old := range files[:len(files)-168] {
			os.Remove(old)
			writeLog(fmt.Sprintf("[Backup] ลบเก่า: %s\n", filepath.Base(old)))
		}
	}
}

func showMainInterface(w fyne.Window, a fyne.App) {
	// ─── หยุด goroutines เก่าก่อน (กัน leak ตอน Disconnect→Login ใหม่) ───
	if backupStopCh != nil {
		close(backupStopCh)
	}
	backupStopCh = make(chan struct{})

	// ─── Auto Backup ทุก 1 ชม. ───
	go func(stop chan struct{}) {
		ticker := time.NewTicker(1 * time.Hour)
		defer ticker.Stop()
		for {
			select {
			case <-ticker.C:
				backupDatabase()
			case <-stop:
				return
			}
		}
	}(backupStopCh)

	// ─── Keep-Alive ป้องกัน OS มองว่าแอปหลับลึก ───
	go func(stop chan struct{}) {
		ticker := time.NewTicker(1 * time.Minute)
		defer ticker.Stop()
		for {
			select {
			case <-ticker.C:
				fyne.Do(func() {
					if w.Content() != nil {
						w.Canvas().Refresh(w.Content())
					}
				})
			case <-stop:
				return
			}
		}
	}(backupStopCh)

	// 1. สร้าง pages
	homePage := getHomeGUI()
	companyPage := getCompanySetupGUI(w)

	ledgerPage, resetLedger := getLedgerGUI(w)

	var localBookPage fyne.CanvasObject
	var localResetBook func()
	localBookPage, localResetBook = getBookGUI(w)
	bookPage = localBookPage
	resetBook = localResetBook

	contentContainer := container.NewStack(homePage)
	var currentPage fyne.CanvasObject = homePage
	changePage = func(newContent fyne.CanvasObject) {
		currentPage = newContent
		contentContainer.Objects = []fyne.CanvasObject{newContent}
		contentContainer.Refresh()
	}

	// ❌ โค้ดเดิม:
	// bookGotoFunc = func(bitem string) {
	// 	if localResetBook != nil {
	// 		localResetBook()
	// 	}
	// 	changePage(bookPage)
	// 	if bookSearchFunc != nil {
	// 		safeGo(func() {
	// 			fyne.Do(func() { bookSearchFunc() })
	// 		})
	// 	}
	// }

	// 2. Setup Menu
	setupMenu := fyne.NewMenu("",
		fyne.NewMenuItem("Company", func() { changePage(companyPage) }),
		fyne.NewMenuItem("Subsidiary Books", func() { changePage(getSubsidiaryBooksGUI(w)) }),
		fyne.NewMenuItem("Special Account Code", func() { changePage(getSpecialAccountGUI(w)) }),
		fyne.NewMenuItem("Capital", func() { changePage(getCapitalGUI(w)) }),
	)

	reportMenu := fyne.NewMenu("",
		fyne.NewMenuItem("Statement of Financial Position", func() { showBalanceSheetDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Statement of Earnings", func() { showProfitAndLossDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Cost of Goods Sold Statements", func() { showCGSDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("Trial Balance V1  [F12]", func() { showTrialBalanceDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Trial Balance V2", func() { showTrialBalanceV2Dialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("Subsidiary Books", func() { showSubbookDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Ledger", func() { showLedgerReportDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Worksheet", func() { showWorksheetDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("Purchases VAT", func() { showVatDialog(w, "P", func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Sales VAT", func() { showVatDialog(w, "S", func() { changePage(companyPage) }) }),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("Account Closing Balance", func() { showClosingBalanceDialog(w, func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Account Receivable", func() { showARAPDialog(w, "AR", func() { changePage(companyPage) }) }),
		fyne.NewMenuItem("Account Payable", func() { showARAPDialog(w, "AP", func() { changePage(companyPage) }) }),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("Account Group", func() { showAcctGroupDialog(w, func() { changePage(companyPage) }) }),
	)

	// 3. NavBar
	btnSetup := widget.NewButtonWithIcon("Setup", theme.SettingsIcon(), nil)
	btnSetup.OnTapped = func() {
		pos := fyne.CurrentApp().Driver().AbsolutePositionForObject(btnSetup)
		pos.Y += btnSetup.Size().Height
		widget.ShowPopUpMenuAtPosition(setupMenu, w.Canvas(), pos)
	}

	btnReport := widget.NewButtonWithIcon("Reports", theme.FileTextIcon(), nil)
	btnReport.OnTapped = func() {
		pos := fyne.CurrentApp().Driver().AbsolutePositionForObject(btnReport)
		pos.Y += btnReport.Size().Height
		widget.ShowPopUpMenuAtPosition(reportMenu, w.Canvas(), pos)
	}

	navBar := container.NewHBox(
		btnSetup,
		widget.NewButtonWithIcon("Ledger [F8]", theme.DocumentIcon(), func() {
			resetLedger()
			changePage(ledgerPage)
		}),
		widget.NewButtonWithIcon("Books [F9]", theme.FolderOpenIcon(), func() {
			resetBook()
			changePage(bookPage)
		}),
		widget.NewButtonWithIcon("Close Period [F10]", theme.FolderOpenIcon(), func() {
			showClosePeriodDialog(w)
		}),
		widget.NewButtonWithIcon("Select Period [F11]", theme.FolderOpenIcon(), func() {
			showSelectPeriodDialog(w)
		}),
		widget.NewButtonWithIcon("Inventory", theme.StorageIcon(), func() {
			inventoryPage := getEndingInventoryGUI(w)
			changePage(inventoryPage)
		}),
		btnReport,
		widget.NewButtonWithIcon("Audit Check", theme.SearchIcon(), func() {
			showAuditCheckDialog(w)
		}),
		layout.NewSpacer(),
		widget.NewButtonWithIcon("Disconnect", theme.LogoutIcon(), func() {
			currentDBPath = ""
			showConnectScreen(w, a)
		}),
	)

	// 4. Layout
	w.SetMainMenu(nil)
	mainLayout := container.NewBorder(
		container.NewVBox(navBar, widget.NewSeparator()),
		nil, nil, nil,
		contentContainer,
	)
	w.SetContent(mainLayout)

	ctrlN := &desktop.CustomShortcut{KeyName: fyne.KeyN, Modifier: fyne.KeyModifierControl}
	w.Canvas().AddShortcut(ctrlN, func(s fyne.Shortcut) {
		if currentPage == bookPage && bookAddFunc != nil {
			bookAddFunc()
		}
	})

	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		switch key.Name {
		case fyne.KeyF8:
			changePage(ledgerPage)
		case fyne.KeyF9:
			changePage(bookPage)
		case fyne.KeyF10:
			showClosePeriodDialog(w)
		case fyne.KeyF11:
			showSelectPeriodDialog(w)
		case fyne.KeyF12:
			showTrialBalanceDialog(w, func() { changePage(companyPage) })
		case fyne.KeyF3:
			if currentPage == ledgerPage && ledgerSearchFunc != nil {
				ledgerSearchFunc()
			} else if currentPage == bookPage && bookSearchFunc != nil {
				bookSearchFunc()
			}
		case fyne.KeyPageDown:
			if currentPage == ledgerPage && ledgerNextFunc != nil {
				ledgerNextFunc()
			} else if currentPage == bookPage && bookNextFunc != nil {
				bookNextFunc()
			}
		case fyne.KeyPageUp:
			if currentPage == ledgerPage && ledgerPrevFunc != nil {
				ledgerPrevFunc()
			} else if currentPage == bookPage && bookPrevFunc != nil {
				bookPrevFunc()
			}
		}
	})
}

func getHomeGUI() fyne.CanvasObject {
	makeRow := func(key, desc string) fyne.CanvasObject {
		return container.NewHBox(
			container.NewGridWrap(fyne.NewSize(110, 24),
				widget.NewLabelWithStyle(key, fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
			container.NewGridWrap(fyne.NewSize(260, 24),
				widget.NewLabel(desc)),
		)
	}

	content := container.NewVBox(
		widget.NewLabelWithStyle("General Ledger (Accounting System)  —  เชื่อมต่อฐานข้อมูลและพร้อมทำงานแล้ว",
			fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		widget.NewSeparator(),
		container.NewCenter(widget.NewLabelWithStyle("Keyboard Shortcuts", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})),
		widget.NewSeparator(),
		makeRow("F3", "ค้นหาเอกสาร (Book Search)"),
		makeRow("F8", "ผังบัญชี (Ledger)"),
		makeRow("F9", "สมุดบัญชีย่อย (Books)"),
		makeRow("F10", "ปิดงวด (Close Period)"),
		makeRow("F11", "เลือกงวด (Select Period)"),
		makeRow("F12", "งบทดลอง (Trial Balance)"),
		widget.NewSeparator(),
		makeRow("Ctrl+N", "เพิ่มเอกสารใหม่ (New Voucher)"),
		makeRow("PageDown", "เอกสารถัดไป"),
		makeRow("PageUp", "เอกสารก่อนหน้า"),
		widget.NewSeparator(),
		makeRow("F4", "ลบบรรทัด (Delete Line)"),
		makeRow("↑ ↓", "เลื่อนรายการใน Popup"),
		makeRow("Enter", "ยืนยันการค้นหา / บันทึก"),
		makeRow("Esc", "ปิด Popup / ยกเลิก"),
		widget.NewSeparator(),
	)

	return container.NewCenter(content)
}

func main() {
	// ── Crash Log: เขียน panic และ output ลงไฟล์ gl_crash.log ──
	logFile, logErr := os.OpenFile("gl_crash.log",
		os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if logErr == nil {
		fmt.Fprintf(logFile, "\n========== SESSION START: %s ==========\n",
			time.Now().Format("2006-01-02 15:04:05"))

		// crashLog = เขียนทั้ง stderr เดิม และ log file พร้อมกัน
		crashLog = io.MultiWriter(os.Stderr, logFile)

		// redirect stderr ผ่าน pipe → goroutine io.Copy
		// goroutine นี้ block ที่ io.Copy — ไม่กิน CPU ขณะ UI รันปกติ
		pr, pw, pipeErr := os.Pipe()
		if pipeErr == nil {
			os.Stderr = pw
			go func() {
				io.Copy(crashLog, pr)
			}()
		}

		defer logFile.Close()
	}
	// ถ้า logFile เปิดไม่ได้ → crashLog = os.Stderr (fallback ตั้งแต่ประกาศ global)

	// ── Panic Recovery: จับ panic ใน main goroutine แล้วเขียน log ──
	// ถ้า logFile เปิดไม่ได้ → writeLog() fallback ไปที่ os.Stderr อัตโนมัติ
	defer func() {
		if r := recover(); r != nil {
			writeLog(fmt.Sprintf("[%s] MAIN PANIC: %v\n--- stack trace ---\n%s\n",
				time.Now().Format("2006-01-02 15:04:05"), r, debug.Stack()))
			if logFile != nil {
				logFile.Sync()
			}
		}
	}()

	// 1. สร้าง App
	myApp := app.NewWithID("com.rinukz.gl")

	// 2. ตั้งค่า Theme
	myApp.Settings().SetTheme(&strongBlackTheme{Theme: theme.DefaultTheme()})

	// 3. สร้าง Window
	myWindow := myApp.NewWindow("General Ledger (Accounting System)")
	myWindow.Resize(fyne.NewSize(1200, 850))

	// 4. แสดงหน้า loading และตรวจ License
	checkingLabel := widget.NewLabelWithStyle(
		"🔐  กำลังตรวจสอบ Machine ID...",
		fyne.TextAlignCenter, fyne.TextStyle{},
	)
	myWindow.SetContent(container.NewCenter(checkingLabel))

	safeGo(func() {
		lic := CheckLicense()
		fyne.Do(func() {
			if !lic.Allowed {
				machineID := getMachineGUID()
				machineEntry := widget.NewEntry()
				machineEntry.SetText(machineID)
				machineEntry.Disable()
				btnCopy := widget.NewButton("📋 Copy Machine ID", func() {
					myWindow.Clipboard().SetContent(machineID)
				})
				btnCopy.Importance = widget.HighImportance
				myWindow.SetContent(container.NewCenter(
					container.NewVBox(
						widget.NewLabelWithStyle(
							"🔒  General Ledger (Accounting System)",
							fyne.TextAlignCenter,
							fyne.TextStyle{Bold: true},
						),
						widget.NewLabel(lic.Message),
						widget.NewLabel("Machine ID ของเครื่องนี้:"),
						machineEntry,
						btnCopy,
						widget.NewButton("ปิด", func() { myApp.Quit() }),
					),
				))
				return
			}
			showConnectScreen(myWindow, myApp)
		})
	})

	// 5. รัน!
	myWindow.ShowAndRun()
}
