package main

import (
	"image/color" // ✅ เพิ่มตัวนี้เพื่อกำหนดสี Black

	"fyne.io/fyne/v2"

	"fyne.io/fyne/v2/theme"
)

// ==globals.go
// ประกาศตัวแปร Global ไว้ที่นี่ เพื่อให้ทุกไฟล์ใน package main เรียกใช้ได้
var (
	currentDBPath      string
	currentCompanyCode string
	// refreshLedgerFunc — เรียกจาก book_ui เมื่อ POST สำเร็จ
	refreshLedgerFunc func()

	// refreshCompanySetupFunc — เรียกหลัง Close Period เพื่ออัพเดท enComNPeriod
	refreshCompanySetupFunc func()

	// bookNextFunc / bookPrevFunc — PageDown/PageUp nav ใน book_ui
	bookNextFunc func()
	// bookAddFunc — Ctrl+N add new voucher ใน book_ui
	bookAddFunc  func()
	bookPrevFunc func()

	// ledgerNextFunc / ledgerPrevFunc — PageDown/PageUp nav ใน ledger_ui
	ledgerNextFunc func()
	ledgerPrevFunc func()

	// ledgerSearchFunc / bookSearchFunc — F3 search ใน ledger_ui / book_ui
	ledgerSearchFunc func()
	bookSearchFunc   func()

	// closeAcSearchPopup — ปิด AC code search popup ก่อน Ctrl+S save (book_ui)
	closeAcSearchPopup func()
	bookSaveFunc       func()

	// bookPage — reference หน้า Book UI (ใช้ changePage จาก inventory)
	bookPage   fyne.CanvasObject
	changePage func(fyne.CanvasObject)
	resetBook  func()

	// inventoryPage — reference หน้า Inventory UI
	inventoryPage       fyne.CanvasObject
	inventorySubmitFunc func() // Ctrl+S = submit form เพิ่มรายการสินค้า
	inventorySearchFunc func() // F3 = เปิด popup ค้นหา/แก้ไขรายการ

	// ledgerResetFunc — reset ledger UI
	ledgerResetFunc func()

	// ─────────────────────────────────────────────────────────────────
	// workingPeriod — งวดที่กำลังทำงานอยู่ใน session นี้ (in-memory only)
	//
	// แยกออกจาก NowPeriod (Company_Profile G2) ซึ่งคือ "งวดบัญชีจริง"
	// ที่ผ่านการ Close Period มาแล้ว และห้ามแก้ด้วย Select Period
	//
	// workingPeriod = 0  → ยังไม่ได้ set (ใช้ NowPeriod จาก Company_Profile แทน)
	// workingPeriod = N  → user เลือก period N ผ่าน Select Period dialog
	//
	// การ reset: เมื่อ Close Period สำเร็จ → workingPeriod = 0 (กลับใช้ NowPeriod ใหม่)
	// ─────────────────────────────────────────────────────────────────
	workingPeriod int // 0 = ใช้ NowPeriod จาก Company_Profile

	// setWorkingPeriodFunc — callback จาก book_ui เพื่อรับ workingPeriod ใหม่
	// book_ui จะ reload period range และ refresh UI ให้อัตโนมัติ
	setWorkingPeriodFunc func(period int)

	// bookGotoFunc — navigate ไปยัง Bitem ที่ระบุ (ใช้จาก inventory Auto-JV)
	bookGotoFunc func(bitem string, bperiod int) // ✅ เพิ่ม bperiod int

	// ═══════════════════════════════════════════════════════════
	// Global Variables - UI References
	// ═══════════════════════════════════════════════════════════

	currentWindow fyne.Window
)

type strongBlackTheme struct {
	fyne.Theme
}

// 1. ปรับสีตัวหนังสือให้ดำสนิท (แก้ Error 'not used' ไปในตัว)
func (m *strongBlackTheme) Color(name fyne.ThemeColorName, variant fyne.ThemeVariant) color.Color {
	if name == theme.ColorNameForeground {
		return color.Black // ตัวหนังสือตอนพิมพ์อยู่ ให้ดำสนิทชัดเจน
	}
	if name == theme.ColorNameDisabled {
		// เปลี่ยนจาก color.Black เป็นสีเทาเข้มแทน จะได้ไม่พร่างตา
		return color.RGBA{R: 120, G: 120, B: 120, A: 255}
	}
	return m.Theme.Color(name, variant)
}

// 2. ⚡ นี่คือจุดที่ปรับ "โดยรวม" ให้เล็กลงครับ
func (m *strongBlackTheme) Size(name fyne.ThemeSizeName) float32 {
	if name == theme.SizeNameText {
		return 12 // 👈 ปรับตรงนี้ที่เดียว ตัวหนังสือใน Body, Entry, Table จะเล็กลงหมดเลยครับ
	}

	// แถม: ถ้าอยากให้ช่องกรอก (Entry) ดูเตี้ยลง ไม่กินที่ ให้ปรับตรงนี้เพิ่มครับ
	if name == theme.SizeNamePadding {
		return 2 // ลดระยะห่างรอบๆ ตัวหนังสือ
	}

	return m.Theme.Size(name)
}
