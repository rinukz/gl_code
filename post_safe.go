// post_safe.go

package main

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────
// PostingLog — ข้อมูล Log สำหรับ recovery
// ─────────────────────────────────────────────────────────

type PostingLog struct {
	Timestamp            string
	Comcode              string
	Bitem                string
	Status               string // IN_PROGRESS, COMPLETED, RECOVERED, FAILED
	LastProcessedAccount string
	ErrorMsg             string
	LogRowNum            int
	Bperiod              int // ✅ เพิ่ม Bperiod
}

// ─────────────────────────────────────────────────────────
// initPostingLogSheet — สร้าง sheet Posting_Log ถ้ายังไม่มี
// ─────────────────────────────────────────────────────────
func initPostingLogSheet(f *excelize.File) error {
	for _, sheet := range f.GetSheetList() {
		if sheet == "Posting_Log" {
			// ✅ อัปเดต Header ให้ไฟล์เก่าที่ยังไม่มีคอลัมน์ Bperiod
			val, _ := f.GetCellValue("Posting_Log", "G1")
			if val == "" {
				f.SetCellValue("Posting_Log", "G1", "Bperiod")
			}
			return nil
		}
	}
	f.NewSheet("Posting_Log")
	headers := []string{
		"Timestamp", "Comcode", "Bitem", "Status",
		"LastProcessedAccount", "ErrorMsg", "Bperiod", // ✅ เพิ่ม Bperiod
	}
	for i, h := range headers {
		col, _ := excelize.ColumnNumberToName(i + 1)
		f.SetCellValue("Posting_Log", fmt.Sprintf("%s1", col), h)
	}
	return nil
}

// ─────────────────────────────────────────────────────────
// addPostingLog — เพิ่ม entry ใหม่ใน Posting_Log
// ─────────────────────────────────────────────────────────
func addPostingLog(f *excelize.File, timestamp, comcode, bitem, status, lastAc, errMsg string, bperiod int) (int, error) {
	if err := initPostingLogSheet(f); err != nil {
		return -1, err
	}
	rows, _ := f.GetRows("Posting_Log")
	newRow := len(rows) + 1
	vals := []interface{}{timestamp, comcode, bitem, status, lastAc, errMsg, bperiod} // ✅ บันทึก bperiod ลง Excel
	for j, v := range vals {
		col, _ := excelize.ColumnNumberToName(j + 1)
		f.SetCellValue("Posting_Log", fmt.Sprintf("%s%d", col, newRow), v)
	}
	return newRow, f.Save()
}

// ─────────────────────────────────────────────────────────
// updatePostingLog — อัปเดต Posting_Log entry
// ─────────────────────────────────────────────────────────
func updatePostingLog(f *excelize.File, logRow int, status, lastAc, errMsg string) error {
	f.SetCellValue("Posting_Log", fmt.Sprintf("D%d", logRow), status)
	f.SetCellValue("Posting_Log", fmt.Sprintf("E%d", logRow), lastAc)
	f.SetCellValue("Posting_Log", fmt.Sprintf("F%d", logRow), errMsg)
	return f.Save()
}

// ─────────────────────────────────────────────────────────
// calcCBAL — คำนวณ Closing Balance ตาม Fox Pro P_SUB.PRG
//
//	Fox Pro:
//	  IF CODE < '4'  → งบดุล
//	    CLOS = GLAC.DR - GLAC.CR + &PFF   (บวก BBAL)
//	  ELSE           → งบกำไรขาดทุน
//	    CLOS = GLAC.DR - GLAC.CR          (ไม่บวก BBAL)
//
//	หมายเหตุ: DR, CR ที่ใส่มาคือค่า **หลัง** อัพเดทแล้ว (newDebit, newCredit)
//
// ─────────────────────────────────────────────────────────
func calcCBAL(acCode string, newDebit, newCredit, bbal float64) float64 {
	if acCode < "4" {
		// งบดุล: Assets / Liabilities / Equity
		return newDebit - newCredit + bbal
	}
	// งบกำไรขาดทุน: Revenue / Expense
	return newDebit - newCredit
}

// ─────────────────────────────────────────────────────────
// collectLines — จัดกลุ่ม BookLine ตาม AcCode
// includeVAT = true  → รวม VAT line (ใช้ตอน POST)
// includeVAT = false → ข้าม VAT line (ใช้ตอน UNPOST แบบ selective)
//
// ✅ Fox Pro loop ทุก line ใน GLDAT โดยไม่แยก VAT
//
//	ดังนั้น includeVAT = true เสมอ ทั้ง POST และ UNPOST
//
// ─────────────────────────────────────────────────────────
func collectLines(bookLines []BookLine) map[string][]BookLine {
	grouped := make(map[string][]BookLine)
	for _, line := range bookLines {
		grouped[line.AcCode] = append(grouped[line.AcCode], line)
	}
	return grouped
}

// ─────────────────────────────────────────────────────────
// postAccountsToLedger — POST หรือ UNPOST แต่ละ account
//
//	sign = +1.0 → POST   (ADD)
//	sign = -1.0 → UNPOST (RECALL ก่อน EDIT/DELETE)
//
// ─────────────────────────────────────────────────────────
func postAccountsToLedger(f *excelize.File, grouped map[string][]BookLine, comcode string, sign float64) error {
	ledgerRows, _ := f.GetRows("Ledger_Master")

	for acCode, lines := range grouped {
		rowNum := findLedgerRowNum(ledgerRows, comcode, acCode)
		if rowNum < 0 {
			continue // ไม่พบ account → ข้าม
		}

		var totalDebit, totalCredit float64
		for _, line := range lines {
			totalDebit += parseFloat(line.Bdebit)
			totalCredit += parseFloat(line.Bcredit)
		}

		oldDebit := parseFloat(safeGet(ledgerRows[rowNum], 7))
		oldCredit := parseFloat(safeGet(ledgerRows[rowNum], 8))
		bbal := parseFloat(safeGet(ledgerRows[rowNum], 5))

		newDebit := oldDebit + sign*totalDebit
		newCredit := oldCredit + sign*totalCredit
		cbal := calcCBAL(acCode, newDebit, newCredit, bbal)

		excelRow := rowNum + 1
		f.SetCellValue("Ledger_Master", fmt.Sprintf("H%d", excelRow), newDebit)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("I%d", excelRow), newCredit)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("G%d", excelRow), cbal)
		// ✅ ไม่ Save ใน loop — Save ครั้งเดียวหลังจบทุก account
	}

	// Save ครั้งเดียวหลัง update ทุก account เสร็จ
	return f.Save()
}

// ─────────────────────────────────────────────────────────
// safePostVoucher — POST voucher → Ledger (Fox Pro S_ADD)
//
//	Flow:
//	  1. initPostingLogSheet
//	  2. markBitemAsPosted(false) → Posted=0  [checkpoint 1]
//	  3. addPostingLog(IN_PROGRESS)
//	  4. postAccountsToLedger(sign=+1)
//	  5. updatePostingLog(COMPLETED)           [checkpoint 2]
//	  6. markBitemAsPosted(true)  → Posted=1  [checkpoint 3]
//
// ─────────────────────────────────────────────────────────

func safePostVoucher(bitem, comcode string, bperiod int, f *excelize.File, xlOptions excelize.Options) error {
	if err := initPostingLogSheet(f); err != nil {
		return fmt.Errorf("init posting log: %v", err)
	}

	// Checkpoint 1: mark Posted=0 ก่อน
	if err := markBitemAsPosted(f, bitem, comcode, bperiod, false, xlOptions); err != nil {
		return err
	}

	// ✅ ส่ง bperiod เข้าไปใน addPostingLog ด้วย
	logRow, err := addPostingLog(f,
		time.Now().Format("2006-01-02 15:04:05"),
		comcode, bitem, "IN_PROGRESS", "", "", bperiod)
	if err != nil {
		return fmt.Errorf("add posting log: %v", err)
	}

	// ดึง lines และจัดกลุ่มตาม AcCode
	bookLines := getBookLines(bitem, comcode, bperiod, f, xlOptions)
	grouped := collectLines(bookLines)

	// POST ทุก account (sign=+1)
	if err := postAccountsToLedger(f, grouped, comcode, +1.0); err != nil {
		_ = updatePostingLog(f, logRow, "FAILED", "", err.Error())
		return err
	}

	// Checkpoint 2: COMPLETED
	if err := updatePostingLog(f, logRow, "COMPLETED", "", ""); err != nil {
		return fmt.Errorf("update posting log COMPLETED: %v", err)
	}

	// Checkpoint 3: mark Posted=1
	if err := markBitemAsPosted(f, bitem, comcode, bperiod, true, xlOptions); err != nil {
		return err
	}
	return f.Save()
}

// ─────────────────────────────────────────────────────────
// unpostVoucherFromLedger — UNPOST voucher (Fox Pro S_EDIT RECALL / S_DELE)
//
//	เรียกก่อน EDIT หรือ DELETE เสมอ
//	sign = -1 → ลบยอดออกจาก Ledger
//	✅ รวม VAT line ด้วย (สมมาตรกับ safePostVoucher)
//
// ─────────────────────────────────────────────────────────
func unpostVoucherFromLedger(bitem, comcode string, bperiod int, f *excelize.File, xlOptions excelize.Options) error {
	if !isBitemPosted(bitem, comcode, bperiod, f, xlOptions) {
		return nil
	}

	bookLines := getBookLines(bitem, comcode, bperiod, f, xlOptions)
	grouped := collectLines(bookLines)

	if err := postAccountsToLedger(f, grouped, comcode, -1.0); err != nil {
		return fmt.Errorf("unpost ledger: %v", err)
	}

	if err := markBitemAsPosted(f, bitem, comcode, bperiod, false, xlOptions); err != nil {
		return err
	}
	return f.Save()
}

// isBitemPosted — เช็คว่า bitem นี้ Posted=1 หรือยัง
func isBitemPosted(bitem, comcode string, bperiod int, f *excelize.File, xlOptions excelize.Options) bool {
	rows, _ := f.GetRows("Book_items")

	cfg, err := loadCompanyPeriod(xlOptions)
	var dr1, dr2 time.Time
	if err == nil {
		dr1, dr2, _ = getCurrentPeriodRange(cfg.YearEnd, cfg.TotalPeriods, bperiod)
	}

	for i, row := range rows {
		if i == 0 || len(row) < 4 || row[0] != comcode || row[3] != bitem {
			continue
		}
		var rowPeriod int
		if len(row) > 21 {
			fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
		}

		isInPeriod := false
		if err != nil {
			isInPeriod = (rowPeriod == 0 || rowPeriod == bperiod)
		} else {
			rowDate := safeGet(row, 1)
			t, dateErr := time.Parse("02/01/06", rowDate)
			isInPeriod = (rowPeriod != 0 && rowPeriod == bperiod) ||
				(rowPeriod == 0 && dateErr == nil && !t.Before(dr1) && !t.After(dr2))
		}

		if isInPeriod {
			return safeGet(row, 19) == "1"
		}
	}
	return false
}

// ─────────────────────────────────────────────────────────
// recoverPosting — Scan Posting_Log หา IN_PROGRESS แล้ว repair
//
//	Strategy: unpost ทั้งหมดก่อน แล้ว repost ใหม่
//	(ปลอดภัยกว่า resume จาก lastAcCode เพราะ map ไม่มี order)
//
// ─────────────────────────────────────────────────────────

func recoverPosting(f *excelize.File, xlOptions excelize.Options) error {
	rows, _ := f.GetRows("Posting_Log")
	var needRecovery []PostingLog

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 4 && safeGet(row, 3) == "IN_PROGRESS" {
			bperiod := 0
			if len(row) >= 7 {
				fmt.Sscanf(safeGet(row, 6), "%d", &bperiod) // ✅ อ่าน Bperiod จาก Log
			}
			needRecovery = append(needRecovery, PostingLog{
				Timestamp:            safeGet(row, 0),
				Comcode:              safeGet(row, 1),
				Bitem:                safeGet(row, 2),
				Status:               safeGet(row, 3),
				LastProcessedAccount: safeGet(row, 4),
				ErrorMsg:             safeGet(row, 5),
				LogRowNum:            i + 1,
				Bperiod:              bperiod, // ✅ เก็บลง Struct
			})
		}
	}

	if len(needRecovery) == 0 {
		return nil
	}

	for _, log := range needRecovery {
		_ = unpostVoucherFromLedger(log.Bitem, log.Comcode, log.Bperiod, f, xlOptions)

		// ✅ Step 2: repost ใหม่ทั้งหมด (ส่ง log.Bperiod เข้าไปด้วย)
		if err := safePostVoucher(log.Bitem, log.Comcode, log.Bperiod, f, xlOptions); err != nil {
			errMsg := fmt.Sprintf("recovery failed: %v", err)
			_ = updatePostingLog(f, log.LogRowNum, "FAILED", "", errMsg)
			return fmt.Errorf("%s", errMsg)
		}

		// Mark RECOVERED
		recoverMsg := "Recovered at " + time.Now().Format("2006-01-02 15:04:05")
		if err := updatePostingLog(f, log.LogRowNum, "RECOVERED", "", recoverMsg); err != nil {
			return err
		}

		fmt.Printf("[RECOVERY] Done: %s/%s\n", log.Comcode, log.Bitem)
	}

	return f.Save()
}

// ─────────────────────────────────────────────────────────
// Helper functions
// ─────────────────────────────────────────────────────────

// findLedgerRowNum — หา row index ใน Ledger_Master (0-based)
func findLedgerRowNum(rows [][]string, comcode, acCode string) int {
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 2 && row[0] == comcode && row[1] == acCode {
			return i
		}
	}
	return -1
}

// markBitemAsPosted — set Posted col T ใน Book_items
func markBitemAsPosted(f *excelize.File, bitem, comcode string, bperiod int, posted bool, xlOptions excelize.Options) error {
	sheet := "Book_items"
	rows, _ := f.GetRows(sheet)

	cfg, err := loadCompanyPeriod(xlOptions)
	var dr1, dr2 time.Time
	if err == nil {
		dr1, dr2, _ = getCurrentPeriodRange(cfg.YearEnd, cfg.TotalPeriods, bperiod)
	}

	for i, row := range rows {
		if i == 0 || len(row) < 4 || row[0] != comcode || row[3] != bitem {
			continue
		}
		var rowPeriod int
		if len(row) > 21 {
			fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
		}

		isInPeriod := false
		if err != nil {
			isInPeriod = (rowPeriod == 0 || rowPeriod == bperiod)
		} else {
			rowDate := safeGet(row, 1)
			t, dateErr := time.Parse("02/01/06", rowDate)
			isInPeriod = (rowPeriod != 0 && rowPeriod == bperiod) ||
				(rowPeriod == 0 && dateErr == nil && !t.Before(dr1) && !t.After(dr2))
		}

		if isInPeriod {
			val := "0"
			if posted {
				val = "1"
			}
			f.SetCellValue(sheet, fmt.Sprintf("T%d", i+1), val)
		}
	}
	return nil
}

// getBookLines — ดึง BookLine ทั้งหมดของ bitem นี้
func getBookLines(bitem, comcode string, bperiod int, f *excelize.File, xlOptions excelize.Options) []BookLine {
	var lines []BookLine
	rows, _ := f.GetRows("Book_items")

	cfg, err := loadCompanyPeriod(xlOptions)
	var dr1, dr2 time.Time
	if err == nil {
		dr1, dr2, _ = getCurrentPeriodRange(cfg.YearEnd, cfg.TotalPeriods, bperiod)
	}

	for i, row := range rows {
		if i == 0 || len(row) < 4 || row[0] != comcode || row[3] != bitem {
			continue
		}
		var rowPeriod int
		if len(row) > 21 {
			fmt.Sscanf(safeGet(row, 21), "%d", &rowPeriod)
		}

		isInPeriod := false
		if err != nil {
			isInPeriod = (rowPeriod == 0 || rowPeriod == bperiod)
		} else {
			rowDate := safeGet(row, 1)
			t, dateErr := time.Parse("02/01/06", rowDate)
			isInPeriod = (rowPeriod != 0 && rowPeriod == bperiod) ||
				(rowPeriod == 0 && dateErr == nil && !t.Before(dr1) && !t.After(dr2))
		}

		if isInPeriod {
			bl := BookLine{
				Comcode: row[0], Bdate: safeGet(row, 1), Bvoucher: safeGet(row, 2),
				Bitem: row[3], AcCode: safeGet(row, 5), AcName: safeGet(row, 6),
				Scode: safeGet(row, 7), Sname: safeGet(row, 8),
				Bdebit: safeGet(row, 9), Bcredit: safeGet(row, 10),
				Bref: safeGet(row, 11), Boff: safeGet(row, 12),
				Bcomtaxid: safeGet(row, 13), Bnote: safeGet(row, 14),
				Bchqno: safeGet(row, 15), Bchqdate: safeGet(row, 16),
				Bnote2: safeGet(row, 20), IsVATLine: safeGet(row, 17) == "1",
			}
			var parentBline int
			fmt.Sscanf(safeGet(row, 18), "%d", &parentBline)
			bl.ParentBline = parentBline

			var blineNum int
			fmt.Sscanf(safeGet(row, 4), "%d", &blineNum)
			bl.Bline = blineNum
			bl.Bperiod = rowPeriod

			lines = append(lines, bl)
		}
	}
	return lines
}
