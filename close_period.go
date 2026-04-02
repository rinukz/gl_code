package main

// close_period.go - CLOSE + OPEN รวมกัน
// Special codes ดึงจาก Special_code sheet จริง

import (
	"fmt"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

const (
	keyVAT1    = "VAT1"
	keyVAT2    = "VAT2"
	keyTVAT    = "TVAT"
	keyPROFIT  = "PROFIT"
	keyPROFITS = "PROFITS"
)

type ClosePeriodResult struct {
	ClosedPeriod int
	IsYearEnd    bool
	NewPeriod    int
	NewYearEnd   time.Time
	ProfitLoss   float64
}

func closePeriod(xlOptions excelize.Options) (*ClosePeriodResult, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, fmt.Errorf("โหลด Period config ไม่ได้: %v", err)
	}
	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}

	nowPeriod := cfg.NowPeriod
	totalPeriods := cfg.TotalPeriods
	isYearEnd := nowPeriod == totalPeriods
	periodField := fmt.Sprintf("P%02d", nowPeriod)
	comCode := getComCodeFromExcel(xlOptions)
	sc := loadSpecialCodeMap(f, comCode)
	ledgerRows, _ := f.GetRows("Ledger_Master")

	findRow := func(acCode string) int {
		for i, row := range ledgerRows {
			if i == 0 {
				continue
			}
			if len(row) >= 2 && row[0] == comCode && row[1] == acCode {
				return i
			}
		}
		return -1
	}
	getClos := func(acCode string) float64 {
		i := findRow(acCode)
		if i < 0 {
			return 0
		}
		return parseFloat(safeGet(ledgerRows[i], 6))
	}
	setCell := func(rowIdx int, colOneBased int, val interface{}) {
		colName, _ := excelize.ColumnNumberToName(colOneBased)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", colName, rowIdx+1), val)
	}
	// P01=col12(idx11), P02=col13(idx12), Bxx=Pxx+totalPeriods
	pColOf := func(pField string) int {
		var pNo int
		fmt.Sscanf(pField, "P%d", &pNo)
		return 11 + pNo // 0-based
	}

	// Step 1: ปิด VAT
	tvatClos := getClos(sc[keyTVAT])
	if r := findRow(sc[keyTVAT]); r >= 0 {
		setCell(r, 7, 0.0)
	}
	if tvatClos > 0 {
		if r := findRow(sc[keyVAT1]); r >= 0 {
			setCell(r, 7, getClos(sc[keyVAT1])+tvatClos)
		}
	} else if tvatClos < 0 {
		if r := findRow(sc[keyVAT2]); r >= 0 {
			setCell(r, 7, getClos(sc[keyVAT2])+tvatClos)
		}
	}
	if err := f.Save(); err != nil {
		f.Close()
		return nil, fmt.Errorf("Step1: %v", err)
	}
	ledgerRows, _ = f.GetRows("Ledger_Master")

	// Step 2: CLOS → Period Field
	pColIdx := pColOf(periodField)
	for i, row := range ledgerRows {
		if i == 0 || len(row) < 2 || row[0] != comCode {
			continue
		}
		clos := parseFloat(safeGet(row, 6))
		cn, _ := excelize.ColumnNumberToName(pColIdx + 1)
		f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", cn, i+1), clos)
	}
	if err := f.Save(); err != nil {
		f.Close()
		return nil, fmt.Errorf("Step2: %v", err)
	}
	ledgerRows, _ = f.GetRows("Ledger_Master")

	// Step 3: P&L
	prevColIdx := pColOf(getPrevPeriodField(totalPeriods, nowPeriod))
	var x1, x2, x3 float64
	for i, row := range ledgerRows {
		if i == 0 || len(row) < 2 || row[0] != comCode {
			continue
		}
		acCode := safeGet(row, 1)
		if acCode >= "4" {
			x1 += parseFloat(safeGet(row, 6))
		}
		if acCode >= "117" && acCode <= "119" {
			x2 += parseFloat(safeGet(row, prevColIdx))
			x3 += parseFloat(safeGet(row, 6))
		}
	}
	x4 := x1 + x2 - x3
	profitRow := findRow(sc[keyPROFIT])
	if profitRow >= 0 {
		cn, _ := excelize.ColumnNumberToName(pColIdx + 1)
		old := parseFloat(safeGet(ledgerRows[profitRow], pColIdx))
		f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", cn, profitRow+1), old+x4)
	}

	// Step 4: ปิดปี → โอน PROFIT → PROFITS
	x5 := x4
	if profitRow >= 0 {
		x5 = parseFloat(safeGet(ledgerRows[profitRow], pColIdx)) + x4
	}
	if isYearEnd {
		if profitRow >= 0 {
			setCell(profitRow, 7, 0.0)
			cn, _ := excelize.ColumnNumberToName(pColIdx + 1)
			f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", cn, profitRow+1), 0.0)
		}
		if r := findRow(sc[keyPROFITS]); r >= 0 {
			cn, _ := excelize.ColumnNumberToName(pColIdx + 1)
			old := parseFloat(safeGet(ledgerRows[r], pColIdx))
			f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", cn, r+1), old+x5)
		}
	}
	if err := f.Save(); err != nil {
		f.Close()
		return nil, fmt.Errorf("Step3/4: %v", err)
	}
	ledgerRows, _ = f.GetRows("Ledger_Master")

	// Step 5: เปิดงวดใหม่ (P_OPEN)
	for i, row := range ledgerRows {
		if i == 0 || len(row) < 2 || row[0] != comCode {
			continue
		}
		acCode := safeGet(row, 1)
		excelRow := i + 1
		f.SetCellValue("Ledger_Master", fmt.Sprintf("H%d", excelRow), 0.0) // DR=0
		f.SetCellValue("Ledger_Master", fmt.Sprintf("I%d", excelRow), 0.0) // CR=0
		if acCode >= "4" {
			f.SetCellValue("Ledger_Master", fmt.Sprintf("G%d", excelRow), 0.0) // P&L CLOS=0
		} else {
			f.SetCellValue("Ledger_Master", fmt.Sprintf("G%d", excelRow), parseFloat(safeGet(row, pColIdx))) // งบดุล ยกยอด
		}
		if isYearEnd {
			for p := 1; p <= totalPeriods; p++ {
				tIdx := pColOf(fmt.Sprintf("P%02d", p))
				bIdx := tIdx + totalPeriods
				val := parseFloat(safeGet(row, tIdx))
				bcn, _ := excelize.ColumnNumberToName(bIdx + 1)
				tcn, _ := excelize.ColumnNumberToName(tIdx + 1)
				f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", bcn, excelRow), val) // Bxx=Pxx
				f.SetCellValue("Ledger_Master", fmt.Sprintf("%s%d", tcn, excelRow), 0.0) // Pxx=0
			}
			bbal := parseFloat(safeGet(row, pColIdx))
			if acCode >= "4" {
				bbal = 0
			}
			f.SetCellValue("Ledger_Master", fmt.Sprintf("F%d", excelRow), bbal) // BBal

			// Bthisyear (col J = index 9) = ยอดยกมาต้นปีใหม่
			// เทียบ FoxPro P_OPEN: REPL ALL BEGI WITH &N  (N = B12 = Lastper12)
			// งบดุล (code < "4") → ยกยอดจาก Lastper12 (B12 ที่เพิ่งถูก set = val ของ P12)
			// P&L (code >= "4") → reset เป็น 0
			bthisYear := bbal                                                        // งบดุล: ใช้ค่าเดียวกับ BBal (= P12 ณ ปิดปี)
			f.SetCellValue("Ledger_Master", fmt.Sprintf("J%d", excelRow), bthisYear) // Bthisyear
		}
	}
	if err := f.Save(); err != nil {
		f.Close()
		return nil, fmt.Errorf("Step5: %v", err)
	}

	// Step 6: เลื่อน ComNPeriod
	newPeriod := nowPeriod + 1
	newYearEnd := cfg.YearEnd
	if isYearEnd {
		newPeriod = 1
		newYearEnd = cfg.YearEnd.AddDate(1, 0, 0)
		f.SetCellValue("Company_Profile", "E2", newYearEnd.Format("02/01/06"))
	}
	f.SetCellValue("Company_Profile", "G2", newPeriod)
	if err := f.Save(); err != nil {
		f.Close()
		return nil, fmt.Errorf("Step6: %v", err)
	}
	f.Close()

	return &ClosePeriodResult{
		ClosedPeriod: nowPeriod, IsYearEnd: isYearEnd,
		NewPeriod: newPeriod, NewYearEnd: newYearEnd, ProfitLoss: x4,
	}, nil
}

// showClosePeriodDialog — dialog ปิดงวด (Enter=ยืนยัน, Esc=ยกเลิก)
func showClosePeriodDialog(w fyne.Window) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}
	dr1, dr2, _ := getCurrentPeriodRange(cfg.YearEnd, cfg.TotalPeriods, cfg.NowPeriod)
	suffix := ""
	if cfg.NowPeriod == cfg.TotalPeriods {
		suffix = "\n⚠️  งวดสุดท้าย — จะปิดปีบัญชีด้วย"
	}
	periodInfo := fmt.Sprintf("งวด %d / %d   (%s — %s)%s",
		cfg.NowPeriod, cfg.TotalPeriods,
		dr1.Format("02/01/06"), dr2.Format("02/01/06"), suffix)

	var d dialog.Dialog

	// btnOK รับ Enter=ยืนยัน, Esc=ยกเลิก (focus อยู่ที่นี่เสมอ)
	btnOK := newEnterEscButton("ยืนยัน (Enter)", func() {
		// Enter
		// Enter
		d.Hide()
		go func() {
			result, err := closePeriod(xlOptions)
			fyne.Do(func() {
				if err != nil {
					showErrDialog(w, "❌ "+err.Error())
					return
				}

				workingPeriod = 0 // ✅ เพิ่มบรรทัดนี้: รีเซ็ตกลับเป็นงวดปัจจุบันเฉพาะตอนปิดงวดสำเร็จ

				sign := "กำไร"
				if result.ProfitLoss < 0 {
					sign = "ขาดทุน"
				}
				msg := fmt.Sprintf("✅ ปิดงวด %d สำเร็จ\n%s: %.2f บาท\nงวดถัดไป: %d",
					result.ClosedPeriod, sign, result.ProfitLoss, result.NewPeriod)

				if result.IsYearEnd {
					msg += fmt.Sprintf("\n🎉 ปิดปีบัญชี — ปีใหม่สิ้นสุด: %s",
						result.NewYearEnd.Format("02/01/06"))
				}
				var done dialog.Dialog
				okBtn := newEnterButton("OK", func() { done.Hide() })
				done = dialog.NewCustomWithoutButtons("สำเร็จ", container.NewVBox(
					widget.NewLabel(msg),
					container.NewCenter(okBtn),
				), w)
				done.Show()
				go func() {
					time.Sleep(50 * time.Millisecond)
					fyne.Do(func() { w.Canvas().Focus(okBtn) })
				}()
				if refreshLedgerFunc != nil {
					refreshLedgerFunc()
				}
				if refreshCompanySetupFunc != nil {
					refreshCompanySetupFunc()
				}
			})
		}()
	}, func() {
		// Esc
		d.Hide()
	})
	btnOK.Importance = widget.DangerImportance
	btnCancel := widget.NewButton("ยกเลิก (Esc)", func() { d.Hide() })

	d = dialog.NewCustomWithoutButtons("Close Period", container.NewVBox(
		widget.NewLabelWithStyle("Close Period", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		widget.NewLabelWithStyle(periodInfo, fyne.TextAlignCenter, fyne.TextStyle{}),
		widget.NewSeparator(),
		widget.NewLabelWithStyle(
			"⚠️  ก่อนปิดงวดบัญชี โปรดตรวจสอบข้อมูลให้เรียบร้อย",
			fyne.TextAlignCenter, fyne.TextStyle{Italic: true},
		),
		widget.NewSeparator(),
		container.NewCenter(container.NewHBox(btnOK, btnCancel)),
	), w)
	d.Show()
	go func() {
		time.Sleep(50 * time.Millisecond)
		fyne.Do(func() { w.Canvas().Focus(btnOK) })
	}()
}

func loadSpecialCodeMap(f *excelize.File, comCode string) map[string]string {
	// default fallback
	m := map[string]string{
		keyVAT1: "120VAT", keyVAT2: "235VAT", keyTVAT: "235TVAT",
		keyPROFIT: "360PLA", keyPROFITS: "350RTE",
	}
	rows, _ := f.GetRows("Special_code")
	// รวบรวม CODE ของ comCode นี้ตามลำดับ
	var data []string
	for i, row := range rows {
		if i == 0 || len(row) < 2 || row[0] != comCode {
			continue
		}
		data = append(data, strings.TrimSpace(row[1]))
	}
	// map ตาม index ลำดับเดียวกับ requiredData ใน VerifySpecialAccounts
	keys := []string{keyVAT1, keyTVAT, keyVAT2, keyPROFITS, keyPROFIT}
	for i, key := range keys {
		if i < len(data) && data[i] != "" {
			m[key] = data[i]
		}
	}
	return m
}
