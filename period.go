package main

// ==period.go==
// Circle Period System — เทียบเท่า GLPER ใน Fox Pro DOS
//
// ทำงานร่วมกับ book_ui.go โดยไม่ต้อง import
// Go compile ทุก .go ใน package main พร้อมกัน
//
// การใช้งานใน book_ui.go:
//   yearEnd, totalPeriods, nowPeriod := loadCompanyPeriod(xlOptions)
//   dr1, dr2, err := getCurrentPeriodRange(yearEnd, totalPeriods, nowPeriod)
//
// ─────────────────────────────────────────────────────────────────

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// PeriodInfo — เทียบเท่า 1 record ใน GLPER ของ Fox Pro
//
//	PNo    = PNO   (1-12)
//	PStart = PSAT  วันเริ่มงวด
//	PEnd   = PEND  วันสิ้นสุดงวด
//	PField = PFIELD เช่น "P01", "P02"
type PeriodInfo struct {
	PNo    int
	PStart time.Time
	PEnd   time.Time
	PField string
}

// CompanyPeriodConfig — ค่าที่อ่านจาก Company_Profile sheet
//
//	YearEnd      = ComYEnd    (col E2) วันสิ้นปีบัญชี เช่น 31/12/2025
//	TotalPeriods = ComPeriod  (col F2) จำนวนงวดทั้งหมด เช่น 12
//	NowPeriod    = ComNPeriod (col G2) งวดปัจจุบัน เช่น 3
type CompanyPeriodConfig struct {
	YearEnd      time.Time
	TotalPeriods int
	NowPeriod    int
}

// ─────────────────────────────────────────────────────────────────
// loadCompanyPeriod — อ่านค่า Period config จาก Company_Profile
//
// Company_Profile layout:
//
//	A2=ComCode  B2=ComName  C2=ComAddr  D2=ComTaxID
//	E2=ComYEnd  F2=ComPeriod  G2=ComNPeriod
//
// ─────────────────────────────────────────────────────────────────
func loadCompanyPeriod(xlOptions excelize.Options) (CompanyPeriodConfig, error) {
	cfg := CompanyPeriodConfig{}

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return cfg, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	yend, _ := f.GetCellValue("Company_Profile", "E2")
	period, _ := f.GetCellValue("Company_Profile", "F2")
	nperiod, _ := f.GetCellValue("Company_Profile", "G2")

	if yend == "" {
		return cfg, fmt.Errorf("ไม่พบ ComYEnd ใน Company_Profile!E2")
	}

	// parse วันสิ้นปีบัญชี — รองรับทั้ง dd/mm/yy และ dd/mm/yyyy
	cfg.YearEnd, err = time.Parse("02/01/06", yend)
	if err != nil {
		cfg.YearEnd, err = time.Parse("02/01/2006", yend)
		if err != nil {
			return cfg, fmt.Errorf("รูปแบบ ComYEnd ไม่ถูกต้อง: %s", yend)
		}
	}

	fmt.Sscanf(period, "%d", &cfg.TotalPeriods)
	fmt.Sscanf(nperiod, "%d", &cfg.NowPeriod)

	if cfg.TotalPeriods <= 0 {
		return cfg, fmt.Errorf("ComPeriod ไม่ถูกต้อง: %s", period)
	}
	if cfg.NowPeriod < 1 || cfg.NowPeriod > cfg.TotalPeriods {
		return cfg, fmt.Errorf("ComNPeriod (%d) ต้องอยู่ระหว่าง 1-%d", cfg.NowPeriod, cfg.TotalPeriods)
	}

	return cfg, nil
}

// ─────────────────────────────────────────────────────────────────
// calcAllPeriods — สร้าง period table ทั้งหมด เทียบเท่า GLPER
//
// Fox Pro logic:
//
//	date0  = GOMO(i_date+1, -12)  = ต้นปีบัญชี
//	psat0  = GOMO(date0, (I-1)*12/I_PERIODE)
//	pend0  = GOMO(date0, I*12/I_PERIODE) - 1
//
// ตัวอย่าง YearEnd=31/12/2025, TotalPeriods=12:
//
//	yearStart = 01/01/2025
//	Period 1  = 01/01/2025 → 31/01/2025
//	Period 12 = 01/12/2025 → 31/12/2025
//
// ตัวอย่าง YearEnd=31/03/2026, TotalPeriods=4:
//
//	yearStart = 01/04/2025
//	Period 1  = 01/04/2025 → 30/06/2025
//	Period 2  = 01/07/2025 → 30/09/2025
//	Period 3  = 01/10/2025 → 31/12/2025
//	Period 4  = 01/01/2026 → 31/03/2026
//
// ─────────────────────────────────────────────────────────────────
func calcAllPeriods(yearEnd time.Time, totalPeriods int) []PeriodInfo {
	if totalPeriods <= 0 || 12%totalPeriods != 0 {
		return nil // รองรับเฉพาะตัวหาร 12 ได้ลงตัว: 1,2,3,4,6,12
	}

	// GOMO(i_date+1, -12) = yearEnd + 1 day - 1 year = ต้นปีบัญชี
	yearStart := yearEnd.AddDate(0, 0, 1).AddDate(-1, 0, 0)

	monthsPerPeriod := 12 / totalPeriods
	periods := make([]PeriodInfo, totalPeriods)

	for i := 0; i < totalPeriods; i++ {
		// GOMO(date0, (I-1)*12/I_PERIODE)
		pStart := addMonths(yearStart, i*monthsPerPeriod)
		// GOMO(date0, I*12/I_PERIODE) - 1
		pEnd := addMonths(yearStart, (i+1)*monthsPerPeriod).AddDate(0, 0, -1)

		periods[i] = PeriodInfo{
			PNo:    i + 1,
			PStart: pStart,
			PEnd:   pEnd,
			PField: fmt.Sprintf("P%02d", i+1),
		}
	}
	return periods
}

// ─────────────────────────────────────────────────────────────────
// getCurrentPeriodRange — คืน DR1, DR2 ของงวดปัจจุบัน
// เทียบกับ Fox Pro: GOTO I_PEMNOW → DR1=PSAT, DR2=PEND
// ─────────────────────────────────────────────────────────────────
func getCurrentPeriodRange(yearEnd time.Time, totalPeriods, nowPeriod int) (dr1, dr2 time.Time, err error) {
	periods := calcAllPeriods(yearEnd, totalPeriods)
	if periods == nil {
		err = fmt.Errorf("ComPeriod (%d) ต้องเป็นตัวหาร 12 ได้ลงตัว (1,2,3,4,6,12)", totalPeriods)
		return
	}
	if nowPeriod < 1 || nowPeriod > len(periods) {
		err = fmt.Errorf("ComNPeriod (%d) ต้องอยู่ระหว่าง 1-%d", nowPeriod, totalPeriods)
		return
	}
	p := periods[nowPeriod-1]
	return p.PStart, p.PEnd, nil
}

// ─────────────────────────────────────────────────────────────────
// validateVoucherDate — ตรวจสอบ date ว่าอยู่ใน current period
// เทียบกับ Fox Pro: FUNC AE0 → IF BETWEEN(SUB(1,2), DR1, DR2)
//
// return: error message หรือ nil ถ้าผ่าน
// ─────────────────────────────────────────────────────────────────
func validateVoucherDate(dateStr string, dr1, dr2 time.Time) error {
	t, err := time.Parse("02/01/06", dateStr)
	if err != nil {
		t, err = time.Parse("02/01/2006", dateStr)
		if err != nil {
			return fmt.Errorf("รูปแบบวันที่ไม่ถูกต้อง (ต้องเป็น dd/mm/yy)")
		}
	}

	// BETWEEN(date, DR1, DR2)
	if t.Before(dr1) || t.After(dr2) {
		return fmt.Errorf(
			"วันที่อยู่นอกงวดปัจจุบัน\nRange: %s ถึง %s",
			dr1.Format("02/01/06"),
			dr2.Format("02/01/06"),
		)
	}
	return nil
}

// ─────────────────────────────────────────────────────────────────
// getPrevPeriodField — คืน field name ของงวดก่อนหน้า
// เทียบกับ Fox Pro: I_PEMBEF
//
//	NowPeriod=1 → "B12" (Beginning of last year's last period)
//	NowPeriod=3 → "P02"
//
// ใช้ใน POST ledger: GLAC.CLOS = GLAC.DR - GLAC.CR + &I_PEMBEF
// ─────────────────────────────────────────────────────────────────
func getPrevPeriodField(totalPeriods, nowPeriod int) string {
	if nowPeriod == 1 {
		// งวดแรก → previous = Beginning balance ของปีที่แล้ว
		return fmt.Sprintf("B%02d", totalPeriods)
	}
	return fmt.Sprintf("P%02d", nowPeriod-1)
}

// ─────────────────────────────────────────────────────────────────
// addMonths — เพิ่มเดือนโดยไม่ overflow วัน
// เทียบกับ FoxPro GOMO() function
//
// time.AddDate(0, n, 0) ใน Go จัดการ overflow อัตโนมัติ
// เช่น 31/01 + 1 month = 03/03 (ไม่ใช่ 28/02)
// ใช้วิธีนี้แทน: เอาวันที่ 1 ของเดือน แล้ว add month
// ─────────────────────────────────────────────────────────────────
func addMonths(t time.Time, months int) time.Time {
	// ใช้ first day of month เพื่อหลีกเลี่ยง overflow
	firstOfMonth := time.Date(t.Year(), t.Month(), 1, 0, 0, 0, 0, t.Location())
	return firstOfMonth.AddDate(0, months, 0)
}

// ─────────────────────────────────────────────────────────────────
// PeriodSummary — ใช้แสดงข้อมูล period ให้ user เห็น
// เรียกใช้ใน UI เพื่อแสดง "งวด 3/12 (01/06/25 - 30/06/25)"
// ─────────────────────────────────────────────────────────────────
func getPeriodSummary(cfg CompanyPeriodConfig) string {
	dr1, dr2, err := getCurrentPeriodRange(cfg.YearEnd, cfg.TotalPeriods, cfg.NowPeriod)
	if err != nil {
		return fmt.Sprintf("งวด %d/%d (ข้อมูลไม่สมบูรณ์)", cfg.NowPeriod, cfg.TotalPeriods)
	}
	return fmt.Sprintf("งวด %d/%d  (%s - %s)",
		cfg.NowPeriod,
		cfg.TotalPeriods,
		dr1.Format("02/01/06"),
		dr2.Format("02/01/06"),
	)
}
