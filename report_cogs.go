package main

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// Data Structure สำหรับ Cost of Goods Sold (CGS)
//
// การ map รหัสบัญชี (ตาม Acct_Group ใน helpers.go):
//
//	500-504 = ซื้อวัตถุดิบ / ค่าใช้จ่ายซื้อ / ตัวหักซื้อ
//	506     = ค่าแรงงานทางตรง
//	508     = โสหุ้ยอุปกรณ์โรงงาน  ← ค่าโสหุ้ยในการผลิต
//	510     = ซื้อสินค้าสำเร็จรูป   ← "ซื้อสินค้า"
//	512     = ส่วนลดรับในการซื้อสินค้า ← "ส่งคืนและส่วนลดรับ"
//	514     = ค่าใช้จ่ายในการซื้อสินค้า (รวมใน ซื้อสินค้า)
//	516     = ต้นทุนขายอื่น
//	600     = ปรับปรุงสินค้า
//
// Inventory accounts:
//
//	117 = สินค้าสำเร็จรูป (Finished Goods)
//	118 = งานระหว่างทาง  (Work in Process)
//	119 = วัตถุดิบ        (Raw Materials)
//
// หมายเหตุ: "ปลายงวดยกไป" = ยอดคงเหลือสุทธิ ณ วันสิ้นงวด
//
//	= Bthisyear + SUM(ทุก tx จนถึง dateTo)
//	ไม่ใช่ begBal + txInPeriod
//
// ─────────────────────────────────────────────────────────────────
type CGSData struct {
	ComName    string
	PeriodFrom int
	PeriodTo   int
	DateFrom   time.Time
	DateTo     time.Time
	PrintTime  time.Time

	// Current Year (ปีปัจจุบัน)
	CY_Prev119 float64 // วัตถุดิบต้นงวดยกมา           (119 ต้นงวด)
	CY_Acc500  float64 // ซื้อวัตถุดิบระหว่างงวด        (500-504)
	CY_End119  float64 // วัตถุดิบปลายงวดยกไป           (119 ปลายงวด)
	CY_Acc506  float64 // ค่าแรงทางตรง                  (506)
	CY_Acc508  float64 // ค่าโสหุ้ยในการผลิต             (508)
	CY_Prev118 float64 // งานระหว่างทำต้นงวดยกมา        (118 ต้นงวด)
	CY_End118  float64 // งานระหว่างทำปลายงวดยกไป       (118 ปลายงวด)
	CY_Prev117 float64 // สินค้าสำเร็จรูปต้นงวดยกมา     (117 ต้นงวด)
	CY_Acc510  float64 // ซื้อสินค้า                     (510+514)
	CY_Acc512  float64 // ส่งคืนและส่วนลดรับ             (512 = ส่วนลดรับ)
	CY_End117  float64 // สินค้าสำเร็จรูปปลายงวดยกไป    (117 ปลายงวด)
	CY_Acc516  float64 // ต้นทุนขายอื่น                  (516)
	CY_Acc600  float64 // ปรับปรุงสินค้า                 (600)

	// Previous Year (ปีที่แล้ว) — ใช้ Bthisyear เป็นยอดปีก่อน
	PY_Prev119 float64
	PY_Acc500  float64
	PY_End119  float64
	PY_Acc506  float64
	PY_Acc508  float64
	PY_Prev118 float64
	PY_End118  float64
	PY_Prev117 float64
	PY_Acc510  float64
	PY_Acc512  float64
	PY_End117  float64
	PY_Acc516  float64
	PY_Acc600  float64
}

// ─────────────────────────────────────────────────────────────────
// Core Logic: คำนวณ CGS
//
// ตรรกะการคำนวณ (เทียบกับ FoxPro P_REPORT.PRG):
//
//	"ต้นงวด" (Beginning):
//	  = Bthisyear + SUM(tx.dr - tx.cr) WHERE date < dateFrom
//	  = ยอดสะสมตั้งแต่ต้นปีจนก่อนงวดที่เลือก
//	  เทียบกับ &RP_SATP ใน FoxPro
//
//	"ปลายงวด" (Ending Balance):
//	  = Bthisyear + SUM(tx.dr - tx.cr) WHERE date <= dateTo
//	  = ยอดคงเหลือสุทธิ ณ วันสิ้นงวด
//	  เทียบกับ &RP_ENDP ใน FoxPro (CLOS field)
//
//	"ยอดในงวด" (Period Amount):
//	  = SUM(tx.dr - tx.cr) WHERE dateFrom <= date <= dateTo
//	  เทียบกับ &RP_ENDPS - &RP_SATP ใน FoxPro
//
// ─────────────────────────────────────────────────────────────────
func buildCGSData(xlOptions excelize.Options, prdFrom, prdTo int) (*CGSData, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, err
	}

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if prdFrom < 1 || prdFrom > len(periods) || prdTo < 1 || prdTo > len(periods) || prdFrom > prdTo {
		return nil, fmt.Errorf("ช่วง Period ไม่ถูกต้อง")
	}

	dateFrom := periods[prdFrom-1].PStart
	dateTo := periods[prdTo-1].PEnd

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")

	cgs := &CGSData{
		ComName:    comName,
		PeriodFrom: prdFrom,
		PeriodTo:   prdTo,
		DateFrom:   dateFrom,
		DateTo:     dateTo,
		PrintTime:  time.Now(),
	}

	// โหลด Book_items ทั้งหมดก่อน แล้ว filter ใน memory
	type bItem struct {
		acCode string
		date   time.Time
		dr     float64
		cr     float64
	}
	var allTx []bItem
	bRows, _ := f.GetRows("Book_items")
	for i, r := range bRows {
		if i == 0 || len(r) < 11 || safeGet(r, 0) != comCode {
			continue
		}
		dtStr := safeGet(r, 1)
		dt, err := parseSubbookDate(dtStr)
		if err != nil {
			continue
		}
		allTx = append(allTx, bItem{
			acCode: safeGet(r, 5),
			date:   dt,
			dr:     parseFloat(safeGet(r, 9)),
			cr:     parseFloat(safeGet(r, 10)),
		})
	}

	// ── helpers ──────────────────────────────────────────────────

	// sumPeriod — SUM(dr-cr) ของ acCode ในช่วง dateFrom..dateTo
	sumPeriod := func(acCode string, from, to time.Time) float64 {
		var total float64
		for _, tx := range allTx {
			if tx.acCode == acCode && !tx.date.Before(from) && !tx.date.After(to) {
				total += tx.dr - tx.cr
			}
		}
		return total
	}

	// sumPeriodMulti — SUM(dr-cr) ของหลาย acCode ในช่วง from..to
	sumPeriodMulti := func(prefix3 int, from, to time.Time) float64 {
		var total float64
		for _, tx := range allTx {
			if len(tx.acCode) < 3 {
				continue
			}
			var p int
			fmt.Sscanf(tx.acCode[:3], "%d", &p)
			if p == prefix3 && !tx.date.Before(from) && !tx.date.After(to) {
				total += tx.dr - tx.cr
			}
		}
		return total
	}

	// sumPeriodRange — SUM(dr-cr) ของรหัสในช่วง prefix p1..p2 ในช่วง from..to
	sumPeriodRange := func(p1, p2 int, from, to time.Time) float64 {
		var total float64
		for _, tx := range allTx {
			if len(tx.acCode) < 3 {
				continue
			}
			var p int
			fmt.Sscanf(tx.acCode[:3], "%d", &p)
			if p >= p1 && p <= p2 && !tx.date.Before(from) && !tx.date.After(to) {
				total += tx.dr - tx.cr
			}
		}
		return total
	}

	// ──────────────────────────────────────────────────────────────
	// คำนวณ "ยอดต้นงวด" และ "ยอดปลายงวด" ของ inventory accounts
	// ตรรกะ:
	//   begBal (ต้นงวด) = Bthisyear + SUM(tx) WHERE date < dateFrom
	//   endBal (ปลายงวด) = Bthisyear + SUM(tx) WHERE date <= dateTo
	// ──────────────────────────────────────────────────────────────
	calcInventoryBegEnd := func(acCode3 int, bThisYear float64) (beg, end float64) {
		for _, tx := range allTx {
			if len(tx.acCode) < 3 {
				continue
			}
			var p int
			fmt.Sscanf(tx.acCode[:3], "%d", &p)
			if p != acCode3 {
				continue
			}
			net := tx.dr - tx.cr
			if tx.date.Before(dateFrom) {
				beg += net // สะสมก่อนงวด
			}
			if !tx.date.After(dateTo) {
				end += net // สะสมถึงสิ้นงวด
			}
		}
		beg += bThisYear
		end += bThisYear
		return
	}

	// ── อ่าน Ledger_Master เพื่อดึง Bthisyear ──────────────────
	lRows, _ := f.GetRows("Ledger_Master")
	for i, r := range lRows {
		if i == 0 || len(r) < 4 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := strings.TrimSpace(safeGet(r, 1))
		if len(acCode) < 3 {
			continue
		}

		var pNum int
		fmt.Sscanf(acCode[:3], "%d", &pNum)

		bThisYear := parseFloat(safeGet(r, 9)) // col J = Bthisyear

		switch {
		// ── Inventory accounts: คำนวณ ต้นงวด และ ปลายงวด ──
		case pNum == 117:
			beg, end := calcInventoryBegEnd(117, bThisYear)
			cgs.CY_Prev117 += beg
			cgs.CY_End117 += end
			// Previous Year: ใช้ Bthisyear เป็นยอดต้นปี (ยอดยกมาก่อนงวด 1)
			// Lastper12 (col AI = index 34) คือยอดสิ้นปีที่แล้ว
			lastper12 := parseFloat(safeGet(r, 34))
			cgs.PY_Prev117 += bThisYear // ต้นปีนี้ = ปลายปีที่แล้ว
			cgs.PY_End117 += lastper12  // ปลายปีที่แล้ว (ยอดสิ้นปีก่อน)

		case pNum == 118:
			beg, end := calcInventoryBegEnd(118, bThisYear)
			cgs.CY_Prev118 += beg
			cgs.CY_End118 += end
			lastper12 := parseFloat(safeGet(r, 34))
			cgs.PY_Prev118 += bThisYear
			cgs.PY_End118 += lastper12

		case pNum == 119:
			beg, end := calcInventoryBegEnd(119, bThisYear)
			cgs.CY_Prev119 += beg
			cgs.CY_End119 += end
			lastper12 := parseFloat(safeGet(r, 34))
			cgs.PY_Prev119 += bThisYear
			cgs.PY_End119 += lastper12
		}
	}

	// ── ยอดในงวด (Period Amounts) สำหรับ P&L accounts ──────────

	// 500-504: ซื้อวัตถุดิบ + ค่าใช้จ่ายซื้อ - ตัวหัก
	//   500 = ซื้อวัตถุดิบ (+)
	//   502 = ค่าใช้จ่ายในการซื้อวัตถุดิบ (+)
	//   504 = ตัวหักในการซื้อวัตถุดิบ (dr-cr จะเป็นลบ ถ้าเป็น credit)
	cgs.CY_Acc500 = sumPeriodRange(500, 504, dateFrom, dateTo)
	cgs.PY_Acc500 = sumPeriodRange(500, 504,
		periods[0].PStart, periods[len(periods)-1].PEnd) // ทั้งปีก่อน

	// 506: ค่าแรงงานทางตรง
	cgs.CY_Acc506 = sumPeriodMulti(506, dateFrom, dateTo)
	_ = sumPeriod // suppress unused warning

	// 508: โสหุ้ยอุปกรณ์โรงงาน = ค่าโสหุ้ยในการผลิต
	cgs.CY_Acc508 = sumPeriodMulti(508, dateFrom, dateTo)

	// 510: ซื้อสินค้าสำเร็จรูป ("ซื้อสินค้า")
	// 514: ค่าใช้จ่ายในการซื้อสินค้า (รวมเข้า ซื้อสินค้า)
	cgs.CY_Acc510 = sumPeriodRange(510, 510, dateFrom, dateTo) +
		sumPeriodRange(514, 514, dateFrom, dateTo)

	// 512: ส่วนลดรับในการซื้อสินค้า = "ส่งคืนและส่วนลดรับ"
	// ใน FoxPro ถูก ABS เพราะบันทึกเป็น Credit (ยอดลดต้นทุน)
	cgs.CY_Acc512 = math.Abs(sumPeriodMulti(512, dateFrom, dateTo))

	// 516: ต้นทุนขายอื่น
	cgs.CY_Acc516 = sumPeriodMulti(516, dateFrom, dateTo)

	// 600: ปรับปรุงสินค้า
	cgs.CY_Acc600 = sumPeriodMulti(600, dateFrom, dateTo)

	// ── Previous Year (ทั้งปี) ──────────────────────────────────
	// ใช้ SUM ทั้งปีจาก Ledger_Master Thisper01-12 หรืออ่านจาก Lastper
	// วิธีง่ายที่สุด: อ่าน Lastper01-12 จาก Ledger_Master
	// (ถ้าระบบ post ครบ) — แต่เนื่องจากอาจไม่ครบ ให้ใช้ Bthisyear ของ P&L
	// สำหรับปีก่อน: P&L accounts มียอดสุทธิทั้งปีอยู่ใน Thisper01-12 ปีก่อน
	// ซึ่งเก็บใน Lastper01-12 (col X-AI)
	lRows2, _ := f.GetRows("Ledger_Master")
	for i, r := range lRows2 {
		if i == 0 || len(r) < 4 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := strings.TrimSpace(safeGet(r, 1))
		if len(acCode) < 3 {
			continue
		}
		var pNum int
		fmt.Sscanf(acCode[:3], "%d", &pNum)

		// SUM Lastper01-12 (col 23-34, index 23..34) = ยอดเคลื่อนไหวปีก่อนทั้งปี
		var lastperSum float64
		for j := 23; j <= 34; j++ {
			lastperSum += parseFloat(safeGet(r, j))
		}

		switch {
		case pNum >= 500 && pNum <= 504:
			cgs.PY_Acc500 += lastperSum
		case pNum == 506:
			cgs.PY_Acc506 += lastperSum
		case pNum == 508:
			cgs.PY_Acc508 += lastperSum
		case pNum == 510 || pNum == 514:
			cgs.PY_Acc510 += lastperSum
		case pNum == 512:
			cgs.PY_Acc512 += math.Abs(lastperSum)
		case pNum == 516:
			cgs.PY_Acc516 += lastperSum
		case pNum == 600:
			cgs.PY_Acc600 += lastperSum
		}
	}

	return cgs, nil
}

// ─────────────────────────────────────────────────────────────────
// Export to PDF (gopdf) - Layout งบต้นทุนสินค้าที่ขาย
// ─────────────────────────────────────────────────────────────────
func exportCGSPDF(cgs *CGSData, savePath string) (string, error) {
	userFontsDir := filepath.Join(os.Getenv("LOCALAPPDATA"), "Microsoft", "Windows", "Fonts")
	sysFontsDir := filepath.Join(os.Getenv("WINDIR"), "Fonts")
	if sysFontsDir == "Fonts" || sysFontsDir == "" {
		sysFontsDir = filepath.Join("C:", "Windows", "Fonts")
	}
	fontsDir := userFontsDir
	if _, err := os.Stat(filepath.Join(userFontsDir, "Sarabun-Regular.ttf")); os.IsNotExist(err) {
		fontsDir = sysFontsDir
	}
	fontPath := ""
	for _, name := range []string{"Sarabun-Regular.ttf", "Sarabun-Medium.ttf", "THSarabunNew.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	if fontPath == "" {
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s", fontsDir)
	}

	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})
	pdf.AddPage()
	if err := pdf.AddTTFFont("thai", fontPath); err != nil {
		return "", err
	}
	pdf.AddTTFFont("thai-bold", boldPath)

	const (
		lm    = 30.0  // left margin
		col1  = 420.0 // ปีปัจจุบัน
		col2  = 540.0 // ปีที่แล้ว
		fs    = 9
		pageH = 800.0
	)
	y := 30.0

	nl := func(h float64) {
		y += h
		if y > pageH {
			pdf.AddPage()
			y = 30
		}
	}
	sf := func(size float64) { pdf.SetFont("thai", "", size) }

	drawSolidLine := func(x1, x2, yy float64) {
		pdf.SetLineWidth(0.5)
		pdf.Line(x1, yy, x2, yy)
	}

	drawDoubleSolidLine := func(x1, x2, yy float64) {
		drawSolidLine(x1, x2, yy)
		drawSolidLine(x1, x2, yy+2)
	}

	fmtNum := func(v float64) string {
		if v == 0 {
			return "0.00"
		}
		neg := v < 0
		if neg {
			v = -v
		}
		intPart := int64(v)
		dec := int64(math.Round((v - float64(intPart)) * 100))
		if dec >= 100 {
			intPart++
			dec -= 100
		}
		s := fmt.Sprintf("%d", intPart)
		res := ""
		for i, ch := range s {
			if i > 0 && (len(s)-i)%3 == 0 {
				res += ","
			}
			res += string(ch)
		}
		out := fmt.Sprintf("%s.%02d", res, dec)
		if neg {
			return "(" + out + ")"
		}
		return out
	}

	printRight := func(x float64, text string) {
		w, _ := pdf.MeasureTextWidth(text)
		pdf.SetXY(x-w, y)
		pdf.Cell(nil, text)
	}

	// ── Header ──
	pdf.SetFont("thai-bold", "", 13)
	w, _ := pdf.MeasureTextWidth(cgs.ComName)
	pdf.SetXY((595-w)/2, y)
	pdf.Cell(nil, cgs.ComName)
	nl(13 + 8)

	pdf.SetFont("thai", "", 10)
	title := "งบต้นทุนสินค้าที่ขาย"
	w, _ = pdf.MeasureTextWidth(title)
	pdf.SetXY((595-w)/2, y)
	pdf.Cell(nil, title)
	nl(10 + 6)

	pdf.SetFont("thai", "", 9)
	periodStr := fmt.Sprintf("ตั้งแต่ %s ถึง %s", cgs.DateFrom.Format("02/01/06"), cgs.DateTo.Format("02/01/06"))
	w, _ = pdf.MeasureTextWidth(periodStr)
	pdf.SetXY((595-w)/2, y)
	pdf.Cell(nil, periodStr)
	nl(9 + 20)

	sf(fs)
	printRight(col1, "ปีปัจจุบัน")
	printRight(col2, "ปีที่แล้ว")
	nl(fs + 8)

	// ═══════════════════════════════════════════════════════
	// ส่วนที่ 1: ต้นทุนสินค้าที่ผลิตระหว่างงวด
	// ═══════════════════════════════════════════════════════
	pdf.SetXY(lm, y)
	pdf.Cell(nil, "ต้นทุนสินค้าที่ผลิตระหว่างงวด")
	nl(fs + 6)

	// ── วัตถุดิบ (รหัส 119) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "วัตถุดิบ")
	nl(fs + 4)

	pdf.SetXY(lm+30, y)
	pdf.Cell(nil, "วัตถุดิบต้นงวดยกมา")
	printRight(col1, fmtNum(cgs.CY_Prev119))
	printRight(col2, fmtNum(cgs.PY_Prev119))
	nl(fs + 4)

	//  (500-504)
	pdf.SetXY(lm+30, y)
	pdf.Cell(nil, "บวก ซื้อวัตถุดิบระหว่างงวด")
	printRight(col1, fmtNum(cgs.CY_Acc500))
	printRight(col2, fmtNum(cgs.PY_Acc500))
	nl(fs + 4)

	cyTotalRawMatAvail := cgs.CY_Prev119 + cgs.CY_Acc500
	pyTotalRawMatAvail := cgs.PY_Prev119 + cgs.PY_Acc500
	pdf.SetXY(lm+30, y)
	pdf.Cell(nil, "รวมวัตถุดิบที่มีไว้เพื่อผลิต")
	printRight(col1, fmtNum(cyTotalRawMatAvail))
	printRight(col2, fmtNum(pyTotalRawMatAvail))
	nl(fs + 4)

	//  (119 ปลายงวด)
	pdf.SetXY(lm+30, y)
	pdf.Cell(nil, "หัก วัตถุดิบปลายงวดยกไป")
	printRight(col1, fmtNum(cgs.CY_End119))
	printRight(col2, fmtNum(cgs.PY_End119))
	drawSolidLine(col1-100, col1, y+fs+2)
	drawSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 4)

	cyRawMatUsed := cyTotalRawMatAvail - cgs.CY_End119
	pyRawMatUsed := pyTotalRawMatAvail - cgs.PY_End119
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "รวมต้นทุนวัตถุดิบ")
	printRight(col1, fmtNum(cyRawMatUsed))
	printRight(col2, fmtNum(pyRawMatUsed))
	nl(fs + 4)

	// ── ค่าแรงทางตรง (506) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "ค่าแรงทางตรง")
	printRight(col1, fmtNum(cgs.CY_Acc506))
	printRight(col2, fmtNum(cgs.PY_Acc506))
	nl(fs + 4)

	// ── ค่าโสหุ้ยในการผลิต (508) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "ค่าโสหุ้ยในการผลิต")
	printRight(col1, fmtNum(cgs.CY_Acc508))
	printRight(col2, fmtNum(cgs.PY_Acc508))
	drawSolidLine(col1-100, col1, y+fs+2)
	drawSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 4)

	cyTotalMfgCost := cyRawMatUsed + cgs.CY_Acc506 + cgs.CY_Acc508
	pyTotalMfgCost := pyRawMatUsed + cgs.PY_Acc506 + cgs.PY_Acc508
	pdf.SetXY(lm, y)
	pdf.Cell(nil, "รวมต้นทุนสินค้าที่ผลิต")
	printRight(col1, fmtNum(cyTotalMfgCost))
	printRight(col2, fmtNum(pyTotalMfgCost))
	nl(fs + 4)

	// ── งานระหว่างทำ   (118 ต้นงวด) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "บวก งานระหว่างทำต้นงวดยกมา")
	printRight(col1, fmtNum(cgs.CY_Prev118))
	printRight(col2, fmtNum(cgs.PY_Prev118))
	nl(fs + 4)

	//  (118 ปลายงวด)
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "หัก งานระหว่างทำปลายงวดยกไป")
	printRight(col1, fmtNum(cgs.CY_End118))
	printRight(col2, fmtNum(cgs.PY_End118))
	drawSolidLine(col1-100, col1, y+fs+2)
	drawSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 4)

	cyCostOfGoodsMfg := cyTotalMfgCost + cgs.CY_Prev118 - cgs.CY_End118
	pyCostOfGoodsMfg := pyTotalMfgCost + cgs.PY_Prev118 - cgs.PY_End118
	pdf.SetXY(lm, y)
	pdf.Cell(nil, "รวมต้นทุนสินค้าที่ผลิตระหว่างงวด")
	printRight(col1, fmtNum(cyCostOfGoodsMfg))
	printRight(col2, fmtNum(pyCostOfGoodsMfg))
	drawDoubleSolidLine(col1-100, col1, y+fs+2)
	drawDoubleSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 12)

	// ═══════════════════════════════════════════════════════
	// ส่วนที่ 2: ต้นทุนสินค้าที่ขาย
	// ═══════════════════════════════════════════════════════
	pdf.SetXY(lm, y)
	pdf.Cell(nil, "ต้นทุนสินค้าที่ขาย")
	nl(fs + 6)

	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "สินค้าที่ผลิตระหว่างงวด")
	printRight(col1, fmtNum(cyCostOfGoodsMfg))
	printRight(col2, fmtNum(pyCostOfGoodsMfg))
	nl(fs + 4)

	// ── สินค้าสำเร็จรูป (117) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "บวก สินค้าสำเร็จรูปต้นงวดยกมา")
	printRight(col1, fmtNum(cgs.CY_Prev117))
	printRight(col2, fmtNum(cgs.PY_Prev117))
	nl(fs + 4)

	// ── ซื้อสินค้า (510+514) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "บวก ซื้อสินค้า")
	printRight(col1, fmtNum(cgs.CY_Acc510))
	printRight(col2, fmtNum(cgs.PY_Acc510))
	nl(fs + 4)

	// ── ส่งคืนและส่วนลดรับ (512) ──
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "หัก ส่งคืนและส่วนลดรับ")
	printRight(col1, fmtNum(cgs.CY_Acc512))
	printRight(col2, fmtNum(cgs.PY_Acc512))
	drawSolidLine(col1-100, col1, y+fs+2)
	drawSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 4)

	cyGoodsAvail := cyCostOfGoodsMfg + cgs.CY_Prev117 + cgs.CY_Acc510 - cgs.CY_Acc512
	pyGoodsAvail := pyCostOfGoodsMfg + cgs.PY_Prev117 + cgs.PY_Acc510 - cgs.PY_Acc512
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "คงเหลือสินค้าสำเร็จรูปที่มีไว้ขาย")
	printRight(col1, fmtNum(cyGoodsAvail))
	printRight(col2, fmtNum(pyGoodsAvail))
	nl(fs + 4)

	//  (117 ปลายงวด)
	pdf.SetXY(lm+15, y)
	pdf.Cell(nil, "หัก สินค้าสำเร็จรูปปลายงวดยกไป")
	printRight(col1, fmtNum(cgs.CY_End117))
	printRight(col2, fmtNum(cgs.PY_End117))
	drawSolidLine(col1-100, col1, y+fs+2)
	drawSolidLine(col2-100, col2, y+fs+2)
	nl(fs + 4)

	cyCOGS := cyGoodsAvail - cgs.CY_End117
	pyCOGS := pyGoodsAvail - cgs.PY_End117

	// ── ต้นทุนขายอื่น (516) ──
	if cgs.CY_Acc516 != 0 || cgs.PY_Acc516 != 0 {
		pdf.SetXY(lm+15, y)
		pdf.Cell(nil, "บวก ต้นทุนขายอื่น")
		printRight(col1, fmtNum(cgs.CY_Acc516))
		printRight(col2, fmtNum(cgs.PY_Acc516))
		nl(fs + 4)
		cyCOGS += cgs.CY_Acc516
		pyCOGS += cgs.PY_Acc516
	}

	// ── ปรับปรุงสินค้า (600) ──
	if cgs.CY_Acc600 != 0 || cgs.PY_Acc600 != 0 {
		pdf.SetXY(lm+15, y)
		pdf.Cell(nil, "บวก(หัก) ปรับปรุงสินค้า")
		printRight(col1, fmtNum(cgs.CY_Acc600))
		printRight(col2, fmtNum(cgs.PY_Acc600))
		nl(fs + 4)
		cyCOGS += cgs.CY_Acc600
		pyCOGS += cgs.PY_Acc600
	}

	pdf.SetXY(lm, y)
	pdf.Cell(nil, "รวมต้นทุนสินค้าที่ขาย")
	printRight(col1, fmtNum(cyCOGS))
	printRight(col2, fmtNum(pyCOGS))
	drawDoubleSolidLine(col1-100, col1, y+fs+2)
	drawDoubleSolidLine(col2-100, col2, y+fs+2)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// Export to Excel (Excelize) - Layout งบต้นทุนสินค้าที่ขาย
// ─────────────────────────────────────────────────────────────────
func exportCGSExcel(cgs *CGSData, savePath string) (string, error) {
	fx := excelize.NewFile()
	sn := "Cost of Goods Sold"
	fx.SetSheetName("Sheet1", sn)

	fx.SetSheetView(sn, 0, &excelize.ViewOptions{ShowGridLines: func() *bool { b := false; return &b }()})

	fx.SetColWidth(sn, "A", "A", 3)
	fx.SetColWidth(sn, "B", "B", 45)
	fx.SetColWidth(sn, "C", "D", 20)

	stCtrBold, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 11, Bold: true}, Alignment: &excelize.Alignment{Horizontal: "center"}})
	stCtr, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10}, Alignment: &excelize.Alignment{Horizontal: "center"}})
	stLeft, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Bold: true}, Alignment: &excelize.Alignment{Horizontal: "left"}})
	stLeftI1, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10}, Alignment: &excelize.Alignment{Horizontal: "left", Indent: 2}})
	stLeftI2, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10}, Alignment: &excelize.Alignment{Horizontal: "left", Indent: 4}})

	stNum, _ := fx.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10}, Alignment: &excelize.Alignment{Horizontal: "right"}, NumFmt: 4})
	stNumDash, _ := fx.NewStyle(&excelize.Style{
		Font: &excelize.Font{Size: 10}, Alignment: &excelize.Alignment{Horizontal: "right"}, NumFmt: 4,
		Border: []excelize.Border{{Type: "bottom", Color: "000000", Style: 1}},
	})
	stNumDoubleDash, _ := fx.NewStyle(&excelize.Style{
		Font: &excelize.Font{Size: 10, Bold: true}, Alignment: &excelize.Alignment{Horizontal: "right"}, NumFmt: 4,
		Border: []excelize.Border{{Type: "bottom", Color: "000000", Style: 6}},
	})

	row := 1
	set := func(col string, val interface{}, st int) {
		cell := fmt.Sprintf("%s%d", col, row)
		fx.SetCellValue(sn, cell, val)
		fx.SetCellStyle(sn, cell, cell, st)
	}
	setF := func(col string, val float64, st int) {
		cell := fmt.Sprintf("%s%d", col, row)
		fx.SetCellFloat(sn, cell, val, 2, 64)
		fx.SetCellStyle(sn, cell, cell, st)
	}
	mc := func(c1, c2 string) { fx.MergeCell(sn, fmt.Sprintf("%s%d", c1, row), fmt.Sprintf("%s%d", c2, row)) }
	br := func() { row++ }

	// ── Header ──
	mc("B", "D")
	set("B", cgs.ComName, stCtrBold)
	br()
	mc("B", "D")
	set("B", "งบต้นทุนสินค้าที่ขาย", stCtr)
	br()
	mc("B", "D")
	set("B", fmt.Sprintf("ตั้งแต่ %s ถึง %s", cgs.DateFrom.Format("02/01/06"), cgs.DateTo.Format("02/01/06")), stCtr)
	br()
	br()
	set("C", "ปีปัจจุบัน", stCtr)
	set("D", "ปีที่แล้ว", stCtr)
	br()

	// ═══════════════════════════════════════════════════════
	// ส่วนที่ 1: ต้นทุนสินค้าที่ผลิตระหว่างงวด
	// ═══════════════════════════════════════════════════════
	set("B", "ต้นทุนสินค้าที่ผลิตระหว่างงวด", stLeft)
	br()

	// ── วัตถุดิบ (119) ──
	set("B", "วัตถุดิบ", stLeftI1)
	br()

	//  (119 ต้นงวด)
	set("B", "วัตถุดิบต้นงวดยกมา", stLeftI2)
	setF("C", cgs.CY_Prev119, stNum)
	setF("D", cgs.PY_Prev119, stNum)
	br()

	//  (500-504)
	set("B", "บวก ซื้อวัตถุดิบระหว่างงวด", stLeftI2)
	setF("C", cgs.CY_Acc500, stNum)
	setF("D", cgs.PY_Acc500, stNum)
	br()

	cyTotalRawMatAvail := cgs.CY_Prev119 + cgs.CY_Acc500
	pyTotalRawMatAvail := cgs.PY_Prev119 + cgs.PY_Acc500
	set("B", "รวมวัตถุดิบที่มีไว้เพื่อผลิต", stLeftI2)
	setF("C", cyTotalRawMatAvail, stNum)
	setF("D", pyTotalRawMatAvail, stNum)
	br()

	//  (119 ปลายงวด)
	set("B", "หัก วัตถุดิบปลายงวดยกไป", stLeftI2)
	setF("C", cgs.CY_End119, stNumDash)
	setF("D", cgs.PY_End119, stNumDash)
	br()

	cyRawMatUsed := cyTotalRawMatAvail - cgs.CY_End119
	pyRawMatUsed := pyTotalRawMatAvail - cgs.PY_End119
	set("B", "รวมต้นทุนวัตถุดิบ", stLeftI1)
	setF("C", cyRawMatUsed, stNum)
	setF("D", pyRawMatUsed, stNum)
	br()

	// ── ค่าแรงทางตรง (506) ──
	set("B", "ค่าแรงทางตรง", stLeftI1)
	setF("C", cgs.CY_Acc506, stNum)
	setF("D", cgs.PY_Acc506, stNum)
	br()

	// ── ค่าโสหุ้ยในการผลิต (508) ──
	set("B", "ค่าโสหุ้ยในการผลิต", stLeftI1)
	setF("C", cgs.CY_Acc508, stNumDash)
	setF("D", cgs.PY_Acc508, stNumDash)
	br()

	cyTotalMfgCost := cyRawMatUsed + cgs.CY_Acc506 + cgs.CY_Acc508
	pyTotalMfgCost := pyRawMatUsed + cgs.PY_Acc506 + cgs.PY_Acc508
	set("B", "รวมต้นทุนสินค้าที่ผลิต", stLeft)
	setF("C", cyTotalMfgCost, stNum)
	setF("D", pyTotalMfgCost, stNum)
	br()

	// ── งานระหว่างทำ (118) ──
	set("B", "บวก งานระหว่างทำต้นงวดยกมา", stLeftI1)
	setF("C", cgs.CY_Prev118, stNum)
	setF("D", cgs.PY_Prev118, stNum)
	br()
	//  (118 ปลายงวด)
	set("B", "หัก งานระหว่างทำปลายงวดยกไป", stLeftI1)
	setF("C", cgs.CY_End118, stNumDash)
	setF("D", cgs.PY_End118, stNumDash)
	br()

	cyCostOfGoodsMfg := cyTotalMfgCost + cgs.CY_Prev118 - cgs.CY_End118
	pyCostOfGoodsMfg := pyTotalMfgCost + cgs.PY_Prev118 - cgs.PY_End118
	set("B", "รวมต้นทุนสินค้าที่ผลิตระหว่างงวด", stLeft)
	setF("C", cyCostOfGoodsMfg, stNumDoubleDash)
	setF("D", pyCostOfGoodsMfg, stNumDoubleDash)
	br()
	br()

	// ═══════════════════════════════════════════════════════
	// ส่วนที่ 2: ต้นทุนสินค้าที่ขาย
	// ═══════════════════════════════════════════════════════
	set("B", "ต้นทุนสินค้าที่ขาย", stLeft)
	br()

	set("B", "สินค้าที่ผลิตระหว่างงวด", stLeftI1)
	setF("C", cyCostOfGoodsMfg, stNum)
	setF("D", pyCostOfGoodsMfg, stNum)
	br()

	// ── สินค้าสำเร็จรูป (117) ──
	set("B", "บวก สินค้าสำเร็จรูปต้นงวดยกมา", stLeftI1)
	setF("C", cgs.CY_Prev117, stNum)
	setF("D", cgs.PY_Prev117, stNum)
	br()

	// ── ซื้อสินค้า (510+514) ──
	set("B", "บวก ซื้อสินค้า", stLeftI1)
	setF("C", cgs.CY_Acc510, stNum)
	setF("D", cgs.PY_Acc510, stNum)
	br()

	// ── ส่งคืนและส่วนลดรับ (512) ──
	set("B", "หัก ส่งคืนและส่วนลดรับ", stLeftI1)
	setF("C", cgs.CY_Acc512, stNumDash)
	setF("D", cgs.PY_Acc512, stNumDash)
	br()

	cyGoodsAvail := cyCostOfGoodsMfg + cgs.CY_Prev117 + cgs.CY_Acc510 - cgs.CY_Acc512
	pyGoodsAvail := pyCostOfGoodsMfg + cgs.PY_Prev117 + cgs.PY_Acc510 - cgs.PY_Acc512
	set("B", "คงเหลือสินค้าสำเร็จรูปที่มีไว้ขาย", stLeftI1)
	setF("C", cyGoodsAvail, stNum)
	setF("D", pyGoodsAvail, stNum)
	br()

	// ── สินค้าสำเร็จรูปปลายงวด (117) ──
	set("B", "หัก สินค้าสำเร็จรูปปลายงวดยกไป", stLeftI1)
	setF("C", cgs.CY_End117, stNumDash)
	setF("D", cgs.PY_End117, stNumDash)
	br()

	cyCOGS := cyGoodsAvail - cgs.CY_End117
	pyCOGS := pyGoodsAvail - cgs.PY_End117

	// ── ต้นทุนขายอื่น (516) ──
	if cgs.CY_Acc516 != 0 || cgs.PY_Acc516 != 0 {
		set("B", "บวก ต้นทุนขายอื่น", stLeftI1)
		setF("C", cgs.CY_Acc516, stNumDash)
		setF("D", cgs.PY_Acc516, stNumDash)
		br()
		cyCOGS += cgs.CY_Acc516
		pyCOGS += cgs.PY_Acc516
	}

	// ── ปรับปรุงสินค้า (600) ──
	if cgs.CY_Acc600 != 0 || cgs.PY_Acc600 != 0 {
		set("B", "บวก(หัก) ปรับปรุงสินค้า", stLeftI1)
		setF("C", cgs.CY_Acc600, stNumDash)
		setF("D", cgs.PY_Acc600, stNumDash)
		br()
		cyCOGS += cgs.CY_Acc600
		pyCOGS += cgs.PY_Acc600
	}

	set("B", "รวมต้นทุนสินค้าที่ขาย", stLeft)
	setF("C", cyCOGS, stNumDoubleDash)
	setF("D", pyCOGS, stNumDoubleDash)
	br()

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return fx.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// UI Dialog สำหรับ CGS
// ─────────────────────────────────────────────────────────────────
func showCGSDialog(w fyne.Window, onGoSetup func()) {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOpts)
	if err != nil {
		dialog.ShowError(err, w)
		return
	}

	reportDir := getReportDir(xlOpts)
	if strings.HasSuffix(filepath.ToSlash(reportDir), "/Desktop") ||
		reportDir == filepath.ToSlash(filepath.Dir(currentDBPath)) {
		var warn dialog.Dialog
		btnGo := newEnterButton("ไปตั้งค่า (Enter)", func() {
			warn.Hide()
			if onGoSetup != nil {
				onGoSetup()
			}
		})
		btnGo.Importance = widget.HighImportance
		btnCancel2 := newEscButton("ยกเลิก (Esc)", func() { warn.Hide() })
		warn = dialog.NewCustomWithoutButtons(
			"⚠️  ยังไม่ได้ตั้งค่า Report Path",
			container.NewVBox(
				widget.NewLabel("กรุณาตั้งค่า Report Path ที่ Setup > Company Profile"),
				widget.NewLabel("เพื่อให้ไฟล์รายงานเก็บในที่เดียวกันทุกครั้ง"),
				widget.NewSeparator(),
				container.NewCenter(container.NewHBox(btnGo, btnCancel2)),
			), w)
		warn.Show()
		w.Canvas().Focus(btnGo)
		return
	}

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	options := make([]string, showUpTo)
	for i, p := range periods[:showUpTo] {
		options[i] = fmt.Sprintf("Period %d (%s)", i+1, p.PEnd.Format("02/01/06"))
	}

	selPeriodFrom := widget.NewSelect(options, nil)
	selPeriodFrom.SetSelectedIndex(0)

	selPeriodTo := widget.NewSelect(options, nil)
	selPeriodTo.SetSelectedIndex(showUpTo - 1)

	btnExcelTH := widget.NewButton("📊 Excel", nil)
	btnPDFTH := widget.NewButton("📄 PDF", nil)
	btnCancel := widget.NewButton("❌ ปิด", nil)

	var pop *widget.PopUp
	prevKey := w.Canvas().OnTypedKey()
	closePopup := func() {
		if pop != nil {
			pop.Hide()
		}
		w.Canvas().SetOnTypedKey(prevKey)
	}

	pop = widget.NewModalPopUp(
		container.NewVBox(
			widget.NewLabelWithStyle("งบต้นทุนสินค้าที่ขาย / Cost of Goods Sold",
				fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			widget.NewSeparator(),
			container.NewHBox(
				widget.NewLabel("ตั้งแต่งวด:"),
				selPeriodFrom,
				widget.NewLabel("  ถึงงวด:"),
				selPeriodTo,
			),
			widget.NewSeparator(),
			container.NewCenter(container.NewHBox(btnExcelTH, btnPDFTH, btnCancel)),
		),
		w.Canvas(),
	)
	pop.Resize(fyne.NewSize(560, 180))

	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		if key.Name == fyne.KeyEscape {
			closePopup()
		}
	})
	btnCancel.OnTapped = closePopup

	showDone := func(pathToOpen string) {
		ext2 := filepath.Ext(pathToOpen)
		isTmp := strings.HasSuffix(strings.TrimSuffix(pathToOpen, ext2), "_tmp")
		title := "✅ บันทึกรายงานแล้ว"
		note := ""
		if isTmp {
			title = "⚠️ เปิดรายงานชั่วคราว"
			note = "ปิดไฟล์เดิมก่อน แล้วกด Export ใหม่เพื่อบันทึกถาวร"
		}
		var done dialog.Dialog
		ok2 := newEnterButton("OK — เปิดไฟล์", func() {
			done.Hide()
			openFile(pathToOpen)
		})
		btnClose := newEscButton("ปิด", func() { done.Hide() })
		body := container.NewVBox(
			widget.NewLabel(title),
			widget.NewLabel(filepath.Base(pathToOpen)),
		)
		if note != "" {
			body.Add(widget.NewLabel(note))
		}
		body.Add(widget.NewSeparator())
		body.Add(container.NewCenter(container.NewHBox(ok2, btnClose)))
		done = dialog.NewCustomWithoutButtons("รายงาน", body, w)
		done.Show()
		w.Canvas().Focus(ok2)
	}

	run := func(isPDF bool) {
		prdFrom := selPeriodFrom.SelectedIndex() + 1
		prdTo := selPeriodTo.SelectedIndex() + 1

		if prdFrom < 1 || prdTo < 1 {
			dialog.ShowInformation("แจ้งเตือน", "กรุณาเลือก Period ให้ครบถ้วน", w)
			return
		}
		if prdFrom > prdTo {
			dialog.ShowInformation("แจ้งเตือน", "Period เริ่มต้นต้องไม่มากกว่า Period สิ้นสุด", w)
			return
		}
		closePopup()

		cgs, err := buildCGSData(xlOpts, prdFrom, prdTo)
		if err != nil {
			dialog.ShowError(err, w)
			return
		}

		var savePath, pathToOpen string
		var exportErr error
		timestamp := time.Now().Format("20060102_150405")
		if isPDF {
			savePath = filepath.Join(reportDir, fmt.Sprintf("CGS_P%02d-%02d_%s.pdf", prdFrom, prdTo, timestamp))
			pathToOpen, exportErr = exportCGSPDF(cgs, savePath)
		} else {
			savePath = filepath.Join(reportDir, fmt.Sprintf("CGS_P%02d-%02d_%s.xlsx", prdFrom, prdTo, timestamp))
			pathToOpen, exportErr = exportCGSExcel(cgs, savePath)
		}
		if exportErr != nil {
			dialog.ShowError(exportErr, w)
			return
		}
		showDone(pathToOpen)
	}

	btnExcelTH.OnTapped = func() { run(false) }
	btnPDFTH.OnTapped = func() { run(true) }
	pop.Show()
}
