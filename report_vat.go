package main

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
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
// VatRow — 1 รายการ VAT
// ─────────────────────────────────────────────────────────────────
type VatRow struct {
	Bdate     string
	Bitem     string
	Bref      string
	Boff      string
	Bcomtaxid string
	BaseAmt   float64
	VatAmt    float64
}

// ─────────────────────────────────────────────────────────────────
// buildVatReport
// vatType = "P" → Purchases VAT (235TVAT Bdebit > 0)
// vatType = "S" → Sales VAT     (235TVAT Bcredit > 0)
// ─────────────────────────────────────────────────────────────────
func buildVatReport(xlOpts excelize.Options, periodNo int, vatType string) ([]VatRow, CompanyPeriodConfig, error) {
	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err != nil {
		return nil, CompanyPeriodConfig{}, err
	}
	defer f.Close()

	cfg, err := loadCompanyPeriod(xlOpts)
	if err != nil {
		return nil, cfg, err
	}

	comCode := getComCodeFromExcel(xlOpts)
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if periodNo < 1 || periodNo > len(periods) {
		return nil, cfg, fmt.Errorf("period %d ไม่ถูกต้อง", periodNo)
	}
	p := periods[periodNo-1]

	rows, err := f.GetRows("Book_items")
	if err != nil {
		return nil, cfg, fmt.Errorf("อ่าน Book_items ไม่ได้: %v", err)
	}

	type rawLine struct {
		Bdate     string
		Bitem     string
		Bref      string
		Boff      string
		Bcomtaxid string
		AcCode    string
		Bdebit    float64
		Bcredit   float64
	}

	voucherMap := map[string][]rawLine{}
	for i, row := range rows {
		if i == 0 || len(row) < 11 || safeGet(row, 0) != comCode {
			continue
		}
		bdate := strings.TrimSpace(safeGet(row, 1))
		t, err := time.Parse("02/01/06", bdate)
		if err != nil {
			continue
		}
		if t.Before(p.PStart) || t.After(p.PEnd) {
			continue
		}
		rl := rawLine{
			Bdate:     bdate,
			Bitem:     safeGet(row, 2),
			Bref:      safeGet(row, 11),
			Boff:      safeGet(row, 12),
			Bcomtaxid: safeGet(row, 13),
			AcCode:    safeGet(row, 5),
			Bdebit:    parseFloat(safeGet(row, 9)),
			Bcredit:   parseFloat(safeGet(row, 10)),
		}
		voucherMap[rl.Bitem] = append(voucherMap[rl.Bitem], rl)
	}

	var result []VatRow
	for _, lines := range voucherMap {
		for _, vatLine := range lines {
			if vatLine.AcCode != "235TVAT" {
				continue
			}
			isPurchase := vatLine.Bdebit > 0
			isSale := vatLine.Bcredit > 0
			if vatType == "P" && !isPurchase {
				continue
			}
			if vatType == "S" && !isSale {
				continue
			}
			var vatAmt float64
			if isPurchase {
				vatAmt = vatLine.Bdebit
			} else {
				vatAmt = vatLine.Bcredit
			}
			var baseAmt float64
			for _, other := range lines {
				if other.AcCode == "235TVAT" || other.AcCode == "120VAT" {
					continue
				}
				if vatType == "P" {
					baseAmt += other.Bdebit
				} else {
					baseAmt += other.Bcredit
				}
			}
			if baseAmt == 0 && vatAmt == 0 {
				continue
			}
			bref, boff, bcomtaxid := vatLine.Bref, vatLine.Boff, vatLine.Bcomtaxid
			for _, other := range lines {
				if bref == "" && other.Bref != "" {
					bref = other.Bref
				}
				if boff == "" && other.Boff != "" {
					boff = other.Boff
				}
				if bcomtaxid == "" && other.Bcomtaxid != "" {
					bcomtaxid = other.Bcomtaxid
				}
			}
			result = append(result, VatRow{
				Bdate:     vatLine.Bdate,
				Bitem:     vatLine.Bitem,
				Bref:      bref,
				Boff:      boff,
				Bcomtaxid: bcomtaxid,
				BaseAmt:   baseAmt,
				VatAmt:    vatAmt,
			})
			break
		}
	}

	sort.Slice(result, func(i, j int) bool {
		if result[i].Bdate != result[j].Bdate {
			return result[i].Bdate < result[j].Bdate
		}
		return result[i].Bitem < result[j].Bitem
	})
	return result, cfg, nil
}

// ─────────────────────────────────────────────────────────────────
// exportVatExcel
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
// exportVatExcel
// ─────────────────────────────────────────────────────────────────
func exportVatExcel(xlOpts excelize.Options, periodNo int, vatType string, savePath string) (string, error) {
	vatRows, cfg, err := buildVatReport(xlOpts, periodNo, vatType)
	if err != nil {
		return "", err
	}

	// แก้ไข: เช็ค error จาก OpenFile เพื่อป้องกัน panic
	f, err := excelize.OpenFile(currentDBPath, xlOpts)
	if err != nil {
		return "", fmt.Errorf("ไม่สามารถเปิดไฟล์ฐานข้อมูลเพื่ออ่านข้อมูลบริษัทได้: %v", err)
	}
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comAddr, _ := f.GetCellValue("Company_Profile", "C2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	// ป้องกัน index out of range เผื่อไว้
	if periodNo < 1 || periodNo > len(periods) {
		return "", fmt.Errorf("period %d ไม่ถูกต้อง", periodNo)
	}
	pEnd := periods[periodNo-1].PEnd

	var titleTH, titleEN string
	if vatType == "P" {
		titleTH = "รายงานภาษีซื้อ"
		titleEN = "Purchases VAT Report"
	} else {
		titleTH = "รายงานภาษีขาย"
		titleEN = "Sales VAT Report"
	}

	wb := excelize.NewFile()
	sh := "Sheet1"
	wb.SetSheetName("Sheet1", sh)

	colBorder := []excelize.Border{
		{Type: "left", Color: "CCCCCC", Style: 1},
		{Type: "right", Color: "CCCCCC", Style: 1},
	}

	stBold := func(sz float64, halign string) int {
		s, _ := wb.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: sz},
			Alignment: &excelize.Alignment{Horizontal: halign, WrapText: true},
		})
		return s
	}
	stNormal := func(halign string) int {
		s, _ := wb.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{Horizontal: halign},
			Border:    colBorder,
		})
		return s
	}
	stNum := func() int {
		s, _ := wb.NewStyle(&excelize.Style{
			NumFmt:    4,
			Alignment: &excelize.Alignment{Horizontal: "right"},
			Border:    colBorder,
		})
		return s
	}
	stHdr := func() int {
		s, _ := wb.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: 10},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D9E1F2"}, Pattern: 1},
			Alignment: &excelize.Alignment{Horizontal: "center", WrapText: true, Vertical: "center"},
			Border: []excelize.Border{
				{Type: "top", Color: "000000", Style: 1},
				{Type: "bottom", Color: "000000", Style: 1},
				{Type: "left", Color: "000000", Style: 1},
				{Type: "right", Color: "000000", Style: 1},
			},
		})
		return s
	}
	stTotal := func() int {
		s, _ := wb.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true},
			NumFmt:    4,
			Border:    []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 6}},
			Alignment: &excelize.Alignment{Horizontal: "right"},
		})
		return s
	}
	stTotalLbl := func() int {
		s, _ := wb.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true},
			Border:    []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 6}},
			Alignment: &excelize.Alignment{Horizontal: "center"},
		})
		return s
	}

	set := func(cell, val string, style int) {
		wb.SetCellValue(sh, cell, val)
		wb.SetCellStyle(sh, cell, cell, style)
	}
	setF := func(cell string, val float64, style int) {
		wb.SetCellValue(sh, cell, val)
		wb.SetCellStyle(sh, cell, cell, style)
	}
	merge := func(c1, c2 string) { wb.MergeCell(sh, c1, c2) }
	rc := func(r int, col string) string { return fmt.Sprintf("%s%d", col, r) }

	r := 1
	merge(rc(r, "A"), rc(r, "G"))
	set(rc(r, "A"), comName, stBold(13, "center"))
	r++
	merge(rc(r, "A"), rc(r, "G"))
	set(rc(r, "A"), comAddr, stNormal("center"))
	r++
	merge(rc(r, "A"), rc(r, "G"))
	set(rc(r, "A"), fmt.Sprintf("%s / %s", titleTH, titleEN), stBold(12, "center"))
	r++
	merge(rc(r, "A"), rc(r, "D"))
	set(rc(r, "A"), fmt.Sprintf("ณ วันที่  %s", pEnd.Format("02/01/2006")), stNormal("center"))
	merge(rc(r, "E"), rc(r, "G"))
	set(rc(r, "E"), fmt.Sprintf("เลขประจำตัวผู้เสียภาษีอากร    %s", comTax), stNormal("center"))
	r++
	r++ // blank

	// header บรรทัดแรก
	hdrRow1 := r
	merge(rc(r, "A"), rc(r, "B"))
	set(rc(r, "A"), "ในกำกับภาษี", stHdr())
	var xlHdrParty, xlHdrTaxParty string
	if vatType == "P" {
		xlHdrParty = "ชื่อผู้ขาย\nหรือผู้ให้บริการ"
		xlHdrTaxParty = "เลขประจำตัวผู้เสียภาษี\nของผู้ขาย\nหรือผู้ให้บริการ"
	} else {
		xlHdrParty = "ชื่อผู้ซื้อ\nหรือผู้รับบริการ"
		xlHdrTaxParty = "เลขประจำตัวผู้เสียภาษี\nของผู้ซื้อ\nหรือผู้รับบริการ"
	}
	set(rc(r, "C"), xlHdrParty, stHdr())
	set(rc(r, "D"), xlHdrTaxParty, stHdr())
	set(rc(r, "E"), "สถานประกอบการ\nสนง./สาขาที่", stHdr())
	set(rc(r, "F"), "มูลค่าสินค้า\nหรือบริการ", stHdr())
	set(rc(r, "G"), "จำนวนเงิน\nภาษีมูลค่าเพิ่ม", stHdr())
	r++
	// header บรรทัดสอง
	set(rc(r, "A"), "วันเดือนปี", stHdr())
	set(rc(r, "B"), "เลขที่", stHdr())
	set(rc(r, "C"), "", stHdr())
	set(rc(r, "D"), "", stHdr())
	set(rc(r, "E"), "", stHdr())
	set(rc(r, "F"), "", stHdr())
	set(rc(r, "G"), "", stHdr())
	r++

	ns := stNum()
	nl := stNormal("left")
	nc := stNormal("center")
	var sumBase, sumVat float64

	for _, vr := range vatRows {
		set(rc(r, "A"), vr.Bdate, nc)
		set(rc(r, "B"), vr.Bref, nl)
		set(rc(r, "C"), vr.Boff, nl)
		set(rc(r, "D"), vr.Bcomtaxid, nc)
		set(rc(r, "E"), "00000", nc)
		setF(rc(r, "F"), vr.BaseAmt, ns)
		setF(rc(r, "G"), vr.VatAmt, ns)
		sumBase += vr.BaseAmt
		sumVat += vr.VatAmt
		r++
	}

	ts := stTotal()
	tl := stTotalLbl()
	merge(rc(r, "A"), rc(r, "E"))
	set(rc(r, "A"), fmt.Sprintf("รวมทั้งสิ้น   %d   รายการ", len(vatRows)), tl)
	setF(rc(r, "F"), sumBase, ts)
	setF(rc(r, "G"), sumVat, ts)

	wb.SetColWidth(sh, "A", "A", 10) // วันเดือนปี
	wb.SetColWidth(sh, "B", "B", 14) // เลขที่
	wb.SetColWidth(sh, "C", "C", 26) // ชื่อผู้ขาย/ผู้ซื้อ
	wb.SetColWidth(sh, "D", "D", 16) // เลขประจำตัวผู้เสียภาษี (13 หลัก)
	wb.SetColWidth(sh, "E", "E", 9)  // สาขา (00000)
	wb.SetColWidth(sh, "F", "F", 13) // มูลค่าสินค้า
	wb.SetColWidth(sh, "G", "G", 12) // ภาษีมูลค่าเพิ่ม
	wb.SetRowHeight(sh, hdrRow1, 36)

	orientation := "portrait"
	paperSize := 9   // 9 = A4
	fitToWidth := 1  // บีบให้พอดีความกว้าง 1 หน้า (ป้องกันคอลัมน์ล้น)
	fitToHeight := 0 // 0 = ไม่จำกัดความสูง (ปล่อยให้รันขึ้นหน้าใหม่ได้ถ้ารายการเยอะ)

	errLayout := wb.SetPageLayout(sh, &excelize.PageLayoutOptions{
		Orientation: &orientation,
		Size:        &paperSize,
		FitToWidth:  &fitToWidth,
		FitToHeight: &fitToHeight,
	})
	if errLayout != nil {
		fmt.Printf("Warning: ไม่สามารถตั้งค่าหน้ากระดาษได้: %v\n", errLayout)
	}
	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return wb.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// exportVatPDF — Portrait A4
// ─────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────
// exportVatPDF — Portrait A4
// ─────────────────────────────────────────────────────────────────
func exportVatPDF(xlOpts excelize.Options, periodNo int, vatType string, savePath string) (string, error) {
	vatRows, cfg, err := buildVatReport(xlOpts, periodNo, vatType)
	if err != nil {
		return "", err
	}
	f, _ := excelize.OpenFile(currentDBPath, xlOpts)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comAddr, _ := f.GetCellValue("Company_Profile", "C2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	pEnd := periods[periodNo-1].PEnd

	var titleTH string
	if vatType == "P" {
		titleTH = "รายงานภาษีซื้อ"
	} else {
		titleTH = "รายงานภาษีขาย"
	}

	// font
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
	for _, name := range []string{"Sarabun-Regular.ttf", "Sarabun-Medium.ttf", "THSarabunNew.ttf", "arial.ttf", "Arial.ttf"} {
		if pp := filepath.Join(fontsDir, name); fileExists(pp) {
			fontPath = pp
			break
		}
	}
	if fontPath == "" {
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s", fontsDir)
	}

	// A4 Portrait
	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})
	pdf.AddPage()
	if err := pdf.AddTTFFont("thai", fontPath); err != nil {
		return "", err
	}

	const (
		lm    = 25.0
		pageW = 595.28
		fs    = 7.5
		fsH   = 8.0
		// ปรับความกว้างคอลัมน์ให้สมดุลตามภาพตัวอย่าง (ใช้เป็นพิกัด X ของเส้นแนวตั้ง)
		cDate   = 80.0       // วันเดือนปี
		cRef    = 145.0      // เลขที่
		cName   = 305.0      // ชื่อผู้ขาย/ผู้ซื้อ
		cTaxID  = 405.0      // เลขประจำตัวผู้เสียภาษี
		cBranch = 450.0      // สนง./สาขาที่
		cBase   = 510.0      // มูลค่าสินค้า
		cVat    = pageW - lm // ภาษีมูลค่าเพิ่ม (ชิดขอบขวาพอดี)
	)

	y := 25.0
	pageNo := 1

	addNl := func(h float64) { y += h }
	sf := func(sz float64) { pdf.SetFont("thai", "", sz) }
	hln := func(x1, x2, yy, lw float64) {
		pdf.SetLineWidth(lw)
		pdf.Line(x1, yy, x2, yy)
	}
	putR := func(text string, rightX float64) {
		w, _ := pdf.MeasureTextWidth(text)
		pdf.SetXY(rightX-w, y)
		pdf.Cell(nil, text)
	}
	putL := func(text string, x, maxW float64) {
		for {
			w, _ := pdf.MeasureTextWidth(text)
			if w <= maxW || len(text) == 0 {
				break
			}
			// ตัดทีละ rune
			runes := []rune(text)
			text = string(runes[:len(runes)-1])
		}
		pdf.SetXY(x, y)
		pdf.Cell(nil, text)
	}

	// putC — วาง text กึ่งกลางระหว่าง x1 ถึง x2
	putC := func(text string, x1, x2 float64) {
		if text == "" {
			return
		}
		tw, _ := pdf.MeasureTextWidth(text)
		mid := x1 + (x2-x1)/2 - tw/2
		if mid < x1 {
			mid = x1
		}
		pdf.SetXY(mid, y)
		pdf.Cell(nil, text)
	}

	var yHdrTop, yHdrMid, yHdrBot float64

	printHeader := func() {
		// ── Page number (มุมขวาบน) ──
		sf(8)
		pgStr := fmt.Sprintf("Page  %d", pageNo)
		pw, _ := pdf.MeasureTextWidth(pgStr)
		pdf.SetXY(pageW-lm-pw, y)
		pdf.Cell(nil, pgStr)
		addNl(14)

		// ── ชื่อบริษัท (center, bold) ──
		sf(13)
		w, _ := pdf.MeasureTextWidth(comName)
		pdf.SetXY((pageW-w)/2, y)
		pdf.Cell(nil, comName)
		addNl(16)

		// ── ชื่อรายงาน (center) ──
		sf(11)
		w, _ = pdf.MeasureTextWidth(titleTH)
		pdf.SetXY((pageW-w)/2, y)
		pdf.Cell(nil, titleTH)
		addNl(15)

		// ── วันที่ (ชิดซ้าย) | เลขผู้เสียภาษีบริษัท (ชิดขวา) ──
		sf(fs)
		pdf.SetXY(lm, y)
		pdf.Cell(nil, fmt.Sprintf("ณ วันที่  %s", pEnd.Format("02/01/2006")))
		taxStr := fmt.Sprintf("เลขประจำตัวผู้เสียภาษีอากร    %s", comTax)
		tw, _ := pdf.MeasureTextWidth(taxStr)
		pdf.SetXY(pageW-lm-tw, y)
		pdf.Cell(nil, taxStr)
		addNl(13)

		// ── ที่อยู่บริษัท (center, เล็กลง) ──
		sf(7.5)
		addrW, _ := pdf.MeasureTextWidth(comAddr)
		if addrW > pageW-lm*2 {
			runes := []rune(comAddr)
			for addrW > pageW-lm*2 && len(runes) > 0 {
				runes = runes[:len(runes)-1]
				addrW, _ = pdf.MeasureTextWidth(string(runes))
			}
			comAddr = string(runes)
		}
		pdf.SetXY((pageW-addrW)/2, y)
		pdf.Cell(nil, comAddr)
		addNl(13)

		// ── เส้นบนสุดของ Header ──
		yHdrTop = y
		hln(lm, pageW-lm, y, 0.5)
		addNl(2)

		// ── Column headers 2 rows ──
		sf(fsH)
		// row 1: group headers
		putC("ในกำกับภาษี", lm, cRef)
		var pdfHdrParty string
		if vatType == "P" {
			pdfHdrParty = "ชื่อผู้ขาย/ผู้ให้บริการ"
		} else {
			pdfHdrParty = "ชื่อผู้ซื้อ/ผู้รับบริการ"
		}
		putC(pdfHdrParty, cRef, cName)
		putC("เลขประจำตัวผู้เสียภาษี", cName, cTaxID)
		putC("สนง./สาขาที่", cTaxID, cBranch)
		putC("มูลค่าสินค้า", cBranch, cBase)
		putC("จำนวนเงิน", cBase, cVat)
		addNl(fsH + 2)

		// เส้นคั่นกลางเฉพาะใต้ "ในกำกับภาษี"
		yHdrMid = y
		hln(lm, cRef, y, 0.5)
		addNl(2)

		// row 2: sub headers
		putC("วันเดือนปี", lm, cDate)
		putC("เลขที่", cDate, cRef)
		putC("", cRef, cName)
		var pdfHdrTaxSub string
		if vatType == "P" {
			pdfHdrTaxSub = "ของผู้ขาย/ผู้ให้บริการ"
		} else {
			pdfHdrTaxSub = "ของผู้ซื้อ/ผู้รับบริการ"
		}
		putC(pdfHdrTaxSub, cName, cTaxID)
		putC("", cTaxID, cBranch)
		putC("หรือบริการ", cBranch, cBase)
		putC("ภาษีมูลค่าเพิ่ม", cBase, cVat)
		addNl(fsH + 2)

		// เส้นล่างสุดของ Header
		yHdrBot = y
		hln(lm, pageW-lm, y, 0.5)
		addNl(3)

		// วาดเส้นแนวตั้งใน Header
		pdf.SetLineWidth(0.5)
		pdf.SetStrokeColor(0, 0, 0)              // สีดำ
		pdf.Line(cDate, yHdrMid, cDate, yHdrBot) // เส้นคั่นระหว่าง วันเดือนปี กับ เลขที่ (เริ่มจากเส้นกลาง)
		pdf.Line(cRef, yHdrTop, cRef, yHdrBot)
		pdf.Line(cName, yHdrTop, cName, yHdrBot)
		pdf.Line(cTaxID, yHdrTop, cTaxID, yHdrBot)
		pdf.Line(cBranch, yHdrTop, cBranch, yHdrBot)
		pdf.Line(cBase, yHdrTop, cBase, yHdrBot)
	}

	printHeader()

	var sumBase, sumVat float64
	rowsPerPage := 38

	for idx, vr := range vatRows {
		if idx > 0 && idx%rowsPerPage == 0 {
			// วาดเส้นแนวตั้งก่อนขึ้นหน้าใหม่ (จากใต้ Header ถึงบรรทัดสุดท้ายของหน้า)
			pdf.SetLineWidth(0.5)
			pdf.SetStrokeColor(0, 0, 0)
			for _, x := range []float64{cDate, cRef, cName, cTaxID, cBranch, cBase} {
				pdf.Line(x, yHdrBot, x, y)
			}

			pdf.AddPage()
			y = 25.0
			pageNo++
			printHeader()
		}
		sf(fs)
		// พิมพ์ข้อมูลโดยเว้นระยะ (Padding) ซ้ายขวา เพื่อไม่ให้ชิดเส้นตาราง
		putL(vr.Bdate, lm+2, cDate-lm-4)
		putL(vr.Bref, cDate+2, cRef-cDate-4)
		putL(vr.Boff, cRef+2, cName-cRef-4)
		putR(vr.Bcomtaxid, cTaxID-2)
		putR("00000", cBranch-2)
		putR(formatComma(vr.BaseAmt), cBase-2)
		putR(formatComma(vr.VatAmt), cVat-2)
		addNl(fs + 4)
		sumBase += vr.BaseAmt
		sumVat += vr.VatAmt
	}

	// วาดเส้นแนวตั้งหลังจบ loop ทีเดียวยาวตลอด (จากใต้ Header ถึงบรรทัดสุดท้ายของข้อมูล)
	pdf.SetLineWidth(0.5)
	pdf.SetStrokeColor(0, 0, 0)
	for _, x := range []float64{cDate, cRef, cName, cTaxID, cBranch, cBase} {
		pdf.Line(x, yHdrBot, x, y)
	}

	// total
	hln(lm, pageW-lm, y, 0.5) // เส้นบนของแถวรวม
	addNl(4)
	sf(fsH)
	pdf.SetXY(lm+2, y)
	pdf.Cell(nil, fmt.Sprintf("รวมทั้งสิ้น   %d   รายการ", len(vatRows)))
	putR(formatComma(sumBase), cBase-2)
	putR(formatComma(sumVat), cVat-2)
	addNl(fsH + 4)
	hln(lm, pageW-lm, y, 0.5) // เส้นล่างของแถวรวม

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// showVatDialog — ModalPopUp เหมือน Balance Sheet
// ─────────────────────────────────────────────────────────────────
func showVatDialog(w fyne.Window, vatType string, onGoSetup func()) {
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
				widget.NewSeparator(),
				container.NewCenter(container.NewHBox(btnGo, btnCancel2)),
			), w)
		warn.Show()
		w.Canvas().Focus(btnGo)
		return
	}

	var titleTH, titleEN, prefix string
	if vatType == "P" {
		titleTH = "รายงานภาษีซื้อ"
		titleEN = "Purchases VAT Report"
		prefix = "PurchasesVAT"
	} else {
		titleTH = "รายงานภาษีขาย"
		titleEN = "Sales VAT Report"
		prefix = "SalesVAT"
	}

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	var opts []string
	for _, p := range periods[:showUpTo] {
		opts = append(opts, fmt.Sprintf("งวด %d  (%s - %s)",
			p.PNo, p.PStart.Format("02/01/06"), p.PEnd.Format("02/01/06")))
	}
	selPeriod := widget.NewSelect(opts, nil)
	selPeriod.SetSelected(opts[showUpTo-1])

	getPeriod := func() int {
		for i, o := range opts {
			if o == selPeriod.Selected {
				return i + 1
			}
		}
		return showUpTo
	}

	btnExcel := widget.NewButton("📊 Excel", nil)
	btnPDF := widget.NewButton("📄 PDF", nil)
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
			widget.NewLabelWithStyle(
				fmt.Sprintf("%s / %s", titleTH, titleEN),
				fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			widget.NewLabel("เลือก Period:"),
			selPeriod,
			widget.NewSeparator(),
			container.NewHBox(btnExcel, btnPDF, btnCancel),
		),
		w.Canvas(),
	)
	pop.Resize(fyne.NewSize(480, 190))

	w.Canvas().SetOnTypedKey(func(key *fyne.KeyEvent) {
		if key.Name == fyne.KeyEscape {
			closePopup()
		}
	})
	btnCancel.OnTapped = closePopup

	showDone := func(pathToOpen string) {
		isTmp := strings.HasSuffix(pathToOpen, ".tmp")
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
		pNo := getPeriod()
		closePopup()
		if isPDF {
			savePath := filepath.Join(reportDir, fmt.Sprintf("%s_P%02d.pdf", prefix, pNo))
			go func() {
				pathToOpen, err := exportVatPDF(xlOpts, pNo, vatType, savePath)
				fyne.Do(func() {
					if err != nil {
						dialog.ShowError(fmt.Errorf("Export PDF ล้มเหลว: %v", err), w)
					} else {
						showDone(pathToOpen)
					}
				})
			}()
		} else {
			savePath := filepath.Join(reportDir, fmt.Sprintf("%s_P%02d.xlsx", prefix, pNo))
			go func() {
				pathToOpen, err := exportVatExcel(xlOpts, pNo, vatType, savePath)
				fyne.Do(func() {
					if err != nil {
						dialog.ShowError(fmt.Errorf("Export Excel ล้มเหลว: %v", err), w)
					} else {
						showDone(pathToOpen)
					}
				})
			}()
		}
	}

	btnExcel.OnTapped = func() { run(false) }
	btnPDF.OnTapped = func() { run(true) }
	pop.Show()
	w.Canvas().Focus(nil)
}
