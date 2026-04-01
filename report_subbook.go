package main

// report_subbook.go
// Subsidiary Books Report — สมุดรายวัน
// Layout: A4 Portrait — ตรงตาม VB.NET / Crystal Report ต้นแบบ
// Column order: รายการที่ | วันที่ | ใบคุม | เลขที่เอกสาร | เลขที่เช็ค | วันที่ส่งจ่าย | คำอธิบาย | เดบิต | เครดิต
// Account lines: รหัสบัญชี (col1) | ชื่อบัญชี (col7) | เดบิต | เครดิต

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
// SubbookLine — 1 account line ใน voucher
// ─────────────────────────────────────────────────────────────────
type SubbookLine struct {
	AcCode string
	AcName string
	Debit  float64
	Credit float64
}

// SubbookVoucher — 1 voucher (BITEM) พร้อม lines
type SubbookVoucher struct {
	Bitem     string // col3  เลขที่ sequence (001, 002...)
	Bdate     string // col1  วันที่
	Bvoucher  string // col2  เลขที่เอกสาร เช่น RV12/001
	Bref      string // col11 ใบคุมเอกสาร เช่น 6312-01
	Bnote     string // col14 คำอธิบาย / ชื่อลูกค้า
	Lines     []SubbookLine
	SumDebit  float64
	SumCredit float64
}

// ─────────────────────────────────────────────────────────────────
// parseSubbookDate
// ─────────────────────────────────────────────────────────────────
func parseSubbookDate(s string) (time.Time, error) {
	for _, layout := range []string{"02/01/2006", "02/01/06", "2006-01-02"} {
		if t, err := time.Parse(layout, s); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("parse date: %q", s)
}

// ─────────────────────────────────────────────────────────────────
// loadSubbookData — อ่าน Book_items แล้ว group ตาม Bitem
// periodNo = 0 → All Periods
// ─────────────────────────────────────────────────────────────────
func loadSubbookData(xlOptions excelize.Options, periodNo int) ([]SubbookVoucher, CompanyPeriodConfig, error) {
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		return nil, cfg, err
	}

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil, cfg, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	comCode := getComCodeFromExcel(xlOptions)

	// ── โหลด acName map จาก Ledger_Master (เปิดไฟล์แค่ครั้งเดียว) ──
	acctMap := getAccountNameMapFromFile(f, comCode)

	var periodStart, periodEnd time.Time
	allPeriods := (periodNo == 0)
	if !allPeriods {
		periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
		if periodNo < 1 || periodNo > len(periods) {
			return nil, cfg, fmt.Errorf("period %d ไม่ถูกต้อง", periodNo)
		}
		periodStart = periods[periodNo-1].PStart
		periodEnd = periods[periodNo-1].PEnd
	}

	rows, _ := f.GetRows("Book_items")
	voucherMap := make(map[string]*SubbookVoucher)
	var voucherOrder []string

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) < 11 {
			continue
		}
		if safeGet(row, 0) != comCode {
			continue
		}
		if !allPeriods {
			bdate, err2 := parseSubbookDate(safeGet(row, 1))
			if err2 != nil {
				continue
			}
			if bdate.Before(periodStart) || bdate.After(periodEnd) {
				continue
			}
		}

		bitem := safeGet(row, 3)
		if bitem == "" {
			continue
		}
		if _, ok := voucherMap[bitem]; !ok {
			voucherMap[bitem] = &SubbookVoucher{
				Bitem:    bitem,
				Bdate:    safeGet(row, 1),
				Bvoucher: safeGet(row, 2),
				Bref:     safeGet(row, 11),
				Bnote:    safeGet(row, 14),
			}
			voucherOrder = append(voucherOrder, bitem)
		}

		acCode := safeGet(row, 5)

		// ── Fallback: ใช้ชื่อจาก Ledger_Master ก่อน
		//    ถ้าไม่พบ (account ถูกลบแล้ว) ใช้ snapshot จาก Book_items col 6
		//    และเพิ่มวงเล็บ [] บ่งบอกว่าชื่อนี้มาจาก snapshot เก่า ──
		acName := acctMap[acCode]
		if acName == "" {
			acName = strings.TrimSpace(safeGet(row, 6))
			if acName != "" {
				acName = "[" + acName + "]"
			}
		}

		dr := parseFloat(safeGet(row, 9))
		cr := parseFloat(safeGet(row, 10))
		v := voucherMap[bitem]
		v.Lines = append(v.Lines, SubbookLine{
			AcCode: acCode,
			AcName: acName,
			Debit:  dr,
			Credit: cr,
		})
		v.SumDebit += dr
		v.SumCredit += cr
	}

	sort.Slice(voucherOrder, func(i, j int) bool {
		vi := voucherMap[voucherOrder[i]]
		vj := voucherMap[voucherOrder[j]]
		di, _ := parseSubbookDate(vi.Bdate)
		dj, _ := parseSubbookDate(vj.Bdate)
		if di.Equal(dj) {
			return voucherOrder[i] < voucherOrder[j]
		}
		return di.Before(dj)
	})

	result := make([]SubbookVoucher, 0, len(voucherOrder))
	for _, key := range voucherOrder {
		v := *voucherMap[key]
		sortSubbookLines(v.Lines)
		result = append(result, v)
	}
	return result, cfg, nil
}

// sortSubbookLines — เรียง account lines ภายใน voucher ตามหมวดบัญชีสากล
// (1=สินทรัพย์ → 2=หนี้สิน → 3=ทุน → 4=รายได้ → 5=ค่าใช้จ่าย → 6+=อื่นๆ)
func sortSubbookLines(lines []SubbookLine) {
	acGroup := func(acCode string) int {
		if len(acCode) == 0 {
			return 9
		}
		switch acCode[0] {
		case '1':
			return 1
		case '2':
			return 2
		case '3':
			return 3
		case '4':
			return 4
		case '5':
			return 5
		default:
			return 6
		}
	}
	for i := 1; i < len(lines); i++ {
		key := lines[i]
		gi := acGroup(key.AcCode)
		j := i - 1
		for j >= 0 {
			gj := acGroup(lines[j].AcCode)
			if gj > gi || (gj == gi && lines[j].AcCode > key.AcCode) {
				lines[j+1] = lines[j]
				j--
			} else {
				break
			}
		}
		lines[j+1] = key
	}
}

// truncRune — ตัด string ถ้าเกิน maxRune
func truncRune(s string, maxRune int) string {
	runes := []rune(s)
	if len(runes) <= maxRune {
		return s
	}
	return string(runes[:maxRune]) + ".."
}

// ─────────────────────────────────────────────────────────────────
// exportSubbookPDF — A4 Portrait
//
// Column layout (ตรง VB.NET ต้นแบบ):
//
//	xSeq   xDate   xRef     xVoucher  xCheque  xCheqDt  xDesc        xDebit  xCredit
//	รายการที่ วันที่  ใบคุม   เลขที่เอกสาร เลขที่เช็ค วันที่ส่งจ่าย  คำอธิบาย  เดบิต  เครดิต
//	รหัสบัญชี         ชื่อบัญชี(col2)                              ชื่อบัญชี(col7)
//
// Header row 1 (top labels), row 2 (bottom labels)
// Voucher header: seq|date|ref|voucher|cheque|cheqdt|note  (ไม่มีตัวเลข)
// Account lines:  accode(xSeq)|acname(xDesc)|debit|credit
// Subtotal:       เส้น + ตัวเลขชิดขวา
// ─────────────────────────────────────────────────────────────────
// dottedLine — วาดเส้น dotted โดย loop วาด dot เล็กๆ ห่างกัน
func dottedLine(pdf *gopdf.GoPdf, x1, y, x2, dotLen, gap, lineWidth float64) {
	pdf.SetLineWidth(lineWidth)
	for x := x1; x < x2; x += dotLen + gap {
		end := x + dotLen
		if end > x2 {
			end = x2
		}
		pdf.Line(x, y, end, y)
	}
}

func exportSubbookPDF(xlOptions excelize.Options, periodNo int, savePath string) (string, error) {
	vouchers, cfg, err := loadSubbookData(xlOptions, periodNo)
	if err != nil {
		return "", err
	}

	f2, err2 := excelize.OpenFile(currentDBPath, xlOptions)
	if err2 != nil {
		return "", fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err2)
	}
	comName, _ := f2.GetCellValue("Company_Profile", "B2")
	comAddr, _ := f2.GetCellValue("Company_Profile", "C2")
	comTax, _ := f2.GetCellValue("Company_Profile", "D2")
	f2.Close()

	var periodLabel string
	if periodNo == 0 {
		periodLabel = "ทุกงวด (All Periods)"
	} else {
		periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
		p := periods[periodNo-1]
		periodLabel = fmt.Sprintf("ณ วันที่ %s", p.PEnd.Format("02/01/2006"))
	}

	// ─── Font ──────────────────────────────────────────────────────
	userFontsDir := filepath.Join(os.Getenv("LOCALAPPDATA"), "Microsoft", "Windows", "Fonts")
	sysFontsDir := filepath.Join(os.Getenv("WINDIR"), "Fonts")
	if sysFontsDir == "Fonts" || sysFontsDir == "" {
		sysFontsDir = filepath.Join("C:", "Windows", "Fonts")
	}
	fontsDir := userFontsDir
	if _, statErr := os.Stat(filepath.Join(userFontsDir, "Sarabun-Regular.ttf")); os.IsNotExist(statErr) {
		fontsDir = sysFontsDir
	}
	fontPath := ""
	for _, name := range []string{"Sarabun-Regular.ttf", "Sarabun-Medium.ttf", "THSarabunNew.ttf", "arial.ttf", "Arial.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	if fontPath == "" {
		return "", fmt.Errorf("ไม่พบ font ไทยใน %s", fontsDir)
	}
	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "Sarabun-SemiBold.ttf", "THSarabunNew Bold.ttf", "arialbd.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	// ─── A4 Portrait layout constants ─────────────────────────────
	//
	// ต้นแบบ VB.NET (portrait ~595pt wide, usable ~535pt):
	//
	//  xSeq=30  xDate=65  xRef=117  xVoucher=175  xCheque=248  xCheqDt=308  xDesc=370  xDebit=470  xCredit=535(end)
	//  w=35     w=52      w=58      w=73           w=60         w=62          w=100      w=65         w=~60
	//
	// Account lines: AcCode วางที่ xSeq, AcName วางที่ xDesc
	//
	const (
		pageW   = 595.28
		pageH   = 841.89
		marginL = 30.0
		marginR = 20.0
		marginT = 28.0
		marginB = 30.0
		rowH    = 13.0
		rowHHdr = 11.0

		// Column X & W
		xSeq    = marginL // รายการที่ / รหัสบัญชี
		wSeq    = 35.0
		xDate   = xSeq + wSeq // วันที่
		wDate   = 52.0
		xRef    = xDate + wDate // ใบคุมเอกสาร
		wRef    = 58.0
		xVch    = xRef + wRef // เลขที่เอกสาร / ชื่อบัญชี(line)
		wVch    = 65.0
		xCheque = xVch + wVch // เลขที่เช็ค
		wCheque = 40.0
		xCheqDt = xCheque + wCheque // วันที่ส่งจ่าย
		wCheqDt = 40.0
		xDesc   = xCheqDt + wCheqDt               // คำอธิบาย / ชื่อบัญชี(line)
		wDesc   = pageW - marginR - xDesc - 130.0 // ~80pt

		// Debit / Credit ชิดขวา
		xDebit  = pageW - marginR - 160.0
		wDebit  = 80.0
		xCredit = xDebit + wDebit
		wCredit = 80.0
	)

	pdf := gopdf.GoPdf{}
	ps := gopdf.Rect{W: pageW, H: pageH}
	pdf.Start(gopdf.Config{PageSize: ps, Unit: gopdf.UnitPT})
	if err = pdf.AddTTFFont("normal", fontPath); err != nil {
		return "", fmt.Errorf("โหลด font: %v", err)
	}
	pdf.AddTTFFont("bold", boldPath)

	setFont := func(bold bool, size float64) {
		n := "normal"
		if bold {
			n = "bold"
		}
		pdf.SetFont(n, "", size)
	}

	pageNo := 0
	y := pageH // trigger first page

	// ─── printPageHeader ───────────────────────────────────────────
	printPageHeader := func() {
		pageNo++
		pdf.AddPageWithOption(gopdf.PageOption{PageSize: &gopdf.Rect{W: pageW, H: pageH}})
		y = marginT

		// ชื่อบริษัท (bold, center)
		setFont(true, 13)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - marginL - marginR, H: 15},
			comName, gopdf.CellOption{Align: gopdf.Center})
		// "Page N" ขวาบน
		setFont(false, 9)
		pdf.SetXY(pageW-marginR-55, y)
		pdf.CellWithOption(&gopdf.Rect{W: 55, H: 15},
			fmt.Sprintf("Page:  %d", pageNo), gopdf.CellOption{Align: gopdf.Right})
		y += 15

		// ที่อยู่ / Tax ID
		if comAddr != "" || comTax != "" {
			setFont(false, 9)
			pdf.SetXY(marginL, y)
			addrLine := comAddr
			if comTax != "" {
				addrLine = fmt.Sprintf("เลขประจำตัวผู้เสียภาษีอากร %s", comTax)
			}
			pdf.CellWithOption(&gopdf.Rect{W: pageW - marginL - marginR, H: 11},
				addrLine, gopdf.CellOption{Align: gopdf.Center})
			y += 11
		}

		// ชื่อรายงาน
		setFont(false, 11)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - marginL - marginR, H: 14},
			"รายการรายวัน เรียงตามลำดับที่", gopdf.CellOption{Align: gopdf.Center})
		y += 14

		// period label
		setFont(false, 9)
		pdf.SetXY(marginL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - marginL - marginR, H: 12},
			periodLabel, gopdf.CellOption{Align: gopdf.Center})
		y += 12 + 4

		// ─── Column headers — 2 แถว (ตรงต้นแบบ VB.NET) ─────────────
		pdf.SetLineWidth(0.6)
		pdf.Line(marginL, y, pageW-marginR, y)
		y += 3

		setFont(false, 8)
		// แถว 1: top labels
		cells1 := []struct {
			x, w float64
			txt  string
		}{
			{xSeq, wSeq, "รายการที่"},
			{xDate, wDate, "วันที่"},
			{xRef, wRef, "ใบคุมเอกสาร"},
			{xVch, wVch, "เลขที่เอกสาร"},
			{xCheque, wCheque, "เลขที่เช็ค"},
			{xCheqDt, wCheqDt, "วันที่สั่งจ่าย"},
			{xDesc, wDesc, "คำอธิบาย"},
			{xDebit, wDebit, "เดบิต"},
			{xCredit, wCredit, "เครดิต"},
		}
		for _, c := range cells1 {
			pdf.SetXY(c.x, y)
			align := gopdf.Left
			if c.x >= xDebit {
				align = gopdf.Right
			}
			pdf.CellWithOption(&gopdf.Rect{W: c.w, H: rowHHdr}, c.txt, gopdf.CellOption{Align: align})
		}
		y += rowHHdr

		// แถว 2: bottom labels
		cells2 := []struct {
			x, w float64
			txt  string
		}{
			{xSeq, wSeq, "รหัสบัญชี"},
			{xVch, wVch, "ชื่อบัญชี"},
		}
		for _, c := range cells2 {
			pdf.SetXY(c.x, y)
			pdf.CellWithOption(&gopdf.Rect{W: c.w, H: rowHHdr}, c.txt, gopdf.CellOption{})
		}
		y += rowHHdr

		pdf.SetLineWidth(0.6)
		pdf.Line(marginL, y, pageW-marginR, y)
		y += 3
	}

	checkNewPage := func(needed float64) {
		if y+needed > pageH-marginB {
			printPageHeader()
		}
	}

	printPageHeader()

	// ─── Accumulators ──────────────────────────────────────────────
	var grandDebit, grandCredit float64
	var countDR, countCR int

	// ─── Print each voucher ────────────────────────────────────────
	for _, v := range vouchers {
		estH := rowH + float64(len(v.Lines))*rowH + rowH + 5
		if estH > pageH-marginT-marginB-70 {
			estH = rowH * 4
		}
		checkNewPage(estH)

		// ── Voucher header row ──────────────────────────────────────
		// รายการที่ | วันที่ | ใบคุม(Bref) | เลขที่เอกสาร(Bvoucher) | คำอธิบาย(Bnote)
		setFont(false, 8.5)
		pdf.SetXY(xSeq, y)
		pdf.CellWithOption(&gopdf.Rect{W: wSeq, H: rowH}, v.Bitem, gopdf.CellOption{})
		pdf.SetXY(xDate, y)
		pdf.CellWithOption(&gopdf.Rect{W: wDate, H: rowH}, v.Bdate, gopdf.CellOption{})
		pdf.SetXY(xRef, y)
		pdf.CellWithOption(&gopdf.Rect{W: wRef, H: rowH}, v.Bvoucher, gopdf.CellOption{})
		pdf.SetXY(xVch, y)
		pdf.CellWithOption(&gopdf.Rect{W: wVch, H: rowH}, v.Bref, gopdf.CellOption{})
		// Bnote อยู่ใน column คำอธิบาย (xDesc)
		pdf.SetXY(xDesc, y)
		pdf.CellWithOption(&gopdf.Rect{W: wDesc + 130, H: rowH}, truncRune(v.Bnote, 30), gopdf.CellOption{})
		y += rowH

		// ── Account lines ───────────────────────────────────────────
		// AcCode ที่ xSeq, AcName ที่ xVch
		for _, line := range v.Lines {
			checkNewPage(rowH)
			setFont(false, 8.5)
			pdf.SetXY(xSeq, y)
			pdf.CellWithOption(&gopdf.Rect{W: wSeq + wDate, H: rowH}, line.AcCode, gopdf.CellOption{})
			pdf.SetXY(xVch, y)
			pdf.CellWithOption(&gopdf.Rect{W: wVch + wCheque + wCheqDt + wDesc, H: rowH}, truncRune(line.AcName, 60), gopdf.CellOption{})
			if line.Debit != 0 {
				pdf.SetXY(xDebit, y)
				pdf.CellWithOption(&gopdf.Rect{W: wDebit, H: rowH}, formatNum(line.Debit), gopdf.CellOption{Align: gopdf.Right})
				countDR++
			}
			if line.Credit != 0 {
				pdf.SetXY(xCredit, y)
				pdf.CellWithOption(&gopdf.Rect{W: wCredit, H: rowH}, formatNum(line.Credit), gopdf.CellOption{Align: gopdf.Right})
				countCR++
			}
			y += rowH
		}

		// ── Subtotal — เส้น dot บนๆ + เส้น solid ล่าง ─────────────
		checkNewPage(rowH + 10)

		// เส้น dotted บน (เหนือตัวเลข)
		// ── Subtotal — เส้น dot บนๆ + เส้น solid ล่าง ─────────────
		checkNewPage(rowH + 10)

		// เส้น dotted บน (เหนือตัวเลข)
		dotY := y
		dotStartX := xDebit - 2
		dotEndX := pageW - marginR
		dotSpacing := 2.0 // ระยะห่างระหว่างจุด (ปรับให้ถี่/ห่างได้ตามต้องการ)

		pdf.SetLineWidth(0.5) // ความหนาของจุด
		for dx := dotStartX; dx < dotEndX; dx += dotSpacing {
			// วาดเส้นสั้นมากๆ (0.5 pt) เพื่อให้ดูเหมือนจุด
			pdf.Line(dx, dotY, dx+0.5, dotY)
		}
		y += 2

		setFont(false, 8.5)
		pdf.SetXY(xDebit, y)
		pdf.CellWithOption(&gopdf.Rect{W: wDebit, H: rowH}, formatNum(v.SumDebit), gopdf.CellOption{Align: gopdf.Right})
		pdf.SetXY(xCredit, y)
		pdf.CellWithOption(&gopdf.Rect{W: wCredit, H: rowH}, formatNum(v.SumCredit), gopdf.CellOption{Align: gopdf.Right})
		y += rowH + 1

		// เส้น solid ล่าง (ใต้ตัวเลข)
		pdf.SetLineWidth(0.5)
		pdf.Line(xDebit-2, y, pageW-marginR, y)
		y += 5

		grandDebit += v.SumDebit
		grandCredit += v.SumCredit
	}

	// ─── Grand Total (ตรงต้นแบบ VB.NET) ─────────────────────────
	// DR  166   ITEM.   AMOUNT   17,554,438.15
	// CR  122   ITEM.   AMOUNT   17,554,438.15
	checkNewPage(rowH*3 + 10)
	pdf.SetLineWidth(0.6)
	pdf.Line(marginL, y, pageW-marginR, y)
	y += 3

	setFont(false, 8.5)
	// แถว DR
	pdf.SetXY(xVch, y)
	pdf.CellWithOption(&gopdf.Rect{W: wVch * 0.4, H: rowH}, "DR", gopdf.CellOption{})
	pdf.SetXY(xVch+wVch*0.4, y)
	pdf.CellWithOption(&gopdf.Rect{W: wVch * 0.6, H: rowH},
		fmt.Sprintf("%d", countDR), gopdf.CellOption{Align: gopdf.Right})
	pdf.SetXY(xCheque, y)
	pdf.CellWithOption(&gopdf.Rect{W: wCheque, H: rowH}, "ITEM.", gopdf.CellOption{Align: gopdf.Center})
	pdf.SetXY(xCheqDt, y)
	pdf.CellWithOption(&gopdf.Rect{W: wCheqDt + wDesc, H: rowH}, "AMOUNT", gopdf.CellOption{Align: gopdf.Center})
	pdf.SetXY(xDebit, y)
	pdf.CellWithOption(&gopdf.Rect{W: wDebit + wCredit, H: rowH},
		formatNum(grandDebit), gopdf.CellOption{Align: gopdf.Right})
	y += rowH

	// แถว CR
	pdf.SetXY(xVch, y)
	pdf.CellWithOption(&gopdf.Rect{W: wVch * 0.4, H: rowH}, "CR", gopdf.CellOption{})
	pdf.SetXY(xVch+wVch*0.4, y)
	pdf.CellWithOption(&gopdf.Rect{W: wVch * 0.6, H: rowH},
		fmt.Sprintf("%d", countCR), gopdf.CellOption{Align: gopdf.Right})
	pdf.SetXY(xCheque, y)
	pdf.CellWithOption(&gopdf.Rect{W: wCheque, H: rowH}, "ITEM.", gopdf.CellOption{Align: gopdf.Center})
	pdf.SetXY(xCheqDt, y)
	pdf.CellWithOption(&gopdf.Rect{W: wCheqDt + wDesc, H: rowH}, "AMOUNT", gopdf.CellOption{Align: gopdf.Center})
	pdf.SetXY(xDebit, y)
	pdf.CellWithOption(&gopdf.Rect{W: wDebit + wCredit, H: rowH},
		formatNum(grandCredit), gopdf.CellOption{Align: gopdf.Right})
	y += rowH

	pdf.SetLineWidth(0.6)
	pdf.Line(marginL, y, pageW-marginR, y)

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// showSubbookDialog — UI เลือก Period แล้ว export PDF
// ─────────────────────────────────────────────────────────────────
func showSubbookDialog(w fyne.Window, onGoSetup func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}

	reportDir := getReportDir(xlOptions)
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

	var opts []string
	opts = append(opts, "ทุกงวด (All Periods)")
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	for _, p := range periods[:showUpTo] {
		opts = append(opts, fmt.Sprintf("งวด %d  (%s - %s)",
			p.PNo, p.PStart.Format("02/01/06"), p.PEnd.Format("02/01/06")))
	}

	cbo := widget.NewSelect(opts, nil)
	if showUpTo < len(opts) {
		cbo.SetSelected(opts[showUpTo])
	} else {
		cbo.SetSelected(opts[0])
	}

	getPeriodNo := func() int {
		for i, o := range opts {
			if o == cbo.Selected {
				return i
			}
		}
		return 0
	}

	// showDone — แสดง success แล้วเปิดไฟล์เมื่อกด OK
	// ถ้า pathToOpen ลงท้าย .tmp แสดงว่าไฟล์เดิมถูกล็อกอยู่
	showDone := func(pathToOpen string) {
		ext2 := filepath.Ext(pathToOpen)
		isTmp := strings.HasSuffix(strings.TrimSuffix(pathToOpen, ext2), "_tmp")
		title := "✅ บันทึกรายงานแล้ว"
		note := ""
		if isTmp {
			title = "⚠️ เปิดรายงานชั่วคราว"
			note = "ปิดไฟล์ PDF เดิมก่อน แล้วกด Export ใหม่เพื่อบันทึกถาวร"
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

	var d dialog.Dialog
	btnOK := newEnterButton("สร้างรายงาน (Enter)", func() {
		periodNo := getPeriodNo()
		var fileName string
		if periodNo == 0 {
			fileName = "SubBook_AllPeriods.pdf"
		} else {
			fileName = fmt.Sprintf("SubBook_P%02d.pdf", periodNo)
		}
		savePath := filepath.Join(reportDir, fileName)
		d.Hide()
		pathToOpen, err2 := exportSubbookPDF(xlOptions, periodNo, savePath)
		if err2 != nil {
			showErrDialog(w, "สร้าง PDF ไม่ได้: "+err2.Error())
			return
		}
		showDone(pathToOpen)
	})
	btnOK.Importance = widget.HighImportance
	btnCancel := newEscButton("ยกเลิก (Esc)", func() { d.Hide() })

	form := container.NewVBox(
		widget.NewLabelWithStyle("สมุดรายวัน (Subsidiary Books)", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		widget.NewSeparator(),
		container.NewHBox(widget.NewLabel("เลือกงวด:"), cbo),
		widget.NewSeparator(),
		container.NewCenter(container.NewHBox(btnOK, btnCancel)),
	)
	d = dialog.NewCustomWithoutButtons("รายงานสมุดรายวัน", form, w)
	d.Resize(fyne.NewSize(480, 180))
	d.Show()
	w.Canvas().Focus(btnOK)
}
