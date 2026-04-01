package main

// report_ledger.go
// Ledger Report — บัญชีแยกประเภท (Layout ตามต้นแบบ VB.NET)

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// Data Structures
// ─────────────────────────────────────────────────────────────────
type LedgerLine struct {
	Date   string
	Book   string // สมุดบัญชี (เช่น รวร, รวจ) ดึงจากตัวอักษรนำหน้า Voucher
	Seq    string // ลำดับที่ (Bitem)
	Ref    string // อ้างอิง (Bvoucher / Bref)
	Desc   string // คำอธิบาย
	Debit  float64
	Credit float64
}

type LedgerAccount struct {
	AcCode      string
	AcName      string
	Lines       []LedgerLine
	TotalDebit  float64 // รวมเดบิตทั้งหมด (ยอดยกมา + เคลื่อนไหว)
	TotalCredit float64 // รวมเครดิตทั้งหมด (ยอดยกมา + เคลื่อนไหว)
	CloseDebit  float64 // ยอดยกไป (ฝั่งเดบิต)
	CloseCredit float64 // ยอดยกไป (ฝั่งเครดิต)
	GrandTotal  float64 // ยอดรวมสุทธิที่ดุลกัน (บรรทัด "รวม")
}

// ─────────────────────────────────────────────────────────────────
// Core Logic: ดึงข้อมูลและคำนวณยอดยกมา/ยกไป ให้ดุลกัน
// ─────────────────────────────────────────────────────────────────
func buildLedgerData(xlOptions excelize.Options, pFrom, pTo int, glFrom, glTo string) ([]LedgerAccount, CompanyPeriodConfig, error) {
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
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)

	if pFrom < 1 {
		pFrom = 1
	}
	if pTo > len(periods) {
		pTo = len(periods)
	}
	dateFrom := periods[pFrom-1].PStart
	dateTo := periods[pTo-1].PEnd

	type bItem struct {
		acCode string
		date   time.Time
		dateSt string
		vch    string
		ref    string
		seq    string
		scode  string
		note   string
		dr     float64
		cr     float64
	}
	var allTx []bItem
	bRows, _ := f.GetRows("Book_items")
	for i, r := range bRows {
		if i == 0 || len(r) < 11 || safeGet(r, 0) != comCode {
			continue
		}
		ac := safeGet(r, 5)
		if ac < glFrom || ac > glTo {
			continue
		}
		dtStr := safeGet(r, 1)
		dt, err := parseSubbookDate(dtStr)
		if err != nil {
			continue
		}
		allTx = append(allTx, bItem{
			acCode: ac,
			date:   dt,
			dateSt: dtStr,
			vch:    safeGet(r, 2),
			seq:    safeGet(r, 3),
			scode:  safeGet(r, 7),
			ref:    safeGet(r, 11),
			note:   safeGet(r, 14),
			dr:     parseFloat(safeGet(r, 9)),
			cr:     parseFloat(safeGet(r, 10)),
		})
	}

	var result []LedgerAccount
	lRows, _ := f.GetRows("Ledger_Master")

	for i, r := range lRows {
		if i == 0 || len(r) < 3 || safeGet(r, 0) != comCode {
			continue
		}
		acCode := safeGet(r, 1)
		acName := safeGet(r, 2)
		if acCode < glFrom || acCode > glTo {
			continue
		}

		isPL := strings.HasPrefix(acCode, "4") || strings.HasPrefix(acCode, "5") || strings.HasPrefix(acCode, "6")

		// 360PLA — FoxPro PRNAC logic:
		// ไม่มี transaction ตรงใน GLDAT แสดงแค่ยอดยกมา/ยอดยกไป 0.00
		if acCode == "360PLA" {
			acc := LedgerAccount{AcCode: acCode, AcName: acName}
			acc.Lines = append(acc.Lines, LedgerLine{
				Date: dateFrom.Format("02/01/06"), Seq: "0", Desc: "ยอดยกมา",
			})
			acc.Lines = append(acc.Lines, LedgerLine{
				Date: dateTo.Format("02/01/06"), Desc: "ยอดยกไป",
			})
			result = append(result, acc)
			continue
		}

		openingBal := 0.0
		if !isPL {
			openingBal = parseFloat(safeGet(r, 9))
			for _, tx := range allTx {
				if tx.acCode == acCode && tx.date.Before(dateFrom) {
					openingBal += tx.dr - tx.cr
				}
			}
		}

		var currentTx []bItem
		for _, tx := range allTx {
			if tx.acCode == acCode && !tx.date.Before(dateFrom) && !tx.date.After(dateTo) {
				currentTx = append(currentTx, tx)
			}
		}

		if openingBal == 0 && len(currentTx) == 0 {
			continue
		}

		sort.Slice(currentTx, func(i, j int) bool {
			return currentTx[i].date.Before(currentTx[j].date)
		})

		acc := LedgerAccount{AcCode: acCode, AcName: acName}

		// 1. ยอดยกมา
		openDr, openCr := 0.0, 0.0
		if openingBal > 0 {
			openDr = openingBal
		} else if openingBal < 0 {
			openCr = -openingBal
		}
		acc.Lines = append(acc.Lines, LedgerLine{
			Date:   dateFrom.Format("02/01/06"),
			Seq:    "0", // ใส่ลำดับที่ 0 สำหรับยอดยกมา
			Desc:   "ยอดยกมา",
			Debit:  openDr,
			Credit: openCr,
		})
		acc.TotalDebit += openDr
		acc.TotalCredit += openCr

		// 2. รายการเคลื่อนไหว
		for _, tx := range currentTx {
			acc.Lines = append(acc.Lines, LedgerLine{
				Date:   tx.dateSt,
				Book:   tx.scode, // <--- ใช้ tx.scode ที่ดึงมาจากตารางโดยตรง
				Seq:    tx.seq,
				Ref:    tx.ref,
				Desc:   tx.note,
				Debit:  tx.dr,
				Credit: tx.cr,
			})
			acc.TotalDebit += tx.dr
			acc.TotalCredit += tx.cr
		}

		// 3. คำนวณยอดยกไป (Balancing Figure)
		if acc.TotalDebit > acc.TotalCredit {
			acc.CloseCredit = acc.TotalDebit - acc.TotalCredit
			acc.GrandTotal = acc.TotalDebit
		} else {
			acc.CloseDebit = acc.TotalCredit - acc.TotalDebit
			acc.GrandTotal = acc.TotalCredit
		}

		descClose := "ยอดยกไป"
		if isPL {
			descClose = "ปิดบัญชี"
		}
		acc.Lines = append(acc.Lines, LedgerLine{
			Date:   dateTo.Format("02/01/06"),
			Desc:   descClose,
			Debit:  acc.CloseDebit,
			Credit: acc.CloseCredit,
		})

		result = append(result, acc)
	}

	sortLedgerAccounts(result)
	return result, cfg, nil
}

// sortLedgerAccounts — เรียงตามหมวดบัญชีสากล: 1=สินทรัพย์, 2=หนี้สิน, 3=ทุน, 4=รายได้, 5=ค่าใช้จ่าย, 6+=อื่นๆ
// ภายในแต่ละหมวดเรียงตาม AcCode ตามลำดับ
func sortLedgerAccounts(accounts []LedgerAccount) {
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
	for i := 1; i < len(accounts); i++ {
		key := accounts[i]
		gi := acGroup(key.AcCode)
		j := i - 1
		for j >= 0 {
			gj := acGroup(accounts[j].AcCode)
			if gj > gi || (gj == gi && accounts[j].AcCode > key.AcCode) {
				accounts[j+1] = accounts[j]
				j--
			} else {
				break
			}
		}
		accounts[j+1] = key
	}
}

// ─────────────────────────────────────────────────────────────────
// Export PDF (Layout ตามรูปต้นแบบ)
// ─────────────────────────────────────────────────────────────────
func exportLedgerPDF(xlOptions excelize.Options, pFrom, pTo int, glFrom, glTo, savePath string) (string, error) {
	data, cfg, err := buildLedgerData(xlOptions, pFrom, pTo, glFrom, glTo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	dFrom := periods[pFrom-1].PStart.Format("02/01/06")
	dTo := periods[pTo-1].PEnd.Format("02/01/06")

	// ─── Font Setup ───
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
	for _, name := range []string{"Sarabun-Regular.ttf", "THSarabunNew.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			fontPath = p
			break
		}
	}
	boldPath := fontPath
	for _, name := range []string{"Sarabun-Bold.ttf", "THSarabunNew Bold.ttf"} {
		if p := filepath.Join(fontsDir, name); fileExists(p) {
			boldPath = p
			break
		}
	}

	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4, Unit: gopdf.UnitPT}) // A4 Portrait
	pdf.AddTTFFont("normal", fontPath)
	pdf.AddTTFFont("bold", boldPath)

	const (
		pageW = 595.28
		pageH = 841.89
		mL    = 30.0
		mR    = 30.0
		mT    = 40.0
		mB    = 40.0
		rowH  = 14.0

		// Column X positions (ขยายให้รองรับหลักหมื่นล้าน)
		xDate = mL
		xBook = xDate + 50
		xSeq  = xBook + 40
		xRef  = xSeq + 40
		xDesc = xRef + 60        // ลดความกว้างอ้างอิงลงนิดนึง
		xDr   = pageW - mR - 170 // ขยับเดบิตไปทางซ้ายเพื่อให้มีพื้นที่กว้างขึ้น (กว้าง 85 pt)
		xCr   = pageW - mR - 85  // ขยับเครดิตไปทางซ้ายเพื่อให้มีพื้นที่กว้างขึ้น (กว้าง 85 pt)
	)

	pageNo := 0
	y := pageH

	printHeader := func() {
		pageNo++
		pdf.AddPage()
		y = mT

		pdf.SetFont("bold", "", 12)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: 300, H: 15}, comName, gopdf.CellOption{Align: gopdf.Left})
		pdf.SetFont("normal", "", 10)
		pdf.SetXY(pageW-mR-50, y)
		pdf.CellWithOption(&gopdf.Rect{W: 50, H: 15}, fmt.Sprintf("Page: %d", pageNo), gopdf.CellOption{Align: gopdf.Right})
		y += 15

		pdf.SetFont("normal", "", 9)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: 300, H: 15}, "เลขประจำตัวผู้เสียภาษีอากร  "+comTax, gopdf.CellOption{Align: gopdf.Left})
		y += 25

		pdf.SetFont("normal", "", 11)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - mL - mR, H: 15}, "บัญชีแยกประเภท", gopdf.CellOption{Align: gopdf.Center})
		y += 15
		pdf.SetFont("normal", "", 9)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: pageW - mL - mR, H: 15}, fmt.Sprintf("ณ วันที่ %s ถึง %s", dFrom, dTo), gopdf.CellOption{Align: gopdf.Center})
		y += 25
	}

	checkPage := func(need float64) {
		if y+need > pageH-mB {
			printHeader()
		}
	}

	printHeader()

	for _, acc := range data {
		checkPage(rowH * 5)

		// Account Header
		pdf.SetFont("bold", "", 9)
		pdf.SetXY(mL, y)
		pdf.CellWithOption(&gopdf.Rect{W: 100, H: rowH}, acc.AcCode, gopdf.CellOption{})
		pdf.SetXY(mL+80, y)
		pdf.CellWithOption(&gopdf.Rect{W: 300, H: rowH}, acc.AcName, gopdf.CellOption{})
		y += rowH + 5

		// Table Header
		pdf.SetLineWidth(0.5)
		pdf.Line(mL, y, pageW-mR, y)
		y += 3
		pdf.SetFont("normal", "", 8.5)
		pdf.SetXY(xDate, y)
		pdf.CellWithOption(&gopdf.Rect{W: 50, H: rowH}, "วันที่", gopdf.CellOption{})
		pdf.SetXY(xBook, y)
		pdf.CellWithOption(&gopdf.Rect{W: 40, H: rowH}, "สมุดบัญชี", gopdf.CellOption{})
		pdf.SetXY(xSeq, y)
		pdf.CellWithOption(&gopdf.Rect{W: 40, H: rowH}, "ลำดับที่", gopdf.CellOption{Align: gopdf.Center})
		pdf.SetXY(xRef, y)
		pdf.CellWithOption(&gopdf.Rect{W: 70, H: rowH}, "อ้างอิง", gopdf.CellOption{})
		pdf.SetXY(xDesc, y)
		pdf.CellWithOption(&gopdf.Rect{W: 150, H: rowH}, "คำอธิบาย", gopdf.CellOption{})
		pdf.SetXY(xDr, y)
		pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, "เดบิต", gopdf.CellOption{Align: gopdf.Right}) // เปลี่ยน W เป็น 85
		pdf.SetXY(xCr, y)
		pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, "เครดิต", gopdf.CellOption{Align: gopdf.Right}) // เปลี่ยน W เป็น 85
		y += rowH

		// Dotted line under header
		for dx := mL; dx < pageW-mR; dx += 2.0 {
			pdf.Line(dx, y, dx+0.5, y)
		}
		y += 3

		// Lines
		pdf.SetFont("normal", "", 8.5)
		for _, line := range acc.Lines {
			checkPage(rowH)
			pdf.SetXY(xDate, y)
			pdf.CellWithOption(&gopdf.Rect{W: 50, H: rowH}, line.Date, gopdf.CellOption{})
			pdf.SetXY(xBook, y)
			pdf.CellWithOption(&gopdf.Rect{W: 40, H: rowH}, line.Book, gopdf.CellOption{})
			pdf.SetXY(xSeq, y)
			pdf.CellWithOption(&gopdf.Rect{W: 40, H: rowH}, line.Seq, gopdf.CellOption{Align: gopdf.Center})
			pdf.SetXY(xRef, y)
			pdf.CellWithOption(&gopdf.Rect{W: 70, H: rowH}, line.Ref, gopdf.CellOption{})

			desc := line.Desc
			if len([]rune(desc)) > 35 {
				desc = string([]rune(desc)[:35]) + ".."
			}
			pdf.SetXY(xDesc, y)
			pdf.CellWithOption(&gopdf.Rect{W: 180, H: rowH}, desc, gopdf.CellOption{})

			// ถ้าเป็นยอดยกมา และเป็น 0 ให้พิมพ์ 0.00 ฝั่งเดบิต
			if line.Desc == "ยอดยกมา" && line.Debit == 0 && line.Credit == 0 {
				pdf.SetXY(xDr, y)
				pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, "0.00", gopdf.CellOption{Align: gopdf.Right})
			} else {
				if line.Debit != 0 {
					pdf.SetXY(xDr, y)
					pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, formatNum(line.Debit), gopdf.CellOption{Align: gopdf.Right})
				}
				if line.Credit != 0 {
					pdf.SetXY(xCr, y)
					pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, formatNum(line.Credit), gopdf.CellOption{Align: gopdf.Right})
				}
			}
			y += rowH
		}

		// Footer (รวม)
		checkPage(rowH + 15)
		y += 2
		for dx := mL; dx < pageW-mR; dx += 2.0 {
			pdf.Line(dx, y, dx+0.5, y)
		} // Dotted line
		y += 3

		pdf.SetFont("normal", "", 8.5)
		pdf.SetXY(xDesc, y)
		pdf.CellWithOption(&gopdf.Rect{W: 100, H: rowH}, "รวม", gopdf.CellOption{})
		pdf.SetXY(xDr, y)
		pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, formatNum(acc.GrandTotal), gopdf.CellOption{Align: gopdf.Right})
		pdf.SetXY(xCr, y)
		pdf.CellWithOption(&gopdf.Rect{W: 85, H: rowH}, formatNum(acc.GrandTotal), gopdf.CellOption{Align: gopdf.Right})
		y += rowH + 2

		pdf.SetLineWidth(1.0) // Thick line
		pdf.Line(mL, y, pageW-mR, y)
		y += 20
	}

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return pdf.WritePdf(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// Export Excel (Layout ตามรูปต้นแบบ)
// ─────────────────────────────────────────────────────────────────
func exportLedgerExcel(xlOptions excelize.Options, pFrom, pTo int, glFrom, glTo, savePath string) (string, error) {
	data, cfg, err := buildLedgerData(xlOptions, pFrom, pTo, glFrom, glTo)
	if err != nil {
		return "", err
	}

	f, _ := excelize.OpenFile(currentDBPath, xlOptions)
	comName, _ := f.GetCellValue("Company_Profile", "B2")
	comTax, _ := f.GetCellValue("Company_Profile", "D2")
	f.Close()

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	dFrom := periods[pFrom-1].PStart.Format("02/01/06")
	dTo := periods[pTo-1].PEnd.Format("02/01/06")

	wb := excelize.NewFile()
	sh := "Ledger"
	wb.SetSheetName("Sheet1", sh)

	boldL, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true, Size: 11, Family: "TH Sarabun New"}})
	boldC, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true, Size: 12, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "center"}})
	hdrStyle, _ := wb.NewStyle(&excelize.Style{
		Font:   &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 3}}, // 3 = dashed
	})
	hdrRight, _ := wb.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "right"},
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 3}},
	})
	hdrCenter, _ := wb.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "center"},
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 1}, {Type: "bottom", Color: "000000", Style: 3}},
	})
	txtStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}})
	txtCenterStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "center"}})
	txtRightStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}, Alignment: &excelize.Alignment{Horizontal: "right"}})
	numStyle, _ := wb.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 10, Family: "TH Sarabun New"}, CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;""`)})

	totTxtStyle, _ := wb.NewStyle(&excelize.Style{
		Font:   &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"},
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 3}, {Type: "bottom", Color: "000000", Style: 2}}, // 2 = thick
	})
	totNumStyle, _ := wb.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 10, Family: "TH Sarabun New"}, CustomNumFmt: strPtr(`#,##0.00;-#,##0.00;0.00`),
		Border: []excelize.Border{{Type: "top", Color: "000000", Style: 3}, {Type: "bottom", Color: "000000", Style: 2}},
	})

	wb.SetColWidth(sh, "A", "A", 10) // วันที่
	wb.SetColWidth(sh, "B", "B", 8)  // สมุด
	wb.SetColWidth(sh, "C", "C", 8)  // ลำดับ
	wb.SetColWidth(sh, "D", "D", 15) // อ้างอิง
	wb.SetColWidth(sh, "E", "E", 30) // ลดคำอธิบายลงนิดนึง
	wb.SetColWidth(sh, "F", "G", 18) // ขยาย DR, CR ให้กว้างขึ้น (จาก 15 เป็น 18)

	// Page Header
	wb.SetCellValue(sh, "A1", comName)
	wb.SetCellStyle(sh, "A1", "A1", boldL)
	wb.SetCellValue(sh, "A2", "เลขประจำตัวผู้เสียภาษีอากร "+comTax)
	wb.SetCellStyle(sh, "A2", "A2", txtStyle)
	wb.MergeCell(sh, "A4", "G4")
	wb.SetCellValue(sh, "A4", "บัญชีแยกประเภท")
	wb.SetCellStyle(sh, "A4", "A4", boldC)
	wb.MergeCell(sh, "A5", "G5")
	wb.SetCellValue(sh, "A5", fmt.Sprintf("ณ วันที่ %s ถึง %s", dFrom, dTo))
	wb.SetCellStyle(sh, "A5", "A5", boldC)

	rowIdx := 7
	for _, acc := range data {
		rs := strconv.Itoa(rowIdx)
		wb.SetCellValue(sh, "A"+rs, acc.AcCode)
		wb.SetCellStyle(sh, "A"+rs, "A"+rs, boldL)
		wb.SetCellValue(sh, "B"+rs, acc.AcName)
		wb.SetCellStyle(sh, "B"+rs, "B"+rs, boldL)
		rowIdx++

		rs = strconv.Itoa(rowIdx)
		wb.SetCellValue(sh, "A"+rs, "วันที่")
		wb.SetCellStyle(sh, "A"+rs, "A"+rs, hdrStyle)
		wb.SetCellValue(sh, "B"+rs, "สมุดบัญชี")
		wb.SetCellStyle(sh, "B"+rs, "B"+rs, hdrStyle)
		wb.SetCellValue(sh, "C"+rs, "ลำดับที่")
		wb.SetCellStyle(sh, "C"+rs, "C"+rs, hdrCenter)
		wb.SetCellValue(sh, "D"+rs, "อ้างอิง")
		wb.SetCellStyle(sh, "D"+rs, "D"+rs, hdrStyle)
		wb.SetCellValue(sh, "E"+rs, "คำอธิบาย")
		wb.SetCellStyle(sh, "E"+rs, "E"+rs, hdrStyle)
		wb.SetCellValue(sh, "F"+rs, "เดบิต")
		wb.SetCellStyle(sh, "F"+rs, "F"+rs, hdrRight)
		wb.SetCellValue(sh, "G"+rs, "เครดิต")
		wb.SetCellStyle(sh, "G"+rs, "G"+rs, hdrRight)
		rowIdx++

		for _, line := range acc.Lines {
			rs = strconv.Itoa(rowIdx)
			wb.SetCellValue(sh, "A"+rs, line.Date)
			wb.SetCellValue(sh, "B"+rs, line.Book)
			wb.SetCellValue(sh, "C"+rs, line.Seq)
			wb.SetCellValue(sh, "D"+rs, line.Ref)
			wb.SetCellValue(sh, "E"+rs, line.Desc)

			// ถ้าเป็นยอดยกมา และเป็น 0 ให้พิมพ์ 0.00 ฝั่งเดบิต
			if line.Desc == "ยอดยกมา" && line.Debit == 0 && line.Credit == 0 {
				wb.SetCellValue(sh, "F"+rs, "0.00")
				wb.SetCellStyle(sh, "F"+rs, "F"+rs, txtRightStyle)
			} else {
				wb.SetCellFloat(sh, "F"+rs, roundF(line.Debit), 2, 64)
				wb.SetCellFloat(sh, "G"+rs, roundF(line.Credit), 2, 64)
				wb.SetCellStyle(sh, "F"+rs, "G"+rs, numStyle)
			}

			wb.SetCellStyle(sh, "A"+rs, "B"+rs, txtStyle)
			wb.SetCellStyle(sh, "C"+rs, "C"+rs, txtCenterStyle)
			wb.SetCellStyle(sh, "D"+rs, "E"+rs, txtStyle)
			rowIdx++
		}

		rs = strconv.Itoa(rowIdx)
		wb.SetCellValue(sh, "E"+rs, "รวม")
		wb.SetCellStyle(sh, "A"+rs, "E"+rs, totTxtStyle)
		wb.SetCellFloat(sh, "F"+rs, roundF(acc.GrandTotal), 2, 64)
		wb.SetCellStyle(sh, "F"+rs, "F"+rs, totNumStyle)
		wb.SetCellFloat(sh, "G"+rs, roundF(acc.GrandTotal), 2, 64)
		wb.SetCellStyle(sh, "G"+rs, "G"+rs, totNumStyle)
		rowIdx += 3 // เว้นบรรทัดขึ้นบัญชีใหม่
	}

	pathToOpen, err := safeWriteFile(savePath, func(tmp string) error {
		return wb.SaveAs(tmp)
	})
	return pathToOpen, err
}

// ─────────────────────────────────────────────────────────────────
// UI Dialog
// ─────────────────────────────────────────────────────────────────
func showLedgerReportDialog(w fyne.Window, onGoSetup func()) {
	xlOptions := excelize.Options{Password: "@A123456789a"}
	cfg, err := loadCompanyPeriod(xlOptions)
	if err != nil {
		showErrDialog(w, "โหลดข้อมูลงวดไม่ได้: "+err.Error())
		return
	}

	reportDir := getReportDir(xlOptions)
	if strings.HasSuffix(filepath.ToSlash(reportDir), "/Desktop") || reportDir == filepath.ToSlash(filepath.Dir(currentDBPath)) {
		var warn dialog.Dialog
		btnGo := newEnterButton("ไปตั้งค่า (Enter)", func() {
			warn.Hide()
			if onGoSetup != nil {
				onGoSetup()
			}
		})
		btnCancel2 := newEscButton("ยกเลิก (Esc)", func() { warn.Hide() })
		warn = dialog.NewCustomWithoutButtons("⚠️ แจ้งเตือน", container.NewVBox(
			widget.NewLabel("กรุณาตั้งค่า Report Path ที่ Setup > Company Profile"),
			container.NewCenter(container.NewHBox(btnGo, btnCancel2)),
		), w)
		warn.Show()
		return
	}

	var opts []string
	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	showUpTo := cfg.NowPeriod
	if showUpTo > len(periods) {
		showUpTo = len(periods)
	}
	for _, p := range periods[:showUpTo] {
		opts = append(opts, fmt.Sprintf("Period %d", p.PNo))
	}

	cboFrom := widget.NewSelect(opts, nil)
	cboTo := widget.NewSelect(opts, nil)
	cboFrom.SetSelected(opts[0])
	cboTo.SetSelected(opts[showUpTo-1])

	getPeriodIdx := func(sel string) int {
		for i, o := range opts {
			if o == sel {
				return i + 1
			}
		}
		return 1
	}

	// ─── เตรียมตัวเลือกรหัสบัญชีจาก Ledger_Master ───
	comCode := getComCodeFromExcel(xlOptions)
	var acOpts []string
	acMap := make(map[string]string)

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err == nil {
		rows, _ := f.GetRows("Ledger_Master")
		for i, row := range rows {
			if i == 0 {
				continue // ข้ามบรรทัด Header
			}
			// ใช้ Logic เช็คแบบเดียวกับ ledger_search_ui.go
			if len(row) >= 3 && row[0] == comCode {
				acCode := row[1]
				acName := row[2]
				if acCode != "" {
					displayTxt := fmt.Sprintf("%s | %s", acCode, acName)
					acOpts = append(acOpts, displayTxt)
					acMap[displayTxt] = acCode
				}
			}
		}
		f.Close()
	}

	if len(acOpts) == 0 {
		acOpts = []string{"ไม่มีข้อมูลบัญชี"}
		acMap["ไม่มีข้อมูลบัญชี"] = ""
	} else {
		// เรียง accounting order: 1xx → 2xx → 3xx → 4xx → 5xx → 6xx
		sort.Slice(acOpts, func(i, j int) bool {
			return acMap[acOpts[i]] < acMap[acOpts[j]]
		})
	}

	cboAcFrom := widget.NewSelect(acOpts, nil)
	cboAcTo := widget.NewSelect(acOpts, nil)
	cboAcFrom.SetSelected(acOpts[0])
	cboAcTo.SetSelected(acOpts[len(acOpts)-1])

	radioAcMode := widget.NewRadioGroup([]string{"All Code", "Select Code"}, nil)
	radioAcMode.Horizontal = false
	radioAcMode.OnChanged = func(sel string) {
		if sel == "All Code" {
			cboAcFrom.Disable()
			cboAcTo.Disable()
		} else {
			cboAcFrom.Enable()
			cboAcTo.Enable()
		}
	}
	radioAcMode.SetSelected("All Code")

	var pop *widget.PopUp
	closePopup := func() {
		if pop != nil {
			pop.Hide()
		}
	}

	runExport := func(isPDF bool) {
		pFrom := getPeriodIdx(cboFrom.Selected)
		pTo := getPeriodIdx(cboTo.Selected)
		if pFrom > pTo {
			dialog.ShowError(fmt.Errorf("งวดเริ่มต้น ต้องน้อยกว่าหรืองวดสิ้นสุด"), w)
			return
		}

		glFrom, glTo := "", "ZZZZZZZZZZ"
		if radioAcMode.Selected == "Select Code" {
			glFrom = acMap[cboAcFrom.Selected]
			glTo = acMap[cboAcTo.Selected]
			if glFrom > glTo {
				dialog.ShowError(fmt.Errorf("รหัสบัญชีเริ่มต้น ต้องน้อยกว่ารหัสบัญชีสิ้นสุด"), w)
				return
			}
		}

		ext := ".xlsx"
		if isPDF {
			ext = ".pdf"
		}
		savePath := filepath.Join(reportDir, fmt.Sprintf("Ledger_P%02d_P%02d%s", pFrom, pTo, ext))
		closePopup()

		var pathToOpen string
		var err error
		if isPDF {
			pathToOpen, err = exportLedgerPDF(xlOptions, pFrom, pTo, glFrom, glTo, savePath)
		} else {
			pathToOpen, err = exportLedgerExcel(xlOptions, pFrom, pTo, glFrom, glTo, savePath)
		}

		if err != nil {
			showErrDialog(w, "สร้างรายงานไม่ได้: "+err.Error())
		} else {
			isTmp := strings.HasSuffix(pathToOpen, ".tmp")
			if isTmp {
				showErrDialog(w, "เปิดรายงานชั่วคราว\nปิดไฟล์เดิมก่อน แล้วกด Export ใหม่เพื่อบันทึกถาวร")
			}
			openFile(pathToOpen)
		}
	}

	btnExcel := widget.NewButton("📊 Excel", func() { runExport(false) })
	btnExcel.Importance = widget.HighImportance
	btnPDF := widget.NewButton("📄 PDF", func() { runExport(true) })
	btnPDF.Importance = widget.HighImportance
	btnCancel := widget.NewButton("Cancel", closePopup)

	form := container.NewVBox(
		container.NewHBox(widget.NewLabelWithStyle("Select End Date", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}), cboFrom, widget.NewLabel("To"), cboTo),
		widget.NewSeparator(),
		radioAcMode,
		container.NewHBox(layout.NewSpacer(), cboAcFrom, widget.NewLabel("To"), cboAcTo),
		widget.NewSeparator(),
		container.NewHBox(layout.NewSpacer(), btnExcel, btnPDF, btnCancel, layout.NewSpacer()),
	)

	pop = widget.NewModalPopUp(
		container.NewVBox(
			container.NewHBox(widget.NewLabelWithStyle("Select Period", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}), layout.NewSpacer(), widget.NewButton("X", closePopup)),
			widget.NewSeparator(),
			container.NewPadded(form),
		), w.Canvas(),
	)
	pop.Resize(fyne.NewSize(550, 300))
	pop.Show()
}
