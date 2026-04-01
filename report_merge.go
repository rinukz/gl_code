package main

// report_merge.go  v2 — Master Marker Edition
// ─────────────────────────────────────────────────────────────────
// Consolidated (Merge) Report — รวมงบหลายสาขา/บริษัทในเครือ
//
// KEY DESIGN: Master Marker (MDM — Master Data Management)
//   - ไฟล์ที่ถูก mark เป็น Master → ใช้เป็น "พจนานุกรม" ผังบัญชี
//     (AcName, Gcode, Gname, ลำดับ, GroupLabel)
//   - ไฟล์สาขาอื่น → ดึงแต่ตัวเลข (CBal, BBal)
//   - AC Code ที่สาขามีแต่ Master ไม่มี → รวมยอดแต่แยกไปหมวด "⚠ บัญชีนอกผัง"
//
// Flow:
//   1. LoadBranchInfo()       — อ่านข้อมูลแต่ละไฟล์
//   2. BuildMasterDictionary()— สร้าง map[acCode]AccountMeta จากไฟล์ Master
//   3. AggregateLedgers()     — รวมตัวเลขทุกสาขา, เทียบกับ Master dict
//   4. ExportMergeReport()    — เขียน Excel loop จาก masterOrder (ลำดับ Master)
// ─────────────────────────────────────────────────────────────────

import (
	"fmt"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────
// Structs
// ─────────────────────────────────────────────────────────────────

// BranchInfo — ข้อมูลแต่ละสาขา
type BranchInfo struct {
	FilePath     string
	BranchName   string
	TaxID        string
	ComCode      string
	YearEnd      time.Time
	TotalPeriods int
	NowPeriod    int
	IsMaster     bool // ★ Master Marker — ยึดเป็นพจนานุกรมผังบัญชี
}

// AccountMeta — ข้อมูล Meta ของบัญชี (ดึงจาก Master เท่านั้น)
// ใช้เป็น "source of truth" สำหรับ AcName / Gcode / Gname / ลำดับ
type AccountMeta struct {
	AcName    string
	Gcode     string
	Gname     string
	Group     int    // 1=Asset 2=Liab 3=Equity 4=Rev 5=Exp 6=InvAdj
	SortOrder int    // ลำดับตามผัง Master (ใช้เรียงรายงาน)
}

// ConsolidatedAccount — ยอดรวม 1 AC Code จากทุกสาขา
type ConsolidatedAccount struct {
	AcCode     string
	Meta       AccountMeta        // ดึงจาก Master dict (ถ้าไม่มีใน Master = นอกผัง)
	IsOffChart bool               // true = สาขามีแต่ Master ไม่มี (⚠ Off-chart)
	BranchBal  map[string]float64 // key=BranchName, val=CBal งวดที่เลือก
	BranchBBal map[string]float64 // key=BranchName, val=BBal งวดที่เลือก
	TotalBBal  float64
	TotalCBal  float64
}

// MergeOptions — ตัวเลือก Export
type MergeOptions struct {
	IncludeTrialBalance  bool
	IncludePnL           bool
	IncludeBalanceSheet  bool
	MatrixLayout         bool     // แสดงแยกคอลัมน์ตามสาขา
	EliminateInterBranch bool
	InterBranchCodes     []string
	NowPeriod            int      // 0 = ใช้ NowPeriod ของแต่ละไฟล์
	IncludeOffChart      bool     // รวมบัญชีนอกผัง Master ในรายงานหรือไม่
}

// AggregateResult — ผลลัพธ์จาก AggregateLedgers
type AggregateResult struct {
	Consolidated map[string]*ConsolidatedAccount
	MasterOrder  []string // ลำดับจาก Master (ใช้เรียง output)
	OffChart     []string // AC Code ที่สาขามีแต่ Master ไม่มี
	Warnings     []string // คำเตือนต่างๆ
}

// ─────────────────────────────────────────────────────────────────
// getAcctGroupMerge — แยก Group บัญชีจาก AC Code (prefix แรก)
// ─────────────────────────────────────────────────────────────────
func getAcctGroupMerge(acCode string) int {
	if len(acCode) == 0 {
		return 0
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
	case '6':
		return 6
	}
	return 0
}

// groupLabel — ชื่อหมวดบัญชีภาษาไทย
func groupLabelMerge(grp int) string {
	switch grp {
	case 1:
		return "หมวด 1 — สินทรัพย์ (Assets)"
	case 2:
		return "หมวด 2 — หนี้สิน (Liabilities)"
	case 3:
		return "หมวด 3 — ส่วนของผู้ถือหุ้น (Equity)"
	case 4:
		return "หมวด 4 — รายได้ (Revenue)"
	case 5:
		return "หมวด 5 — ค่าใช้จ่าย / ต้นทุน (Expenses & Cost)"
	case 6:
		return "หมวด 6 — ปรับปรุงสินค้าคงเหลือ (Inventory Adj.)"
	}
	return "ไม่ระบุหมวด"
}

// ─────────────────────────────────────────────────────────────────
// LoadBranchInfo — อ่านข้อมูลสาขาจากไฟล์
// ─────────────────────────────────────────────────────────────────
func LoadBranchInfo(filePath string) (*BranchInfo, error) {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	f, err := excelize.OpenFile(filePath, xlOpts)
	if err != nil {
		return nil, fmt.Errorf("เปิดไฟล์ไม่ได้: %v", err)
	}
	defer f.Close()

	info := &BranchInfo{FilePath: filePath}
	info.ComCode, _ = f.GetCellValue("Company_Profile", "A2")
	info.BranchName, _ = f.GetCellValue("Company_Profile", "B2")
	info.TaxID, _ = f.GetCellValue("Company_Profile", "D2")

	yend, _ := f.GetCellValue("Company_Profile", "E2")
	if yend == "" {
		return nil, fmt.Errorf("ไม่พบ ComYEnd ใน %s", filePath)
	}
	t, err2 := time.Parse("02/01/06", yend)
	if err2 != nil {
		t, err2 = time.Parse("02/01/2006", yend)
		if err2 != nil {
			return nil, fmt.Errorf("รูปแบบ ComYEnd ไม่ถูกต้องใน %s", filePath)
		}
	}
	info.YearEnd = t

	perStr, _ := f.GetCellValue("Company_Profile", "F2")
	nperStr, _ := f.GetCellValue("Company_Profile", "G2")
	fmt.Sscanf(perStr, "%d", &info.TotalPeriods)
	fmt.Sscanf(nperStr, "%d", &info.NowPeriod)

	if info.BranchName == "" {
		info.BranchName = info.ComCode
	}
	if info.BranchName == "" {
		info.BranchName = filePath
	}
	return info, nil
}

// ─────────────────────────────────────────────────────────────────
// ValidateFilesForMerge — ตรวจว่า YearEnd + TotalPeriods ตรงกัน
// ─────────────────────────────────────────────────────────────────
func ValidateFilesForMerge(branches []*BranchInfo) error {
	if len(branches) < 2 {
		return fmt.Errorf("ต้องเลือกอย่างน้อย 2 ไฟล์เพื่อทำ Merge")
	}
	// ตรวจว่ามี Master หนึ่งตัว
	masterCount := 0
	for _, b := range branches {
		if b.IsMaster {
			masterCount++
		}
	}
	if masterCount == 0 {
		return fmt.Errorf("กรุณากำหนด Master (ผังบัญชีหลัก) 1 ไฟล์\nคลิกที่ปุ่ม 👑 ในตารางเพื่อกำหนด")
	}
	if masterCount > 1 {
		return fmt.Errorf("กำหนด Master ได้เพียง 1 ไฟล์เท่านั้น (พบ %d ไฟล์)", masterCount)
	}

	ref := branches[0]
	for _, b := range branches {
		if !b.YearEnd.Equal(ref.YearEnd) {
			return fmt.Errorf(
				"วันสิ้นปีบัญชีไม่ตรงกัน:\n  %s: %s\n  %s: %s",
				ref.BranchName, ref.YearEnd.Format("02/01/2006"),
				b.BranchName, b.YearEnd.Format("02/01/2006"),
			)
		}
		if b.TotalPeriods != ref.TotalPeriods {
			return fmt.Errorf(
				"จำนวน Period ไม่ตรงกัน:\n  %s: %d\n  %s: %d",
				ref.BranchName, ref.TotalPeriods,
				b.BranchName, b.TotalPeriods,
			)
		}
	}
	return nil
}

// ─────────────────────────────────────────────────────────────────
// BuildMasterDictionary — Step 1: สร้างพจนานุกรมผังบัญชีจาก Master
//
// Return:
//   masterDict  map[acCode]AccountMeta  — ข้อมูล Meta ยึดจาก Master
//   masterOrder []string                — ลำดับตามผัง Master (เรียง sort)
// ─────────────────────────────────────────────────────────────────
func BuildMasterDictionary(master *BranchInfo) (map[string]AccountMeta, []string, error) {
	xlOpts := excelize.Options{Password: "@A123456789a"}
	f, err := excelize.OpenFile(master.FilePath, xlOpts)
	if err != nil {
		return nil, nil, fmt.Errorf("เปิดไฟล์ Master ไม่ได้: %v", err)
	}
	defer f.Close()

	masterDict := make(map[string]AccountMeta)
	var masterOrder []string

	rows, _ := f.GetRows("Ledger_Master")
	sortIdx := 0
	for i, row := range rows {
		if i == 0 || len(row) < 5 {
			continue
		}
		if safeGet(row, 0) != master.ComCode {
			continue
		}
		acCode := strings.TrimSpace(safeGet(row, 1))
		if acCode == "" {
			continue
		}
		grp := getAcctGroupMerge(acCode)
		masterDict[acCode] = AccountMeta{
			AcName:    strings.TrimSpace(safeGet(row, 2)),
			Gcode:     strings.TrimSpace(safeGet(row, 3)),
			Gname:     strings.TrimSpace(safeGet(row, 4)),
			Group:     grp,
			SortOrder: sortIdx,
		}
		masterOrder = append(masterOrder, acCode)
		sortIdx++
	}

	// เรียงตาม accounting sort (prefix 3 หลักเป็น int)
	sort.SliceStable(masterOrder, func(i, j int) bool {
		return acCodeLess(masterOrder[i], masterOrder[j])
	})

	return masterDict, masterOrder, nil
}

// ─────────────────────────────────────────────────────────────────
// acCodeLess — comparator สำหรับ accounting sort
// ─────────────────────────────────────────────────────────────────
func acCodeLess(ai, aj string) bool {
	pi, pj := ai, aj
	if len(pi) > 3 {
		pi = pi[:3]
	}
	if len(pj) > 3 {
		pj = pj[:3]
	}
	var ni, nj int
	_, erri := fmt.Sscanf(pi, "%d", &ni)
	_, errj := fmt.Sscanf(pj, "%d", &nj)
	if erri == nil && errj == nil {
		if ni != nj {
			return ni < nj
		}
		return ai < aj
	}
	return ai < aj
}

// ─────────────────────────────────────────────────────────────────
// AggregateLedgers — Step 2+3: รวมตัวเลขทุกสาขา + เทียบ Master dict
//
// Ledger_Master column (0-based slice index):
//   0=Comcode  1=Ac_code  2=Ac_name  3=Gcode  4=Gname
//   5=BBAL  6=CBAL  7=Debit  8=Credit  9=Bthisyear
//   10-21=Thisper01-12  22=Blastyear  23-34=Lastper01-12
//
// Thisper คือ cumulative balance ณ สิ้นงวด N (YTD)
// CBal งวด N = Thisper[N-1]
// BBal งวด N = Thisper[N-2] ถ้า N>1  หรือ Bthisyear ถ้า N=1
// ─────────────────────────────────────────────────────────────────
func AggregateLedgers(branches []*BranchInfo, masterDict map[string]AccountMeta, opts MergeOptions) (*AggregateResult, error) {
	xlOpts := excelize.Options{Password: "@A123456789a"}

	result := &AggregateResult{
		Consolidated: make(map[string]*ConsolidatedAccount),
	}
	offChartSet := make(map[string]bool) // AC Code ที่สาขามีแต่ Master ไม่มี

	for _, branch := range branches {
		f, err := excelize.OpenFile(branch.FilePath, xlOpts)
		if err != nil {
			return nil, fmt.Errorf("เปิดไฟล์ %s ไม่ได้: %v", branch.BranchName, err)
		}

		rows, _ := f.GetRows("Ledger_Master")

		targetPeriod := branch.NowPeriod
		if opts.NowPeriod > 0 && opts.NowPeriod <= branch.TotalPeriods {
			targetPeriod = opts.NowPeriod
		}
		if targetPeriod < 1 {
			targetPeriod = 1
		}

		for i, row := range rows {
			if i == 0 || len(row) < 5 {
				continue
			}
			if safeGet(row, 0) != branch.ComCode {
				continue
			}
			acCode := strings.TrimSpace(safeGet(row, 1))
			if acCode == "" {
				continue
			}

			// Eliminate inter-branch
			if opts.EliminateInterBranch {
				skip := false
				for _, ic := range opts.InterBranchCodes {
					if acCode == ic {
						skip = true
						break
					}
				}
				if skip {
					continue
				}
			}

			// ──── ตรวจว่า AC Code นี้อยู่ใน Master dict ไหม ────────
			meta, inMaster := masterDict[acCode]
			isOffChart := !inMaster

			if isOffChart {
				// สาขามีแต่ Master ไม่มี — บันทึก warning
				if !offChartSet[acCode] {
					offChartSet[acCode] = true
					// สร้าง Meta จากข้อมูลสาขาเอง (fallback)
					acNameFallback := strings.TrimSpace(safeGet(row, 2))
					meta = AccountMeta{
						AcName:    acNameFallback,
						Gcode:     strings.TrimSpace(safeGet(row, 3)),
						Gname:     strings.TrimSpace(safeGet(row, 4)),
						Group:     getAcctGroupMerge(acCode),
						SortOrder: 9999 + len(offChartSet), // ท้ายสุด
					}
					result.Warnings = append(result.Warnings,
						fmt.Sprintf("⚠️  [%s] %s — มีใน %s แต่ไม่มีในผัง Master",
							acCode, acNameFallback, branch.BranchName))
				}
			}

			// ──── คำนวณ BBal / CBal ─────────────────────────────────
			bthis := parseFloat(safeGet(row, 9))
			grp := meta.Group
			if grp >= 4 {
				bthis = 0
			}

			// BBal = ยอดต้นงวด targetPeriod
			//   targetPeriod=1  → BBal = Bthisyear
			//   targetPeriod=N  → BBal = Thisper[N-2]  (cumulative ณ สิ้น period N-1)
			bbal := bthis
			if targetPeriod > 1 {
				prevIdx := 10 + targetPeriod - 2 // Thisper[N-2], 0-based slice
				if prevIdx < len(row) {
					if grp < 4 {
						bbal = parseFloat(safeGet(row, prevIdx))
					}
					// P&L: BBal = 0 (แสดง YTD ตั้งแต่ต้นปี)
				}
			}

			// CBal = Thisper[targetPeriod-1]  (cumulative ณ สิ้น period N)
			cbal := float64(0)
			cbalIdx := 10 + targetPeriod - 1
			if cbalIdx < len(row) {
				cbal = parseFloat(safeGet(row, cbalIdx))
			}

			// ──── สะสมใน map ────────────────────────────────────────
			if _, exists := result.Consolidated[acCode]; !exists {
				result.Consolidated[acCode] = &ConsolidatedAccount{
					AcCode:     acCode,
					Meta:       meta,
					IsOffChart: isOffChart,
					BranchBal:  make(map[string]float64),
					BranchBBal: make(map[string]float64),
				}
			}
			ca := result.Consolidated[acCode]
			// ถ้าเป็น Master → ชื่อ/หมวดจาก Master เสมอ (อย่าให้สาขาทับ)
			if branch.IsMaster {
				ca.Meta = meta
			}
			ca.BranchBal[branch.BranchName] += cbal
			ca.BranchBBal[branch.BranchName] += bbal
			ca.TotalBBal += bbal
			ca.TotalCBal += cbal
		}
		f.Close()
	}

	// ──── สร้าง offChart slice ───────────────────────────────────
	for acCode := range offChartSet {
		result.OffChart = append(result.OffChart, acCode)
	}
	sort.Slice(result.OffChart, func(i, j int) bool {
		return acCodeLess(result.OffChart[i], result.OffChart[j])
	})

	return result, nil
}

// ─────────────────────────────────────────────────────────────────
// ExportMergeReport — Step 4: Export Excel
// Loop จาก masterOrder (ลำดับ Master) เพื่อรักษา layout ของสำนักงานใหญ่
// ─────────────────────────────────────────────────────────────────
func ExportMergeReport(
	branches []*BranchInfo,
	masterDict map[string]AccountMeta,
	masterOrder []string,
	aggResult *AggregateResult,
	opts MergeOptions,
	outputPath string,
) error {
	out := excelize.NewFile()
	out.DeleteSheet("Sheet1")

	var branchNames []string
	for _, b := range branches {
		branchNames = append(branchNames, b.BranchName)
	}

	// ──── Styles ─────────────────────────────────────────────────
	titleStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 13, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"1F4E79"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	subtitleStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 10, Italic: true, Color: "404040"},
		Alignment: &excelize.Alignment{Horizontal: "left"},
	})
	headerStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10, Color: "1F4E79"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"DEEAF1"}, Pattern: 1},
		Border:    []excelize.Border{{Type: "bottom", Color: "1F4E79", Style: 2}},
		Alignment: &excelize.Alignment{Horizontal: "center", WrapText: true, Vertical: "center"},
	})
	numStyle, _ := out.NewStyle(&excelize.Style{
		NumFmt:    4,
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})
	negStyle, _ := out.NewStyle(&excelize.Style{
		NumFmt:    4,
		Font:      &excelize.Font{Color: "C00000"},
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})
	groupStyle, _ := out.NewStyle(&excelize.Style{
		Font:  &excelize.Font{Bold: true, Size: 10, Color: "1F4E79"},
		Fill:  excelize.Fill{Type: "pattern", Color: []string{"D9E1F2"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "1F4E79", Style: 1},
			{Type: "bottom", Color: "1F4E79", Style: 1},
		},
	})
	subtotalStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E2EFDA"}, Pattern: 1},
		NumFmt:    4,
		Alignment: &excelize.Alignment{Horizontal: "right"},
		Border:    []excelize.Border{{Type: "top", Color: "70AD47", Style: 1}},
	})
	totalStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 11, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"375623"}, Pattern: 1},
		NumFmt:    4,
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})
	offChartStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "C00000", Italic: true},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFE7E7"}, Pattern: 1},
		NumFmt:    4,
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})
	masterTagStyle, _ := out.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "7B3F00"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFF2CC"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	_ = masterTagStyle

	// ──── Period info ─────────────────────────────────────────────
	var masterBranch *BranchInfo
	for _, b := range branches {
		if b.IsMaster {
			masterBranch = b
			break
		}
	}
	targetPeriod := masterBranch.NowPeriod
	if opts.NowPeriod > 0 {
		targetPeriod = opts.NowPeriod
	}
	periodInfo := fmt.Sprintf(
		"Period %d/%d  |  YearEnd: %s  |  Master: %s  |  จำนวนสาขา: %d",
		targetPeriod, masterBranch.TotalPeriods,
		masterBranch.YearEnd.Format("02/01/2006"),
		masterBranch.BranchName,
		len(branches),
	)

	// ──── Sheet: Trial Balance ────────────────────────────────────
	if opts.IncludeTrialBalance {
		sh := "Trial Balance รวม"
		out.NewSheet(sh)
		styles := mergeStyles{titleStyle, subtitleStyle, headerStyle, numStyle, negStyle, groupStyle, subtotalStyle, totalStyle, offChartStyle}
		writeMergeTB(out, sh, aggResult, masterOrder, branchNames, opts, periodInfo, styles)
	}

	// ──── Sheet: P&L ─────────────────────────────────────────────
	if opts.IncludePnL {
		sh := "P&L รวม"
		out.NewSheet(sh)
		styles := mergeStyles{titleStyle, subtitleStyle, headerStyle, numStyle, negStyle, groupStyle, subtotalStyle, totalStyle, offChartStyle}
		writeMergePnL(out, sh, aggResult, masterOrder, branchNames, opts, periodInfo, styles)
	}

	// ──── Sheet: Balance Sheet ────────────────────────────────────
	if opts.IncludeBalanceSheet {
		sh := "Balance Sheet รวม"
		out.NewSheet(sh)
		styles := mergeStyles{titleStyle, subtitleStyle, headerStyle, numStyle, negStyle, groupStyle, subtotalStyle, totalStyle, offChartStyle}
		writeMergeBS(out, sh, aggResult, masterOrder, branchNames, opts, periodInfo, styles)
	}

	// ──── Sheet: ⚠ Off-Chart (ถ้ามี) ─────────────────────────────
	if len(aggResult.OffChart) > 0 {
		sh := "⚠ บัญชีนอกผัง"
		out.NewSheet(sh)
		styles := mergeStyles{titleStyle, subtitleStyle, headerStyle, numStyle, negStyle, groupStyle, subtotalStyle, totalStyle, offChartStyle}
		writeOffChartSheet(out, sh, aggResult, branchNames, periodInfo, styles)
	}

	// ──── Sheet: สรุปสาขา ────────────────────────────────────────
	{
		sh := "สรุปสาขา"
		out.NewSheet(sh)
		writeBranchSummary(out, sh, branches, aggResult.Warnings, periodInfo,
			titleStyle, headerStyle, subtitleStyle)
	}

	return out.SaveAs(outputPath, excelize.Options{})
}

// ─────────────────────────────────────────────────────────────────
// mergeStyles — bundle styles เพื่อส่งต่อ helper functions
// ─────────────────────────────────────────────────────────────────
type mergeStyles struct {
	title, subtitle, header, num, neg, group, subtotal, total, offChart int
}

// ─────────────────────────────────────────────────────────────────
// writeMergeTB — Trial Balance Sheet
// Loop จาก masterOrder เพื่อรักษา layout ของ Master
// ─────────────────────────────────────────────────────────────────
func writeMergeTB(
	f *excelize.File, sh string,
	aggResult *AggregateResult,
	masterOrder []string,
	branchNames []string,
	opts MergeOptions,
	periodInfo string,
	s mergeStyles,
) {
	consolidated := aggResult.Consolidated
	// fixed: Code(1), Name(2), Gcode(3) + matrix + TotalBBal + TotalCBal
	fixedCols := 3
	matrixCols := 0
	if opts.MatrixLayout {
		matrixCols = len(branchNames) * 2
	}
	lastCol := fixedCols + matrixCols + 2

	row := 1
	// Title
	f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
	f.SetCellValue(sh, cellRef(1, row), "งบทดลองรวม (Consolidated Trial Balance)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.title)
	f.SetRowHeight(sh, row, 24)
	row++
	f.SetCellValue(sh, cellRef(1, row), periodInfo)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtitle)
	row += 2

	// Column headers
	c := 1
	setH := func(v string) { f.SetCellValue(sh, cellRef(c, row), v); c++ }
	setH("AC Code")
	setH("ชื่อบัญชี (ตามผัง Master)")
	setH("หมวด")
	if opts.MatrixLayout {
		for _, bn := range branchNames {
			setH(bn + "\n(BBal)")
			setH(bn + "\n(CBal)")
		}
	}
	setH("รวม BBal")
	setH("รวม CBal (งวดนี้)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.header)
	f.SetRowHeight(sh, row, 30)
	row++

	currentGroup := 0
	groupBBal := map[int]float64{}
	groupCBal := map[int]float64{}
	var grandBBal, grandCBal float64

	flushGroupRow := func(grp int) {
		if grp == 0 {
			return
		}
		lbl := "  รวม" + groupLabelMerge(grp)
		f.SetCellValue(sh, cellRef(1, row), lbl)
		tBBal := fixedCols + matrixCols + 1
		tCBal := fixedCols + matrixCols + 2
		f.SetCellValue(sh, cellRef(tBBal, row), groupBBal[grp])
		f.SetCellValue(sh, cellRef(tCBal, row), groupCBal[grp])
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtotal)
		row++
		row++ // blank separator
	}

	// ─── Loop ตาม masterOrder (ลำดับจาก Master) ───
	for _, acCode := range masterOrder {
		ca, exists := consolidated[acCode]
		if !exists {
			continue // Master มีแต่ไม่มีข้อมูลใดๆ → ข้าม (หรือจะแสดง 0 ก็ได้)
		}
		grp := ca.Meta.Group

		if grp != currentGroup && grp > 0 {
			flushGroupRow(currentGroup)
			currentGroup = grp
			f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
			f.SetCellValue(sh, cellRef(1, row), groupLabelMerge(grp))
			f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.group)
			row++
		}

		c2 := 1
		f.SetCellValue(sh, cellRef(c2, row), ca.AcCode); c2++
		// ชื่อจาก Master dict เสมอ (ป้องกันสาขา override)
		f.SetCellValue(sh, cellRef(c2, row), ca.Meta.AcName); c2++
		f.SetCellValue(sh, cellRef(c2, row), ca.Meta.Gcode); c2++

		if opts.MatrixLayout {
			for _, bn := range branchNames {
				bb := ca.BranchBBal[bn]
				cb := ca.BranchBal[bn]
				f.SetCellValue(sh, cellRef(c2, row), bb)
				ns := s.num; if bb < 0 { ns = s.neg }
				f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns); c2++
				f.SetCellValue(sh, cellRef(c2, row), cb)
				ns = s.num; if cb < 0 { ns = s.neg }
				f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns); c2++
			}
		}
		f.SetCellValue(sh, cellRef(c2, row), ca.TotalBBal)
		ns := s.num; if ca.TotalBBal < 0 { ns = s.neg }
		f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns); c2++
		f.SetCellValue(sh, cellRef(c2, row), ca.TotalCBal)
		ns = s.num; if ca.TotalCBal < 0 { ns = s.neg }
		f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns)

		groupBBal[grp] += ca.TotalBBal
		groupCBal[grp] += ca.TotalCBal
		grandBBal += ca.TotalBBal
		grandCBal += ca.TotalCBal
		row++
	}
	flushGroupRow(currentGroup)

	// ─── Off-Chart items (ท้ายสุด) ───────────────────────────────
	if opts.IncludeOffChart && len(aggResult.OffChart) > 0 {
		row++
		f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
		f.SetCellValue(sh, cellRef(1, row), "⚠  บัญชีที่ไม่มีในผัง Master (สาขากำหนดเอง)")
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.offChart)
		row++
		var offBBal, offCBal float64
		for _, acCode := range aggResult.OffChart {
			ca := consolidated[acCode]
			if ca == nil {
				continue
			}
			c2 := 1
			f.SetCellValue(sh, cellRef(c2, row), ca.AcCode); c2++
			f.SetCellValue(sh, cellRef(c2, row), "⚠ "+ca.Meta.AcName); c2++
			f.SetCellValue(sh, cellRef(c2, row), ca.Meta.Gcode); c2++
			if opts.MatrixLayout {
				for _, bn := range branchNames {
					f.SetCellValue(sh, cellRef(c2, row), ca.BranchBBal[bn])
					f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart); c2++
					f.SetCellValue(sh, cellRef(c2, row), ca.BranchBal[bn])
					f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart); c2++
				}
			}
			f.SetCellValue(sh, cellRef(c2, row), ca.TotalBBal)
			f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart); c2++
			f.SetCellValue(sh, cellRef(c2, row), ca.TotalCBal)
			f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart)
			offBBal += ca.TotalBBal
			offCBal += ca.TotalCBal
			grandBBal += ca.TotalBBal
			grandCBal += ca.TotalCBal
			row++
		}
		tBBal := fixedCols + matrixCols + 1
		tCBal := fixedCols + matrixCols + 2
		f.SetCellValue(sh, cellRef(1, row), "  รวมบัญชีนอกผัง")
		f.SetCellValue(sh, cellRef(tBBal, row), offBBal)
		f.SetCellValue(sh, cellRef(tCBal, row), offCBal)
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtotal)
		row++
	}

	// ─── Grand Total ──────────────────────────────────────────────
	row++
	tBBal := fixedCols + matrixCols + 1
	tCBal := fixedCols + matrixCols + 2
	f.MergeCell(sh, cellRef(1, row), cellRef(tBBal-1, row))
	f.SetCellValue(sh, cellRef(1, row), "★  ยอดรวมทั้งสิ้น")
	f.SetCellValue(sh, cellRef(tBBal, row), grandBBal)
	f.SetCellValue(sh, cellRef(tCBal, row), grandCBal)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.total)

	// ─── Column widths ────────────────────────────────────────────
	f.SetColWidth(sh, "A", "A", 12)
	f.SetColWidth(sh, "B", "B", 40)
	f.SetColWidth(sh, "C", "C", 8)
	startDataCol := 4
	if opts.MatrixLayout {
		for i := 0; i < len(branchNames)*2+2; i++ {
			cn, _ := excelize.ColumnNumberToName(startDataCol + i)
			f.SetColWidth(sh, cn, cn, 16)
		}
	} else {
		f.SetColWidth(sh, "D", "E", 18)
	}
}

// ─────────────────────────────────────────────────────────────────
// writeMergePnL — P&L Sheet (Group 4 Revenue + Group 5 Expenses)
// ─────────────────────────────────────────────────────────────────
func writeMergePnL(
	f *excelize.File, sh string,
	aggResult *AggregateResult,
	masterOrder []string,
	branchNames []string,
	opts MergeOptions,
	periodInfo string,
	s mergeStyles,
) {
	consolidated := aggResult.Consolidated
	matrixCols := 0
	if opts.MatrixLayout {
		matrixCols = len(branchNames)
	}
	totalCol := 2 + matrixCols + 1
	lastCol := totalCol

	row := 1
	f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
	f.SetCellValue(sh, cellRef(1, row), "งบกำไรขาดทุนรวม (Consolidated P&L Statement)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.title)
	f.SetRowHeight(sh, row, 24)
	row++
	f.SetCellValue(sh, cellRef(1, row), periodInfo)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtitle)
	row += 2

	c := 1
	f.SetCellValue(sh, cellRef(c, row), "AC Code"); c++
	f.SetCellValue(sh, cellRef(c, row), "ชื่อบัญชี (ตามผัง Master)"); c++
	if opts.MatrixLayout {
		for _, bn := range branchNames {
			f.SetCellValue(sh, cellRef(c, row), bn); c++
		}
	}
	f.SetCellValue(sh, cellRef(c, row), "รวม (Consolidated)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.header)
	f.SetRowHeight(sh, row, 26)
	row++

	writeSec := func(grp int, title string, negateForDisplay bool) float64 {
		f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
		f.SetCellValue(sh, cellRef(1, row), title)
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.group)
		row++

		var secTotal float64
		for _, acCode := range masterOrder {
			ca, exists := consolidated[acCode]
			if !exists || ca.Meta.Group != grp {
				continue
			}
			c2 := 1
			f.SetCellValue(sh, cellRef(c2, row), ca.AcCode); c2++
			f.SetCellValue(sh, cellRef(c2, row), ca.Meta.AcName); c2++
			if opts.MatrixLayout {
				for _, bn := range branchNames {
					v := ca.BranchBal[bn]
					if negateForDisplay {
						v = -v
					}
					f.SetCellValue(sh, cellRef(c2, row), v)
					ns := s.num; if v < 0 { ns = s.neg }
					f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns); c2++
				}
			}
			val := ca.TotalCBal
			if negateForDisplay {
				val = -val
			}
			f.SetCellValue(sh, cellRef(c2, row), val)
			ns := s.num; if val < 0 { ns = s.neg }
			f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns)
			secTotal += val
			row++
		}
		f.SetCellValue(sh, cellRef(1, row), "  รวม"+title)
		f.SetCellValue(sh, cellRef(totalCol, row), secTotal)
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtotal)
		row += 2
		return secTotal
	}

	// Revenue: CBal เป็นลบ (Credit nature) → negate เพื่อแสดงเป็นบวก
	totalRev := writeSec(4, "รายได้ (Revenue)", true)
	totalExp := writeSec(5, "ค่าใช้จ่าย / ต้นทุน (Expenses & Cost)", false)

	netPL := totalRev - totalExp
	lbl := "★  กำไรสุทธิ (Net Profit)"
	if netPL < 0 {
		lbl = "★  ขาดทุนสุทธิ (Net Loss)"
	}
	f.MergeCell(sh, cellRef(1, row), cellRef(totalCol-1, row))
	f.SetCellValue(sh, cellRef(1, row), lbl)
	f.SetCellValue(sh, cellRef(totalCol, row), netPL)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.total)

	f.SetColWidth(sh, "A", "A", 12)
	f.SetColWidth(sh, "B", "B", 42)
	if opts.MatrixLayout {
		for i := 0; i <= matrixCols; i++ {
			cn, _ := excelize.ColumnNumberToName(3 + i)
			f.SetColWidth(sh, cn, cn, 18)
		}
	} else {
		f.SetColWidth(sh, "C", "C", 22)
	}
}

// ─────────────────────────────────────────────────────────────────
// writeMergeBS — Balance Sheet (Group 1 Assets + 2 Liab + 3 Equity)
// ─────────────────────────────────────────────────────────────────
func writeMergeBS(
	f *excelize.File, sh string,
	aggResult *AggregateResult,
	masterOrder []string,
	branchNames []string,
	opts MergeOptions,
	periodInfo string,
	s mergeStyles,
) {
	consolidated := aggResult.Consolidated
	matrixCols := 0
	if opts.MatrixLayout {
		matrixCols = len(branchNames)
	}
	totalCol := 2 + matrixCols + 1
	lastCol := totalCol

	row := 1
	f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
	f.SetCellValue(sh, cellRef(1, row), "งบแสดงฐานะการเงินรวม (Consolidated Balance Sheet)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.title)
	f.SetRowHeight(sh, row, 24)
	row++
	f.SetCellValue(sh, cellRef(1, row), periodInfo)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtitle)
	row += 2

	c := 1
	f.SetCellValue(sh, cellRef(c, row), "AC Code"); c++
	f.SetCellValue(sh, cellRef(c, row), "ชื่อบัญชี (ตามผัง Master)"); c++
	if opts.MatrixLayout {
		for _, bn := range branchNames {
			f.SetCellValue(sh, cellRef(c, row), bn); c++
		}
	}
	f.SetCellValue(sh, cellRef(c, row), "รวม (Consolidated)")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.header)
	f.SetRowHeight(sh, row, 26)
	row++

	writeSec := func(grp int, title string) float64 {
		f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
		f.SetCellValue(sh, cellRef(1, row), title)
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.group)
		row++

		var secTotal float64
		for _, acCode := range masterOrder {
			ca, exists := consolidated[acCode]
			if !exists || ca.Meta.Group != grp {
				continue
			}
			c2 := 1
			f.SetCellValue(sh, cellRef(c2, row), ca.AcCode); c2++
			f.SetCellValue(sh, cellRef(c2, row), ca.Meta.AcName); c2++
			if opts.MatrixLayout {
				for _, bn := range branchNames {
					v := ca.BranchBal[bn]
					f.SetCellValue(sh, cellRef(c2, row), v)
					ns := s.num; if v < 0 { ns = s.neg }
					f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns); c2++
				}
			}
			f.SetCellValue(sh, cellRef(c2, row), ca.TotalCBal)
			ns := s.num; if ca.TotalCBal < 0 { ns = s.neg }
			f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), ns)
			secTotal += ca.TotalCBal
			row++
		}
		f.SetCellValue(sh, cellRef(1, row), "  รวม"+title)
		f.SetCellValue(sh, cellRef(totalCol, row), secTotal)
		f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtotal)
		row += 2
		return secTotal
	}

	totalAssets := writeSec(1, "สินทรัพย์ (Assets)")
	totalLiab := writeSec(2, "หนี้สิน (Liabilities)")
	totalEquity := writeSec(3, "ส่วนของผู้ถือหุ้น (Equity)")

	le := totalLiab + totalEquity
	f.MergeCell(sh, cellRef(1, row), cellRef(totalCol-1, row))
	f.SetCellValue(sh, cellRef(1, row), "★  รวมสินทรัพย์ (Total Assets)")
	f.SetCellValue(sh, cellRef(totalCol, row), totalAssets)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.total)
	row++
	f.MergeCell(sh, cellRef(1, row), cellRef(totalCol-1, row))
	f.SetCellValue(sh, cellRef(1, row), "★  รวมหนี้สิน + ทุน (Total L&E)")
	f.SetCellValue(sh, cellRef(totalCol, row), le)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.total)
	row += 2

	diff := totalAssets - le
	if diff > 0.005 || diff < -0.005 {
		f.SetCellValue(sh, cellRef(1, row), fmt.Sprintf("⚠️  ไม่ Balance: ส่วนต่าง = %.2f  (ตรวจสอบรายการนอกผัง Master)", diff))
	} else {
		f.SetCellValue(sh, cellRef(1, row), "✅  งบ Balance เรียบร้อย")
	}

	f.SetColWidth(sh, "A", "A", 12)
	f.SetColWidth(sh, "B", "B", 42)
	if opts.MatrixLayout {
		for i := 0; i <= matrixCols; i++ {
			cn, _ := excelize.ColumnNumberToName(3 + i)
			f.SetColWidth(sh, cn, cn, 18)
		}
	} else {
		f.SetColWidth(sh, "C", "C", 22)
	}
}

// ─────────────────────────────────────────────────────────────────
// writeOffChartSheet — Sheet ⚠ บัญชีนอกผัง
// ─────────────────────────────────────────────────────────────────
func writeOffChartSheet(
	f *excelize.File, sh string,
	aggResult *AggregateResult,
	branchNames []string,
	periodInfo string,
	s mergeStyles,
) {
	lastCol := 2 + len(branchNames) + 1
	row := 1
	f.MergeCell(sh, cellRef(1, row), cellRef(lastCol, row))
	f.SetCellValue(sh, cellRef(1, row), "⚠  บัญชีที่สาขาสร้างเองและไม่มีในผัง Master")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.title)
	f.SetRowHeight(sh, row, 24)
	row++
	f.SetCellValue(sh, cellRef(1, row), periodInfo)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.subtitle)
	row += 2

	c := 1
	f.SetCellValue(sh, cellRef(c, row), "AC Code"); c++
	f.SetCellValue(sh, cellRef(c, row), "ชื่อบัญชี (ของสาขา)"); c++
	for _, bn := range branchNames {
		f.SetCellValue(sh, cellRef(c, row), bn); c++
	}
	f.SetCellValue(sh, cellRef(c, row), "รวม")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(lastCol, row), s.header)
	row++

	for _, acCode := range aggResult.OffChart {
		ca := aggResult.Consolidated[acCode]
		if ca == nil {
			continue
		}
		c2 := 1
		f.SetCellValue(sh, cellRef(c2, row), ca.AcCode); c2++
		f.SetCellValue(sh, cellRef(c2, row), ca.Meta.AcName); c2++
		for _, bn := range branchNames {
			v := ca.BranchBal[bn]
			f.SetCellValue(sh, cellRef(c2, row), v)
			f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart); c2++
		}
		f.SetCellValue(sh, cellRef(c2, row), ca.TotalCBal)
		f.SetCellStyle(sh, cellRef(c2, row), cellRef(c2, row), s.offChart)
		row++
	}

	// Warnings
	if len(aggResult.Warnings) > 0 {
		row += 2
		f.SetCellValue(sh, cellRef(1, row), "คำเตือนทั้งหมด:")
		row++
		for _, w := range aggResult.Warnings {
			f.SetCellValue(sh, cellRef(1, row), w)
			row++
		}
	}

	f.SetColWidth(sh, "A", "A", 12)
	f.SetColWidth(sh, "B", "B", 40)
	for i := 0; i < len(branchNames)+1; i++ {
		cn, _ := excelize.ColumnNumberToName(3 + i)
		f.SetColWidth(sh, cn, cn, 18)
	}
}

// ─────────────────────────────────────────────────────────────────
// writeBranchSummary — Sheet สรุปสาขา + Warnings
// ─────────────────────────────────────────────────────────────────
func writeBranchSummary(
	f *excelize.File, sh string,
	branches []*BranchInfo,
	warnings []string,
	periodInfo string,
	titleStyle, headerStyle, subtitleStyle int,
) {
	row := 1
	f.MergeCell(sh, cellRef(1, row), cellRef(9, row))
	f.SetCellValue(sh, cellRef(1, row), "สรุปข้อมูลสาขาที่รวมงบ")
	f.SetCellStyle(sh, cellRef(1, row), cellRef(9, row), titleStyle)
	f.SetRowHeight(sh, row, 22)
	row++
	f.SetCellValue(sh, cellRef(1, row), periodInfo)
	f.SetCellStyle(sh, cellRef(1, row), cellRef(9, row), subtitleStyle)
	row++
	f.SetCellValue(sh, cellRef(1, row),
		"ออกรายงาน: "+time.Now().Format("02/01/2006 15:04:05"))
	f.SetCellStyle(sh, cellRef(1, row), cellRef(9, row), subtitleStyle)
	row += 2

	headers := []string{"#", "ชื่อบริษัท/สาขา", "Tax ID", "ComCode", "YearEnd", "Periods", "NowPeriod", "Master", "ไฟล์"}
	for i, h := range headers {
		f.SetCellValue(sh, cellRef(i+1, row), h)
	}
	f.SetCellStyle(sh, cellRef(1, row), cellRef(len(headers), row), headerStyle)
	row++

	for i, b := range branches {
		masterTag := ""
		if b.IsMaster {
			masterTag = "👑 Master"
		}
		f.SetCellValue(sh, cellRef(1, row), i+1)
		f.SetCellValue(sh, cellRef(2, row), b.BranchName)
		f.SetCellValue(sh, cellRef(3, row), b.TaxID)
		f.SetCellValue(sh, cellRef(4, row), b.ComCode)
		f.SetCellValue(sh, cellRef(5, row), b.YearEnd.Format("02/01/2006"))
		f.SetCellValue(sh, cellRef(6, row), b.TotalPeriods)
		f.SetCellValue(sh, cellRef(7, row), b.NowPeriod)
		f.SetCellValue(sh, cellRef(8, row), masterTag)
		f.SetCellValue(sh, cellRef(9, row), b.FilePath)
		row++
	}

	// Warnings section
	if len(warnings) > 0 {
		row += 2
		f.SetCellValue(sh, cellRef(1, row), fmt.Sprintf("⚠️  พบบัญชีนอกผัง Master จำนวน %d รายการ:", len(warnings)))
		row++
		for _, w := range warnings {
			f.SetCellValue(sh, cellRef(1, row), w)
			row++
		}
	} else {
		row += 2
		f.SetCellValue(sh, cellRef(1, row), "✅  ผังบัญชีทุกสาขาตรงกับ Master ทั้งหมด")
	}

	f.SetColWidth(sh, "A", "A", 5)
	f.SetColWidth(sh, "B", "B", 32)
	f.SetColWidth(sh, "C", "C", 18)
	f.SetColWidth(sh, "D", "D", 10)
	f.SetColWidth(sh, "E", "E", 12)
	f.SetColWidth(sh, "F", "F", 9)
	f.SetColWidth(sh, "G", "G", 11)
	f.SetColWidth(sh, "H", "H", 14)
	f.SetColWidth(sh, "I", "I", 50)
}

// ─────────────────────────────────────────────────────────────────
// cellRef — สร้าง Excel cell reference  col(1-based), row(1-based)
// ─────────────────────────────────────────────────────────────────
func cellRef(col, row int) string {
	colName, _ := excelize.ColumnNumberToName(col)
	return fmt.Sprintf("%s%d", colName, row)
}

// ─────────────────────────────────────────────────────────────────
// loadDBNamesForMerge — รายชื่อไฟล์ใน ./dont_edit/
// ─────────────────────────────────────────────────────────────────
func loadDBNamesForMerge() []string {
	return loadDBNames()
}
