// ledger_search_ui.go
package main

import (
	"sort"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────
// loadAllLedgerRecords — โหลด Ledger_Master ทั้งหมด พร้อม
// batch-compute BBal และ CBal จาก Book_items ใน pass เดียว
// (เปิด Excel ครั้งเดียว ไม่ช้า)
// ─────────────────────────────────────────────────────────────
func loadAllLedgerRecords(xlOptions excelize.Options) []LedgerRecord {
	comCode := getComCodeFromExcel(xlOptions)
	cfg, cfgErr := loadCompanyPeriod(xlOptions)

	f, err := excelize.OpenFile(currentDBPath, xlOptions)
	if err != nil {
		return nil
	}
	defer f.Close()

	// 1. อ่าน Ledger_Master
	var records []LedgerRecord
	lRows, _ := f.GetRows("Ledger_Master")
	for i, row := range lRows {
		if i == 0 {
			continue
		}
		if len(row) >= 3 && row[0] == comCode {
			r := emptyLedger()
			r.Comcode = row[0]
			r.AcCode = row[1]
			r.AcName = row[2]
			if len(row) > 3 {
				r.Gcode = row[3]
			}
			if len(row) > 4 {
				r.Gname = row[4]
			}
			if len(row) > 9 {
				r.Bthisyear = row[9] // สำหรับ chain BBal
			}
			if len(row) > 34 {
				r.Lastper[11] = ledgerNum(parseFloat(row[34]))
			}
			records = append(records, r)
		}
	}

	// 2. ถ้าอ่าน period config ไม่ได้ → คืน records โดยไม่คำนวณ
	if cfgErr != nil || len(records) == 0 {
		// sort ก่อน return
		sortLedgerRecords(records)
		return records
	}

	periods := calcAllPeriods(cfg.YearEnd, cfg.TotalPeriods)
	if cfg.NowPeriod < 1 || cfg.NowPeriod > len(periods) {
		sortLedgerRecords(records)
		return records
	}

	// 3. อ่าน Book_items ครั้งเดียว สะสมยอดทุก account ทุก period
	type acSum struct{ dr, cr [12]float64 }
	sumMap := make(map[string]*acSum)

	bRows, _ := f.GetRows("Book_items")
	for i, row := range bRows {
		if i == 0 || len(row) < 11 || safeGet(row, 0) != comCode {
			continue
		}
		acCode := safeGet(row, 5)
		if acCode == "" {
			continue
		}
		t, err := parseSubbookDate(safeGet(row, 1))
		if err != nil {
			continue
		}
		dr := parseFloat(safeGet(row, 9))
		cr := parseFloat(safeGet(row, 10))
		for pIdx, p := range periods {
			if !t.Before(p.PStart) && !t.After(p.PEnd) {
				if sumMap[acCode] == nil {
					sumMap[acCode] = &acSum{}
				}
				sumMap[acCode].dr[pIdx] += dr
				sumMap[acCode].cr[pIdx] += cr
				break
			}
		}
	}

	// 4. คำนวณ BBal / CBal สำหรับแต่ละ account (chain เหมือน computeRealTimeLedger)
	nowIdx := cfg.NowPeriod - 1 // 0-based index
	for idx := range records {
		r := &records[idx]
		bthis := parseFloat(r.Bthisyear)
		if r.AcCode >= "4" {
			bthis = 0 // P&L ไม่มียอดยกมา
		}
		// chain: bbal = bthis + net ของ per1..per(now-1)
		bbal := bthis
		s := sumMap[r.AcCode]
		for p := 0; p < nowIdx; p++ {
			if s != nil {
				bbal += s.dr[p] - s.cr[p]
			}
		}
		// cbal = bbal + net ของงวดปัจจุบัน
		cbal := bbal
		if s != nil {
			cbal += s.dr[nowIdx] - s.cr[nowIdx]
		}
		r.BBal = ledgerNum(bbal)
		r.CBal = ledgerNum(cbal)
	}

	// 5. เรียงลำดับ records ตาม AcCode แบบ accounting sort
	sortLedgerRecords(records)

	return records
}

// ─────────────────────────────────────────────────────────────
// sortLedgerRecords — เรียง []LedgerRecord ตาม AcCode
//
// Accounting Sort:
//   - เปรียบ prefix 3 หลักแรกเป็น int ก่อน  (100 < 111 < 120 < 235 < 360)
//   - prefix เท่ากัน → ใช้ lexicographic     (120 < 120VAT < 120WHT)
//
// ผลลัพธ์: 100 → 111 → 120 → 120VAT → 120WHT → 235 → 235TVAT → 235VAT → 235WHT → ...
// ─────────────────────────────────────────────────────────────
func sortLedgerRecords(records []LedgerRecord) {
	sort.Slice(records, func(i, j int) bool {
		ai := records[i].AcCode
		aj := records[j].AcCode

		// ดึง prefix 3 หลักแรก
		pi := ai
		if len(pi) > 3 {
			pi = pi[:3]
		}
		pj := aj
		if len(pj) > 3 {
			pj = pj[:3]
		}

		ni, erri := strconv.Atoi(pi)
		nj, errj := strconv.Atoi(pj)

		// ทั้งคู่เป็นตัวเลข → เปรียบตัวเลขก่อน
		if erri == nil && errj == nil {
			if ni != nj {
				return ni < nj
			}
			// prefix เท่ากัน (เช่น "120" vs "120VAT") → lexicographic
			return ai < aj
		}

		// fallback: lexicographic
		return ai < aj
	})
}

func showLedgerSearch(w fyne.Window, xlOptions excelize.Options, allCodes []string,
	onSelect func(r LedgerRecord, idx int)) {

	allRecords := loadAllLedgerRecords(xlOptions)
	filteredRecords := make([]LedgerRecord, len(allRecords))
	copy(filteredRecords, allRecords)

	selectedIdx := 0

	var list *widget.List
	var pop *widget.PopUp

	// ── doSelect — ยืนยันการเลือก ─────────────────────────────────────
	doSelect := func(id int) {
		if id < 0 || id >= len(filteredRecords) {
			return
		}
		r := filteredRecords[id]
		pop.Hide()
		for i, c := range allCodes {
			if c == r.AcCode {
				onSelect(r, i)
				return
			}
		}
		onSelect(r, -1)
	}

	// ── highlightOnly — เลื่อน highlight ↑↓ โดยไม่ยืนยัน ─────────────
	highlightOnly := func(id int) {
		if id < 0 || id >= len(filteredRecords) {
			return
		}
		selectedIdx = id
		list.Select(widget.ListItemID(selectedIdx))
		list.ScrollTo(widget.ListItemID(selectedIdx))
	}

	searchEntry := newSmartEntry(nil)
	searchEntry.SetPlaceHolder("พิมพ์เพื่อค้นหา... (↑↓=เลือก  Enter=ยืนยัน  Esc=ปิด)")

	list = widget.NewList(
		func() int { return len(filteredRecords) },
		func() fyne.CanvasObject {
			btn := widget.NewButton("", nil)
			btn.Importance = widget.LowImportance
			contentRow := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabel("")), // A/C Code
				container.NewGridWrap(fyne.NewSize(210, 30), widget.NewLabel("")), // A/C Name
				container.NewGridWrap(fyne.NewSize(140, 30), widget.NewLabel("")), // Beg. Balance
				container.NewGridWrap(fyne.NewSize(140, 30), widget.NewLabel("")), // Closing Bal.
				container.NewGridWrap(fyne.NewSize(130, 30), widget.NewLabel("")), // Last Per12
			)
			return container.NewMax(btn, contentRow)
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			r := filteredRecords[id]
			maxCtr := o.(*fyne.Container)
			btn := maxCtr.Objects[0].(*widget.Button)
			btn.OnTapped = func() { doSelect(int(id)) }
			row := maxCtr.Objects[1].(*fyne.Container)
			row.Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.AcCode)
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.AcName)
			row.Objects[2].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.BBal)
			row.Objects[3].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.CBal)
			row.Objects[4].(*fyne.Container).Objects[0].(*widget.Label).SetText(r.Lastper[11])
		},
	)
	// ลบ OnSelected ออก — ดักการคลิกที่ btn.OnTapped แทนแล้ว

	// ── Arrow Down ────────────────────────────────────────────────────
	searchEntry.onDown = func() {
		if len(filteredRecords) == 0 {
			return
		}
		next := selectedIdx + 1
		if next >= len(filteredRecords) {
			next = len(filteredRecords) - 1
		}
		highlightOnly(next)
	}

	// ── Arrow Up ──────────────────────────────────────────────────────
	searchEntry.onUp = func() {
		if len(filteredRecords) == 0 {
			return
		}
		prev := selectedIdx - 1
		if prev < 0 {
			prev = 0
		}
		highlightOnly(prev)
	}

	// ── Enter — ยืนยัน selectedIdx ───────────────────────────────────
	searchEntry.onEnter = func() {
		doSelect(selectedIdx)
	}

	// ── Esc ──────────────────────────────────────────────────────────
	searchEntry.onEsc = func() {
		pop.Hide()
	}

	// ── Filter ───────────────────────────────────────────────────────
	searchEntry.OnChanged = func(keyword string) {
		keyword = strings.ToLower(strings.TrimSpace(keyword))
		filteredRecords = nil
		selectedIdx = 0
		for _, r := range allRecords {
			if keyword == "" ||
				strings.Contains(strings.ToLower(r.AcCode), keyword) ||
				strings.Contains(strings.ToLower(r.AcName), keyword) {
				filteredRecords = append(filteredRecords, r)
			}
		}
		list.UnselectAll()
		list.Refresh()
		if len(filteredRecords) > 0 {
			highlightOnly(0)
		}
	}

	header := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(100, 30), widget.NewLabelWithStyle("A/C CODE", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(210, 30), widget.NewLabelWithStyle("A/C NAME", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(140, 30), widget.NewLabelWithStyle("Beg. Balance", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(140, 30), widget.NewLabelWithStyle("Closing Bal.", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(130, 30), widget.NewLabelWithStyle("Last Per12", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
	)

	closeBtn := widget.NewButton("✕ ปิด", func() { pop.Hide() })

	content := container.NewBorder(
		container.NewVBox(searchEntry, header),
		nil, nil, nil,
		container.NewGridWrap(fyne.NewSize(760, 400), list),
	)

	pop = widget.NewModalPopUp(
		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("ค้นหา Account Code", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				closeBtn,
			),
			content,
		),
		w.Canvas(),
	)

	// highlight แถวแรกโดยไม่ยืนยัน
	// ⚠️ ต้อง UnselectAll ก่อน Show เสมอ เพื่อให้ OnSelected ยิงได้แม้คลิกแถวแรก
	list.UnselectAll()
	pop.Show()
	if len(filteredRecords) > 0 {
		highlightOnly(0)
	}

	// Focus searchEntry หลัง popup render เสร็จ
	go func() {
		time.Sleep(350 * time.Millisecond)
		fyne.Do(func() {
			if w.Canvas() != nil {
				w.Canvas().Focus(nil)
				w.Canvas().Focus(searchEntry)
			}
		})
	}()
}
