package main

// report_merge_ui.go  v2 — Master Marker Edition
// ─────────────────────────────────────────────────────────────────
// UI สำหรับ Merge Report
//
// Key UX:
//   - ตารางไฟล์สาขา มีปุ่ม [👑] เลือก Master (Radio — เลือกได้แค่ 1)
//   - ไฟล์แรกที่ Add → เป็น Master อัตโนมัติ
//   - คลิก [👑] ที่แถวอื่น → ย้าย Master ไป
//   - แสดง Badge "👑 Master" ชัดเจนในตาราง
//   - Warning panel แสดงบัญชีนอกผัง Master ทันทีหลัง Validate
// ─────────────────────────────────────────────────────────────────

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"image/color"
)

// showMergeReportDialog — เปิด dialog หลักของ Merge Report
func showMergeReportDialog(w fyne.Window) {

	// ════════════════════════════════════════════════════
	// State
	// ════════════════════════════════════════════════════
	var branches []*BranchInfo
	masterIdx := -1    // index ใน branches ที่เป็น Master
	selectedIdx := -1  // index ที่ user คลิก highlight ใน list

	// ════════════════════════════════════════════════════
	// Status / Info labels
	// ════════════════════════════════════════════════════
	statusLbl := widget.NewLabel("")
	statusLbl.Wrapping = fyne.TextWrapWord

	warningBox := container.NewVBox() // แสดง off-chart warnings หลัง validate

	lblCount := widget.NewLabelWithStyle(
		"ยังไม่ได้เลือกไฟล์",
		fyne.TextAlignLeading,
		fyne.TextStyle{Italic: true},
	)

	// ════════════════════════════════════════════════════
	// Branch List Widget
	// ════════════════════════════════════════════════════
	var branchList *widget.List

	refreshList := func() {
		n := len(branches)
		if n == 0 {
			lblCount.SetText("ยังไม่ได้เลือกไฟล์")
		} else {
			masterName := ""
			if masterIdx >= 0 && masterIdx < n {
				masterName = branches[masterIdx].BranchName
			}
			lblCount.SetText(fmt.Sprintf("%d ไฟล์  |  Master: %s", n, masterName))
		}
		// sync IsMaster flag
		for i, b := range branches {
			b.IsMaster = (i == masterIdx)
		}
		branchList.Refresh()
		selectedIdx = -1
	}

	// Row layout: [👑 btn][#][BranchName][TaxID][YearEnd][Periods][NowPer][✕ btn]
	branchList = widget.NewList(
		func() int { return len(branches) },
		func() fyne.CanvasObject {
			masterBtn := widget.NewButton("👑", nil)
			masterBtn.Importance = widget.LowImportance
			removeBtn := widget.NewButtonWithIcon("", theme.DeleteIcon(), nil)
			removeBtn.Importance = widget.LowImportance

			row := container.NewHBox(
				container.NewGridWrap(fyne.NewSize(38, 28), masterBtn),
				container.NewGridWrap(fyne.NewSize(22, 28), widget.NewLabel("")),  // #
				container.NewGridWrap(fyne.NewSize(190, 28), widget.NewLabel("")), // name
				container.NewGridWrap(fyne.NewSize(130, 28), widget.NewLabel("")), // taxID
				container.NewGridWrap(fyne.NewSize(90, 28), widget.NewLabel("")),  // yearEnd
				container.NewGridWrap(fyne.NewSize(55, 28), widget.NewLabel("")),  // periods
				container.NewGridWrap(fyne.NewSize(65, 28), widget.NewLabel("")),  // nowPer
				container.NewGridWrap(fyne.NewSize(32, 28), removeBtn),
			)
			return row
		},
		func(id widget.ListItemID, o fyne.CanvasObject) {
			if int(id) >= len(branches) {
				return
			}
			b := branches[id]
			row := o.(*fyne.Container)

			// Master button
			masterBtn := row.Objects[0].(*fyne.Container).Objects[0].(*widget.Button)
			if int(id) == masterIdx {
				masterBtn.SetText("👑")
				masterBtn.Importance = widget.HighImportance
			} else {
				masterBtn.SetText("○")
				masterBtn.Importance = widget.LowImportance
			}
			capturedID := int(id)
			masterBtn.OnTapped = func() {
				masterIdx = capturedID
				statusLbl.SetText(fmt.Sprintf("✅ กำหนด Master: %s", branches[capturedID].BranchName))
				refreshList()
			}

			// # label
			row.Objects[1].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("%d", id+1))

			// BranchName — เพิ่ม suffix ถ้าเป็น Master
			nameText := b.BranchName
			if int(id) == masterIdx {
				nameText = "👑 " + b.BranchName
			}
			row.Objects[2].(*fyne.Container).Objects[0].(*widget.Label).SetText(nameText)
			row.Objects[3].(*fyne.Container).Objects[0].(*widget.Label).SetText(b.TaxID)
			row.Objects[4].(*fyne.Container).Objects[0].(*widget.Label).SetText(b.YearEnd.Format("02/01/06"))
			row.Objects[5].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("%d", b.TotalPeriods))
			row.Objects[6].(*fyne.Container).Objects[0].(*widget.Label).SetText(fmt.Sprintf("งวด %d", b.NowPeriod))

			// Remove button
			removeBtn := row.Objects[7].(*fyne.Container).Objects[0].(*widget.Button)
			removeBtn.OnTapped = func() {
				name := branches[capturedID].BranchName
				wasMaster := (capturedID == masterIdx)
				branches = append(branches[:capturedID], branches[capturedID+1:]...)
				if wasMaster {
					if len(branches) > 0 {
						masterIdx = 0
					} else {
						masterIdx = -1
					}
				} else if capturedID < masterIdx {
					masterIdx--
				}
				statusLbl.SetText(fmt.Sprintf("🗑️  ลบ: %s", name))
				refreshList()
			}
		},
	)
	branchList.OnSelected = func(id widget.ListItemID) {
		selectedIdx = int(id)
	}

	// List Header
	listHeader := container.NewHBox(
		container.NewGridWrap(fyne.NewSize(38, 22), widget.NewLabelWithStyle("Master", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(22, 22), widget.NewLabel("#")),
		container.NewGridWrap(fyne.NewSize(190, 22), widget.NewLabelWithStyle("ชื่อสาขา/บริษัท", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(130, 22), widget.NewLabelWithStyle("Tax ID", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(90, 22), widget.NewLabelWithStyle("YearEnd", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(55, 22), widget.NewLabelWithStyle("Periods", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(65, 22), widget.NewLabelWithStyle("งวดปัจจุบัน", fyne.TextAlignLeading, fyne.TextStyle{Bold: true})),
		container.NewGridWrap(fyne.NewSize(32, 22), widget.NewLabel("")),
	)

	// ════════════════════════════════════════════════════
	// Add File popup
	// ════════════════════════════════════════════════════
	addFileByName := func(dbName string) {
		dbName = strings.TrimSuffix(strings.TrimSpace(dbName), ".xlsx")
		if dbName == "" {
			return
		}
		fullPath := filepath.Join(".", "dont_edit", dbName+".xlsx")
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			statusLbl.SetText("❌ ไม่พบไฟล์: " + fullPath)
			return
		}
		for _, b := range branches {
			if b.FilePath == fullPath {
				statusLbl.SetText("⚠️  ไฟล์นี้เพิ่มไปแล้ว: " + dbName)
				return
			}
		}
		info, err := LoadBranchInfo(fullPath)
		if err != nil {
			statusLbl.SetText("❌ " + err.Error())
			return
		}
		branches = append(branches, info)
		// ไฟล์แรก → เป็น Master อัตโนมัติ
		if len(branches) == 1 {
			masterIdx = 0
			statusLbl.SetText(fmt.Sprintf("✅ เพิ่ม [👑 Master]: %s  (YearEnd: %s, Period: %d/%d)",
				info.BranchName, info.YearEnd.Format("02/01/2006"), info.NowPeriod, info.TotalPeriods))
		} else {
			statusLbl.SetText(fmt.Sprintf("✅ เพิ่ม: %s  (YearEnd: %s, Period: %d/%d)",
				info.BranchName, info.YearEnd.Format("02/01/2006"), info.NowPeriod, info.TotalPeriods))
		}
		refreshList()
	}

	showAddPopup := func() {
		allNames := loadDBNamesForMerge()
		if len(allNames) == 0 {
			statusLbl.SetText("❌ ไม่พบไฟล์ใน ./dont_edit/")
			return
		}
		filteredNames := append([]string{}, allNames...)
		selPopIdx := 0

		var popList *widget.List
		var pop *widget.PopUp

		searchEn := newSmartEntry(nil)
		searchEn.SetPlaceHolder("ค้นหา... (↑↓ เลื่อน, Enter เลือก, Esc ปิด)")

		doSelect := func(idx int) {
			if idx < 0 || idx >= len(filteredNames) {
				return
			}
			pop.Hide()
			addFileByName(filteredNames[idx])
		}

		searchEn.onEsc = func() { pop.Hide() }
		searchEn.onEnter = func() { doSelect(selPopIdx) }
		searchEn.onDown = func() {
			if selPopIdx < len(filteredNames)-1 {
				selPopIdx++
				popList.Select(widget.ListItemID(selPopIdx))
				popList.ScrollTo(widget.ListItemID(selPopIdx))
			}
		}
		searchEn.onUp = func() {
			if selPopIdx > 0 {
				selPopIdx--
				popList.Select(widget.ListItemID(selPopIdx))
				popList.ScrollTo(widget.ListItemID(selPopIdx))
			}
		}
		searchEn.OnChanged = func(kw string) {
			kw = strings.ToLower(strings.TrimSpace(kw))
			filteredNames = nil
			selPopIdx = 0
			for _, n := range allNames {
				if kw == "" || strings.Contains(strings.ToLower(n), kw) {
					filteredNames = append(filteredNames, n)
				}
			}
			popList.UnselectAll()
			popList.Refresh()
			if len(filteredNames) > 0 {
				popList.Select(0)
			}
		}

		popList = widget.NewList(
			func() int { return len(filteredNames) },
			func() fyne.CanvasObject {
				btn := widget.NewButton("", nil)
				btn.Importance = widget.LowImportance
				lbl := widget.NewLabel("")
				return container.NewMax(btn,
					container.NewHBox(container.NewGridWrap(fyne.NewSize(460, 28), lbl)))
			},
			func(id widget.ListItemID, o fyne.CanvasObject) {
				if int(id) >= len(filteredNames) {
					return
				}
				ctr := o.(*fyne.Container)
				ctr.Objects[0].(*widget.Button).OnTapped = func() { doSelect(int(id)) }
				ctr.Objects[1].(*fyne.Container).Objects[0].(*fyne.Container).Objects[0].(*widget.Label).SetText(filteredNames[id])
			},
		)
		if len(filteredNames) > 0 {
			popList.Select(0)
		}

		pop = widget.NewModalPopUp(
			container.NewVBox(
				container.NewHBox(
					widget.NewLabelWithStyle("เลือกฐานข้อมูลสาขา", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
					layout.NewSpacer(),
					widget.NewButton("✕ ปิด", func() { pop.Hide() }),
				),
				searchEn,
				container.NewGridWrap(fyne.NewSize(490, 300), popList),
			),
			w.Canvas(),
		)
		pop.Resize(fyne.NewSize(510, 410))
		pop.Show()
		go func() { fyne.Do(func() { w.Canvas().Focus(searchEn) }) }()
	}

	// ════════════════════════════════════════════════════
	// Options
	// ════════════════════════════════════════════════════
	chkTB := widget.NewCheck("Trial Balance", nil); chkTB.SetChecked(true)
	chkPnL := widget.NewCheck("Profit & Loss", nil); chkPnL.SetChecked(true)
	chkBS := widget.NewCheck("Balance Sheet", nil); chkBS.SetChecked(true)
	chkMatrix := widget.NewCheck("แสดงแยกคอลัมน์ตามสาขา (Matrix)", nil)
	chkOffChart := widget.NewCheck("รวมบัญชีนอกผัง Master ในงบรวม", nil)
	chkEliminate := widget.NewCheck("ตัดรายการระหว่างกัน (Eliminate)", nil)
	enEliminateCodes := newSmartEntry(nil)
	enEliminateCodes.SetPlaceHolder("AC Code คั่นด้วย , เช่น 113001,210001")
	enEliminateCodes.Disable()
	chkEliminate.OnChanged = func(v bool) {
		if v {
			enEliminateCodes.Enable()
		} else {
			enEliminateCodes.Disable()
		}
	}

	periodOpts := []string{"ใช้งวดปัจจุบันของแต่ละสาขา (Auto)"}
	for i := 1; i <= 12; i++ {
		periodOpts = append(periodOpts, fmt.Sprintf("งวด %d", i))
	}
	selPeriod := widget.NewSelect(periodOpts, nil)
	selPeriod.SetSelectedIndex(0)

	// ════════════════════════════════════════════════════
	// Output path
	// ════════════════════════════════════════════════════
	enOut := newSmartEntry(nil)
	enOut.SetText(fmt.Sprintf("./merge_report_%s.xlsx", time.Now().Format("20060102_1504")))

	btnBrowseOut := widget.NewButtonWithIcon("", theme.FolderOpenIcon(), func() {
		dialog.ShowFileSave(func(uri fyne.URIWriteCloser, err error) {
			if err != nil || uri == nil {
				return
			}
			p := uri.URI().Path()
			if !strings.HasSuffix(strings.ToLower(p), ".xlsx") {
				p += ".xlsx"
			}
			enOut.SetText(p)
			uri.Close()
		}, w)
	})

	// ════════════════════════════════════════════════════
	// Validate Button
	// ════════════════════════════════════════════════════
	btnValidate := widget.NewButtonWithIcon("ตรวจสอบ (Validate)", theme.ConfirmIcon(), func() {
		warningBox.Objects = nil
		warningBox.Refresh()

		if len(branches) < 2 {
			statusLbl.SetText("⚠️  ต้องเลือกอย่างน้อย 2 ไฟล์")
			return
		}
		if err := ValidateFilesForMerge(branches); err != nil {
			statusLbl.SetText("❌ " + err.Error())
			return
		}

		// ทดลอง build master dict เพื่อ preview warnings
		var master *BranchInfo
		for _, b := range branches {
			if b.IsMaster {
				master = b
				break
			}
		}
		masterDict, _, err := BuildMasterDictionary(master)
		if err != nil {
			statusLbl.SetText("❌ Build Master Dict: " + err.Error())
			return
		}

		// หา off-chart เร็วๆ
		xlOpts := excelize.Options{Password: "@A123456789a"}
		offChartPreviews := []string{}
		for _, b := range branches {
			if b.IsMaster {
				continue
			}
			f2, err2 := excelize.OpenFile(b.FilePath, xlOpts)
			if err2 != nil {
				continue
			}
			rows, _ := f2.GetRows("Ledger_Master")
			for i, row := range rows {
				if i == 0 || len(row) < 2 || safeGet(row, 0) != b.ComCode {
					continue
				}
				ac := strings.TrimSpace(safeGet(row, 1))
				if ac == "" {
					continue
				}
				if _, inM := masterDict[ac]; !inM {
					offChartPreviews = append(offChartPreviews,
						fmt.Sprintf("  ⚠️  [%s] %s — มีใน %s แต่ไม่มีในผัง Master",
							ac, strings.TrimSpace(safeGet(row, 2)), b.BranchName))
				}
			}
			f2.Close()
		}

		// แสดง warning ใน warningBox
		if len(offChartPreviews) > 0 {
			warningBox.Add(widget.NewLabelWithStyle(
				fmt.Sprintf("⚠️  พบบัญชีนอกผัง Master %d รายการ:", len(offChartPreviews)),
				fyne.TextAlignLeading, fyne.TextStyle{Bold: true},
			))
			// แสดงสูงสุด 8 รายการ
			limit := len(offChartPreviews)
			if limit > 8 {
				limit = 8
			}
			for _, w2 := range offChartPreviews[:limit] {
				warningBox.Add(widget.NewLabel(w2))
			}
			if len(offChartPreviews) > 8 {
				warningBox.Add(widget.NewLabel(fmt.Sprintf("  ... และอีก %d รายการ (ดูใน Sheet ⚠ บัญชีนอกผัง)", len(offChartPreviews)-8)))
			}
			warningBox.Refresh()
			statusLbl.SetText(fmt.Sprintf("✅ Validate ผ่าน — แต่มีบัญชีนอกผัง %d รายการ (ดู Warning ด้านล่าง)", len(offChartPreviews)))
		} else {
			statusLbl.SetText(fmt.Sprintf("✅ Validate ผ่าน — ผังบัญชีทุกสาขาตรงกับ Master ทั้งหมด"))
		}
	})

	// ════════════════════════════════════════════════════
	// Generate Button
	// ════════════════════════════════════════════════════
	btnGenerate := widget.NewButtonWithIcon("Generate Consolidated Report", theme.DocumentSaveIcon(), func() {
		warningBox.Objects = nil
		warningBox.Refresh()

		if len(branches) < 2 {
			statusLbl.SetText("❌ ต้องเลือกอย่างน้อย 2 ไฟล์")
			return
		}
		if !chkTB.Checked && !chkPnL.Checked && !chkBS.Checked {
			statusLbl.SetText("❌ กรุณาเลือกรายงานอย่างน้อย 1 รายการ")
			return
		}
		outPath := strings.TrimSpace(enOut.Text)
		if outPath == "" {
			statusLbl.SetText("❌ กรุณาระบุที่บันทึกไฟล์")
			return
		}
		if !strings.HasSuffix(strings.ToLower(outPath), ".xlsx") {
			outPath += ".xlsx"
		}

		if err := ValidateFilesForMerge(branches); err != nil {
			statusLbl.SetText("❌ " + err.Error())
			return
		}

		targetPeriod := 0
		if selPeriod.SelectedIndex() > 0 {
			targetPeriod = selPeriod.SelectedIndex()
		}
		var interCodes []string
		if chkEliminate.Checked {
			for _, s := range strings.Split(enEliminateCodes.Text, ",") {
				code := strings.TrimSpace(s)
				if code != "" {
					interCodes = append(interCodes, code)
				}
			}
		}
		opts := MergeOptions{
			IncludeTrialBalance:  chkTB.Checked,
			IncludePnL:           chkPnL.Checked,
			IncludeBalanceSheet:  chkBS.Checked,
			MatrixLayout:         chkMatrix.Checked,
			EliminateInterBranch: chkEliminate.Checked,
			InterBranchCodes:     interCodes,
			NowPeriod:            targetPeriod,
			IncludeOffChart:      chkOffChart.Checked,
		}

		statusLbl.SetText("⏳ กำลังรวมข้อมูล...")

		go func() {
			// Step 1: Build Master Dictionary
			var master *BranchInfo
			for _, b := range branches {
				if b.IsMaster {
					master = b
					break
				}
			}
			masterDict, masterOrder, err := BuildMasterDictionary(master)
			if err != nil {
				fyne.Do(func() { statusLbl.SetText("❌ Master Dict: " + err.Error()) })
				return
			}

			// Step 2: Aggregate
			aggResult, err := AggregateLedgers(branches, masterDict, opts)
			if err != nil {
				fyne.Do(func() { statusLbl.SetText("❌ Aggregate: " + err.Error()) })
				return
			}

			// Step 3: Export
			if err := ExportMergeReport(branches, masterDict, masterOrder, aggResult, opts, outPath); err != nil {
				fyne.Do(func() { statusLbl.SetText("❌ Export: " + err.Error()) })
				return
			}

			fyne.Do(func() {
				// แสดง warnings ใน warningBox
				if len(aggResult.Warnings) > 0 {
					warningBox.Add(widget.NewLabelWithStyle(
						fmt.Sprintf("⚠️  พบบัญชีนอกผัง %d รายการ (ดูได้ใน Sheet ⚠ บัญชีนอกผัง):",
							len(aggResult.Warnings)),
						fyne.TextAlignLeading, fyne.TextStyle{Bold: true},
					))
					limit := len(aggResult.Warnings)
					if limit > 5 {
						limit = 5
					}
					for _, ww := range aggResult.Warnings[:limit] {
						warningBox.Add(widget.NewLabel(ww))
					}
					warningBox.Refresh()
				}

				statusLbl.SetText(fmt.Sprintf("✅ สร้างรายงานสำเร็จ → %s", outPath))

				// Success dialog
				var d dialog.Dialog
				okBtn := newEnterButton("OK", func() { d.Hide() })
				okBtn.Importance = widget.HighImportance

				sheets := []string{}
				if opts.IncludeTrialBalance {
					sheets = append(sheets, "• Trial Balance รวม")
				}
				if opts.IncludePnL {
					sheets = append(sheets, "• P&L รวม")
				}
				if opts.IncludeBalanceSheet {
					sheets = append(sheets, "• Balance Sheet รวม")
				}
				if len(aggResult.OffChart) > 0 {
					sheets = append(sheets, fmt.Sprintf("• ⚠ บัญชีนอกผัง (%d รายการ)", len(aggResult.OffChart)))
				}
				sheets = append(sheets, "• สรุปสาขา")

				sheetList := strings.Join(sheets, "\n")
				msg := fmt.Sprintf(
					"สร้างรายงานสำเร็จ!\n\nไฟล์: %s\n\nSheets ที่สร้าง:\n%s",
					filepath.Base(outPath), sheetList,
				)
				if len(aggResult.Warnings) > 0 {
					msg += fmt.Sprintf("\n\n⚠️  มีบัญชีนอกผัง Master %d รายการ\nตรวจสอบใน Sheet ⚠ บัญชีนอกผัง", len(aggResult.Warnings))
				}

				d = dialog.NewCustomWithoutButtons("✅ Generate สำเร็จ",
					container.NewVBox(
						widget.NewLabel(msg),
						container.NewCenter(okBtn),
					), w)
				d.Show()
				go func() { fyne.Do(func() { w.Canvas().Focus(okBtn) }) }()
			})
		}()
	})
	btnGenerate.Importance = widget.HighImportance

	// ════════════════════════════════════════════════════
	// Master Legend / Info box
	// ════════════════════════════════════════════════════
	masterInfoText := canvas.NewText(
		"👑 = Master (ผังบัญชีหลัก)  |  ○ = คลิกเพื่อกำหนดเป็น Master  |  ไฟล์แรกที่เพิ่มจะเป็น Master อัตโนมัติ",
		color.RGBA{R: 100, G: 60, B: 0, A: 255},
	)
	masterInfoText.TextSize = 10

	// ════════════════════════════════════════════════════
	// Full Layout
	// ════════════════════════════════════════════════════

	// Section 1: File list
	fileToolbar := container.NewHBox(
		widget.NewButtonWithIcon("เพิ่มไฟล์สาขา", theme.ContentAddIcon(), func() { showAddPopup() }),
		widget.NewSeparator(),
		widget.NewLabel("Period:"),
		container.NewGridWrap(fyne.NewSize(240, 28), selPeriod),
		layout.NewSpacer(),
		lblCount,
	)

	fileSection := container.NewVBox(
		widget.NewLabelWithStyle("📁 ไฟล์ฐานข้อมูลสาขา", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		masterInfoText,
		fileToolbar,
		listHeader,
		container.NewGridWrap(fyne.NewSize(650, 150), branchList),
	)

	// Section 2: Options
	optSection := container.NewVBox(
		widget.NewLabelWithStyle("📋 รายงานที่ต้องการ", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewHBox(chkTB, chkPnL, chkBS),
		container.NewHBox(chkMatrix, chkOffChart),
		chkEliminate,
		container.NewHBox(
			widget.NewLabel("  AC Codes:"),
			container.NewGridWrap(fyne.NewSize(380, 28), enEliminateCodes),
		),
	)

	// Section 3: Output + Buttons
	outSection := container.NewVBox(
		widget.NewLabelWithStyle("💾 บันทึกรายงาน", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil,
			widget.NewLabel("บันทึกที่:"),
			btnBrowseOut,
			enOut,
		),
		container.NewHBox(
			btnGenerate,
			layout.NewSpacer(),
			btnValidate,
		),
	)

	// Warning / Status section
	statusSection := container.NewVBox(
		widget.NewSeparator(),
		statusLbl,
		warningBox,
	)

	mainContent := container.NewVBox(
		fileSection,
		widget.NewSeparator(),
		optSection,
		widget.NewSeparator(),
		outSection,
		statusSection,
	)

	var d dialog.Dialog
	closeBtn := widget.NewButton("ปิด", func() { d.Hide() })

	d = dialog.NewCustomWithoutButtons(
		"Merge Report — รวมงบหลายสาขา (Master Marker)",
		container.NewVBox(
			mainContent,
			widget.NewSeparator(),
			container.NewHBox(layout.NewSpacer(), closeBtn),
		),
		w,
	)
	d.Resize(fyne.NewSize(720, 660))
	d.Show()
}
