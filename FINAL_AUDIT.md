# Parcel Ops Dashboard - Final Audit v4

## ตรวจระบบที่แก้
- URL สะอาด ไม่มี `?v=xx`
- Dropdown รายงานจัดเป็นหมวดหมู่ด้วย `optgroup`
- ตัวเลือกหลักของรายงาน:
  1. รูปแบบรายงานไลน์
  2. ขนาดรายงาน
  3. กลุ่มงานที่ต้องการรายงาน
  4. เลขพัสดุในรายงาน
- รายการพัสดุรองรับ:
  - ตัวอย่าง 5 เลข/ช่วงเวลา
  - เลขพัสดุทั้งหมด
  - แยกเลขตามช่วงเวลาในคลัง
- สรุปหน้าอื่นรองรับ:
  - เวลาในคลัง
  - จำนวนชิ้น
  - เวลาในคลัง + จำนวนชิ้น
  - กรองเฉพาะงาน 22–24 / >24 / >48
- กราฟมีตัวเลขบนแท่งกราฟ
- ล้างข้อมูลล้าง storage/cache และกลับหน้าแรก

## Final v5
- เปลี่ยน `กลุ่มงานที่ต้องการรายงาน` จาก dropdown เป็น checkbox
- เลือกหลายกลุ่มพร้อมกันได้ เช่น `มากกว่า 24 ชม.` + `มากกว่า 48 ชม.`
- ถ้าเลือก `ทั้งหมด` ระบบจะยกเลิกตัวเลือกกลุ่มเสี่ยงอื่นให้อัตโนมัติ
- ถ้าเลือกกลุ่มเสี่ยง ระบบจะยกเลิก `ทั้งหมด` ให้อัตโนมัติ

## Final v6
- ตรวจรูปแบบข้อความคัดลอกและปรับให้ไม่เกิด `ชม.:`
- มาตรฐานข้อความเวลาในคลัง: `ช่วงเวลา = จำนวน ชิ้น`
- ตัวอย่าง: `≤6 ชม. = 1,093 ชิ้น | ≤12 ชม. = 367 ชิ้น`

## Final v7
- แก้ `captureTarget()` ให้หา target ที่เหมาะสมผ่าน `getBestCaptureTarget()`
- แก้ปัญหาบันทึกภาพมีพื้นที่ว่างด้านขวา
- ตั้งค่า `backgroundColor` เป็นขาว และใช้ขนาด scrollWidth/scrollHeight ของ card จริง

## Final v8
- ตรวจบั๊ก page navigation restore
- แก้ `ReferenceError: saved is not defined` ใน `writeControls`
- เพิ่ม backup workspace ลง `localStorage`
- เพิ่ม restore guard ไม่ให้ UI controls error ทำให้ข้อมูลไฟล์เดิมหาย

## Final v9
- แก้ capture ให้ robust กว่า v7/v8
- ไม่บังคับ width/height เอง เพื่อลด empty canvas / clone element issue

## Final v10
- เพิ่ม `captureTableFallback()`
- เพิ่ม `downloadCanvas()`
- ภาพตารางยังบันทึกได้แม้ html2canvas fail

## Final v11
- เพิ่ม focus แบบทุกช่วงเวลา
- เปลี่ยน default copyLineLimit = all
- เปลี่ยน default copyTrackMode = all
- copyTableForLine ใช้ state.filtered เพื่อแนบเลขพัสดุตามกลุ่มจริง

## Final v12
- เพิ่ม `captureTableCanvas()`
- เพิ่ม `captureChartCanvas()`
- แก้ปัญหาบันทึกภาพไม่สำเร็จด้วยการเลิกพึ่ง html2canvas เป็นหลัก

## Final v13
- ตรวจ copy issue: v12 มี `copyTableForLine()` เรียก `getCopySettings()` แต่ helper หาย
- เพิ่ม missing helpers ทั้งหมดกลับเข้าไป
- เพิ่ม robust clipboard fallback
- ตรวจ navigation restore: ใช้ IndexedDB + localStorage backup และ writeControls ใช้ `c.copyFocus`
- Clear data รวม cache แล้วผ่าน `wipeAllClientData(true)`

## Final v14
- Runtime helper audit:
  - timestampFile: restored
  - formatHours: restored
  - trackingNo: restored
  - shortAgingLabel: restored
- node --check passed

## Final v15
- copyTableForLine no longer always uses all `state.filtered`
- Added `copySourceRowsForTable()` to match table source:
  - branchSummaryTable/fdBranchStatusTable => FD rows for selected base
  - lhOtherTable/lhHubStatusTable => LH other rows for selected base
  - hubSummaryTable => overview + appended FD branch summary

## Final v16
- Added `bindNavigationPersistence()`
- Added `persistWorkspaceNow()`
- Changed restore order: sessionStorage -> IndexedDB -> legacy localStorage
- Prevents losing loaded Excel when navigating left menu pages

## Final v17
- Asset cache busting through renamed JS/CSS files, not visible page query strings
- Added `window.PARCEL_OPS_BUILD`
- Added `window.ParcelOpsStorageDebug()`

## Final v18
- `renderSummaryTables()` now excludes selected base HUB from overview HUB table when base selected
- `hubSummaryTable` copy now groups LH rows only when base selected
- FD branch rows remain appended to overview copy
- Asset cache-busted to v18

## Final v19
- SPA navigation with `bindSpaNavigation()`
- All HTML entrypoints share one combined page shell
- Overview table with selected base = FD branch rows + LH other hub rows in one table
- Asset cache-busted to v19

## Final v20
- Restored missing copy/report helper functions
- Hardened SPA navigation switching
- Exposed debug globals for copy functions
- Asset cache-busted to v20

## Final v21
- Restored Chart.js script include
- Deferred chart drawing with requestAnimationFrame + setTimeout after SPA panel display
- Added `window.ParcelOpsChartDebug()`
- Asset cache-busted to v21
