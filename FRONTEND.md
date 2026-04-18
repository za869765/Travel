# 差旅明細整理系統 — 前端架構

> 檔案：`TRAVEL/index.html`（單檔 1458 行 HTML + 內嵌 CSS + JS）
> 版本：v6.6.7
> 部署：https://travel-8i3.pages.dev/

---

## WHAT — 技術棧

| 項目 | 實作 |
|------|------|
| 架構 | 純前端單頁（SPA），無後端 |
| Excel 解析 | [xlsx-js-style@1.2.0](https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js) CDN |
| 樣式 | 內嵌 CSS（無外部框架） |
| 儲存 | `localStorage`（記憶上次選擇的衛生所） |
| 部署 | Cloudflare Pages |

> **注意**：`CLAUDE.md` 標示 Python 3 / openpyxl / xlrd，但線上版其實是純瀏覽器 JS（`xlsx-js-style`）。Python 版可能是舊版本或另一條路線。

---

## WHY — 目的

讀取臺南市 37 間衛生所繳交的各式差旅／加班費 XLS 檔，自動：
1. 偵測各種檔案／分頁格式（8 種）
2. 拆分「姓名\$金額\*N人」這類合併欄位
3. 金額三方對帳（原始 = 明細 = 個人小計）
4. 匯出 A4 直式列印用 Excel 總表（含樣式／合併儲存格／列印設定）

---

## HOW — 結構

### 1. UI 區塊
| 區塊 ID | 用途 |
|---------|------|
| `uploadCard` | 拖曳 / 點選上傳 XLS |
| `progressArea` + `progressFill` + `logArea` | 進度條 + 逐步 log |
| `resultArea` | 結果總覽（統計卡片） |
| `officeSelect` | 衛生所篩選下拉 |
| `detailCard` / `officeTable` | 明細表 |
| `personCard` / `personDetail` | 個人小計表 |
| `noDataModal` | 查無資料彈窗 |

### 2. 資料流

```
上傳 XLS
  ↓ XLSX.read (ArrayBuffer)
偵測格式 isExpenseReportFile / detectSheetType
  ├─ 代辦經費彙整表 → extractExpenseReportSheet（3 種策略）
  ├─ overtime_detail → extractOvertimeDetail
  ├─ overtime_summary → extractOvertimeSummary
  ├─ travel_detail → extractTravelDetail
  ├─ travel_traffic → extractTravelTraffic
  ├─ travel_list → extractTravelList
  ├─ voucher_detail → extractVoucherDetail
  └─ self_output → extractSelfOutput（讀自己產出的總表）
  ↓
splitPersons + pushSplitRecords（姓名拆分＋浮點校正）
  ↓
resolveOffice / cleanSheetName / cleanPersonName（正規化）
  ↓
過濾 isValidOffice
  ↓
個人聚合 personMap（office|person 為 key）
  ↓
三方對帳 sourceByOffice / detailByOffice / personByOffice
  ↓
showResult → filterByOffice → renderPersonTable
  ↓ （可選）
downloadExcel（產 A4 直式 3 區塊 xlsx）
```

### 3. 核心常數
| 常數 | 內容 |
|------|------|
| `OFFICE_CODE_MAP` | z01~z37 → 衛生所名稱（37 間） |
| `OFFICE_NAME_MAP` | 反向對照 |
| `OFFICE_ORDER` | 顯示順序 |
| `OFFICE_CODE_REVERSE` | 衛生所 → z-code |

### 4. 核心函式

#### 檔案解析
- `handleFile(file)` / `processFile(file)` — 主流程
- `resolveOffice(name)` — 模糊比對衛生所名
- `extractFileTitle(filename)` — 支援 3 種檔名格式抽標題
- `splitPersons(personStr, totalAmount)` — v6.6.3「姓名\$金額\*N人」解析
- `pushSplitRecords(records, base, person, amount)` — v6.6.4 浮點數金額校正

#### 分頁類型偵測
- `detectSheetType(sheet, name, idx)` — 6+2 種類型分類器
- `findHeaderRow(sheet, keyword)` — 智慧標題列偵測
- `detectColumns(sheet, headerRow)` — 智慧欄位偵測（衛生所 / 領受人 / 金額 / 備註）
- `isExpenseReportFile(wb, filename)` — 代辦經費彙整格式偵測

#### 各類型擷取器
- `extractOvertimeDetail` — 加班明細
- `extractOvertimeSummary` — 加班／差旅／其他費用彙整
- `extractTravelDetail` — 差旅明細
- `extractTravelTraffic` — 交通差旅
- `extractTravelList` — 旅費清單
- `extractVoucherDetail` — 傳票明細（col0 姓名 / col10 小計）
- `extractSelfOutput` — 讀本工具產出的整理總表
- `extractExpenseReportSheet` — 代辦經費彙整（3 策略：右側面板 / 簡易 / z-code）

#### 輔助
- `cleanSheetName(name)` — 去除頁碼前綴
- `cleanPersonName(s)` — 姓名正規化
- `isSkipRow(row)` — 合計/總計列過濾
- `isValidOffice(o)` — 衛生所名合法性
- `safeFloat(v)` — 安全數字轉換
- `getRow(sheet, r)` / `sheetRows(sheet)` / `sheetCols(sheet)` — sheet 存取

#### UI 渲染
- `showResult()` — 顯示統計卡片＋衛生所下拉
- `filterByOffice()` — 依衛生所過濾並渲染
- `sortRecordsByPersonOrder(records, personList)` — v6.6.5 同一人列聚合
- `renderPersonTable(filterOffice)` — 個人小計表
- `showNoDataModal(filename)` — 查無資料彈窗

#### Excel 匯出
- `downloadExcel()` — 核心匯出函式
  - 6 欄寬度（MAX_COLS = 6）
  - 3 區塊：全部明細 / 個人小計 / 加班費彙整
  - 樣式：thin / medium / double 三種邊框
  - 金額欄 `z = '#,##0'`（千分位）
  - 合併儲存格：標題跨 A:F、明細事由 E:F、加班費事由 D:F
  - 列印：A4 直式（paper: 9）、fitToWidth: 1、頁尾「第 X 頁，共 Y 頁」

### 5. v6.6.x 版本重點
- **v6.6.3**：`splitPersons` 解析「姓名\$金額\*N人」
- **v6.6.4**：`pushSplitRecords` 浮點校正；三方金額把關
- **v6.6.5**：明細依 personList 順序排序（同一人列聚合）
- **v6.6.7**：當前部署版本

---

## 與 CLAUDE.md 的差異

`TRAVEL/CLAUDE.md` 寫的是：
- Python 3
- openpyxl / xlrd
- `travel.py` 主程式

但實際 `index.html` 是純瀏覽器 JS，並無 `travel.py`。`CLAUDE.md` 可能是舊規劃文件，若要更新建議改寫為前端版本。
