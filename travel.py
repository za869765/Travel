"""
差旅明細整理程式
讀取 XLS 旅費彙整表，整理出：
1. 指定衛生所（佳里區/Z20）的出差明細
2. 全機關加班費彙整
3. 每人小計
輸出成 Excel 總表
"""
import xlrd
import re
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

# ========== 設定 ==========
INPUT_FILE = "input/79,054元115年衛生所第3次旅費彙整表.xls"
OUTPUT_FILE = "output/差旅明細整理總表.xlsx"
TARGET_KEYWORDS = ["佳里", "z20", "Z20"]  # 目標衛生所識別關鍵字


def open_workbook(path):
    """開啟 XLS 檔案"""
    return xlrd.open_workbook(path, formatting_info=True)


def is_target(text):
    """判斷是否為目標衛生所"""
    if not text:
        return False
    text_upper = str(text).upper().strip()
    for kw in TARGET_KEYWORDS:
        if kw.upper() in text_upper:
            return True
    return False


def parse_date(val, wb):
    """嘗試解析日期值"""
    if isinstance(val, float) and val > 40000:
        # Excel 日期序列值
        try:
            date_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
            return f"{date_tuple[0]}/{date_tuple[1]}/{date_tuple[2]}"
        except:
            return str(int(val))
    return str(val).strip() if val else ""


def safe_float(val):
    """安全轉換為數值"""
    try:
        f = float(val)
        return f if f != 0 else 0
    except:
        return 0


def extract_type4_detail(wb, sheet_idx):
    """
    解析類型4: 差旅明細 (分頁 4, 愛滋差旅格式)
    格式: 衛生所 | 領受人 | 金額 | 總計 | 備註
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []
    current_office = ""

    for r in range(5, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]

        # 跳過合計、製表人等
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "製表" in row_text or "股長" in row_text:
            continue

        # 取得衛生所（col 1）
        office = str(row[1]).strip() if len(row) > 1 else ""
        if office and "衛生所" in office:
            current_office = office

        # 取得領受人（col 2）
        person = str(row[2]).strip() if len(row) > 2 else ""

        # 取得金額（col 5）
        amount = safe_float(row[5]) if len(row) > 5 else 0

        # 取得備註/事由（col 7）
        note = str(row[7]).strip() if len(row) > 7 else ""

        if amount > 0 and current_office:
            records.append({
                "分頁": sheet_name,
                "衛生所": current_office,
                "姓名": person if person else "",
                "日期": "",
                "金額": amount,
                "事由": note,
                "類型": "差旅費" if "差旅" in sheet_name else "加班費"
            })

    return records


def extract_type5_detail(wb, sheet_idx):
    """
    解析類型5: 登革熱交通費格式 (分頁 5, 6)
    格式: 姓名 | 出差日期 | 金額 | 區別 | 合計 | 備註
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []

    for r in range(4, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "製表" in row_text or not any(str(v).strip() for v in row):
            continue

        person = str(row[0]).strip() if len(row) > 0 else ""
        date_val = parse_date(row[1], wb) if len(row) > 1 else ""
        amount = safe_float(row[4]) if len(row) > 4 else 0
        district = str(row[5]).strip() if len(row) > 5 else ""
        note = str(row[7]).strip() if len(row) > 7 else ""

        if amount > 0 and person:
            # 從區別找衛生所名
            office = f"{district}衛生所" if district and "衛生所" not in district else district
            records.append({
                "分頁": sheet_name,
                "衛生所": office,
                "姓名": person,
                "日期": date_val,
                "金額": amount,
                "事由": note if note else sheet_name,
                "類型": "差旅費"
            })

    return records


def extract_type7_detail(wb, sheet_idx):
    """
    解析類型7: 教育訓練格式 (分頁 7)
    格式: 衛生所 | 金額 | 備註(姓名) | 教育訓練
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []

    for r in range(3, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "製表" in row_text:
            continue

        office_raw = str(row[0]).strip() if len(row) > 0 else ""
        amount = safe_float(row[1]) if len(row) > 1 else 0
        person = str(row[2]).strip() if len(row) > 2 else ""

        if amount > 0 and office_raw:
            # 移除前面的編號 "01." "02." 等
            office = re.sub(r'^\d+\.?', '', office_raw).strip()
            records.append({
                "分頁": sheet_name,
                "衛生所": office,
                "姓名": person,
                "日期": "",
                "金額": amount,
                "事由": sheet_name,
                "類型": "差旅費"
            })

    return records


def extract_type8_detail(wb, sheet_idx):
    """
    解析類型8: 旅費明細清單格式 (分頁 8~13)
    格式: 衛生所 | 金額 | 備註(含姓名) | ...
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []

    for r in range(3, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "製表" in row_text or "總計" in row_text:
            continue

        office = str(row[0]).strip() if len(row) > 0 else ""
        amount = safe_float(row[1]) if len(row) > 1 else 0
        note = str(row[2]).strip() if len(row) > 2 else ""

        if amount > 0 and office and "衛生所" in office:
            # 從備註提取姓名（通常格式為 "1/21洪紅華" 或 "1/21葉千嬅、嚴家翎$104*2人"）
            person = ""
            if note:
                # 嘗試提取日期後的姓名
                m = re.search(r'\d+/\d+\s*(.+)', note)
                if m:
                    person = m.group(1).strip()
                else:
                    person = note

            records.append({
                "分頁": sheet_name,
                "衛生所": office,
                "姓名": person,
                "日期": "",
                "金額": amount,
                "事由": sheet_name,
                "類型": "差旅費"
            })

    return records


def extract_type1_overtime(wb, sheet_idx):
    """
    解析加班費明細 (分頁 1, 2)
    格式同類型4
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []
    current_office = ""

    for r in range(5, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "製表" in row_text or "股長" in row_text:
            continue

        office = str(row[1]).strip() if len(row) > 1 else ""
        if office and "衛生所" in office:
            current_office = office

        person = str(row[2]).strip() if len(row) > 2 else ""
        amount = safe_float(row[5]) if len(row) > 5 else 0
        note = str(row[7]).strip() if len(row) > 7 else ""

        if amount > 0 and current_office:
            records.append({
                "分頁": sheet_name,
                "衛生所": current_office,
                "姓名": person if person else "",
                "日期": "",
                "金額": amount,
                "事由": note,
                "類型": "加班費"
            })

    return records


def extract_type3_overtime(wb, sheet_idx):
    """
    解析加班費彙整表格式 (分頁 3)
    格式: 編號 | 鄉鎮別 | 金額
    """
    sh = wb.sheet_by_index(sheet_idx)
    sheet_name = wb.sheet_names()[sheet_idx]
    records = []

    for r in range(3, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        row_text = "".join(str(v) for v in row)
        if "合計" in row_text or "總計" in row_text:
            continue

        code = str(row[0]).strip() if len(row) > 0 else ""
        office = str(row[1]).strip() if len(row) > 1 else ""
        amount = safe_float(row[2]) if len(row) > 2 else 0

        if amount > 0 and office:
            records.append({
                "分頁": sheet_name,
                "衛生所": office,
                "姓名": "",
                "日期": "",
                "金額": amount,
                "事由": "covid-19疫苗接種業務加班費",
                "類型": "加班費"
            })

    return records


def extract_summary(wb):
    """解析彙總表 (分頁 0)"""
    sh = wb.sheet_by_index(0)
    records = []
    # 欄位名稱在 R1-R2
    headers = []
    for c in range(3, 16):
        h1 = str(sh.cell_value(1, c)).strip()
        h2 = str(sh.cell_value(2, c)).strip()
        headers.append(h2 if h2 else h1)

    for r in range(3, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        code = str(row[1]).strip()
        office = str(row[2]).strip()
        if not code or "總" in office:
            continue
        total = safe_float(row[16])
        records.append({
            "編號": code,
            "衛生所": office,
            "各項金額": {headers[i]: safe_float(row[3 + i]) for i in range(13)},
            "合計": total
        })

    return records, headers


def write_output(all_records, target_records, overtime_records, summary_data, summary_headers, person_subtotals):
    """產出 Excel 總表"""
    wb_out = Workbook()

    # 樣式
    header_font = Font(bold=True, size=12)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    subtotal_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    subtotal_font = Font(bold=True, size=11)

    def style_header(ws, row, cols):
        for c in range(1, cols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

    def style_data(ws, row, cols):
        for c in range(1, cols + 1):
            cell = ws.cell(row=row, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')

    def auto_width(ws):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    val = str(cell.value) if cell.value else ""
                    # 中文字大約佔2個字元寬
                    length = sum(2 if ord(c) > 127 else 1 for c in val)
                    max_len = max(max_len, length)
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    # ===== Sheet 1: 佳里區(Z20)出差明細 =====
    ws1 = wb_out.active
    ws1.title = "佳里區衛生所明細"
    ws1.merge_cells('A1:G1')
    title_cell = ws1.cell(row=1, column=1, value="佳里區衛生所(Z20) 差旅費及加班費明細")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center')

    cols1 = ["分頁來源", "類型", "姓名", "日期", "金額", "事由", "備註"]
    for c, h in enumerate(cols1, 1):
        ws1.cell(row=3, column=c, value=h)
    style_header(ws1, 3, len(cols1))

    row_num = 4
    subtotal = 0
    for rec in target_records:
        ws1.cell(row=row_num, column=1, value=rec["分頁"])
        ws1.cell(row=row_num, column=2, value=rec["類型"])
        ws1.cell(row=row_num, column=3, value=rec["姓名"])
        ws1.cell(row=row_num, column=4, value=rec["日期"])
        ws1.cell(row=row_num, column=5, value=rec["金額"])
        ws1.cell(row=row_num, column=5).number_format = '#,##0'
        ws1.cell(row=row_num, column=6, value=rec["事由"])
        style_data(ws1, row_num, len(cols1))
        subtotal += rec["金額"]
        row_num += 1

    # 小計
    ws1.cell(row=row_num, column=4, value="合計")
    ws1.cell(row=row_num, column=5, value=subtotal)
    ws1.cell(row=row_num, column=5).number_format = '#,##0'
    for c in range(1, len(cols1) + 1):
        cell = ws1.cell(row=row_num, column=c)
        cell.fill = subtotal_fill
        cell.font = subtotal_font
        cell.border = thin_border

    auto_width(ws1)

    # ===== Sheet 2: 全部明細(含個人小計) =====
    ws2 = wb_out.create_sheet("全部明細")
    ws2.merge_cells('A1:G1')
    title2 = ws2.cell(row=1, column=1, value="115年衛生所第3次差旅費 - 全部明細")
    title2.font = Font(bold=True, size=14)
    title2.alignment = Alignment(horizontal='center')

    cols2 = ["分頁來源", "衛生所", "類型", "姓名", "日期", "金額", "事由"]
    for c, h in enumerate(cols2, 1):
        ws2.cell(row=3, column=c, value=h)
    style_header(ws2, 3, len(cols2))

    row_num = 4
    for rec in all_records:
        ws2.cell(row=row_num, column=1, value=rec["分頁"])
        ws2.cell(row=row_num, column=2, value=rec["衛生所"])
        ws2.cell(row=row_num, column=3, value=rec["類型"])
        ws2.cell(row=row_num, column=4, value=rec["姓名"])
        ws2.cell(row=row_num, column=5, value=rec["日期"])
        ws2.cell(row=row_num, column=6, value=rec["金額"])
        ws2.cell(row=row_num, column=6).number_format = '#,##0'
        ws2.cell(row=row_num, column=7, value=rec["事由"])
        style_data(ws2, row_num, len(cols2))
        row_num += 1

    # 總計
    total_all = sum(r["金額"] for r in all_records)
    ws2.cell(row=row_num, column=5, value="總計")
    ws2.cell(row=row_num, column=6, value=total_all)
    ws2.cell(row=row_num, column=6).number_format = '#,##0'
    for c in range(1, len(cols2) + 1):
        cell = ws2.cell(row=row_num, column=c)
        cell.fill = subtotal_fill
        cell.font = subtotal_font
        cell.border = thin_border

    auto_width(ws2)

    # ===== Sheet 3: 個人小計 =====
    ws3 = wb_out.create_sheet("個人小計")
    ws3.merge_cells('A1:E1')
    title3 = ws3.cell(row=1, column=1, value="各人員費用小計")
    title3.font = Font(bold=True, size=14)
    title3.alignment = Alignment(horizontal='center')

    cols3 = ["衛生所", "姓名", "差旅費小計", "加班費小計", "總計"]
    for c, h in enumerate(cols3, 1):
        ws3.cell(row=3, column=c, value=h)
    style_header(ws3, 3, len(cols3))

    row_num = 4
    grand_travel = 0
    grand_overtime = 0
    for (office, person), amounts in sorted(person_subtotals.items(), key=lambda x: x[0]):
        if not person:
            continue
        travel_amt = amounts.get("差旅費", 0)
        overtime_amt = amounts.get("加班費", 0)
        total = travel_amt + overtime_amt
        ws3.cell(row=row_num, column=1, value=office)
        ws3.cell(row=row_num, column=2, value=person)
        ws3.cell(row=row_num, column=3, value=travel_amt)
        ws3.cell(row=row_num, column=3).number_format = '#,##0'
        ws3.cell(row=row_num, column=4, value=overtime_amt)
        ws3.cell(row=row_num, column=4).number_format = '#,##0'
        ws3.cell(row=row_num, column=5, value=total)
        ws3.cell(row=row_num, column=5).number_format = '#,##0'
        style_data(ws3, row_num, len(cols3))
        grand_travel += travel_amt
        grand_overtime += overtime_amt
        row_num += 1

    # 總計
    ws3.cell(row=row_num, column=2, value="總計")
    ws3.cell(row=row_num, column=3, value=grand_travel)
    ws3.cell(row=row_num, column=3).number_format = '#,##0'
    ws3.cell(row=row_num, column=4, value=grand_overtime)
    ws3.cell(row=row_num, column=4).number_format = '#,##0'
    ws3.cell(row=row_num, column=5, value=grand_travel + grand_overtime)
    ws3.cell(row=row_num, column=5).number_format = '#,##0'
    for c in range(1, len(cols3) + 1):
        cell = ws3.cell(row=row_num, column=c)
        cell.fill = subtotal_fill
        cell.font = subtotal_font
        cell.border = thin_border

    auto_width(ws3)

    # ===== Sheet 4: 加班費彙整 =====
    ws4 = wb_out.create_sheet("加班費彙整")
    ws4.merge_cells('A1:F1')
    title4 = ws4.cell(row=1, column=1, value="全機關加班費彙整")
    title4.font = Font(bold=True, size=14)
    title4.alignment = Alignment(horizontal='center')

    cols4 = ["分頁來源", "衛生所", "姓名", "金額", "事由", "備註"]
    for c, h in enumerate(cols4, 1):
        ws4.cell(row=3, column=c, value=h)
    style_header(ws4, 3, len(cols4))

    row_num = 4
    overtime_total = 0
    for rec in overtime_records:
        ws4.cell(row=row_num, column=1, value=rec["分頁"])
        ws4.cell(row=row_num, column=2, value=rec["衛生所"])
        ws4.cell(row=row_num, column=3, value=rec["姓名"])
        ws4.cell(row=row_num, column=4, value=rec["金額"])
        ws4.cell(row=row_num, column=4).number_format = '#,##0'
        ws4.cell(row=row_num, column=5, value=rec["事由"])
        style_data(ws4, row_num, len(cols4))
        overtime_total += rec["金額"]
        row_num += 1

    # 總計
    ws4.cell(row=row_num, column=3, value="總計")
    ws4.cell(row=row_num, column=4, value=overtime_total)
    ws4.cell(row=row_num, column=4).number_format = '#,##0'
    for c in range(1, len(cols4) + 1):
        cell = ws4.cell(row=row_num, column=c)
        cell.fill = subtotal_fill
        cell.font = subtotal_font
        cell.border = thin_border

    auto_width(ws4)

    # ===== Sheet 5: 彙總表(原始) =====
    ws5 = wb_out.create_sheet("彙總表")
    ws5.merge_cells('A1:P1')
    title5 = ws5.cell(row=1, column=1, value="115年衛生所第3次差旅費等匯款明細")
    title5.font = Font(bold=True, size=14)
    title5.alignment = Alignment(horizontal='center')

    summary_cols = ["編號", "衛生所"] + summary_headers + ["合計"]
    for c, h in enumerate(summary_cols, 1):
        ws5.cell(row=3, column=c, value=h)
    style_header(ws5, 3, len(summary_cols))

    row_num = 4
    for rec in summary_data:
        ws5.cell(row=row_num, column=1, value=rec["編號"])
        ws5.cell(row=row_num, column=2, value=rec["衛生所"])
        for i, h in enumerate(summary_headers):
            val = rec["各項金額"].get(h, 0)
            ws5.cell(row=row_num, column=3 + i, value=val)
            ws5.cell(row=row_num, column=3 + i).number_format = '#,##0'
        ws5.cell(row=row_num, column=3 + len(summary_headers), value=rec["合計"])
        ws5.cell(row=row_num, column=3 + len(summary_headers)).number_format = '#,##0'

        # 標記佳里區
        if is_target(rec["衛生所"]) or is_target(rec["編號"]):
            highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for c in range(1, len(summary_cols) + 1):
                ws5.cell(row=row_num, column=c).fill = highlight

        style_data(ws5, row_num, len(summary_cols))
        row_num += 1

    auto_width(ws5)

    wb_out.save(OUTPUT_FILE)
    print(f"\n✅ 已產出: {OUTPUT_FILE}")


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    print("=" * 60)
    print("差旅明細整理程式")
    print("=" * 60)

    wb = open_workbook(INPUT_FILE)
    all_records = []

    # 解析各分頁
    print("\n📄 解析各分頁...")

    # 分頁 1, 2: 加班費明細
    for idx in [1, 2]:
        recs = extract_type1_overtime(wb, idx)
        all_records.extend(recs)
        print(f"  [{idx}] {wb.sheet_names()[idx]}: {len(recs)} 筆")

    # 分頁 3: 加班費彙整
    recs = extract_type3_overtime(wb, 3)
    all_records.extend(recs)
    print(f"  [3] {wb.sheet_names()[3]}: {len(recs)} 筆")

    # 分頁 4: 愛滋差旅
    recs = extract_type4_detail(wb, 4)
    all_records.extend(recs)
    print(f"  [4] {wb.sheet_names()[4]}: {len(recs)} 筆")

    # 分頁 5, 6: 登革熱交通
    for idx in [5, 6]:
        recs = extract_type5_detail(wb, idx)
        all_records.extend(recs)
        print(f"  [{idx}] {wb.sheet_names()[idx]}: {len(recs)} 筆")

    # 分頁 7: 低碳訓練
    recs = extract_type7_detail(wb, 7)
    all_records.extend(recs)
    print(f"  [7] {wb.sheet_names()[7]}: {len(recs)} 筆")

    # 分頁 8~13: 旅費明細清單
    for idx in range(8, 14):
        recs = extract_type8_detail(wb, idx)
        all_records.extend(recs)
        print(f"  [{idx}] {wb.sheet_names()[idx]}: {len(recs)} 筆")

    print(f"\n📊 共解析 {len(all_records)} 筆明細")

    # 篩選佳里區/Z20
    target_records = [r for r in all_records if is_target(r["衛生所"])]
    print(f"🎯 佳里區衛生所(Z20)相關: {len(target_records)} 筆")

    # 篩選加班費
    overtime_records = [r for r in all_records if r["類型"] == "加班費"]
    print(f"⏰ 全機關加班費: {len(overtime_records)} 筆")

    # 個人小計
    person_subtotals = {}
    for rec in all_records:
        key = (rec["衛生所"], rec["姓名"])
        if key not in person_subtotals:
            person_subtotals[key] = {"差旅費": 0, "加班費": 0}
        person_subtotals[key][rec["類型"]] += rec["金額"]

    print(f"👤 不重複人員: {len([k for k in person_subtotals if k[1]])} 人")

    # 解析彙總表
    summary_data, summary_headers = extract_summary(wb)

    # 產出 Excel
    print("\n📝 產出 Excel 總表...")
    write_output(all_records, target_records, overtime_records, summary_data, summary_headers, person_subtotals)

    # 顯示佳里區摘要
    print("\n" + "=" * 60)
    print("佳里區衛生所(Z20) 摘要:")
    print("=" * 60)
    target_total = sum(r["金額"] for r in target_records)
    for rec in target_records:
        print(f"  {rec['分頁']:<30} {rec['姓名']:<8} {rec['金額']:>8,.0f}  {rec['事由']}")
    print(f"  {'':─<60}")
    print(f"  {'合計':<38} {target_total:>8,.0f}")


if __name__ == "__main__":
    main()
