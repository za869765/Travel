"""
差旅明細整理 - Web 介面
上傳單一 XLS 檔案，自動處理並產出 Excel 總表供下載
不依賴 input/ 資料夾，上傳即處理
"""
import os
import sys
import io
import json
import cgi
import tempfile
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from travel import (
    open_workbook, process_file, extract_summary, is_target,
    Workbook, write_sheet_target, write_sheet_all, write_sheet_person,
    write_sheet_overtime, write_sheet_summary
)

PORT = 8080

# 暫存最後產出的 Excel bytes 和結果
last_result = {"excel_bytes": None, "data": None, "filename": ""}

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>差旅明細整理系統</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; background: #f0f4f8; min-height: 100vh; }
        .container { max-width: 960px; margin: 0 auto; padding: 20px; }
        header { background: linear-gradient(135deg, #1e3a5f 0%, #4a90d9 100%); color: white; padding: 30px; border-radius: 12px; margin-bottom: 24px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.15); }
        header h1 { font-size: 28px; margin-bottom: 8px; }
        header p { opacity: 0.85; font-size: 15px; }
        .card { background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .card h2 { font-size: 18px; color: #1e3a5f; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
        .upload-zone { border: 2px dashed #b0c4de; border-radius: 10px; padding: 40px; text-align: center; cursor: pointer; transition: all 0.3s; background: #f8fafc; }
        .upload-zone:hover, .upload-zone.dragover { border-color: #4a90d9; background: #eef4fb; transform: scale(1.01); }
        .upload-zone .icon { font-size: 48px; margin-bottom: 12px; }
        .upload-zone p { color: #666; margin-bottom: 8px; }
        .upload-zone .hint { font-size: 13px; color: #999; }
        #fileInput { display: none; }
        .file-info { margin-top: 16px; display: flex; align-items: center; justify-content: space-between; padding: 12px 16px; background: #eef4fb; border-radius: 8px; border: 1px solid #d5e3f0; }
        .file-info .name { font-weight: 600; color: #1e3a5f; }
        .file-info .size { color: #888; font-size: 13px; }
        .file-info .remove { color: #e74c3c; cursor: pointer; font-size: 20px; padding: 0 8px; }
        .btn { display: inline-flex; align-items: center; gap: 8px; padding: 12px 28px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.3s; }
        .btn-primary { background: #4a90d9; color: white; }
        .btn-primary:hover { background: #357abd; }
        .btn-primary:disabled { background: #b0c4de; cursor: not-allowed; }
        .btn-success { background: #27ae60; color: white; }
        .btn-success:hover { background: #219a52; }
        .btn-group { display: flex; gap: 12px; margin-top: 16px; justify-content: center; }
        .progress-area { margin-top: 20px; display: none; }
        .progress-bar { height: 6px; background: #e8edf2; border-radius: 3px; overflow: hidden; }
        .progress-bar .fill { height: 100%; background: linear-gradient(90deg, #4a90d9, #27ae60); border-radius: 3px; transition: width 0.5s; }
        .log { margin-top: 12px; background: #1e2a3a; color: #a8d8a8; border-radius: 8px; padding: 16px; font-family: "Consolas", monospace; font-size: 13px; max-height: 250px; overflow-y: auto; line-height: 1.6; }
        .result-area { display: none; }
        .result-summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #f8fafc, #eef4fb); padding: 16px; border-radius: 10px; text-align: center; border: 1px solid #d5e3f0; }
        .stat-card .number { font-size: 26px; font-weight: 700; color: #1e3a5f; }
        .stat-card .label { font-size: 13px; color: #666; margin-top: 4px; }
        table { width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 12px; }
        th { background: #4a90d9; color: white; padding: 10px 8px; text-align: left; position: sticky; top: 0; }
        td { padding: 8px; border-bottom: 1px solid #eee; }
        tr:hover { background: #f0f7ff; }
        .total-row { background: #d9e2f3 !important; font-weight: 700; }
        .section-title { margin: 20px 0 8px; font-size: 16px; color: #1e3a5f; font-weight: 700; }
        .table-wrap { max-height: 400px; overflow-y: auto; border: 1px solid #e8edf2; border-radius: 8px; }
        .text-right { text-align: right; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>差旅明細整理系統</h1>
            <p>上傳單一 XLS 旅費彙整表，自動拆分姓名、整理佳里區(Z20)明細、加班費彙整、個人小計</p>
        </header>

        <div class="card">
            <h2>📤 上傳檔案（一次一份）</h2>
            <div class="upload-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <div class="icon">📁</div>
                <p>點擊或拖曳 XLS 檔案到此處</p>
                <div class="hint">支援 .xls 格式</div>
            </div>
            <input type="file" id="fileInput" accept=".xls">
            <div id="fileInfo" style="display:none"></div>
            <div class="btn-group">
                <button class="btn btn-primary" id="btnProcess" onclick="processFile()" disabled>
                    ⚙️ 開始處理
                </button>
            </div>
        </div>

        <div class="progress-area" id="progressArea">
            <div class="card">
                <h2>⏳ 處理中...</h2>
                <div class="progress-bar"><div class="fill" id="progressFill" style="width:0%"></div></div>
                <div class="log" id="logArea"></div>
            </div>
        </div>

        <div class="result-area" id="resultArea">
            <div class="card">
                <h2>📊 處理結果</h2>
                <div class="result-summary" id="resultSummary"></div>
                <div class="btn-group">
                    <button class="btn btn-success" onclick="window.location.href='/download'">
                        📥 下載 Excel 總表
                    </button>
                    <button class="btn btn-primary" onclick="resetPage()">
                        🔄 處理另一份
                    </button>
                </div>
            </div>

            <div class="card" id="targetCard" style="display:none">
                <h2>🎯 佳里區衛生所(Z20) 明細</h2>
                <div class="table-wrap" id="targetDetail"></div>
            </div>

            <div class="card" id="personCard" style="display:none">
                <h2>👤 個人小計</h2>
                <div class="table-wrap" id="personDetail"></div>
            </div>
        </div>
    </div>

    <script>
        let selectedFile = null;

        const dropZone = document.getElementById('dropZone');
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) setFile(e.dataTransfer.files[0]);
        });
        document.getElementById('fileInput').addEventListener('change', e => {
            if (e.target.files.length > 0) setFile(e.target.files[0]);
        });

        function setFile(file) {
            if (!file.name.endsWith('.xls')) { alert('請上傳 .xls 格式檔案'); return; }
            selectedFile = file;
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('fileInfo').className = 'file-info';
            document.getElementById('fileInfo').innerHTML = `
                <span class="name">📄 ${file.name}</span>
                <span class="size">${(file.size / 1024).toFixed(1)} KB</span>
                <span class="remove" onclick="clearFile()">✕</span>
            `;
            document.getElementById('btnProcess').disabled = false;
        }

        function clearFile() {
            selectedFile = null;
            document.getElementById('fileInfo').style.display = 'none';
            document.getElementById('btnProcess').disabled = true;
            document.getElementById('fileInput').value = '';
        }

        function resetPage() {
            clearFile();
            document.getElementById('progressArea').style.display = 'none';
            document.getElementById('resultArea').style.display = 'none';
        }

        async function processFile() {
            if (!selectedFile) return;
            const btn = document.getElementById('btnProcess');
            btn.disabled = true;
            document.getElementById('progressArea').style.display = 'block';
            document.getElementById('resultArea').style.display = 'none';

            const log = document.getElementById('logArea');
            const fill = document.getElementById('progressFill');
            log.innerHTML = '';
            function addLog(msg) { log.innerHTML += msg + '\n'; log.scrollTop = log.scrollHeight; }

            addLog('📤 上傳檔案中...');
            fill.style.width = '20%';

            const formData = new FormData();
            formData.append('file', selectedFile);

            try {
                const resp = await fetch('/process', { method: 'POST', body: formData });
                fill.style.width = '90%';
                const result = await resp.json();
                fill.style.width = '100%';

                if (result.success) {
                    for (const line of result.log) addLog(line);
                    addLog('');
                    addLog('✅ 處理完成！');
                    showResult(result);
                } else {
                    addLog('❌ ' + result.error);
                }
            } catch (err) {
                addLog('❌ 錯誤: ' + err.message);
            }
            btn.disabled = false;
        }

        function showResult(r) {
            document.getElementById('resultArea').style.display = 'block';

            document.getElementById('resultSummary').innerHTML = `
                <div class="stat-card"><div class="number">${r.total_records}</div><div class="label">全部明細筆數</div></div>
                <div class="stat-card"><div class="number">${r.target_count}</div><div class="label">佳里區(Z20)</div></div>
                <div class="stat-card"><div class="number">${Number(r.target_total).toLocaleString()}</div><div class="label">佳里區合計</div></div>
                <div class="stat-card"><div class="number">${r.overtime_count}</div><div class="label">加班費筆數</div></div>
                <div class="stat-card"><div class="number">${r.person_count}</div><div class="label">不重複人員</div></div>
            `;

            // 佳里區明細
            if (r.target_records && r.target_records.length > 0) {
                document.getElementById('targetCard').style.display = 'block';
                let rows = r.target_records.map(t => `
                    <tr><td>${t.sheet}</td><td>${t.type}</td><td>${t.person}</td>
                    <td class="text-right">${Number(t.amount).toLocaleString()}</td><td>${t.reason}</td></tr>
                `).join('');
                document.getElementById('targetDetail').innerHTML = `
                    <table><tr><th>分頁</th><th>類型</th><th>姓名</th><th>金額</th><th>事由</th></tr>
                    ${rows}
                    <tr class="total-row"><td colspan="3" class="text-right">合計</td>
                    <td class="text-right">${Number(r.target_total).toLocaleString()}</td><td></td></tr></table>
                `;
            }

            // 個人小計
            if (r.person_subtotals && r.person_subtotals.length > 0) {
                document.getElementById('personCard').style.display = 'block';
                let gt = {travel:0, overtime:0, other:0};
                let prows = r.person_subtotals.map(p => {
                    gt.travel += p.travel; gt.overtime += p.overtime; gt.other += p.other;
                    let total = p.travel + p.overtime + p.other;
                    return `<tr><td>${p.office}</td><td>${p.person}</td>
                        <td class="text-right">${p.travel ? Number(p.travel).toLocaleString() : '-'}</td>
                        <td class="text-right">${p.overtime ? Number(p.overtime).toLocaleString() : '-'}</td>
                        <td class="text-right">${p.other ? Number(p.other).toLocaleString() : '-'}</td>
                        <td class="text-right">${Number(total).toLocaleString()}</td></tr>`;
                }).join('');
                let grandTotal = gt.travel + gt.overtime + gt.other;
                document.getElementById('personDetail').innerHTML = `
                    <table><tr><th>衛生所</th><th>姓名</th><th>差旅費</th><th>加班費</th><th>其他</th><th>總計</th></tr>
                    ${prows}
                    <tr class="total-row"><td></td><td>總計</td>
                        <td class="text-right">${Number(gt.travel).toLocaleString()}</td>
                        <td class="text-right">${Number(gt.overtime).toLocaleString()}</td>
                        <td class="text-right">${Number(gt.other).toLocaleString()}</td>
                        <td class="text-right">${Number(grandTotal).toLocaleString()}</td></tr></table>
                `;
            }
        }
    </script>
</body>
</html>"""


class TravelHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def do_GET(self):
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode('utf-8'))
        elif self.path == '/download':
            if last_result["excel_bytes"]:
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                fname = last_result.get("filename", "差旅明細整理總表.xlsx")
                self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{fname}")
                self.end_headers()
                self.wfile.write(last_result["excel_bytes"])
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b'No file yet')
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/process':
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self._json_response(400, {"success": False, "error": "Invalid request"})
                return

            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': content_type}
            )

            file_item = form['file']
            if not file_item.filename:
                self._json_response(400, {"success": False, "error": "No file uploaded"})
                return

            # 寫入暫存檔
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
            tmp.write(file_item.file.read())
            tmp.close()

            try:
                result = run_process(tmp.name, file_item.filename)
                self._json_response(200, result)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._json_response(500, {"success": False, "error": str(e)})
            finally:
                os.unlink(tmp.name)
        else:
            self.send_response(404)
            self.end_headers()

    def _json_response(self, code, data):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))


def run_process(filepath, original_filename):
    """處理單一上傳的 XLS 檔案"""
    file_label = original_filename.replace('.xls', '')
    log_lines = []

    def log_fn(msg):
        log_lines.append(msg)

    log_lines.append(f"📂 處理檔案: {original_filename}")

    # 解析
    all_records, wb = process_file(filepath, file_label, log_fn=log_fn)

    # 彙總表
    all_summaries = {}
    try:
        summary_data, summary_headers = extract_summary(wb, file_label)
        if summary_data:
            all_summaries[file_label] = (summary_data, summary_headers)
    except:
        pass

    log_lines.append(f"\n📊 共解析 {len(all_records)} 筆明細（姓名已拆分）")

    # 篩選
    target_records = [r for r in all_records if is_target(r["衛生所"])]
    overtime_records = [r for r in all_records if r["類型"] == "加班費"]

    # 個人小計
    person_subtotals = {}
    for rec in all_records:
        key = (rec["衛生所"], rec["姓名"])
        if key not in person_subtotals:
            person_subtotals[key] = {}
        fee_type = rec["類型"]
        person_subtotals[key][fee_type] = person_subtotals[key].get(fee_type, 0) + rec["金額"]

    named_persons = {k: v for k, v in person_subtotals.items() if k[1]}
    log_lines.append(f"🎯 佳里區(Z20): {len(target_records)} 筆")
    log_lines.append(f"⏰ 加班費: {len(overtime_records)} 筆")
    log_lines.append(f"👤 不重複人員: {len(named_persons)} 人")

    # 產出 Excel 到 memory
    wb_out = Workbook()
    write_sheet_target(wb_out, target_records, f"佳里區衛生所(Z20) 明細 - {file_label[:30]}", is_first=True)
    write_sheet_all(wb_out, all_records, f"{file_label[:40]} - 全部明細")
    write_sheet_person(wb_out, person_subtotals)
    write_sheet_overtime(wb_out, overtime_records)
    write_sheet_summary(wb_out, all_summaries)

    buf = io.BytesIO()
    wb_out.save(buf)
    excel_bytes = buf.getvalue()

    # 存到全域供下載
    last_result["excel_bytes"] = excel_bytes
    last_result["filename"] = f"{file_label}_整理總表.xlsx"

    target_total = sum(r["金額"] for r in target_records)

    # 個人小計列表 (for JSON)
    person_list = []
    for (office, person), amounts in sorted(named_persons.items(), key=lambda x: x[0]):
        travel = amounts.get("差旅費", 0)
        overtime = amounts.get("加班費", 0)
        other = sum(v for k, v in amounts.items() if k not in ("差旅費", "加班費"))
        person_list.append({
            "office": office,
            "person": person,
            "travel": travel,
            "overtime": overtime,
            "other": other
        })

    return {
        "success": True,
        "log": log_lines,
        "total_records": len(all_records),
        "target_count": len(target_records),
        "target_total": target_total,
        "overtime_count": len(overtime_records),
        "person_count": len(named_persons),
        "target_records": [
            {
                "sheet": r["分頁"],
                "type": r["類型"],
                "person": r["姓名"],
                "amount": r["金額"],
                "reason": r["事由"][:40]
            }
            for r in target_records
        ],
        "person_subtotals": person_list
    }


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    server = HTTPServer(('0.0.0.0', PORT), TravelHandler)
    print(f"差旅明細整理系統已啟動")
    print(f"  http://localhost:{PORT}")
    print(f"  按 Ctrl+C 停止")
    threading.Timer(1, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n伺服器已停止")
        server.server_close()


if __name__ == "__main__":
    main()
