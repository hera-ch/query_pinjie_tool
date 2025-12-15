import os
import io
import pandas as pd
from flask import Flask, render_template_string, request, jsonify, send_file
import webbrowser
from threading import Timer

# ================= 配置 =================
PORT = 5001
# =======================================

app = Flask(__name__)

# ================= 前端 HTML =================
html_template = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AE 链接生成器 (端口5001)</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <style>
        /* 让页面内容撑满高度，保证水印在底部 */
        html, body { height: 100%; margin: 0; display: flex; flex-direction: column; }
        body { background: #f4f7f6; font-family: sans-serif; }

        .main-content { flex: 1; padding: 30px 0; } /* 内容区自动伸缩 */

        .card-custom { border: none; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.05); background: white; padding: 25px; margin-bottom: 20px; }
        .btn-primary { background-color: #764ba2; border-color: #764ba2; border-radius: 50px; padding: 10px 30px; }
        .btn-primary:hover { background-color: #5d3b82; border-color: #5d3b82; }

        .result-table th { background: #f8f9fa; border: none; padding: 15px; color: #555; }
        .result-table td { border-bottom: 1px solid #eee; vertical-align: middle; }

        .code-block {
            background: #f1f3f5; padding: 8px 12px; border-radius: 6px;
            font-family: monospace; color: #d63384; font-size: 0.9rem;
            display: flex; justify-content: space-between; align-items: center;
            word-break: break-all;
        }
        .copy-btn {
            font-size: 12px; border: 1px solid #ddd; background: white;
            padding: 2px 8px; border-radius: 4px; cursor: pointer; color: #333; margin-left: 10px; flex-shrink: 0;
        }
        .copy-btn:hover { background: #eee; }
        .copy-btn.copied { background: #28a745; color: white; border-color: #28a745; }

        /* === 水印样式 === */
        .footer-watermark {
            text-align: center;
            padding: 20px 0;
            color: #000;        /* 暗色 */
            opacity: 0.15;      /* 极低透明度，达成“不明显”的效果 */
            font-size: 14px;
            font-weight: bold;
            font-family: 'Segoe UI', sans-serif;
            pointer-events: none; /* 鼠标穿透，不影响点击 */
            user-select: none;    /* 文字不可选中 */
        }
    </style>
</head>
<body>

    <!-- 主要内容区域 -->
    <div class="container main-content" style="max-width: 1000px;">
        <h2 class="text-center mb-4" style="color:#764ba2;">🔗 query链接配置工具</h2>

        <div class="row">
            <!-- 文本输入 -->
            <div class="col-md-6">
                <div class="card-custom">
                    <h5>✏️ 文本输入</h5>
                    <textarea id="textInput" class="form-control mb-3" rows="4" placeholder="例如：apple, banana"></textarea>
                    <button class="btn btn-primary w-100" onclick="startProcessText()">生成结果</button>
                </div>
            </div>

            <!-- Excel上传 -->
            <div class="col-md-6">
                <div class="card-custom">
                    <h5>📂 Excel 上传</h5>
                    <input class="form-control mb-3" type="file" id="fileInput" accept=".xlsx, .xls">
                    <button class="btn btn-primary w-100" onclick="startProcessExcel()">解析文件</button>
                </div>
            </div>
        </div>

        <!-- 结果区域 -->
        <div id="resultArea" style="display:none;" class="card-custom mt-3">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h5 class="m-0">生成结果 (<span id="count">0</span>)</h5>
                <div>
                    <button class="btn btn-sm btn-success me-2" onclick="exportToExcel()">📥 导出 Excel</button>
                    <button class="btn btn-sm btn-outline-danger" onclick="clearTable()">清空</button>
                </div>
            </div>
            <div class="table-responsive">
                <table class="table result-table">
                    <thead>
                        <tr>
                            <th width="15%">原始 Query</th>
                            <th width="40%">Keywords 格式</th>
                            <th width="45%">DeepLink 格式</th>
                        </tr>
                    </thead>
                    <tbody id="tbody"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- 底部水印 -->
    <div class="footer-watermark">@轻梦</div>

    <script>
        let globalDataList = [];

        // 1. 处理文本点击
        function startProcessText() {
            const raw = document.getElementById('textInput').value;
            if (!raw || !raw.trim()) return alert("⚠️ 请先输入内容！");

            const list = raw.split(/[,，\\n]/).map(s => s.trim()).filter(s => s);
            if(list.length === 0) return;
            renderData(list);
        }

        // 2. 处理Excel点击
        async function startProcessExcel() {
            const fileInput = document.getElementById('fileInput');
            if (!fileInput.files.length) return alert("⚠️ 请先选择 Excel 文件");

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            try {
                const btn = document.querySelector('button[onclick="startProcessExcel()"]');
                const oldText = btn.innerText;
                btn.innerText = "处理中..."; btn.disabled = true;

                const res = await fetch('/upload_excel', { method: 'POST', body: formData });
                const data = await res.json();

                if(data.error) alert("❌ 错误: " + data.error);
                else renderData(data);

                btn.innerText = oldText; btn.disabled = false;
                fileInput.value = ''; 
            } catch(e) {
                alert("请求失败"); console.error(e);
            }
        }

        // 3. 渲染表格
        function renderData(list) {
            const tbody = document.getElementById('tbody');
            const resultArea = document.getElementById('resultArea');

            list.forEach(q => {
                const fmt1 = `q=${q}&show_query=${q}`;
                const fmt2 = `aecmd://list?q=${q}&osf=index&show_query%02${q}`;

                globalDataList.push({
                    "原始Query": q,
                    "Keywords格式": fmt1,
                    "DeepLink格式": fmt2
                });

                const row = `
                    <tr>
                        <td><strong>${q}</strong></td>
                        <td>
                            <div class="code-block">
                                <span>${fmt1}</span>
                                <button class="copy-btn" onclick="doCopy(this, '${fmt1}')">复制</button>
                            </div>
                        </td>
                        <td>
                            <div class="code-block">
                                <span>${fmt2}</span>
                                <button class="copy-btn" onclick="doCopy(this, '${fmt2}')">复制</button>
                            </div>
                        </td>
                    </tr>
                `;
                tbody.insertAdjacentHTML('afterbegin', row);
            });

            resultArea.style.display = 'block';
            document.getElementById('count').innerText = globalDataList.length;
        }

        // 4. 导出 Excel
        async function exportToExcel() {
            if (globalDataList.length === 0) return alert("没有数据可以导出");
            try {
                const res = await fetch('/export_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(globalDataList)
                });
                if (res.ok) {
                    const blob = await res.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = "链接生成结果.xlsx";
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                } else {
                    alert("导出失败");
                }
            } catch (e) {
                console.error(e);
            }
        }

        function doCopy(btn, text) {
            const temp = document.createElement("textarea");
            temp.value = text;
            document.body.appendChild(temp);
            temp.select();
            document.execCommand("copy");
            document.body.removeChild(temp);
            const oldText = btn.innerText;
            btn.innerText = "OK"; btn.classList.add('copied');
            setTimeout(() => { btn.innerText = oldText; btn.classList.remove('copied'); }, 1000);
        }

        function clearTable() {
            document.getElementById('tbody').innerHTML = '';
            document.getElementById('resultArea').style.display = 'none';
            globalDataList = [];
        }
    </script>
</body>
</html>
"""


# ================= 后端逻辑 =================

@app.route('/')
def index():
    return render_template_string(html_template)


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files: return jsonify({"error": "No file"}), 400
    try:
        file = request.files['file']
        df = pd.read_excel(file)
        data = df.iloc[:, 0].dropna().astype(str).tolist()
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        data_list = request.json
        df = pd.DataFrame(data_list)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="result.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    # 注意 host='0.0.0.0'
    app.run(host='0.0.0.0', port=5001, debug=False)
