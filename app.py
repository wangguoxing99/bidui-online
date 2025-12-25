import os
import re
import pandas as pd
import uuid
import logging
from flask import Flask, request, render_template_string, jsonify, send_from_directory
from urllib.parse import quote

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# 配置日志记录器
log_stream = []
class WebLogHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        log_stream.append(log_entry)
        if len(log_stream) > 100: log_stream.pop(0) # 保留最近100条

logger = logging.getLogger('web_logger')
logger.setLevel(logging.INFO)
handler = WebLogHandler()
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S'))
logger.addHandler(handler)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <title>进销项智能比对系统 Pro</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f6; padding: 40px 0; font-family: 'Microsoft YaHei', sans-serif; }
        .card { border: none; border-radius: 15px; box-shadow: 0 8px 20px rgba(0,0,0,0.05); }
        .step-header { background: #eef2ff; padding: 10px 20px; border-radius: 8px; font-weight: bold; color: #4338ca; margin-bottom: 20px; }
        #logConsole { background: #1e1e1e; color: #d4d4d4; font-family: 'Consolas', monospace; font-size: 0.85rem; 
                      height: 200px; overflow-y: auto; padding: 15px; border-radius: 8px; border: 1px solid #333; }
        .log-info { color: #4ec9b0; }
        .log-error { color: #f48771; }
        .log-warn { color: #ce9178; }
    </style>
</head>
<body>
<div class="container" style="max-width: 1000px;">
    <div class="card p-4">
        <h3 class="text-center mb-4">📊 进销项智能比对系统</h3>
        
        <form id="mainForm">
            <div class="step-header">第一步：上传原始 Excel 文件</div>
            <div class="row g-4">
                <div class="col-md-6">
                    <div class="p-3 border rounded">
                        <label class="form-label text-primary fw-bold">进项文件 (Input)</label>
                        <input type="file" class="form-control" name="file_in" id="file_in" required onchange="parseFile('in')">
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="p-3 border rounded">
                        <label class="form-label text-success fw-bold">销项文件 (Output)</label>
                        <input type="file" class="form-control" name="file_out" id="file_out" required onchange="parseFile('out')">
                    </div>
                </div>
            </div>

            <div id="mappingArea" class="mt-5" style="display:none;">
                <div class="step-header">第二步：配置差异比对列</div>
                <div class="table-responsive">
                    <table class="table table-hover align-middle">
                        <thead class="table-light">
                            <tr><th>对比项</th><th>进项文件列 (IN)</th><th>销项文件列 (OUT)</th></tr>
                        </thead>
                        <tbody>
                            <tr><td><strong>对比名称 (Key)</strong></td>
                                <td><select class="form-select" name="map_in_name" id="sel_in_name"></select></td>
                                <td><select class="form-select" name="map_out_name" id="sel_out_name"></select></td>
                            </tr>
                            <tr><td><strong>核算数量</strong></td>
                                <td><select class="form-select" name="map_in_qty" id="sel_in_qty"></select></td>
                                <td><select class="form-select" name="map_out_qty" id="sel_out_qty"></select></td>
                            </tr>
                            <tr><td><strong>核算金额</strong></td>
                                <td><select class="form-select" name="map_in_val" id="sel_in_val"></select></td>
                                <td><select class="form-select" name="map_out_val" id="sel_out_val"></select></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <button type="submit" class="btn btn-primary w-100 py-3 mt-3 fw-bold">执行比对任务</button>
            </div>
        </form>

        <div class="mt-5">
            <div class="d-flex justify-content-between align-items-center mb-2">
                <span class="fw-bold text-secondary">系统运行日志</span>
                <button class="btn btn-sm btn-link text-decoration-none" onclick="document.getElementById('logConsole').innerHTML=''">清空日志</button>
            </div>
            <div id="logConsole"></div>
        </div>

        <div id="resultArea" class="mt-4 text-center" style="display:none;">
            <div class="alert alert-success d-inline-block px-5">✅ 任务处理成功！请点击下方按钮下载报告。</div>
            <br>
            <a id="downloadBtn" href="#" class="btn btn-success btn-lg px-5 shadow-sm">下载差异比对报告.xlsx</a>
        </div>
    </div>
</div>

<script>
// 日志轮询
setInterval(async () => {
    const res = await fetch('/get_logs');
    const logs = await res.json();
    const consoleEl = document.getElementById('logConsole');
    if (logs.length > 0) {
        logs.forEach(log => {
            const p = document.createElement('div');
            if (log.includes('ERROR')) p.className = 'log-error';
            else if (log.includes('WARNING')) p.className = 'log-warn';
            else p.className = 'log-info';
            p.textContent = log;
            consoleEl.appendChild(p);
        });
        consoleEl.scrollTop = consoleEl.scrollHeight;
    }
}, 1000);

async function parseFile(type) {
    const fileInput = document.getElementById('file_' + type);
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    
    try {
        const response = await fetch('/get_headers', { method: 'POST', body: formData });
        const data = await response.json();
        if (data.columns) {
            fillSelects(type, data.columns);
            checkReady();
        }
    } catch (e) { console.error("解析失败", e); }
}

function fillSelects(type, columns) {
    const ids = type === 'in' ? ['sel_in_name', 'sel_in_qty', 'sel_in_val'] : ['sel_out_name', 'sel_out_qty', 'sel_out_val'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        el.innerHTML = '<option value="">--选择列--</option>';
        columns.forEach(col => {
            const opt = document.createElement('option');
            opt.value = col; opt.textContent = col;
            el.appendChild(opt);
        });
    });
}

function checkReady() {
    if (document.getElementById('sel_in_name').options.length > 1 && 
        document.getElementById('sel_out_name').options.length > 1) {
        document.getElementById('mappingArea').style.display = 'block';
    }
}

document.getElementById('mainForm').onsubmit = async (e) => {
    e.preventDefault();
    document.getElementById('resultArea').style.display = 'none';
    const formData = new FormData(e.target);
    const response = await fetch('/process', { method: 'POST', body: formData });
    const data = await response.json();
    if (data.success) {
        document.getElementById('resultArea').style.display = 'block';
        document.getElementById('downloadBtn').href = '/download/' + data.filename;
    }
};
</script>
</body>
</html>
"""

def clean_name(text):
    if pd.isna(text): return ""
    return re.sub(r'\*.*?\*', '', str(text)).strip()

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/get_logs')
def get_logs():
    global log_stream
    logs = list(log_stream)
    log_stream.clear()
    return jsonify(logs)

@app.route('/get_headers', methods=['POST'])
def get_headers():
    f = request.files['file']
    logger.info(f"正在读取文件表头: {f.filename}")
    df = pd.read_excel(f, nrows=1)
    return jsonify({"columns": df.columns.tolist()})

@app.route('/process', methods=['POST'])
def process():
    try:
        f_in = request.files['file_in']
        f_out = request.files['file_out']
        m = request.form
        
        logger.info(">>> 开始比对任务...")
        logger.info(f"配置映射: 进项[{m['map_in_name']}] <-> 销项[{m['map_out_name']}]")

        df_in = pd.read_excel(f_in)
        df_out = pd.read_excel(f_out)
        logger.info(f"数据加载完成。进项: {len(df_in)} 行, 销项: {len(df_out)} 行")

        # 核心逻辑
        df_in['__key__'] = df_in[m['map_in_name']].apply(clean_name)
        df_out['__key__'] = df_out[m['map_out_name']].apply(clean_name)
        
        logger.info("正在执行名称清洗与分组聚合...")
        in_agg = df_in.groupby('__key__')[[m['map_in_qty'], m['map_in_val']]].sum().reset_index()
        out_agg = df_out.groupby('__key__')[[m['map_out_qty'], m['map_out_val']]].sum().reset_index()

        in_agg.columns = ['关联名称', '进项_数量', '进项_金额']
        out_agg.columns = ['关联名称', '销项_数量', '销项_金额']

        logger.info("正在合并并计算差异...")
        merged = pd.merge(in_agg, out_agg, on='关联名称', how='outer').fillna(0)
        merged['数量差异(销-进)'] = merged['销项_数量'] - merged['进项_数量']
        merged['金额差异(销-进)'] = merged['销项_金额'] - merged['进项_金额']

        res_name = f"result_{uuid.uuid4().hex}.xlsx"
        merged.to_excel(os.path.join(RESULT_FOLDER, res_name), index=False)
        
        logger.info(f"SUCCESS: 比对完成。输出结果: {len(merged)} 条记录")
        return jsonify({"success": True, "filename": res_name})
    except Exception as e:
        logger.error(f"FAILED: 处理过程中发生错误: {str(e)}")
        return jsonify({"success": False, "message": str(e)})

@app.route('/download/<filename>')
def download(filename):
    display_name = "进销项比对报告.xlsx"
    response = send_from_directory(RESULT_FOLDER, filename)
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(display_name)}"
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
