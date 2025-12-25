import os
import re
import pandas as pd
import uuid
import logging
from flask import Flask, request, render_template_string, jsonify, send_from_directory
from flask_httpauth import HTTPBasicAuth
from urllib.parse import quote

app = Flask(__name__)
auth = HTTPBasicAuth()

# 环境变量读取账号密码
USER_DATA = {
    os.getenv("WEB_USER", "admin"): os.getenv("WEB_PWD", "admin")
}

@auth.verify_password
def verify(username, password):
    if username in USER_DATA and USER_DATA.get(username) == password:
        return username

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

log_stream = []
class WebLogHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        log_stream.append(log_entry)
        if len(log_stream) > 50: log_stream.pop(0)

logger = logging.getLogger('web_logger')
logger.setLevel(logging.INFO)
handler = WebLogHandler()
handler.setFormatter(logging.Formatter('%H:%M:%S - %(levelname)s - %(message)s'))
logger.addHandler(handler)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <title>进销项智能比对系统</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8fafc; padding: 30px 0; }
        #logConsole { background: #1e293b; color: #38bdf8; font-family: monospace; height: 180px; overflow-y: auto; padding: 15px; border-radius: 8px; font-size: 0.85rem; }
    </style>
</head>
<body>
<div class="container" style="max-width: 900px;">
    <div class="card p-4 shadow-sm">
        <h4 class="mb-4">📊 进销项比对系统</h4>
        <form id="mainForm">
            <div class="row g-3">
                <div class="col-md-6"><label class="form-label">进项文件</label><input type="file" class="form-control" id="file_in" name="file_in" required onchange="parseFile('in')"></div>
                <div class="col-md-6"><label class="form-label">销项文件</label><input type="file" class="form-control" id="file_out" name="file_out" required onchange="parseFile('out')"></div>
            </div>
            <div id="mappingArea" class="mt-4" style="display:none;">
                <table class="table table-sm border">
                    <thead class="table-light"><tr><th>比对项</th><th>进项列</th><th>销项列</th></tr></thead>
                    <tbody>
                        <tr><td>名称列</td><td><select class="form-select" name="m_in_n" id="sel_in_n"></select></td><td><select class="form-select" name="m_out_n" id="sel_out_n"></select></td></tr>
                        <tr><td>数量列</td><td><select class="form-select" name="m_in_q" id="sel_in_q"></select></td><td><select class="form-select" name="m_out_q" id="sel_out_q"></select></td></tr>
                        <tr><td>金额列</td><td><select class="form-select" name="m_in_v" id="sel_in_v"></select></td><td><select class="form-select" name="m_out_v" id="sel_out_v"></select></td></tr>
                    </tbody>
                </table>
                <button type="submit" class="btn btn-primary w-100">执行比对</button>
            </div>
        </form>
        <div class="mt-4"><div id="logConsole"></div></div>
        <div id="resultArea" class="mt-3 text-center" style="display:none;"><a id="downloadBtn" href="#" class="btn btn-success">下载报告</a></div>
    </div>
</div>
<script>
async function parseFile(t) {
    const f = document.getElementById('file_'+t).files[0];
    const fd = new FormData(); fd.append('file', f);
    const res = await fetch('/get_headers', {method:'POST', body:fd});
    const data = await res.json();
    const sels = t==='in'?['sel_in_n','sel_in_q','sel_in_v']:['sel_out_n','sel_out_q','sel_out_v'];
    sels.forEach(id=>{ const el=document.getElementById(id); el.innerHTML=''; data.columns.forEach(c=>el.add(new Option(c,c))); });
    if(document.getElementById('sel_in_n').options.length && document.getElementById('sel_out_n').options.length)
        document.getElementById('mappingArea').style.display='block';
}
document.getElementById('mainForm').onsubmit = async (e) => {
    e.preventDefault();
    const res = await fetch('/process', {method:'POST', body:new FormData(e.target)});
    const data = await res.json();
    if(data.success) { document.getElementById('resultArea').style.display='block'; document.getElementById('downloadBtn').href='/download/'+data.filename; }
};
setInterval(async () => {
    const res = await fetch('/get_logs'); const logs = await res.json();
    const con = document.getElementById('logConsole');
    logs.forEach(l=>{ const d=document.createElement('div'); d.textContent=l; con.appendChild(d); });
    if(logs.length) con.scrollTop = con.scrollHeight;
}, 1000);
</script>
</body>
</html>
"""

def clean_name(t): return re.sub(r'\*.*?\*', '', str(t)).strip() if pd.notna(t) else ""

@app.route('/')
@auth.login_required
def index(): return render_template_string(HTML_TEMPLATE)

@app.route('/get_logs')
@auth.login_required
def get_logs():
    global log_stream
    l=list(log_stream); log_stream.clear(); return jsonify(l)

@app.route('/get_headers', methods=['POST'])
@auth.login_required
def get_headers():
    f = request.files['file']
    return jsonify({"columns": pd.read_excel(f, nrows=1).columns.tolist()})

@app.route('/process', methods=['POST'])
@auth.login_required
def process():
    try:
        logger.info("开始处理任务...")
        m = request.form
        df_in = pd.read_excel(request.files['file_in'])
        df_out = pd.read_excel(request.files['file_out'])
        df_in['__k'] = df_in[m['m_in_n']].apply(clean_name)
        df_out['__k'] = df_out[m['m_out_n']].apply(clean_name)
        in_agg = df_in.groupby('__k')[[m['m_in_q'], m['m_in_v']]].sum().reset_index()
        out_agg = df_out.groupby('__k')[[m['m_out_q'], m['m_out_v']]].sum().reset_index()
        in_agg.columns = ['关联名称', '进项_数量', '进项_金额']
        out_agg.columns = ['关联名称', '销项_数量', '销项_金额']
        merged = pd.merge(in_agg, out_agg, on='关联名称', how='outer').fillna(0)
        merged['数量差异'] = merged['销项_数量'] - merged['进项_数量']
        merged['金额差异'] = merged['销项_金额'] - merged['进项_金额']
        fname = f"res_{uuid.uuid4().hex}.xlsx"
        merged.to_excel(os.path.join(RESULT_FOLDER, fname), index=False)
        logger.info("处理完成。")
        return jsonify({"success":True, "filename":fname})
    except Exception as e:
        logger.error(str(e)); return jsonify({"success":False})

@app.route('/download/<filename>')
@auth.login_required
def download(filename):
    res = send_from_directory(RESULT_FOLDER, filename)
    res.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote('比对报告.xlsx')}"
    return res

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
