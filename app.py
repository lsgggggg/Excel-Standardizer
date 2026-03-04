#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Standardizer Web v4.0
============================
Flask Web 服务器 + 嵌入式前端界面
用法: python app.py → 浏览器打开 http://localhost:5000
"""

import os, sys, json, copy, uuid, shutil, traceback
from datetime import datetime
from pathlib import Path

try:
    from flask import Flask, request, jsonify, send_file, Response
except ImportError:
    print("缺少 Flask，请运行: pip install flask")
    sys.exit(1)

# 导入核心标准化引擎
try:
    import excel_standardizer as es
except ImportError:
    print("错误: 找不到 excel_standardizer.py，请确保它与 app.py 在同一目录")
    sys.exit(1)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# 全局会话存储(本地单用户,无需数据库)
sessions = {}


def get_session(sid):
    if sid not in sessions:
        sessions[sid] = {
            'settings': copy.deepcopy(es.DEFAULT_SETTINGS),
            'filepath': None, 'wb': None, 'proposals': [],
            'col_types': {}, 'logger': None, 'sheets': [],
        }
    return sessions[sid]


# ============================================================
#  API 路由
# ============================================================

@app.route('/')
def index():
    return Response(HTML_TEMPLATE, mimetype='text/html')


@app.route('/api/session', methods=['POST'])
def create_session():
    sid = str(uuid.uuid4())[:8]
    get_session(sid)
    return jsonify({'session_id': sid})


@app.route('/api/settings/<sid>', methods=['GET'])
def get_settings(sid):
    s = get_session(sid)
    # 构建分类结构
    cats = [
        ("一、字符编码与不可见字符", "1."),
        ("二、大小写标准化", "2."),
        ("三、中文相关标准化", "3."),
        ("四、标点符号与特殊字符", "4."),
        ("五、数字与数值标准化", "5."),
        ("六、日期与时间标准化", "6."),
        ("七、公司/机构名称标准化", "7."),
        ("八、通用文本标准化", "8."),
        ("九、电话号码与证件号", "9."),
        ("十、Excel格式与样式", "10."),
        ("十一、数据结构与完整性", "11."),
        ("十二、地址与地理信息", "12."),
    ]
    result = []
    for cat_name, prefix in cats:
        items = []
        for k, v in s['settings'].items():
            if k.startswith(prefix):
                level = 'safe' if k in es.SAFETY_LEVELS.get('SAFE', []) else \
                        'moderate' if k in es.SAFETY_LEVELS.get('MODERATE', []) else 'dangerous'
                col_req = list(es.RULE_COLUMN_APPLICABILITY.get(k, []))
                items.append({
                    'key': k, 'enabled': v['enabled'], 'desc': v['desc'],
                    'level': level, 'col_types': col_req,
                    'option': v.get('option'), 'options': v.get('options'),
                })
        if items:
            result.append({'category': cat_name, 'rules': items})
    return jsonify(result)


@app.route('/api/settings/<sid>', methods=['POST'])
def update_settings(sid):
    s = get_session(sid)
    data = request.json
    if 'key' in data and 'enabled' in data:
        k = data['key']
        if k in s['settings']:
            s['settings'][k]['enabled'] = data['enabled']
            if 'option' in data and data['option']:
                s['settings'][k]['option'] = data['option']
    elif 'preset' in data:
        p = data['preset']
        if p == 'safe':
            for k in s['settings']:
                s['settings'][k]['enabled'] = k in es.SAFETY_LEVELS.get('SAFE', [])
            for k in ["1.01_fullwidth_to_halfwidth", "1.02_fullwidth_space",
                       "1.04_tab_to_space", "1.10_unicode_normalize",
                       "8.09_email_normalize", "11.01_header_clean",
                       "4.02_tilde_normalize", "4.05_ellipsis_normalize",
                       "5.06_negative_accounting"]:
                if k in s['settings']:
                    s['settings'][k]['enabled'] = True
        elif p == 'all_on':
            for k in s['settings']:
                s['settings'][k]['enabled'] = True
        elif p == 'all_off':
            for k in s['settings']:
                s['settings'][k]['enabled'] = False
        elif p == 'default':
            s['settings'] = copy.deepcopy(es.DEFAULT_SETTINGS)
    return jsonify({'ok': True})


@app.route('/api/settings/<sid>/batch', methods=['POST'])
def batch_settings(sid):
    s = get_session(sid)
    updates = request.json.get('updates', {})
    for k, v in updates.items():
        if k in s['settings']:
            if isinstance(v, bool):
                s['settings'][k]['enabled'] = v
            elif isinstance(v, dict):
                s['settings'][k].update(v)
    return jsonify({'ok': True})


@app.route('/api/upload/<sid>', methods=['POST'])
def upload_file(sid):
    s = get_session(sid)
    if 'file' not in request.files:
        return jsonify({'error': '未选择文件'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': '仅支持 .xlsx / .xlsm 格式'}), 400

    # 保存文件
    sess_dir = UPLOAD_DIR / sid
    sess_dir.mkdir(exist_ok=True)
    filepath = sess_dir / f.filename
    f.save(str(filepath))
    s['filepath'] = str(filepath)

    # 加载工作簿
    try:
        wb = es.load_workbook(str(filepath), data_only=False, keep_links=True)
        s['wb'] = wb
        warnings = es.pre_check(wb)
        sheets_info = []
        for name in wb.sheetnames:
            ws = wb[name]
            sheets_info.append({
                'name': name,
                'rows': ws.max_row or 0,
                'cols': ws.max_column or 0,
            })
        s['sheets'] = [si['name'] for si in sheets_info]
        return jsonify({
            'filename': f.filename,
            'sheets': sheets_info,
            'warnings': warnings,
        })
    except Exception as e:
        return jsonify({'error': f'读取文件失败: {str(e)}'}), 400


@app.route('/api/analyze/<sid>', methods=['POST'])
def analyze(sid):
    s = get_session(sid)
    if not s['wb']:
        return jsonify({'error': '请先上传文件'}), 400

    data = request.json or {}
    selected_sheets = data.get('sheets', s['sheets'])
    col_type_overrides = data.get('col_types', {})

    wb = s['wb']
    logger = es.ChangeLogger()
    s['logger'] = logger
    all_col_types = {}

    # 阶段1: 列类型检测
    for sn in selected_sheets:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        ct = es.ColumnTypeDetector.detect_all_columns(ws)
        # 应用用户覆盖
        if sn in col_type_overrides:
            for col, ctype in col_type_overrides[sn].items():
                ct[col] = ctype
        all_col_types[sn] = ct

    s['col_types'] = all_col_types

    # 阶段2: 计算变更提案
    all_proposals = []
    for sn in selected_sheets:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        es.process_header(ws, s['settings'], logger, sn)
        ct = all_col_types.get(sn, {})
        proposals = es.compute_proposals(ws, sn, s['settings'], ct, logger)
        all_proposals.extend(proposals)

    s['proposals'] = all_proposals

    # 构建返回数据
    finals = [p for p in all_proposals if p.rule_id == "_FINAL_"]
    details = [p for p in all_proposals if p.rule_id != "_FINAL_"]

    col_types_display = {}
    for sn, ct in all_col_types.items():
        col_types_display[sn] = {col: t for col, t in ct.items() if t != 'general_text'}

    # 按规则统计
    rule_stats = {}
    for p in details:
        rule_stats[p.rule_id] = rule_stats.get(p.rule_id, 0) + 1

    proposals_data = []
    for i, p in enumerate(finals):
        rules = [d.rule_id for d in details
                 if d.sheet == p.sheet and d.cell_ref == p.cell_ref]
        proposals_data.append({
            'idx': i, 'sheet': p.sheet, 'cell': p.cell_ref,
            'original': p.original[:200], 'proposed': p.proposed[:200],
            'rules': rules, 'accepted': True,
            'row': p.row, 'col': p.col,
        })

    return jsonify({
        'col_types': col_types_display,
        'total_changes': len(finals),
        'rule_stats': rule_stats,
        'proposals': proposals_data,
        'warnings': [{'sheet': w['sheet'], 'cell': w['cell'],
                       'message': w['message']} for w in logger.warnings],
        'skipped': len(logger.skipped),
    })


@app.route('/api/apply/<sid>', methods=['POST'])
def apply_changes(sid):
    s = get_session(sid)
    if not s['wb'] or not s['proposals']:
        return jsonify({'error': '请先分析文件'}), 400

    data = request.json or {}
    decisions = data.get('decisions', {})  # {idx: {accepted: bool, override: str|null}}

    finals = [p for p in s['proposals'] if p.rule_id == "_FINAL_"]

    # 应用用户决策
    for idx_str, dec in decisions.items():
        idx = int(idx_str)
        if 0 <= idx < len(finals):
            finals[idx].accepted = dec.get('accepted', True)
            override = dec.get('override')
            if override is not None and override != '':
                finals[idx].user_override = override

    # 重新加载工作簿(避免header已被改)
    wb = es.load_workbook(s['filepath'], data_only=False, keep_links=True)

    # 先处理表头
    logger = es.ChangeLogger()
    for sn in s['sheets']:
        if sn in wb.sheetnames:
            es.process_header(wb[sn], s['settings'], logger, sn)

    applied = es.apply_confirmed_changes(wb, s['proposals'])

    # 保存输出文件
    sess_dir = UPLOAD_DIR / sid
    orig = Path(s['filepath'])
    out_name = f"{orig.stem}_标准化{orig.suffix}"
    out_path = sess_dir / out_name
    wb.save(str(out_path))

    # 保存变更日志
    for p in s['proposals']:
        if p.rule_id == "_FINAL_" and (p.accepted or p.user_override):
            logger.log(p.sheet, p.cell_ref, p.original, p.final_value, "最终变更")

    log_name = f"{orig.stem}_变更日志.xlsx"
    log_path = sess_dir / log_name
    log_count = 0
    if logger.logs:
        log_count = logger.export_to_workbook(str(log_path))

    confirmed = len([p for p in finals if p.accepted or p.user_override])
    rejected = len(finals) - confirmed

    return jsonify({
        'applied': applied, 'confirmed': confirmed, 'rejected': rejected,
        'output_file': out_name, 'log_file': log_name if log_count > 0 else None,
        'log_count': log_count,
    })


@app.route('/api/download/<sid>/<filename>')
def download(sid, filename):
    filepath = UPLOAD_DIR / sid / filename
    if filepath.exists():
        return send_file(str(filepath.resolve()), as_attachment=True,
                         download_name=filename)
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/generate_test/<sid>', methods=['POST'])
def generate_test(sid):
    sess_dir = UPLOAD_DIR / sid
    sess_dir.mkdir(exist_ok=True)
    # 使用引擎的测试生成器
    old_cwd = os.getcwd()
    os.chdir(str(sess_dir))
    try:
        es.generate_comprehensive_test_file()
    finally:
        os.chdir(old_cwd)
    return jsonify({'filename': 'test_standardization_v3.xlsx'})


# ============================================================
#  新增 API: Excel 预览(返回指定sheet的数据用于网页展示)
# ============================================================

@app.route('/api/preview/<sid>', methods=['POST'])
def preview_excel(sid):
    """返回原始 Excel 文件的 sheet 数据, 用于在审核界面中预览"""
    s = get_session(sid)
    if not s['filepath']:
        return jsonify({'error': '无文件'}), 400

    data = request.json or {}
    sheet_name = data.get('sheet', '')
    target_row = data.get('row', 1)
    target_col = data.get('col', 1)

    # 始终从原始文件读取(保持原始数据)
    try:
        wb = es.load_workbook(s['filepath'], data_only=False, keep_links=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # 计算要返回的区域: 目标单元格周围的上下文
    context_rows = 8
    context_cols = 4

    row_start = max(1, target_row - context_rows)
    row_end = min(max_row, target_row + context_rows)
    col_start = max(1, target_col - context_cols)
    col_end = min(max_col, target_col + context_cols)

    # 始终包含第1行(表头)
    include_header = (row_start > 1)

    rows_data = []

    # 如果需要, 先加表头行
    if include_header:
        header_cells = []
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=1, column=c)
            val = cell.value
            if val is None:
                val = ''
            else:
                val = str(val)[:100]
            header_cells.append({
                'value': val,
                'col': c,
                'col_letter': es.get_column_letter(c),
                'is_target': False,
                'is_header': True,
            })
        rows_data.append({
            'row': 1,
            'cells': header_cells,
            'is_separator': False,
        })
        # 添加省略标记
        if row_start > 2:
            rows_data.append({'row': -1, 'cells': [], 'is_separator': True})

    # 数据行
    for r in range(row_start, row_end + 1):
        cells = []
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None:
                val = ''
            elif isinstance(val, (datetime,)):
                val = val.strftime('%Y-%m-%d %H:%M:%S')
            elif isinstance(val, bool):
                val = str(val)
            else:
                val = str(val)[:200]

            is_formula = isinstance(cell.value, str) and str(cell.value).startswith('=')

            cells.append({
                'value': val,
                'col': c,
                'col_letter': es.get_column_letter(c),
                'is_target': (r == target_row and c == target_col),
                'is_header': (r == 1),
                'is_formula': is_formula,
            })
        rows_data.append({
            'row': r,
            'cells': cells,
            'is_separator': False,
        })

    # 列头信息
    col_headers = []
    for c in range(col_start, col_end + 1):
        col_headers.append({
            'col': c,
            'letter': es.get_column_letter(c),
            'is_target_col': (c == target_col),
        })

    wb.close()

    return jsonify({
        'sheet': sheet_name,
        'sheets': list(wb.sheetnames) if hasattr(wb, 'sheetnames') else [sheet_name],
        'target_row': target_row,
        'target_col': target_col,
        'target_cell': f"{es.get_column_letter(target_col)}{target_row}",
        'col_headers': col_headers,
        'rows': rows_data,
        'total_rows': max_row,
        'total_cols': max_col,
    })


# ============================================================
#  嵌入式 HTML 前端 — v4.0 全新高端白色专业主题
# ============================================================

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel Standardizer v4.0</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
/* ═══════════════════════════════════════════════════
   ROOT & VARIABLES — v4.0 Premium White Theme
   ═══════════════════════════════════════════════════ */
:root {
  --font-main: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
  --font-mono: 'JetBrains Mono', 'SF Mono', 'Fira Code', monospace;
  --font-cn: 'Inter', -apple-system, 'PingFang SC', 'Microsoft YaHei', 'Noto Sans SC', sans-serif;

  --white: #ffffff;
  --gray-25: #fdfdfe;
  --gray-50: #f8f9fc;
  --gray-75: #f3f4f8;
  --gray-100: #eef0f4;
  --gray-150: #e4e7ed;
  --gray-200: #d8dce6;
  --gray-300: #b8bfcc;
  --gray-400: #8e96a6;
  --gray-500: #6b7280;
  --gray-600: #4b5563;
  --gray-700: #374151;
  --gray-800: #1f2937;
  --gray-900: #111827;

  --blue-50: #eff6ff;
  --blue-100: #dbeafe;
  --blue-200: #bfdbfe;
  --blue-300: #93c5fd;
  --blue-400: #60a5fa;
  --blue-500: #3b82f6;
  --blue-600: #2563eb;
  --blue-700: #1d4ed8;

  --green-50: #f0fdf4;
  --green-100: #dcfce7;
  --green-200: #bbf7d0;
  --green-500: #22c55e;
  --green-600: #16a34a;
  --green-700: #15803d;

  --amber-50: #fffbeb;
  --amber-100: #fef3c7;
  --amber-200: #fde68a;
  --amber-500: #f59e0b;
  --amber-600: #d97706;
  --amber-700: #b45309;

  --red-50: #fef2f2;
  --red-100: #fee2e2;
  --red-200: #fecaca;
  --red-400: #f87171;
  --red-500: #ef4444;
  --red-600: #dc2626;

  --violet-50: #f5f3ff;
  --violet-100: #ede9fe;
  --violet-500: #8b5cf6;
  --violet-600: #7c3aed;

  --bg: #f5f6fa;
  --surface: var(--white);
  --surface-raised: var(--white);
  --border: #e5e7eb;
  --border-light: #f0f1f4;
  --text: var(--gray-800);
  --text-secondary: var(--gray-500);
  --text-tertiary: var(--gray-400);
  --primary: #4f6ef7;
  --primary-hover: #3b5de5;
  --primary-light: #eef2ff;
  --primary-border: var(--blue-200);

  --radius-xs: 4px;
  --radius-sm: 6px;
  --radius: 10px;
  --radius-lg: 14px;
  --radius-xl: 20px;
  --shadow-xs: 0 1px 2px rgba(0,0,0,.03);
  --shadow-sm: 0 1px 3px rgba(0,0,0,.04), 0 1px 2px rgba(0,0,0,.03);
  --shadow-md: 0 4px 16px rgba(0,0,0,.05), 0 1px 4px rgba(0,0,0,.03);
  --shadow-lg: 0 8px 30px rgba(0,0,0,.07), 0 2px 8px rgba(0,0,0,.03);
  --shadow-xl: 0 20px 50px rgba(0,0,0,.08), 0 4px 12px rgba(0,0,0,.04);
  --shadow-card: 0 1px 3px rgba(0,0,0,.04), 0 1px 2px rgba(0,0,0,.02);
  --shadow-card-hover: 0 4px 12px rgba(0,0,0,.06), 0 1px 4px rgba(0,0,0,.03);
  --transition: .2s cubic-bezier(.4,0,.2,1);
  --transition-fast: .12s ease;
}

/* ═══ RESET & BASICS ═══ */
*, *::before, *::after { margin:0; padding:0; box-sizing:border-box; }
html { font-size: 14px; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; scroll-behavior: smooth; }
body {
  font-family: var(--font-cn);
  background: var(--bg);
  color: var(--text);
  line-height: 1.6;
  min-height: 100vh;
  overflow-x: hidden;
}

/* ═══ LAYOUT ═══ */
.app { display: flex; min-height: 100vh; }

/* ── SIDEBAR ── */
.sidebar {
  width: 256px; min-width: 256px;
  background: var(--white);
  border-right: 1px solid var(--border);
  display: flex; flex-direction: column;
  position: sticky; top: 0;
  max-height: 100vh; overflow-y: auto;
  z-index: 50;
  box-shadow: 1px 0 4px rgba(0,0,0,.02);
}
.sidebar-brand {
  padding: 28px 24px 24px;
  border-bottom: 1px solid var(--border-light);
}
.sidebar-brand h1 {
  font-size: 15px; font-weight: 700; letter-spacing: -.3px;
  color: var(--gray-800); display: flex; align-items: center; gap: 12px;
}
.sidebar-brand h1 .brand-icon {
  width: 36px; height: 36px; border-radius: 10px;
  background: linear-gradient(135deg, #4f6ef7, #7c5cf6);
  display: flex; align-items: center; justify-content: center;
  font-size: 17px; color: #fff; flex-shrink: 0;
  box-shadow: 0 2px 8px rgba(79,110,247,.25);
}
.sidebar-brand p {
  font-size: 11px; color: var(--text-tertiary); margin-top: 6px;
  font-weight: 400; padding-left: 48px; letter-spacing: .2px;
}

.nav { padding: 16px 12px; flex: 1; }
.nav-item {
  display: flex; align-items: center; gap: 12px;
  padding: 10px 14px; cursor: pointer;
  transition: var(--transition);
  font-size: 13.5px; font-weight: 500;
  color: var(--gray-500);
  border-radius: var(--radius);
  margin-bottom: 2px;
  position: relative;
  user-select: none;
}
.nav-item:hover {
  background: var(--gray-50);
  color: var(--gray-700);
}
.nav-item.active {
  background: var(--primary-light);
  color: var(--primary);
  font-weight: 600;
}
.nav-item.active::before {
  content: '';
  position: absolute; left: 0; top: 50%; transform: translateY(-50%);
  width: 3px; height: 18px; border-radius: 0 3px 3px 0;
  background: var(--primary);
}
.nav-item .icon {
  font-size: 16px; width: 22px; text-align: center;
  flex-shrink: 0; opacity: .85;
}
.nav-item .badge {
  margin-left: auto;
  background: var(--primary);
  color: #fff;
  font-size: 10.5px; font-weight: 700;
  padding: 1px 8px; border-radius: 10px;
  line-height: 18px; min-width: 26px; text-align: center;
}
.nav-item .badge-check {
  margin-left: auto;
  color: var(--green-500); font-size: 16px;
}

.sidebar-footer {
  padding: 18px 24px;
  border-top: 1px solid var(--border-light);
  font-size: 11.5px; color: var(--text-tertiary);
  line-height: 1.8;
  background: var(--gray-25);
}
.sidebar-footer strong { color: var(--gray-600); font-weight: 600; }

/* ── MAIN CONTENT ── */
.main {
  flex: 1; padding: 28px 36px; overflow-y: auto;
  min-width: 0;
}
@media (max-width: 1400px) { .main { padding: 24px 24px; } }

/* ═══ PAGE HEADER ═══ */
.page-header {
  margin-bottom: 24px;
}
.page-header h2 {
  font-size: 20px; font-weight: 700; color: var(--gray-900);
  letter-spacing: -.4px; display: flex; align-items: center; gap: 10px;
}
.page-header h2 .ph-icon {
  font-size: 22px; opacity: .8;
}
.page-header .ph-desc {
  font-size: 13px; color: var(--text-secondary); margin-top: 4px;
  font-weight: 400;
}

/* ═══ STEPS INDICATOR ═══ */
.steps {
  display: flex; gap: 0; margin-bottom: 24px;
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  overflow: hidden;
  box-shadow: var(--shadow-xs);
}
.step {
  flex: 1; text-align: center; padding: 12px 8px;
  font-size: 12.5px; font-weight: 600;
  color: var(--text-tertiary);
  border-right: 1px solid var(--border-light);
  transition: var(--transition);
  position: relative;
  letter-spacing: .2px;
}
.step:last-child { border-right: none; }
.step.active {
  color: var(--primary);
  background: var(--primary-light);
}
.step.done {
  color: var(--green-600);
  background: var(--green-50);
}
.step .num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 20px; height: 20px; border-radius: 50%;
  font-size: 10.5px; font-weight: 700;
  background: var(--gray-150); color: var(--gray-500);
  margin-right: 6px; vertical-align: middle;
}
.step.active .num { background: var(--primary); color: #fff; }
.step.done .num { background: var(--green-500); color: #fff; }

/* ═══ CARDS ═══ */
.card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  padding: 24px;
  margin-bottom: 20px;
  box-shadow: var(--shadow-card);
  transition: box-shadow var(--transition);
}
.card:hover { box-shadow: var(--shadow-card-hover); }
.card-header {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 18px; gap: 16px;
}
.card-title {
  font-size: 15px; font-weight: 700; color: var(--gray-800);
  display: flex; align-items: center; gap: 8px;
  letter-spacing: -.2px;
}
.card-title .title-icon { font-size: 17px; opacity: .8; }
.card-subtitle {
  font-size: 12.5px; color: var(--text-secondary); margin-top: 3px;
}

/* ═══ BUTTONS ═══ */
.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 8px 18px; border-radius: var(--radius);
  border: none; font-family: var(--font-cn);
  font-size: 13px; font-weight: 600; cursor: pointer;
  transition: var(--transition);
  text-decoration: none;
  line-height: 1.4;
  letter-spacing: .1px;
}
.btn:active { transform: scale(.97); }
.btn-primary {
  background: var(--primary); color: #fff;
  box-shadow: 0 1px 4px rgba(79,110,247,.2), 0 1px 2px rgba(79,110,247,.1);
}
.btn-primary:hover { background: var(--primary-hover); box-shadow: 0 3px 12px rgba(79,110,247,.25); }
.btn-success {
  background: var(--green-600); color: #fff;
  box-shadow: 0 1px 3px rgba(22,163,74,.2);
}
.btn-success:hover { background: var(--green-700); }
.btn-danger { background: var(--red-500); color: #fff; }
.btn-danger:hover { background: var(--red-600); }
.btn-outline {
  background: var(--white);
  border: 1px solid var(--gray-200);
  color: var(--gray-600);
}
.btn-outline:hover {
  border-color: var(--primary);
  color: var(--primary);
  background: var(--primary-light);
}
.btn-sm { padding: 6px 14px; font-size: 12px; border-radius: var(--radius-sm); }
.btn-group { display: flex; gap: 8px; flex-wrap: wrap; }
.btn-icon {
  width: 34px; height: 34px; padding: 0;
  display: inline-flex; align-items: center; justify-content: center;
  border-radius: var(--radius); border: 1px solid var(--border);
  background: var(--white); color: var(--gray-500); cursor: pointer;
  transition: var(--transition); font-size: 15px;
}
.btn-icon:hover { border-color: var(--primary); color: var(--primary); background: var(--primary-light); }

/* ═══ UPLOAD ZONE ═══ */
.upload-zone {
  border: 2px dashed var(--gray-200);
  border-radius: var(--radius-lg);
  padding: 52px 40px; text-align: center;
  cursor: pointer; transition: var(--transition);
  background: var(--gray-25);
  position: relative;
}
.upload-zone:hover, .upload-zone.dragover {
  border-color: var(--primary);
  background: var(--primary-light);
  box-shadow: 0 0 0 4px rgba(79,110,247,.06);
}
.upload-zone .upload-icon {
  width: 60px; height: 60px; border-radius: 16px;
  background: linear-gradient(135deg, var(--blue-100), var(--violet-100));
  display: flex; align-items: center; justify-content: center;
  font-size: 28px; margin: 0 auto 18px;
  box-shadow: 0 2px 8px rgba(79,110,247,.1);
}
.upload-zone p {
  color: var(--gray-600); font-size: 14.5px; font-weight: 600;
  letter-spacing: -.1px;
}
.upload-zone .hint {
  font-size: 12px; color: var(--text-tertiary); margin-top: 8px;
  font-weight: 400;
}

/* ═══ FILE INFO ═══ */
.file-info {
  display: flex; align-items: center; gap: 14px;
  padding: 14px 18px;
  background: var(--green-50);
  border: 1px solid var(--green-100);
  border-radius: var(--radius);
  margin-bottom: 14px;
}
.file-info .fi-icon {
  width: 42px; height: 42px; border-radius: 11px;
  background: var(--green-100);
  display: flex; align-items: center; justify-content: center;
  font-size: 20px; flex-shrink: 0;
}
.file-info .name { font-weight: 600; font-size: 14px; color: var(--gray-800); }
.file-info .meta { font-size: 12px; color: var(--gray-500); margin-top: 2px; }

/* ═══ SETTINGS ═══ */
.settings-category {
  margin-bottom: 8px;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow: hidden;
  background: var(--white);
  transition: box-shadow var(--transition);
}
.settings-category:hover { box-shadow: var(--shadow-sm); }
.settings-cat-header {
  display: flex; align-items: center; justify-content: space-between;
  padding: 12px 18px;
  background: var(--gray-25);
  cursor: pointer;
  font-size: 13.5px; font-weight: 600;
  color: var(--gray-700);
  user-select: none;
  transition: var(--transition);
  letter-spacing: -.1px;
}
.settings-cat-header:hover { background: var(--gray-75); }
.settings-cat-header .arrow {
  transition: transform .2s ease; font-size: 11px; color: var(--gray-400);
}
.settings-cat-header.open .arrow { transform: rotate(90deg); }
.settings-cat-body { display: none; }
.settings-cat-body.open { display: block; }

.rule-item {
  display: flex; align-items: center; gap: 12px;
  padding: 11px 18px;
  border-top: 1px solid var(--border-light);
  font-size: 13px;
  transition: var(--transition);
}
.rule-item:hover { background: var(--gray-25); }
.rule-desc { flex: 1; color: var(--gray-600); line-height: 1.5; }
.rule-col-tag {
  font-size: 10px; padding: 2px 8px; border-radius: 4px;
  background: var(--violet-50); color: var(--violet-600);
  font-weight: 600; white-space: nowrap;
}
.level-dot {
  width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0;
}
.level-dot.safe { background: var(--green-500); }
.level-dot.moderate { background: var(--amber-500); }
.level-dot.dangerous { background: var(--red-500); }

/* TOGGLE */
.toggle { position: relative; width: 38px; height: 21px; flex-shrink: 0; }
.toggle input { display: none; }
.toggle-slider {
  position: absolute; top:0; left:0; right:0; bottom:0;
  background: var(--gray-300); border-radius: 11px; cursor: pointer;
  transition: var(--transition);
}
.toggle-slider:before {
  content:''; position:absolute; width:15px; height:15px;
  left:3px; top:3px; background:#fff; border-radius:50%;
  transition: var(--transition);
  box-shadow: 0 1px 3px rgba(0,0,0,.12);
}
.toggle input:checked + .toggle-slider { background: var(--primary); }
.toggle input:checked + .toggle-slider:before { transform: translateX(17px); }

.rule-option {
  font-size: 11.5px; padding: 4px 10px;
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  color: var(--text); cursor: pointer; outline: none;
  font-family: var(--font-cn);
}
.rule-option:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(79,110,247,.1); }

/* ═══ SEARCH BOX ═══ */
.search-box {
  width: 100%; padding: 10px 14px 10px 40px;
  background: var(--gray-25);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  color: var(--text); font-size: 13px;
  font-family: var(--font-cn);
  outline: none; transition: var(--transition);
  margin-bottom: 14px;
}
.search-box:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(79,110,247,.08); background: var(--white); }
.search-wrap {
  position: relative;
}
.search-wrap::before {
  content: '🔍'; position: absolute; left: 13px; top: 50%;
  transform: translateY(-50%); font-size: 13px; pointer-events: none; z-index: 1;
  opacity: .6;
}

/* ═══ PROPOSALS TABLE ═══ */
.proposals-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.proposals-table th {
  text-align: left; padding: 10px 14px;
  background: var(--gray-50);
  color: var(--gray-500); font-weight: 600; font-size: 11.5px;
  text-transform: uppercase; letter-spacing: .5px;
  border-bottom: 2px solid var(--border);
  position: sticky; top: 0; z-index: 2;
}
.proposals-table td {
  padding: 10px 14px;
  border-bottom: 1px solid var(--border-light);
  vertical-align: top;
}
.proposals-table tr { transition: var(--transition-fast); cursor: pointer; }
.proposals-table tr:hover td { background: #f7f9ff; }
.proposals-table tr.selected-row td { background: #edf2ff; border-left: 3px solid var(--primary); }
.proposals-table .val-old {
  color: var(--red-500); text-decoration: line-through;
  font-family: var(--font-mono); font-size: 12px;
  background: var(--red-50); padding: 3px 8px; border-radius: 5px;
  display: inline-block; word-break: break-all;
  max-width: 260px; overflow: hidden; text-overflow: ellipsis;
}
.proposals-table .val-new {
  color: var(--green-600); font-weight: 500;
  font-family: var(--font-mono); font-size: 12px;
  background: var(--green-50); padding: 3px 8px; border-radius: 5px;
  display: inline-block; word-break: break-all;
  max-width: 260px; overflow: hidden; text-overflow: ellipsis;
}
.proposals-table .rules-tags { display: flex; flex-wrap: wrap; gap: 3px; }
.proposals-table .rule-tag {
  font-size: 10px; padding: 2px 7px; border-radius: 4px;
  background: var(--blue-50); color: var(--blue-600);
  font-weight: 600; white-space: nowrap;
  border: 1px solid var(--blue-100);
}
.proposals-table .edit-input {
  width: 100%; padding: 6px 10px;
  background: var(--white); border: 1.5px solid var(--primary);
  border-radius: var(--radius-sm);
  color: var(--text); font-size: 12px; outline: none;
  font-family: var(--font-mono);
  margin-top: 6px;
}
.proposals-table .edit-input:focus { box-shadow: 0 0 0 3px rgba(79,110,247,.1); }
.proposals-table .edit-btn {
  font-size: 11px; padding: 3px 10px; border-radius: var(--radius-sm);
  border: 1px solid var(--border); background: var(--white);
  color: var(--gray-500); cursor: pointer; margin-left: 4px;
  transition: var(--transition); font-family: var(--font-cn);
}
.proposals-table .edit-btn:hover { border-color: var(--primary); color: var(--primary); }
.proposals-table .edit-btn.active {
  border-color: var(--amber-500); color: var(--amber-700);
  background: var(--amber-50);
}
.proposals-table .override-tag {
  font-size: 10px; padding: 2px 7px; border-radius: 4px;
  background: var(--amber-50); color: var(--amber-700);
  font-weight: 600; margin-top: 3px; display: inline-block;
  border: 1px solid var(--amber-200);
}
.table-wrap {
  max-height: 520px; overflow-y: auto;
  border: 1px solid var(--border);
  border-radius: var(--radius);
}

/* ═══ REVIEW LAYOUT (split view) ═══ */
.review-split {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 20px;
  align-items: start;
}
@media (max-width: 1200px) {
  .review-split { grid-template-columns: 1fr; }
}
.review-left { min-width: 0; }
.review-right { position: sticky; top: 24px; }

/* ═══ EXCEL PREVIEW PANEL ═══ */
.preview-panel {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-md);
  overflow: hidden;
}
.preview-header {
  padding: 14px 18px;
  background: linear-gradient(180deg, var(--gray-25), var(--gray-50));
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between;
}
.preview-header .ph-title {
  font-size: 13px; font-weight: 700; color: var(--gray-700);
  display: flex; align-items: center; gap: 8px;
  letter-spacing: -.1px;
}
.preview-header .ph-cell {
  font-family: var(--font-mono); font-size: 12px;
  background: var(--primary-light); color: var(--primary);
  padding: 3px 12px; border-radius: 6px; font-weight: 700;
  border: 1px solid rgba(79,110,247,.15);
}
.preview-header .ph-sheet {
  font-size: 11.5px; color: var(--text-secondary); font-weight: 500;
}
.preview-empty {
  padding: 56px 24px; text-align: center; color: var(--text-tertiary);
  font-size: 13px;
}
.preview-empty .pe-icon { font-size: 36px; margin-bottom: 14px; opacity: .4; }
.preview-body {
  overflow: auto; max-height: 560px;
}

/* CHANGE INFO CARD in preview */
.preview-change-info {
  padding: 12px 16px; margin: 12px 14px;
  background: linear-gradient(135deg, var(--amber-50), #fff9eb);
  border: 1px solid var(--amber-200);
  border-radius: var(--radius);
  font-size: 12px;
}
.preview-change-info .pci-label { font-weight: 600; color: var(--amber-700); margin-bottom: 6px; }
.preview-change-info .pci-row { display: flex; gap: 8px; align-items: center; margin-top: 4px; }
.preview-change-info .pci-old {
  font-family: var(--font-mono); background: var(--red-50); color: var(--red-500);
  padding: 2px 8px; border-radius: 4px; text-decoration: line-through; font-size: 11.5px;
  max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.preview-change-info .pci-arrow { color: var(--gray-400); font-size: 14px; }
.preview-change-info .pci-new {
  font-family: var(--font-mono); background: var(--green-50); color: var(--green-600);
  padding: 2px 8px; border-radius: 4px; font-weight: 600; font-size: 11.5px;
  max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}

.excel-table {
  width: 100%; border-collapse: collapse;
  font-family: var(--font-mono); font-size: 11.5px;
}
.excel-table th {
  background: var(--gray-75); color: var(--gray-500);
  font-size: 10.5px; font-weight: 600;
  padding: 6px 10px; border: 1px solid var(--gray-150);
  text-align: center; position: sticky; top: 0; z-index: 1;
  letter-spacing: .3px;
}
.excel-table td {
  padding: 5px 10px; border: 1px solid var(--gray-150);
  white-space: nowrap; max-width: 180px; overflow: hidden;
  text-overflow: ellipsis; color: var(--gray-700);
  background: var(--white);
}
.excel-table .row-num {
  background: var(--gray-75); color: var(--gray-400);
  text-align: center; font-size: 10px; font-weight: 600;
  min-width: 36px;
}
.excel-table .target-cell {
  background: #fff8e1 !important;
  outline: 2px solid var(--amber-500);
  outline-offset: -1px;
  font-weight: 700; color: var(--gray-900) !important;
  position: relative;
  z-index: 1;
  animation: cellPulse 2s ease infinite;
}
@keyframes cellPulse {
  0%, 100% { outline-color: var(--amber-500); box-shadow: 0 0 0 2px rgba(245,158,11,.15); }
  50% { outline-color: var(--red-400); box-shadow: 0 0 0 4px rgba(245,158,11,.1); }
}
.excel-table .target-col-header {
  background: var(--amber-100) !important;
  color: var(--amber-700) !important;
}
.excel-table .header-row td {
  background: var(--blue-50) !important;
  font-weight: 600; color: var(--blue-700) !important;
}
.excel-table .separator-row td {
  background: var(--gray-50);
  text-align: center; color: var(--gray-400);
  font-size: 10px; border: 1px dashed var(--gray-200);
}
.excel-table .formula-cell { color: var(--violet-600) !important; font-style: italic; }

/* ═══ STATS GRID ═══ */
.stats-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 12px;
}
.stat-card {
  padding: 20px 16px; border-radius: var(--radius);
  text-align: center;
  border: 1px solid transparent;
  transition: transform var(--transition), box-shadow var(--transition);
}
.stat-card:hover { transform: translateY(-1px); box-shadow: var(--shadow-sm); }
.stat-card .num {
  font-size: 28px; font-weight: 800; line-height: 1;
  letter-spacing: -1px;
  font-family: var(--font-main);
}
.stat-card .label {
  font-size: 12px; color: var(--text-secondary); margin-top: 6px;
  font-weight: 500;
}

/* ═══ WARNING BOX ═══ */
.warning-box {
  padding: 10px 14px;
  background: var(--amber-50);
  border: 1px solid var(--amber-100);
  border-radius: var(--radius-sm);
  font-size: 12.5px; color: var(--amber-700);
  margin-bottom: 8px;
  display: flex; align-items: flex-start; gap: 8px;
}

/* ═══ COLUMN TYPES ═══ */
.col-type-grid { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
.col-type-chip {
  font-size: 12px; padding: 5px 13px; border-radius: 20px;
  background: var(--gray-50); border: 1px solid var(--border);
  font-weight: 500;
  transition: var(--transition-fast);
}
.col-type-chip:hover { background: var(--gray-75); }
.col-type-chip b { color: var(--primary); }

/* ═══ PROGRESS BAR ═══ */
.progress-bar {
  height: 4px; background: var(--gray-150); border-radius: 2px;
  overflow: hidden; margin-top: 14px;
}
.progress-fill {
  height: 100%; border-radius: 2px;
  background: linear-gradient(90deg, var(--primary), var(--green-500));
  transition: width .5s ease;
}

/* ═══ RESULT CARD ═══ */
.result-card {
  background: linear-gradient(135deg, #f0fdf4, #eef2ff, #f5f3ff);
  border: 1px solid var(--green-200);
  padding: 48px; text-align: center;
  border-radius: var(--radius-xl);
  box-shadow: var(--shadow-md);
}
.result-card h2 {
  font-size: 22px; margin-bottom: 8px; color: var(--gray-800);
  letter-spacing: -.3px;
}
.result-card .result-icon {
  width: 68px; height: 68px; border-radius: 50%;
  background: linear-gradient(135deg, var(--green-100), var(--green-200));
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 34px; margin-bottom: 18px;
  box-shadow: 0 4px 12px rgba(22,163,74,.12);
}

/* ═══ LEGEND ═══ */
.legend {
  display: flex; gap: 16px; font-size: 12px;
  color: var(--text-secondary); align-items: center;
}
.legend-item { display: flex; align-items: center; gap: 5px; }

/* ═══ TOAST NOTIFICATIONS ═══ */
.toast-container { position: fixed; top: 20px; right: 20px; z-index: 9999; }
.toast {
  padding: 12px 20px; border-radius: var(--radius);
  margin-bottom: 8px; font-size: 13px; font-weight: 500;
  animation: slideIn .3s ease;
  box-shadow: var(--shadow-lg);
  display: flex; align-items: center; gap: 8px;
  max-width: 360px;
  backdrop-filter: blur(10px);
}
.toast-success { background: var(--green-600); color: #fff; }
.toast-error { background: var(--red-500); color: #fff; }
.toast-info { background: var(--primary); color: #fff; }
@keyframes slideIn {
  from { transform: translateX(100%); opacity: 0; }
  to { transform: none; opacity: 1; }
}

/* ═══ SPINNER ═══ */
.spinner {
  width: 18px; height: 18px;
  border: 2px solid rgba(255,255,255,.3);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin .6s linear infinite;
  display: inline-block;
}
.spinner-dark {
  border-color: var(--gray-200);
  border-top-color: var(--primary);
}
@keyframes spin { to { transform: rotate(360deg); } }

/* ═══ CHECKBOX STYLED ═══ */
.chk-styled {
  width: 18px; height: 18px; border-radius: 5px;
  border: 2px solid var(--gray-300);
  appearance: none; -webkit-appearance: none;
  cursor: pointer; transition: var(--transition);
  position: relative; vertical-align: middle;
}
.chk-styled:checked {
  background: var(--primary); border-color: var(--primary);
}
.chk-styled:checked::after {
  content: '✓'; position: absolute;
  top: 50%; left: 50%; transform: translate(-50%, -50%);
  color: #fff; font-size: 11px; font-weight: 800;
}
.chk-styled:hover { border-color: var(--primary); }

/* ═══ SCROLLBAR ═══ */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--gray-200); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--gray-300); }

/* ═══ UTILITY ═══ */
.hidden { display: none !important; }
.text-mono { font-family: var(--font-mono); }
.text-sm { font-size: 12px; }
.text-muted { color: var(--text-secondary); }
.mt-4 { margin-top: 16px; }
.mt-3 { margin-top: 12px; }
.mb-3 { margin-bottom: 12px; }
.flex-between { display: flex; align-items: center; justify-content: space-between; }

/* ═══ SHEET CHECKBOX ═══ */
.sheet-label {
  display: flex; align-items: center; gap: 10px;
  padding: 9px 14px; font-size: 13.5px; cursor: pointer;
  border-radius: var(--radius-sm);
  transition: var(--transition);
}
.sheet-label:hover { background: var(--gray-50); }
.sheet-label .sheet-meta {
  font-size: 11.5px; color: var(--text-tertiary); font-weight: 400;
}

/* ═══ PAGE-LEVEL ANIMATION ═══ */
[id^="page-"]:not(.hidden) {
  animation: pageFadeIn .3s ease;
}
@keyframes pageFadeIn {
  from { opacity: 0; transform: translateY(6px); }
  to { opacity: 1; transform: none; }
}

/* ═══ EMPTY STATE ═══ */
.empty-state {
  text-align: center; padding: 40px 20px; color: var(--text-tertiary);
}
.empty-state .es-icon { font-size: 40px; opacity: .4; margin-bottom: 12px; }
.empty-state .es-text { font-size: 13px; line-height: 1.6; }
</style>
</head>
<body>
<div class="app">
  <!-- ══════ SIDEBAR ══════ -->
  <div class="sidebar">
    <div class="sidebar-brand">
      <h1>
        <span class="brand-icon">📊</span>
        Excel Standardizer
      </h1>
      <p>v4.0 · 安全优先 · 两阶段处理</p>
    </div>
    <div class="nav">
      <div class="nav-item active" onclick="showPage('upload')">
        <span class="icon">📂</span> 上传文件
        <span class="badge-check" id="badge-upload" style="display:none">✓</span>
      </div>
      <div class="nav-item" onclick="showPage('settings')">
        <span class="icon">⚙️</span> 规则设置
        <span class="badge" id="badge-settings"></span>
      </div>
      <div class="nav-item" onclick="showPage('analyze')">
        <span class="icon">🔍</span> 扫描分析
      </div>
      <div class="nav-item" onclick="showPage('review')">
        <span class="icon">👁️</span> 审核变更
        <span class="badge" id="badge-review" style="display:none">0</span>
      </div>
      <div class="nav-item" onclick="showPage('result')">
        <span class="icon">✅</span> 导出结果
      </div>
    </div>
    <div class="sidebar-footer">
      <strong>核心原则</strong><br>
      宁可漏改，不可误改<br>
      公式不动 · 格式不改
    </div>
  </div>

  <!-- ══════ MAIN CONTENT ══════ -->
  <div class="main">
    <div class="toast-container" id="toasts"></div>

    <!-- ────── PAGE: Upload ────── -->
    <div id="page-upload">
      <div class="steps">
        <div class="step active"><span class="num">1</span>上传</div>
        <div class="step"><span class="num">2</span>设置</div>
        <div class="step"><span class="num">3</span>分析</div>
        <div class="step"><span class="num">4</span>审核</div>
        <div class="step"><span class="num">5</span>导出</div>
      </div>
      <div class="page-header">
        <h2><span class="ph-icon">📂</span> 上传 Excel 文件</h2>
        <div class="ph-desc">支持 .xlsx / .xlsm 格式，最大 100MB</div>
      </div>
      <div class="card">
        <div class="upload-zone" id="dropZone"
             onclick="document.getElementById('fileInput').click()"
             ondragover="event.preventDefault();this.classList.add('dragover')"
             ondragleave="this.classList.remove('dragover')"
             ondrop="handleDrop(event)">
          <div class="upload-icon">📄</div>
          <p>点击选择文件或拖拽到此处</p>
          <div class="hint">仅支持 .xlsx / .xlsm 格式 · 最大 100MB</div>
        </div>
        <input type="file" id="fileInput" accept=".xlsx,.xlsm" style="display:none" onchange="handleFile(this.files[0])">
        <div id="fileInfoBox" class="hidden" style="margin-top:16px"></div>
      </div>
    </div>

    <!-- ────── PAGE: Settings ────── -->
    <div id="page-settings" class="hidden">
      <div class="steps">
        <div class="step done"><span class="num">1</span>上传</div>
        <div class="step active"><span class="num">2</span>设置</div>
        <div class="step"><span class="num">3</span>分析</div>
        <div class="step"><span class="num">4</span>审核</div>
        <div class="step"><span class="num">5</span>导出</div>
      </div>
      <div class="page-header">
        <h2><span class="ph-icon">⚙️</span> 规则设置</h2>
        <div class="ph-desc">勾选要启用的标准化规则 · 红色标记的规则需确认列类型</div>
      </div>
      <div class="card">
        <div class="card-header">
          <div class="legend">
            <div class="legend-item"><div class="level-dot safe"></div> 安全</div>
            <div class="legend-item"><div class="level-dot moderate"></div> 中等</div>
            <div class="legend-item"><div class="level-dot dangerous"></div> 危险</div>
          </div>
          <div class="btn-group">
            <button class="btn btn-outline btn-sm" onclick="presetSafe()">🛡️ 仅安全</button>
            <button class="btn btn-outline btn-sm" onclick="presetAll(true)">全部开启</button>
            <button class="btn btn-outline btn-sm" onclick="presetAll(false)">全部关闭</button>
            <button class="btn btn-outline btn-sm" onclick="presetDefault()">恢复默认</button>
          </div>
        </div>
        <div class="search-wrap">
          <input class="search-box" id="ruleSearch" placeholder="搜索规则名称或编号..." oninput="filterRules(this.value)">
        </div>
        <div id="settingsContainer"></div>
      </div>
      <div style="text-align:right;margin-top:14px">
        <button class="btn btn-primary" onclick="showPage('analyze')">下一步 → 扫描分析</button>
      </div>
    </div>

    <!-- ────── PAGE: Analyze ────── -->
    <div id="page-analyze" class="hidden">
      <div class="steps">
        <div class="step done"><span class="num">1</span>上传</div>
        <div class="step done"><span class="num">2</span>设置</div>
        <div class="step active"><span class="num">3</span>分析</div>
        <div class="step"><span class="num">4</span>审核</div>
        <div class="step"><span class="num">5</span>导出</div>
      </div>
      <div class="page-header">
        <h2><span class="ph-icon">🔍</span> 扫描分析</h2>
        <div class="ph-desc">扫描所有单元格，计算变更提案（不会修改原始文件）</div>
      </div>
      <div class="card" id="sheetsCard">
        <div class="card-title"><span class="title-icon">📋</span> 选择工作表</div>
        <div id="sheetsContainer" style="margin-top:12px"></div>
      </div>
      <div class="card">
        <div class="card-header">
          <div class="card-title"><span class="title-icon">🔍</span> 开始扫描</div>
          <button class="btn btn-primary" id="analyzeBtn" onclick="runAnalysis()">开始扫描</button>
        </div>
        <div class="progress-bar hidden" id="analyzeProgress"><div class="progress-fill" style="width:0%"></div></div>
        <div id="analyzeResults" class="hidden" style="margin-top:20px"></div>
      </div>
    </div>

    <!-- ────── PAGE: Review (Split View) ────── -->
    <div id="page-review" class="hidden">
      <div class="steps">
        <div class="step done"><span class="num">1</span>上传</div>
        <div class="step done"><span class="num">2</span>设置</div>
        <div class="step done"><span class="num">3</span>分析</div>
        <div class="step active"><span class="num">4</span>审核</div>
        <div class="step"><span class="num">5</span>导出</div>
      </div>
      <div class="page-header">
        <h2><span class="ph-icon">👁️</span> 审核变更</h2>
        <div class="ph-desc">点击左侧条目可在右侧预览原始文件中该单元格位置</div>
      </div>
      <div class="review-split">
        <!-- LEFT: proposals table -->
        <div class="review-left">
          <div class="card" style="margin-bottom:0;padding:18px">
            <div class="card-header" style="margin-bottom:14px">
              <div class="card-title"><span class="title-icon">📝</span> 变更提案</div>
              <div class="btn-group">
                <button class="btn btn-success btn-sm" onclick="setAllProposals(true)">✓ 全部接受</button>
                <button class="btn btn-danger btn-sm" onclick="setAllProposals(false)">✗ 全部拒绝</button>
              </div>
            </div>
            <div class="search-wrap">
              <input class="search-box" placeholder="搜索变更内容..." oninput="filterProposals(this.value)">
            </div>
            <div class="table-wrap" id="proposalsTableWrap"></div>
            <div style="margin-top:16px;text-align:right">
              <button class="btn btn-primary" onclick="applyChanges()">✅ 应用选中的变更</button>
            </div>
          </div>
        </div>

        <!-- RIGHT: Excel preview -->
        <div class="review-right">
          <div class="preview-panel" id="previewPanel">
            <div class="preview-header">
              <div class="ph-title">📊 原始文件预览</div>
              <div id="previewCellInfo"></div>
            </div>
            <div id="previewChangeInfo"></div>
            <div id="previewContent">
              <div class="preview-empty">
                <div class="pe-icon">📋</div>
                <div>点击左侧任意变更条目<br>即可预览该单元格在原始文件中的位置</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ────── PAGE: Result ────── -->
    <div id="page-result" class="hidden">
      <div class="steps">
        <div class="step done"><span class="num">1</span>上传</div>
        <div class="step done"><span class="num">2</span>设置</div>
        <div class="step done"><span class="num">3</span>分析</div>
        <div class="step done"><span class="num">4</span>审核</div>
        <div class="step active"><span class="num">5</span>导出</div>
      </div>
      <div class="page-header">
        <h2><span class="ph-icon">✅</span> 导出结果</h2>
        <div class="ph-desc">下载标准化后的文件及变更日志</div>
      </div>
      <div id="resultContent"></div>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════
//  Global State
// ═══════════════════════════════════════════════════
let SID = null;
let settingsData = [];
let proposalsData = [];
let fileUploaded = false;
let analysisData = null;
let currentPreviewIdx = -1;

async function init() {
  const r = await fetch('/api/session', {method:'POST'});
  const d = await r.json();
  SID = d.session_id;
  loadSettings();
}

// ═══════════════════════════════════════════════════
//  Navigation
// ═══════════════════════════════════════════════════
function showPage(name) {
  document.querySelectorAll('[id^="page-"]').forEach(p => p.classList.add('hidden'));
  document.getElementById('page-' + name).classList.remove('hidden');
  document.querySelectorAll('.nav-item').forEach((n, i) => {
    n.classList.remove('active');
    if (['upload','settings','analyze','review','result'][i] === name) n.classList.add('active');
  });
}

function toast(msg, type='info') {
  const c = document.getElementById('toasts');
  const t = document.createElement('div');
  t.className = 'toast toast-' + type;
  const icons = { success: '✓', error: '✗', info: 'ℹ' };
  t.innerHTML = `<span>${icons[type] || 'ℹ'}</span> ${msg}`;
  c.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(30px)'; setTimeout(() => t.remove(), 200); }, 3500);
}

// ═══════════════════════════════════════════════════
//  File Upload
// ═══════════════════════════════════════════════════
function handleDrop(e) {
  e.preventDefault();
  e.currentTarget.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
}

async function handleFile(file) {
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  toast('正在上传...', 'info');
  try {
    const r = await fetch(`/api/upload/${SID}`, {method:'POST', body: fd});
    const d = await r.json();
    if (d.error) { toast(d.error, 'error'); return; }
    fileUploaded = true;
    document.getElementById('badge-upload').style.display = '';

    let html = `<div class="file-info">
      <div class="fi-icon">📊</div>
      <div><div class="name">${d.filename}</div>
      <div class="meta">${d.sheets.length} 个工作表</div></div></div>`;
    if (d.warnings.length > 0) {
      html += d.warnings.map(w => `<div class="warning-box"><span>⚠️</span> ${w}</div>`).join('');
    }
    html += '<div style="margin-top:14px;font-size:13px;font-weight:600;color:var(--gray-600)">工作表:</div>';
    html += '<div class="col-type-grid">';
    html += d.sheets.map(s => `<div class="col-type-chip">${s.name} <span style="color:var(--text-tertiary)">(${s.rows}行×${s.cols}列)</span></div>`).join('');
    html += '</div>';
    document.getElementById('fileInfoBox').innerHTML = html;
    document.getElementById('fileInfoBox').classList.remove('hidden');
    renderSheets(d.sheets);
    toast('文件上传成功!', 'success');
  } catch(e) { toast('上传失败: ' + e.message, 'error'); }
}

function renderSheets(sheets) {
  let html = '';
  sheets.forEach(s => {
    html += `<label class="sheet-label">
      <input type="checkbox" class="chk-styled sheet-check" value="${s.name}" checked>
      <b>${s.name}</b>
      <span class="sheet-meta">(${s.rows}行×${s.cols}列)</span>
    </label>`;
  });
  document.getElementById('sheetsContainer').innerHTML = html;
}

async function generateTest() {
  toast('正在生成测试文件...', 'info');
  const r = await fetch(`/api/generate_test/${SID}`, {method:'POST'});
  const d = await r.json();
  toast('测试文件已生成', 'success');
  window.open(`/api/download/${SID}/${d.filename}`);
}

// ═══════════════════════════════════════════════════
//  Settings
// ═══════════════════════════════════════════════════
async function loadSettings() {
  const r = await fetch(`/api/settings/${SID}`);
  settingsData = await r.json();
  renderSettings();
  updateSettingsBadge();
}

function renderSettings() {
  const c = document.getElementById('settingsContainer');
  let html = '';
  settingsData.forEach((cat, ci) => {
    const enabledCount = cat.rules.filter(r => r.enabled).length;
    html += `<div class="settings-category" data-cat="${ci}">
      <div class="settings-cat-header" onclick="toggleCat(this)">
        <span>${cat.category} <span style="color:var(--text-tertiary);font-weight:400;font-size:12px">(${enabledCount}/${cat.rules.length})</span></span>
        <span class="arrow">▶</span>
      </div>
      <div class="settings-cat-body">`;
    cat.rules.forEach(rule => {
      const colTags = rule.col_types.length > 0
        ? rule.col_types.map(t => `<span class="rule-col-tag">仅:${t}</span>`).join(' ')
        : '';
      const optHtml = rule.options
        ? `<select class="rule-option" onchange="updateRuleOption('${rule.key}',this.value)" ${!rule.enabled?'disabled':''}>${rule.options.map(o=>`<option value="${o}" ${o===rule.option?'selected':''}>${o}</option>`).join('')}</select>`
        : '';
      html += `<div class="rule-item" data-key="${rule.key}" data-desc="${rule.desc}">
        <div class="level-dot ${rule.level}"></div>
        <div class="rule-desc">${rule.desc} ${colTags}</div>
        ${optHtml}
        <label class="toggle">
          <input type="checkbox" ${rule.enabled?'checked':''} onchange="toggleRule('${rule.key}',this.checked)">
          <span class="toggle-slider"></span>
        </label>
      </div>`;
    });
    html += '</div></div>';
  });
  c.innerHTML = html;
}

function toggleCat(el) {
  el.classList.toggle('open');
  el.nextElementSibling.classList.toggle('open');
}

async function toggleRule(key, enabled) {
  await fetch(`/api/settings/${SID}`, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({key, enabled})
  });
  settingsData.forEach(cat => cat.rules.forEach(r => { if(r.key===key) r.enabled=enabled; }));
  updateSettingsBadge();
  const item = document.querySelector(`.rule-item[data-key="${key}"] select`);
  if (item) item.disabled = !enabled;
}

async function updateRuleOption(key, option) {
  await fetch(`/api/settings/${SID}`, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({key, enabled:true, option})
  });
}

function updateSettingsBadge() {
  let n = 0;
  settingsData.forEach(cat => cat.rules.forEach(r => { if(r.enabled) n++; }));
  document.getElementById('badge-settings').textContent = n;
}

async function presetSafe() {
  await fetch(`/api/settings/${SID}`, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({preset:'safe'})
  });
  await loadSettings(); toast('已切换为安全模式', 'success');
}
async function presetAll(on) {
  await fetch(`/api/settings/${SID}`, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({preset: on?'all_on':'all_off'})
  });
  await loadSettings(); toast(on?'全部开启':'全部关闭', 'info');
}
async function presetDefault() {
  await fetch(`/api/settings/${SID}`, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({preset:'default'})
  });
  await loadSettings(); toast('已恢复默认设置', 'success');
}

function filterRules(q) {
  q = q.toLowerCase();
  document.querySelectorAll('.rule-item').forEach(el => {
    const match = el.dataset.desc.toLowerCase().includes(q) || el.dataset.key.toLowerCase().includes(q);
    el.style.display = match ? '' : 'none';
  });
}

// ═══════════════════════════════════════════════════
//  Analysis
// ═══════════════════════════════════════════════════
async function runAnalysis() {
  if (!fileUploaded) { toast('请先上传文件', 'error'); return; }
  const sheets = [...document.querySelectorAll('.sheet-check:checked')].map(c => c.value);
  if (!sheets.length) { toast('请至少选择一个工作表', 'error'); return; }

  const btn = document.getElementById('analyzeBtn');
  btn.innerHTML = '<span class="spinner"></span> 分析中...';
  btn.disabled = true;
  const pb = document.getElementById('analyzeProgress');
  pb.classList.remove('hidden');
  pb.querySelector('.progress-fill').style.width = '60%';

  try {
    const r = await fetch(`/api/analyze/${SID}`, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({sheets})
    });
    const d = await r.json();
    if (d.error) { toast(d.error, 'error'); return; }
    analysisData = d;
    proposalsData = d.proposals;
    pb.querySelector('.progress-fill').style.width = '100%';

    let html = '<div class="stats-grid">';
    html += `<div class="stat-card" style="background:var(--blue-50);border-color:var(--blue-200)"><div class="num" style="color:var(--blue-600)">${d.total_changes}</div><div class="label">待变更单元格</div></div>`;
    html += `<div class="stat-card" style="background:var(--amber-50);border-color:var(--amber-100)"><div class="num" style="color:var(--amber-600)">${d.warnings.length}</div><div class="label">警告</div></div>`;
    html += `<div class="stat-card" style="background:var(--violet-50);border-color:var(--violet-100)"><div class="num" style="color:var(--violet-600)">${d.skipped}</div><div class="label">跳过(公式等)</div></div>`;
    html += `<div class="stat-card" style="background:var(--green-50);border-color:var(--green-100)"><div class="num" style="color:var(--green-600)">${Object.keys(d.rule_stats).length}</div><div class="label">触发规则数</div></div>`;
    html += '</div>';

    // Column types
    const ctEntries = Object.entries(d.col_types).filter(([_, v]) => Object.keys(v).length > 0);
    if (ctEntries.length > 0) {
      html += '<div style="margin-top:18px"><b style="color:var(--gray-700)">🔍 列类型推断:</b></div>';
      ctEntries.forEach(([sn, cols]) => {
        html += `<div style="margin-top:6px;font-size:13px"><span style="color:var(--text-secondary)">[${sn}]</span>`;
        html += '<div class="col-type-grid">';
        Object.entries(cols).forEach(([col, type]) => {
          html += `<div class="col-type-chip">列${col}: <b>${type}</b></div>`;
        });
        html += '</div></div>';
      });
    }

    // Rule stats
    if (Object.keys(d.rule_stats).length > 0) {
      html += '<div style="margin-top:18px"><b style="color:var(--gray-700)">📋 规则触发统计:</b></div>';
      html += '<div class="col-type-grid">';
      Object.entries(d.rule_stats).sort((a,b) => b[1]-a[1]).forEach(([rule, cnt]) => {
        html += `<div class="col-type-chip">${rule}: <b>${cnt}</b></div>`;
      });
      html += '</div>';
    }

    // Warnings
    if (d.warnings.length > 0) {
      html += '<div style="margin-top:18px"><b style="color:var(--gray-700)">⚠️ 警告:</b></div>';
      d.warnings.slice(0,10).forEach(w => {
        html += `<div class="warning-box"><span>⚠️</span> [${w.sheet}] ${w.cell}: ${w.message}</div>`;
      });
    }

    if (d.total_changes > 0) {
      html += `<div style="margin-top:22px;text-align:right">
        <button class="btn btn-primary" onclick="showPage('review')">下一步 → 审核变更 (${d.total_changes}处)</button></div>`;
    }

    document.getElementById('analyzeResults').innerHTML = html;
    document.getElementById('analyzeResults').classList.remove('hidden');
    document.getElementById('badge-review').style.display = '';
    document.getElementById('badge-review').textContent = d.total_changes;

    renderProposals();
    toast(`分析完成: ${d.total_changes} 处变更`, 'success');
  } catch(e) { toast('分析失败: ' + e.message, 'error'); }
  finally {
    btn.innerHTML = '开始扫描';
    btn.disabled = false;
  }
}

// ═══════════════════════════════════════════════════
//  Review & Excel Preview
// ═══════════════════════════════════════════════════
function renderProposals() {
  if (!proposalsData.length) {
    document.getElementById('proposalsTableWrap').innerHTML = '<div class="empty-state"><div class="es-icon">📋</div><div class="es-text">没有变更提案</div></div>';
    return;
  }
  let html = `<table class="proposals-table"><thead><tr>
    <th style="width:36px">✓</th><th>位置</th><th>原始值</th><th>建议修改</th><th>规则</th>
  </tr></thead><tbody>`;
  proposalsData.forEach((p, i) => {
    p.override = p.override || null;
    html += `<tr data-idx="${i}" data-searchable="${p.original} ${p.proposed} ${p.sheet} ${p.cell}"
             onclick="onProposalClick(${i}, event)">
      <td><input type="checkbox" class="chk-styled proposal-check" data-idx="${i}" ${p.accepted?'checked':''}
          onclick="event.stopPropagation()" onchange="onProposalToggle(${i},this.checked)"></td>
      <td><div style="font-weight:600;color:var(--gray-700);font-size:13px">${p.cell}</div>
          <div style="font-size:11px;color:var(--text-tertiary)">${p.sheet}</div></td>
      <td><div class="val-old">${escHtml(p.original.slice(0,60))}</div></td>
      <td>
        <div class="val-new" id="proposed-${i}">${escHtml(p.proposed.slice(0,60))}</div>
        <div id="override-show-${i}" style="display:none"><span class="override-tag">✏️ 自定义</span> <span id="override-val-${i}" class="text-mono text-sm"></span></div>
        <div id="edit-area-${i}" style="display:none">
          <input class="edit-input" id="edit-input-${i}" placeholder="输入自定义值..." onkeydown="if(event.key==='Enter')saveEdit(${i})" onclick="event.stopPropagation()">
          <div style="margin-top:4px" onclick="event.stopPropagation()">
            <button class="edit-btn" onclick="saveEdit(${i})">✓ 保存</button>
            <button class="edit-btn" onclick="cancelEdit(${i})">✗ 取消</button>
            <button class="edit-btn" onclick="clearEdit(${i})">↩ 还原</button>
          </div>
        </div>
        <button class="edit-btn" id="edit-trigger-${i}" onclick="event.stopPropagation();openEdit(${i})" style="margin-top:4px">✏️ 编辑</button>
      </td>
      <td><div class="rules-tags">${p.rules.map(r=>`<span class="rule-tag">${r}</span>`).join('')}</div></td>
    </tr>`;
  });
  html += '</tbody></table>';
  document.getElementById('proposalsTableWrap').innerHTML = html;
}

function onProposalToggle(idx, checked) {
  proposalsData[idx].accepted = checked;
}

// ── Excel Preview on click ──
async function onProposalClick(idx, event) {
  // Don't trigger on checkbox or button clicks
  if (event.target.tagName === 'INPUT' || event.target.tagName === 'BUTTON') return;

  const p = proposalsData[idx];
  currentPreviewIdx = idx;

  // Highlight selected row
  document.querySelectorAll('.proposals-table tr.selected-row').forEach(tr => tr.classList.remove('selected-row'));
  const row = document.querySelector(`.proposals-table tr[data-idx="${idx}"]`);
  if (row) row.classList.add('selected-row');

  // Update header info
  document.getElementById('previewCellInfo').innerHTML = `
    <span class="ph-sheet">${p.sheet}</span>
    <span class="ph-cell">${p.cell}</span>
  `;

  // Show change info card
  document.getElementById('previewChangeInfo').innerHTML = `
    <div class="preview-change-info">
      <div class="pci-label">📌 变更详情</div>
      <div class="pci-row">
        <span class="pci-old" title="${escHtml(p.original)}">${escHtml(p.original.slice(0,40))}</span>
        <span class="pci-arrow">→</span>
        <span class="pci-new" title="${escHtml(p.proposed)}">${escHtml(p.proposed.slice(0,40))}</span>
      </div>
    </div>
  `;

  // Show loading
  document.getElementById('previewContent').innerHTML = `
    <div class="preview-empty"><span class="spinner spinner-dark" style="width:24px;height:24px"></span><div style="margin-top:10px">加载预览...</div></div>
  `;

  try {
    const r = await fetch(`/api/preview/${SID}`, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        sheet: p.sheet,
        row: p.row,
        col: p.col,
      })
    });
    const d = await r.json();
    if (d.error) {
      document.getElementById('previewContent').innerHTML = `<div class="preview-empty">⚠️ ${d.error}</div>`;
      return;
    }
    renderExcelPreview(d);
  } catch(e) {
    document.getElementById('previewContent').innerHTML = `<div class="preview-empty">加载失败: ${e.message}</div>`;
  }
}

function renderExcelPreview(data) {
  let html = '<div class="preview-body"><table class="excel-table">';

  // Column headers row
  html += '<tr><th class="row-num"></th>';
  data.col_headers.forEach(ch => {
    const cls = ch.is_target_col ? 'target-col-header' : '';
    html += `<th class="${cls}">${ch.letter}</th>`;
  });
  html += '</tr>';

  // Data rows
  data.rows.forEach(row => {
    if (row.is_separator) {
      html += `<tr class="separator-row"><td class="row-num">⋮</td>`;
      data.col_headers.forEach(() => { html += '<td>⋯</td>'; });
      html += '</tr>';
      return;
    }

    const isHeaderRow = row.row === 1;
    html += `<tr class="${isHeaderRow ? 'header-row' : ''}"><td class="row-num">${row.row}</td>`;
    row.cells.forEach(cell => {
      let cls = '';
      if (cell.is_target) cls += ' target-cell';
      if (cell.is_formula) cls += ' formula-cell';
      const val = escHtml(cell.value || '');
      html += `<td class="${cls}" title="${escHtml(cell.value || '')}">${val || '<span style="color:var(--gray-300)">—</span>'}</td>`;
    });
    html += '</tr>';
  });

  html += '</table></div>';

  // Info bar at bottom
  html += `<div style="padding:10px 16px;background:var(--gray-25);border-top:1px solid var(--border);font-size:11px;color:var(--text-tertiary);display:flex;justify-content:space-between">
    <span>Sheet: ${data.sheet} · 共 ${data.total_rows} 行 × ${data.total_cols} 列</span>
    <span>目标: ${data.target_cell}</span>
  </div>`;

  document.getElementById('previewContent').innerHTML = html;
}

function openEdit(idx) {
  document.getElementById('edit-area-' + idx).style.display = '';
  document.getElementById('edit-trigger-' + idx).style.display = 'none';
  const input = document.getElementById('edit-input-' + idx);
  input.value = proposalsData[idx].override || proposalsData[idx].proposed;
  input.focus();
}

function saveEdit(idx) {
  const val = document.getElementById('edit-input-' + idx).value;
  proposalsData[idx].override = val;
  proposalsData[idx].accepted = true;
  document.querySelector(`.proposal-check[data-idx="${idx}"]`).checked = true;
  document.getElementById('proposed-' + idx).style.display = 'none';
  document.getElementById('override-show-' + idx).style.display = '';
  document.getElementById('override-val-' + idx).textContent = val.slice(0, 60);
  document.getElementById('edit-area-' + idx).style.display = 'none';
  document.getElementById('edit-trigger-' + idx).style.display = '';
  document.getElementById('edit-trigger-' + idx).classList.add('active');
  document.getElementById('edit-trigger-' + idx).textContent = '✏️ 已自定义';
  toast(`[${proposalsData[idx].cell}] 已自定义修改`, 'success');
}

function cancelEdit(idx) {
  document.getElementById('edit-area-' + idx).style.display = 'none';
  document.getElementById('edit-trigger-' + idx).style.display = '';
}

function clearEdit(idx) {
  proposalsData[idx].override = null;
  document.getElementById('proposed-' + idx).style.display = '';
  document.getElementById('override-show-' + idx).style.display = 'none';
  document.getElementById('edit-area-' + idx).style.display = 'none';
  document.getElementById('edit-trigger-' + idx).style.display = '';
  document.getElementById('edit-trigger-' + idx).classList.remove('active');
  document.getElementById('edit-trigger-' + idx).textContent = '✏️ 编辑';
  toast(`[${proposalsData[idx].cell}] 已还原为建议值`, 'info');
}

function setAllProposals(accepted) {
  proposalsData.forEach(p => p.accepted = accepted);
  document.querySelectorAll('.proposal-check').forEach(c => c.checked = accepted);
  toast(accepted ? '全部接受' : '全部拒绝', 'info');
}

function filterProposals(q) {
  q = q.toLowerCase();
  document.querySelectorAll('.proposals-table tbody tr').forEach(tr => {
    tr.style.display = (tr.dataset.searchable||'').toLowerCase().includes(q) ? '' : 'none';
  });
}

function escHtml(s) {
  if (!s) return '';
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ═══════════════════════════════════════════════════
//  Apply Changes
// ═══════════════════════════════════════════════════
async function applyChanges() {
  const decisions = {};
  proposalsData.forEach((p, i) => {
    decisions[i] = {accepted: p.accepted};
    if (p.override) decisions[i].override = p.override;
  });

  toast('正在应用变更...', 'info');
  try {
    const r = await fetch(`/api/apply/${SID}`, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({decisions})
    });
    const d = await r.json();
    if (d.error) { toast(d.error, 'error'); return; }

    let html = `<div class="result-card">
      <div class="result-icon">✅</div>
      <h2>处理完成!</h2>
      <p style="color:var(--text-secondary);margin-bottom:24px;font-size:14px">
        共应用 <b style="color:var(--green-600)">${d.confirmed}</b> 处变更，
        拒绝 <b style="color:var(--red-500)">${d.rejected}</b> 处
      </p>
      <div class="btn-group" style="justify-content:center">
        <a class="btn btn-success" href="/api/download/${SID}/${encodeURIComponent(d.output_file)}" target="_blank">
          📥 下载标准化文件
        </a>`;
    if (d.log_file) {
      html += `<a class="btn btn-outline" href="/api/download/${SID}/${encodeURIComponent(d.log_file)}" target="_blank">
          📋 下载变更日志 (${d.log_count}条)
        </a>`;
    }
    html += `</div></div>`;

    html += `<div class="card" style="margin-top:20px">
      <div class="card-title"><span class="title-icon">📊</span> 处理统计</div>
      <div class="stats-grid" style="margin-top:14px">
        <div class="stat-card" style="background:var(--green-50);border-color:var(--green-100)"><div class="num" style="color:var(--green-600)">${d.confirmed}</div><div class="label">已应用</div></div>
        <div class="stat-card" style="background:var(--red-50);border-color:var(--red-100)"><div class="num" style="color:var(--red-500)">${d.rejected}</div><div class="label">已拒绝</div></div>
        <div class="stat-card" style="background:var(--blue-50);border-color:var(--blue-200)"><div class="num" style="color:var(--blue-600)">${d.log_count}</div><div class="label">日志记录</div></div>
      </div>
    </div>`;

    document.getElementById('resultContent').innerHTML = html;
    showPage('result');
    toast('处理完成!', 'success');
  } catch(e) { toast('应用失败: ' + e.message, 'error'); }
}

// ═══════════════════════════════════════════════════
//  Init
// ═══════════════════════════════════════════════════
init();
</script>
</body>
</html>"""


# ============================================================
#  启动服务器
# ============================================================

if __name__ == '__main__':
    import webbrowser, threading
    port = 5001
    url = f'http://localhost:{port}'
    print("=" * 60)
    print("  Excel Standardizer Web v4.0")
    print(f"  🌐 打开浏览器访问: {url}")
    print("  按 Ctrl+C 停止服务器")
    print("=" * 60)
    # 自动打开浏览器
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host='0.0.0.0', port=port, debug=False)