import os
import json
import re
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for

# pandas is imported lazily inside parse_xls() to avoid OpenBLAS OOM at startup

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True            # always reload from disk
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DATA_FOLDER'] = 'data'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DATA_FOLDER'], exist_ok=True)

# data/files_index.json  →  { "filename": { uploaded_at, school_count, record_count } }
# data/schools_<hash>.json → list of aggregated school dicts for that file
FILES_INDEX = os.path.join(app.config['DATA_FOLDER'], 'files_index.json')

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}


def detect_year(filename, records=None):
    """Try to extract a 4-digit year (2010-2035) from filename first, then from program names."""
    match = re.search(r'20[1-3][0-9]', filename)
    if match:
        return int(match.group())
    if records:
        years = set()
        for r in records:
            m = re.search(r'20[1-3][0-9]', r.get('destinacao', ''))
            if m:
                years.add(int(m.group()))
        if years:
            return max(years)
    return None


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_key(filename):
    """Filesystem-safe key derived from filename."""
    return re.sub(r'[^a-zA-Z0-9_\-.]', '_', filename)


def schools_path(filename):
    return os.path.join(app.config['DATA_FOLDER'], f'schools_{safe_key(filename)}.json')


# ── Parsing ──────────────────────────────────────────────────────────────────

def to_float(value):
    if value is None or (isinstance(value, float) and value != value):  # None or NaN
        return 0.0
    if value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def classify_pdde(destinacao):
    if destinacao is None or (isinstance(destinacao, float) and destinacao != destinacao):
        return None
    d = str(destinacao)
    if re.search(r'EQUIDADE', d, re.IGNORECASE):
        return 'PDDE EQUIDADE'
    if re.search(r'QUALIDADE', d, re.IGNORECASE):
        return 'PDDE QUALIDADE'
    if re.search(r'PDDE\s*/\s*PDDE', d, re.IGNORECASE):
        return 'PDDE Básico'
    return None


def parse_xls(filepath):
    import pandas as pd  # lazy import — keeps Flask startup free of OpenBLAS
    ext = filepath.rsplit('.', 1)[-1].lower()
    if ext == 'xlsx':
        df = pd.read_excel(filepath, engine='openpyxl')
    else:
        tables = pd.read_html(filepath, encoding='latin1')
        df = None
        for t in tables:
            cols = [str(c) for c in t.columns]
            if any('escola' in c.lower() or 'município' in c.lower() for c in cols):
                df = t
                break
        if df is None:
            df = tables[0]

    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'nome escola' in cl:
            col_map['nome_escola'] = c
        elif 'município' in cl or 'municipio' in cl:
            col_map['municipio'] = c
        elif cl == 'uf':
            col_map['uf'] = c
        elif 'destinação' in cl or 'destinacao' in cl:
            col_map['destinacao'] = c
        elif 'valor pago total' in cl or 'vl pago total' in cl:
            col_map['valor_pago'] = c
        elif 'vl final devido total' in cl:
            col_map['valor_devido'] = c
        elif 'vl devido total' in cl and 'vl final' not in cl:
            col_map['valor_devido_raw'] = c
        elif 'cod' in cl and 'escola' in cl:
            col_map['cod_escola'] = c
        elif 'cnpj' in cl:
            col_map['cnpj'] = c
        elif 'rede' in cl:
            col_map['rede'] = c

    records = []
    for _, row in df.iterrows():
        dest = row.get(col_map.get('destinacao', ''), None)
        cat = classify_pdde(dest)
        if cat is None:
            continue
        nome = str(row.get(col_map.get('nome_escola', ''), '')).strip()
        if not nome or nome in ('nan', 'NaN', 'Subtotal', 'Total Geral'):
            continue

        valor_pago_col = col_map.get('valor_pago', col_map.get('valor_devido_raw', ''))
        valor_devido_col = col_map.get('valor_devido', col_map.get('valor_devido_raw', ''))

        records.append({
            'nome_escola': nome,
            'municipio': str(row.get(col_map.get('municipio', ''), '')).strip(),
            'uf': str(row.get(col_map.get('uf', ''), '')).strip(),
            'rede': str(row.get(col_map.get('rede', ''), '')).strip(),
            'cod_escola': str(row.get(col_map.get('cod_escola', ''), '')).strip(),
            'destinacao': str(dest).strip(),
            'categoria': cat,
            'valor_pago': to_float(row.get(valor_pago_col, 0)),
            'valor_devido': to_float(row.get(valor_devido_col, 0)),
        })

    return records


def aggregate_schools(records):
    schools = {}
    for r in records:
        key = r['nome_escola']
        if key not in schools:
            schools[key] = {
                'nome_escola': r['nome_escola'],
                'municipio': r['municipio'],
                'uf': r['uf'],
                'rede': r['rede'],
                'cod_escola': r['cod_escola'],
                'pdde_basico': 0.0,
                'pdde_equidade': 0.0,
                'pdde_qualidade': 0.0,
                'total': 0.0,
                'programas': [],
            }
        val = r['valor_pago']
        schools[key]['total'] += val
        schools[key]['programas'].append({'programa': r['destinacao'], 'valor': val})
        if r['categoria'] == 'PDDE Básico':
            schools[key]['pdde_basico'] += val
        elif r['categoria'] == 'PDDE EQUIDADE':
            schools[key]['pdde_equidade'] += val
        elif r['categoria'] == 'PDDE QUALIDADE':
            schools[key]['pdde_qualidade'] += val

    return list(schools.values())


# ── Index (files registry) ────────────────────────────────────────────────────

def load_index():
    if not os.path.exists(FILES_INDEX):
        return {}
    with open(FILES_INDEX, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_index(index):
    with open(FILES_INDEX, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def get_year(filename, meta):
    """Return the year for a file entry — from metadata or re-detected from filename."""
    return meta.get('year') or detect_year(filename)


def load_all_schools(years_filter=None):
    """Merge schools from every registered file, optionally filtered by year list."""
    index = load_index()
    merged = {}
    for filename, meta in index.items():
        # Apply year filter when provided
        if years_filter is not None:
            file_year = get_year(filename, meta)
            if file_year not in years_filter:
                continue
        path = schools_path(filename)
        if not os.path.exists(path):
            continue
        with open(path, 'r', encoding='utf-8') as f:
            file_schools = json.load(f)
        for s in file_schools:
            key = s['nome_escola']
            if key not in merged:
                merged[key] = {**s, 'pdde_basico': 0.0, 'pdde_equidade': 0.0,
                               'pdde_qualidade': 0.0, 'total': 0.0, 'programas': []}
            merged[key]['pdde_basico']   += s['pdde_basico']
            merged[key]['pdde_equidade'] += s['pdde_equidade']
            merged[key]['pdde_qualidade'] += s['pdde_qualidade']
            merged[key]['total']         += s['total']
            merged[key]['programas'].extend(s.get('programas', []))
    return list(merged.values())


# ── Summary ───────────────────────────────────────────────────────────────────

def build_summary(schools):
    total_pdde    = sum(s['pdde_basico']    for s in schools)
    total_eq      = sum(s['pdde_equidade']  for s in schools)
    total_qual    = sum(s['pdde_qualidade'] for s in schools)
    total_geral   = sum(s['total']          for s in schools)
    n             = len(schools)
    top5 = sorted(schools, key=lambda s: s['total'], reverse=True)[:5]

    municipios = {}
    for s in schools:
        m = s['municipio']
        if m and m != 'nan':
            municipios[m] = municipios.get(m, 0) + s['total']
    top_mun = sorted(municipios.items(), key=lambda x: x[1], reverse=True)[:10]

    return {
        'total_schools':          n,
        'total_pdde_basico':      total_pdde,
        'total_pdde_equidade':    total_eq,
        'total_pdde_qualidade':   total_qual,
        'total_geral':            total_geral,
        'schools_with_basico':    sum(1 for s in schools if s['pdde_basico']    > 0),
        'schools_with_equidade':  sum(1 for s in schools if s['pdde_equidade']  > 0),
        'schools_with_qualidade': sum(1 for s in schools if s['pdde_qualidade'] > 0),
        'top5_schools':           [{'nome': s['nome_escola'], 'total': s['total']} for s in top5],
        'top_municipios':         [{'municipio': m, 'total': v} for m, v in top_mun],
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    file_index = load_index()
    return render_template('index.html', file_index=file_index)


@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('file')
    if not files or all(f.filename == '' for f in files):
        return redirect(url_for('index'))

    index = load_index()
    errors = []

    for file in files:
        if not file or file.filename == '':
            continue
        if not allowed_file(file.filename):
            errors.append(f'{file.filename}: formato inválido (use .xls ou .xlsx)')
            continue

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        try:
            records  = parse_xls(filepath)
            schools  = aggregate_schools(records)

            # Persist schools for this file
            with open(schools_path(file.filename), 'w', encoding='utf-8') as f:
                json.dump(schools, f, ensure_ascii=False)

            year = detect_year(file.filename, records)
            index[file.filename] = {
                'uploaded_at':  datetime.now().isoformat(),
                'record_count': len(records),
                'school_count': len(schools),
                'year':         year,
            }
        except Exception as e:
            errors.append(f'{file.filename}: {e}')

    save_index(index)

    if errors:
        return jsonify({'errors': errors}), 400

    return redirect(url_for('dashboard'))


@app.route('/files/<filename>/delete', methods=['POST'])
def delete_file(filename):
    index = load_index()
    if filename in index:
        del index[filename]
        save_index(index)
        path = schools_path(filename)
        if os.path.exists(path):
            os.remove(path)
    return redirect(url_for('index'))


@app.route('/dashboard')
def dashboard():
    index = load_index()
    if not index:
        return redirect(url_for('index'))
    schools   = load_all_schools()
    summary   = build_summary(schools)
    all_years = sorted({get_year(fn, m) for fn, m in index.items() if get_year(fn, m)})
    return render_template('dashboard.html', summary=summary, file_index=index,
                           all_years=all_years)


def parse_years_param():
    """Parse ?years=2023,2024 query param into a set of ints, or None for all."""
    raw = request.args.get('years', '').strip()
    if not raw:
        return None
    try:
        return {int(y) for y in raw.split(',') if y.strip()}
    except ValueError:
        return None


@app.route('/api/data')
def api_data():
    index = load_index()
    if not index:
        return jsonify({'error': 'Nenhum dado carregado'}), 404
    years_filter = parse_years_param()
    schools  = load_all_schools(years_filter)
    summary  = build_summary(schools)
    all_years = sorted({get_year(fn, m) for fn, m in index.items() if get_year(fn, m)})
    return jsonify({'schools': schools, 'summary': summary, 'files': index,
                    'all_years': all_years})


@app.route('/api/schools')
def api_schools():
    years_filter = parse_years_param()
    schools   = load_all_schools(years_filter)
    page      = int(request.args.get('page', 1))
    per_page  = int(request.args.get('per_page', 20))
    search    = request.args.get('search', '').lower()
    if search:
        schools = [s for s in schools
                   if search in s['nome_escola'].lower() or search in s['municipio'].lower()]
    total = len(schools)
    start = (page - 1) * per_page
    return jsonify({'schools': schools[start:start + per_page], 'total': total,
                    'page': page, 'per_page': per_page})


@app.route('/como-usar')
def como_usar():
    file_index = load_index()
    return render_template('como-usar.html', file_index=file_index)


@app.route('/api/files')
def api_files():
    return jsonify(load_index())


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8089))
    app.run(host='0.0.0.0', port=port, debug=False)
