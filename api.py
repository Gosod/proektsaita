#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
API для Фотомеханика Time Tracking
Назначение: обслуживает /api/* запросы от WebApp (index.html).
- /api/init      — данные для инициализации WebApp
- /api/reports   — список отчётов (только для админов)
- /api/report    — создание отчёта за сотрудника (только для админов)
- /api/report/<id> PUT/DELETE — редактирование / удаление
- /api/project   — добавление / удаление проекта
- /api/assign    — назначение проектов пользователю

Отчёты обычных сотрудников пишутся ботом напрямую в Sheets.
Этот API нужен только для AdminPanel и команды /start (init).
"""

import os
import json
import time
import logging
import hashlib
import secrets
import hmac
import calendar

import gspread
import pytz

from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, request, jsonify
from flask_cors import CORS
from google.oauth2.service_account import Credentials

# ══════════════════════════════════════════════════════
# КОНФИГ
# ══════════════════════════════════════════════════════
BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
USERS_FILE       = os.path.join(BASE_DIR, 'users.json')
PROJECTS_FILE    = os.path.join(BASE_DIR, 'projects.json')
USER_PROJECTS_FILE = os.path.join(BASE_DIR, 'user_projects.json')
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
REPORTS_LOCAL_FILE = os.path.join(BASE_DIR, 'reports_local.json')
VACATIONS_FILE   = os.path.join(BASE_DIR, 'vacations.json')
REPORT_FLAGS_FILE = os.path.join(BASE_DIR, 'report_flags.json')

SPREADSHEET_ID   = os.environ.get('SPREADSHEET_ID', '')
SHEET_REPORTS    = 'Отчёты'

# ── ПРИЗНАК ВРЕМЕНИ (report_flags.json) ──
# Админ выставляет при корректировке отчёта сотрудника; используется для
# подсветки ячеек мини-табеля (жёлтый — переработка, синий — отработка невыхода).
TIME_TYPES = {'overtime', 'dayoff'}

ADMIN_IDS = [int(x) for x in os.environ.get('ADMIN_IDS', '699229724,924261386').split(',') if x.strip()]

MSK              = pytz.timezone('Europe/Moscow')
SHEETS_RETRY     = 3
SHEETS_DELAY     = 2

# JWT
JWT_SECRET = os.environ.get('JWT_SECRET', 'phm-secret-change-me-in-production')
JWT_EXPIRE_DAYS = 30

# ══════════════════════════════════════════════════════
# AUTH HELPERS
# ══════════════════════════════════════════════════════

def hash_password(password: str) -> str:
    """SHA-256 + salt хэш пароля."""
    salt = secrets.token_hex(16)
    h = hashlib.sha256((salt + password).encode()).hexdigest()
    return f"{salt}:{h}"

def verify_password(password: str, stored: str) -> bool:
    try:
        salt, h = stored.split(':', 1)
        return hmac.compare_digest(
            hashlib.sha256((salt + password).encode()).hexdigest(), h
        )
    except Exception:
        return False

def make_token(user_id: str, username: str) -> str:
    """Простой JWT-подобный токен без внешних зависимостей."""
    payload = {
        'uid':  user_id,
        'unm':  username,
        'exp':  (datetime.now(MSK) + timedelta(days=JWT_EXPIRE_DAYS)).timestamp()
    }
    data = json.dumps(payload, separators=(',', ':'))
    data_b64 = data.encode().hex()
    sig = hmac.new(JWT_SECRET.encode(), data_b64.encode(), hashlib.sha256).hexdigest()
    return f"{data_b64}.{sig}"

def verify_token(token: str):
    """Проверяет токен. Возвращает payload или None."""
    try:
        data_b64, sig = token.rsplit('.', 1)
        expected = hmac.new(JWT_SECRET.encode(), data_b64.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        payload = json.loads(bytes.fromhex(data_b64).decode())
        if payload['exp'] < datetime.now(MSK).timestamp():
            return None
        return payload
    except Exception:
        return None

def auth_required(f):
    """Декоратор — проверяет Bearer токен."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        auth = request.headers.get('Authorization', '')
        token = auth.replace('Bearer ', '').strip()
        payload = verify_token(token)
        if not payload:
            return jsonify({'error': 'Unauthorized'}), 401
        request.current_user = payload
        return f(*args, **kwargs)
    return wrapper

# ── КОДЫ-ПРИГЛАШЕНИЯ ──────────────────────────────────
# Алфавит без похожих символов: без 0/O и 1/I.
INVITE_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'

def gen_invite_code(length: int = 4) -> str:
    """Персональный одноразовый код (латиница+цифры, без 0/O, 1/I)."""
    return ''.join(secrets.choice(INVITE_ALPHABET) for _ in range(length))

# ── НОРМЫ РАБОЧИХ ЧАСОВ ПО РАСПИСАНИЯМ ───────────────
MONTHLY_NORMS: dict[str, dict[str, int]] = {
    '5/2': {
        '2026-01': 120, '2026-02': 152, '2026-03': 168,
        '2026-04': 176, '2026-05': 152, '2026-06': 168,
        '2026-07': 184, '2026-08': 168, '2026-09': 176,
        '2026-10': 176, '2026-11': 160, '2026-12': 176,
    },
    '2/2A': {
        '2026-01': 144, '2026-02': 156, '2026-03': 180,
        '2026-04': 180, '2026-05': 180, '2026-06': 168,
        '2026-07': 192, '2026-08': 180, '2026-09': 180,
        '2026-10': 192, '2026-11': 168, '2026-12': 180,
    },
    '2/2B': {
        '2026-01': 132, '2026-02': 168, '2026-03': 180,
        '2026-04': 180, '2026-05': 168, '2026-06': 180,
        '2026-07': 180, '2026-08': 192, '2026-09': 180,
        '2026-10': 180, '2026-11': 180, '2026-12': 180,
    },
}
VALID_SCHEDULES = set(MONTHLY_NORMS.keys())

# ── РАБОЧИЕ ДНИ ПО ГРАФИКАМ (для подсветки выходных в табеле) ──
# Дни месяца, когда сотрудник по графику работает (1-индексация).
# Все остальные дни месяца — выходные/праздники для этого графика.
SCHEDULE_WORKDAYS: dict[str, dict[str, list[int]]] = {
    '5/2': {
        '2026-01': [12, 13, 14, 15, 16, 19, 20, 21, 22, 23, 26, 27, 28, 29, 30],
        '2026-02': [2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 24, 25, 26, 27],
        '2026-03': [2, 3, 4, 5, 6, 10, 11, 12, 13, 16, 17, 18, 19, 20, 23, 24, 25, 26, 27, 30, 31],
        '2026-04': [1, 2, 3, 6, 7, 8, 9, 10, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 27, 28, 29, 30],
        '2026-05': [4, 5, 6, 7, 8, 12, 13, 14, 15, 18, 19, 20, 21, 22, 25, 26, 27, 28, 29],
        '2026-06': [1, 2, 3, 4, 5, 8, 9, 10, 11, 15, 16, 17, 18, 19, 22, 23, 24, 25, 26, 29, 30],
        '2026-07': [1, 2, 3, 6, 7, 8, 9, 10, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 27, 28, 29, 30, 31],
        '2026-08': [3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 17, 18, 19, 20, 21, 24, 25, 26, 27, 28, 31],
        '2026-09': [1, 2, 3, 4, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29, 30],
        '2026-10': [1, 2, 5, 6, 7, 8, 9, 12, 13, 14, 15, 16, 19, 20, 21, 22, 23, 26, 27, 28, 29, 30],
        '2026-11': [2, 3, 5, 6, 9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 23, 24, 25, 26, 27, 30],
        '2026-12': [1, 2, 3, 4, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29, 30],
    },
    '2/2A': {
        '2026-01': [10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31],
        '2026-02': [3, 4, 7, 8, 11, 12, 15, 16, 19, 20, 24, 25, 28],
        '2026-03': [1, 4, 5, 9, 10, 13, 14, 17, 18, 21, 22, 25, 26, 29, 30],
        '2026-04': [2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30],
        '2026-05': [2, 5, 6, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31],
        '2026-06': [3, 4, 7, 8, 11, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-07': [2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31],
        '2026-08': [3, 4, 7, 8, 11, 12, 15, 16, 19, 20, 23, 24, 27, 28, 31],
        '2026-09': [1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-10': [2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31],
        '2026-11': [3, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-12': [2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31],
    },
    '2/2B': {
        '2026-01': [9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-02': [1, 2, 5, 6, 9, 10, 13, 14, 17, 18, 21, 22, 26, 27],
        '2026-03': [2, 3, 6, 7, 11, 12, 15, 16, 19, 20, 23, 24, 27, 28, 31],
        '2026-04': [1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-05': [3, 4, 7, 8, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-06': [1, 2, 5, 6, 9, 10, 14, 15, 18, 19, 22, 23, 26, 27, 30],
        '2026-07': [1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-08': [1, 2, 5, 6, 9, 10, 13, 14, 17, 18, 21, 22, 25, 26, 29, 30],
        '2026-09': [2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30],
        '2026-10': [1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
        '2026-11': [1, 2, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30],
        '2026-12': [1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21, 24, 25, 28, 29],
    },
}

def get_monthly_norm(user_record: dict, year_month: str) -> int:
    schedule = user_record.get('schedule', '')
    if schedule in MONTHLY_NORMS:
        return MONTHLY_NORMS[schedule].get(year_month, 160)
    return int(user_record.get('monthly_norm', 160) or 160)

# ══════════════════════════════════════════════════════
# ЛОГИРОВАНИЕ
# ══════════════════════════════════════════════════════
logging.basicConfig(
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    level=logging.INFO,
)
log = logging.getLogger('phm_api')

# ══════════════════════════════════════════════════════
# FLASK APP
# ══════════════════════════════════════════════════════
app = Flask(__name__)
CORS(app)

# ── Security-заголовки (HSTS, CSP и т.п.) ──
_CSP_POLICY = (
    "default-src 'self'; "
    "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
    "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
    "font-src 'self' https://fonts.gstatic.com; "
    "img-src 'self' data:; "
    "connect-src 'self'; "
    "frame-ancestors 'self'; "
    "base-uri 'self'; "
    "form-action 'self'"
)

@app.after_request
def _security_headers(response):
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    response.headers['Content-Security-Policy'] = _CSP_POLICY
    return response

# ── Локальная отдача статики (index.html / manifest.json) ──
from flask import send_from_directory

@app.route('/')
@app.route('/app.html')
def _serve_index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/manifest.json')
def _serve_manifest():
    return send_from_directory(BASE_DIR, 'manifest.json')

@app.route('/.well-known/security.txt')
def _serve_security_txt():
    return send_from_directory(BASE_DIR, 'security.txt')

# Иконки и фон — если файла нет, тихий 204 вместо 404 в логах
@app.route('/favicon.ico')
@app.route('/bg.jpg')
@app.route('/icon-192.png')
@app.route('/icon-512.png')
def _serve_optional_static():
    fname = request.path.lstrip('/')
    fpath = os.path.join(BASE_DIR, fname)
    if os.path.exists(fpath):
        return send_from_directory(BASE_DIR, fname)
    return '', 204

# ══════════════════════════════════════════════════════
# JSON HELPERS (атомарная запись)
# ══════════════════════════════════════════════════════
def load_json(path: str, default):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        log.error(f"load_json({path}): {e}")
    return default


def save_json(path: str, data) -> bool:
    tmp = path + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return True
    except Exception as e:
        log.error(f"save_json({path}): {e}")
        return False

# ══════════════════════════════════════════════════════
# GOOGLE SHEETS
# ══════════════════════════════════════════════════════
def _sheets_client():
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def _hours_for_sheets(hours) -> str:
    """
    Форматирует часы для записи в Sheets с valueInputOption=USER_ENTERED.
    Таблица настроена на русскую локаль (запятая — десятичный разделитель,
    точка — разделитель тысяч), поэтому "2.1" Sheets превращает в 21.
    Заменяем точку на запятую: "2.1" -> "2,1" -> корректно 2.1.
    """
    return str(hours).replace('.', ',')


def sheets_append(report: dict) -> bool:
    """Добавить строку в Google Sheets с retry."""
    dt_str = report.get('datetime', '')
    try:
        dt = datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
        date_str = dt.strftime('%d.%m.%Y')
        time_str = dt.strftime('%H:%M:%S')
    except Exception:
        date_str = report.get('date', '')
        time_str = ''

    row = [
        date_str,
        time_str,
        report.get('username', ''),
        report.get('project', ''),
        _hours_for_sheets(report.get('hours', 0)),
        report.get('comments', ''),
    ]

    for attempt in range(1, SHEETS_RETRY + 1):
        try:
            client = _sheets_client()
            sheet  = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_REPORTS)
            sheet.append_row(row, value_input_option='USER_ENTERED')
            log.info(f"SHEETS_OK | attempt={attempt} | project={report.get('project')} | user={report.get('username')}")
            return True
        except Exception as e:
            log.warning(f"SHEETS_FAIL | attempt={attempt}/{SHEETS_RETRY} | error={e}")
            if attempt < SHEETS_RETRY:
                time.sleep(SHEETS_DELAY)
    return False


def sheets_read_all() -> list:
    """Прочитать все строки из листа Отчёты."""
    try:
        client  = _sheets_client()
        sheet   = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_REPORTS)
        # Колонка 5 (Часы) — не даём gspread численно интерпретировать значение,
        # т.к. он считает запятую разделителем тысяч ("5,9" -> 59).
        # Парсинг с учётом запятой как десятичного разделителя — в _normalize_sheets_records.
        records = sheet.get_all_records(numericise_ignore=[5])
        log.info(f"SHEETS_READ | rows={len(records)}")
        return records
    except Exception as e:
        log.error(f"SHEETS_READ_FAIL | error={e}")
        return []


def local_reports_for_user(uid_int: int, username: str) -> list:
    """Отчёты тестовых сотрудников, сохранённые локально."""
    uname = username.lower()
    return [
        r for r in load_json(REPORTS_LOCAL_FILE, [])
        if r.get('user_id') == uid_int or r.get('username', '').lower() == uname
    ]


def append_local_report(report: dict) -> None:
    """Добавить отчёт в локальный файл (для тест-сотрудников)."""
    records = load_json(REPORTS_LOCAL_FILE, [])
    records.append(report)
    save_json(REPORTS_LOCAL_FILE, records)

# ══════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════
def is_admin(user_id) -> bool:
    try:
        return int(user_id) in ADMIN_IDS
    except (TypeError, ValueError):
        return False


def msk_now() -> datetime:
    return datetime.now(MSK)


def gen_id() -> str:
    return msk_now().strftime('%Y%m%d%H%M%S%f')


def get_user_projects(user_id: int) -> list:
    all_projects  = load_json(PROJECTS_FILE, [])
    user_proj_map = load_json(USER_PROJECTS_FILE, {})
    abbrs         = user_proj_map.get(str(user_id))
    if not abbrs:
        return all_projects
    filtered = [p for p in all_projects if p['abbr'] in abbrs]
    return filtered or all_projects


def build_user_stats(user_id: int, all_reports: list, all_users: dict = None) -> dict:
    if all_users is None:
        all_users = load_json(USERS_FILE, {})
    username  = all_users.get(str(user_id), {}).get('username', '').lower()

    user_reports = [
        r for r in all_reports
        if r.get('user_id') == user_id
        or (username and r.get('username', '').lower() == username)
    ]

    # Часы за текущий календарный месяц (по московскому времени)
    month_prefix = msk_now().strftime('%Y-%m')
    month_hours  = round(sum(
        r.get('hours', 0) for r in user_reports
        if r.get('date', '').startswith(month_prefix)
    ), 2)

    by_project = {}
    for r in user_reports:
        proj = r.get('project', '?')
        by_project[proj] = round(by_project.get(proj, 0) + r.get('hours', 0), 2)
    months_with_reports = sorted(
        {r.get('date', '')[:7] for r in user_reports if len(r.get('date', '')) >= 7},
        reverse=True,
    )
    return {
        'total_hours':         round(sum(r.get('hours', 0) for r in user_reports), 2),
        'total_reports':       len(user_reports),
        'by_project':          by_project,
        'month_hours':         month_hours,
        'months_with_reports': months_with_reports,
    }


def build_admin_stats(all_reports: list, all_projects: list) -> dict:
    employees  = {}
    proj_stats = {}

    for r in all_reports:
        uid   = r.get('user_id')
        uname = r.get('username', '?')
        proj  = r.get('project', '?')
        hours = r.get('hours', 0)

        if uid not in employees:
            employees[uid] = {'username': uname, 'hours': 0, 'reports': 0, 'projects': {}}
        employees[uid]['hours']   += hours
        employees[uid]['reports'] += 1
        employees[uid]['projects'][proj] = employees[uid]['projects'].get(proj, 0) + hours
        proj_stats[proj] = proj_stats.get(proj, 0) + hours

    recent = sorted(all_reports, key=lambda x: x.get('datetime', ''), reverse=True)[:50]
    unique_projects = sorted(
        {r.get('project', '') for r in all_reports if r.get('project')},
        key=lambda x: x.lower()
    )

    return {
        'total_hours':           sum(r.get('hours', 0) for r in all_reports),
        'total_reports':         len(all_reports),
        'employees':             list(employees.values()),
        'projects':              proj_stats,
        'recent_reports':        recent,
        'unique_report_projects': unique_projects,
    }


# ══════════════════════════════════════════════════════
# ДЕКОРАТОР: требует подписанный токен админа
# ══════════════════════════════════════════════════════
def admin_token_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        auth = request.headers.get('Authorization', '')
        token = auth.replace('Bearer ', '').strip()
        payload = verify_token(token)
        if not payload or not is_admin(payload.get('uid')):
            log.warning(f"ADMIN_FORBIDDEN | endpoint={request.path}")
            return jsonify({'error': 'Forbidden'}), 403
        request.current_user = payload
        return f(*args, **kwargs)
    return wrapper


# ══════════════════════════════════════════════════════
# ENDPOINTS
# ══════════════════════════════════════════════════════

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        'status':    'ok',
        'version':   '4.1',
        'timestamp': msk_now().isoformat(),
    })


# ══════════════════════════════════════════════════════
# AUTH ENDPOINTS
# ══════════════════════════════════════════════════════

@app.route('/api/auth/users', methods=['GET'])
def get_user_list():
    """Список сотрудников для выпадающего списка при регистрации."""
    show_test = request.args.get('show_test') == '1'
    users = load_json(USERS_FILE, {})
    result = []
    for uid, udata in users.items():
        if udata.get('is_test') and not show_test:
            continue
        result.append({
            'id':           uid,
            'username':     udata.get('username', ''),
            'display_name': udata.get('display_name', udata.get('username', '')),
            'has_password': bool(udata.get('password_hash')),
        })
    result.sort(key=lambda x: x['display_name'])
    return jsonify({'users': result})


@app.route('/api/auth/register', methods=['POST'])
def register():
    """Первичная установка пароля сотрудником."""
    data     = request.get_json(silent=True) or {}
    uid      = str(data.get('user_id', ''))
    password = data.get('password', '').strip()

    if not uid or not password:
        return jsonify({'error': 'user_id и password обязательны'}), 400
    if len(password) < 6:
        return jsonify({'error': 'Пароль минимум 6 символов'}), 400

    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404

    if users[uid].get('password_hash'):
        return jsonify({'error': 'Пароль уже установлен'}), 400

    # ── Персональный код-приглашение (защита от регистрации за чужой аккаунт) ──
    code        = str(data.get('invite_code', '')).strip()
    stored_code = users[uid].get('invite_code')
    invite_used = bool(users[uid].get('invite_used', False))
    if not code:
        return jsonify({'error': 'Введите код-приглашение'}), 400
    if not stored_code or invite_used or code.upper() != str(stored_code).upper():
        log.warning(f"AUTH_REGISTER_BADCODE | uid={uid}")
        return jsonify({'error': 'Неверный или уже использованный код-приглашение'}), 403

    users[uid]['password_hash'] = hash_password(password)
    users[uid]['status'] = 'active'
    users[uid]['invite_used'] = True   # код сгорел — повторно не сработает
    save_json(USERS_FILE, users)

    username = users[uid].get('username', '')
    token    = make_token(uid, username)

    log.info(f"AUTH_REGISTER | uid={uid} | username={username}")
    return jsonify({
        'success':      True,
        'token':        token,
        'user_id':      int(uid),
        'username':     username,
        'display_name': users[uid].get('display_name', username),
        'admin':        is_admin(uid),
    })


@app.route('/api/auth/login', methods=['POST'])
def login():
    """Вход по логину + паролю."""
    data     = request.get_json(silent=True) or {}
    uid      = str(data.get('user_id', ''))
    password = data.get('password', '').strip()

    if not uid or not password:
        return jsonify({'error': 'user_id и password обязательны'}), 400

    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404

    stored = users[uid].get('password_hash')
    if not stored:
        return jsonify({'error': 'Пароль не установлен — зарегистрируйтесь'}), 400

    if not verify_password(password, stored):
        log.warning(f"AUTH_FAIL | uid={uid}")
        return jsonify({'error': 'Неверный пароль'}), 401

    username = users[uid].get('username', '')
    token    = make_token(uid, username)

    log.info(f"AUTH_LOGIN | uid={uid} | username={username}")
    return jsonify({
        'success':      True,
        'token':        token,
        'user_id':      int(uid),
        'username':     username,
        'display_name': users[uid].get('display_name', username),
        'admin':        is_admin(uid),
    })


# ── INIT ─────────────────────────────────────────────
@app.route('/api/init', methods=['POST'])
@auth_required
def init_data():
    user_id = request.current_user['uid']

    all_users    = load_json(USERS_FILE, {})
    all_projects = load_json(PROJECTS_FILE, [])
    user_projects = get_user_projects(int(user_id))

    # Читаем отчёты из Sheets для всех (тест-аккаунты исключены)
    all_reports = sheets_read_all()
    all_reports = _normalize_sheets_records(all_reports, all_users)

    # Для личной статистики добавляем локальные отчёты (тест-сотрудники)
    _urec_pre = all_users.get(str(user_id), {})
    local_extra = (
        local_reports_for_user(int(user_id), _urec_pre.get('username', ''))
        if _urec_pre.get('is_test') else []
    )
    user_stats = build_user_stats(int(user_id), all_reports + local_extra, all_users)

    _urec = all_users.get(str(user_id), {})
    resp = {
        'admin':        is_admin(user_id),
        'user_id':      user_id,
        'username':     _urec.get('username', ''),
        'display_name': _urec.get('display_name', _urec.get('username', '')),
        'projects':     user_projects,
        'user_stats':   user_stats,
        'monthly_norm': get_monthly_norm(_urec, msk_now().strftime('%Y-%m')),
        'schedule':     _urec.get('schedule', ''),
    }

    if is_admin(user_id):
        resp['all_projects'] = all_projects
        # Тестовых сотрудников исключаем из обычных админских списков
        resp['all_users']    = [
            {'id': int(uid), 'username': udata.get('username', '?')}
            for uid, udata in all_users.items()
            if not udata.get('is_test')
        ]
        resp['admin_stats']      = build_admin_stats(all_reports, all_projects)
        resp['user_assignments'] = load_json(USER_PROJECTS_FILE, {})

    log.info(f"INIT | user_id={user_id} | admin={resp['admin']}")
    response = jsonify(resp)
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    return response


def _normalize_sheets_records(records: list, all_users: dict) -> list:
    """
    Преобразует строки из gspread (dict по заголовкам) в наш внутренний формат.
    Ожидаемые заголовки: Дата, Время, Сотрудник, Проект, Часы, Комментарий
    """
    normalized = []
    # Строим обратный словарь username → user_id для обогащения
    uname_to_id = {v.get('username', '').lower(): int(k) for k, v in all_users.items()}
    # Признаки времени (переработка/отработка невыхода), выставленные админом
    # при корректировке отчёта — report_flags.json: {uid: {date: {project: type}}}
    report_flags = load_json(REPORT_FLAGS_FILE, {})
    # Ники тестовых сотрудников — их строки исключаем из статистики/отчётов
    test_unames = {
        str(v.get('username', '')).lower()
        for v in all_users.values() if v.get('is_test')
    }

    for i, row in enumerate(records):
        try:
            date_raw = str(row.get('Дата', '') or row.get('date', '')).strip()
            time_raw = str(row.get('Время', '') or row.get('time', '')).strip()
            username = str(row.get('Сотрудник', '') or row.get('username', '')).strip()
            if username.lower() in test_unames:
                continue  # тестовый сотрудник — пропускаем
            # Сотрудник, которого нет в users.json (опечатка / тестовая запись
            # вроде «ГЛЕБ КРУПСКИЙ ОЧЕНЬ ОТВЕТСТВЕННЫЙ!»), либо пустое имя —
            # не должен попадать в статистику как фантомный сотрудник («?»).
            # Пропускаем и логируем.
            if username.lower() not in uname_to_id:
                log.info(f"SHEETS_SKIP_UNKNOWN | row={i} | user='{username}' нет в users.json — пропущен")
                continue
            project  = str(row.get('Проект', '')  or row.get('project', '')).strip()
            # Парсим часы — учитываем разные форматы: 7.25, 7,25, 725 (ошибка)
            raw_hours = str(row.get('Часы', 0) or row.get('hours', 0)).strip()
            try:
                if ',' in raw_hours and '.' not in raw_hours:
                    # Запятая как десятичный разделитель: "7,25" → 7.25
                    hours = float(raw_hours.replace(',', '.'))
                else:
                    hours = float(raw_hours or 0)
                # Санитарная проверка — не может быть больше 24ч за запись
                if hours > 24:
                    log.warning(f"Подозрительное значение часов: {hours} (raw='{raw_hours}'), строка {i}")
                    hours = 0.0
            except (ValueError, TypeError):
                hours = 0.0
            comments = str(row.get('Комментарий', '') or row.get('comments', '')).strip()

            # Дата: dd.mm.yyyy → yyyy-mm-dd
            try:
                dt = datetime.strptime(date_raw, '%d.%m.%Y')
                date_iso = dt.strftime('%Y-%m-%d')
                datetime_str = f"{date_iso} {time_raw}" if time_raw else f"{date_iso} 00:00:00"
            except Exception:
                date_iso     = date_raw
                datetime_str = f"{date_raw} {time_raw}".strip()

            uid = uname_to_id.get(username.lower(), 0)
            time_type = report_flags.get(str(uid), {}).get(date_iso, {}).get(project, '')

            normalized.append({
                'id':        f'sheets_{i}',
                'user_id':   uid,
                'username':  username,
                'project':   project,
                'hours':     hours,
                'comments':  comments,
                'date':      date_iso,
                'datetime':  datetime_str,
                'time_type': time_type,
            })
        except Exception as e:
            log.warning(f"Пропуск строки Sheets row={i}: {e}")

    return normalized


# ── CREATE REPORT (только для Admin — за сотрудника) ──
@app.route('/api/report/submit', methods=['POST'])
@auth_required
def submit_report_pwa():
    """Отправка отчёта из PWA напрямую в Sheets."""
    data    = request.get_json(silent=True) or {}
    uid     = request.current_user['uid']
    users   = load_json(USERS_FILE, {})
    user_rec = users.get(str(uid), {})
    # Берём никнейм из users.json — не из payload
    username = user_rec.get('username', data.get('username', '?'))
    is_test_user = bool(user_rec.get('is_test', False))

    projects    = data.get('projects', [])
    general_cmt = data.get('comments', '-')
    custom_date = data.get('custom_date')

    if not projects:
        return jsonify({'error': 'projects required'}), 400

    if custom_date:
        try:
            dt = MSK.localize(datetime.strptime(custom_date, '%Y-%m-%d'))
        except Exception:
            dt = msk_now()
    else:
        dt = msk_now()

    # Конфликт: день отмечен как отпуск → отчёт сдавать нельзя
    date_str = dt.strftime('%Y-%m-%d')
    user_vac = load_json(VACATIONS_FILE, {}).get(str(uid), {})
    if date_str in user_vac:
        return jsonify({
            'error':    f'День {date_str} отмечен как отпуск. Обратитесь к админу, чтобы снять отметку.',
            'conflict': 'vacation',
            'date':     date_str,
        }), 409

    errors = 0
    saved  = []
    for item in projects:
        proj_cmt = item.get('comment', '').strip()
        report = {
            'user_id':  int(uid),
            'username': username,
            'project':  item.get('project', '?'),
            'hours':    float(item.get('hours', 0)),
            'comments': proj_cmt or general_cmt,
            'date':     dt.strftime('%Y-%m-%d'),
            'datetime': dt.strftime('%Y-%m-%d %H:%M:%S'),
        }
        # Тестовый сотрудник — сохраняем локально, не в Sheets
        if is_test_user:
            saved.append(report)
            append_local_report({**report, 'user_id': int(uid)})
            log.info(f"PWA_REPORT_TEST | uid={uid} | project={report['project']} | (локально сохранён)")
            continue
        ok = sheets_append(report)
        if not ok:
            errors += 1
        saved.append(report)
        log.info(f"PWA_REPORT | uid={uid} | username={username} | project={report['project']} | hours={report['hours']}")

    return jsonify({'success': True, 'saved': len(saved), 'sheets_errors': errors})


# ── CREATE REPORT за сотрудника (только Admin) ────────
@app.route('/api/report', methods=['POST'])
@admin_token_required
def admin_create_report():
    """Создание отчёта администратором за выбранного сотрудника."""
    data       = request.get_json(silent=True) or {}
    target_uid = str(data.get('on_behalf_of_user_id', '')).strip()
    if not target_uid:
        return jsonify({'error': 'on_behalf_of_user_id required'}), 400

    users    = load_json(USERS_FILE, {})
    user_rec = users.get(target_uid)
    if user_rec is None:
        return jsonify({'error': 'Пользователь не найден'}), 404

    username     = user_rec.get('username', '').strip()
    is_test_user = bool(user_rec.get('is_test', False))
    # Без ника строка в Sheets не будет привязана к сотруднику
    # (_normalize_sheets_records пропускает записи с неизвестным/пустым ником),
    # поэтому такой отчёт «молча исчез» бы при чтении — отклоняем сразу.
    if not username:
        return jsonify({'error': 'У сотрудника не задан ник — отчёт нельзя создать'}), 400

    projects    = data.get('projects', [])
    general_cmt = data.get('comments', '-')
    custom_date = data.get('custom_date')
    if not projects:
        return jsonify({'error': 'projects required'}), 400

    if custom_date:
        try:
            dt = MSK.localize(datetime.strptime(custom_date, '%Y-%m-%d'))
        except Exception:
            dt = msk_now()
    else:
        dt = msk_now()

    # Конфликт: день отмечен как отпуск → отчёт сдавать нельзя
    date_str = dt.strftime('%Y-%m-%d')
    user_vac = load_json(VACATIONS_FILE, {}).get(target_uid, {})
    if date_str in user_vac:
        return jsonify({
            'error':    f'День {date_str} отмечен как отпуск. Снимите отметку, чтобы добавить отчёт.',
            'conflict': 'vacation',
            'date':     date_str,
        }), 409

    errors = 0
    saved  = []
    for item in projects:
        proj_cmt = str(item.get('comment', '')).strip()
        try:
            hours = float(item.get('hours', 0))
        except (ValueError, TypeError):
            hours = 0.0
        report = {
            'user_id':  int(target_uid),
            'username': username,
            'project':  item.get('project', '?'),
            'hours':    hours,
            'comments': proj_cmt or general_cmt,
            'date':     date_str,
            'datetime': dt.strftime('%Y-%m-%d %H:%M:%S'),
        }
        # Тестовый сотрудник — сохраняем локально, не в Sheets
        if is_test_user:
            append_local_report(report)
            saved.append(report)
            continue
        if not sheets_append(report):
            errors += 1
        saved.append(report)

    log.info(f"ADMIN_REPORT_CREATE | by={request.current_user.get('uid')} | for={target_uid} | n={len(saved)} | errors={errors}")
    return jsonify({'success': True, 'saved': len(saved), 'sheets_errors': errors})


# ── GET REPORTS ────────────────────────────────────────
@app.route('/api/user/timesheet', methods=['GET'])
@auth_required
def user_timesheet():
    """Данные табеля для сотрудника за указанный месяц."""
    uid = request.current_user['uid']
    try:
        year  = int(request.args.get('year',  0))
        month = int(request.args.get('month', 0))
        if not year or not (1 <= month <= 12):
            raise ValueError()
    except (ValueError, TypeError):
        return jsonify({'error': 'year и month обязательны'}), 400

    all_users   = load_json(USERS_FILE, {})
    all_reports = sheets_read_all()
    all_reports = _normalize_sheets_records(all_reports, all_users)

    urec  = all_users.get(uid, {})
    uname = urec.get('username', '').lower()
    prefix = f'{year:04d}-{month:02d}'

    local_extra = (
        local_reports_for_user(int(uid), uname)
        if urec.get('is_test') else []
    )
    combined = all_reports + local_extra

    user_reports = [
        r for r in combined
        if (r.get('user_id') == int(uid) or (uname and r.get('username', '').lower() == uname))
        and r.get('date', '').startswith(prefix)
    ]

    days_in_month = calendar.monthrange(year, month)[1]
    rows: dict = {}
    cell_flags: dict = {}
    projects_set: set = set()
    for r in user_reports:
        date_str = r.get('date', '')
        try:
            day = int(date_str[8:10])
        except (ValueError, IndexError):
            continue
        proj  = r.get('project', '?')
        hours = float(r.get('hours', 0))
        projects_set.add(proj)
        rows.setdefault(str(day), {})
        rows[str(day)][proj] = round(rows[str(day)].get(proj, 0) + hours, 2)
        # Подсветка ячейки (переработка/отработка невыхода), выставляется
        # админом при корректировке отчёта — см. TIME_TYPE_LABELS
        if r.get('time_type'):
            cell_flags.setdefault(str(day), {})[proj] = r['time_type']

    # Отпускные дни за этот месяц (только даты, без часов — для отметки в табеле)
    vacations     = load_json(VACATIONS_FILE, {})
    vacation_days = sorted(
        d[8:10].lstrip('0') or '0'   # "01" → "1"
        for d in vacations.get(uid, {})
        if d.startswith(prefix)
    )

    # Выходные/праздники по графику сотрудника (для подсветки в табеле).
    # Если для графика и месяца данных нет — фронтенд подсветит сб/вс сам.
    workdays = SCHEDULE_WORKDAYS.get(urec.get('schedule', ''), {}).get(prefix)
    off_days = (
        [d for d in range(1, days_in_month + 1) if d not in workdays]
        if workdays is not None else None
    )

    return jsonify({
        'year':          year,
        'month':         month,
        'days_in_month': days_in_month,
        'norm':          get_monthly_norm(urec, prefix),
        'projects':      sorted(projects_set),
        'rows':          rows,
        'cell_flags':    cell_flags,       # {"5": {"Озон Кемерово": "overtime"}, ...}
        'vacation_days': vacation_days,   # ["3","10","11"] — просто список дней
        'off_days':      off_days,        # [1,2,9,...] или null (сб/вс по умолчанию)
    })


# ── Отпуска (admin) ───────────────────────────────

@app.route('/api/admin/employee/vacation', methods=['GET'])
@admin_token_required
def admin_get_vacation():
    uid = str(request.args.get('user_id', ''))
    if not uid:
        return jsonify({'error': 'user_id обязателен'}), 400
    vacations = load_json(VACATIONS_FILE, {})
    return jsonify({'vacation': vacations.get(uid, {})})


@app.route('/api/admin/employee/vacation', methods=['POST'])
@admin_token_required
def admin_set_vacation():
    data   = request.get_json(silent=True) or {}
    uid    = str(data.get('user_id', ''))
    d_from = str(data.get('date_from', ''))
    d_to   = str(data.get('date_to',   ''))
    try:
        cur = datetime.strptime(d_from, '%Y-%m-%d').date()
        end = datetime.strptime(d_to,   '%Y-%m-%d').date()
        if end < cur:
            raise ValueError()
    except Exception:
        return jsonify({'error': 'Некорректные даты'}), 400
    all_users = load_json(USERS_FILE, {})
    if uid not in all_users:
        return jsonify({'error': 'Пользователь не найден'}), 404

    # Конфликт: если в диапазоне у сотрудника уже есть отчёты — блокируем
    urec  = all_users[uid]
    uname = urec.get('username', '').lower()
    range_dates = set()
    _cur = cur
    while _cur <= end:
        range_dates.add(_cur.strftime('%Y-%m-%d'))
        _cur += timedelta(days=1)
    combined = _normalize_sheets_records(sheets_read_all(), all_users)
    if urec.get('is_test'):
        combined = combined + local_reports_for_user(int(uid), uname)
    conflict_dates = sorted({
        r.get('date', '') for r in combined
        if (r.get('user_id') == int(uid) or (uname and r.get('username', '').lower() == uname))
        and r.get('date', '') in range_dates
    })
    if conflict_dates:
        return jsonify({
            'error':    f'У сотрудника есть отчёты за: {", ".join(conflict_dates)}. Сначала удалите их.',
            'conflict': 'reports',
            'dates':    conflict_dates,
        }), 409

    vacations = load_json(VACATIONS_FILE, {})
    user_vac  = vacations.get(uid, {})
    added = 0
    while cur <= end:
        user_vac[cur.strftime('%Y-%m-%d')] = True
        cur += timedelta(days=1)
        added += 1
    vacations[uid] = user_vac
    if not save_json(VACATIONS_FILE, vacations):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    log.info(f"ADMIN_VAC_ADD | by={request.current_user.get('uid')} | uid={uid} | from={d_from} to={d_to} | days={added}")
    return jsonify({'success': True, 'added': added})


@app.route('/api/admin/employee/vacation', methods=['DELETE'])
@admin_token_required
def admin_del_vacation():
    data = request.get_json(silent=True) or {}
    uid  = str(data.get('user_id', ''))
    date = str(data.get('date', ''))
    vacations = load_json(VACATIONS_FILE, {})
    if uid in vacations and date in vacations[uid]:
        del vacations[uid][date]
        save_json(VACATIONS_FILE, vacations)
        log.info(f"ADMIN_VAC_DEL | by={request.current_user.get('uid')} | uid={uid} | date={date}")
    return jsonify({'success': True})


@app.route('/api/reports', methods=['GET'])
@admin_token_required
def get_reports():
    filter_user    = request.args.get('user_id')
    filter_project = request.args.get('project', '').lower()
    filter_date    = request.args.get('date')

    all_users = load_json(USERS_FILE, {})
    reports   = sheets_read_all()
    reports   = _normalize_sheets_records(reports, all_users)

    if filter_user:
        reports = [r for r in reports if str(r.get('user_id')) == str(filter_user)]
    if filter_project:
        reports = [r for r in reports if filter_project in r.get('project', '').lower()]
    if filter_date:
        reports = [r for r in reports if r.get('date') == filter_date]

    reports = sorted(reports, key=lambda x: x.get('datetime', ''), reverse=True)
    return jsonify({'reports': reports, 'total': len(reports)})


# ── UPDATE REPORT (редактирование строки в Sheets) ────
@app.route('/api/report/<report_id>', methods=['PUT'])
@admin_token_required
def update_report(report_id):
    """
    Редактирование строки в Sheets по индексу.
    report_id формата 'sheets_N' где N — 0-based индекс в массиве записей.
    """
    data = request.get_json(silent=True) or {}

    if not report_id.startswith('sheets_'):
        return jsonify({'error': 'Invalid report_id format'}), 400

    try:
        row_idx = int(report_id.split('_')[1])  # 0-based
    except (IndexError, ValueError):
        return jsonify({'error': 'Invalid report_id'}), 400

    try:
        client  = _sheets_client()
        sheet   = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_REPORTS)
        records = sheet.get_all_values()  # включая заголовок

        # row_idx — индекс в нормализованном массиве (без заголовка)
        # В Sheets: строка 1 = заголовок, строка 2 = первая запись
        sheet_row = row_idx + 2  # 1-based + заголовок

        if sheet_row > len(records):
            return jsonify({'error': 'Row not found'}), 404

        # Обновляем нужные ячейки
        if 'date' in data:
            try:
                dt = datetime.strptime(data['date'], '%Y-%m-%d')
                sheet.update_cell(sheet_row, 1, dt.strftime('%d.%m.%Y'))
            except Exception:
                pass
        if 'hours' in data:
            sheet.update_cell(sheet_row, 5, _hours_for_sheets(float(data['hours'])))
        if 'project' in data:
            sheet.update_cell(sheet_row, 4, data['project'])
        if 'comments' in data:
            sheet.update_cell(sheet_row, 6, data['comments'])
        if 'time_type' in data:
            time_type = data['time_type'] if data['time_type'] in TIME_TYPES else ''
            row_vals = records[sheet_row - 1]
            username = row_vals[2].strip() if len(row_vals) > 2 else ''
            all_users   = load_json(USERS_FILE, {})
            uname_to_id = {v.get('username', '').lower(): int(k) for k, v in all_users.items()}
            uid = uname_to_id.get(username.lower())
            if uid is not None:
                if 'date' in data:
                    date_iso = data['date']
                else:
                    try:
                        date_iso = datetime.strptime(row_vals[0], '%d.%m.%Y').strftime('%Y-%m-%d')
                    except Exception:
                        date_iso = row_vals[0]
                project = data.get('project', row_vals[3] if len(row_vals) > 3 else '')

                flags = load_json(REPORT_FLAGS_FILE, {})
                uid_flags  = flags.setdefault(str(uid), {})
                date_flags = uid_flags.setdefault(date_iso, {})
                if time_type:
                    date_flags[project] = time_type
                else:
                    date_flags.pop(project, None)
                    if not date_flags:
                        uid_flags.pop(date_iso, None)
                    if not uid_flags:
                        flags.pop(str(uid), None)
                save_json(REPORT_FLAGS_FILE, flags)

        log.info(f"REPORT_UPDATED | report_id={report_id} | sheet_row={sheet_row}")
        return jsonify({'success': True, 'report_id': report_id})

    except Exception as e:
        log.error(f"REPORT_UPDATE_FAIL | report_id={report_id} | error={e}")
        return jsonify({'error': str(e)}), 500


# ── DELETE REPORT ──────────────────────────────────────
@app.route('/api/report/<report_id>', methods=['DELETE'])
@admin_token_required
def delete_report(report_id):
    """Удаление строки из Sheets по индексу."""
    if not report_id.startswith('sheets_'):
        return jsonify({'error': 'Invalid report_id format'}), 400

    try:
        row_idx   = int(report_id.split('_')[1])
        sheet_row = row_idx + 2
    except (IndexError, ValueError):
        return jsonify({'error': 'Invalid report_id'}), 400

    try:
        client = _sheets_client()
        sheet  = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_REPORTS)
        sheet.delete_rows(sheet_row)
        log.info(f"REPORT_DELETED | report_id={report_id} | sheet_row={sheet_row}")
        return jsonify({'success': True, 'deleted': report_id})
    except Exception as e:
        log.error(f"REPORT_DELETE_FAIL | report_id={report_id} | error={e}")
        return jsonify({'error': str(e)}), 500


# ── PROJECT MANAGEMENT ─────────────────────────────────
@app.route('/api/project', methods=['POST'])
@admin_token_required
def add_project():
    data = request.get_json(silent=True) or {}

    abbr = str(data.get('abbr', '')).strip().upper()
    full = str(data.get('full', '')).strip()

    if len(abbr) < 2:
        return jsonify({'error': 'abbr минимум 2 символа'}), 400
    if len(full) < 3:
        return jsonify({'error': 'full минимум 3 символа'}), 400

    projects = load_json(PROJECTS_FILE, [])
    if any(p['abbr'] == abbr for p in projects):
        return jsonify({'error': 'Проект уже существует'}), 400

    projects.append({'abbr': abbr, 'full': full})
    save_json(PROJECTS_FILE, projects)

    log.info(f"PROJECT_ADDED | admin={request.current_user['uid']} | abbr={abbr} | full={full}")
    return jsonify({'success': True, 'project': {'abbr': abbr, 'full': full}})


@app.route('/api/project/<abbr>', methods=['DELETE'])
@admin_token_required
def remove_project(abbr):
    projects = load_json(PROJECTS_FILE, [])
    before   = len(projects)
    projects = [p for p in projects if p['abbr'] != abbr]

    if len(projects) == before:
        return jsonify({'error': 'Проект не найден'}), 404

    save_json(PROJECTS_FILE, projects)
    log.info(f"PROJECT_REMOVED | admin={request.current_user['uid']} | abbr={abbr}")
    return jsonify({'success': True, 'deleted': abbr})


# ── ASSIGN PROJECTS ────────────────────────────────────
@app.route('/api/assign', methods=['POST'])
@admin_token_required
def assign_projects():
    data    = request.get_json(silent=True) or {}
    user_id = data.get('user_id')

    if not user_id:
        return jsonify({'error': 'user_id required'}), 400

    abbrs       = data.get('abbrs', [])
    assignments = load_json(USER_PROJECTS_FILE, {})
    assignments[str(user_id)] = abbrs
    save_json(USER_PROJECTS_FILE, assignments)

    log.info(f"PROJECTS_ASSIGNED | admin={request.current_user['uid']} | user={user_id} | abbrs={abbrs}")
    return jsonify({'success': True, 'user_id': user_id, 'projects': abbrs})


# ══════════════════════════════════════════════════════
# УПРАВЛЕНИЕ СОТРУДНИКАМИ (только админ, по токену)
# ══════════════════════════════════════════════════════

def _gen_unique_user_id(users: dict) -> str:
    """Генерит 10-значный id, не пересекающийся с существующими."""
    while True:
        nid = str(9_000_000_000 + secrets.randbelow(999_999_999))
        if nid not in users:
            return nid


@app.route('/api/admin/employees', methods=['GET'])
@admin_token_required
def admin_employees():
    """Список всех сотрудников со статусом регистрации и кода."""
    users  = load_json(USERS_FILE, {})
    result = []
    for uid, u in users.items():
        result.append({
            'id':            uid,
            'username':      u.get('username', ''),
            'display_name':  u.get('display_name', u.get('username', '')),
            'registered':    bool(u.get('password_hash')),
            'invite_code':   u.get('invite_code'),
            'invite_used':   bool(u.get('invite_used', False)),
            'is_test':       bool(u.get('is_test', False)),
            'is_admin':      is_admin(uid),
            'monthly_norm':  get_monthly_norm(u, msk_now().strftime('%Y-%m')),
            'schedule':      u.get('schedule', ''),
        })
    # Тестовые сверху, далее по ФИО
    result.sort(key=lambda x: (not x['is_test'], x['display_name'].lower()))
    return jsonify({'employees': result})


@app.route('/api/admin/employee/add', methods=['POST'])
@admin_token_required
def admin_employee_add():
    """Добавить нового сотрудника (без пароля, ждёт кода-приглашения)."""
    data         = request.get_json(silent=True) or {}
    display_name = str(data.get('display_name', '')).strip()
    username     = str(data.get('username', '')).strip()
    if not display_name or not username:
        return jsonify({'error': 'ФИО и ник обязательны'}), 400

    users = load_json(USERS_FILE, {})
    if any(u.get('username', '').lower() == username.lower() for u in users.values()):
        return jsonify({'error': 'Такой ник уже существует'}), 400

    schedule = str(data.get('schedule', '')).strip()
    if schedule not in VALID_SCHEDULES:
        schedule = ''

    try:
        monthly_norm = int(data.get('monthly_norm', 160) or 160)
        if monthly_norm < 0 or monthly_norm > 744:
            monthly_norm = 160
    except (ValueError, TypeError):
        monthly_norm = 160

    new_id = _gen_unique_user_id(users)
    rec = {
        'username':      username,
        'display_name':  display_name,
        'registered_at': '',
        'status':        'active',
        'password_hash': None,
        'invite_code':   None,
        'invite_used':   False,
    }
    if schedule:
        rec['schedule'] = schedule
    else:
        rec['monthly_norm'] = monthly_norm
    users[new_id] = rec
    if not save_json(USERS_FILE, users):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    log.info(f"ADMIN_EMP_ADD | by={request.current_user.get('uid')} | new_id={new_id} | username={username}")
    return jsonify({'success': True, 'id': new_id})


@app.route('/api/admin/employee/delete', methods=['POST'])
@admin_token_required
def admin_employee_delete():
    """Удалить сотрудника. Администратора удалить нельзя."""
    data = request.get_json(silent=True) or {}
    uid  = str(data.get('user_id', ''))
    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404
    if is_admin(uid):
        return jsonify({'error': 'Нельзя удалить администратора'}), 400
    removed = users.pop(uid)
    if not save_json(USERS_FILE, users):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    log.info(f"ADMIN_EMP_DEL | by={request.current_user.get('uid')} | uid={uid} | username={removed.get('username')}")
    return jsonify({'success': True})


@app.route('/api/admin/employee/invite', methods=['POST'])
@admin_token_required
def admin_employee_invite():
    """Сгенерировать/перевыпустить персональный код. Возвращается админу."""
    data = request.get_json(silent=True) or {}
    uid  = str(data.get('user_id', ''))
    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404

    # Код уникален среди действующих, чтобы не пересекался
    existing = {
        str(u.get('invite_code', '')).upper()
        for u in users.values() if u.get('invite_code')
    }
    code = gen_invite_code(4)
    while code.upper() in existing:
        code = gen_invite_code(4)

    users[uid]['invite_code'] = code
    users[uid]['invite_used'] = False   # перевыпуск делает код снова рабочим
    if not save_json(USERS_FILE, users):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    # Сам код НЕ пишем в лог
    log.info(f"ADMIN_EMP_INVITE | by={request.current_user.get('uid')} | uid={uid} | code=***")
    return jsonify({'success': True, 'invite_code': code})


@app.route('/api/admin/employee/reset', methods=['POST'])
@admin_token_required
def admin_employee_reset():
    """Сбросить пароль и PIN. Сотрудник регистрируется заново по новому коду."""
    data = request.get_json(silent=True) or {}
    uid  = str(data.get('user_id', ''))
    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404
    users[uid]['password_hash'] = None
    if not save_json(USERS_FILE, users):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    log.info(f"ADMIN_EMP_RESET | by={request.current_user.get('uid')} | uid={uid}")
    return jsonify({'success': True})


@app.route('/api/admin/employee/schedule', methods=['POST'])
@admin_token_required
def admin_employee_schedule():
    """Установить тип расписания сотрудника (5/2, 2/2A, 2/2B) или сбросить."""
    data     = request.get_json(silent=True) or {}
    uid      = str(data.get('user_id', ''))
    schedule = str(data.get('schedule', '')).strip()
    if schedule and schedule not in VALID_SCHEDULES:
        return jsonify({'error': 'Неверный тип расписания'}), 400
    users = load_json(USERS_FILE, {})
    if uid not in users:
        return jsonify({'error': 'Пользователь не найден'}), 404
    if schedule:
        users[uid]['schedule'] = schedule
        users[uid].pop('monthly_norm', None)
    else:
        users[uid].pop('schedule', None)
        users[uid]['monthly_norm'] = int(data.get('monthly_norm', 160) or 160)
    if not save_json(USERS_FILE, users):
        return jsonify({'error': 'Не удалось сохранить'}), 500
    log.info(f"ADMIN_EMP_SCHED | by={request.current_user.get('uid')} | uid={uid} | schedule={schedule}")
    return jsonify({'success': True})


# ══════════════════════════════════════════════════════
# ЗАПУСК
# ══════════════════════════════════════════════════════
if __name__ == '__main__':
    log.info("══ API запущен ══════════════════════════")
    log.info(f"Sheets ID: {SPREADSHEET_ID or '⚠️  НЕ ЗАДАН'}")
    log.info(f"Admins:    {ADMIN_IDS}")
    if JWT_SECRET == 'phm-secret-change-me-in-production':
        log.warning("⚠️  JWT_SECRET использует дефолт — задайте переменную окружения в проде!")
    log.info("═════════════════════════════════════════")
    app.run(host='0.0.0.0', port=5000, debug=False)
