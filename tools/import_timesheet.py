#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер мастер-табеля (Excel) в формат приложения Фотомеханика.

Вход:  xlsx с помесячными листами (янв, мар, апр, май, июн ...), где
       строка 1 — номера дней, колонка A — сотрудник, колонка B — проект,
       далее по дням — часы (числа) либо буквенные пометки.

Выход (каталог --out):
  - reports_import.tsv   — строки для листа «Отчёты» Google Sheets:
        Дата<TAB>Время<TAB>Сотрудник<TAB>Проект<TAB>Часы<TAB>Комментарий<TAB>ID
  - vacations_import.json — отпуска по user_id: {uid: {"YYYY-MM-DD": true}}
  - unmatched.txt        — сотрудники из Excel, которых не нашли в users.json

Привязка к сотрудникам — по username из users.json (как в
_normalize_sheets_records: строка с неизвестным ником молча отбрасывается).
Поэтому ВАЖНО передать актуальный users.json (--users).

По договорённости:
  • импортируются только ЧИСЛОВЫЕ часы и пометка «отпуск» (отп/отпуск → зелёный);
  • прочие буквенные пометки (sl, бл, отг, с/с, командировка) игнорируются;
  • комментарии из таблицы не переносятся;
  • перед импортом соответствующие месяцы в таблице очищаются вручную.

Запуск:
  python3 tools/import_timesheet.py --xlsx табель.xlsx --users users.json \
      --year 2026 --out import_out
"""
import argparse
import json
import os
import re
import uuid
import collections

import openpyxl

MONTH_NUM = {
    'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'май': 5, 'июн': 6,
    'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12,
}

VACATION_TOKENS = {'отп', 'отпуск', 'отпуск/сиклив'}

SUR_SUF = re.compile(r'(ов|ёв|ев|ин|ын|ский|цкий|ской|их|ых|ко|юк|ук|ян|дзе|швили)$', re.I)
PATR    = re.compile(r'(ович|евич|ьич|инич|овна|евна|ична)$', re.I)


def name_key(raw: str):
    """Ключ кластеризации (фамилия, первая буква имени) — сводит короткие и
    полные формы имени одного человека к одному ключу."""
    toks = raw.split()
    sur = None
    for t in toks:
        if PATR.search(t):
            continue
        if SUR_SUF.search(t):
            sur = t
            break
    if sur is None:
        sur = toks[0]
    given = None
    for t in toks:
        if t == sur or PATR.search(t):
            continue
        given = t
        break
    if given is None:
        given = toks[-1]
    return (sur.lower(), given[:1].lower())


def day_columns(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, (int, float)) and 1 <= int(v) <= 31 and float(v) == int(v):
            cols[c] = int(v)
    return cols


def fmt_hours(v) -> str:
    """8.0 → '8', 7.5 → '7,5' (запятая — десятичный разделитель, как в Sheets)."""
    if float(v) == int(v):
        return str(int(v))
    return str(v).replace('.', ',')


def build_user_index(users: dict):
    """key(display_name) → {'username':..., 'id':...} для авто-сопоставления."""
    idx = {}
    for uid, u in users.items():
        dn = (u.get('display_name') or u.get('username') or '').strip()
        if not dn:
            continue
        idx[name_key(dn)] = {'username': u.get('username', ''), 'id': str(uid)}
    return idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--xlsx', required=True)
    ap.add_argument('--users', help='users.json для привязки имён к username/id')
    ap.add_argument('--year', type=int, default=2026)
    ap.add_argument('--out', default='import_out')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    month_sheets = [s for s in wb.sheetnames if s.lower() in MONTH_NUM]

    users = {}
    if args.users and os.path.exists(args.users):
        users = json.load(open(args.users, encoding='utf-8'))
    user_idx = build_user_index(users)

    # raw-имя сотрудника → разрешённый (username, id) либо заглушка
    resolved = {}
    canon_name = {}   # key → представительное ФИО (самое длинное)

    def resolve(raw):
        k = name_key(raw)
        if len(raw) > len(canon_name.get(k, '')):
            canon_name[k] = raw
        if k in user_idx:
            return user_idx[k]['username'], user_idx[k]['id'], True
        return canon_name[k], '', False   # заглушка: ФИО как ник, без id

    reports = []                       # (date_iso, username, project, hours)
    vacations = collections.defaultdict(dict)   # uid → {date_iso: True}
    vac_by_name = collections.defaultdict(set)   # для превью без users.json
    markers_skipped = collections.Counter()
    work_days = collections.defaultdict(set)     # (username, date_iso) с часами

    for name in month_sheets:
        ws = wb[name]
        month = MONTH_NUM[name.lower()]
        dcols = day_columns(ws)
        cur = None
        for r in range(3, ws.max_row + 1):
            a = ws.cell(r, 1).value
            b = ws.cell(r, 2).value
            if a and isinstance(a, str) and a.strip().endswith('ИТОГ'):
                continue
            if a and isinstance(a, str) and a.strip():
                cur = a.strip()
            if not cur or not (b and str(b).strip()):
                continue
            project = str(b).strip()
            username, uid, matched = resolve(cur)
            for c, day in dcols.items():
                v = ws.cell(r, c).value
                if v is None or v == '':
                    continue
                date_iso = f"{args.year:04d}-{month:02d}-{day:02d}"
                if isinstance(v, (int, float)):
                    if v:
                        reports.append((date_iso, username, project, float(v), matched))
                        work_days[(username, date_iso)].add(project)
                else:
                    tok = str(v).strip().lower()
                    if tok in VACATION_TOKENS:
                        if uid:
                            vacations[uid][date_iso] = True
                        vac_by_name[username].add(date_iso)
                    else:
                        markers_skipped[tok] += 1

    # Конфликт «отпуск + часы в один день» — отчёты приоритетнее, отпуск убираем
    conflicts = 0
    for uid, days in list(vacations.items()):
        # найдём username по uid
        uname = next((u['username'] for u in user_idx.values() if u['id'] == uid), None)
        for d in list(days):
            if uname and (uname, d) in work_days:
                del days[d]
                conflicts += 1

    os.makedirs(args.out, exist_ok=True)
    HEADER = 'Дата\tВремя\tСотрудник\tПроект\tЧасы\tКомментарий\tID\n'

    def write_tsv(path, rows):
        with open(path, 'w', encoding='utf-8') as f:
            f.write(HEADER)
            for date_iso, username, project, hours in rows:
                y, m, d = date_iso.split('-')
                f.write(f"{d}.{m}.{y}\t\t{username}\t{project}\t{fmt_hours(hours)}\t\t{uuid.uuid4().hex[:12]}\n")

    matched_rows   = [(d, u, p, h) for d, u, p, h, m in reports if m]
    unmatched_rows = [(d, u, p, h) for d, u, p, h, m in reports if not m]
    # Готов к вставке в лист «Отчёты» — только привязываемые сотрудники
    tsv_path = os.path.join(args.out, 'reports_import.tsv')
    write_tsv(tsv_path, matched_rows)
    # Остальные (нет аккаунта на сайте) — отдельно, на случай заведения аккаунтов
    if unmatched_rows:
        write_tsv(os.path.join(args.out, 'reports_unmatched.tsv'), unmatched_rows)

    vac_path = os.path.join(args.out, 'vacations_import.json')
    json.dump(vacations, open(vac_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

    print(f"Месяцы:            {month_sheets}")
    print(f"Строк-отчётов:     {len(reports)}  (привязано: {len(matched_rows)} | без аккаунта: {len(unmatched_rows)})")
    matched_keys = [k for k in canon_name if k in user_idx]
    unmatched_names = sorted(canon_name[k] for k in canon_name if k not in user_idx)
    print(f"Сотрудников:       {len(canon_name)}  (сопоставлено в users.json: {len(matched_keys)})")
    vac_total = sum(len(d) for d in vac_by_name.values())
    print(f"Дней отпуска:      {vac_total}  (привязано к id: {sum(len(d) for d in vacations.values())})")
    print(f"Конфликтов отпуск/часы (отпуск снят): {conflicts}")
    print(f"Пропущено пометок: {sum(markers_skipped.values())}  {dict(markers_skipped)}")
    print(f"\nФайлы: {tsv_path}\n       {vac_path}")
    if unmatched_names:
        with open(os.path.join(args.out, 'unmatched.txt'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(unmatched_names))
        print(f"\n⚠️  Не нашли в users.json ({len(unmatched_names)}): {unmatched_names}")
        print("    Эти строки попадут в TSV с ФИО вместо username и НЕ привяжутся в приложении.")


if __name__ == '__main__':
    main()
