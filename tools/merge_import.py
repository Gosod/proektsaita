#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Слияние импортных данных в рабочие JSON-файлы приложения (с бэкапом).

Берёт из каталога импорта:
  vacations_import.json     → сливает в <data>/vacations.json
  report_flags_import.json  → сливает в <data>/report_flags.json

Слияние аддитивное (существующие записи сохраняются, новые добавляются/обновляются).
Перед записью делается резервная копия *.bak-YYYYmmddHHMMSS.

Запуск на сервере из каталога приложения:
  python3 tools/merge_import.py --import-dir import_out --data-dir .
"""
import argparse
import datetime
import json
import os
import shutil


def load(path):
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return {}


def backup(path):
    if os.path.exists(path):
        b = f"{path}.bak-{datetime.datetime.now():%Y%m%d%H%M%S}"
        shutil.copy2(path, b)
        print(f"  бэкап: {b}")


def save(path, data):
    backup(path)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  записан: {path}")


def merge_vacations(dst, src):
    added = 0
    for uid, days in src.items():
        d = dst.setdefault(uid, {})
        for day, val in days.items():
            if day not in d:
                added += 1
            d[day] = val
    return dst, added


def merge_flags(dst, src):
    added = 0
    for uid, dates in src.items():
        u = dst.setdefault(uid, {})
        for date, projs in dates.items():
            dd = u.setdefault(date, {})
            for proj, t in projs.items():
                if proj not in dd:
                    added += 1
                dd[proj] = t
    return dst, added


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--import-dir', default='import_out')
    ap.add_argument('--data-dir', default='.')
    args = ap.parse_args()

    vac_src = os.path.join(args.import_dir, 'vacations_import.json')
    flags_src = os.path.join(args.import_dir, 'report_flags_import.json')
    vac_dst = os.path.join(args.data_dir, 'vacations.json')
    flags_dst = os.path.join(args.data_dir, 'report_flags.json')

    if os.path.exists(vac_src):
        merged, added = merge_vacations(load(vac_dst), load(vac_src))
        print(f"Отпуска: +{added} дней")
        save(vac_dst, merged)
    else:
        print(f"нет {vac_src} — пропуск отпусков")

    if os.path.exists(flags_src):
        merged, added = merge_flags(load(flags_dst), load(flags_src))
        print(f"Флаги переработка/отработка: +{added}")
        save(flags_dst, merged)
    else:
        print(f"нет {flags_src} — пропуск флагов")

    print("Готово.")


if __name__ == '__main__':
    main()
