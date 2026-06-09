#!/usr/bin/env python3
"""
Сброс тест-сотрудников в исходное «незарегистрированное» состояние.
Используется только для локального тестирования.

Запуск:  py reset_test.py
"""
import json
import os

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, 'users.json')

TEST_ACCOUNTS = [
    {'uid': '100000001', 'code': 'TEST',    'display': 'ТЕСТ Тестовый',       'norm': 160},
    {'uid': '100000002', 'code': 'TESTADM', 'display': 'ТЕСТ Администратор',  'norm': 0},
]


def main():
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        users = json.load(f)

    for acc in TEST_ACCOUNTS:
        uid = acc['uid']
        if uid not in users:
            print(f'⚠️  {acc["display"]} (id {uid}) не найден — пропускаю')
            continue
        u = users[uid]
        u['password_hash'] = None
        u['pin_hash']      = None
        u['invite_code']   = acc['code']
        u['invite_used']   = False
        u['registered_at'] = ''
        u['status']        = 'active'
        u['is_test']       = True
        u['monthly_norm']  = acc['norm']
        print(f'✅ {acc["display"]} (id {uid}) — сброшен, код: {acc["code"]}')

    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

    print('\nПерезапуск сервера не нужен.')


if __name__ == '__main__':
    main()
