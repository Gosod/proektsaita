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
    {'uid': '100000001', 'code': 'TEST',    'display': 'ТЕСТ Тестовый',      'schedule': '5/2'},
    {'uid': '100000002', 'code': 'TESTADM', 'display': 'ТЕСТ Администратор', 'norm': 0},
    {'uid': '100000003', 'code': 'TESTA',   'display': 'ТЕСТ Смена А',       'schedule': '2/2A'},
    {'uid': '100000004', 'code': 'TESTB',   'display': 'ТЕСТ Смена Б',       'schedule': '2/2B'},
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
        if 'schedule' in acc:
            u['schedule'] = acc['schedule']
            u.pop('monthly_norm', None)
        else:
            u.pop('schedule', None)
            u['monthly_norm'] = acc['norm']
        print(f'✅ {acc["display"]} (id {uid}) — сброшен, код: {acc["code"]}')

    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

    print('\nПерезапуск сервера не нужен.')


if __name__ == '__main__':
    main()
