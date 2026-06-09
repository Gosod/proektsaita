#!/usr/bin/env python3
"""
Сброс тест-сотрудника в исходное «незарегистрированное» состояние.
Используется только для локального тестирования цикла регистрации по коду.

Запуск:  py reset_test.py
"""
import json
import os

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, 'users.json')

TEST_UID  = '100000001'
TEST_CODE = 'TEST'


def main():
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        users = json.load(f)

    if TEST_UID not in users:
        print(f'❌ Тест-сотрудник {TEST_UID} не найден в users.json')
        return

    u = users[TEST_UID]
    u['password_hash'] = None
    u['pin_hash']      = None
    u['invite_code']   = TEST_CODE
    u['invite_used']   = False
    u['registered_at'] = ''
    u['status']        = 'active'
    u['is_test']       = True

    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

    print('✅ Тест-сотрудник сброшен:')
    print(f'   {u.get("display_name")} (id {TEST_UID})')
    print(f'   код-приглашение: {TEST_CODE}, пароль очищен')
    print('   Можно снова регистрироваться. Перезапуск сервера не нужен.')


if __name__ == '__main__':
    main()
