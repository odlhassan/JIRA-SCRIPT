import gc
import json
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

from flask import Flask
from epr_auth import OWNER_ACCOUNT_ID, normalize_game_shortcut, register_epr_auth


class ShortcutTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(gc.collect)
        self.directory = Path(self.tmp.name) / 'oauth'
        self.directory.mkdir()
        env = patch.dict(os.environ, {'EPR_AUTH_ENABLED': '1', 'JIRA_OAUTH_CLIENT_ID': '', 'JIRA_OAUTH_CLIENT_SECRET': '', 'JIRA_OAUTH_REDIRECT_URI': ''})
        env.start(); self.addCleanup(env.stop)
        self.app = Flask(__name__)
        register_epr_auth(self.app, self.directory, {'report.html'})
        self.client = self.app.test_client()
        self.headers = {'Origin': 'http://localhost'}

    def login(self, owner=True):
        account, sid = (OWNER_ACCOUNT_ID, 'a' * 64) if owner else ('other-admin', 'b' * 64)
        (self.directory / (sid + '.json')).write_text(json.dumps({'session_expires_at': time.time()+86400, 'epr_signed_at': time.time(), 'epr_csrf': 'secret', 'epr_identity': {'account_id': account, 'display_name': account, 'avatar_url': ''}}))
        self.client.set_cookie('epr_jira_session', sid)
        return self.client.get('/api/auth/me').json['user']

    def test_owner_changes_shortcut_and_approved_users_read_same_saved_setting(self):
        self.assertEqual(self.client.get('/api/game-shortcut').status_code, 401)
        owner = self.login()
        self.assertTrue(owner['is_superadmin'])
        self.assertEqual(self.client.get('/api/game-shortcut').json['shortcut'], 'Ctrl+G')
        result = self.client.put('/api/auth/game-shortcut', json={'shortcut': 'shift + ctrl + h'}, headers=self.headers)
        self.assertEqual(result.json, {'shortcut': 'Ctrl+Shift+H'})
        path = self.directory.parent / 'epr_access' / 'game_shortcut.json'
        self.assertEqual(json.loads(path.read_text()), result.json)
        self.login(False)
        self.login()
        self.client.put('/api/auth/users/other-admin', json={'role': 'member', 'approved': True, 'active': True}, headers=self.headers)
        self.login(False)
        self.assertEqual(self.client.get('/api/game-shortcut').json['shortcut'], 'Ctrl+Shift+H')
        self.assertEqual(self.client.get('/api/auth/game-shortcut').status_code, 403)
        self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': 'Alt+P'}, headers=self.headers).status_code, 403)

    def test_ordinary_admin_cannot_change_owner_preference(self):
        self.login(False)
        self.login()
        self.client.put('/api/auth/users/other-admin', json={'role': 'admin', 'approved': True, 'active': True}, headers=self.headers)
        user = self.login(False)
        self.assertFalse(user['is_superadmin'])
        self.assertEqual(self.client.get('/api/auth/game-shortcut').status_code, 403)
        self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': 'Alt+P'}, headers=self.headers).status_code, 403)

    def test_validation_and_csrf_preserve_last_saved_value(self):
        self.login()
        self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': 'Alt+P'}).status_code, 403)
        self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': 'Alt+P'}, headers={'Origin': 'https://evil.test'}).status_code, 403)
        for value in ['G', 'Shift+G', 'Ctrl+Ctrl+G', 'Ctrl++', 'Ctrl+Enter', '<script>', None, ['Ctrl+G']]:
            with self.subTest(value=value):
                self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': value}, headers=self.headers).status_code, 400)
        self.assertEqual(self.client.get('/api/game-shortcut').json['shortcut'], 'Ctrl+G')
        self.assertEqual(self.client.put('/api/auth/game-shortcut', json={'shortcut': 'Alt+P'}, headers={'X-EPR-CSRF': 'secret'}).status_code, 200)
        self.assertEqual(normalize_game_shortcut(' Ctrl + Alt + 7 '), 'Ctrl+Alt+7')


if __name__ == '__main__':
    unittest.main()
