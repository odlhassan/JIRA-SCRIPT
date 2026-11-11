import gc
import json
import math
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from flask import Flask, jsonify
from werkzeug.test import Client
from werkzeug.wrappers import Response

import maintenance_service as service
from maintenance_service import MaintenanceGateway, Playground, write_status


class GameTests(unittest.TestCase):
    def setUp(self):
        self.now = 100.0
        self.game = Playground(lambda: self.now)

    def join(self, name):
        return self.game.join(name)

    def test_first_police_second_runner_no_tokens_leak(self):
        cop = self.join('Officer')
        runner = self.join('<script>runner</script>')
        state = self.game.snapshot(runner)
        self.assertEqual(self.game.players[cop]['role'], 'police')
        self.assertEqual(self.game.players[runner]['role'], 'runner')
        self.assertNotIn(cop, json.dumps(state))
        self.assertNotIn(runner, json.dumps(state))
        self.assertEqual(len(self.game.players[runner]['name']), 18)

    def test_approval_is_police_only_and_atomic_max_two(self):
        cop, one, two = [self.join(n) for n in ['Cop', 'One', 'Two']]
        for token in [one, two]:
            self.game.command(token, {'action': 'request'})
        with self.assertRaises(ValueError):
            self.game.command(one, {'action': 'accept', 'player_id': self.game.players[two]['id']})
        def accept(token):
            try:
                self.game.command(cop, {'action': 'accept', 'player_id': self.game.players[token]['id']})
            except ValueError:
                pass
        threads = [threading.Thread(target=accept, args=(token,)) for token in [one, two]]
        for t in threads: t.start()
        for t in threads: t.join()
        self.assertEqual(len(self.game._police()), 2)

    def test_capture_jail_rescue_and_immunity(self):
        cop, caught, rescuer = [self.join(n) for n in ['Cop', 'Caught', 'Rescuer']]
        self.now += 4
        self.game.players[cop].update(x=360, y=300)
        self.game.players[caught].update(x=365, y=300)
        self.game.snapshot()
        self.assertTrue(self.game.players[caught]['jailed'])
        self.assertTrue(450 <= self.game.players[caught]['x'] <= 550)
        self.game.players[rescuer].update(x=570, y=300)
        self.game.snapshot()
        self.assertFalse(self.game.players[caught]['jailed'])
        self.assertGreater(self.game.players[caught]['immune_until'], self.now)

    def test_all_caught_starts_new_round(self):
        cop, runner = self.join('Cop'), self.join('Runner')
        self.now += 4
        self.game.players[runner].update(x=90, y=300)
        self.game.snapshot()
        self.assertTrue(self.game.players[runner]['jailed'])
        self.now += 5.1
        state = self.game.snapshot()
        self.assertEqual(state['round'], 2)
        self.assertFalse(self.game.players[runner]['jailed'])

    def test_server_movement_clamp_collision_invalid_input(self):
        token = self.join('Cop')
        p = self.game.players[token]
        p.update(x=160, y=150)
        self.game.command(token, {'action': 'move', 'dx': 1, 'dy': 0, 'x': 990})
        self.now += .1
        self.game.snapshot()
        self.assertGreater(p['x'], 160)  # Small steps approach the wall without crossing it.
        self.assertLessEqual(p['x'], 165)
        for dx in [float('nan'), float('inf'), 100, '1', None]:
            with self.subTest(dx=dx), self.assertRaises(ValueError):
                self.game.command(token, {'action': 'move', 'dx': dx})
        p.update(x=90, y=300)
        self.game.command(token, {'action': 'move', 'dx': 1, 'dy': 1})
        self.now += .1
        self.game.snapshot()
        self.assertAlmostEqual(((p['x'] - 90)**2 + (p['y'] - 300)**2)**.5, 18)
        self.now += 1
        before = p['x']
        before_y = p['y']
        self.game.snapshot()
        self.assertAlmostEqual(math.hypot(p['x']-before, p['y']-before_y), 90)
        stopped = (p['x'], p['y'])
        self.now += 1
        self.game.snapshot()
        self.assertEqual(stopped, (p['x'], p['y']))  # Stops at input expiry, not at poll boundaries.

    def test_speed_does_not_depend_on_poll_frequency_or_number_of_readers(self):
        positions = []
        for steps in [1, 2, 12, 40]:
            self.setUp()
            token = self.join('Cop')
            self.game.command(token, {'action': 'move', 'dx': 1, 'dy': 0})
            for _ in range(steps):
                self.now += .2 / steps
                self.game.snapshot()
            positions.append(self.game.players[token]['x'])
        for x in positions:
            self.assertAlmostEqual(x, 126)

    def test_large_time_gap_cannot_tunnel_through_a_house(self):
        token = self.join('Cop')
        self.game.players[token].update(x=120, y=150)
        self.game.command(token, {'action': 'move', 'dx': 1, 'dy': 0})
        self.now += 2
        state = self.game.snapshot(token)
        self.assertLessEqual(state['players'][0]['x'], 165)
        self.assertEqual(state['players'][0]['dx'], 0)
        self.assertEqual(state['motion']['speed'], self.game.SPEED)

    def test_disconnect_promotes_oldest_and_jail_cannot_request(self):
        cop, runner = self.join('Cop'), self.join('Runner')
        self.game.players[runner]['jailed'] = True
        with self.assertRaises(ValueError):
            self.game.command(runner, {'action': 'request'})
        self.now += 19
        self.game.snapshot(runner)
        self.now += 2
        self.game.snapshot()
        self.assertNotIn(cop, self.game.players)
        self.assertEqual(self.game.players[runner]['role'], 'police')
        self.assertFalse(self.game.players[runner]['jailed'])

    def test_rejoin_idempotent_and_capacity_bounded(self):
        token = self.join('Cop')
        self.assertEqual(self.game.join('New', token), token)
        self.assertEqual(len(self.game.players), 1)
        for n in range(39): self.join(str(n))
        with self.assertRaises(ValueError): self.join('Overflow')


class GatewayTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(gc.collect)
        self.marker = Path(self.tmp.name) / 'state.json'
        core = Flask(__name__)
        self.calls = []
        @core.route('/report', methods=['GET', 'POST'])
        def report():
            self.calls.append('core')
            return jsonify(report=True)
        @core.get('/protected')
        def protected():
            return '', 401
        self.gate = MaintenanceGateway(core, marker=self.marker, base=Path(self.tmp.name))
        self.client = Client(self.gate, Response)
        self.headers = {'Origin': 'http://localhost'}

    def pause(self, phase='deploying'):
        write_status(self.marker, phase, 'Deployment in progress', 'run-1')

    def test_gate_blocks_document_and_mutation_before_core_and_reopens(self):
        self.pause()
        page = self.client.get('/report', headers={'Accept': 'text/html'})
        self.assertEqual(page.status_code, 503)
        self.assertIn(b'Chorr', page.data)
        self.assertEqual(page.headers['Retry-After'], '5')
        page.close()
        self.assertEqual(self.client.post('/report', json={'write': 1}).status_code, 503)
        self.assertEqual(self.calls, [])
        self.pause('available')
        self.assertEqual(self.client.get('/report').status_code, 200)
        self.assertEqual(self.client.get('/protected').status_code, 401)

    def test_failed_or_invalid_marker_stays_paused(self):
        for raw in ['broken', '[]', '{"phase":"typo"}']:
            self.marker.write_text(raw)
            self.assertFalse(self.client.get('/api/maintenance/status').json['available'])
            self.assertEqual(self.client.post('/report').status_code, 503)
        self.pause('failed')
        self.assertFalse(self.client.get('/api/maintenance/status').json['available'])

    def test_public_shell_game_and_health_no_report_auth_bypass(self):
        self.pause()
        for path in ['/maintenance', '/maintenance/game.css', '/maintenance/game.js', '/maintenance/motion.js', '/maintenance/app-shortcuts.js', '/healthz']:
            with self.subTest(path=path), self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)
        self.assertEqual(self.client.get('/maintenance/secrets').status_code, 404)
        self.assertEqual(self.client.get('/healthz').json, {'ok': True})
        self.assertEqual(self.client.post('/api/maintenance/game', json={'action': 'join'}, headers=self.headers).status_code, 200)
        self.assertTrue(self.client.get('/api/maintenance/game').json['you'])
        self.assertEqual(self.calls, [])

    def test_game_origin_json_limits_and_opaque_cookie(self):
        self.assertEqual(self.client.post('/api/maintenance/game', json={'action': 'join'}).status_code, 403)
        self.assertEqual(self.client.post('/api/maintenance/game', json={'action': 'join'}, headers={'Origin': 'https://evil.test'}).status_code, 403)
        self.assertEqual(self.client.post('/api/maintenance/game', json=[], headers=self.headers).status_code, 400)
        response = self.client.post('/api/maintenance/game', json={'action': 'join', 'name': 'Visitor'}, headers=self.headers)
        self.assertIn('HttpOnly', response.headers['Set-Cookie'])
        self.assertIn('SameSite=Strict', response.headers['Set-Cookie'])
        self.assertEqual(self.client.post('/api/maintenance/game', json={'action': 'move', 'dx': 200}, headers=self.headers).status_code, 400)
        self.assertEqual(self.client.post('/api/maintenance/game', json={'action': 'join', 'name': 'x'*3000}, headers=self.headers).status_code, 413)

    def test_inflight_stream_counts_until_closed(self):
        def streaming(environ, start_response):
            start_response('200 OK', [('Content-Type', 'text/plain')])
            return iter([b'first', b'second'])
        self.gate.core = streaming
        response = self.client.get('/report', buffered=False)
        self.pause('draining')
        self.assertEqual(self.client.get('/api/maintenance/status').json['active_requests'], 1)
        self.assertEqual(self.client.get('/report').status_code, 503)
        response.close()
        self.assertEqual(self.client.get('/api/maintenance/status').json['active_requests'], 0)

    def test_background_work_drains_after_response(self):
        release = threading.Event()
        job = service.start_background_job(release.wait)
        try:
            self.assertEqual(self.gate.status()['background_jobs'], 1)
        finally:
            release.set(); job.join(timeout=2)
        self.assertEqual(self.gate.status()['background_jobs'], 0)

    def test_startup_shell_available_while_core_warms_and_failure_keeps_gate(self):
        release = threading.Event()
        started = threading.Event()
        def build():
            started.set(); release.wait(2)
            return self.gate.core
        gate = MaintenanceGateway(factory=build, marker=self.marker)
        started.wait(1)
        client = Client(gate, Response)
        try:
            with client.get('/maintenance') as response:
                self.assertEqual(response.status_code, 200)
            self.assertFalse(client.get('/api/maintenance/status').json['ready'])
            self.assertEqual(client.post('/report').status_code, 503)
        finally:
            release.set()
        with patch('maintenance_service.logging.exception'):
            gate.failed = True
            self.assertFalse(gate.status()['available'])

    def test_real_reporting_app_gate_precedes_auth_and_reopening_restores_auth(self):
        from report_server import create_report_server_app
        root = Path(self.tmp.name)
        with patch.dict('os.environ', {'EPR_AUTH_ENABLED': '1', 'JIRA_PROJECT_KEYS': '', 'EPR_MAINTENANCE_FILE': str(self.marker)}):
            app = create_report_server_app(base_dir=root, folder_raw='report_html')
            client = app.test_client()
            preferences = app.config['EPR_AUTH_DIRECTORY'] / 'game_shortcut.json'
            preferences.parent.mkdir(parents=True, exist_ok=True)
            preferences.write_text(json.dumps({'shortcut':'Ctrl+Shift+H'}))
            self.pause()
            self.assertEqual(client.get('/api/maintenance/status').json['shortcut'], 'Ctrl+Shift+H')
            with client.get('/login', headers={'Accept': 'text/html'}) as response:
                self.assertEqual(response.status_code, 503)
            self.assertEqual(client.post('/api/capacity', json={}).status_code, 503)
            self.assertEqual(client.post('/api/maintenance/game', json={'action': 'join'}, headers=self.headers).status_code, 200)
            self.pause('available')
            self.assertEqual(client.post('/api/capacity', json={}).status_code, 401)
            with client.get('/login') as response:
                self.assertEqual(response.status_code, 200)


if __name__ == '__main__':
    unittest.main()
