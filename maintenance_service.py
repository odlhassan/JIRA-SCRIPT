"""Database-free deployment gate and anonymous, server-authoritative playground."""
from __future__ import annotations

import argparse
import json
import logging
import math
import os
from pathlib import Path
import secrets
import threading
import time
from urllib.parse import urlsplit

from flask import Flask, jsonify, request, send_from_directory
from werkzeug.wsgi import ClosingIterator

ROOT = Path(__file__).resolve().parent
_jobs_lock = threading.Lock()
_jobs = 0


def start_background_job(target, args=()):
    """Count a report job before starting it so deployment can drain it."""
    global _jobs
    with _jobs_lock:
        _jobs += 1

    def run():
        global _jobs
        try:
            target(*args)
        finally:
            with _jobs_lock:
                _jobs -= 1

    thread = threading.Thread(target=run, daemon=True)
    try:
        thread.start()
    except BaseException:
        with _jobs_lock:
            _jobs -= 1
        raise
    return thread


def marker_path(base=ROOT):
    default = Path('/home/data/deployment-maintenance.json') if os.getenv('WEBSITE_SITE_NAME') else base / '.deployment-maintenance.json'
    return Path(os.getenv('EPR_MAINTENANCE_FILE', str(default)))


def write_status(path, phase, message, deployment=''):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + secrets.token_hex(6) + '.tmp')
    temporary.write_text(json.dumps({'phase': phase, 'message': message[:240], 'deployment': deployment, 'updated_at': time.time()}), encoding='utf-8')
    os.replace(temporary, path)


class Playground:
    WIDTH, HEIGHT = 1000, 600
    SPEED, TTL, LIMIT = 180, 20, 40
    JAIL = (450, 240, 100, 120)
    OBSTACLES = ((180, 125, 145, 70), (675, 125, 145, 70), (180, 405, 145, 70), (675, 405, 145, 70))

    def __init__(self, clock=time.monotonic):
        self.clock = clock
        self.lock = threading.RLock()
        self.players = {}
        self.round = 1
        self.resume_at = 0
        self.room_id = secrets.token_hex(8)
        self.message = 'First to join becomes police. Everyone else: run!'
        self.updated = clock()

    def _police(self):
        return [p for p in self.players.values() if p['role'] == 'police']

    def _spawn(self, p):
        badge_index = sum(cop['order'] < p['order'] for cop in self._police())
        p['x'], p['y'] = (90, 300 + badge_index * 65) if p['role'] == 'police' else (900, 75 + (p['order'] * 83) % 450)
        p['jailed'] = False
        p['dx'], p['dy'], p['input_at'] = 0, 0, self.clock()
        p['immune_until'] = self.clock() + 3

    def _tick(self):
        now = self.clock()
        begin = self.updated
        self.updated = now
        for token, p in list(self.players.items()):
            if now - p['seen'] > self.TTL:
                del self.players[token]
        if self.players and not self._police():
            p = min(self.players.values(), key=lambda v: v['order'])
            p['role'], p['request'] = 'police', False
            self._spawn(p)
        if self.resume_at:
            if now < self.resume_at:
                return
            self.resume_at = 0
            self.round += 1
            for p in self.players.values():
                self._spawn(p)
            self.message = 'New round. Run, rescue, repeat!'
            begin = now
        # Integrate elapsed time in small collision-safe steps. A 200 ms poll must
        # cover the same distance as twelve 60 Hz steps, even with only one player.
        cursor = begin
        end = min(now, begin + .6)
        while cursor < end - 1e-9:
            following = min(end, cursor + 1 / 60)
            self._move_players(cursor, following)
            self._contacts(following)
            if self.resume_at:
                break
            cursor = following
        self._contacts(now)

    def _move_players(self, begin, end):
        for p in self.players.values():
            if p['jailed']:
                continue
            input_at = p.get('input_at', 0)
            dt = max(0, min(end, input_at + .6) - max(begin, input_at))
            if dt == 0:
                continue
            dx, dy = p.get('dx', 0), p.get('dy', 0)
            length = max(math.hypot(dx, dy), 1)
            x = min(982, max(18, p['x'] + dx / length * self.SPEED * dt))
            y = min(582, max(18, p['y'] + dy / length * self.SPEED * dt))
            boxes = (*self.OBSTACLES, self.JAIL)
            def blocked(px, py):
                return any(bx - 15 < px < bx + bw + 15 and by - 15 < py < by + bh + 15 for bx, by, bw, bh in boxes)
            if not blocked(x, p['y']):
                p['x'] = x
            if not blocked(p['x'], y):
                p['y'] = y

    def _contacts(self, now):
        runners = [p for p in self.players.values() if p['role'] == 'runner']
        for p in runners:
            if not p['jailed'] and now >= p['immune_until'] and any(math.hypot(p['x'] - cop['x'], p['y'] - cop['y']) < 34 for cop in self._police()):
                p['jailed'] = True
                p['x'], p['y'] = 475 + (p['order'] % 3) * 24, 270 + (p['order'] % 4) * 20
                self.message = p['name'] + ' was caught! Touch the jail to rescue them.'
        free = [p for p in runners if not p['jailed']]
        if any(425 <= p['x'] <= 575 and 215 <= p['y'] <= 385 for p in free):
            for p in runners:
                if p['jailed']:
                    self._spawn(p)
                    p['x'], p['y'] = 590, 300
                    self.message = 'Jailbreak! Rescued players have a 3-second head start.'
        if runners and all(p['jailed'] for p in runners) and not self.resume_at:
            self.resume_at = now + 5
            self.message = 'Police caught everyone! Next round in 5 seconds.'

    def join(self, name, token=''):
        with self.lock:
            self._tick()
            if token in self.players:
                self.players[token]['seen'] = self.clock()
                return token
            if len(self.players) >= self.LIMIT:
                raise ValueError('This playground is full (40 players). Please try again shortly.')
            token = secrets.token_urlsafe(24)
            order = time.time_ns()
            p = {'id': secrets.token_hex(5), 'name': str(name or 'Guest').strip()[:18] or 'Guest', 'role': 'runner' if self._police() else 'police', 'order': order, 'seen': self.clock(), 'request': False, 'dx': 0, 'dy': 0}
            self._spawn(p)
            self.players[token] = p
            return token

    def command(self, token, data):
        with self.lock:
            self._tick()
            if token not in self.players:
                raise LookupError('Please rejoin the playground.')
            p = self.players[token]
            p['seen'] = self.clock()
            action = data.get('action', 'move')
            if action == 'move':
                dx, dy = data.get('dx', 0), data.get('dy', 0)
                if not all(isinstance(v, (int, float)) and math.isfinite(v) and -1 <= v <= 1 for v in (dx, dy)):
                    raise ValueError('Invalid movement.')
                p['dx'], p['dy'], p['input_at'] = dx, dy, self.clock()
            elif action == 'request':
                if p['role'] != 'runner' or p['jailed'] or len(self._police()) >= 2:
                    raise ValueError('Only free runners can request an open police place.')
                p['request'] = True
            elif action in ('accept', 'decline'):
                if p['role'] != 'police':
                    raise ValueError('Only police can decide team requests.')
                candidate = next((v for v in self.players.values() if v['id'] == data.get('player_id')), None)
                if not candidate or not candidate['request']:
                    raise ValueError('This request is no longer available.')
                if action == 'accept':
                    if len(self._police()) >= 2 or candidate['jailed']:
                        raise ValueError('Police team is full or this player is jailed.')
                    candidate['role'] = 'police'
                    self._spawn(candidate)
                candidate['request'] = False
            elif action == 'leave':
                del self.players[token]
            else:
                raise ValueError('Unknown action.')

    def snapshot(self, token=''):
        with self.lock:
            self._tick()
            if token in self.players:
                self.players[token]['seen'] = self.clock()
            now = self.clock()
            paused = self.resume_at > now
            players = []
            for p in self.players.values():
                row = {k: p[k] for k in ('id', 'name', 'role', 'x', 'y', 'jailed', 'request')}
                moving = not paused and not p['jailed'] and now - p.get('input_at', 0) <= .6
                row.update(dx=p.get('dx', 0) if moving else 0, dy=p.get('dy', 0) if moving else 0)
                players.append(row)
            return {'players': players, 'you': self.players.get(token, {}).get('id'), 'round': self.round, 'message': self.message, 'obstacles': self.OBSTACLES,
                    'room_id': self.room_id, 'server_time': self.updated, 'paused': paused,
                    'motion': {'speed': self.SPEED, 'width': self.WIDTH, 'height': self.HEIGHT, 'radius': 15, 'edge': 18, 'jail': self.JAIL}}


class MaintenanceGateway:
    """WSGI gate outside app authentication; active responses and report jobs drain."""
    def __init__(self, app=None, *, factory=None, base=ROOT, marker=None, shortcut_reader=None):
        self.core = app
        self.shortcut_reader = shortcut_reader
        self.base = Path(base)
        self.marker = Path(marker) if marker else marker_path(self.base)
        revision_file = self.base / 'DEPLOYMENT_REVISION'
        self.revision = revision_file.read_text().strip() if revision_file.exists() else 'local'
        self.lock = threading.RLock()
        self.active = 0
        self.failed = False
        self.game = Playground()
        self.front = self._front_app()
        if factory:
            def warm():
                try:
                    self.core = factory()
                except Exception:
                    self.failed = True
                    logging.exception('Application startup failed; maintenance remains enabled')
            threading.Thread(target=warm, daemon=True, name='app-warmup').start()

    def status(self):
        try:
            state = json.loads(self.marker.read_text(encoding='utf-8'))
            if not isinstance(state, dict) or state.get('phase') not in ('available', 'draining', 'deploying', 'checking', 'failed'):
                raise ValueError('Invalid deployment state')
        except FileNotFoundError:
            state = {'phase': 'available', 'message': 'The app is available.'}
        except (OSError, ValueError):
            state = {'phase': 'failed', 'message': 'Availability is being checked. Please wait.'}
        ready = self.core is not None and not self.failed
        if not ready:
            state = {**state, 'phase': 'failed' if self.failed else 'checking', 'message': 'The app is starting. Please wait.' if not self.failed else 'Deployment needs attention. The app is safely paused.'}
        reader = self.shortcut_reader or getattr(self.core, 'extensions', {}).get('epr_game_shortcut')
        shortcut = reader() if reader else None
        with self.lock, _jobs_lock:
            return {**state, 'available': ready and state['phase'] == 'available', 'ready': ready, 'revision': self.revision, 'active_requests': self.active, 'background_jobs': _jobs, 'shortcut': shortcut}

    def _front_app(self):
        app = Flask('maintenance', static_folder=None)
        app.config['MAX_CONTENT_LENGTH'] = 2048

        @app.after_request
        def headers(response):
            response.headers['Cache-Control'] = 'no-store'
            response.headers['X-Content-Type-Options'] = 'nosniff'
            response.headers['X-Frame-Options'] = 'SAMEORIGIN'
            response.headers['Content-Security-Policy'] = "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'self'; base-uri 'none'; form-action 'self'"
            return response

        @app.get('/maintenance')
        def page():
            return send_from_directory(ROOT / 'maintenance', 'index.html')

        @app.get('/maintenance/<asset>')
        def asset(asset):
            if asset not in ('game.js', 'game.css', 'motion.js', 'app-shortcuts.js'):
                return '', 404
            return send_from_directory(ROOT if asset == 'app-shortcuts.js' else ROOT / 'maintenance', asset)

        @app.get('/healthz')
        def health():
            return jsonify(ok=True)

        @app.get('/api/maintenance/status')
        def status():
            return jsonify(self.status())

        @app.get('/api/maintenance/game')
        def state():
            return jsonify(self.game.snapshot(request.cookies.get('epr_play', '')))

        @app.post('/api/maintenance/game')
        def command():
            origin = urlsplit(request.headers.get('Origin', ''))
            if origin.netloc != request.host or origin.scheme not in ('http', 'https'):
                return jsonify(error='Use the playground on this site.'), 403
            data = request.get_json(silent=True)
            if not isinstance(data, dict):
                return jsonify(error='Expected a JSON object.'), 400
            token = request.cookies.get('epr_play', '')
            try:
                if data.get('action') == 'join':
                    token = self.game.join(data.get('name', ''), token)
                else:
                    self.game.command(token, data)
                response = jsonify(self.game.snapshot(token))
                response.set_cookie('epr_play', token, httponly=True, secure=request.is_secure, samesite='Strict', max_age=3600, path='/api/maintenance/game')
                return response
            except LookupError as exc:
                return jsonify(error=str(exc)), 401
            except ValueError as exc:
                return jsonify(error=str(exc)), 400

        @app.route('/', defaults={'path': ''}, methods=['GET', 'HEAD', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'])
        @app.route('/<path:path>', methods=['GET', 'HEAD', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'])
        def paused(path):
            if request.method in ('GET', 'HEAD') and not request.path.startswith('/api/') and ('text/html' in request.headers.get('Accept', '') or request.headers.get('Sec-Fetch-Dest') == 'document'):
                response = send_from_directory(ROOT / 'maintenance', 'index.html')
                response.status_code = 503
            else:
                response = jsonify(error='App temporarily unavailable during deployment.', maintenance=True)
                response.status_code = 503
            response.headers['Retry-After'] = '5'
            return response

        return app

    def __call__(self, environ, start_response):
        path = environ.get('PATH_INFO', '')
        if path in ('/maintenance', '/healthz', '/api/maintenance/status', '/api/maintenance/game') or path.startswith('/maintenance/'):
            return self.front(environ, start_response)
        with self.lock:
            if not self.status()['available']:
                return self.front(environ, start_response)
            self.active += 1

        finished = False

        def finish():
            nonlocal finished
            with self.lock:
                if not finished:
                    self.active -= 1
                    finished = True
        try:
            return ClosingIterator(self.core(environ, start_response), finish)
        except BaseException:
            finish()
            raise


def main():
    parser = argparse.ArgumentParser(description='Control deployment maintenance locally (no Jira or database writes).')
    parser.add_argument('phase', choices=['available', 'draining', 'deploying', 'checking', 'failed'])
    parser.add_argument('--message', default='An update is in progress. Play a round while we get ready.')
    parser.add_argument('--file', type=Path, default=marker_path())
    args = parser.parse_args()
    write_status(args.file, args.phase, args.message if args.phase != 'available' else 'The app is available.')
    print(f'{args.phase}: {args.file}')


if __name__ == '__main__':
    main()
