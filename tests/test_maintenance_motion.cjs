const test = require('node:test');
const assert = require('node:assert/strict');
const {advance, Motion, defaults} = require('../maintenance/motion.js');
const player = (id, x=90, y=300) => ({id, x, y, dx: 1, dy: 0, role: 'police', jailed: false});
const snapshot = (players, time, changes={}) => ({players, you: 'me', server_time: time, room_id: 'room', round: 1, paused: false, obstacles: defaults.obstacles, motion: defaults, ...changes});

test('prediction moves every display frame without a network response', () => {
  const motion = new Motion();
  motion.receive(snapshot([player('me')], 0), 0);
  motion.setInput({dx: 1, dy: 0}, 0);
  const positions = [];
  for (let ms=1000/120; ms<=200; ms+=1000/120) positions.push(motion.frame(ms)[0].x);
  assert.ok(positions.length >= 23);
  for (let n=1; n<positions.length; n++) {
    assert.ok(positions[n] > positions[n-1], 'each frame advances instead of holding for a packet');
    assert.ok(positions[n]-positions[n-1] < 2, 'no 200 ms jump');
  }
  assert.ok(Math.abs(positions.at(-1)-126) < 2);
});

test('release and reversal respond on the next frame', () => {
  const motion = new Motion();
  motion.receive(snapshot([player('me')], 0), 0);
  motion.setInput({dx: 1, dy: 0}, 0);
  for (let t=16; t<=160; t+=16) motion.frame(t);
  motion.setInput({dx: 0, dy: 0}, 160);
  const stop = motion.frame(176)[0].x;
  assert.equal(motion.frame(192)[0].x, stop);
  motion.setInput({dx: -1, dy: 0}, 192);
  assert.ok(motion.frame(208)[0].x < stop);
});

test('other players interpolate smoothly between sparse packets', () => {
  const motion = new Motion();
  motion.receive(snapshot([player('other', 90)], 0, {you: null}), 0);
  motion.receive(snapshot([player('other', 126)], .2, {you: null}), 200);
  const frames = [200, 216, 232, 248, 264, 280, 296, 312, 328, 344, 360, 376, 392].map(t => motion.frame(t)[0].x);
  for (let n=1; n<frames.length; n++) {
    assert.ok(frames[n] > frames[n-1]);
    assert.ok(frames[n]-frames[n-1] <= 3);
  }
});

test('frame rate cannot change speed, diagonals or wall collisions', () => {
  const travel = fps => {
    let p = {x: 90, y: 300};
    for (let n=0; n<fps; n++) p=advance(p, {dx: 1, dy: 0}, 1/fps);
    return p;
  };
  for (const fps of [30, 60, 120]) assert.ok(Math.abs(travel(fps).x-270)<1e-6);
  const diagonal = advance({x: 90, y: 300}, {dx: 1, dy: 1}, .1);
  assert.ok(Math.abs(Math.hypot(diagonal.x-90, diagonal.y-300)-18)<1e-6);
  assert.ok(advance({x:120, y:150}, {dx:1,dy:0}, .6).x<=165);
});

test('jail, rescue and a new round snap to authoritative locations', () => {
  const motion = new Motion();
  motion.receive(snapshot([player('me')], 0), 0);
  motion.setInput({dx:1,dy:0}, 0);
  motion.receive(snapshot([{...player('me',475,270), jailed:true}], .1), 100);
  assert.equal(motion.frame(116)[0].x, 475);
  motion.receive(snapshot([player('me',590,300)], .2), 200);
  assert.equal(motion.frame(200)[0].x, 590);
  motion.receive(snapshot([player('me')], .3, {round:2, paused:true}), 300);
  assert.equal(motion.frame(350)[0].x, 90);
});

test('stale responses are ignored and a disconnected client stops prediction', () => {
  const motion = new Motion();
  motion.receive(snapshot([player('me')], 1), 0);
  assert.equal(motion.receive(snapshot([player('me',800)], .5), 100), false);
  motion.setInput({dx:1,dy:0}, 0);
  for (let t=16; t<=700; t+=16) motion.frame(t);
  const stopped = motion.frame(720)[0].x;
  assert.equal(motion.frame(736)[0].x, stopped);
});
