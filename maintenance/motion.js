/* Shared pure movement math: animation never waits for the network. */
(function (root, factory) {
  const motion = factory();
  if (typeof module === 'object' && module.exports) module.exports = motion;
  else root.ChorrMotion = motion;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const defaults = {speed: 180, width: 1000, height: 600, radius: 15, edge: 18, jail: [450, 240, 100, 120], obstacles: [[180,125,145,70],[675,125,145,70],[180,405,145,70],[675,405,145,70]]};
  function advance(position, input, seconds, rules = defaults) {
    let {x, y} = position;
    const dx = input.dx || 0, dy = input.dy || 0;
    if (!dx && !dy) return {x, y};
    const length = Math.max(1, Math.hypot(dx, dy));
    const boxes = [...rules.obstacles, rules.jail];
    const blocked = (px, py) => boxes.some(([bx, by, w, h]) => bx-rules.radius < px && px < bx+w+rules.radius && by-rules.radius < py && py < by+h+rules.radius);
    // Small steps prevent prediction from tunnelling through a house or jail.
    for (let remaining = Math.max(0, Math.min(seconds, .6)); remaining > 1e-9;) {
      const dt = Math.min(1/60, remaining); remaining -= dt;
      const nx = Math.min(rules.width-rules.edge, Math.max(rules.edge, x + dx/length*rules.speed*dt));
      const ny = Math.min(rules.height-rules.edge, Math.max(rules.edge, y + dy/length*rules.speed*dt));
      if (!blocked(nx, y)) x = nx;
      if (!blocked(x, ny)) y = ny;
    }
    return {x, y};
  }

  class Motion {
    constructor() {
      this.rules = defaults;
      this.players = [];
      this.history = new Map();
      this.input = {dx: 0, dy: 0};
      this.local = null;
      this.error = {x: 0, y: 0};
      this.lastFrame = null;
      this.lastReceived = -Infinity;
    }
    setInput(input, now) {
      this.frame(now); // Integrate the previous direction up to this input edge.
      this.input = {dx: input.dx || 0, dy: input.dy || 0};
      this.error = {x: 0, y: 0};
    }
    receive(snapshot, now, roundTrip = 0, command = null) {
      if (snapshot.room_id === this.room && snapshot.server_time < this.serverTime) return false;
      const reset = snapshot.room_id !== this.room || snapshot.round !== this.round || snapshot.you !== this.you;
      this.frame(now);
      if (reset) { this.history.clear(); this.local = null; this.error = {x: 0, y: 0}; }
      this.room = snapshot.room_id; this.round = snapshot.round; this.you = snapshot.you;
      this.serverTime = snapshot.server_time;
      this.paused = snapshot.paused;
      this.rules = {...defaults, ...snapshot.motion, obstacles: snapshot.obstacles || defaults.obstacles};
      this.lastReceived = now;
      this.interpolationDelay = 200+Math.min(Math.max(roundTrip, 0), 200);
      const previous = new Map(this.players.map(p => [p.id, p]));
      this.players = snapshot.players;
      for (const p of this.players) {
        const old = previous.get(p.id);
        const teleport = reset || !old || old.jailed !== p.jailed || old.role !== p.role || Math.hypot(p.x-old.x, p.y-old.y) > 120;
        const samples = teleport ? [] : this.history.get(p.id) || [];
        samples.push({...p, time: now});
        this.history.set(p.id, samples.slice(-8));
        if (p.id === this.you) {
          if (!this.local || teleport || p.jailed || snapshot.paused) {
            this.local = {x: p.x, y: p.y}; this.error = {x: 0, y: 0};
          } else if (!command || command.action !== 'move' || (command.dx === this.input.dx && command.dy === this.input.dy)) {
            const target = advance(p, this.input, Math.min(roundTrip/2000, .1), this.rules);
            this.error = {x: target.x-this.local.x, y: target.y-this.local.y};
          }
        }
      }
      for (const id of this.history.keys()) if (!this.players.some(p => p.id === id)) this.history.delete(id);
      return true;
    }
    frame(now) {
      const dt = this.lastFrame === null ? 0 : Math.min(Math.max((now-this.lastFrame)/1000, 0), .05);
      this.lastFrame = now;
      const me = this.players.find(p => p.id === this.you);
      if (this.local && me && !me.jailed && !this.paused) {
        if (now-this.lastReceived <= 600) this.local = advance(this.local, this.input, dt, this.rules);
        // Resolve small network disagreement gradually, rather than snapping back.
        const blend = 1-Math.exp(-dt/0.12);
        const correction = {dx: this.error.x*blend, dy: this.error.y*blend};
        const length = Math.hypot(correction.dx, correction.dy);
        if (length > 0) {
          const corrected = advance(this.local, {dx: correction.dx/length, dy: correction.dy/length}, length/this.rules.speed, this.rules);
          this.error.x -= corrected.x-this.local.x; this.error.y -= corrected.y-this.local.y;
          this.local = corrected;
        }
      }
      return this.players.map(p => {
        if (p.id === this.you && this.local) return {...p, ...this.local};
        const samples = this.history.get(p.id) || [];
        const at = now-this.interpolationDelay; // Buffer one network interval, including latency.
        let before = samples[0] || p, after = before;
        for (const sample of samples) {
          if (sample.time <= at) before = sample;
          if (sample.time >= at) { after = sample; break; }
          after = sample;
        }
        if (before.time < at && after.time > before.time) {
          const mix = Math.min(1, Math.max(0, (at-before.time)/(after.time-before.time)));
          return {...p, x: before.x+(after.x-before.x)*mix, y: before.y+(after.y-before.y)*mix};
        }
        const extra = Math.min(Math.max((at-after.time)/1000, 0), .1);
        return {...p, ...advance(after, after, extra, this.rules)};
      });
    }
  }
  return {advance, Motion, defaults};
});
