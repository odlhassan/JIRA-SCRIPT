/* Decorative ambience only. Authentication stays in the existing server OAuth flow. */
(() => {
  'use strict';
  const reduced=matchMedia('(prefers-reduced-motion: reduce)'), toggle=document.querySelector('#motion-toggle');
  const canvas=document.querySelector('#constellation'), ctx=canvas.getContext('2d'), root=document.documentElement;
  let paused=reduced.matches, frame=0, previous=0, time=0, width=0, height=0, points=[], ripples=[];
  let pointer={x:-1000,y:-1000}, parallax={x:0,y:0}, target={x:0,y:0};
  function setPaused(value){
    paused=value;document.body.classList.toggle('motion-paused',paused);
    toggle.setAttribute('aria-pressed',String(paused));
    toggle.setAttribute('aria-label',paused?'Resume background motion':'Pause background motion');
    toggle.title=paused?'Resume background motion':'Pause background motion';
    if(paused){cancelAnimationFrame(frame);frame=0;root.style.setProperty('--mx','0');root.style.setProperty('--my','0');document.querySelector('.connections')?.pauseAnimations?.();}
    else{document.querySelector('.connections')?.unpauseAnimations?.();start();}
  }
  function resize(){
    width=innerWidth;height=innerHeight;
    const ratio=Math.min(devicePixelRatio||1,1.5);canvas.width=width*ratio;canvas.height=height*ratio;
    if(!ctx)return;ctx.setTransform(ratio,0,0,ratio,0,0);
    points=Array.from({length:width<760?30:65},(_,i)=>({x:((i*.61803398875+.12)%1)*width,y:((i*.41421356237+.08)%1)*height,phase:i*1.77,size:i%7===0?1.5:.65+(i%3)*.2}));paint(0);
  }
  function paint(delta){
    if(!ctx)return;time+=delta;ctx.clearRect(0,0,width,height);
    const stars=points.map(p=>{
      let x=p.x+Math.sin(time*.16+p.phase)*17,y=p.y+Math.cos(time*.12+p.phase)*14;
      const dx=x-pointer.x,dy=y-pointer.y,dist=Math.hypot(dx,dy);
      if(!paused&&dist<150&&dist>0){const force=(1-dist/150)*14;x+=dx/dist*force;y+=dy/dist*force;}return {...p,x,y};
    });
    stars.forEach((p,i)=>{
      ctx.fillStyle=`rgba(130,154,193,${.25+Math.sin(time*.7+p.phase)*.13})`;ctx.beginPath();ctx.arc(p.x,p.y,p.size,0,Math.PI*2);ctx.fill();
      if(i%7===0){ctx.strokeStyle='#a3b4d42e';ctx.lineWidth=.65;ctx.beginPath();ctx.moveTo(p.x-4,p.y);ctx.lineTo(p.x+4,p.y);ctx.moveTo(p.x,p.y-4);ctx.lineTo(p.x,p.y+4);ctx.stroke();}
      for(let j=i+1;j<stars.length;j++){const q=stars[j],dist=Math.hypot(p.x-q.x,p.y-q.y);if(dist<135){ctx.strokeStyle=`rgba(134,161,200,${(1-dist/135)*.12})`;ctx.lineWidth=.6;ctx.beginPath();ctx.moveTo(p.x,p.y);ctx.lineTo(q.x,q.y);ctx.stroke();}}
    });
    ripples=ripples.filter(r=>r.age<1.7);ripples.forEach(r=>{r.age+=delta;ctx.strokeStyle=`rgba(135,152,213,${Math.max(0,1-r.age/1.7)*.16})`;ctx.lineWidth=.8;ctx.beginPath();ctx.arc(r.x,r.y,8+r.age*70,0,Math.PI*2);ctx.stroke();});
  }
  function animate(stamp){
    frame=0;if(paused||document.hidden)return;
    const elapsed=stamp-previous;
    // Canvas is capped at 30fps and device pixel ratio 1.5; CSS moves the larger forms.
    if(elapsed>=32){parallax.x+=(target.x-parallax.x)*.08;parallax.y+=(target.y-parallax.y)*.08;root.style.setProperty('--mx',parallax.x.toFixed(3));root.style.setProperty('--my',parallax.y.toFixed(3));paint(Math.min(elapsed/1000,.06));previous=stamp;}
    frame=requestAnimationFrame(animate);
  }
  function start(){if(!frame&&!paused&&!document.hidden){previous=performance.now();frame=requestAnimationFrame(animate);}}
  toggle.addEventListener('click',()=>setPaused(!paused));reduced.addEventListener('change',e=>setPaused(e.matches));
  addEventListener('resize',resize,{passive:true});
  addEventListener('pointermove',e=>{if(paused||e.pointerType==='touch')return;pointer={x:e.clientX,y:e.clientY};target={x:(e.clientX/width-.5)*2,y:(e.clientY/height-.5)*2};},{passive:true});
  document.addEventListener('pointerleave',()=>{pointer={x:-1000,y:-1000};target={x:0,y:0};});
  addEventListener('pointerdown',e=>{if(!paused&&ripples.length<4)ripples.push({x:e.clientX,y:e.clientY,age:0});},{passive:true});
  document.addEventListener('visibilitychange',()=>{if(document.hidden){cancelAnimationFrame(frame);frame=0;}else start();});
  resize();setPaused(paused);
  fetch('/api/jira-auth/status',{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error();return r.json();}).then(s=>{if(!s.configured){document.querySelector('#signin').setAttribute('aria-disabled','true');document.querySelector('#status').textContent='Atlassian sign-in needs administrator configuration.';}}).catch(()=>{document.querySelector('#status').textContent='Could not check sign-in availability. Please reload.';});
})();
