/* Nhạc và hiệu ứng tổng hợp tại thiết bị; không cần mạng. */
(()=>{
 let ctx,master,loop=null,step=0;
 let enabled=localStorage.getItem('stemSound')!=='off',music=localStorage.getItem('stemMusic')==='on';
 const pauseReasons=new Set(),active=new Set();
 function context(){if(!enabled)return false;try{ctx=ctx||new(window.AudioContext||window.webkitAudioContext)();if(!master){master=ctx.createGain();master.gain.value=.35;master.connect(ctx.destination);}if(ctx.state==='suspended')ctx.resume().catch(()=>{});return true;}catch(e){return false;}}
 function tone(freq,time,duration,volume,type='sine'){
  if(!ctx||!master)return;const o=ctx.createOscillator(),g=ctx.createGain();o.type=type;o.frequency.setValueAtTime(freq,time);g.gain.setValueAtTime(.0001,time);g.gain.exponentialRampToValueAtTime(Math.max(.0002,volume),time+.012);g.gain.exponentialRampToValueAtTime(.0001,time+duration);o.connect(g);g.connect(master);o.start(time);o.stop(time+duration+.02);active.add(o);o.onended=()=>active.delete(o);
 }
 function play(kind){if(!context())return;const at=ctx.currentTime+.01;
  const notes=kind==='unlock'?[523,659,784,1047,1319,1568]:kind==='success'?[659,784,1047]:kind==='error'?[330,262]:[784,1047];
  notes.forEach((n,i)=>tone(n,at+i*(kind==='click'?.035:.09),kind==='click'?.075:.23,kind==='click'?.045:.12,'triangle'));
 }
 const tune=[72,76,79,76,74,77,81,77,76,79,84,79,74,77,79,71,72,76,79,84,81,79,76,74,72,74,76,79,77,76,74,72];
 function stopLoop(){if(loop!==null){clearInterval(loop);loop=null;}}
 function refreshLoop(){stopLoop();if(!enabled||!music||document.hidden||pauseReasons.size||!ctx)return;const tick=()=>{const t=ctx.currentTime+.015,n=tune[step%tune.length];tone(440*Math.pow(2,(n-69)/12),t,.19,.065,'triangle');if(step%2===0)tone(130.81,t,.22,.08,'sine');if(step%4===0){tone(70,t,.075,.11,'sine');tone(1568,t+.12,.025,.015,'square');}step++;};tick();loop=setInterval(tick,240);}
 function sendToResource(){const f=document.getElementById('resourceFrame');if(f)f.contentWindow.postMessage({stemResource:true,type:'sound',enabled},location.origin==='null'?'*':location.origin);}
 function render(){const s=document.getElementById('soundToggle'),m=document.getElementById('musicToggle');if(s){s.textContent=enabled?'🔊 Âm thanh: Bật':'🔇 Âm thanh: Tắt';s.setAttribute('aria-pressed',String(enabled));}if(m){m.textContent=music?'🎵 Nhạc nền: Bật':'🎵 Nhạc nền: Tắt';m.setAttribute('aria-pressed',String(music));}}
 function stopAll(){active.forEach(o=>{try{o.stop();}catch(e){}});active.clear();}
 document.addEventListener('pointerdown',()=>{if(context())refreshLoop();},{once:true});
 document.addEventListener('click',e=>{
  const button=e.target.closest('button');if(!button)return;
  if(button.id==='soundToggle'){enabled=!enabled;localStorage.setItem('stemSound',enabled?'on':'off');if(!enabled){stopLoop();stopAll();if(master)master.gain.value=0;}else{context();if(master)master.gain.value=.35;play('success');refreshLoop();}sendToResource();render();return;}
  if(button.id==='musicToggle'){music=!music;localStorage.setItem('stemMusic',music?'on':'off');context();refreshLoop();render();return;}
  if(!button.disabled&&button.id!=='speakBtn')play('click');
 });
 document.addEventListener('play',e=>{if(e.target.tagName==='VIDEO'){pauseReasons.add('video');refreshLoop();}},true);
 document.addEventListener('pause',e=>{if(e.target.tagName==='VIDEO'){pauseReasons.delete('video');refreshLoop();}},true);
 document.addEventListener('ended',e=>{if(e.target.tagName==='VIDEO'){pauseReasons.delete('video');refreshLoop();}},true);
 document.addEventListener('visibilitychange',()=>{if(document.hidden){stopLoop();stopAll();}else refreshLoop();});
 window.STEM_SOUND={play,enabled:()=>enabled,learning(stage){pauseReasons.delete('video');stage===2?pauseReasons.add('resource'):pauseReasons.delete('resource');refreshLoop();sendToResource();}};
 render();
})();
