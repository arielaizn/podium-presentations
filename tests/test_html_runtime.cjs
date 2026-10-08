/* Dependency-free behavioral unit tests. Browser visual testing is still required. */
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const code = fs.readFileSync(path.join(__dirname,'../assets/html-runtime.js'),'utf8');
function fixture(rtl=false) {
  const listeners={},globalListeners={};
  const controls=Object.fromEntries(['prev-slide','next-slide','slide-status','fullscreen'].map(id=>[id,{disabled:false,textContent:'',addEventListener(k,fn){this[k]=fn}}]));
  const slides=Array.from({length:3},()=>({classList:{toggle(k,v){this[k]=v}},setAttribute(k,v){this[k]=v},scrollTop:20}));
  const document={documentElement:{dir:rtl?'rtl':'ltr'},querySelectorAll(){return slides},getElementById(id){return controls[id]},addEventListener(k,f){listeners[k]=f},dispatchEvent(){}};
  const location={hash:''},history={replaceState(a,b,h){location.hash=h}};
  vm.runInNewContext(code,{document,location,history,CustomEvent:class{},addEventListener(k,f){globalListeners[k]=f}});
  const key=(key,tag='MAIN',interactive=false)=>listeners.keydown({key,target:{tagName:tag,closest:()=>interactive,isContentEditable:false},preventDefault(){}});
  return {slides,controls,key,location,listeners};
}
for (const rtl of [false,true]) {
  const f=fixture(rtl);
  assert.equal(f.controls['slide-status'].textContent,'1 / 3');
  f.key(rtl?'ArrowLeft':'ArrowRight','SPAN');
  assert.equal(f.controls['slide-status'].textContent,'2 / 3');
  assert.equal(f.slides[0].inert,true);assert.equal(f.slides[1].inert,false);
  f.key('End');assert.equal(f.controls['slide-status'].textContent,'3 / 3');
  assert.equal(f.controls['next-slide'].disabled,true);
  f.key('Home');assert.equal(f.controls['slide-status'].textContent,'1 / 3');
  f.key(' ','INPUT',true);assert.equal(f.controls['slide-status'].textContent,'1 / 3');
  f.controls['next-slide'].click();assert.equal(f.location.hash,'#slide-2');
  f.controls['prev-slide'].click();assert.equal(f.controls['prev-slide'].disabled,true);
  f.listeners.touchstart({touches:[{clientX:200,clientY:100}]});
  f.listeners.touchend({changedTouches:[{clientX:rtl?300:100,clientY:105}]});
  assert.equal(f.controls['slide-status'].textContent,'2 / 3');
}
console.log('PASS: LTR/RTL controls, keyboard, touch, bounds, hash, hidden-slide accessibility and interactive fields');
