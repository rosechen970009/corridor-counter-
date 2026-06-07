<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>Vertical Grid Analyzer</title>
<style>
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;padding:14px;font-family:-apple-system,BlinkMacSystemFont,"Noto Sans TC",Arial,sans-serif;background:#eefcff;color:#173b3f}
.wrap{max-width:900px;margin:auto}
h1{text-align:center;color:#00796b;margin:8px 0 12px;font-size:26px}.panel{background:#fff;border-radius:20px;padding:14px;box-shadow:0 6px 18px #0002;margin-bottom:12px}
label{font-weight:900;font-size:14px;display:block;margin:8px 0 4px}input,select{width:100%;font-size:18px;padding:12px;border-radius:12px;border:1px solid #cbdede;background:#fff}.grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
video{width:100%;max-height:52vh;background:#000;border-radius:14px}.time{font-size:22px;font-weight:900;text-align:center;background:#e2f6f4;border-radius:14px;padding:10px;margin-top:8px}.btnGrid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:10px}
button{border:0;border-radius:16px;color:white;font-weight:900;font-size:18px;min-height:58px;touch-action:manipulation;cursor:pointer}.mark{background:#1976d2}.save{background:#00897b}.warn{background:#fb8c00}.danger{background:#666}.step{background:#6a1b9a}.small{font-size:15px;min-height:48px}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:10px}.card{text-align:center;background:#f8fbfb;border-radius:16px;padding:10px;border:1px solid #e1eeee}.card .num{font-size:30px;font-weight:1000;color:#00796b}.checks{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:10px}.check{background:#f8fbfb;border-radius:12px;padding:10px;font-weight:800}.check input{width:auto;transform:scale(1.25);margin-right:6px}.note{font-size:14px;color:#555;line-height:1.5;margin-top:10px}table{width:100%;border-collapse:collapse;background:#fff;font-size:14px}th{background:#d8f3ee;color:#234;padding:8px;white-space:nowrap}td{border-bottom:1px solid #e7eeee;padding:8px;text-align:center;white-space:nowrap}.tableWrap{overflow:auto;margin-top:10px;border-radius:14px}.ok{color:#087b31}.na{color:#999}
@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}.btnGrid{grid-template-columns:1fr 1fr}.stats{grid-template-columns:1fr 1fr}.checks{grid-template-columns:1fr 1fr}button{font-size:16px}.card .num{font-size:26px}}
</style>
</head>
<body>
<div class="wrap">
<h1>Vertical Grid Analyzer</h1>
<div class="panel">
  <label>上傳影片</label>
  <input id="videoFile" type="file" accept="video/*">
  <video id="video" controls playsinline></video>
  <div id="videoTime" class="time">影片時間：0.00 s</div>
</div>

<div class="panel">
  <div class="grid">
    <div><label>小鼠編號</label><input id="mouseId" placeholder="M01"></div>
    <div><label>Trial</label><select id="trial"><option>1</option><option>2</option><option>3</option></select></div>
    <div><label>備註</label><input id="note" placeholder="Day 14 / pre-test"></div>
  </div>

  <div class="btnGrid">
    <button class="mark" onclick="markStart()">1 Start<br><span class="small">開始/放上格子</span></button>
    <button class="mark" onclick="markTurn()">2 Turn<br><span class="small">完成轉頭</span></button>
    <button class="mark" onclick="markLanding()">3 Landing<br><span class="small">四肢落地</span></button>
  </div>

  <div class="stats">
    <div class="card"><div>Start</div><div id="startT" class="num na">--</div></div>
    <div class="card"><div>Turn</div><div id="turnT" class="num na">--</div></div>
    <div class="card"><div>Landing</div><div id="landingT" class="num na">--</div></div>
    <div class="card"><div>Hindlimb Steps</div><div id="steps" class="num">0</div></div>
  </div>

  <div class="stats">
    <div class="card"><div>Turn Time</div><div id="turnTime" class="num">--</div></div>
    <div class="card"><div>Descent Time</div><div id="descentTime" class="num">--</div></div>
    <div class="card"><div>Total Time</div><div id="totalTime" class="num">--</div></div>
    <div class="card"><div>Status</div><div id="status" class="num">--</div></div>
  </div>

  <div class="btnGrid">
    <button class="step" onclick="addStep()">H 後腳步數 +1</button>
    <button class="warn" onclick="minusStep()">後腳步數 -1</button>
    <button class="danger" onclick="resetTrial()">重設目前 Trial</button>
  </div>

  <div class="checks">
    <label class="check"><input id="failedTurn" type="checkbox">Failed turn</label>
    <label class="check"><input id="failedDescent" type="checkbox">Failed descent</label>
    <label class="check"><input id="assisted" type="checkbox">Assisted by hand</label>
    <label class="check"><input id="exclude" type="checkbox">Exclude</label>
  </div>

  <div class="btnGrid">
    <button class="save" onclick="saveTrial()">S 儲存 Trial</button>
    <button class="save" onclick="exportExcel()">匯出 Excel (.xls)</button>
    <button class="danger" onclick="clearAll()">清空全部紀錄</button>
  </div>

  <div class="note">
    後腳步數定義：完成轉頭後開始計算，在下爬過程中左後腳＋右後腳踩格子的總次數，直到四肢落地。<br>
    快捷鍵：Space 播放/暫停；1=Start；2=Turn；3=Landing；H=後腳步數+1；S=儲存。
  </div>

  <div class="tableWrap">
    <table>
      <thead><tr><th>Time</th><th>Mouse</th><th>Trial</th><th>Start</th><th>Turn</th><th>Landing</th><th>Turn time</th><th>Descent time</th><th>Total time</th><th>Hindlimb steps</th><th>Failed turn</th><th>Failed descent</th><th>Assisted</th><th>Exclude</th><th>Note</th></tr></thead>
      <tbody id="records"></tbody>
    </table>
  </div>
</div>
</div>

<script>
let start=null, turn=null, landing=null, hindSteps=0, rows=[];
const $=id=>document.getElementById(id);
const v=$('video');
function now(){return Number(v.currentTime||0)}
function fmt(x){return x===null||isNaN(x)?'--':Number(x).toFixed(2)}
function calc(a,b){return (a!==null && b!==null && b>=a)?(b-a):null}
function update(){
 $('videoTime').textContent='影片時間：'+fmt(now())+' s';
 $('startT').textContent=fmt(start); $('turnT').textContent=fmt(turn); $('landingT').textContent=fmt(landing); $('steps').textContent=hindSteps;
 $('turnTime').textContent=fmt(calc(start,turn)); $('descentTime').textContent=fmt(calc(turn,landing)); $('totalTime').textContent=fmt(calc(start,landing));
 let st='--'; if($('failedTurn').checked) st='Failed turn'; else if($('failedDescent').checked) st='Failed descent'; else if(start!==null&&turn!==null&&landing!==null) st='Success';
 $('status').textContent=st;
}
$('videoFile').addEventListener('change',e=>{const f=e.target.files[0]; if(f){v.src=URL.createObjectURL(f); v.load();}});
v.addEventListener('timeupdate',update); v.addEventListener('loadedmetadata',update); v.addEventListener('seeked',update);
function markStart(){start=now(); update()}
function markTurn(){turn=now(); update()}
function markLanding(){landing=now(); update()}
function addStep(){hindSteps++; update()}
function minusStep(){hindSteps=Math.max(0,hindSteps-1); update()}
function resetTrial(){start=null;turn=null;landing=null;hindSteps=0;['failedTurn','failedDescent','assisted','exclude'].forEach(id=>$(id).checked=false);update()}
function saveTrial(){
 rows.push({time:new Date().toLocaleString(),mouse:$('mouseId').value||'',trial:$('trial').value||'',start:fmt(start),turn:fmt(turn),landing:fmt(landing),turnTime:fmt(calc(start,turn)),descentTime:fmt(calc(turn,landing)),totalTime:fmt(calc(start,landing)),steps:hindSteps,failedTurn:$('failedTurn').checked?'Yes':'No',failedDescent:$('failedDescent').checked?'Yes':'No',assisted:$('assisted').checked?'Yes':'No',exclude:$('exclude').checked?'Yes':'No',note:$('note').value||''});
 render(); resetTrial();
}
function render(){
 $('records').innerHTML=rows.map(r=>`<tr><td>${r.time}</td><td>${r.mouse}</td><td>${r.trial}</td><td>${r.start}</td><td>${r.turn}</td><td>${r.landing}</td><td>${r.turnTime}</td><td>${r.descentTime}</td><td>${r.totalTime}</td><td>${r.steps}</td><td>${r.failedTurn}</td><td>${r.failedDescent}</td><td>${r.assisted}</td><td>${r.exclude}</td><td>${r.note}</td></tr>`).join('');
}
function exportExcel(){
 const table=`<table><tr><th>Time</th><th>Mouse</th><th>Trial</th><th>Start_s</th><th>Turn_s</th><th>Landing_s</th><th>Turn_time_s</th><th>Descent_time_s</th><th>Total_time_s</th><th>Hindlimb_steps</th><th>Failed_turn</th><th>Failed_descent</th><th>Assisted_by_hand</th><th>Exclude</th><th>Note</th></tr>${rows.map(r=>`<tr><td>${r.time}</td><td>${r.mouse}</td><td>${r.trial}</td><td>${r.start}</td><td>${r.turn}</td><td>${r.landing}</td><td>${r.turnTime}</td><td>${r.descentTime}</td><td>${r.totalTime}</td><td>${r.steps}</td><td>${r.failedTurn}</td><td>${r.failedDescent}</td><td>${r.assisted}</td><td>${r.exclude}</td><td>${r.note}</td></tr>`).join('')}</table>`;
 const blob=new Blob(['\ufeff<html><head><meta charset="UTF-8"></head><body>'+table+'</body></html>'],{type:'application/vnd.ms-excel'});
 const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='vertical_grid_results.xls'; a.click(); URL.revokeObjectURL(url);
}
function clearAll(){if(confirm('確定清空全部紀錄？')){rows=[];render();resetTrial();}}
['failedTurn','failedDescent','assisted','exclude'].forEach(id=>$(id).addEventListener('change',update));
document.addEventListener('keydown',e=>{
 const tag=(document.activeElement&&document.activeElement.tagName)||''; if(tag==='INPUT'||tag==='SELECT') return;
 const k=e.key.toLowerCase();
 if(e.code==='Space'){e.preventDefault(); if(v.paused)v.play(); else v.pause();}
 if(k==='1')markStart(); if(k==='2')markTurn(); if(k==='3')markLanding(); if(k==='h')addStep(); if(k==='s')saveTrial();
});
update();
</script>
</body>
</html>
