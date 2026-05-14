// ==UserScript==
// @name         pay_mena Ticket Reporter
// @namespace    http://tampermonkey.net/
// @version      12.0
// @description  pay_mena — прямые API запросы, X-Requested-With fix, retry, диагностика
// @match        *://th-managment.com/*
// @match        *://*.th-managment.com/*
// @match        *://webmanegment.com/*
// @match        *://*.webmanegment.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        GM_addStyle
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  var AGENTS = {
    'pay_mena_63':'Жанат Мейржан','pay_mena_106':'Шаймарданов Мадияр','pay_mena_110':'Айтбатыр Айжан',
    'pay_mena_116':'Кадиров Марлен','pay_mena_141':'Дарвакулов Дильмурат','pay_mena_184':'Чугункова Ульяна',
    'pay_mena_193':'Тарасов Вячеслав','pay_mena_201':'Ильясов Димаш','pay_mena_230':'Дзюбан Егор',
    'pay_mena_245':'Игорь Китайгородский','pay_mena_246':'Виктория Куклева','pay_mena_272':'Исмаилова Мария',
    'pay_mena_279':'Васинский Кирилл','pay_mena_323':'Кириенко Мария','pay_mena_330':'Гамза Кирилл',
    'pay_mena_333':'Романова Ильина','pay_mena_342':'Байзен Жамал','pay_mena_348':'Дюсенбаева Инара',
    'pay_mena_379':'Гайнетдинов Руслан','pay_mena_390':'Виктория Зотова','pay_mena_410':'Когай Екатерина',
    'pay_mena_429':'Савостенко Денис','pay_mena_431':'Лытнев Сергей','pay_mena_433':'Недилько Александра',
    'pay_mena_436':'Поляничев Михаил','pay_mena_441':'Питаев Анатолий','pay_mena_442':'Партенко Егор',
    'pay_mena_447':'Ефремова Вероника','pay_mena_448':'Голиней Богдан','pay_mena_459':'Байканов Ильдар',
    'pay_mena_470':'Доброва Надежда','pay_mena_471':'Имамутдинов Глеб','pay_mena_476':'Ткаченко Наталья',
    'pay_mena_486':'Бабак Виктория','pay_mena_495':'Костюхин Алексей','pay_mena_496':'Магдинов Алмаз',
    'pay_mena_509':'Маликова Елизавета','pay_mena_511':'Ушаков Артур','pay_mena_513':'Белов Федор',
    'pay_mena_516':'Платонов Кирилл','pay_mena_525':'Матева Анастасия','pay_mena_529':'Красюк Никита',
    'pay_mena_533':'Ларуков Давид','pay_mena_535':'Лежнев Макар','pay_mena_537':'Иваненко Олег',
    'pay_mena_542':'Багаутдинова Диана','pay_mena_543':'Васильев Максим','pay_mena_544':'Шестаков Александр',
    'pay_mena_545':'Дмитриев Никита','pay_mena_593':'Спесивцева Арина','pay_mena_601':'Матрук Сергей',
    'pay_mena_614':'Синельникова Софья','pay_mena_618':'Тарасов Илья','pay_mena_630':'Локтяжнов Артём',
    'pay_mena_645':'Алтухов Даниель','pay_mena_646':'Мухаметдинов Руслан Ринатович','pay_mena_652':'Яцкова Дарья',
    'pay_mena_653':'Кучкарева Алена','pay_mena_654':'Шагабутдинов Марсель Айратович',
    'pay_mena_658':'Богуцкая Арина Дмитриевна','pay_mena_659':'Соболев Вячеслав Владиславович',
    'pay_mena_662':'Сулейманов Ринат','pay_mena_663':'Силантьев Илья','pay_mena_665':'Мингазов Артур',
    'pay_mena_668':'Киреев Сергей','pay_mena_669':'Онуфрейчук Никита','pay_mena_670':'Бородай Иван',
    'pay_mena_671':'Мифтахова Айгуль','pay_mena_681':'Якубов Максад','pay_mena_682':'Кадаев Ильдар'
  };

  var INTERVAL_GREEN   = 50;
  var INTERVAL_YELLOW  = 60;
  var BREAK_LIMIT_MS   = 60 * 60 * 1000;
  var MIN_BREAK_SHOW   = 5 * 60 * 1000;
  var PAGE_WAIT_MS     = 1000;
  var DELAY_MS         = 350;
  var MAX_RETRIES      = 5;
  var RETRY_DELAY_MS   = 3500;

  var state = {
    agentLogin:null, agentName:null, shiftStart:null, shiftEnd:null,
    results:[], countryStat:{}, statusStat:{},
    currentPage:1, running:false, total:0, done:0, skipped:0,
    dbg:false
  };

  var agentMode='auto', dateMode='auto';
  var selectedLogin=null, parsedStart=null, parsedEnd=null;

  var allLogins = Object.keys(AGENTS).sort(function(a,b){
    return AGENTS[a].localeCompare(AGENTS[b],'ru');
  });

  GM_addStyle([
    '#tm{position:fixed;bottom:20px;right:20px;z-index:2147483647;width:390px;background:#0f1a2b;border:1px solid #1e3a5f;border-radius:14px;box-shadow:0 8px 32px rgba(0,0,0,.6);font-family:Arial,sans-serif;font-size:13px;color:#d0e8ff;overflow:hidden;}',
    '#tm.col #tmb{display:none;}#tm.col{width:200px;}',
    '#tmh{display:flex;align-items:center;justify-content:space-between;padding:11px 14px 10px;background:linear-gradient(90deg,#1a3a5c,#0f2540);border-bottom:1px solid #1e3a5f;cursor:pointer;user-select:none;}',
    '#tmht{display:flex;align-items:center;gap:7px;font-weight:700;font-size:13px;color:#7ecfff;}',
    '#tmht .dot{width:8px;height:8px;border-radius:50%;background:#3fa;display:inline-block;}',
    '#tmht .dot.busy{background:#f93;}',
    '#tmtog{background:none;border:none;color:#5a8ab0;font-size:17px;cursor:pointer;padding:0;}',
    '#tmtog:hover{color:#7ecfff;}',
    '#tmb{padding:14px 14px 12px;}',
    '.tms{margin-bottom:10px;}',
    '.tml-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:5px;}',
    '.tml{font-size:11px;color:#5a8ab0;font-weight:600;letter-spacing:.5px;text-transform:uppercase;}',
    '.tmg{display:flex;gap:4px;}',
    '.tmbtn{background:none;border:1px solid #1e3a5f;border-radius:5px;color:#5a8ab0;font-size:11px;cursor:pointer;padding:2px 7px;line-height:1.5;white-space:nowrap;}',
    '.tmbtn:hover{border-color:#2e75b6;color:#7ecfff;background:#0a2040;}',
    '.tmbtn.on{border-color:#2e75b6;background:#0a2040;color:#7ecfff;}',
    '.tmad{padding:9px 11px;border-radius:8px;font-size:12px;border:1px solid #1e3a5f;background:#071426;color:#5a8ab0;min-height:36px;display:flex;align-items:center;word-break:break-word;line-height:1.35;}',
    '.tmad.ok{border-color:#2a6a3a;background:#061510;color:#5fd876;}',
    '.tmad.err{border-color:#6a2a2a;background:#160606;color:#f06060;}',
    '#ags{width:100%;box-sizing:border-box;padding:8px 10px;border-radius:8px;border:1px solid #1e3a5f;background:#071426;color:#d0e8ff;font-size:12px;outline:none;}',
    '#ags:focus{border-color:#2e75b6;}#ags::placeholder{color:#3a5a7a;}',
    '#agd{position:absolute;top:calc(100% + 4px);left:0;right:0;background:#0d1f33;border:1px solid #1e3a5f;border-radius:8px;max-height:180px;overflow-y:auto;z-index:999;box-shadow:0 4px 16px rgba(0,0,0,.5);display:none;}',
    '#agd.open{display:block;}',
    '.agdi{padding:7px 10px;cursor:pointer;font-size:12px;border-bottom:1px solid #132030;color:#b0d4f0;}',
    '.agdi:last-child{border-bottom:none;}.agdi:hover{background:#1a3a5c;color:#7ecfff;}',
    '.agdi .lg{color:#4a7a9a;font-size:11px;margin-left:5px;}',
    '#agw{position:relative;}',
    '#agsel{margin-top:5px;padding:5px 10px;border-radius:6px;background:#0a2040;border:1px solid #2e75b6;font-size:12px;color:#7ecfff;display:flex;align-items:center;justify-content:space-between;}',
    '#agsel.h{display:none;}',
    '#agclr{background:none;border:none;color:#5a8ab0;cursor:pointer;font-size:13px;padding:0;}#agclr:hover{color:#f55;}',
    '#dti{width:100%;box-sizing:border-box;padding:8px 10px;border-radius:8px;border:1px solid #1e3a5f;background:#071426;color:#d0e8ff;font-size:12px;outline:none;}',
    '#dti:focus{border-color:#2e75b6;}#dti::placeholder{color:#3a5a7a;}',
    '#dthi{font-size:10px;color:#3a6a8a;margin-top:4px;}',
    '#dtp{margin-top:5px;padding:5px 10px;border-radius:6px;background:#0a2040;border:1px solid #1e3a5f;font-size:11px;color:#5ab0d0;display:none;}',
    '#dtp.ok{border-color:#2e75b6;color:#7ecfff;display:block;}#dtp.err{border-color:#8b2020;color:#f06060;display:block;}',
    '.tmback{width:100%;margin-top:7px;padding:5px 10px;border:1px dashed #1e3a5f;border-radius:6px;background:none;color:#4a7a9a;font-size:11px;cursor:pointer;}',
    '.tmback:hover{border-color:#2e75b6;color:#7ecfff;}',
    '#tmstart{width:100%;padding:10px;border:none;border-radius:8px;background:linear-gradient(135deg,#2e75b6,#1a5276);color:#fff;font-size:13px;font-weight:700;cursor:pointer;margin-top:4px;}',
    '#tmstart:hover:not(:disabled){opacity:.9;}#tmstart:disabled{opacity:.4;cursor:not-allowed;}',
    '#tmpeg{margin-top:10px;background:#071426;border-radius:6px;overflow:hidden;height:6px;display:none;}',
    '#tmpb{height:100%;width:0%;background:linear-gradient(90deg,#2e75b6,#7ecfff);transition:width .3s ease;}',
    '#tmst{margin-top:8px;font-size:11px;color:#4a8ab0;min-height:30px;line-height:1.5;word-break:break-word;}',
    '#agd::-webkit-scrollbar{width:4px;}',
    '#agd::-webkit-scrollbar-track{background:#0d1f33;}',
    '#agd::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:2px;}',
  ].join(''));

  var panel=document.createElement('div');
  panel.id='tm';
  panel.innerHTML=[
    '<div id="tmh">',
    '  <div id="tmht"><span class="dot" id="tmdot"></span>📊 pay_mena Reporter</div>',
    '  <button id="tmtog">▲</button>',
    '</div>',
    '<div id="tmb">',
    '  <div class="tms">',
    '    <div class="tml-row"><div class="tml">Агент</div>',
    '      <div class="tmg">',
    '        <button class="tmbtn on" id="agab">🔄 Авто</button>',
    '        <button class="tmbtn" id="agmb">✏️ Вручную</button>',
    '      </div></div>',
    '    <div id="agaw"><div class="tmad" id="agav">⏳ Определяю...</div></div>',
    '    <div id="agmw" style="display:none">',
    '      <div id="agw">',
    '        <input id="ags" type="text" placeholder="Введи имя или pay_mena_..." autocomplete="off">',
    '        <div id="agd"></div>',
    '      </div>',
    '      <div id="agsel" class="h"><span id="agst"></span><button id="agclr">✕</button></div>',
    '      <button class="tmback" id="agbb">← Вернуться к автодетекту</button>',
    '    </div>',
    '  </div>',
    '  <div class="tms">',
    '    <div class="tml-row"><div class="tml">Период</div>',
    '      <div class="tmg">',
    '        <button class="tmbtn on" id="dtab">🔄 Авто</button>',
    '        <button class="tmbtn" id="dtmb">✏️ Вручную</button>',
    '      </div></div>',
    '    <div id="dtaw"><div class="tmad" id="dtav">⏳ Определяю...</div></div>',
    '    <div id="dtmw" style="display:none">',
    '      <input id="dti" type="text" placeholder="12.05.2026 00:00 ~ 12.05.2026 23:59" autocomplete="off">',
    '      <div id="dthi">Вставь или введи диапазон — парсится автоматически</div>',
    '      <div id="dtp"></div>',
    '      <button class="tmback" id="dtbb">← Вернуться к автодетекту</button>',
    '    </div>',
    '  </div>',
    '  <button id="tmstart">▶ Запустить сбор</button>',
    '  <div id="tmpeg"><div id="tmpb"></div></div>',
    '  <div id="tmst">Ожидаю загрузки страницы...</div>',
    '</div>',
  ].join('');
  document.body.appendChild(panel);

  var col=false;
  document.getElementById('tmh').addEventListener('click',function(e){
    if(e.target.id==='tmtog'||e.target.closest('#tmtog')){
      col=!col;panel.classList.toggle('col',col);
      document.getElementById('tmtog').textContent=col?'▼':'▲';
    }
  });

  function pad2(v){return('0'+v).slice(-2);}
  function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
  function setStatus(m){var el=document.getElementById('tmst');if(el)el.textContent=m;}
  function setProgress(d,t){
    var p=t>0?Math.round(d/t*100):0;
    var b=document.getElementById('tmpb');if(b)b.style.width=p+'%';
    var sk=state.skipped>0?' | ⚠'+state.skipped:'';
    setStatus('Стр.'+state.currentPage+' | '+d+'/'+t+' ('+p+'%)'+sk);
  }

  function parseRuDate(s){
    s=s.trim();
    var m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if(!m)return null;
    var d=new Date(+m[3],+m[2]-1,+m[1],m[4]?+m[4]:0,m[5]?+m[5]:0,m[6]?+m[6]:0);
    return isNaN(d.getTime())?null:d;
  }
  function parseRange(raw){
    var p=raw.split(/[~—–]/);if(p.length!==2)return null;
    var s=parseRuDate(p[0]),e=parseRuDate(p[1]);
    return(s&&e)?{start:s,end:e}:null;
  }
  function fmtDate(d){
    return pad2(d.getDate())+'.'+pad2(d.getMonth()+1)+'.'+d.getFullYear()+' '+pad2(d.getHours())+':'+pad2(d.getMinutes());
  }
  function msToHMS(ms){
    if(!ms||ms<0)return'00:00:00';
    return pad2(Math.floor(ms/3600000))+':'+pad2(Math.floor(ms/60000)%60)+':'+pad2(Math.floor(ms/1000)%60);
  }
  function sumMs(arr){
    var ts=arr.reduce(function(a,v){return a+Math.floor(v/1000);},0);
    return{ms:ts*1000,hms:pad2(Math.floor(ts/3600))+':'+pad2(Math.floor(ts/60)%60)+':'+pad2(ts%60)};
  }
  function nameMatch(a,b){
    if(!a||!b)return false;
    var wa=a.toLowerCase().split(/\s+/).filter(function(w){return w.length>1;});
    var wb=b.toLowerCase().split(/\s+/).filter(function(w){return w.length>1;});
    var n=0;wa.forEach(function(w){if(wb.indexOf(w)!==-1)n++;});return n>=2;
  }
  function inShift(ds){
    var d=new Date(ds.replace(' ','T'));
    return d>=state.shiftStart&&d<=state.shiftEnd;
  }
  function isSenior(txt){
    if(!txt)return false;
    var t=txt.toLowerCase().replace(/[^a-z0-9\s]/g,'');
    return(t.indexOf('wait')!==-1||t.indexOf('waing')!==-1)&&(t.indexOf('senior')!==-1||t.indexOf('senoir')!==-1);
  }

  // ── Автодетект ──────────────────────────────────────────────
  function detectAgent(){
    var ins=document.querySelectorAll('input');
    for(var i=0;i<ins.length;i++){
      if(ins[i].placeholder==='Admins'){
        var sp=ins[i].parentElement.querySelector('.multiselect__tag span');
        if(sp){var nm=sp.textContent.trim();for(var k=0;k<allLogins.length;k++){if(nameMatch(nm,AGENTS[allLogins[k]]))return allLogins[k];}}
        break;
      }
    }
    return null;
  }
  function detectDate(){
    var el=document.querySelector('input.mx-input');
    return(el&&el.value.trim())?el.value.trim():null;
  }
  function refreshAgent(){
    var lg=detectAgent(),el=document.getElementById('agav');
    if(lg){selectedLogin=lg;el.className='tmad ok';el.textContent='✓  '+AGENTS[lg]+'  ('+lg+')';}
    else{selectedLogin=null;el.className='tmad err';el.textContent='⚠  Агент не найден — выбери вручную';}
  }
  function refreshDate(){
    var raw=detectDate(),el=document.getElementById('dtav');
    if(raw){var r=parseRange(raw);if(r){parsedStart=r.start;parsedEnd=r.end;el.className='tmad ok';el.textContent='✓  '+fmtDate(r.start)+'  →  '+fmtDate(r.end);return;}}
    parsedStart=null;parsedEnd=null;el.className='tmad err';el.textContent='⚠  Период не найден — введи вручную';
  }
  setTimeout(function(){refreshAgent();refreshDate();setStatus('Готов к запуску');},1500);
  setInterval(function(){if(agentMode==='auto')refreshAgent();if(dateMode==='auto')refreshDate();},2000);

  // ── Переключение режимов ────────────────────────────────────
  function setAM(m){
    agentMode=m;
    document.getElementById('agab').classList.toggle('on',m==='auto');
    document.getElementById('agmb').classList.toggle('on',m==='manual');
    document.getElementById('agaw').style.display=m==='auto'?'':'none';
    document.getElementById('agmw').style.display=m==='manual'?'':'none';
    if(m==='auto')refreshAgent();
    else{selectedLogin=null;document.getElementById('agsel').className='h';document.getElementById('ags').value='';setTimeout(function(){document.getElementById('ags').focus();},100);}
  }
  function setDM(m){
    dateMode=m;
    document.getElementById('dtab').classList.toggle('on',m==='auto');
    document.getElementById('dtmb').classList.toggle('on',m==='manual');
    document.getElementById('dtaw').style.display=m==='auto'?'':'none';
    document.getElementById('dtmw').style.display=m==='manual'?'':'none';
    if(m==='auto')refreshDate();
    else{parsedStart=null;parsedEnd=null;setTimeout(function(){document.getElementById('dti').focus();},100);}
  }
  document.getElementById('agab').addEventListener('click',function(){setAM('auto');});
  document.getElementById('agmb').addEventListener('click',function(){setAM('manual');});
  document.getElementById('agbb').addEventListener('click',function(){setAM('auto');});
  document.getElementById('dtab').addEventListener('click',function(){setDM('auto');});
  document.getElementById('dtmb').addEventListener('click',function(){setDM('manual');});
  document.getElementById('dtbb').addEventListener('click',function(){setDM('auto');});

  // ── Дропдаун ────────────────────────────────────────────────
  var ags=document.getElementById('ags'),agd=document.getElementById('agd');
  var agsel=document.getElementById('agsel'),agst=document.getElementById('agst');
  function buildDrop(q){
    q=(q||'').toLowerCase().trim();
    var f=allLogins.filter(function(l){var n=AGENTS[l].toLowerCase();return!q||n.indexOf(q)!==-1||l.indexOf(q)!==-1;});
    agd.innerHTML='';
    if(!f.length){agd.innerHTML='<div class="agdi" style="color:#4a6a8a;cursor:default">Не найдено</div>';}
    else{f.slice(0,40).forEach(function(l){var it=document.createElement('div');it.className='agdi';it.innerHTML=AGENTS[l]+'<span class="lg">('+l+')</span>';it.addEventListener('mousedown',function(e){e.preventDefault();selectedLogin=l;ags.value='';ags.placeholder='Изменить агента...';agd.classList.remove('open');agst.textContent=AGENTS[l]+'  ('+l+')';agsel.className='';});agd.appendChild(it);});}
    agd.classList.add('open');
  }
  ags.addEventListener('focus',function(){buildDrop(this.value);});
  ags.addEventListener('input',function(){buildDrop(this.value);});
  ags.addEventListener('blur',function(){setTimeout(function(){agd.classList.remove('open');},200);});
  document.getElementById('agclr').addEventListener('click',function(){selectedLogin=null;ags.value='';ags.placeholder='Введи имя или pay_mena_...';agsel.className='h';agd.classList.remove('open');});

  // ── Ручная дата ─────────────────────────────────────────────
  var dti=document.getElementById('dti'),dtp=document.getElementById('dtp');
  function tryDate(v){
    v=v.trim();if(!v){dtp.className='';dtp.textContent='';parsedStart=parsedEnd=null;return;}
    var r=parseRange(v);
    if(!r){dtp.className='err';dtp.textContent=v.split(/[~—–]/).length!==2?'⚠ Нужно два значения через ~':'⚠ Не удалось распознать дату';parsedStart=parsedEnd=null;return;}
    parsedStart=r.start;parsedEnd=r.end;dtp.className='ok';dtp.textContent='✓  '+fmtDate(r.start)+'  →  '+fmtDate(r.end);
  }
  dti.addEventListener('input',function(){tryDate(this.value);});
  dti.addEventListener('paste',function(){var s=this;setTimeout(function(){tryDate(s.value);},0);});
  dti.addEventListener('change',function(){tryDate(this.value);});

  // ── Кнопка запуска ──────────────────────────────────────────
  document.getElementById('tmstart').addEventListener('click',function(){
    if(state.running)return;
    if(!selectedLogin){setStatus('⚠ Агент не определён');return;}
    if(!parsedStart||!parsedEnd){setStatus('⚠ Период не определён');return;}
    state.agentLogin=selectedLogin; state.agentName=AGENTS[selectedLogin];
    state.shiftStart=parsedStart; state.shiftEnd=parsedEnd;
    state.results=[]; state.countryStat={}; state.statusStat={};
    state.currentPage=1; state.done=0; state.total=0; state.skipped=0; state.dbg=false;

    console.log('%c[pay_mena v12] СТАРТ','color:#3fa;font-weight:bold');
    console.log('  Агент:  ',state.agentLogin,state.agentName);
    console.log('  Смена:  ',state.shiftStart.toLocaleString('ru'),'→',state.shiftEnd.toLocaleString('ru'));

    this.disabled=true;
    document.getElementById('tmpeg').style.display='block';
    document.getElementById('tmdot').classList.add('busy');
    state.running=true;
    setStatus('Запуск...');
    collectPage();
  });

  // ============================================================
  // FETCH С RETRY — X-Requested-With FIX
  // ============================================================
  function fetchTicket(ticketId, attempt) {
    attempt=attempt||1;

    // Заголовки как у оригинального XHR (ключевой фикс: X-Requested-With)
    var headers = {
      'Content-Type':    'application/x-www-form-urlencoded',
      'X-Requested-With':'XMLHttpRequest',
      'Accept':          'application/json, */*; q=0.01'
    };
    var csrf=document.head.querySelector('meta[name="csrf-token"]');
    if(csrf) headers['X-CSRF-TOKEN']=csrf.getAttribute('content');

    return fetch('/admin/backoffice/paymentsupporthistory',{
      method:'POST',
      headers:headers,
      credentials:'include',
      body:'ticketId='+encodeURIComponent(ticketId)+'&is_iframe=1'
    })
    .then(function(r){
      if(r.status===529||r.status===503||r.status===429){
        if(attempt<=MAX_RETRIES){
          var ws=attempt*2;
          setStatus('⏳ Сервер '+r.status+' — тикет '+ticketId+', попытка '+attempt+'/'+MAX_RETRIES+', жду '+ws+'с...');
          return sleep(ws*1000).then(function(){return fetchTicket(ticketId,attempt+1);});
        }
        console.warn('[pay_mena] ПРОПУЩЕН',ticketId,'— статус',r.status);
        state.skipped++;return null;
      }
      return r.text().then(function(text){
        var t=text.replace(/^\s+/,'');
        if(t.indexOf('<!DOCTYPE')===0||t.indexOf('<html')===0){
          console.error('[pay_mena] HTML-ответ (сессия или редирект). Первые 300 символов:',t.slice(0,300));
          state.running=false;
          document.getElementById('tmdot').classList.remove('busy');
          document.getElementById('tmstart').disabled=false;
          setStatus('🔐 СЕССИЯ ИСТЕКЛА — войди в аккаунт заново. Собрано: '+state.results.length+' записей.');
          throw new Error('SESSION_EXPIRED');
        }
        try{return JSON.parse(text);}
        catch(e){
          console.error('[pay_mena] JSON parse error тикет '+ticketId+':',text.slice(0,150));
          if(attempt<=MAX_RETRIES) return sleep(RETRY_DELAY_MS).then(function(){return fetchTicket(ticketId,attempt+1);});
          state.skipped++;return null;
        }
      });
    })
    .catch(function(e){
      if(e.message==='SESSION_EXPIRED')throw e;
      if(attempt<=MAX_RETRIES){
        setStatus('🔁 Сетевая ошибка — тикет '+ticketId+', retry '+attempt+'/'+MAX_RETRIES);
        return sleep(RETRY_DELAY_MS).then(function(){return fetchTicket(ticketId,attempt+1);});
      }
      state.skipped++;return null;
    });
  }

  // ============================================================
  // ПАРСИНГ ТИКЕТА
  // ============================================================
  function processTkt(info, resp) {
    if(!resp||!resp.success||!resp.data){
      if(!state.dbg){console.warn('[pay_mena] ДИАГНОЗ тикет '+info.ticketId+': пустой/неуспешный ответ',resp);}
      return[];
    }
    var rows=resp.data.map(function(r){return{date:r.dateEdit||'',comment:r.commentSupport||'',status:r.nameExternalStatus||'',admin:r.adminProcessedLogin||''};});
    rows.sort(function(a,b){return new Date(a.date.replace(' ','T'))-new Date(b.date.replace(' ','T'));});

    if(!state.dbg){
      state.dbg=true;
      var admins=[];rows.forEach(function(r){if(admins.indexOf(r.admin)===-1)admins.push(r.admin);});
      console.log('%c[pay_mena] ДИАГНОЗ — тикет '+info.ticketId,'color:#f93;font-weight:bold');
      console.log('  Ищем агента:        "'+state.agentLogin+'"');
      console.log('  Смена:              '+state.shiftStart.toLocaleString('ru')+' → '+state.shiftEnd.toLocaleString('ru'));
      console.log('  Строк в истории:    '+rows.length);
      console.log('  Админы в истории:   ',JSON.stringify(admins));
      console.log('  Даты:               '+(rows[0]?rows[0].date:'?')+' ... '+(rows[rows.length-1]?rows[rows.length-1].date:'?'));
      console.log('  По агенту:          '+rows.filter(function(r){return r.admin===state.agentLogin;}).length);
      console.log('  По смене:           '+rows.filter(function(r){return inShift(r.date);}).length);
      console.log('  По обоим (итог):    '+rows.filter(function(r){return r.admin===state.agentLogin&&inShift(r.date);}).length);
    }

    var ar=rows.filter(function(r){return r.admin===state.agentLogin&&inShift(r.date);});
    if(!ar.length)return[];

    var pairs=[],ip=null;
    ar.forEach(function(cur){
      var isIP=cur.status.toLowerCase().indexOf('in progress')!==-1;
      if(isIP){ip=cur;}
      else{
        pairs.push(ip?{sd:ip.date,ss:ip.status,ed:cur.date,es:cur.status,cm:cur.comment,hp:true}
                     :{sd:'',ss:'',ed:cur.date,es:cur.status,cm:cur.comment,hp:false});
        ip=null;
      }
    });
    if(ip)pairs.push({sd:ip.date,ss:ip.status,ed:'',es:'',cm:ip.comment,hp:true});

    var res=pairs.map(function(p){
      var dur=(p.sd&&p.ed)?msToHMS(new Date(p.ed.replace(' ','T'))-new Date(p.sd.replace(' ','T'))):'';
      var isDup=p.es&&p.es.toLowerCase().indexOf('duplicat')!==-1;
      var isSen=isSenior(p.cm);
      var note=isSen?'Ожидает ответа в чате Q&A Hub':isDup?'Дубликат':!p.hp?'БЕЗ IN PROGRESS':'';
      if(p.es)state.statusStat[p.es]=(state.statusStat[p.es]||0)+1;
      return{ticketId:info.ticketId,country:info.country,login:state.agentLogin,agentName:state.agentName,
             startDate:p.sd,startStatus:p.ss,endDate:p.ed,endStatus:p.es,
             duration:dur,note:note,isDuplicate:isDup,isSenior:isSen,hadInProgress:p.hp};
    });
    if(info.country)state.countryStat[info.country]=(state.countryStat[info.country]||0)+1;
    return res;
  }

  // ============================================================
  // ПОСЛЕДОВАТЕЛЬНАЯ ОБРАБОТКА
  // ============================================================
  function runSeq(tickets,i){
    if(i>=tickets.length)return Promise.resolve();
    var t=tickets[i];
    setStatus('Стр.'+state.currentPage+' | тикет '+(state.done+1)+'/'+state.total+' → '+t.ticketId);
    return fetchTicket(t.ticketId)
      .then(function(resp){
        processTkt(t,resp).forEach(function(r){state.results.push(r);});
        state.done++;setProgress(state.done,state.total);
        return sleep(DELAY_MS);
      })
      .then(function(){return runSeq(tickets,i+1);})
      .catch(function(e){if(e.message==='SESSION_EXPIRED')return;throw e;});
  }

  // ============================================================
  // СТРАНИЦЫ
  // ============================================================
  function collectPage(){
    var rows=Array.from(document.querySelectorAll('table tbody tr')),tickets=[];
    rows.forEach(function(row){
      var tds=row.querySelectorAll('td');if(!tds.length)return;
      var id=tds[0]?tds[0].textContent.trim():'';
      if(!id||!/^\d{5,}$/.test(id))return;
      tickets.push({ticketId:id,country:tds[12]?tds[12].textContent.trim():''});
    });
    state.total+=tickets.length;
    console.log('[pay_mena] Стр.'+state.currentPage+': '+tickets.length+' тикетов, первый: '+(tickets[0]&&tickets[0].ticketId));
    setProgress(state.done,state.total);
    runSeq(tickets,0).then(nextPage);
  }

  function nextPage(){
    var links=Array.from(document.querySelectorAll('a'));
    var num=String(state.currentPage+1),btn=null;
    for(var k=0;k<links.length;k++){if(links[k].textContent.trim()===num){btn=links[k];break;}}
    if(!btn){
      state.running=false;
      document.getElementById('tmdot').classList.remove('busy');
      setStatus('✅ Генерирую файл...');
      console.log('[pay_mena] Готово. Записей:',state.results.length,'Пропущено:',state.skipped);
      genXLSX();return;
    }
    state.currentPage++;btn.click();setTimeout(collectPage,PAGE_WAIT_MS);
  }

  // ============================================================
  // XLSX
  // ============================================================
  function genXLSX(){
    var wb=XLSX.utils.book_new();
    var C={nB:'FF1F4E79',nF:'FFFFFFFF',bB:'FFDAE8F0',bF:'FF1F4E79',gB:'FFC6EFCE',gF:'FF375623',yB:'FFFFEB9C',yF:'FF9C6500',rB:'FFFFC7CE',rF:'FF9C0006',oB:'FFFCE4D6',oF:'FFB85C00',grB:'FFF2F2F2',grF:'FF595959',W:'FFFFFFFF',DT:'FF1F1F1F',alt:'FFF7FBFF'};
    function hdr(){return{font:{bold:true,color:{rgb:C.nF},name:'Arial',sz:10},fill:{fgColor:{rgb:C.nB}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'medium',color:{rgb:C.nB}},right:{style:'thin',color:{rgb:'FFAAAAAA'}}}};}
    function cs(bg,fg,ce){return{font:{color:{rgb:fg||C.DT},name:'Arial',sz:10},fill:{fgColor:{rgb:bg||C.W}},alignment:{horizontal:ce?'center':'left',vertical:'center',wrapText:true},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}
    function sh(){return{font:{bold:true,color:{rgb:C.nF},name:'Arial',sz:11},fill:{fgColor:{rgb:'FF2E75B6'}},alignment:{horizontal:'left',vertical:'center'},border:{bottom:{style:'medium',color:{rgb:'FF1F4E79'}}}};}
    function lb(){return{font:{bold:true,color:{rgb:C.grF},name:'Arial',sz:10},fill:{fgColor:{rgb:C.grB}},alignment:{horizontal:'left',vertical:'center'},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}
    function vl(bg,fg){return{font:{color:{rgb:fg||C.DT},name:'Arial',sz:10},fill:{fgColor:{rgb:bg||C.W}},alignment:{horizontal:'left',vertical:'center'},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}

    // Raw
    var rh=['Ticket ID','Country','Login','Agent Name','Start Date','Start Status','End Date','End Status','Duration','Note'];
    var rd=[rh].concat(state.results.map(function(r){return[r.ticketId,r.country,r.login,r.agentName,r.startDate,r.startStatus,r.endDate,r.endStatus,r.duration,r.note];}));
    var wR=XLSX.utils.aoa_to_sheet(rd);
    wR['!cols']=[{wch:14},{wch:12},{wch:18},{wch:26},{wch:20},{wch:22},{wch:20},{wch:24},{wch:12},{wch:30}];
    var rr=XLSX.utils.decode_range(wR['!ref']);
    for(var c=rr.s.c;c<=rr.e.c;c++){var ha=XLSX.utils.encode_cell({r:0,c:c});if(wR[ha])wR[ha].s=hdr();}
    for(var r=1;r<=rr.e.r;r++){
      var na=XLSX.utils.encode_cell({r:r,c:9}),nv=wR[na]?wR[na].v:'',bg=r%2===0?C.alt:C.W;
      for(var ci=rr.s.c;ci<=rr.e.c;ci++){var a=XLSX.utils.encode_cell({r:r,c:ci});if(!wR[a])wR[a]={t:'s',v:''};wR[a].s=nv==='Ожидает ответа в чате Q&A Hub'?cs(C.oB,C.oF):nv==='Дубликат'?cs(C.bB,C.bF):nv==='БЕЗ IN PROGRESS'?cs(C.yB,C.yF):cs(bg);}
    }
    wR['!freeze']={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,wR,'Raw');

    // Processed
    var sorted=state.results.filter(function(r){return r.startDate&&r.endDate&&r.hadInProgress;}).sort(function(a,b){return new Date(a.startDate.replace(' ','T'))-new Date(b.startDate.replace(' ','T'));});
    var itvs=[],prts=[],sCnt={},cCnt={},pEnd=null;
    var pd=[['Ticket ID','Country','Start Time','End Time','Processing Time','Interval Since Previous','End Status','Note']];
    sorted.forEach(function(row,pi){
      var s=new Date(row.startDate.replace(' ','T')),e=new Date(row.endDate.replace(' ','T'));
      var pm=e-s,im=(pEnd&&pi>0)?Math.max(0,s-pEnd):0;pEnd=e;
      prts.push({ticket:row.ticketId,ms:pm});if(pi>0)itvs.push({ticket:row.ticketId,ms:im});
      if(row.endStatus)sCnt[row.endStatus]=(sCnt[row.endStatus]||0)+1;
      if(row.country)cCnt[row.country]=(cCnt[row.country]||0)+1;
      pd.push([row.ticketId,row.country,row.startDate,row.endDate,msToHMS(pm),pi>0?msToHMS(im):'—',row.endStatus,row.note]);
    });
    var wP=XLSX.utils.aoa_to_sheet(pd);
    wP['!cols']=[{wch:14},{wch:12},{wch:20},{wch:20},{wch:16},{wch:22},{wch:24},{wch:30}];
    var pr=XLSX.utils.decode_range(wP['!ref']);
    for(var c2=pr.s.c;c2<=pr.e.c;c2++){var hp=XLSX.utils.encode_cell({r:0,c:c2});if(wP[hp])wP[hp].s=hdr();}
    for(var r2=1;r2<=pr.e.r;r2++){
      var ic=XLSX.utils.encode_cell({r:r2,c:5}),iv=wP[ic]?wP[ic].v:'—',im2=0;
      if(iv&&iv!=='—'){var pt=iv.split(':');im2=(+pt[0])*60+(+pt[1]);}
      var ibg=iv==='—'||im2===0?C.W:im2<=INTERVAL_GREEN?C.gB:im2<=INTERVAL_YELLOW?C.yB:C.rB;
      var ifg=iv==='—'||im2===0?C.DT:im2<=INTERVAL_GREEN?C.gF:im2<=INTERVAL_YELLOW?C.yF:C.rF;
      for(var c3=pr.s.c;c3<=pr.e.c;c3++){var pa=XLSX.utils.encode_cell({r:r2,c:c3});if(!wP[pa])wP[pa]={t:'s',v:''};wP[pa].s=cs(ibg,ifg);}
    }
    wP['!freeze']={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,wP,'Processed');

    // Dashboard
    var si=sumMs(itvs.map(function(i){return i.ms;}));
    var ai=itvs.length>0?si.ms/itvs.length:0;
    var mp=prts.reduce(function(m,i){return i.ms>m.ms?i:m;},{ms:0,ticket:'-'});
    var ap=prts.length>0?sumMs(prts.map(function(i){return i.ms;})).ms/prts.length:0;
    var tsk='',tsv=0;Object.keys(sCnt).forEach(function(k){if(sCnt[k]>tsv){tsv=sCnt[k];tsk=k;}});
    var noIP=state.results.filter(function(r){return!r.hadInProgress&&!r.isDuplicate;}).length;
    var dup=state.results.filter(function(r){return r.isDuplicate;}).length;
    var sen=state.results.filter(function(r){return r.isSenior;}).length;
    var inp=state.results.filter(function(r){return r.startDate&&!r.endDate;}).sort(function(a,b){return new Date(a.startDate.replace(' ','T'))-new Date(b.startDate.replace(' ','T'));});
    var bex=si.ms>=BREAK_LIMIT_MS;
    var sArr=Object.keys(state.statusStat).map(function(k){return{k:k,v:state.statusStat[k]};}).sort(function(a,b){return b.v-a.v;});
    var cArr=Object.keys(state.countryStat).map(function(k){return{k:k,v:state.countryStat[k]};}).sort(function(a,b){return b.v-a.v;});
    var topI=itvs.filter(function(i){return i.ms>=MIN_BREAK_SHOW;}).sort(function(a,b){return b.ms-a.ms;});

    var ws={},ROW=0;
    function w(col,v,sty,tp){var a=XLSX.utils.encode_cell({r:ROW,c:col});ws[a]={t:tp||'s',v:v==null?'':v};if(sty)ws[a].s=sty;}
    function nr(){ROW++;} function br(){ROW++;}

    w(0,'ОБЩАЯ СТАТИСТИКА',sh());w(1,'',sh());w(2,'',sh());w(3,'',sh());nr();
    [['Агент',state.agentName+' ('+state.agentLogin+')',null],
     ['Период',state.shiftStart.toLocaleString('ru')+' — '+state.shiftEnd.toLocaleString('ru'),null],
     ['Всего записей',state.results.length,'n'],['Тикетов с полными датами',sorted.length,'n'],
     ['Самый частый статус',tsk+' ('+tsv+')',null],['Дубликатов',dup,'n'],
     ['Без In Progress',noIP,'n'],['Ожидали Q&A Hub',sen,'n'],
     ['Пропущено (ошибки)',state.skipped,'n'],
     ['Среднее время обработки',msToHMS(ap),null],['Макс. время обработки',msToHMS(mp.ms)+' (тикет '+mp.ticket+')',null],
     ['Средний интервал',msToHMS(ai),null],
     ['Суммарные перерывы',si.hms,bex?[C.rB,C.rF]:[C.gB,C.gF]],
     ['Лимит перерывов',bex?'ПРЕВЫШЕН':'В норме',bex?[C.rB,C.rF]:[C.gB,C.gF]],
    ].forEach(function(s){w(0,s[0],lb());w(1,s[1],vl(s[2]?s[2][0]:C.W,s[2]?s[2][1]:C.DT),s[3]||null);w(2,'',cs(C.W));w(3,'',cs(C.W));nr();});
    br();
    w(0,'СТАТУСЫ',sh());w(1,'',sh());w(2,'',sh());w(3,'',sh());nr();
    w(0,'Статус',lb());w(1,'Кол-во',lb());w(2,'',cs(C.W));w(3,'',cs(C.W));nr();
    sArr.forEach(function(s,i){var rb=i%2===0?C.W:C.alt;w(0,s.k,cs(rb));w(1,s.v,cs(rb,C.DT,true),'n');w(2,'',cs(C.W));w(3,'',cs(C.W));nr();});
    w(0,'ИТОГО',lb());w(1,state.results.length,vl(C.bB,C.bF),'n');nr();br();
    w(0,'ТОП СТРАН',sh());w(1,'',sh());w(2,'',sh());w(3,'',sh());nr();
    w(0,'Страна',lb());w(1,'Тикетов',lb());w(2,'',cs(C.W));w(3,'',cs(C.W));nr();
    cArr.slice(0,10).forEach(function(c,i){var cb=i%2===0?C.W:C.alt;w(0,c.k,cs(cb));w(1,c.v,cs(cb,C.DT,true),'n');w(2,'',cs(C.W));w(3,'',cs(C.W));nr();});
    br();
    w(0,'В РАБОТЕ ПРЯМО СЕЙЧАС',sh());w(1,'',sh());w(2,'',sh());w(3,'',sh());nr();
    w(0,'Ticket ID',lb());w(1,'Country',lb());w(2,'Взят в работу',lb());w(3,'Висит уже',lb());nr();
    var now=Date.now();
    if(!inp.length){w(0,'Нет активных тикетов',cs(C.gB,C.gF));w(1,'',cs(C.gB));w(2,'',cs(C.gB));w(3,'',cs(C.gB));nr();}
    else{inp.forEach(function(r){var hg=now-new Date(r.startDate.replace(' ','T')).getTime();var hm=Math.floor(hg/60000);var ib=hm>60?C.rB:hm>30?C.yB:C.gB;var if2=hm>60?C.rF:hm>30?C.yF:C.gF;w(0,r.ticketId,cs(ib,if2));w(1,r.country,cs(ib,if2));w(2,r.startDate,cs(ib,if2));w(3,msToHMS(hg),cs(ib,if2,true));nr();});}

    var sv=ROW;ROW=0;
    w(5,'ПЕРЕРЫВЫ (>= 5 мин)',sh());w(6,'',sh());w(7,'',sh());w(8,'',sh());nr();
    w(5,'#',lb());w(6,'Ticket ID',lb());w(7,'Интервал',lb());w(8,'Оценка',lb());nr();
    if(!topI.length){w(5,'Нет перерывов >= 5 мин',cs(C.grB));w(6,'',cs(C.grB));w(7,'',cs(C.grB));w(8,'',cs(C.grB));nr();}
    else{topI.forEach(function(x,ti){var xm=Math.floor(x.ms/60000);var tb=xm<=INTERVAL_GREEN?C.gB:xm<=INTERVAL_YELLOW?C.yB:C.rB;var tf=xm<=INTERVAL_GREEN?C.gF:xm<=INTERVAL_YELLOW?C.yF:C.rF;var ax=xm<=INTERVAL_GREEN?'Норма':xm<=INTERVAL_YELLOW?'Внимание':'Нарушение';w(5,ti+1,cs(tb,tf,true),'n');w(6,x.ticket,cs(tb,tf));w(7,msToHMS(x.ms),cs(tb,tf,true));w(8,ax,cs(tb,tf,true));nr();});}
    w(5,'Суммарно (все)',lb());w(6,'',cs(C.W));
    w(7,si.hms,vl(bex?C.rB:C.gB,bex?C.rF:C.gF));
    w(8,bex?'ПРЕВЫШЕН':'В норме',vl(bex?C.rB:C.gB,bex?C.rF:C.gF));
    ROW=Math.max(ROW,sv);

    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:ROW+5,c:8}});
    ws['!cols']=[{wch:28},{wch:32},{wch:16},{wch:22},{wch:3},{wch:6},{wch:16},{wch:14},{wch:14}];
    var ra=[];for(var ri=0;ri<=ROW+5;ri++)ra.push({hpt:20});
    ws['!rows']=ra;
    XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

    var fn='report_'+state.agentLogin+'_'+new Date().toISOString().slice(0,10)+'.xlsx';
    XLSX.writeFile(wb,fn);
    document.getElementById('tmstart').disabled=false;
    setProgress(state.done,state.done);
    setStatus('✅ '+fn+' | Записей: '+state.results.length+(state.skipped?' | ⚠ пропущено: '+state.skipped:''));
  }

})();
