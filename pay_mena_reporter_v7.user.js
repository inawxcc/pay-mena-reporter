// ==UserScript==
// @name         pay_mena Ticket Reporter
// @namespace    http://tampermonkey.net/
// @version      7.0
// @description  Сбор и выгрузка тикетов агентов pay_mena в XLSX
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

  // ============================================================
  // КОНФИГ
  // ============================================================
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

  var INTERVAL_GREEN       = 50;
  var INTERVAL_YELLOW      = 60;
  var BREAK_LIMIT_MS       = 60 * 60 * 1000;
  var MIN_INTERVAL_SHOW_MS = 5 * 60 * 1000;
  var MODAL_WAIT_MS        = 3000;
  var AFTER_CLOSE_MS       = 300;
  var PAGE_WAIT_MS         = 2000;

  // ============================================================
  // СОСТОЯНИЕ
  // ============================================================
  var state = {
    agentLogin:  null,
    agentName:   null,
    shiftStart:  null,
    shiftEnd:    null,
    results:     [],
    countryStat: {},
    statusStat:  {},
    currentPage: 1,
    running:     false,
    total:       0,
    done:        0
  };

  // ============================================================
  // СТИЛИ
  // ============================================================
  GM_addStyle([
    /* ── Панель ── */
    '#tm-panel{',
    '  position:fixed;bottom:20px;right:20px;z-index:2147483647;',
    '  width:380px;',
    '  background:#0f1a2b;',
    '  border:1px solid #1e3a5f;',
    '  border-radius:14px;',
    '  box-shadow:0 8px 32px rgba(0,0,0,.6);',
    '  font-family:Arial,sans-serif;',
    '  font-size:13px;',
    '  color:#d0e8ff;',
    '  overflow:hidden;',
    '  transition:all .25s ease;',
    '}',

    /* свёрнутый режим */
    '#tm-panel.collapsed #tm-body{display:none;}',
    '#tm-panel.collapsed{width:200px;}',

    /* ── Шапка ── */
    '#tm-header{',
    '  display:flex;align-items:center;justify-content:space-between;',
    '  padding:11px 14px 10px;',
    '  background:linear-gradient(90deg,#1a3a5c,#0f2540);',
    '  border-bottom:1px solid #1e3a5f;',
    '  cursor:pointer;user-select:none;',
    '}',
    '#tm-header-title{display:flex;align-items:center;gap:7px;font-weight:700;font-size:13px;color:#7ecfff;letter-spacing:.4px;}',
    '#tm-header-title span.dot{width:8px;height:8px;border-radius:50%;background:#3fa;display:inline-block;flex-shrink:0;}',
    '#tm-header-title span.dot.busy{background:#f93;}',
    '#tm-toggle-btn{background:none;border:none;color:#5a8ab0;font-size:17px;cursor:pointer;line-height:1;padding:0;flex-shrink:0;}',
    '#tm-toggle-btn:hover{color:#7ecfff;}',

    /* ── Тело ── */
    '#tm-body{padding:14px 14px 12px;}',

    /* ── Секции ── */
    '.tm-section{margin-bottom:10px;}',
    '.tm-label{font-size:11px;color:#5a8ab0;font-weight:600;letter-spacing:.5px;text-transform:uppercase;margin-bottom:5px;}',

    /* ── Агент ── */
    '#tm-agent-wrap{position:relative;}',
    '#tm-agent-search{',
    '  width:100%;box-sizing:border-box;',
    '  padding:8px 10px;',
    '  border-radius:8px;border:1px solid #1e3a5f;',
    '  background:#071426;color:#d0e8ff;font-size:12px;',
    '  outline:none;transition:border-color .2s;',
    '}',
    '#tm-agent-search:focus{border-color:#2e75b6;}',
    '#tm-agent-search::placeholder{color:#3a5a7a;}',
    '#tm-agent-dropdown{',
    '  position:absolute;top:calc(100% + 4px);left:0;right:0;',
    '  background:#0d1f33;border:1px solid #1e3a5f;border-radius:8px;',
    '  max-height:180px;overflow-y:auto;z-index:999;',
    '  box-shadow:0 4px 16px rgba(0,0,0,.5);',
    '  display:none;',
    '}',
    '#tm-agent-dropdown.open{display:block;}',
    '.tm-drop-item{',
    '  padding:7px 10px;cursor:pointer;font-size:12px;',
    '  border-bottom:1px solid #132030;color:#b0d4f0;',
    '  transition:background .15s;',
    '}',
    '.tm-drop-item:last-child{border-bottom:none;}',
    '.tm-drop-item:hover,.tm-drop-item.active{background:#1a3a5c;color:#7ecfff;}',
    '.tm-drop-item span.login{color:#4a7a9a;font-size:11px;margin-left:5px;}',
    '#tm-agent-selected{',
    '  margin-top:5px;padding:5px 10px;border-radius:6px;',
    '  background:#0a2040;border:1px solid #2e75b6;',
    '  font-size:12px;color:#7ecfff;display:none;',
    '  display:flex;align-items:center;justify-content:space-between;',
    '}',
    '#tm-agent-selected.hidden{display:none;}',
    '#tm-agent-clear{background:none;border:none;color:#5a8ab0;cursor:pointer;font-size:13px;padding:0;}',
    '#tm-agent-clear:hover{color:#f55;}',

    /* ── Дата ── */
    '#tm-date-input{',
    '  width:100%;box-sizing:border-box;',
    '  padding:8px 10px;',
    '  border-radius:8px;border:1px solid #1e3a5f;',
    '  background:#071426;color:#d0e8ff;font-size:12px;',
    '  outline:none;transition:border-color .2s;',
    '}',
    '#tm-date-input:focus{border-color:#2e75b6;}',
    '#tm-date-input::placeholder{color:#3a5a7a;}',
    '#tm-date-hint{font-size:10px;color:#3a6a8a;margin-top:4px;}',
    '#tm-date-parsed{',
    '  margin-top:5px;padding:5px 10px;border-radius:6px;',
    '  background:#0a2040;border:1px solid #1e3a5f;',
    '  font-size:11px;color:#5ab0d0;display:none;',
    '}',
    '#tm-date-parsed.ok{border-color:#2e75b6;color:#7ecfff;display:block;}',
    '#tm-date-parsed.err{border-color:#8b2020;color:#f06060;display:block;}',

    /* ── Кнопка ── */
    '#tm-btn-start{',
    '  width:100%;padding:10px;border:none;border-radius:8px;',
    '  background:linear-gradient(135deg,#2e75b6,#1a5276);',
    '  color:#fff;font-size:13px;font-weight:700;cursor:pointer;',
    '  letter-spacing:.4px;transition:opacity .2s,transform .1s;',
    '  margin-top:4px;',
    '}',
    '#tm-btn-start:hover:not(:disabled){opacity:.9;transform:translateY(-1px);}',
    '#tm-btn-start:active:not(:disabled){transform:translateY(0);}',
    '#tm-btn-start:disabled{opacity:.4;cursor:not-allowed;}',

    /* ── Прогресс ── */
    '#tm-progress-wrap{',
    '  margin-top:10px;background:#071426;border-radius:6px;',
    '  overflow:hidden;height:6px;display:none;',
    '}',
    '#tm-progress-bar{',
    '  height:100%;width:0%;',
    '  background:linear-gradient(90deg,#2e75b6,#7ecfff);',
    '  transition:width .3s ease;',
    '}',

    /* ── Статус ── */
    '#tm-status{',
    '  margin-top:8px;font-size:11px;color:#4a8ab0;',
    '  min-height:15px;line-height:1.4;word-break:break-word;',
    '}',

    /* скроллбар дропдауна */
    '#tm-agent-dropdown::-webkit-scrollbar{width:4px;}',
    '#tm-agent-dropdown::-webkit-scrollbar-track{background:#0d1f33;}',
    '#tm-agent-dropdown::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:2px;}',
  ].join(''));

  // ============================================================
  // HTML
  // ============================================================
  var panel = document.createElement('div');
  panel.id = 'tm-panel';
  panel.innerHTML = [
    '<div id="tm-header">',
    '  <div id="tm-header-title">',
    '    <span class="dot" id="tm-dot"></span>',
    '    📊 pay_mena Reporter',
    '  </div>',
    '  <button id="tm-toggle-btn" title="Свернуть/развернуть">▲</button>',
    '</div>',

    '<div id="tm-body">',

    '  <div class="tm-section">',
    '    <div class="tm-label">Агент</div>',
    '    <div id="tm-agent-wrap">',
    '      <input id="tm-agent-search" type="text" placeholder="Введи имя или pay_mena_..." autocomplete="off">',
    '      <div id="tm-agent-dropdown"></div>',
    '    </div>',
    '    <div id="tm-agent-selected" class="hidden">',
    '      <span id="tm-agent-selected-text"></span>',
    '      <button id="tm-agent-clear" title="Сбросить">✕</button>',
    '    </div>',
    '  </div>',

    '  <div class="tm-section">',
    '    <div class="tm-label">Период</div>',
    '    <input id="tm-date-input" type="text"',
    '      placeholder="12.05.2026 00:00 ~ 12.05.2026 23:59"',
    '      autocomplete="off">',
    '    <div id="tm-date-hint">Вставь или введи диапазон — парсится автоматически</div>',
    '    <div id="tm-date-parsed"></div>',
    '  </div>',

    '  <button id="tm-btn-start">▶ Запустить сбор</button>',

    '  <div id="tm-progress-wrap"><div id="tm-progress-bar"></div></div>',
    '  <div id="tm-status">Выбери агента и период</div>',

    '</div>',
  ].join('');
  document.body.appendChild(panel);

  // ============================================================
  // ЛОГИКА ПАНЕЛИ
  // ============================================================

  // Свернуть/развернуть
  var collapsed = false;
  document.getElementById('tm-header').addEventListener('click', function(e) {
    if (e.target.id === 'tm-toggle-btn' || e.target.closest('#tm-toggle-btn')) {
      collapsed = !collapsed;
      panel.classList.toggle('collapsed', collapsed);
      document.getElementById('tm-toggle-btn').textContent = collapsed ? '▼' : '▲';
    }
  });

  // ── Список агентов ──────────────────────────────────────────
  var selectedLogin = null;
  var allLogins = Object.keys(AGENTS).sort(function(a,b) {
    return AGENTS[a].localeCompare(AGENTS[b], 'ru');
  });

  var searchEl   = document.getElementById('tm-agent-search');
  var dropEl     = document.getElementById('tm-agent-dropdown');
  var selectedEl = document.getElementById('tm-agent-selected');
  var selectedTx = document.getElementById('tm-agent-selected-text');

  function buildDropdown(query) {
    var q = (query || '').toLowerCase().trim();
    var filtered = allLogins.filter(function(login) {
      var name = AGENTS[login].toLowerCase();
      return !q || name.indexOf(q) !== -1 || login.indexOf(q) !== -1;
    });
    dropEl.innerHTML = '';
    if (filtered.length === 0) {
      dropEl.innerHTML = '<div class="tm-drop-item" style="color:#4a6a8a;cursor:default">Не найдено</div>';
    } else {
      filtered.slice(0, 40).forEach(function(login) {
        var item = document.createElement('div');
        item.className = 'tm-drop-item';
        item.innerHTML = AGENTS[login] + '<span class="login">(' + login + ')</span>';
        item.addEventListener('mousedown', function(e) {
          e.preventDefault();
          selectAgent(login);
        });
        dropEl.appendChild(item);
      });
    }
    dropEl.classList.add('open');
  }

  function selectAgent(login) {
    selectedLogin = login;
    searchEl.value = '';
    searchEl.placeholder = 'Изменить агента...';
    dropEl.classList.remove('open');
    selectedTx.textContent = AGENTS[login] + '  (' + login + ')';
    selectedEl.classList.remove('hidden');
  }

  function clearAgent() {
    selectedLogin = null;
    searchEl.value = '';
    searchEl.placeholder = 'Введи имя или pay_mena_...';
    selectedEl.classList.add('hidden');
    dropEl.classList.remove('open');
  }

  searchEl.addEventListener('focus', function() { buildDropdown(this.value); });
  searchEl.addEventListener('input', function() { buildDropdown(this.value); });
  searchEl.addEventListener('blur',  function() { setTimeout(function(){ dropEl.classList.remove('open'); }, 200); });
  document.getElementById('tm-agent-clear').addEventListener('click', clearAgent);

  // ── Парсинг даты ────────────────────────────────────────────
  var parsedStart = null;
  var parsedEnd   = null;

  // Формат: DD.MM.YYYY HH:MM  (время опционально)
  function parseRuDate(str) {
    str = str.trim();
    // DD.MM.YYYY HH:MM или DD.MM.YYYY HH:MM:SS
    var m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (!m) return null;
    var d = new Date(
      +m[3], +m[2]-1, +m[1],
      m[4] ? +m[4] : 0,
      m[5] ? +m[5] : 0,
      m[6] ? +m[6] : 0
    );
    return isNaN(d.getTime()) ? null : d;
  }

  function fmtDate(d) {
    return pad2(d.getDate())+'.'+pad2(d.getMonth()+1)+'.'+d.getFullYear()+
           ' '+pad2(d.getHours())+':'+pad2(d.getMinutes());
  }
  function pad2(v){ return ('0'+v).slice(-2); }

  var dateInputEl  = document.getElementById('tm-date-input');
  var dateParsedEl = document.getElementById('tm-date-parsed');

  function tryParseDate(val) {
    var raw = val.trim();
    if (!raw) {
      dateParsedEl.className = '';
      dateParsedEl.textContent = '';
      parsedStart = parsedEnd = null;
      return;
    }
    // Разделитель ~ или —
    var parts = raw.split(/[~—–]/);
    if (parts.length !== 2) {
      dateParsedEl.className = 'err';
      dateParsedEl.textContent = '⚠ Нужно два значения через ~';
      parsedStart = parsedEnd = null;
      return;
    }
    var s = parseRuDate(parts[0]);
    var e = parseRuDate(parts[1]);
    if (!s || !e) {
      dateParsedEl.className = 'err';
      dateParsedEl.textContent = '⚠ Не удалось распознать дату';
      parsedStart = parsedEnd = null;
      return;
    }
    parsedStart = s;
    parsedEnd   = e;
    dateParsedEl.className = 'ok';
    dateParsedEl.textContent = '✓  ' + fmtDate(s) + '  →  ' + fmtDate(e);
  }

  dateInputEl.addEventListener('input',  function(){ tryParseDate(this.value); });
  dateInputEl.addEventListener('paste',  function(){
    var self = this;
    setTimeout(function(){ tryParseDate(self.value); }, 0);
  });
  dateInputEl.addEventListener('change', function(){ tryParseDate(this.value); });

  // ── Автозаполнение из фильтров страницы ─────────────────────
  setTimeout(autoFillFromPage, 1400);

  function autoFillFromPage() {
    // Агент
    var inputs = document.querySelectorAll('input');
    var adminsInput = null;
    for (var i = 0; i < inputs.length; i++) {
      if (inputs[i].placeholder === 'Admins') { adminsInput = inputs[i]; break; }
    }
    var tagSpan = adminsInput ? adminsInput.parentElement.querySelector('.multiselect__tag span') : null;
    if (tagSpan) {
      var name = tagSpan.textContent.trim();
      for (var k = 0; k < allLogins.length; k++) {
        if (nameMatch(name, AGENTS[allLogins[k]])) { selectAgent(allLogins[k]); break; }
      }
    }

    // Дата
    var dateInput = document.querySelector('input.mx-input');
    if (dateInput && dateInput.value.trim()) {
      dateInputEl.value = dateInput.value.trim();
      tryParseDate(dateInputEl.value);
    }

    if (selectedLogin || parsedStart) {
      setStatus('Автозаполнено из фильтров. Проверь и запускай.');
    }
  }

  // ── Кнопка запуска ──────────────────────────────────────────
  document.getElementById('tm-btn-start').addEventListener('click', function() {
    if (state.running) return;

    if (!selectedLogin)  { setStatus('⚠ Выбери агента'); return; }
    if (!parsedStart || !parsedEnd) { setStatus('⚠ Укажи корректный период'); return; }

    state.agentLogin  = selectedLogin;
    state.agentName   = AGENTS[selectedLogin];
    state.shiftStart  = parsedStart;
    state.shiftEnd    = parsedEnd;
    state.results     = [];
    state.countryStat = {};
    state.statusStat  = {};
    state.currentPage = 1;
    state.done        = 0;
    state.total       = 0;

    this.disabled = true;
    document.getElementById('tm-progress-wrap').style.display = 'block';
    document.getElementById('tm-dot').classList.add('busy');
    state.running = true;
    setStatus('Запуск...');
    startCollection();
  });

  // ============================================================
  // УТИЛИТЫ
  // ============================================================
  function setStatus(msg) {
    var el = document.getElementById('tm-status');
    if (el) el.textContent = msg;
  }
  function setProgress(done, total) {
    var pct = total > 0 ? Math.round(done / total * 100) : 0;
    var bar = document.getElementById('tm-progress-bar');
    if (bar) bar.style.width = pct + '%';
    setStatus('Стр. ' + state.currentPage + ' | ' + done + ' / ' + total + ' тикетов  (' + pct + '%)');
  }
  function sleep(ms) { return new Promise(function(r){ setTimeout(r, ms); }); }

  function msToHMS(ms) {
    if (!ms || ms < 0) return '00:00:00';
    var s = Math.floor(ms/1000)%60, m = Math.floor(ms/60000)%60, h = Math.floor(ms/3600000);
    return pad2(h)+':'+pad2(m)+':'+pad2(s);
  }
  function sumMsArray(arr) {
    var ts = arr.reduce(function(a,v){ return a+Math.floor(v/1000); }, 0);
    return { ms: ts*1000, hms: pad2(Math.floor(ts/3600))+':'+pad2(Math.floor(ts/60)%60)+':'+pad2(ts%60) };
  }
  function nameMatch(a, b) {
    if (!a||!b) return false;
    var wa = a.toLowerCase().split(/\s+/).filter(function(w){ return w.length>1; });
    var wb = b.toLowerCase().split(/\s+/).filter(function(w){ return w.length>1; });
    var n=0; wa.forEach(function(w){ if(wb.indexOf(w)!==-1) n++; });
    return n>=2;
  }
  function inShift(dateStr) {
    var d = new Date(dateStr.replace(' ','T'));
    return d >= state.shiftStart && d <= state.shiftEnd;
  }
  function isWaitingForSenior(text) {
    if (!text) return false;
    var t = text.toLowerCase().replace(/[^a-z0-9\s]/g,'');
    return (t.indexOf('wait')!==-1||t.indexOf('waing')!==-1||t.indexOf('wating')!==-1) &&
           (t.indexOf('senior')!==-1||t.indexOf('senoir')!==-1||t.indexOf('senor')!==-1);
  }
  function getCountryFromRow(btn) {
    var row = btn.closest('tr'); if (!row) return '';
    var cells = row.querySelectorAll('td');
    return cells[12] ? cells[12].textContent.trim() : '';
  }
  function getTicketId(btn) {
    var row = btn.closest('tr'); if (!row) return '?';
    var td = row.querySelector('td'); return td ? td.textContent.trim() : '?';
  }
  function closeModal() {
    document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',keyCode:27,bubbles:true}));
  }
  function waitForModal(ticketId, timeout) {
    return new Promise(function(resolve) {
      var elapsed=0, t=setInterval(function(){
        var modal=document.querySelector('.modal_content');
        var rows=modal?modal.querySelectorAll('table tr'):[];
        if (modal&&rows.length>=2){clearInterval(t);resolve(modal);return;}
        elapsed+=100;
        if (elapsed>=timeout){clearInterval(t);resolve(null);}
      },100);
    });
  }

  // ============================================================
  // ПАРСИНГ ТИКЕТА
  // ============================================================
  function processSingleTicket(btn) {
    var ticketId = getTicketId(btn);
    var country  = getCountryFromRow(btn);
    btn.click();
    return waitForModal(ticketId, MODAL_WAIT_MS).then(function(modal) {
      if (!modal) { closeModal(); return []; }

      var allRows=[];
      modal.querySelectorAll('table tr').forEach(function(row,idx){
        if (idx===0) return;
        var cells=row.querySelectorAll('td'); if (cells.length<10) return;
        var date=cells[0]?cells[0].textContent.trim():'';
        var comment=cells[4]?cells[4].textContent.trim():'';
        var status=cells[5]?cells[5].textContent.trim():'';
        var admin=cells[9]?cells[9].textContent.trim():'';
        if (date) allRows.push({date:date,insideComment:comment,externalStatus:status,adminUsername:admin});
      });
      allRows.sort(function(a,b){ return new Date(a.date.replace(' ','T'))-new Date(b.date.replace(' ','T')); });

      var agentRows=allRows.filter(function(r){ return r.adminUsername===state.agentLogin&&inShift(r.date); });
      closeModal();
      if (agentRows.length===0) return [];

      var pairs=[],ipRow=null;
      agentRows.forEach(function(cur){
        var isIP=cur.externalStatus.toLowerCase().indexOf('in progress')!==-1;
        if (isIP) { ipRow=cur; }
        else {
          pairs.push(ipRow
            ? {startDate:ipRow.date,startStatus:ipRow.externalStatus,endDate:cur.date,endStatus:cur.externalStatus,insideComment:cur.insideComment,hadInProgress:true}
            : {startDate:'',startStatus:'',endDate:cur.date,endStatus:cur.externalStatus,insideComment:cur.insideComment,hadInProgress:false}
          );
          ipRow=null;
        }
      });
      if (ipRow) pairs.push({startDate:ipRow.date,startStatus:ipRow.externalStatus,endDate:'',endStatus:'',insideComment:ipRow.insideComment,hadInProgress:true});

      var ticketResults=pairs.map(function(pair){
        var duration=(pair.startDate&&pair.endDate)?msToHMS(new Date(pair.endDate.replace(' ','T'))-new Date(pair.startDate.replace(' ','T'))):'';
        var isDuplicate=pair.endStatus&&pair.endStatus.toLowerCase().indexOf('duplicat')!==-1;
        var isSenior=isWaitingForSenior(pair.insideComment);
        var note=isSenior?'Ожидает ответа в чате Q&A Hub':isDuplicate?'Дубликат':!pair.hadInProgress?'БЕЗ IN PROGRESS':'';
        if (pair.endStatus) state.statusStat[pair.endStatus]=(state.statusStat[pair.endStatus]||0)+1;
        return {ticketId:ticketId,country:country,login:state.agentLogin,agentName:state.agentName,
                startDate:pair.startDate,startStatus:pair.startStatus,endDate:pair.endDate,endStatus:pair.endStatus,
                duration:duration,note:note,isDuplicate:isDuplicate,isSenior:isSenior,hadInProgress:pair.hadInProgress};
      });
      if (country) state.countryStat[country]=(state.countryStat[country]||0)+1;
      return ticketResults;
    });
  }

  // ============================================================
  // ЦИКЛ ОБХОДА
  // ============================================================
  function processPageBatch(buttons, offset) {
    if (offset>=buttons.length) return Promise.resolve();
    var btn=buttons[offset];
    state.done++; setProgress(state.done,state.total);
    return processSingleTicket(btn).then(function(data){
      data.forEach(function(r){ state.results.push(r); });
      return sleep(AFTER_CLOSE_MS);
    }).then(function(){ return processPageBatch(buttons,offset+1); });
  }

  function startCollection() { collectPage(); }

  function collectPage() {
    var buttons=Array.from(document.querySelectorAll('a')).filter(function(a){ return a.textContent.trim()==='Show'; });
    state.total+=buttons.length; setProgress(state.done,state.total);
    processPageBatch(buttons,0).then(goNextPage);
  }

  function goNextPage() {
    var links=Array.from(document.querySelectorAll('a'));
    var nextBtn=null, nextNum=String(state.currentPage+1);
    for (var k=0;k<links.length;k++){
      if (links[k].textContent.trim()===nextNum){nextBtn=links[k];break;}
    }
    if (!nextBtn) {
      state.running=false;
      document.getElementById('tm-dot').classList.remove('busy');
      setStatus('✅ Готово! Генерирую файл...');
      generateXLSX(); return;
    }
    state.currentPage++; nextBtn.click();
    setTimeout(collectPage, PAGE_WAIT_MS);
  }

  // ============================================================
  // ГЕНЕРАЦИЯ XLSX (таблицы без изменений)
  // ============================================================
  function generateXLSX() {
    var wb=XLSX.utils.book_new();
    var C={
      navyBg:'FF1F4E79',navyFg:'FFFFFFFF',
      blueBg:'FFDAE8F0',blueFg:'FF1F4E79',
      green:'FFC6EFCE',greenFg:'FF375623',
      yellow:'FFFFEB9C',yellowFg:'FF9C6500',
      red:'FFFFC7CE',redFg:'FF9C0006',
      orange:'FFFCE4D6',orangeFg:'FFB85C00',
      grey:'FFF2F2F2',greyFg:'FF595959',
      white:'FFFFFFFF',darkText:'FF1F1F1F',
      altRow:'FFF7FBFF'
    };
    function hdr(){return{font:{bold:true,color:{rgb:C.navyFg},name:'Arial',sz:10},fill:{fgColor:{rgb:C.navyBg}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'medium',color:{rgb:C.navyBg}},right:{style:'thin',color:{rgb:'FFAAAAAA'}}}};}
    function cellS(bg,fg,center){return{font:{color:{rgb:fg||C.darkText},name:'Arial',sz:10},fill:{fgColor:{rgb:bg||C.white}},alignment:{horizontal:center?'center':'left',vertical:'center',wrapText:true},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}
    function secHdr(){return{font:{bold:true,color:{rgb:C.navyFg},name:'Arial',sz:11},fill:{fgColor:{rgb:'FF2E75B6'}},alignment:{horizontal:'left',vertical:'center'},border:{bottom:{style:'medium',color:{rgb:'FF1F4E79'}}}};}
    function labelS(){return{font:{bold:true,color:{rgb:C.greyFg},name:'Arial',sz:10},fill:{fgColor:{rgb:C.grey}},alignment:{horizontal:'left',vertical:'center'},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}
    function valS(bg,fg){return{font:{color:{rgb:fg||C.darkText},name:'Arial',sz:10},fill:{fgColor:{rgb:bg||C.white}},alignment:{horizontal:'left',vertical:'center'},border:{right:{style:'thin',color:{rgb:'FFD9D9D9'}},bottom:{style:'thin',color:{rgb:'FFD9D9D9'}}}};}

    // Raw
    var rawHdrs=['Ticket ID','Country','Login','Agent Name','Start Date','Start Status','End Date','End Status','Duration','Note'];
    var rawData=[rawHdrs].concat(state.results.map(function(r){return[r.ticketId,r.country,r.login,r.agentName,r.startDate,r.startStatus,r.endDate,r.endStatus,r.duration,r.note];}));
    var wsRaw=XLSX.utils.aoa_to_sheet(rawData);
    wsRaw['!cols']=[{wch:14},{wch:12},{wch:18},{wch:26},{wch:20},{wch:22},{wch:20},{wch:24},{wch:12},{wch:30}];
    var rr=XLSX.utils.decode_range(wsRaw['!ref']);
    for(var c0=rr.s.c;c0<=rr.e.c;c0++){var ha=XLSX.utils.encode_cell({r:0,c:c0});if(wsRaw[ha])wsRaw[ha].s=hdr();}
    for(var r0=1;r0<=rr.e.r;r0++){
      var nAddr=XLSX.utils.encode_cell({r:r0,c:9});var nVal=wsRaw[nAddr]?wsRaw[nAddr].v:'';var bg0=r0%2===0?C.altRow:C.white;
      for(var c1=rr.s.c;c1<=rr.e.c;c1++){
        var addr=XLSX.utils.encode_cell({r:r0,c:c1});if(!wsRaw[addr])wsRaw[addr]={t:'s',v:''};
        wsRaw[addr].s=nVal==='Ожидает ответа в чате Q&A Hub'?cellS(C.orange,C.orangeFg):nVal==='Дубликат'?cellS(C.blueBg,C.blueFg):nVal==='БЕЗ IN PROGRESS'?cellS(C.yellow,C.yellowFg):cellS(bg0);
      }
    }
    wsRaw['!freeze']={xSplit:0,ySplit:1};
    XLSX.utils.book_append_sheet(wb,wsRaw,'Raw');

    // Processed
    var sorted=state.results.filter(function(r){return r.startDate&&r.endDate&&r.hadInProgress;}).sort(function(a,b){return new Date(a.startDate.replace(' ','T'))-new Date(b.startDate.replace(' ','T'));});
    var intervals=[],procTimes=[],statusCount={},countryCount={},prevEnd=null;
    var procData=[['Ticket ID','Country','Start Time','End Time','Processing Time','Interval Since Previous','End Status','Note']];
    sorted.forEach(function(row,pi){
      var start=new Date(row.startDate.replace(' ','T')),end=new Date(row.endDate.replace(' ','T'));
      var procMs=end-start,intrvMs=(prevEnd&&pi>0)?Math.max(0,start-prevEnd):0;
      prevEnd=end;
      procTimes.push({ticket:row.ticketId,ms:procMs});
      if(pi>0)intervals.push({ticket:row.ticketId,ms:intrvMs});
      if(row.endStatus)statusCount[row.endStatus]=(statusCount[row.endStatus]||0)+1;
      if(row.country)countryCount[row.country]=(countryCount[row.country]||0)+1;
      procData.push([row.ticketId,row.country,row.startDate,row.endDate,msToHMS(procMs),pi>0?msToHMS(intrvMs):'—',row.endStatus,row.note]);
    });
    var wsProc=XLSX.utils.aoa_to_sheet(procData);
    wsProc['!cols']=[{wch:14},{wch:12},{wch:20},{wch:20},{wch:16},{wch:22},{wch:24},{wch:30}];
    var pr=XLSX.utils.decode_range(wsProc['!ref']);
    for(var c2=pr.s.c;c2<=pr.e.c;c2++){var hp=XLSX.utils.encode_cell({r:0,c:c2});if(wsProc[hp])wsProc[hp].s=hdr();}
    for(var r1=1;r1<=pr.e.r;r1++){
      var iCell=XLSX.utils.encode_cell({r:r1,c:5}),iVal=wsProc[iCell]?wsProc[iCell].v:'—',iMin=0;
      if(iVal&&iVal!=='—'){var pts=iVal.split(':');iMin=(+pts[0])*60+(+pts[1]);}
      var ibg=iVal==='—'||iMin===0?C.white:iMin<=INTERVAL_GREEN?C.green:iMin<=INTERVAL_YELLOW?C.yellow:C.red;
      var ifg=iVal==='—'||iMin===0?C.darkText:iMin<=INTERVAL_GREEN?C.greenFg:iMin<=INTERVAL_YELLOW?C.yellowFg:C.redFg;
      for(var c3=pr.s.c;c3<=pr.e.c;c3++){var pa=XLSX.utils.encode_cell({r:r1,c:c3});if(!wsProc[pa])wsProc[pa]={t:'s',v:''};wsProc[pa].s=cellS(ibg,ifg);}
    }
    wsProc['!freeze']={xSplit:0,ySplit:1};
    XLSX.utils.book_append_sheet(wb,wsProc,'Processed');

    // Dashboard
    var sumIntv=sumMsArray(intervals.map(function(i){return i.ms;}));
    var avgIntrvMs=intervals.length>0?sumIntv.ms/intervals.length:0;
    var maxIntv=intervals.reduce(function(m,i){return i.ms>m.ms?i:m;},{ms:0,ticket:'-'});
    var sumProc=sumMsArray(procTimes.map(function(i){return i.ms;}));
    var avgProcMs=procTimes.length>0?sumProc.ms/procTimes.length:0;
    var maxProc=procTimes.reduce(function(m,i){return i.ms>m.ms?i:m;},{ms:0,ticket:'-'});
    var topSK='',topSV=0;
    Object.keys(statusCount).forEach(function(k){if(statusCount[k]>topSV){topSV=statusCount[k];topSK=k;}});
    var noIPCount=state.results.filter(function(r){return!r.hadInProgress&&!r.isDuplicate;}).length;
    var dupCount=state.results.filter(function(r){return r.isDuplicate;}).length;
    var seniorCount=state.results.filter(function(r){return r.isSenior;}).length;
    var inProgNow=state.results.filter(function(r){return r.startDate&&!r.endDate;}).sort(function(a,b){return new Date(a.startDate.replace(' ','T'))-new Date(b.startDate.replace(' ','T'));});
    var breakExceeded=sumIntv.ms>=BREAK_LIMIT_MS;
    var statusArr=Object.keys(state.statusStat).map(function(k){return{k:k,v:state.statusStat[k]};}).sort(function(a,b){return b.v-a.v;});
    var countryArr=Object.keys(state.countryStat).map(function(k){return{k:k,v:state.countryStat[k]};}).sort(function(a,b){return b.v-a.v;});
    var topIntervals=intervals.filter(function(i){return i.ms>=MIN_INTERVAL_SHOW_MS;}).sort(function(a,b){return b.ms-a.ms;});

    var ws={},ROW=0;
    function w(col,v,style,type){var a=XLSX.utils.encode_cell({r:ROW,c:col});ws[a]={t:type||'s',v:v==null?'':v};if(style)ws[a].s=style;}
    function nr(){ROW++;}function br(){ROW++;}

    w(0,'ОБЩАЯ СТАТИСТИКА',secHdr());w(1,'',secHdr());w(2,'',secHdr());w(3,'',secHdr());nr();
    [
      ['Агент',state.agentName+' ('+state.agentLogin+')',null],
      ['Период',state.shiftStart.toLocaleString('ru')+' — '+state.shiftEnd.toLocaleString('ru'),null],
      ['Всего записей',state.results.length,'n'],
      ['Тикетов с полными датами',sorted.length,'n'],
      ['Самый частый статус',topSK+' ('+topSV+')',null],
      ['Дубликатов',dupCount,'n'],
      ['Без In Progress',noIPCount,'n'],
      ['Ожидали Q&A Hub',seniorCount,'n'],
      ['Среднее время обработки',msToHMS(avgProcMs),null],
      ['Макс. время обработки',msToHMS(maxProc.ms)+' (тикет '+maxProc.ticket+')',null],
      ['Средний интервал',msToHMS(avgIntrvMs),null],
      ['Суммарные перерывы',sumIntv.hms,breakExceeded?[C.red,C.redFg]:[C.green,C.greenFg]],
      ['Лимит перерывов',breakExceeded?'ПРЕВЫШЕН':'В норме',breakExceeded?[C.red,C.redFg]:[C.green,C.greenFg]],
    ].forEach(function(s){
      w(0,s[0],labelS());w(1,s[1],valS(s[2]?s[2][0]:C.white,s[2]?s[2][1]:C.darkText),s[3]||null);
      w(2,'',cellS(C.white));w(3,'',cellS(C.white));nr();
    });
    br();
    w(0,'СТАТУСЫ',secHdr());w(1,'',secHdr());w(2,'',secHdr());w(3,'',secHdr());nr();
    w(0,'Статус',labelS());w(1,'Кол-во',labelS());w(2,'',cellS(C.white));w(3,'',cellS(C.white));nr();
    statusArr.forEach(function(s,i){var rb=i%2===0?C.white:C.altRow;w(0,s.k,cellS(rb));w(1,s.v,cellS(rb,C.darkText,true),'n');w(2,'',cellS(C.white));w(3,'',cellS(C.white));nr();});
    w(0,'ИТОГО',labelS());w(1,state.results.length,valS(C.blueBg,C.blueFg),'n');nr();br();
    w(0,'ТОП СТРАН',secHdr());w(1,'',secHdr());w(2,'',secHdr());w(3,'',secHdr());nr();
    w(0,'Страна',labelS());w(1,'Тикетов',labelS());w(2,'',cellS(C.white));w(3,'',cellS(C.white));nr();
    countryArr.slice(0,10).forEach(function(c,i){var cb=i%2===0?C.white:C.altRow;w(0,c.k,cellS(cb));w(1,c.v,cellS(cb,C.darkText,true),'n');w(2,'',cellS(C.white));w(3,'',cellS(C.white));nr();});
    br();
    w(0,'В РАБОТЕ ПРЯМО СЕЙЧАС',secHdr());w(1,'',secHdr());w(2,'',secHdr());w(3,'',secHdr());nr();
    w(0,'Ticket ID',labelS());w(1,'Country',labelS());w(2,'Взят в работу',labelS());w(3,'Висит уже',labelS());nr();
    var nowMs=Date.now();
    if(inProgNow.length===0){w(0,'Нет активных тикетов',cellS(C.green,C.greenFg));w(1,'',cellS(C.green));w(2,'',cellS(C.green));w(3,'',cellS(C.green));nr();}
    else{inProgNow.forEach(function(ipr){var hang=nowMs-new Date(ipr.startDate.replace(' ','T')).getTime();var hMin=Math.floor(hang/60000);var ibg=hMin>60?C.red:hMin>30?C.yellow:C.green;var ifg=hMin>60?C.redFg:hMin>30?C.yellowFg:C.greenFg;w(0,ipr.ticketId,cellS(ibg,ifg));w(1,ipr.country,cellS(ibg,ifg));w(2,ipr.startDate,cellS(ibg,ifg));w(3,msToHMS(hang),cellS(ibg,ifg,true));nr();});}

    var savedROW=ROW;ROW=0;
    w(5,'ПЕРЕРЫВЫ (>= 5 мин)',secHdr());w(6,'',secHdr());w(7,'',secHdr());w(8,'',secHdr());nr();
    w(5,'#',labelS());w(6,'Ticket ID',labelS());w(7,'Интервал',labelS());w(8,'Оценка',labelS());nr();
    if(topIntervals.length===0){w(5,'Нет перерывов >= 5 мин',cellS(C.grey));w(6,'',cellS(C.grey));w(7,'',cellS(C.grey));w(8,'',cellS(C.grey));nr();}
    else{topIntervals.forEach(function(intv,ti){var iMin=Math.floor(intv.ms/60000);var tbg=iMin<=INTERVAL_GREEN?C.green:iMin<=INTERVAL_YELLOW?C.yellow:C.red;var tfg=iMin<=INTERVAL_GREEN?C.greenFg:iMin<=INTERVAL_YELLOW?C.yellowFg:C.redFg;var assess=iMin<=INTERVAL_GREEN?'Норма':iMin<=INTERVAL_YELLOW?'Внимание':'Нарушение';w(5,ti+1,cellS(tbg,tfg,true),'n');w(6,intv.ticket,cellS(tbg,tfg));w(7,msToHMS(intv.ms),cellS(tbg,tfg,true));w(8,assess,cellS(tbg,tfg,true));nr();});}
    w(5,'Суммарно (все)',labelS());w(6,'',cellS(C.white));
    w(7,sumIntv.hms,valS(breakExceeded?C.red:C.green,breakExceeded?C.redFg:C.greenFg));
    w(8,breakExceeded?'ПРЕВЫШЕН':'В норме',valS(breakExceeded?C.red:C.green,breakExceeded?C.redFg:C.greenFg));
    ROW=Math.max(ROW,savedROW);

    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:ROW+5,c:8}});
    ws['!cols']=[{wch:28},{wch:32},{wch:16},{wch:22},{wch:3},{wch:6},{wch:16},{wch:14},{wch:14}];
    var rowsArr=[];for(var ri=0;ri<=ROW+5;ri++)rowsArr.push({hpt:20});
    ws['!rows']=rowsArr;
    XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

    var filename='report_'+state.agentLogin+'_'+new Date().toISOString().slice(0,10)+'.xlsx';
    XLSX.writeFile(wb,filename);
    document.getElementById('tm-btn-start').disabled=false;
    setProgress(state.done,state.done);
    setStatus('✅ Скачан: '+filename+' | Записей: '+state.results.length);
    console.log('XLSX скачан: '+filename);
  }

})();
