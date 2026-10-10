/* WhisperFly — live dashboard.
   Classic script (no ES modules): `import` is blocked over file://, and this page
   has to work when someone double-clicks dashboard.html. No libraries, no network.
   Every number on the page comes from the DATA block below, so there is exactly
   one place to change an assumption. */
(function () {
  'use strict';

  /* ── DATA ─────────────────────────────────────────────────────────────
     Figures mirror the nine slides of WhisperFly-Pitch-RU.pptx. Keep them in
     step: the deck and this dashboard must never disagree in public. */
  var FX = 530;          /* ₸ per $1, as stated on the funding slide      */
  var EUR_USD = 1.08;    /* only used to price MacWhisper's €59            */
  var HOURS_PER_FTE = 1750;
  var SECONDS_PER_YEAR = 365 * 24 * 60 * 60;

  var BUDGET = [
    { name: 'Оплата труда основателя', value: 1200000, note: '200 000 ₸ × 6 мес' },
    { name: 'AI-модели и облачные API', value: 530000, note: '$1 000 · уже оплачено' },
    { name: 'Продвижение лендинга', value: 250000, note: 'первые установки' },
    { name: 'Домен, хостинг и аналитика', value: 60000 },
    { name: 'Apple Developer Program', value: 54000, note: '$99 в год' },
    { name: 'Резерв на непредвиденные', value: 106000 }
  ];

  var GRANT_TENGE = 10000 * FX;          /* ≈ 5 300 000 ₸ */
  var BUDGET_TOTAL = BUDGET.reduce(function (sum, item) { return sum + item.value; }, 0);

  /* Стоимость владения. `sub` = ежемесячная подписка, `once` = разовая покупка. */
  var COMPETITORS = [
    { name: 'WhisperFly', note: 'открытый код · 0 ₸', kind: 'free' },
    { name: 'Superwhisper', note: '$8,99 в месяц', kind: 'sub', usd: 8.99 },
    { name: 'MacWhisper', note: '€59 разово', kind: 'once', eur: 59 },
    { name: 'Wispr Flow', note: '$12 в месяц', kind: 'sub', usd: 12 },
    { name: 'Apple «Диктовка»', note: '0 ₸ · только микрофон', kind: 'free' }
  ];

  var SCENARIOS = {
    conservative: [260, 520, 820, 1080, 1350, 1650, 1950, 2250, 2560, 2870, 3140, 3400],
    base:         [333, 667, 1000, 1300, 1650, 2050, 2500, 3000, 3500, 4000, 4500, 5000],
    optimistic:   [420, 840, 1250, 1750, 2350, 3050, 3850, 4750, 5750, 6850, 8050, 9350]
  };
  var SCENARIO_LABEL = {
    conservative: 'Консервативный',
    base: 'Базовый',
    optimistic: 'Оптимистичный'
  };
  var MONTHS = ['Окт 26', 'Ноя 26', 'Дек 26', 'Янв 27', 'Фев 27', 'Мар 27',
                'Апр 27', 'Май 27', 'Июн 27', 'Июл 27', 'Авг 27', 'Сен 27'];
  var TARGET = 5000;

  var ROADMAP = [
    {
      q: 'Q4 2026', state: 'Этап закрыт', cls: 'is-done', done: 1,
      items: ['Релиз 2.0 выпущен', 'Системный звук и файлы', 'История и 11 языков',
              'Сборки: ARM, Intel, универсальная'],
      note: 'Продукт уже в руках пользователей: 2.0 вышел со всем заявленным функционалом — ' +
            'включая транскрипцию системного звука и готовых файлов. Это первый этап, который ' +
            'не нужно защищать словами, его можно показать.'
    },
    {
      q: 'Q1 2027', state: 'В работе сейчас', cls: 'is-now', done: 0.35,
      items: ['Лендинг продукта', 'Заявка на грант GitHub', 'Первые 1 000 установок',
              'Обратная связь и метрики'],
      note: 'Ключевой квартал: он заканчивается цифрой в 1 000 установок. Именно по ней ' +
            'будет видно, работает ли лендинг и хватает ли продукту одной сарафанной волны.'
    },
    {
      q: 'Q2 2027', state: 'Запланировано', cls: '', done: 0,
      items: ['WhisperFly Pro', 'Оплата и лицензии', 'Свой словарь',
              'Найм: маркетинг'],
      note: 'Здесь появляется выручка. Ядро при этом остаётся бесплатным — платной становится ' +
            'только расширенная версия, а себестоимость распознавания измеряется центами на пользователя.'
    },
    {
      q: 'Q3 2027', state: 'Запланировано', cls: '', done: 0,
      items: ['Потоковая запись', 'iOS-компаньон', 'Версия для Windows',
              'Публичный API'],
      note: 'Расширение за пределы macOS. Потоковая запись убирает ожидание после остановки, ' +
            'а публичный API открывает продукт для сторонних интеграций.'
    },
    {
      q: 'Q4 2027', state: 'Запланировано', cls: '', done: 0,
      items: ['Админ-панель и SSO', 'Общие словари', 'Развёртывание on-prem',
              'Аудит безопасности'],
      note: 'Выход в корпоративный сегмент: общие словари и SSO превращают личный инструмент ' +
            'в инфраструктуру команды, а аудит безопасности снимает последнее возражение закупки.'
    }
  ];

  var PRESETS = {
    cautious:  { users: 2000,  minutes: 20, days: 200, rate: 1200 },
    base:      { users: 5000,  minutes: 40, days: 220, rate: 1500 },
    ambitious: { users: 12000, minutes: 60, days: 240, rate: 2500 }
  };

  var DEMO_LINES = [
    'Привет, команда. Отправь, пожалуйста, отчёт до пятницы — я подпишу его сегодня.',
    'Напомни завтра в девять утра созвон с маркетингом и подготовь короткую повестку.',
    'Продиктуй это в чат: встреча переносится на четырнадцать тридцать, зал на третьем этаже.'
  ];

  /* ── Helpers ──────────────────────────────────────────────────────────── */
  function $(id) { return document.getElementById(id); }

  function all(selector, root) {
    return Array.prototype.slice.call((root || document).querySelectorAll(selector));
  }

  /* 1234567 → "1 234 567"; decimals are joined with a Russian comma. */
  function nf(value, decimals) {
    var places = decimals || 0;
    var parts = Math.abs(value).toFixed(places).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    var out = parts.length > 1 ? parts[0] + ',' + parts[1] : parts[0];
    return (value < 0 ? '−' : '') + out;
  }

  function money(value) { return nf(value) + ' ₸'; }

  function moneyShort(value) {
    if (value >= 1e9) { return nf(value / 1e9, 1) + ' млрд ₸'; }
    if (value >= 1e6) { return nf(value / 1e6, 1) + ' млн ₸'; }
    if (value >= 1e3) { return nf(value / 1e3, 0) + ' тыс ₸'; }
    return money(value);
  }

  /* 733333 → 0,73 млн — the deck's headline figure is 0,7 млн, so large values stay
     in millions rather than degrading to "733,3 тыс". */
  function hoursShort(hours) {
    if (hours >= 1e6) { return nf(hours / 1e6, 1) + ' млн'; }
    if (hours >= 2e5) { return nf(hours / 1e6, 2) + ' млн'; }
    if (hours >= 1e3) { return nf(hours / 1e3, 1) + ' тыс'; }
    return nf(hours);
  }

  function svgEl(name, attrs) {
    var node = document.createElementNS('http://www.w3.org/2000/svg', name);
    Object.keys(attrs || {}).forEach(function (key) { node.setAttribute(key, attrs[key]); });
    return node;
  }

  function clamp(value, min, max) { return Math.min(max, Math.max(min, value)); }

  var reduceMotion = window.matchMedia &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ── 1 · Reveal on scroll ─────────────────────────────────────────────── */
  function initReveal() {
    var targets = all('[data-rv]');
    if (!targets.length) { return; }

    if (!('IntersectionObserver' in window)) {
      targets.forEach(function (node) { node.classList.add('in'); });
      return;
    }

    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) { return; }
        entry.target.classList.add('in');
        observer.unobserve(entry.target);
      });
    }, { rootMargin: '0px 0px -8% 0px', threshold: 0.08 });

    targets.forEach(function (node) { observer.observe(node); });
  }

  /* ── 2 · Count-up numbers ─────────────────────────────────────────────── */
  function formatCount(value, node) {
    var decimals = parseInt(node.getAttribute('data-decimals') || '0', 10);
    var suffix = node.getAttribute('data-suffix') || '';
    return nf(value, decimals) + suffix;
  }

  function runCounter(node) {
    var to = parseFloat(node.getAttribute('data-count'));
    if (isNaN(to)) { return; }

    if (reduceMotion) {
      node.textContent = formatCount(to, node);
      return;
    }

    var duration = 1500;
    var start = null;

    function step(timestamp) {
      if (start === null) { start = timestamp; }
      var progress = clamp((timestamp - start) / duration, 0, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      node.textContent = formatCount(to * eased, node);
      if (progress < 1) { window.requestAnimationFrame(step); }
    }

    node.textContent = formatCount(0, node);
    window.requestAnimationFrame(step);
  }

  function initCounters() {
    var nodes = all('[data-count]');
    if (!nodes.length) { return; }

    if (!('IntersectionObserver' in window)) {
      nodes.forEach(runCounter);
      return;
    }

    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) { return; }
        runCounter(entry.target);
        observer.unobserve(entry.target);
      });
    }, { threshold: 0.4 });

    nodes.forEach(function (node) { observer.observe(node); });
  }

  /* ── 3 · Live dictation demo (hero) ───────────────────────────────────── */
  var STATUS = {
    listen: { icon: '#i-mic', text: 'Слушаю микрофон…' },
    think:  { icon: '#i-cloud', text: 'Распознаю в облаке…' },
    done:   { icon: '#i-check', text: 'Вставлено в активное поле' }
  };

  function setStatus(key) {
    var node = $('demoStatus');
    if (!node) { return; }
    var state = STATUS[key];
    node.classList.toggle('is-done', key === 'done');
    node.innerHTML =
      '<svg class="ds-ico" aria-hidden="true"><use href="' + state.icon + '"/></svg>' +
      '<span>' + state.text + '</span>';
  }

  function initDemo() {
    var wave = $('wave');
    var field = $('demoText');
    var lang = $('demoLang');
    if (!wave || !field) { return; }

    var BARS = 28;
    var bars = [];
    for (var i = 0; i < BARS; i += 1) {
      var bar = document.createElement('span');
      wave.appendChild(bar);
      bars.push(bar);
    }

    if (reduceMotion) {
      bars.forEach(function (bar, index) {
        bar.style.height = (18 + ((index * 37) % 62)) + '%';
      });
      field.textContent = DEMO_LINES[0];
      setStatus('done');
      return;
    }

    var listening = true;
    var tick = 0;

    window.setInterval(function () {
      tick += 1;
      bars.forEach(function (bar, index) {
        var base = listening ? 16 : 6;
        var spread = listening ? 78 : 10;
        var wave1 = Math.sin((tick + index * 1.7) / 2.4);
        var wave2 = Math.sin((tick * 1.9 + index) / 5.1);
        var level = base + (0.55 + 0.45 * wave1) * spread * (0.65 + 0.35 * Math.abs(wave2));
        bar.style.height = clamp(level, 6, 100) + '%';
      });
    }, 90);

    var lineIndex = 0;
    var charIndex = 0;
    var phase = 'typing';

    function advance() {
      var line = DEMO_LINES[lineIndex];

      if (phase === 'typing') {
        charIndex += 2 + Math.floor(Math.random() * 3);
        field.textContent = line.slice(0, charIndex);
        if (charIndex >= line.length) {
          field.textContent = line;
          listening = false;
          setStatus('think');
          phase = 'thinking';
          window.setTimeout(advance, 420);
          return;
        }
        window.setTimeout(advance, 28 + Math.random() * 45);
        return;
      }

      if (phase === 'thinking') {
        setStatus('done');
        phase = 'done';
        window.setTimeout(advance, 1900);
        return;
      }

      /* done → clear and start the next phrase */
      field.textContent = '';
      charIndex = 0;
      lineIndex = (lineIndex + 1) % DEMO_LINES.length;
      phase = 'typing';
      listening = true;
      setStatus('listen');
      if (lang) {
        lang.textContent = ['RU', 'EN', 'KK'][lineIndex % 3];
      }
      window.setTimeout(advance, 700);
    }

    setStatus('listen');
    window.setTimeout(advance, 600);
  }

  /* ── 4 · Live clock + live savings ticker ─────────────────────────────── */
  var live = { users: 5000, minutes: 40, days: 220, rate: 1500 };
  var startedAt = Date.now();

  function tickClock() {
    var node = $('liveClock');
    if (!node) { return; }
    var total = Math.floor((Date.now() - startedAt) / 1000);
    var h = Math.floor(total / 3600);
    var m = Math.floor((total % 3600) / 60);
    var s = total % 60;
    node.textContent = (h < 10 ? '0' : '') + h + ':' +
                       (m < 10 ? '0' : '') + m + ':' +
                       (s < 10 ? '0' : '') + s;
  }

  function hoursPerYear() {
    return (live.users * live.minutes * live.days) / 60;
  }

  function tickLive() {
    var node = $('liveHours');
    if (!node) { return; }
    var perSecond = hoursPerYear() / SECONDS_PER_YEAR;
    var elapsed = (Date.now() - startedAt) / 1000;
    node.textContent = nf(elapsed * perSecond, 3);
  }

  /* Recomputed whenever the calculator moves — the strip and the model share state. */
  function renderLiveStats() {
    var hours = hoursPerYear();
    var perSecond = hours / SECONDS_PER_YEAR;

    var rate = $('liveRate');
    if (rate) { rate.textContent = nf(perSecond * 60, 2) + ' мин'; }

    var perDay = $('livePerDay');
    if (perDay) { perDay.textContent = nf(live.users * live.minutes / 60); }

    var fte = $('liveFte');
    if (fte) { fte.textContent = nf(hours / HOURS_PER_FTE); }
  }

  function initLive() {
    tickClock();
    tickLive();
    renderLiveStats();
    window.setInterval(tickClock, 1000);
    window.setInterval(tickLive, 120);
  }

  /* ── 5 · Scroll progress + section spy ────────────────────────────────── */
  function initScroll() {
    var rail = $('scrollProgress');
    var links = all('.tb-links a[href^="#"]');
    var sections = links
      .map(function (link) { return document.querySelector(link.getAttribute('href')); })
      .filter(Boolean);

    function onScroll() {
      var doc = document.documentElement;
      var max = doc.scrollHeight - window.innerHeight;
      var ratio = max > 0 ? (window.scrollY || doc.scrollTop) / max : 0;
      if (rail) { rail.style.width = (clamp(ratio, 0, 1) * 100).toFixed(2) + '%'; }

      var marker = window.scrollY + window.innerHeight * 0.32;
      var activeIndex = -1;
      sections.forEach(function (section, index) {
        if (section.offsetTop <= marker) { activeIndex = index; }
      });
      links.forEach(function (link, index) {
        link.classList.toggle('is-current', index === activeIndex);
      });
    }

    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('resize', onScroll);
    onScroll();
  }
/* ── 6 · Cost-of-ownership bars ───────────────────────────────────────── */
  var compareYears = 1;

  function owningCost(item, years) {
    if (item.kind === 'sub') { return item.usd * 12 * years * FX; }
    if (item.kind === 'once') { return item.eur * EUR_USD * FX; }
    return 0;
  }

  function initCompare() {
    var host = $('compareBars');
    if (!host) { return; }

    var rows = COMPETITORS.map(function (item, index) {
      var row = document.createElement('div');
      row.className = 'bar-row' + (index === 0 ? ' is-us' : '');
      row.setAttribute('data-key', 'c' + index);

      var label = document.createElement('span');
      label.className = 'bar-label';
      label.textContent = item.name;
      var note = document.createElement('span');
      note.className = 'bar-note';
      note.textContent = item.note;
      label.appendChild(note);

      var track = document.createElement('span');
      track.className = 'bar-track';
      var fill = document.createElement('span');
      fill.className = 'bar-fill';
      track.appendChild(fill);

      var value = document.createElement('span');
      value.className = 'bar-val';

      row.appendChild(label);
      row.appendChild(track);
      row.appendChild(value);
      host.appendChild(row);

      return { row: row, fill: fill, value: value, item: item };
    });

    function paint(years) {
      var costs = COMPETITORS.map(function (item) { return owningCost(item, years); });
      var max = costs.reduce(function (a, b) { return Math.max(a, b); }, 1);

      rows.forEach(function (entry, index) {
        var cost = costs[index];
        var share = cost === 0 ? 0 : Math.max(2.5, (cost / max) * 100);
        entry.fill.style.width = share + '%';
        entry.value.textContent = cost === 0 ? '0 ₸' : money(cost);
      });

      var noteNode = $('compareNote');
      if (noteNode) {
        noteNode.textContent = 'Стоимость владения за ' + years + ' ' +
          (years === 1 ? 'год' : (years < 5 ? 'года' : 'лет')) + ', ₸';
      }
    }

    paint(compareYears);

    all('.seg-btn[data-years]').forEach(function (button) {
      button.addEventListener('click', function () {
        all('.seg-btn[data-years]').forEach(function (other) {
          other.classList.toggle('active', other === button);
        });
        compareYears = parseInt(button.getAttribute('data-years'), 10) || 1;
        paint(compareYears);
      });
    });

    return rows;
  }

  /* ── 7 · Comparison matrix ⇄ bar chart cross-highlight ────────────────── */
  function initMatrix(rows) {
    var table = $('matrix');
    if (!table) { return; }

    /* Animate the score bars in once the table is on screen. */
    var fills = all('.score-fill', table);
    fills.forEach(function (fill) {
      var target = parseInt(fill.style.getPropertyValue('--p'), 10) || 0;
      fill.setAttribute('data-target', target);
      fill.style.setProperty('--p', 0);
    });

    var revealed = false;
    function revealScores() {
      if (revealed) { return; }
      revealed = true;
      window.requestAnimationFrame(function () {
        fills.forEach(function (fill) {
          fill.style.setProperty('--p', fill.getAttribute('data-target'));
        });
      });
    }

    if ('IntersectionObserver' in window) {
      var observer = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) { revealScores(); observer.disconnect(); }
        });
      }, { threshold: 0.25 });
      observer.observe(table);
    } else {
      revealScores();
    }

    /* Hover a product row → its bar in the cost chart lights up. */
    if (!rows) { return; }
    all('tbody tr', table).forEach(function (tr, index) {
      var key = 'c' + index;
      var linked = rows.filter(function (entry) { return entry.row.getAttribute('data-key') === key; })[0];
      if (!linked) { return; }
      tr.addEventListener('mouseenter', function () { linked.row.classList.add('is-hover'); });
      tr.addEventListener('mouseleave', function () { linked.row.classList.remove('is-hover'); });
    });
  }

  /* ── 8 · КПД calculator ───────────────────────────────────────────────── */
  var sliderIds = ['calcUsers', 'calcMinutes', 'calcDays', 'calcRate'];

  function renderCalc() {
    var hours = hoursPerYear();
    var fte = hours / HOURS_PER_FTE;
    var tenge = hours * live.rate;
    var subscriptions = live.users * 8.99 * 12 * FX;
    var dayShare = (live.minutes / (8 * 60)) * 100;

    var set = function (id, text) { var node = $(id); if (node) { node.textContent = text; } };

    set('outUsers', nf(live.users));
    set('outMinutes', nf(live.minutes) + ' мин');
    set('outDays', nf(live.days));
    set('outRate', nf(live.rate) + ' ₸');

    set('calcHours', hoursShort(hours));
    set('calcHoursExact', nf(hours) + ' часов');

    set('fUsers', nf(live.users));
    set('fMinutes', nf(live.minutes));
    set('fDays', nf(live.days));
    set('fResult', nf(hours));

    set('calcFte', nf(fte));
    set('calcTenge', moneyShort(tenge));
    set('calcSaved', moneyShort(subscriptions));
    set('meterNote', nf(live.minutes) + ' мин из 8-часового дня = ' + nf(dayShare, 1) + ' %');

    var meter = $('meterFill');
    if (meter) { meter.style.width = clamp(dayShare, 1.5, 100) + '%'; }

    renderLiveStats();
  }

  function moveSlider(id, value) {
    var input = $(id);
    if (!input) { return; }
    input.value = value;
  }

  function initCalc() {
    var inputs = sliderIds.map($).filter(Boolean);
    if (!inputs.length) { return; }

    inputs.forEach(function (input) {
      input.addEventListener('input', function () {
        live.users = parseInt($('calcUsers').value, 10);
        live.minutes = parseInt($('calcMinutes').value, 10);
        live.days = parseInt($('calcDays').value, 10);
        live.rate = parseInt($('calcRate').value, 10);
        all('.preset-btn').forEach(function (button) { button.classList.remove('active'); });
        renderCalc();
      });
    });

    all('.preset-btn').forEach(function (button) {
      button.addEventListener('click', function () {
        var preset = PRESETS[button.getAttribute('data-preset')];
        if (!preset) { return; }
        all('.preset-btn').forEach(function (other) {
          other.classList.toggle('active', other === button);
        });
        moveSlider('calcUsers', preset.users);
        moveSlider('calcMinutes', preset.minutes);
        moveSlider('calcDays', preset.days);
        moveSlider('calcRate', preset.rate);
        live.users = preset.users;
        live.minutes = preset.minutes;
        live.days = preset.days;
        live.rate = preset.rate;
        renderCalc();
      });
    });

    renderCalc();
  }

  /* ── 9 · Growth projection chart ─────────────────────────────────────── */
  var scenario = 'base';
  var growthHost = null;
  var growthTimer = null;

  function renderGrowth() {
    var host = growthHost;
    if (!host) { return; }

    var data = SCENARIOS[scenario];
    var tip = $('growthTip');
    var width = Math.max(300, host.clientWidth || 640);
    var height = width < 520 ? 240 : 300;

    var pad = { t: 18, r: 18, b: width < 520 ? 46 : 38, l: width < 520 ? 38 : 52 };
    var innerW = width - pad.l - pad.r;
    var innerH = height - pad.t - pad.b;

    var peak = Math.max(TARGET, data[data.length - 1]);
    var step = peak > 6000 ? 2000 : (peak > 3000 ? 1250 : 1000);
    var top = Math.ceil((peak * 1.06) / step) * step;

    function x(index) { return pad.l + (innerW * index) / (data.length - 1); }
    function y(value) { return pad.t + innerH * (1 - value / top); }

    while (host.firstChild) { host.removeChild(host.firstChild); }

    var svg = svgEl('svg', {
      viewBox: '0 0 ' + width + ' ' + height,
      width: width,
      height: height,
      role: 'img',
      'aria-label': 'Накопительный рост установок: ' + SCENARIO_LABEL[scenario]
    });

    /* gradient for the area fill */
    var defs = svgEl('defs', {});
    var gradient = svgEl('linearGradient', { id: 'growthFill', x1: '0', y1: '0', x2: '0', y2: '1' });
    /* Colours come from CSS classes (.g-stop-*), not inline var() — the CSSOM does
       not resolve a custom property inside a presentation attribute, which would
       leave the area fill black. */
    var stopTop = svgEl('stop', { offset: '0%', class: 'g-stop-top' });
    var stopBottom = svgEl('stop', { offset: '100%', class: 'g-stop-bottom' });
    gradient.appendChild(stopTop);
    gradient.appendChild(stopBottom);
    defs.appendChild(gradient);
    svg.appendChild(defs);

    /* horizontal grid + y labels */
    for (var value = 0; value <= top; value += step) {
      var line = svgEl('line', { x1: pad.l, x2: width - pad.r, y1: y(value), y2: y(value), class: 'g-grid' });
      svg.appendChild(line);
      var label = svgEl('text', { x: pad.l - 10, y: y(value) + 4, class: 'g-label', 'text-anchor': 'end' });
      label.textContent = nf(value);
      svg.appendChild(label);
    }

    /* target line */
    var targetLine = svgEl('line', {
      x1: pad.l, x2: width - pad.r, y1: y(TARGET), y2: y(TARGET), class: 'g-target'
    });
    svg.appendChild(targetLine);
    var targetLabel = svgEl('text', {
      x: width - pad.r, y: y(TARGET) - 8, class: 'g-label', 'text-anchor': 'end'
    });
    targetLabel.textContent = 'Цель ' + nf(TARGET);
    svg.appendChild(targetLabel);

    /* baseline */
    svg.appendChild(svgEl('line', {
      x1: pad.l, x2: width - pad.r, y1: y(0), y2: y(0), class: 'g-axis'
    }));

    /* month ticks */
    [0, 3, 6, 9, 11].forEach(function (index) {
      var tick = svgEl('text', {
        x: x(index), y: height - pad.b + 20, class: 'g-label',
        'text-anchor': index === 0 ? 'start' : (index === 11 ? 'end' : 'middle')
      });
      tick.textContent = MONTHS[index];
      svg.appendChild(tick);
    });

    /* area + line */
    var areaPath = 'M' + x(0) + ' ' + y(0);
    data.forEach(function (point, index) { areaPath += ' L' + x(index) + ' ' + y(point); });
    areaPath += ' L' + x(data.length - 1) + ' ' + y(0) + ' Z';

    var area = svgEl('path', { d: areaPath, class: 'g-area' });
    svg.appendChild(area);

    var linePath = '';
    data.forEach(function (point, index) {
      linePath += (index === 0 ? 'M' : ' L') + x(index) + ' ' + y(point);
    });
    var line = svgEl('path', { d: linePath, class: 'g-line' });
    svg.appendChild(line);

    /* points */
    var dots = data.map(function (point, index) {
      var dot = svgEl('circle', { cx: x(index), cy: y(point), r: 3.6, class: 'g-dot' });
      svg.appendChild(dot);
      return dot;
    });

    /* invisible hover bands */
    var band = innerW / (data.length - 1);
    data.forEach(function (point, index) {
      var hit = svgEl('rect', {
        x: x(index) - band / 2, y: pad.t, width: band, height: innerH, class: 'g-hit'
      });
      hit.addEventListener('mouseenter', function () {
        dots.forEach(function (dot, i) { dot.setAttribute('r', i === index ? 6 : 3.6); });
        if (!tip) { return; }
        $('tipInstalls').textContent = nf(point) + ' установок';
        $('tipMonth').textContent = MONTHS[index] + ' · ' + SCENARIO_LABEL[scenario];
        tip.hidden = false;
        tip.style.left = x(index) + 'px';
        tip.style.top = y(point) + 'px';
      });
      svg.appendChild(hit);
    });

    svg.addEventListener('mouseleave', function () {
      dots.forEach(function (dot) { dot.setAttribute('r', 3.6); });
      if (tip) { tip.hidden = true; }
    });

    host.appendChild(svg);

    /* draw the line on, and fade the area in */
    if (!reduceMotion && line.getTotalLength) {
      var length = line.getTotalLength();
      line.style.strokeDasharray = length + ' ' + length;
      line.style.strokeDashoffset = length;
      window.requestAnimationFrame(function () {
        line.style.transition = 'stroke-dashoffset 1.35s cubic-bezier(.34,1.4,.64,1)';
        line.style.strokeDashoffset = '0';
      });
    }
    window.requestAnimationFrame(function () { area.classList.add('in'); });

    /* headline metrics */
    var end = $('growthEnd');
    var q1 = $('growthQ1');
    var peakNode = $('growthPeak');
    var monthly = data.map(function (point, index) {
      return index === 0 ? point : point - data[index - 1];
    });
    var maxMonth = monthly.reduce(function (a, b) { return Math.max(a, b); }, 0);

    if (end) { end.textContent = nf(data[data.length - 1]); }
    if (q1) { q1.textContent = nf(data[2]); }
    if (peakNode) { peakNode.textContent = '+' + nf(maxMonth); }
  }

  function initGrowth() {
    growthHost = $('growthPlot');
    if (!growthHost) { return; }

    renderGrowth();

    all('#scenarioSeg .seg-btn').forEach(function (button) {
      button.addEventListener('click', function () {
        all('#scenarioSeg .seg-btn').forEach(function (other) {
          other.classList.toggle('active', other === button);
        });
        scenario = button.getAttribute('data-scenario') || 'base';
        renderGrowth();
      });
    });

    window.addEventListener('resize', function () {
      window.clearTimeout(growthTimer);
      growthTimer = window.setTimeout(renderGrowth, 160);
    });
  }

  /* ── 10 · Roadmap ─────────────────────────────────────────────────────── */
  function initRoadmap() {
    var grid = $('roadGrid');
    var fill = $('roadRailFill');
    if (!grid) { return; }

    var buttons = all('.road-q', grid);

    function select(index) {
      var stage = ROADMAP[index];
      if (!stage) { return; }

      buttons.forEach(function (button, i) {
        button.classList.toggle('active', i === index);
        button.setAttribute('aria-pressed', i === index ? 'true' : 'false');
      });

      var detail = $('roadDetail');
      if (detail) {
        detail.classList.remove('is-done', 'is-now');
        if (stage.cls) { detail.classList.add(stage.cls); }
      }

      var quarter = $('rdQuarter');
      if (quarter) { quarter.textContent = stage.q; }
      var state = $('rdState');
      if (state) { state.textContent = stage.state; }

      var list = $('rdItems');
      if (list) {
        list.innerHTML = '';
        stage.items.forEach(function (item) {
          var li = document.createElement('li');
          li.textContent = item;
          list.appendChild(li);
        });
      }

      var note = $('rdNote');
      if (note) { note.textContent = stage.note; }
    }

    var elapsed = ROADMAP.reduce(function (sum, stage) { return sum + stage.done; }, 0);
    if (fill) {
      fill.style.width = clamp((elapsed / ROADMAP.length) * 100, 0, 100).toFixed(1) + '%';
    }

    buttons.forEach(function (button, index) {
      button.addEventListener('click', function () { select(index); });
    });

    select(0);
  }

  /* ── 11 · Goal cards with quarterly splits ────────────────────────────── */
  function initGoals() {
    all('.qbar').forEach(function (bar) {
      var split = (bar.getAttribute('data-split') || '').split(',').map(function (part) {
        return parseFloat(part) || 0;
      });
      if (!split.length) { return; }

      var labels = (bar.getAttribute('data-labels') || '').split(',');
      var max = split.reduce(function (a, b) { return Math.max(a, b); }, 0) || 1;

      var values = document.createElement('div');
      values.className = 'qvals';

      split.forEach(function (value, index) {
        var seg = document.createElement('span');
        seg.className = 'qseg';
        var inner = document.createElement('span');
        seg.appendChild(inner);
        bar.appendChild(seg);

        var caption = document.createElement('span');
        caption.textContent = labels[index] || nf(value);
        values.appendChild(caption);

        window.setTimeout(function () {
          inner.style.width = (value === 0 ? 0 : Math.max(8, (value / max) * 100)) + '%';
        }, 120 + index * 90);
      });

      bar.parentNode.insertBefore(values, bar.nextSibling);
    });
  }

  /* ── 12 · Budget donut ⇄ table ⇄ legend ───────────────────────────────── */
  function initBudget() {
    var group = $('donutSegs');
    if (!group) { return; }

    var radius = 72;
    var circumference = 2 * Math.PI * radius;
    var offset = 0;

    var segments = BUDGET.map(function (item, index) {
      var share = item.value / BUDGET_TOTAL;
      var length = share * circumference;
      var gap = length > 6 ? 2 : 0;

      var circle = svgEl('circle', {
        cx: 100, cy: 100, r: radius,
        class: 'donut-seg c' + index,
        transform: 'rotate(-90 100 100)',
        'stroke-dasharray': (length - gap) + ' ' + (circumference - length + gap),
        'stroke-dashoffset': -offset
      });
      group.appendChild(circle);
      offset += length;
      return circle;
    });

    var legend = $('donutLegend');
    var legendItems = [];
    if (legend) {
      BUDGET.forEach(function (item, index) {
        var row = document.createElement('div');
        row.className = 'dl';
        row.setAttribute('data-cat', index);
        row.innerHTML =
          '<span class="dl-sw c' + index + '"></span>' +
          '<span class="dl-n">' + item.name + '</span>' +
          '<span class="dl-v">' + money(item.value) + '</span>';
        legend.appendChild(row);
        legendItems.push(row);
      });
    }

    var tableRows = all('#budgetTable tbody tr[data-cat]');
    var label = $('dcLabel');
    var value = $('dcValue');
    var shareNode = $('dcShare');

    function setActive(index) {
      segments.forEach(function (circle, i) {
        circle.classList.toggle('is-dim', index !== null && i !== index);
        circle.classList.toggle('is-active', index === i);
      });
      legendItems.forEach(function (row, i) {
        row.classList.toggle('is-active', i === index);
      });
      tableRows.forEach(function (row) {
        row.classList.toggle('is-active', parseInt(row.getAttribute('data-cat'), 10) === index);
      });

      if (index === null) {
        if (label) { label.textContent = 'Вся смета'; }
        if (value) { value.textContent = money(BUDGET_TOTAL); }
        if (shareNode) { shareNode.textContent = '100 %'; }
        return;
      }

      var item = BUDGET[index];
      var share = (item.value / BUDGET_TOTAL) * 100;
      if (label) { label.textContent = item.name; }
      if (value) { value.textContent = money(item.value); }
      if (shareNode) {
        shareNode.textContent = nf(share, 1) + ' %' + (item.note ? ' · ' + item.note : '');
      }
    }

    function bind(node, index) {
      node.addEventListener('mouseenter', function () { setActive(index); });
      node.addEventListener('focus', function () { setActive(index); });
    }

    segments.forEach(function (circle, index) { bind(circle, index); });
    legendItems.forEach(function (row, index) { bind(row, index); });
    tableRows.forEach(function (row, index) {
      bind(row, parseInt(row.getAttribute('data-cat'), 10) || index);
    });

    var card = $('donut');
    if (card) {
      card.addEventListener('mouseleave', function () { setActive(null); });
    }
    if (legend) {
      legend.addEventListener('mouseleave', function () { setActive(null); });
    }
    var table = $('budgetTable');
    if (table) {
      table.addEventListener('mouseleave', function () { setActive(null); });
    }

    /* Reveal the ring by sweeping it on. */
    if (!reduceMotion) {
      segments.forEach(function (circle) {
        circle.style.opacity = '0';
      });
      window.setTimeout(function () {
        segments.forEach(function (circle, index) {
          window.setTimeout(function () { circle.style.opacity = ''; }, index * 90);
        });
      }, 200);
    }

    /* Coverage bar — the grant against the six-month budget. */
    var coverage = (GRANT_TENGE / BUDGET_TOTAL) * 100;
    var coverageNode = $('covValue');
    if (coverageNode) { coverageNode.textContent = nf(coverage) + ' %'; }
    var need = $('covNeed');
    if (need) { need.style.width = clamp((BUDGET_TOTAL / GRANT_TENGE) * 100, 0, 100) + '%'; }
    var have = $('covHave');
    if (have) { have.style.width = clamp(100 - (BUDGET_TOTAL / GRANT_TENGE) * 100, 0, 100) + '%'; }

    setActive(null);
  }

  /* ── Boot ─────────────────────────────────────────────────────────────── */
  function boot() {
    initReveal();
    initCounters();
    initDemo();
    initLive();
    initScroll();
    initCalc();
    var rows = initCompare();
    initMatrix(rows);
    initGrowth();
    initRoadmap();
    initGoals();
    initBudget();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
