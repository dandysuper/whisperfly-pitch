/* WhisperFly — pitch deck navigation.
   Deliberately a classic script, not an ES module: `import` is blocked over file://,
   and this deck has to work when someone just double-clicks index.html. */
(function () {
  'use strict';

  var stage = document.getElementById('stage');
  if (!stage) return;

  var slides = Array.prototype.slice.call(stage.querySelectorAll('.slide'));
  var total = slides.length;
  if (!total) return;

  var prevBtn  = document.getElementById('prev');
  var nextBtn  = document.getElementById('next');
  var printBtn = document.getElementById('print');
  var counter  = document.getElementById('counter');
  var dotsWrap = document.getElementById('dots');
  var progress = document.getElementById('progress');

  var current = 0;

  function pad(n) {
    return (n < 10 ? '0' : '') + n;
  }

  /* Build one dot per slide. */
  var dots = slides.map(function (_, i) {
    var dot = document.createElement('button');
    dot.type = 'button';
    dot.className = 'dot';
    dot.setAttribute('aria-label', 'Слайд ' + (i + 1));
    dot.addEventListener('click', function () { go(i); });
    if (dotsWrap) dotsWrap.appendChild(dot);
    return dot;
  });

  function render() {
    slides.forEach(function (slide, i) {
      slide.classList.toggle('active', i === current);
    });

    dots.forEach(function (dot, i) {
      dot.setAttribute('aria-current', i === current ? 'true' : 'false');
    });

    if (counter) counter.textContent = pad(current + 1) + ' / ' + pad(total);
    if (prevBtn) prevBtn.disabled = current === 0;
    if (nextBtn) nextBtn.disabled = current === total - 1;
    if (progress) progress.style.width = ((current + 1) / total * 100) + '%';

    writeHash(current + 1);
  }

  function go(index) {
    current = Math.max(0, Math.min(total - 1, index));
    render();
  }

  function next() { go(current + 1); }
  function prev() { go(current - 1); }

  /* Hash sync: deep-linkable slides that survive a reload or a shared URL.
     replaceState does not fire hashchange, so this cannot loop.
     Some browsers refuse history writes on a file:// document, so the call is
     guarded: opening index.html by double-click still navigates, it just does
     not rewrite the address bar. */
  function writeHash(n) {
    if (!history.replaceState) return;
    try {
      history.replaceState(null, '', '#' + n);
    } catch (error) {
      /* file:// origin — navigation works, the address bar simply stays put. */
    }
  }

  function readHash() {
    var n = parseInt((window.location.hash || '').replace('#', ''), 10);
    return (n >= 1 && n <= total) ? n - 1 : 0;
  }

  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      if (document.documentElement.requestFullscreen) {
        document.documentElement.requestFullscreen();
      }
    } else if (document.exitFullscreen) {
      document.exitFullscreen();
    }
  }

  if (prevBtn) prevBtn.addEventListener('click', prev);
  if (nextBtn) nextBtn.addEventListener('click', next);
  if (printBtn) printBtn.addEventListener('click', function () { window.print(); });

  document.addEventListener('keydown', function (event) {
    if (event.metaKey || event.ctrlKey || event.altKey) return;

    switch (event.key) {
      case 'ArrowRight':
      case 'PageDown':
      case ' ':
      case 'Enter':
        event.preventDefault();
        next();
        break;
      case 'ArrowLeft':
      case 'PageUp':
        event.preventDefault();
        prev();
        break;
      case 'Home':
        event.preventDefault();
        go(0);
        break;
      case 'End':
        event.preventDefault();
        go(total - 1);
        break;
      case 'f':
      case 'F':
        if (!event.target.closest('input, textarea')) toggleFullscreen();
        break;
      default:
        break;
    }
  });

  /* Swipe on trackpads and touch screens. */
  var touchStartX = 0;
  var touchStartY = 0;

  document.addEventListener('touchstart', function (event) {
    touchStartX = event.changedTouches[0].clientX;
    touchStartY = event.changedTouches[0].clientY;
  }, { passive: true });

  document.addEventListener('touchend', function (event) {
    var dx = event.changedTouches[0].clientX - touchStartX;
    var dy = event.changedTouches[0].clientY - touchStartY;
    if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy)) {
      if (dx < 0) { next(); } else { prev(); }
    }
  }, { passive: true });

  window.addEventListener('hashchange', function () { go(readHash()); });

  /* Print only the current slide unless the user asked for the whole deck.
     Before printing we reveal every slide so the print stylesheet can lay them out. */
  window.addEventListener('beforeprint', function () {
    slides.forEach(function (slide) { slide.classList.add('active'); });
  });
  window.addEventListener('afterprint', function () {
    slides.forEach(function (slide, i) { slide.classList.toggle('active', i === current); });
  });

  go(readHash());
})();
