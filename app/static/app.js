document.querySelectorAll('[data-abrir]').forEach(function (b) {
  b.addEventListener('click', function () { var d = document.getElementById(b.dataset.abrir); if (d.showModal) d.showModal(); else d.setAttribute('open', ''); });
});
document.querySelectorAll('[data-cerrar]').forEach(function (b) {
  b.addEventListener('click', function () { b.closest('dialog').close(); });
});
document.querySelectorAll('form').forEach(function (f) {
  f.addEventListener('submit', function () {
    var b = f.querySelector('[data-espera]');
    if (b) setTimeout(function () { b.disabled = true; b.textContent = b.dataset.espera; }, 0);
  });
});

/* Pestanas accesibles (sin JS se ven ambas secciones) */
document.querySelectorAll('.revision').forEach(function (cont) {
  var tabs = Array.prototype.slice.call(cont.querySelectorAll('[role=tab]'));
  function activar(tab, enfocar) {
    tabs.forEach(function (t) {
      var on = t === tab, p = document.getElementById(t.getAttribute('aria-controls'));
      t.setAttribute('aria-selected', on ? 'true' : 'false'); t.tabIndex = on ? 0 : -1; p.hidden = !on;
    });
    if (enfocar) tab.focus();
  }
  tabs.forEach(function (t, i) {
    t.addEventListener('click', function () { activar(t, false); });
    t.addEventListener('keydown', function (e) {
      var n = null;
      if (e.key === 'ArrowRight') n = tabs[(i + 1) % tabs.length];
      else if (e.key === 'ArrowLeft') n = tabs[(i - 1 + tabs.length) % tabs.length];
      else if (e.key === 'Home') n = tabs[0]; else if (e.key === 'End') n = tabs[tabs.length - 1];
      if (n) { e.preventDefault(); activar(n, true); }
    });
  });
  var ini = tabs.filter(function (t) { return t.id === 'tab-' + cont.dataset.tabInicial; })[0] || tabs[0];
  activar(ini, false);
});

/* Copiar filas TSV al portapapeles (con alternativa: cuadro de texto seleccionado) */
document.querySelectorAll('[data-copiar]').forEach(function (b) {
  b.addEventListener('click', function () {
    var texto = JSON.parse(document.getElementById(b.dataset.copiar).textContent);
    var est = document.getElementById(b.dataset.estado), ta = document.getElementById('copia-manual');
    function listo() { ta.hidden = true; est.textContent = 'Copiado. Péguelo en la hoja TRINIDAD del Excel.'; }
    function manual() {
      ta.hidden = false; ta.value = texto; ta.focus(); ta.select();
      var ok = false; try { ok = document.execCommand('copy'); } catch (e) {}
      if (ok) listo(); else est.textContent = 'Seleccione el texto de abajo y cópielo (Ctrl+C).';
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(texto).then(listo, manual); else manual();
  });
});
