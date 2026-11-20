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
