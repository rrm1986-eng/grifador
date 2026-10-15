/* Grifador — marca-texto fixo para o Word (web, iPad e Windows). Versão 5. */
"use strict";

(function () {
  var VERSION = "5";
  // Cores padrão do marca-texto do Word.
  var COLORS = { "#FFFF00": 1, "#00FFFF": 1, "#00FF00": 1, "#FF0000": 1, "none": 1 };
  // Quando o Word avisa que a seleção mudou, espera só este tanto para ter certeza
  // de que você terminou de arrastar (os avisos chegam em rajada durante o arraste).
  var EVENT_SETTLE_MS = 90;
  // Plano B, só se o Word não avisar a seleção: confere a seleção neste intervalo.
  var POLL_MS = 150;

  var store = {
    get: function (k, d) {
      try { var v = localStorage.getItem("grifador." + k); return v === null ? d : JSON.parse(v); }
      catch (e) { return d; }
    },
    set: function (k, v) {
      try { localStorage.setItem("grifador." + k, JSON.stringify(v)); } catch (e) { /* sem armazenamento */ }
    }
  };

  var auto = store.get("auto", true);
  var color = store.get("color", "#FFFF00");
  if (!COLORS[color]) color = "#FFFF00";

  var api13 = false;
  var ctx = null;          // contexto do Word de longa duração
  var temps = [];          // objetos rastreados só durante uma operação
  var lastApplied = null;  // { range, original } — para o botão desfazer
  var candidate = null;    // plano B: seleção observada { range, handled }
  var eventsSeen = 0;
  var queue = Promise.resolve();
  var eventTimer = null, pollTimer = null, flashTimer = null;

  var $ = function (id) { return document.getElementById(id); };

  // ---------- Infraestrutura ----------

  function enqueue(fn) {
    queue = queue.then(fn).then(releaseTemps, onError);
    return queue;
  }
  function getCtx() { if (!ctx) ctx = new Word.RequestContext(); return ctx; }
  function track(c, o) { c.trackedObjects.add(o); return o; }
  function untrack(c, o) { if (o) { try { c.trackedObjects.remove(o); } catch (e) { /* ok */ } } }
  function releaseTemps() {
    if (ctx) for (var i = 0; i < temps.length; i++) untrack(ctx, temps[i]);
    temps = [];
  }

  function onError(err) {
    temps = []; ctx = null; lastApplied = null; candidate = null;
    var msg = (err && (err.code || err.message)) || String(err);
    console.error("[Grifador " + VERSION + "]", err);
    $("bar").classList.add("err");
    $("bar").title = "Erro: " + msg;
  }
  function clearErr() { $("bar").classList.remove("err"); $("bar").title = ""; }

  function snapshot(v) {
    return (v === "" || v === undefined) ? { uniform: false, value: null } : { uniform: true, value: v || null };
  }

  // Prepara a leitura da seleção (sem ir ao Word ainda). No Word na web o objeto
  // precisa ser rastreado antes do primeiro sync, senão "morre" (InvalidObjectPath).
  function querySel(c) {
    var sel = c.document.getSelection();
    c.trackedObjects.add(sel);
    temps.push(sel);
    sel.load(api13 ? "isEmpty" : "text");
    sel.font.load("highlightColor");
    return sel;
  }
  function isEmpty(sel) {
    return api13 ? sel.isEmpty : (sel.text || "").replace(/[\s​ ]/g, "").length === 0;
  }

  // Aplica a cor numa única ida ao Word. collapse = soltar a seleção e pôr o cursor no fim.
  async function apply(c, sel, value, collapse) {
    var original = snapshot(sel.font.highlightColor);
    sel.font.highlightColor = value === "none" ? null : value;
    var anchor = track(c, sel.getRange("Whole"));
    if (lastApplied) untrack(c, lastApplied.range);
    if (collapse) sel.select("End");
    await c.sync();
    lastApplied = { range: anchor, original: original };
    clearErr();
    flash(value);
  }

  // ---------- Automático: pelo aviso do Word (caminho rápido) ----------

  function onSelectionEvent() {
    eventsSeen++;
    stopPolling();                       // o Word avisa: o plano B não é mais necessário
    if (!auto) return;
    clearTimeout(eventTimer);
    eventTimer = setTimeout(function () { enqueue(autoFire); }, EVENT_SETTLE_MS);
  }

  async function autoFire() {
    if (!auto) return;
    var c = getCtx();
    var sel = querySel(c);
    await c.sync();                      // 1ª ida: a seleção está vazia? qual o grifo atual?
    if (isEmpty(sel)) return;            // inclui o aviso causado pelo próprio "cursor no fim"
    await apply(c, sel, color, true);    // 2ª ida: grifa e põe o cursor no fim
  }

  // ---------- Automático: plano B (se o Word não avisar a seleção) ----------

  function dropCandidate(c) { if (candidate) untrack(c, candidate.range); candidate = null; }

  async function pollTick() {
    if (!auto) return;
    var c = getCtx();
    var sel = querySel(c);
    var cmp = candidate ? sel.compareLocationWith(candidate.range) : null;
    await c.sync();
    if (isEmpty(sel)) { dropCandidate(c); return; }
    if (cmp && cmp.value === "Equal") {  // ficou parada desde a última conferida: você soltou
      if (!candidate.handled) {
        candidate.handled = true;
        await apply(c, sel, color, true);
      }
      return;
    }
    dropCandidate(c);
    candidate = { range: track(c, sel.getRange("Whole")), handled: false };
    await c.sync();
  }

  // Ao ligar: a seleção que já está na tela não deve ser grifada sozinha.
  async function primeCurrent() {
    var c = getCtx();
    var sel = querySel(c);
    await c.sync();
    dropCandidate(c);
    if (isEmpty(sel)) return;
    candidate = { range: track(c, sel.getRange("Whole")), handled: true };
    await c.sync();
  }

  function startPolling() {
    if (pollTimer || eventsSeen > 0 || !auto || !api13) return;
    var loop = function () {
      pollTimer = setTimeout(function () {
        enqueue(pollTick).then(function () { if (pollTimer) loop(); });
      }, POLL_MS);
    };
    loop();
  }
  function stopPolling() { clearTimeout(pollTimer); pollTimer = null; }

  // ---------- Botões ----------

  function onSwatch(c0) {
    color = c0;
    store.set("color", color);
    renderSwatches();
    enqueue(async function () {
      var c = getCtx();
      var sel = querySel(c);
      await c.sync();
      if (isEmpty(sel)) return;
      await apply(c, sel, color, auto);
      if (candidate) candidate.handled = true;
    });
  }

  function onUndo() {
    enqueue(async function () {
      if (!lastApplied) return;
      var c = getCtx();
      if (lastApplied.original.uniform) lastApplied.range.font.highlightColor = lastApplied.original.value;
      untrack(c, lastApplied.range);
      await c.sync();
      lastApplied = null;
    });
  }

  function setAuto(on) {
    if (on && !api13) on = false;
    auto = on;
    store.set("auto", on);
    renderAuto();
    clearTimeout(eventTimer);
    if (on) { enqueue(primeCurrent); startPolling(); } else { stopPolling(); }
  }

  // ---------- Interface ----------

  function flash(value) {
    var el = document.querySelector('.sw[data-color="' + value + '"]');
    if (!el) return;
    el.classList.add("flash");
    clearTimeout(flashTimer);
    flashTimer = setTimeout(function () { el.classList.remove("flash"); }, 180);
  }
  function renderSwatches() {
    var btns = document.querySelectorAll(".sw");
    for (var i = 0; i < btns.length; i++) {
      var on = btns[i].getAttribute("data-color") === color;
      btns[i].classList.toggle("is-active", on);
      btns[i].setAttribute("aria-pressed", on ? "true" : "false");
    }
  }
  function renderAuto() {
    $("autoSwitch").setAttribute("aria-checked", auto ? "true" : "false");
    $("autoSwitch").title = "Modo marca-texto: " + (auto ? "ligado" : "desligado");
  }

  // ---------- Início ----------

  Office.onReady(function (info) {
    var btns = document.querySelectorAll(".sw");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function (ev) { onSwatch(ev.currentTarget.getAttribute("data-color")); });
    }
    $("autoSwitch").addEventListener("click", function () { setAuto(!auto); });
    $("undoBtn").addEventListener("click", onUndo);
    renderSwatches();

    if (info.host !== Office.HostType.Word) { onError(new Error("abra o Grifador dentro do Word")); return; }
    api13 = Office.context.requirements.isSetSupported("WordApi", "1.3");
    if (!api13) auto = false;
    renderAuto();
    if (auto) enqueue(primeCurrent);

    var doc = Office.context.document;
    if (doc && doc.addHandlerAsync) {
      doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionEvent, function () { startPolling(); });
    } else {
      startPolling();
    }
  });
})();
