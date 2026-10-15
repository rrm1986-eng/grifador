/* Grifador — marca-texto fixo para o Word (web, iPad e Windows). Versão 7. */
"use strict";

(function () {
  var VERSION = "7";
  // Cores padrão do marca-texto do Word.
  var COLORS = {
    "#FFFF00": "Amarelo",
    "#00FFFF": "Azul claro",
    "#00FF00": "Verde claro",
    "#FF0000": "Vermelho",
    "none": "Borracha"
  };
  var POLL_MS = 150;              // verificação contínua: intervalo entre conferidas
  var MIN_EVENT_MS = 40;          // mínimo para juntar a rajada de avisos de um arraste
  var ADJUST_WINDOW_MS = 6000;    // janela para "corrigir" o grifo quando a seleção é ajustada
  var OVERLAP = {
    Contains: 1, ContainsStart: 1, ContainsEnd: 1,
    Inside: 1, InsideStart: 1, InsideEnd: 1,
    OverlapsBefore: 1, OverlapsAfter: 1
  };

  var store = {
    get: function (k, d) {
      try { var v = localStorage.getItem("grifador." + k); return v === null ? d : JSON.parse(v); }
      catch (e) { return d; }
    },
    set: function (k, v) {
      try { localStorage.setItem("grifador." + k, JSON.stringify(v)); } catch (e) { /* sem armazenamento */ }
    }
  };

  var state = {
    auto: store.get("auto", true),
    color: store.get("color", "#FFFF00"),
    delay: store.get("delay7", 300),
    detect: store.get("detect", "auto"),
    collapse: store.get("collapse", true),
    adjust: store.get("adjust", true),
    api13: false,
    platform: "?",
    eventsRegistered: false,
    eventsSeen: 0,
    polls: 0,
    applied: 0,
    errors: 0,
    lastError: ""
  };
  if (!COLORS[state.color]) state.color = "#FFFF00";

  var ctx = null;          // contexto do Word de longa duração
  var temps = [];          // objetos rastreados só durante uma operação
  var lastApplied = null;  // { range, original, time } — último grifo (desfazer e correção)
  var candidate = null;    // verificação contínua: { range, since, handled }
  var queue = Promise.resolve();
  var eventTimer = null, pollTimer = null, flashTimer = null;

  var $ = function (id) { return document.getElementById(id); };

  // ---------- Infraestrutura ----------

  function enqueue(fn) {
    queue = queue.then(fn).then(releaseTemps, handleError);
    return queue;
  }
  function getCtx() { if (!ctx) ctx = new Word.RequestContext(); return ctx; }
  function track(c, o) { c.trackedObjects.add(o); return o; }
  function untrack(c, o) { if (o) { try { c.trackedObjects.remove(o); } catch (e) { /* ok */ } } }
  function releaseTemps() {
    if (ctx) for (var i = 0; i < temps.length; i++) untrack(ctx, temps[i]);
    temps = [];
  }

  function handleError(err) {
    temps = []; ctx = null; lastApplied = null; candidate = null;
    var msg = err && err.code ? err.code : "";
    if (err && err.message) msg += (msg ? ": " : "") + err.message;
    state.errors++;
    state.lastError = msg || String(err);
    console.error("[Grifador " + VERSION + "]", err);
    setStatus("Algo falhou. Detalhes em Ajustes e diagnóstico.", true);
    renderDiag();
  }

  function snapshot(v) {
    return (v === "" || v === undefined) ? { uniform: false, value: null } : { uniform: true, value: v || null };
  }

  // Prepara a leitura da seleção (ainda sem ir ao Word). No Word na web o objeto
  // precisa ser rastreado antes do primeiro sync, senão "morre" (InvalidObjectPath).
  function querySel(c) {
    var sel = c.document.getSelection();
    c.trackedObjects.add(sel);
    temps.push(sel);
    sel.load(state.api13 ? "isEmpty" : "text");
    sel.font.load("highlightColor");
    return sel;
  }
  function isEmpty(sel) {
    return state.api13 ? sel.isEmpty : (sel.text || "").replace(/[\s​ ]/g, "").length === 0;
  }
  // Se for corrigir, já pede a comparação com o último grifo na mesma ida ao Word.
  function queueCompare(sel, wantAdjust) {
    if (!lastApplied || !state.api13) return null;
    if (wantAdjust && Date.now() - lastApplied.time < ADJUST_WINDOW_MS) return sel.compareLocationWith(lastApplied.range);
    return null;
  }

  // Aplica a cor numa única ida ao Word.
  // opts.collapse: soltar a seleção e pôr o cursor no fim do grifo
  // opts.cmp: comparação com o último grifo (para corrigir quando a seleção foi ajustada)
  async function apply(c, sel, color, opts) {
    var original = snapshot(sel.font.highlightColor);
    var rel = opts.cmp ? opts.cmp.value : null;
    if (rel === "Equal" && lastApplied) {
      original = lastApplied.original;                       // mesmo trecho: só troca a cor
    } else if (rel && OVERLAP[rel] && lastApplied && lastApplied.original.uniform) {
      lastApplied.range.font.highlightColor = lastApplied.original.value;  // seleção ajustada: desfaz o anterior
    }
    sel.font.highlightColor = color === "none" ? null : color;
    var anchor = track(c, sel.getRange("Whole"));
    if (lastApplied) untrack(c, lastApplied.range);
    if (opts.collapse) sel.select("End");
    await c.sync();
    lastApplied = { range: anchor, original: original, time: Date.now() };
    state.applied++;
    flash(color);
    setStatus(color === "none" ? "Grifo removido." : "Grifado em " + COLORS[color].toLowerCase() + ".");
    renderDiag();
  }

  function autoOpts(cmp) { return { collapse: state.collapse, cmp: cmp }; }

  // ---------- Detecção por evento do Word ----------

  function onSelectionEvent() {
    state.eventsSeen++;
    if (state.detect === "auto" && pollTimer) stopPolling();   // o Word avisa: dispensa a verificação
    renderDiag();
    if (!state.auto || state.detect === "poll") return;
    clearTimeout(eventTimer);
    eventTimer = setTimeout(function () { enqueue(eventFire); }, Math.max(state.delay, MIN_EVENT_MS));
  }

  async function eventFire() {
    if (!state.auto) return;
    var c = getCtx();
    var sel = querySel(c);
    var cmp = queueCompare(sel, state.adjust);
    await c.sync();                      // 1ª ida: vazia? grifo atual? relação com o último grifo
    if (isEmpty(sel)) return;            // inclui o aviso causado pelo próprio "cursor no fim"
    await apply(c, sel, state.color, autoOpts(cmp));   // 2ª ida: grifa (e põe o cursor no fim)
  }

  // ---------- Detecção por verificação contínua ----------

  function dropCandidate(c) { if (candidate) untrack(c, candidate.range); candidate = null; }

  async function pollTick() {
    if (!state.auto) return;
    state.polls++;
    var c = getCtx();
    var sel = querySel(c);
    var same = candidate ? sel.compareLocationWith(candidate.range) : null;
    var cmp = queueCompare(sel, state.adjust);
    await c.sync();
    if (isEmpty(sel)) { dropCandidate(c); return; }
    if (same && same.value === "Equal") {
      if (!candidate.handled && Date.now() - candidate.since >= state.delay) {
        candidate.handled = true;
        await apply(c, sel, state.color, autoOpts(cmp));
      }
      return;
    }
    dropCandidate(c);
    candidate = { range: track(c, sel.getRange("Whole")), since: Date.now(), handled: false };
    await c.sync();
    if (state.polls % 10 === 0) renderDiag();
  }

  // Ao ligar: a seleção que já está na tela não deve ser grifada sozinha.
  async function primeCurrent() {
    var c = getCtx();
    var sel = querySel(c);
    await c.sync();
    dropCandidate(c);
    if (isEmpty(sel)) return;
    candidate = { range: track(c, sel.getRange("Whole")), since: Date.now(), handled: true };
    await c.sync();
  }

  function startPolling() {
    if (pollTimer) return;
    var loop = function () {
      pollTimer = setTimeout(function () {
        enqueue(pollTick).then(function () { if (pollTimer) loop(); });
      }, POLL_MS);
    };
    loop();
  }
  function stopPolling() { clearTimeout(pollTimer); pollTimer = null; }

  function applyMode() {
    if (!state.auto || !state.api13) { stopPolling(); renderDiag(); return; }
    if (state.detect === "poll") startPolling();
    else if (state.detect === "event") stopPolling();
    else if (state.eventsRegistered && state.eventsSeen > 0) stopPolling();
    else startPolling();
    renderDiag();
  }

  function usingNow() {
    if (!state.auto) return "nenhuma (modo desligado)";
    if (pollTimer) return "verificação contínua";
    return state.eventsRegistered ? "eventos do Word" : "nenhuma";
  }

  // ---------- Ações do painel ----------

  function onSwatch(color) {
    state.color = color;
    store.set("color", color);
    renderSwatches();
    renderAuto();
    enqueue(async function () {
      var c = getCtx();
      var sel = querySel(c);
      var cmp = queueCompare(sel, false);
      await c.sync();
      if (isEmpty(sel)) {
        setStatus(state.auto ? "Cor ativa: " + COLORS[color].toLowerCase() + "." : "Selecione um trecho e toque na cor.");
        return;
      }
      await apply(c, sel, color, { collapse: state.auto && state.collapse, cmp: cmp });
      if (candidate) candidate.handled = true;
    });
  }

  function onUndo() {
    enqueue(async function () {
      if (!lastApplied) { setStatus("Nada para desfazer."); return; }
      var c = getCtx();
      var ok = lastApplied.original.uniform;
      if (ok) lastApplied.range.font.highlightColor = lastApplied.original.value;
      untrack(c, lastApplied.range);
      await c.sync();
      lastApplied = null;
      setStatus(ok ? "Último grifo desfeito." : "Esse trecho tinha grifos de várias cores; use o Desfazer do Word.");
    });
  }

  function setAuto(on) {
    if (on && !state.api13) {
      setStatus("Este Word não oferece o modo automático. Use os botões de cor.", true);
      on = false;
    }
    state.auto = on;
    store.set("auto", on);
    clearTimeout(eventTimer);
    if (on) enqueue(primeCurrent);
    renderAuto();
    applyMode();
    setStatus(on ? "Selecione um trecho e solte: ele será grifado." : "Modo desligado. Selecione e toque na cor.");
  }

  // ---------- Interface ----------

  function setStatus(msg, isErr) {
    var el = $("status");
    el.textContent = msg;
    el.className = "status" + (isErr ? " err" : "");
  }

  function flash(color) {
    var el = document.querySelector('.sw[data-color="' + color + '"]');
    if (!el) return;
    el.classList.add("flash");
    clearTimeout(flashTimer);
    flashTimer = setTimeout(function () { el.classList.remove("flash"); }, 180);
  }

  function renderSwatches() {
    var btns = document.querySelectorAll(".sw");
    for (var i = 0; i < btns.length; i++) {
      var on = btns[i].getAttribute("data-color") === state.color;
      btns[i].classList.toggle("is-active", on);
      btns[i].setAttribute("aria-pressed", on ? "true" : "false");
    }
  }

  function renderAuto() {
    $("autoSwitch").setAttribute("aria-checked", state.auto ? "true" : "false");
    $("autoRow").classList.toggle("is-on", state.auto);
    $("autoSub").textContent = state.auto
      ? (state.color === "none" ? "Ligado — apagando grifos" : "Ligado — " + COLORS[state.color].toLowerCase())
      : "Desligado";
  }

  function renderDiag() {
    $("diag").textContent =
      "Versão: " + VERSION + " | Plataforma: " + state.platform + "\n" +
      "Word API 1.3: " + (state.api13 ? "sim" : "não") + "\n" +
      "Evento de seleção: " + (state.eventsRegistered ? "registrado" : "não disponível") + "\n" +
      "Eventos recebidos: " + state.eventsSeen + "\n" +
      "Detecção em uso: " + usingNow() + "\n" +
      "Verificações: " + state.polls + " | Grifos: " + state.applied + " | Erros: " + state.errors +
      (state.lastError ? "\nÚltimo erro: " + state.lastError : "");
  }

  function bindUi() {
    var btns = document.querySelectorAll(".sw");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function (ev) { onSwatch(ev.currentTarget.getAttribute("data-color")); });
    }
    $("autoSwitch").addEventListener("click", function () { setAuto(!state.auto); });
    $("undoBtn").addEventListener("click", onUndo);

    $("delay").value = state.delay;
    $("delayOut").textContent = state.delay;
    $("delay").addEventListener("input", function (e) {
      state.delay = Math.round(Number(e.target.value));
      $("delayOut").textContent = state.delay;
      store.set("delay7", state.delay);
    });

    $("detect").value = state.detect;
    $("detect").addEventListener("change", function (e) {
      state.detect = e.target.value;
      store.set("detect", state.detect);
      applyMode();
    });

    $("collapse").checked = state.collapse;
    $("collapse").addEventListener("change", function (e) {
      state.collapse = e.target.checked;
      store.set("collapse", state.collapse);
    });

    $("adjust").checked = state.adjust;
    $("adjust").addEventListener("change", function (e) {
      state.adjust = e.target.checked;
      store.set("adjust", state.adjust);
    });
  }

  // ---------- Início ----------

  Office.onReady(function (info) {
    bindUi();
    renderSwatches();

    if (info.host !== Office.HostType.Word) {
      setStatus("Abra este suplemento dentro do Word.", true);
      return;
    }
    try { state.platform = String(Office.context.diagnostics.platform); } catch (e) { /* ok */ }
    state.api13 = Office.context.requirements.isSetSupported("WordApi", "1.3");
    if (!state.api13) state.auto = false;
    renderAuto();
    if (state.auto) enqueue(primeCurrent);
    setStatus(state.auto ? "Modo ligado. Selecione um trecho e solte." : "Selecione um trecho e toque na cor.");

    var doc = Office.context.document;
    if (doc && doc.addHandlerAsync) {
      doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionEvent, function (res) {
        state.eventsRegistered = res.status === Office.AsyncResultStatus.Succeeded;
        if (!state.eventsRegistered && res.error) state.lastError = "evento: " + res.error.message;
        applyMode();
      });
    } else {
      applyMode();
    }
  });
})();
