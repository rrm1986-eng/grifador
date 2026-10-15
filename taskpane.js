/* Grifador — suplemento de realce para o Word (iPad, web e Windows). */
"use strict";

(function () {
  // Cores padrão do marca-texto do Word (valores exatos da paleta do Word).
  var COLORS = {
    "#FFFF00": "Amarelo",
    "#00FFFF": "Azul claro",
    "#00FF00": "Verde claro",
    "#FF0000": "Vermelho",
    "none": "Borracha"
  };
  var MAX_CHARS = 20000;          // seleções maiores que isso não são grifadas no automático
  var ADJUST_WINDOW_MS = 6000;    // janela para "corrigir" o grifo quando a seleção é ajustada
  var POLL_MS = 300;              // intervalo da verificação contínua
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
    auto: store.get("auto", false),
    color: store.get("color", "#FFFF00"),
    delay: store.get("delay", 600),
    detect: store.get("detect", "auto"),
    adjust: store.get("adjust", true),
    api13: false,
    platform: "?",
    eventsRegistered: false,
    eventsSeen: 0,
    polls: 0,
    applied: 0,
    lastError: ""
  };
  if (!COLORS[state.color]) state.color = "#FFFF00";

  var ctx = null;          // contexto do Word de longa duração
  var candidate = null;    // { range, since, handled } — seleção sendo observada
  var lastApplied = null;  // { range, original, color, time } — último grifo feito
  var lastHandled = null;  // seleção que já foi tratada (evita grifar duas vezes)
  var queue = Promise.resolve();
  var eventTimer = null;
  var pollTimer = null;

  var $ = function (id) { return document.getElementById(id); };

  // ---------- Infraestrutura ----------

  function enqueue(fn) {
    queue = queue.then(fn).catch(handleError);
    return queue;
  }

  function getCtx() {
    if (!ctx) ctx = new Word.RequestContext();
    return ctx;
  }

  function track(c, obj) { c.trackedObjects.add(obj); return obj; }

  function untrack(c, obj) {
    if (!obj) return;
    try { c.trackedObjects.remove(obj); } catch (e) { /* já liberado */ }
  }

  function resetContext() {
    ctx = null; candidate = null; lastApplied = null; lastHandled = null;
  }

  function handleError(err) {
    var msg = (err && (err.message || err.code)) ? (err.code ? err.code + ": " : "") + (err.message || "") : String(err);
    state.lastError = msg;
    console.error("[Grifador]", err);
    setStatus("Algo falhou: " + msg, true);
    resetContext();
    renderDiag();
  }

  async function readSelection(c) {
    var sel = c.document.getSelection();
    sel.load("text");
    await c.sync();
    var text = sel.text || "";
    var empty = text.replace(/[\s​ ]/g, "").length === 0;
    return { sel: sel, text: text, empty: empty };
  }

  async function relation(c, a, b) {
    var r = a.compareLocationWith(b);
    await c.sync();
    return r.value;
  }

  function snapshot(v) {
    // null = sem grifo; "" ou undefined = grifo misto (não dá para restaurar exatamente)
    if (v === "" || v === undefined) return { uniform: false, value: null };
    return { uniform: true, value: v || null };
  }

  // Devolve o trecho ao estado anterior. Se antes havia grifos de cores misturadas,
  // não mexe (para não apagar grifos antigos) e devolve false.
  function revert(L) {
    if (!L.original.uniform) return false;
    L.range.font.highlightColor = L.original.value;
    return true;
  }

  // ---------- Núcleo: aplicar o grifo ----------

  async function applyColor(c, sel, color, allowAdjust) {
    var value = color === "none" ? null : color;

    if (!state.api13) {
      sel.font.highlightColor = value;
      await c.sync();
      state.applied++;
      afterApply(color);
      return;
    }

    var keepOriginal = null;
    if (lastApplied) {
      var rel = await relation(c, sel, lastApplied.range);
      if (rel === "Equal") {
        keepOriginal = lastApplied.original;          // só trocou a cor do mesmo trecho
      } else if (allowAdjust && state.adjust &&
                 Date.now() - lastApplied.time < ADJUST_WINDOW_MS && OVERLAP[rel]) {
        revert(lastApplied);                           // seleção ajustada: desfaz o grifo anterior
      }
    }

    var original = keepOriginal;
    if (!original) {
      sel.font.load("highlightColor");
      await c.sync();
      original = snapshot(sel.font.highlightColor);
    }

    sel.font.highlightColor = value;
    var anchor = track(c, sel.getRange("Whole"));
    var handled = track(c, sel.getRange("Whole"));
    if (lastApplied) untrack(c, lastApplied.range);
    if (lastHandled) untrack(c, lastHandled);
    await c.sync();

    lastApplied = { range: anchor, original: original, color: color, time: Date.now() };
    lastHandled = handled;
    state.applied++;
    afterApply(color);
  }

  function afterApply(color) {
    setStatus(color === "none" ? "Grifo removido." : "Grifado em " + COLORS[color].toLowerCase() + ".");
    renderDiag();
  }

  async function handleStable(c, sel, text) {
    if (!state.auto) return;
    if (lastHandled) {
      var rel = await relation(c, sel, lastHandled);
      if (rel === "Equal") return;
    }
    if (text.length > MAX_CHARS) {
      setStatus("Seleção grande demais para grifar sozinho. Use os botões de cor.");
      if (lastHandled) untrack(c, lastHandled);
      lastHandled = track(c, sel.getRange("Whole"));
      await c.sync();
      return;
    }
    await applyColor(c, sel, state.color, true);
  }

  // ---------- Detecção por evento ----------

  function onSelectionEvent() {
    state.eventsSeen++;
    if (state.detect === "auto" && pollTimer) stopPolling();   // eventos funcionam: dispensa a verificação
    renderDiag();
    if (!state.auto || state.detect === "poll") return;
    clearTimeout(eventTimer);
    eventTimer = setTimeout(function () { enqueue(eventFire); }, state.delay);
  }

  async function eventFire() {
    if (!state.auto) return;
    var c = getCtx();
    var r = await readSelection(c);
    if (r.empty) { forgetHandled(c); return; }
    await handleStable(c, r.sel, r.text);
  }

  // Quando a seleção some (só o cursor), selecionar o mesmo trecho de novo volta a grifar.
  function forgetHandled(c) {
    if (lastHandled) untrack(c, lastHandled);
    lastHandled = null;
  }

  // ---------- Detecção por verificação contínua (reserva) ----------

  function dropCandidate(c) {
    if (candidate) untrack(c, candidate.range);
    candidate = null;
  }

  async function pollTick() {
    if (!state.auto) return;
    state.polls++;
    var c = getCtx();
    var r = await readSelection(c);
    if (r.empty) { dropCandidate(c); forgetHandled(c); return; }

    if (candidate) {
      var rel = await relation(c, r.sel, candidate.range);
      if (rel === "Equal") {
        if (!candidate.handled && Date.now() - candidate.since >= state.delay) {
          candidate.handled = true;
          await handleStable(c, r.sel, r.text);
        }
        return;
      }
      dropCandidate(c);
    }
    candidate = { range: track(c, r.sel.getRange("Whole")), since: Date.now(), handled: false };
    await c.sync();
    if (state.polls % 10 === 0) renderDiag();
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

  function stopPolling() {
    clearTimeout(pollTimer);
    pollTimer = null;
  }

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
      var r = await readSelection(c);
      if (r.empty) {
        setStatus(state.auto
          ? "Cor ativa: " + COLORS[color].toLowerCase() + "."
          : "Selecione um trecho e toque na cor.");
        return;
      }
      await applyColor(c, r.sel, color, false);
      if (candidate) candidate.handled = true;
    });
  }

  function onUndo() {
    enqueue(async function () {
      if (!lastApplied) { setStatus("Nada para desfazer."); return; }
      var c = getCtx();
      var ok = revert(lastApplied);
      untrack(c, lastApplied.range);
      await c.sync();
      lastApplied = null;
      setStatus(ok ? "Último grifo desfeito."
                   : "Esse trecho tinha grifos de várias cores; use o Desfazer do próprio Word.");
    });
  }

  function setAuto(on) {
    if (on && !state.api13) {
      setStatus("Este Word não oferece os recursos do modo automático. Use os botões de cor.", true);
      on = false;
    }
    state.auto = on;
    store.set("auto", on);
    if (on) {
      // ao ligar, a seleção que já estava na tela não deve ser grifada sozinha
      enqueue(async function () {
        var c = getCtx();
        var r = await readSelection(c);
        if (!r.empty) {
          if (lastHandled) untrack(c, lastHandled);
          lastHandled = track(c, r.sel.getRange("Whole"));
          await c.sync();
        }
      });
    }
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
      "Plataforma: " + state.platform + "\n" +
      "Word API 1.3: " + (state.api13 ? "sim" : "não") + "\n" +
      "Evento de seleção: " + (state.eventsRegistered ? "registrado" : "não disponível") + "\n" +
      "Eventos recebidos: " + state.eventsSeen + "\n" +
      "Detecção em uso: " + usingNow() + "\n" +
      "Verificações: " + state.polls + " | Grifos: " + state.applied +
      (state.lastError ? "\nÚltimo erro: " + state.lastError : "");
  }

  function bindUi() {
    var btns = document.querySelectorAll(".sw");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function (ev) {
        onSwatch(ev.currentTarget.getAttribute("data-color"));
      });
    }
    $("autoSwitch").addEventListener("click", function () { setAuto(!state.auto); });
    $("undoBtn").addEventListener("click", onUndo);

    $("delay").value = state.delay;
    $("delayOut").textContent = state.delay;
    $("delay").addEventListener("input", function (e) {
      state.delay = Math.round(Number(e.target.value));
      $("delayOut").textContent = state.delay;
      store.set("delay", state.delay);
    });

    $("detect").value = state.detect;
    $("detect").addEventListener("change", function (e) {
      state.detect = e.target.value;
      store.set("detect", state.detect);
      applyMode();
    });

    $("adjust").checked = state.adjust;
    $("adjust").addEventListener("change", function (e) {
      state.adjust = e.target.checked;
      store.set("adjust", state.adjust);
    });
  }

  // ---------- Inicialização ----------

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
