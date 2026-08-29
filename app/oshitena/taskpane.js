// オシテナ タスクペイン
/* global Office, Word, OshitenaOoxml */

(function () {
  "use strict";

  var STORAGE_KEY = "oshitena-settings-v2";

  var state = {
    sizeMm: 15,
    caption: "認印",
    color: "7F7F7F",
    dash: "sysDot",
    strokePt: 0.75,
    offsetXmm: 0,
    offsetYmm: 0,
    count: 1,
  };

  function $(id) {
    return document.getElementById(id);
  }

  function loadSettings() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var saved = JSON.parse(raw);
        Object.keys(state).forEach(function (k) {
          if (typeof saved[k] === typeof state[k]) state[k] = saved[k];
        });
      }
    } catch (e) {
      /* localStorageが使えない環境では既定値で動く */
    }
  }

  function saveSettings() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      /* 保存できなくても支障なし */
    }
  }

  function syncChips(containerId, dataAttr, value) {
    var chips = $(containerId).querySelectorAll("[data-" + dataAttr + "]");
    chips.forEach(function (chip) {
      chip.setAttribute(
        "aria-pressed",
        chip.getAttribute("data-" + dataAttr) === String(value) ? "true" : "false"
      );
    });
  }

  function syncPresetChips() {
    var chips = $("preset-chips").querySelectorAll("[data-preset]");
    chips.forEach(function (chip) {
      var match =
        Number(chip.getAttribute("data-size")) === state.sizeMm &&
        chip.getAttribute("data-caption") === state.caption;
      chip.setAttribute("aria-pressed", match ? "true" : "false");
    });
  }

  function renderPreview() {
    var canvas = $("preview");
    var ctx = canvas.getContext("2d");
    var scale = 3.2; // 1mm → 3.2px
    var d = state.sizeMm * scale;
    var pad = 8;
    var captionSpace = state.caption ? 18 : 0;
    canvas.width = Math.max(90, Math.ceil(d + pad * 2));
    canvas.height = Math.ceil(d + pad * 2 + captionSpace);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    var cx = canvas.width / 2;
    var cy = pad + d / 2;
    var r = d / 2;
    ctx.strokeStyle = "#" + state.color;
    ctx.lineWidth = Math.max(1, state.strokePt * scale * 0.353);
    if (state.dash === "dash") ctx.setLineDash([6, 4]);
    else if (state.dash === "sysDot") ctx.setLineDash([2, 3]);
    else ctx.setLineDash([]);
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.stroke();
    if (state.caption) {
      ctx.setLineDash([]);
      ctx.fillStyle = "#" + state.color;
      var fontPx = OshitenaOoxml.CAPTION_FONT_PT * scale * 0.353;
      ctx.font = fontPx + 'px "Yu Mincho", "游明朝", serif';
      ctx.textAlign = "center";
      ctx.textBaseline = "top";
      ctx.fillText(
        state.caption,
        cx,
        cy + r + OshitenaOoxml.CAPTION_GAP_MM * scale + 2
      );
    }
  }

  function refreshUi() {
    syncPresetChips();
    syncChips("caption-chips", "caption", state.caption);
    syncChips("color-chips", "color", state.color);
    $("size-custom").value = state.sizeMm;
    $("dash-style").value = state.dash;
    $("stroke-weight").value = String(state.strokePt);
    $("offset-x").value = state.offsetXmm;
    $("offset-y").value = state.offsetYmm;
    $("count").value = state.count;
    renderPreview();
  }

  function setStatus(message, isError) {
    var el = $("status");
    el.textContent = message || "";
    el.className = isError ? "status err" : "status";
  }

  function clampNumber(value, min, max, fallback) {
    var n = Number(value);
    if (!isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
  }

  async function insertStamp() {
    var btn = $("insert-btn");
    btn.disabled = true;
    setStatus("挿入中…");
    try {
      var ooxml = OshitenaOoxml.buildStampOoxml(
        {
          diameterMm: state.sizeMm,
          color: state.color,
          dash: state.dash,
          strokePt: state.strokePt,
          caption: state.caption,
          offsetXmm: state.offsetXmm,
          offsetYmm: state.offsetYmm,
        },
        state.count,
        4
      );
      await Word.run(async function (context) {
        var selection = context.document.getSelection();
        selection.insertOoxml(ooxml, Word.InsertLocation.end);
        await context.sync();
      });
      setStatus(
        state.count > 1
          ? state.count + "個の押印欄を挿入しました"
          : "押印欄を挿入しました"
      );
      saveSettings();
    } catch (error) {
      var detail =
        error && error.debugInfo ? " (" + error.debugInfo.errorLocation + ")" : "";
      setStatus("挿入できませんでした：" + (error.message || error) + detail, true);
    } finally {
      btn.disabled = false;
    }
  }

  function wireEvents() {
    $("preset-chips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-preset]");
      if (!chip) return;
      state.sizeMm = Number(chip.getAttribute("data-size"));
      state.caption = chip.getAttribute("data-caption");
      $("caption-custom").value = "";
      refreshUi();
      saveSettings();
    });

    $("size-custom").addEventListener("change", function () {
      state.sizeMm = clampNumber(this.value, 6, 60, 15);
      refreshUi();
      saveSettings();
    });

    $("caption-chips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-caption]");
      if (!chip) return;
      state.caption = chip.getAttribute("data-caption");
      $("caption-custom").value = "";
      refreshUi();
      saveSettings();
    });

    $("caption-custom").addEventListener("input", function () {
      state.caption = this.value.trim();
      syncChips("caption-chips", "caption", state.caption);
      syncPresetChips();
      renderPreview();
      saveSettings();
    });

    $("color-chips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-color]");
      if (!chip) return;
      state.color = chip.getAttribute("data-color");
      refreshUi();
      saveSettings();
    });

    $("dash-style").addEventListener("change", function () {
      state.dash = this.value;
      renderPreview();
      saveSettings();
    });

    $("stroke-weight").addEventListener("change", function () {
      state.strokePt = Number(this.value);
      renderPreview();
      saveSettings();
    });

    $("offset-x").addEventListener("change", function () {
      state.offsetXmm = clampNumber(this.value, -100, 100, 0);
      saveSettings();
    });

    $("offset-y").addEventListener("change", function () {
      state.offsetYmm = clampNumber(this.value, -100, 100, 0);
      saveSettings();
    });

    $("count").addEventListener("change", function () {
      state.count = Math.round(clampNumber(this.value, 1, 10, 1));
      this.value = state.count;
      saveSettings();
    });

    $("insert-btn").addEventListener("click", insertStamp);
  }

  Office.onReady(function (info) {
    loadSettings();
    refreshUi();
    wireEvents();
    if (info.host === Office.HostType.Word) {
      if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
        $("insert-btn").disabled = false;
        setStatus("");
      } else {
        setStatus("このバージョンのWordは未対応です（Word 2016以降が必要）", true);
      }
    } else {
      setStatus("Word上で開いてください", true);
    }
  });
})();
