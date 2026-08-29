// オシテナ タスクペイン
/* global Office, Word, OshitenaOoxml */

(function () {
  "use strict";

  var STORAGE_KEY = "oshitena-settings-v1";

  var state = {
    sizeMm: 12,
    label: "",
    color: "D93A22",
    dash: "dash",
    strokePt: 1,
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

  function renderPreview() {
    var canvas = $("preview");
    var ctx = canvas.getContext("2d");
    var scale = 3.2; // 1mm → 3.2px
    var d = state.sizeMm * scale;
    var pad = 8;
    canvas.width = Math.max(80, Math.ceil(d + pad * 2));
    canvas.height = canvas.width;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    var cxy = canvas.width / 2;
    var r = d / 2;
    ctx.strokeStyle = "#" + state.color;
    ctx.lineWidth = Math.max(1, state.strokePt * scale * 0.353);
    if (state.dash === "dash") ctx.setLineDash([6, 4]);
    else if (state.dash === "sysDot") ctx.setLineDash([2, 3]);
    else ctx.setLineDash([]);
    ctx.beginPath();
    ctx.arc(cxy, cxy, r, 0, Math.PI * 2);
    ctx.stroke();
    if (state.label) {
      var fontPt = OshitenaOoxml.labelFontPt(state.sizeMm, state.label);
      ctx.setLineDash([]);
      ctx.fillStyle = "#" + state.color;
      ctx.font = fontPt * scale * 0.353 + 'px "Yu Mincho", "游明朝", serif';
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(state.label, cxy, cxy);
    }
  }

  function refreshUi() {
    syncChips("size-chips", "size", state.sizeMm);
    syncChips("label-chips", "label", state.label);
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
          label: state.label,
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
    $("size-chips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-size]");
      if (!chip) return;
      state.sizeMm = Number(chip.getAttribute("data-size"));
      refreshUi();
      saveSettings();
    });

    $("size-custom").addEventListener("change", function () {
      state.sizeMm = clampNumber(this.value, 6, 60, 12);
      refreshUi();
      saveSettings();
    });

    $("label-chips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-label]");
      if (!chip) return;
      state.label = chip.getAttribute("data-label");
      $("label-custom").value = "";
      refreshUi();
      saveSettings();
    });

    $("label-custom").addEventListener("input", function () {
      state.label = this.value.trim();
      syncChips("label-chips", "label", state.label);
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
