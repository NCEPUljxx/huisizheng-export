/* 慧思政题库导出 - content script v6 */
(function () {
  "use strict";

  var panel = null;
  var observer = null;
  var _obsDebounce = null;
  var _routeDebounce = null;
  var _pollTimer = null;
  var $ = function (sel, ctx) { return (ctx || document).querySelector(sel); };
  var $$ = function (sel, ctx) { return Array.from((ctx || document).querySelectorAll(sel)); };

  // ── Extract exam id from hash ──
  function hashExamId() {
    var m = location.hash.match(/[?&]id=(\d+)/);
    return m ? m[1] : "";
  }

  function hashRecordId() {
    // only used for logging/debug, not needed for name resolution
    var m = location.hash.match(/[?&]recordId=(\d+)/);
    return m ? m[1] : "";
  }

  function hashName() {
    var m = location.hash.match(/[?&]name=([^&]+)/);
    if (!m) return "";
    try { return decodeURIComponent(m[1]); } catch (e) { return ""; }
  }

  function isAnalysisPage() { return /\/exam\/analysis/.test(location.hash); }
  function isResultPage()   { return /\/exam\/result/.test(location.hash); }

  // ── Name: simply remember the most recently seen URL name ──
  // Whenever we see a result page URL with ?name=XXX, we overwrite the stored name.
  // The analysis page then uses that last-seen name.

  function storeLatestName(name) {
    if (!name) return;
    try { sessionStorage.setItem("hsz_latest_name", name); } catch (e) {}
  }

  function getLatestName() {
    try { return sessionStorage.getItem("hsz_latest_name") || ""; } catch (e) { return ""; }
  }

  function resolveExamName() {
    var urlName = hashName();

    // Rule: if the URL contains a name, THAT is the current exam name. Always.
    if (urlName) {
      storeLatestName(urlName);
      return urlName;
    }

    // Rule: on analysis page, use the last name we saw (from the parent result page)
    if (isAnalysisPage()) {
      var latest = getLatestName();
      if (latest) return latest;
    }

    // Rule: clear stale name when entering a new result page that has NO name
    // (shouldn't happen, but just in case)
    if (isResultPage() && !urlName) {
      // Don't clear — keep whatever was last seen
    }

    return getLatestName() || "";
  }

  // ── Extract questions from DOM ──
  function extractQuestions() {
    var items = $$(".item");
    if (!items.length) return [];

    return items.map(function (item, idx) {
      var topicName = $(".topic-name", item);
      var qText = topicName ? topicName.textContent.trim() : "题目" + (idx + 1);
      var optionEls = $$(".topic-option", item);
      var options = [];
      var correctAnswer = "";

      optionEls.forEach(function (optEl) {
        var letter = "", text = "", isCorrect = false;
        var spans = $$("span", optEl);
        spans.forEach(function (span) {
          var c = span.textContent.trim();
          if (/^[A-J]\.?\s*$/.test(c)) letter = c.replace(/\./g, "").trim();
          else if (c === "正确答案") isCorrect = true;
          else if (c && c !== "正确答案" && !/^[A-J]\.?\s*$/.test(c)) text = c;
        });
        if (letter && text) {
          options.push({ letter: letter, text: text, isCorrect: isCorrect });
          if (isCorrect) correctAnswer = letter;
        }
      });

      return {
        number: idx + 1,
        type: options.length === 2 ? "判断题" : "单选题",
        text: qText,
        options: options,
        correctAnswer: correctAnswer,
      };
    });
  }

  // ── Build Word HTML ──
  function esc(str) {
    if (!str) return "";
    return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function buildDocHTML(questions, mode, examTitle) {
    var sections = [
      { title: "一、单选题", filter: function (q) { return q.type === "单选题"; } },
      { title: "二、判断题", filter: function (q) { return q.type === "判断题"; } },
    ];

    var activeSections = [];
    sections.forEach(function (sec) {
      var qs = questions.filter(sec.filter);
      if (qs.length) activeSections.push({ title: sec.title, questions: qs });
    });

    var titleText = examTitle || "慧思政题库";
    var labels = { "full": " — 带答案", "no-answer": " — 无答案", "answers-only": " — 仅答案" };
    var body = '<p class="MsoTitle">' + esc(titleText) + esc(labels[mode]) + '</p>';

    var markCorrect = (mode === "full");
    var showCorrectOnly = (mode === "answers-only");

    for (var s = 0; s < activeSections.length; s++) {
      var sec = activeSections[s];
      body += '<h2>' + esc(sec.title) + '</h2>';
      for (var i = 0; i < sec.questions.length; i++) {
        var q = sec.questions[i];
        body += '<p class="QuestionText">' + esc(q.text) + '</p>';
        if (showCorrectOnly) {
          var co = q.options.find(function (o) { return o.isCorrect; });
          if (co) body += '<p class="OptionText"><b style="color:#CC0000;">' + esc(co.letter) + '. ' + esc(co.text) + '</b></p>';
        } else {
          for (var j = 0; j < q.options.length; j++) {
            var opt = q.options[j];
            if (opt.isCorrect && markCorrect) {
              body += '<p class="OptionText"><b style="color:#CC0000;">' + esc(opt.letter) + '. ' + esc(opt.text) + '（正确答案）</b></p>';
            } else {
              body += '<p class="OptionText">' + esc(opt.letter) + '. ' + esc(opt.text) + '</p>';
            }
          }
        }
      }
    }

    return [
      '<html xmlns:o="urn:schemas-microsoft-com:office:office"',
      '      xmlns:w="urn:schemas-microsoft-com:office:word"',
      '      xmlns="http://www.w3.org/TR/REC-html40">',
      '<head><meta charset="utf-8">',
      '<!--[if gte mso 9]><xml>',
      '<w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument>',
      '</xml><![endif]-->',
      '<style>',
      '  @page WordSection1 { size: 595.3pt 841.9pt; margin: 70.85pt; mso-page-orientation: portrait;',
      '    mso-header-margin: 42.55pt; mso-footer-margin: 42.55pt; mso-paper-source: 0; }',
      '  div.WordSection1 { page: WordSection1; }',
      '  body { font-family: SimSun, serif; font-size: 12.0pt; line-height: 1.8; color: #000; }',
      '  p.MsoTitle { text-align: center; font-family: SimHei, sans-serif; font-size: 22.0pt; font-weight: bold; margin-bottom: 6pt; }',
      '  p.MsoSubtitle { text-align: center; font-family: SimSun, serif; font-size: 10.5pt; color: #888; margin-top: 0; margin-bottom: 24pt; }',
      '  h2 { font-family: SimHei, sans-serif; font-size: 16.0pt; font-weight: bold; margin-top: 18pt; margin-bottom: 12pt; }',
      '  p.QuestionText { font-family: SimSun, serif; font-size: 12.0pt; font-weight: bold; margin-top: 14pt; margin-bottom: 4pt; }',
      '  p.OptionText { font-family: SimSun, serif; font-size: 12.0pt; margin-top: 1pt; margin-bottom: 1pt; margin-left: 24.0pt; text-indent: 0; }',
      '</style></head><body><div class="WordSection1">',
      body,
      '</div></body></html>'
    ].join("\n");
  }

  function downloadDoc(content, filename) {
    var blob = new Blob([content], { type: "application/msword" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function handleExport(mode) {
    var questions = extractQuestions();
    if (!questions.length) {
      alert("未检测到题目，请等待页面加载完毕后再试。");
      return;
    }
    var name = resolveExamName() || "慧思政题库";
    var html = buildDocHTML(questions, mode, name);
    var labels = { "full": "带答案", "no-answer": "无答案", "answers-only": "仅答案" };
    var safe = name.replace(/[\\/:*?"<>|]/g, "_") || "慧思政题库";
    downloadDoc(html, safe + "_" + labels[mode] + ".doc");
    toast("已导出 " + questions.length + " 题 (" + labels[mode] + ")");
  }

  function toast(msg) {
    var el = document.createElement("div");
    el.textContent = msg;
    Object.assign(el.style, {
      position: "fixed", bottom: "80px", right: "20px",
      background: "#4CAF50", color: "#fff", padding: "12px 20px",
      borderRadius: "6px", fontSize: "14px", zIndex: "99999",
      boxShadow: "0 2px 12px rgba(0,0,0,0.3)", transition: "opacity 0.3s",
    });
    document.body.appendChild(el);
    setTimeout(function () { el.style.opacity = "0"; setTimeout(function () { el.remove(); }, 300); }, 2000);
  }

  // ── Panel ──
  function ensurePanel() {
    if (panel) return;
    panel = document.createElement("div");
    panel.id = "huisizheng-export-panel";
    panel.innerHTML =
      '<div class="hse-header">' +
        '<span class="hse-title">📋 慧思政导出</span>' +
        '<span class="hse-count" id="hse-count">检测中...</span>' +
      '</div>' +
      '<div class="hse-buttons">' +
        '<button class="hse-btn hse-btn-full">✅ 导出 Word (带答案)</button>' +
        '<button class="hse-btn hse-btn-noans">📝 导出 Word (无答案)</button>' +
        '<button class="hse-btn hse-btn-key">🔑 导出 Word (仅答案)</button>' +
      '</div>' +
      '<div class="hse-footer"><span class="hse-id" id="hse-exam-id"></span></div>';
    document.body.appendChild(panel);

    var btns = panel.querySelectorAll("button");
    btns[0].addEventListener("click", function () { handleExport("full"); });
    btns[1].addEventListener("click", function () { handleExport("no-answer"); });
    btns[2].addEventListener("click", function () { handleExport("answers-only"); });
  }

  function updatePanel() {
    var count = $$(".item").length;
    var countEl = document.getElementById("hse-count");
    var idEl = document.getElementById("hse-exam-id");

    if (countEl) {
      countEl.textContent = count ? count + " 题" : "检测中...";
      countEl.style.color = count > 0 ? "#4CAF50" : "#999";
    }
    if (idEl) {
      var name = resolveExamName();
      idEl.textContent = name || ("试卷 ID: " + hashExamId());
    }

    if (count > 0 && panel) panel.style.display = "";
  }

  // ── Observer (watch DOM for question appearance) ──
  function startObserver() {
    if (observer) return;
    observer = new MutationObserver(function (mutations) {
      for (var i = 0; i < mutations.length; i++) {
        var t = mutations[i].target;
        var skip = (t.id === "huisizheng-export-panel");
        if (!skip && t.closest) skip = !!t.closest("#huisizheng-export-panel");
        if (!skip) {
          clearTimeout(_obsDebounce);
          _obsDebounce = setTimeout(updatePanel, 200);
          return;
        }
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  // ── Route handler (called on every hash change + periodically) ──
  var _lastHash = "";

  function handleRoute() {
    // Always capture name if URL has one (result page)
    var urlName = hashName();
    if (urlName) storeLatestName(urlName);

    _lastHash = location.hash;

    if (isAnalysisPage()) {
      ensurePanel();
      startObserver();
      updatePanel();
      startPolling();
    } else {
      if (panel) panel.style.display = "none";
      stopPolling();
    }
  }

  // ── URL Polling: catch SPA nav that bypasses hashchange ──
  function startPolling() {
    if (_pollTimer) return;
    _pollTimer = setInterval(function () {
      if (location.hash !== _lastHash) {
        handleRoute();
        return;
      }
      // Even if hash hasn't changed, items may have loaded
      if (isAnalysisPage()) {
        var count = $$(".item").length;
        if (count > 0) updatePanel();
      }
    }, 500);
  }

  function stopPolling() {
    if (_pollTimer) { clearInterval(_pollTimer); _pollTimer = null; }
  }

  // ── Init ──
  function init() {
    // Capture name if already on a result page with ?name=
    var urlName = hashName();
    if (urlName) storeLatestName(urlName);

    // Set up for current page
    handleRoute();

    // Listen for hash changes
    window.addEventListener("hashchange", function () {
      clearTimeout(_routeDebounce);
      _routeDebounce = setTimeout(handleRoute, 100);
    });

    // Poll: some SPA routers modify hash without firing hashchange
    startPolling();

    // Aggressive retry for slow SPA rendering
    [300, 800, 1500, 3000, 6000, 10000].forEach(function (ms) {
      setTimeout(function () {
        var n = hashName();
        if (n) storeLatestName(n);
        if (isAnalysisPage()) {
          ensurePanel();
          startObserver();
          updatePanel();
        }
      }, ms);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
