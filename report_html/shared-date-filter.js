(function () {
  if ((window.location.protocol || "").startsWith("file")) return;

  var GLOBAL_DATE_FILTER_API = "/api/report-date-filter";
  var FILE_TO_PAGE_KEY = {
    "dashboard.html": "dashboard",
    "executive_dashboard.html": "executive_dashboard",
    "nested_view_report.html": "nested_view_report",
    "employee_performance_report.html": "employee_performance_report",
    "assignee_hours_report.html": "assignee_hours_report",
    "rnd_data_story.html": "rnd_data_story",
    "phase_rmi_gantt_report.html": "phase_rmi_gantt_report",
    "planned_rmis_report.html": "planned_rmis_report",
    "planned_vs_dispensed_report.html": "approved_vs_planned_hours_report",
    "approved_vs_planned_hours_report.html": "approved_vs_planned_hours_report"
  };
  function currentPageKey() {
    var file = String((window.location.pathname || "").split("/").pop() || "").toLowerCase();
    return FILE_TO_PAGE_KEY[file] || file || "unknown_report";
  }

  function getInput(id) {
    return document.getElementById(id);
  }

  function parseIsoDate(value) {
    var text = String(value || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
    var d = new Date(text + "T00:00:00");
    if (!Number.isFinite(d.getTime())) return null;
    return d;
  }

  function toIsoDate(d) {
    if (!d || !Number.isFinite(d.getTime())) return "";
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  function toMonthValue(isoDate) {
    return String(isoDate || "").slice(0, 7);
  }

  function monthBounds(monthValue, endOfMonth) {
    var text = String(monthValue || "").trim();
    if (!/^\d{4}-\d{2}$/.test(text)) return "";
    var year = Number(text.slice(0, 4));
    var month = Number(text.slice(5, 7));
    if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return "";
    if (!endOfMonth) return text + "-01";
    var end = new Date(year, month, 0);
    return toIsoDate(end);
  }

  function detectPair(ids) {
    var fromEl = getInput(ids[0]);
    var toEl = getInput(ids[1]);
    if (!fromEl || !toEl) return null;
    return { fromEl: fromEl, toEl: toEl, fromId: ids[0], toId: ids[1] };
  }

  function primaryPair() {
    return (
      detectPair(["date-filter-from", "date-filter-to"]) ||
      detectPair(["from-date", "to-date"]) ||
      detectPair(["from", "to"])
    );
  }

  function page4Pair() {
    return detectPair(["from-date-page4", "to-date-page4"]);
  }

  function readNormalizedRange(pair) {
    if (!pair) return null;
    var fromRaw = String(pair.fromEl.value || "").trim();
    var toRaw = String(pair.toEl.value || "").trim();
    if (!fromRaw || !toRaw) return null;
    var isMonth = String(pair.fromEl.type || "").toLowerCase() === "month" || String(pair.toEl.type || "").toLowerCase() === "month";
    var fromDate = isMonth ? monthBounds(fromRaw, false) : fromRaw;
    var toDate = isMonth ? monthBounds(toRaw, true) : toRaw;
    var fromParsed = parseIsoDate(fromDate);
    var toParsed = parseIsoDate(toDate);
    if (!fromParsed || !toParsed || toParsed < fromParsed) return null;
    return { fromDate: toIsoDate(fromParsed), toDate: toIsoDate(toParsed), isMonth: isMonth };
  }

  function setRangeOnPair(pair, fromIso, toIso) {
    if (!pair) return;
    var isMonth = String(pair.fromEl.type || "").toLowerCase() === "month" || String(pair.toEl.type || "").toLowerCase() === "month";
    if (isMonth) {
      pair.fromEl.value = toMonthValue(fromIso);
      pair.toEl.value = toMonthValue(toIso);
      return;
    }
    pair.fromEl.value = String(fromIso || "");
    pair.toEl.value = String(toIso || "");
  }

  async function saveGlobalRange(fromDate, toDate, sourcePage) {
    var payload = {
      from_date: String(fromDate || ""),
      to_date: String(toDate || ""),
      source_page: String(sourcePage || currentPageKey())
    };
    if (!payload.from_date || !payload.to_date) return;
    await fetch(GLOBAL_DATE_FILTER_API, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
  }

  async function loadGlobalRange() {
    try {
      var response = await fetch(GLOBAL_DATE_FILTER_API, { cache: "no-store" });
      if (!response.ok) return null;
      var body = await response.json().catch(function () { return {}; });
      var filter = body && body.filter ? body.filter : null;
      if (!filter) return null;
      var fromDate = String(filter.from_date || "");
      var toDate = String(filter.to_date || "");
      if (!parseIsoDate(fromDate) || !parseIsoDate(toDate) || parseIsoDate(toDate) < parseIsoDate(fromDate)) return null;
      return { fromDate: fromDate, toDate: toDate };
    } catch (_err) {
      return null;
    }
  }

  function clickIfExists(id) {
    var el = document.getElementById(id);
    if (!el || typeof el.click !== "function") return false;
    el.click();
    return true;
  }

  function refreshAfterApply() {
    if (clickIfExists("date-filter-apply")) return;
    if (clickIfExists("apply-btn")) return;
    if (clickIfExists("apply")) return;
    if (clickIfExists("apply-page4-btn")) return;
    var pair = primaryPair();
    if (!pair) return;
    pair.fromEl.dispatchEvent(new Event("change", { bubbles: true }));
    pair.toEl.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function attachApplySave(buttonId, pairResolver, sourceSuffix) {
    var btn = document.getElementById(buttonId);
    if (!btn) return;
    btn.addEventListener("click", function () {
      var pair = pairResolver();
      var range = readNormalizedRange(pair);
      if (!range) return;
      var src = currentPageKey() + (sourceSuffix || "");
      saveGlobalRange(range.fromDate, range.toDate, src).catch(function () {});
    });
  }

  function attachResetSave(buttonId, pairResolver, sourceSuffix) {
    var btn = document.getElementById(buttonId);
    if (!btn) return;
    btn.addEventListener("click", function () {
      window.setTimeout(function () {
        var pair = pairResolver();
        var range = readNormalizedRange(pair);
        if (!range) return;
        var src = currentPageKey() + (sourceSuffix || "");
        saveGlobalRange(range.fromDate, range.toDate, src).catch(function () {});
      }, 0);
    });
  }

  function attachMonthChangeSave() {
    var pair = primaryPair();
    if (!pair) return;
    var isMonth = String(pair.fromEl.type || "").toLowerCase() === "month" || String(pair.toEl.type || "").toLowerCase() === "month";
    if (!isMonth) return;
    var onChange = function () {
      window.setTimeout(function () {
        var range = readNormalizedRange(primaryPair());
        if (!range) return;
        saveGlobalRange(range.fromDate, range.toDate, currentPageKey()).catch(function () {});
      }, 0);
    };
    pair.fromEl.addEventListener("change", onChange);
    pair.toEl.addEventListener("change", onChange);
  }

  function bindSaveHooks() {
    attachApplySave("date-filter-apply", primaryPair, "");
    attachApplySave("apply-btn", primaryPair, "");
    attachApplySave("apply", primaryPair, "");
    attachApplySave("apply-page4-btn", page4Pair, ":page4");

    attachResetSave("date-filter-reset", primaryPair, "");
    attachResetSave("reset-btn", primaryPair, "");
    attachResetSave("reset", primaryPair, "");
    attachResetSave("reset-page4-btn", page4Pair, ":page4");
    attachMonthChangeSave();
  }

  async function initGlobalDateFilter() {
    bindSaveHooks();
    var globalRange = await loadGlobalRange();
    if (!globalRange) return;
    setRangeOnPair(primaryPair(), globalRange.fromDate, globalRange.toDate);
    setRangeOnPair(page4Pair(), globalRange.fromDate, globalRange.toDate);
    refreshAfterApply();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      initGlobalDateFilter().catch(function () {});
    });
  } else {
    initGlobalDateFilter().catch(function () {});
  }
})();
