/* Interactive dumbbell: Annex I / Annex II shares of world cumulative emissions.
 * Data payload built by scripts/12_prepare_interactive_dumbbell.py.
 * Share maths mirrors the discontinued dynamic-date dumbbell workbook (SUMIFS
 * over the selected year range; world = annex 1 + non-annex 1; GMST via
 * cumulative differencing). QA'd against static Figures 3/4/8.
 */
(function () {
  "use strict";

  var DATA = window.EMISSIONS_DATA;

  // Groups are configurable so more rows (G20, custom aggregates) can be
  // added later: each entry just needs a key present in the data payload.
  var GROUPS = [
    { key: "annex2", label: "Annex II" },
    { key: "annex1", label: "Annex I" }
  ];

  // A source is shown when it covers the selected start year and reaches at
  // least (end year - TAIL_TOLERANCE): Climate Watch ends in 2023, and the
  // static Figure 4 (1990-2024) includes it on the same basis.
  var TAIL_TOLERANCE = 2;

  // Slider floor. OWID/PRIMAP data reach back to 1750 but pre-1850 coverage is
  // thin, so the chart starts at 1850 to match the static figures.
  var MIN_YEAR = 1850;

  // Palette from cgd-brand-reference.md (CGD Interactive Toolkit). The three pale
  // categorical entries are unusable as 12px dots on white, so the five series use
  // the readable subset plus the two status colours; GMST is a "+" glyph rather
  // than a dot, so Teal Black keeps it legible without competing for a hue.
  var SOURCE_META = {
    OWID:         { label: "OWID",                  legend: "OWID",          color: "#006970" },
    PRIMAP:       { label: "PRIMAP-hist",           legend: "PRIMAP",        color: "#FFB52C" },
    GCP_fossil:   { label: "GCP (fossil)",          legend: "GCP",           color: "#2D99B5" },
    GCP_BLUE:     { label: "GCP (BLUE model)",      legend: "GCP",           color: "#2D99B5" },
    GCP_OSCAR:    { label: "GCP (OSCAR model)",     legend: "GCP",           color: "#2D99B5" },
    GCP_LUCE:     { label: "GCP (LUCE model)",      legend: "GCP",           color: "#2D99B5" },
    EDGAR:        { label: "EDGAR",                 legend: "EDGAR",         color: "#D15553" },
    ClimateWatch: { label: "Climate Watch",         legend: "Climate Watch", color: "#00896C" }
  };
  var LEGEND_ORDER = ["OWID", "PRIMAP", "GCP", "EDGAR", "Climate Watch"];
  var GMST_COLOR = "#1A272A";

  var MEASURE_TITLES = {
    co2_excLUC: "CO₂ emissions (excluding LULUCF)",
    co2_incLUC: "CO₂ emissions (including LULUCF)",
    ghg_excLUC: "GHG emissions (excluding LULUCF)",
    ghg_incLUC: "GHG emissions (including LULUCF)"
  };

  var state = {
    measure: "ghg_incLUC",
    attribution: "asis",
    start: 1850,
    end: 2024,
    customGroups: []   // {id, name, isos: [iso, ...]}
  };
  var MAX_CUSTOM_GROUPS = 4;
  var customIdSeq = 1;

  // Country-level payloads (countries_[measure].js), lazy-loaded on first use
  // so blog embeds without the picker never fetch them.
  var countryData = {};
  var countryLoading = {};

  function ensureCountryData(measure) {
    if (countryData[measure]) return true;
    if (countryLoading[measure]) return false;
    countryLoading[measure] = true;
    var s = document.createElement("script");
    s.src = "countries_" + measure + ".js";
    s.onload = function () {
      countryData[measure] = (window.EMISSIONS_COUNTRIES || {})[measure];
      countryLoading[measure] = false;
      render();
    };
    s.onerror = function () { countryLoading[measure] = false; };
    document.body.appendChild(s);
    return false;
  }

  // ---- data helpers -------------------------------------------------------

  function attrBlock() { return DATA.attributions[state.attribution]; }

  function yearBounds(attribution) {
    var block = DATA.attributions[attribution];
    var lo = Infinity, hi = -Infinity;
    Object.keys(block.measures).forEach(function (msr) {
      var srcs = block.measures[msr];
      Object.keys(srcs).forEach(function (s) {
        lo = Math.min(lo, srcs[s].start);
        hi = Math.max(hi, srcs[s].end);
      });
    });
    return { min: Math.max(lo, MIN_YEAR), max: hi };
  }

  function sumRange(src, grpKey, start, end) {
    var vals = src[grpKey];
    var total = 0;
    var y0 = Math.max(start, src.start);
    var y1 = Math.min(end, src.end);
    for (var y = y0; y <= y1; y++) {
      var v = vals[y - src.start];
      if (v !== null) total += v;
    }
    return total;
  }

  function worldSum(src, start, end) {
    return sumRange(src, "annex1", start, end) + sumRange(src, "nona1", start, end);
  }

  function share(src, grpKey, start, end) {
    var den = worldSum(src, start, end);
    return den === 0 ? null : sumRange(src, grpKey, start, end) / den * 100;
  }

  // ---- custom-group helpers (country-level payload) -------------------------

  function entrySum(e, start, end) {
    var total = 0, any = false;
    var last = e.s + e.v.length - 1;
    for (var y = Math.max(start, e.s); y <= Math.min(end, last); y++) {
      var v = e.v[y - e.s];
      if (v !== null) { total += v; any = true; }
    }
    return { sum: total, any: any };
  }

  function gmstCumAt(e, y) {
    if (y < e.s) return 0;
    var v = e.v[Math.min(y, e.s + e.v.length - 1) - e.s];
    return v === null ? 0 : v;
  }

  function customPoints(group) {
    var cd = countryData[state.measure];
    var em = cd[state.attribution];
    var srcs = attrBlock().measures[state.measure];
    var points = [];
    activeSources().forEach(function (name) {
      var entries = em[name];
      if (!entries) return;
      var covered = 0, total = 0;
      group.isos.forEach(function (iso) {
        var e = entries[iso];
        if (!e) return;
        var r = entrySum(e, state.start, state.end);
        if (r.any) { covered++; total += r.sum; }
      });
      if (covered === 0) return;
      var den = worldSum(srcs[name], state.start, state.end);
      if (den === 0) return;
      var notes = [];
      var s = srcs[name];
      if (s.end < state.end) notes.push("Data available to " + s.end + ".");
      if (s.start > state.start) notes.push("Data from " + s.start + ".");
      if (covered < group.isos.length) {
        notes.push("Includes " + covered + " of " + group.isos.length + " selected countries.");
      }
      points.push({ source: name, value: total / den * 100, notes: notes });
    });
    return points;
  }

  function customGmst(group) {
    if (!gmstActive()) return { value: null };
    var g = attrBlock().gmst;
    var cg = countryData[state.measure].gmst[state.attribution];
    var series = g[state.measure];
    function cumW(y) { return y < g.start ? 0 : series.world[Math.min(y, g.end) - g.start]; }
    var preW = state.start > 1851 ? cumW(state.start - 1) : 0;
    var den = cumW(state.end) - preW;
    if (!den) return { value: null };
    var total = 0, covered = 0;
    group.isos.forEach(function (iso) {
      var e = cg[iso];
      if (!e) return;
      covered++;
      var pre = state.start > 1851 ? gmstCumAt(e, state.start - 1) : 0;
      total += gmstCumAt(e, state.end) - pre;
    });
    if (covered === 0) return { value: null };
    var notes = state.start < 1851 ? ["GMST response measured from 1851."] : [];
    if (covered < group.isos.length) {
      notes.push("Includes " + covered + " of " + group.isos.length + " selected countries.");
    }
    return { value: total / den * 100, notes: notes };
  }

  function gmstShare(grpKey, start, end) {
    var g = attrBlock().gmst;
    if (end < g.start) return null;
    var series = g[state.measure];
    function cum(arr, y) {
      if (y < g.start) return 0;
      return arr[Math.min(y, g.end) - g.start];
    }
    var preG = start > 1851 ? cum(series[grpKey], start - 1) : 0;
    var preW = start > 1851 ? cum(series.world, start - 1) : 0;
    var den = cum(series.world, end) - preW;
    if (!den) return null;
    return (cum(series[grpKey], end) - preG) / den * 100;
  }

  function activeSources() {
    var srcs = attrBlock().measures[state.measure];
    return Object.keys(srcs).filter(function (name) {
      var s = srcs[name];
      return s.start <= state.start && s.end >= state.end - TAIL_TOLERANCE;
    });
  }

  function gmstActive() {
    return state.end >= attrBlock().gmst.start;
  }

  // Values for the current state:
  // [{label, custom?, id?, loading?, points: [{source, value, notes}], gmst, gmstNotes}]
  function currentValues() {
    var srcs = attrBlock().measures[state.measure];
    var names = activeSources();
    var gmstNotes = state.start < 1851 ? ["GMST response measured from 1851."] : [];

    var rows = GROUPS.map(function (grp) {
      var points = names.map(function (name) {
        var s = srcs[name];
        var v = share(s, grp.key, state.start, state.end);
        var notes = [];
        if (s.end < state.end) notes.push("Data available to " + s.end + ".");
        if (s.start > state.start) notes.push("Data from " + s.start + ".");
        return { source: name, value: v, notes: notes };
      }).filter(function (p) { return p.value !== null; });
      var gmst = gmstActive() ? gmstShare(grp.key, state.start, state.end) : null;
      return { label: grp.label, points: points, gmst: gmst, gmstNotes: gmstNotes };
    });

    if (state.customGroups.length) {
      var loaded = ensureCountryData(state.measure);
      state.customGroups.forEach(function (group) {
        if (!loaded) {
          rows.push({ label: group.name, custom: true, id: group.id,
                      loading: true, points: [], gmst: null, gmstNotes: [] });
          return;
        }
        var g = customGmst(group);
        rows.push({
          label: group.name, custom: true, id: group.id,
          points: customPoints(group),
          gmst: g.value === undefined ? null : g.value,
          gmstNotes: g.notes || []
        });
      });
    }
    return rows;
  }

  // ---- DOM ------------------------------------------------------------------

  var el = {
    title: document.getElementById("chartTitle"),
    measure: document.getElementById("measureSelect"),
    colonial: document.getElementById("colonialToggle"),
    startRange: document.getElementById("startRange"),
    endRange: document.getElementById("endRange"),
    rangeFill: document.getElementById("rangeFill"),
    rangeMinLabel: document.getElementById("rangeMinLabel"),
    rangeMaxLabel: document.getElementById("rangeMaxLabel"),
    yearReadout: document.getElementById("yearReadout"),
    yearReset: document.getElementById("yearReset"),
    groupList: document.getElementById("groupList"),
    legend: document.getElementById("legend"),
    svg: document.getElementById("chart"),
    tooltip: document.getElementById("tooltip"),
    wrap: document.querySelector(".chart-wrap"),
    footnote: document.getElementById("footnote"),
    table: document.getElementById("dataTable"),
    tableDetails: document.querySelector(".data-table")
  };

  // ---- analytics --------------------------------------------------------------

  // Thin wrapper so a failed tracking.js load cannot break the chart.
  // Events are documented in TRACKING.md — update it alongside any change here.
  function track(actionType, actionLabel, actionValue) {
    if (window.CGDTracking) {
      window.CGDTracking.trackEngagement(actionType, actionLabel, actionValue);
    }
  }

  var SVG_NS = "http://www.w3.org/2000/svg";
  function svgEl(tag, attrs) {
    var node = document.createElementNS(SVG_NS, tag);
    Object.keys(attrs || {}).forEach(function (k) { node.setAttribute(k, attrs[k]); });
    return node;
  }

  // ---- slider ---------------------------------------------------------------

  function setSliderBounds(attribution) {
    var b = yearBounds(attribution);
    [el.startRange, el.endRange].forEach(function (input) {
      input.min = b.min;
      input.max = b.max;
    });
    state.start = Math.min(Math.max(state.start, b.min), b.max);
    state.end = Math.min(Math.max(state.end, state.start), b.max);
    el.startRange.value = state.start;
    el.endRange.value = state.end;
    el.rangeMinLabel.textContent = b.min;
    el.rangeMaxLabel.textContent = b.max;
  }

  function updateSliderVisuals() {
    var min = +el.startRange.min, max = +el.startRange.max;
    var lo = (state.start - min) / (max - min) * 100;
    var hi = (state.end - min) / (max - min) * 100;
    el.rangeFill.style.left = lo + "%";
    el.rangeFill.style.width = (hi - lo) + "%";
    el.yearReadout.textContent = state.start + "–" + state.end;

    // With both handles on the same year the top one wins every pointer grab.
    // At the maximum only the start handle has anywhere to go, so lift it above
    // the end handle; otherwise the end handle stays on top as usual.
    var stuckAtMax = state.start === state.end && state.end >= max;
    el.startRange.style.zIndex = stuckAtMax ? 4 : 2;
    el.endRange.style.zIndex = stuckAtMax ? 2 : 3;

    el.yearReset.disabled = state.start === min && state.end === max;
    el.yearReset.title = el.yearReset.disabled ? "" :
      "Reset to the full range (" + min + "–" + max + ")";
  }

  // Dragging fires "input" roughly every 16ms. A debounce is wrong here: each event
  // resets the timer, so the chart never redraws until the drag stops and the dots
  // jump to their final position. This throttles instead — redraw immediately, then
  // at most once per REDRAW_MS, with a trailing call so the last value always lands.
  // 40ms is the coding standard's lower bound, which it allows when the update is
  // cheap; this one only recomputes sums over the selected year range.
  var REDRAW_MS = 40;
  var renderTimer = null;
  var lastRenderAt = -Infinity;

  function renderThrottled() {
    updateSliderVisuals();          // readout and track fill follow the handle exactly
    clearTimeout(renderTimer);
    var wait = REDRAW_MS - (performance.now() - lastRenderAt);
    if (wait <= 0) {
      lastRenderAt = performance.now();
      render();
    } else {
      renderTimer = setTimeout(function () {
        lastRenderAt = performance.now();
        render();
      }, wait);
    }
  }

  el.startRange.addEventListener("input", function () {
    state.start = Math.min(+el.startRange.value, state.end);
    el.startRange.value = state.start;
    renderThrottled();
  });
  el.endRange.addEventListener("input", function () {
    state.end = Math.max(+el.endRange.value, state.start);
    el.endRange.value = state.end;
    renderThrottled();
  });
  el.yearReset.addEventListener("click", function () {
    state.start = +el.startRange.min;
    state.end = +el.startRange.max;
    el.startRange.value = state.start;
    el.endRange.value = state.end;
    track("preset", "year_range_reset");
    render();
  });

  // Tracking fires on "change" (pointer release), not on the "input" events that
  // drive rendering — one event per adjustment rather than one per pixel dragged.
  [el.startRange, el.endRange].forEach(function (input) {
    input.addEventListener("change", function () {
      track("filter", "year_range", periodText());
    });
  });

  // ---- controls ---------------------------------------------------------------

  el.measure.addEventListener("change", function () {
    state.measure = el.measure.value;
    track("filter", "measure_select", state.measure);
    render();
  });

  el.colonial.addEventListener("change", function () {
    state.attribution = el.colonial.checked ? "colonial" : "asis";
    setSliderBounds(state.attribution);
    track("filter", "colonial_attribution", el.colonial.checked ? "on" : "off");
    render();
  });

  // ---- tooltip ----------------------------------------------------------------

  // entries: [{valueText, labelText, color}] — several when marks overlap
  function showTooltip(anchorX, anchorY, entries, noteLines) {
    var tt = el.tooltip;
    tt.textContent = "";
    entries.forEach(function (e) {
      var row = document.createElement("div");
      row.className = "tt-row";
      var key = document.createElement("span");
      key.className = "tt-key";
      key.style.background = e.color;
      var v = document.createElement("span");
      v.className = "tt-value";
      v.textContent = e.valueText;
      var l = document.createElement("span");
      l.className = "tt-label";
      l.textContent = " " + e.labelText;
      row.appendChild(key);
      row.appendChild(v);
      row.appendChild(l);
      tt.appendChild(row);
    });
    (noteLines || []).forEach(function (n) {
      var d = document.createElement("div");
      d.className = "tt-note";
      d.textContent = n;
      tt.appendChild(d);
    });
    tt.hidden = false;
    var wrapRect = el.wrap.getBoundingClientRect();
    var x = anchorX - wrapRect.left + 14;
    var y = anchorY - wrapRect.top - 10;
    if (x + tt.offsetWidth > wrapRect.width - 4) x = x - tt.offsetWidth - 28;
    if (x < 0) x = 4;
    tt.style.left = x + "px";
    tt.style.top = Math.max(0, y - tt.offsetHeight / 2) + "px";
  }

  function hideTooltip() { el.tooltip.hidden = true; }

  // ---- country picker -------------------------------------------------------

  var picker = {
    section: document.getElementById("customGroups"),
    btn: document.getElementById("addGroupBtn"),
    panel: document.getElementById("picker"),
    search: document.getElementById("pickerSearch"),
    groupBy: document.getElementById("pickerGroupBy"),
    list: document.getElementById("pickerList"),
    count: document.getElementById("pickerCount"),
    name: document.getElementById("pickerName"),
    add: document.getElementById("pickerAdd"),
    cancel: document.getElementById("pickerCancel")
  };
  var pickerSelection = {};   // iso -> true

  var ANNEX_LABELS = { a2: "Annex II", a1: "Annex I (other)", na1: "Non-Annex I" };
  var OTHER_BUCKET = "Not in these groups";
  // Institutional (G20/G7/EU27/OECD/BRICS) and development (LDC/LLDC/SIDS) sets
  // come from data.js (script 14, flags in data/country_groups.csv). They overlap,
  // so a country can appear under more than one heading in those modes.
  var SET_GROUPS = DATA.setGroups || {};
  var BUCKET_ORDERS = {
    annex: ["Annex II", "Annex I (other)", "Non-Annex I"],
    income: ["High income", "Upper-middle income", "Lower-middle income",
             "Low income", "Unclassified"],
    region: ["Europe & Central Asia", "North America", "Latin America & Caribbean",
             "East Asia & Pacific", "South Asia", "Middle East & North Africa",
             "Sub-Saharan Africa", "Other"],
    institutional: (SET_GROUPS.institutional || []).concat([OTHER_BUCKET]),
    development: (SET_GROUPS.development || []).concat([OTHER_BUCKET])
  };

  function bucketsOf(c, mode) {
    if (mode === "annex") return [ANNEX_LABELS[c.annex]];
    if (mode === "income") return [c.income];
    if (mode === "region") return [c.region];
    if (SET_GROUPS[mode]) {
      var wanted = SET_GROUPS[mode];
      var mine = (c.sets || []).filter(function (s) { return wanted.indexOf(s) !== -1; });
      return mine.length ? mine : [OTHER_BUCKET];
    }
    return ["All countries"];
  }

  // Selecting a whole bucket is the common case, so offer its name as the row
  // label instead of "Custom group N". Keyed on the exact ISO set.
  var namedSetIndex = null;
  function namedSets() {
    if (namedSetIndex) return namedSetIndex;
    namedSetIndex = [];
    ["annex", "income", "region", "institutional", "development"].forEach(function (mode) {
      var buckets = {};
      DATA.countries.forEach(function (c) {
        bucketsOf(c, mode).forEach(function (b) {
          (buckets[b] = buckets[b] || []).push(c.iso);
        });
      });
      Object.keys(buckets).forEach(function (b) {
        if (b === OTHER_BUCKET) return;
        namedSetIndex.push({ name: b, key: buckets[b].slice().sort().join(",") });
      });
    });
    return namedSetIndex;
  }

  function matchNamedSet(isos) {
    var key = isos.slice().sort().join(",");
    var hit = null;
    namedSets().forEach(function (s) { if (!hit && s.key === key) hit = s.name; });
    return hit;
  }

  function updatePickerFooter() {
    var isos = Object.keys(pickerSelection);
    picker.count.textContent = isos.length + " selected";
    picker.add.disabled = isos.length === 0;
    var match = isos.length ? matchNamedSet(isos) : null;
    picker.name.placeholder = match || "Group name (optional)";
  }

  function buildPickerList() {
    var mode = picker.groupBy.value;
    var q = picker.search.value.trim().toLowerCase();
    var scroll = picker.list.scrollTop;
    picker.list.textContent = "";

    var buckets = {};
    DATA.countries.forEach(function (c) {
      if (q && c.name.toLowerCase().indexOf(q) === -1) return;
      bucketsOf(c, mode).forEach(function (b) {
        (buckets[b] = buckets[b] || []).push(c);
      });
    });

    var order = mode === "none" ? ["All countries"] : BUCKET_ORDERS[mode];
    order.forEach(function (bname) {
      var items = buckets[bname];
      if (!items || !items.length) return;

      var header = document.createElement("label");
      header.className = "picker-bucket";
      var cb = document.createElement("input");
      cb.type = "checkbox";
      var allSel = items.every(function (c) { return pickerSelection[c.iso]; });
      var someSel = items.some(function (c) { return pickerSelection[c.iso]; });
      cb.checked = allSel;
      cb.indeterminate = someSel && !allSel;
      cb.addEventListener("change", function () {
        items.forEach(function (c) {
          if (cb.checked) pickerSelection[c.iso] = true;
          else delete pickerSelection[c.iso];
        });
        buildPickerList();
        updatePickerFooter();
      });
      header.appendChild(cb);
      header.appendChild(document.createTextNode(bname + " (" + items.length + ")"));
      picker.list.appendChild(header);

      items.forEach(function (c) {
        var item = document.createElement("label");
        item.className = "picker-item";
        var icb = document.createElement("input");
        icb.type = "checkbox";
        icb.checked = !!pickerSelection[c.iso];
        icb.addEventListener("change", function () {
          if (icb.checked) pickerSelection[c.iso] = true;
          else delete pickerSelection[c.iso];
          buildPickerList();
          updatePickerFooter();
        });
        item.appendChild(icb);
        item.appendChild(document.createTextNode(c.name));
        picker.list.appendChild(item);
      });
    });

    if (!picker.list.childNodes.length) {
      var none = document.createElement("div");
      none.className = "picker-empty";
      none.textContent = "No countries match.";
      picker.list.appendChild(none);
    }
    picker.list.scrollTop = scroll;
  }

  function openPicker() {
    pickerSelection = {};
    picker.search.value = "";
    picker.name.value = "";
    picker.panel.hidden = false;
    picker.btn.hidden = true;
    buildPickerList();
    updatePickerFooter();
    ensureCountryData(state.measure);   // warm the cache while the user picks
    picker.search.focus();
  }

  function closePicker() {
    picker.panel.hidden = true;
    picker.btn.hidden = false;
  }

  picker.btn.addEventListener("click", function () {
    track("detail_open", "country_picker");
    openPicker();
  });
  picker.cancel.addEventListener("click", function () {
    track("detail_close", "country_picker");
    closePicker();
  });
  picker.search.addEventListener("input", buildPickerList);
  picker.groupBy.addEventListener("change", function () {
    track("view_control", "picker_group_by", picker.groupBy.value);
    buildPickerList();
  });
  picker.add.addEventListener("click", function () {
    var isos = Object.keys(pickerSelection);
    if (!isos.length) return;
    var namedSet = matchNamedSet(isos);
    var name = picker.name.value.trim() || namedSet || "Custom group " + customIdSeq;
    state.customGroups.push({ id: customIdSeq++, name: name, isos: isos });
    // Only a matched named set is reported: user-typed and default names are
    // unbounded, and the standard's cardinality rule says omit rather than guess.
    track("preset", "add_country_group", namedSet || undefined);
    closePicker();
    render();
  });

  // ---- added groups: membership list ----------------------------------------

  var isoNameIndex = null;
  function isoNames() {
    if (isoNameIndex) return isoNameIndex;
    isoNameIndex = {};
    DATA.countries.forEach(function (c) { isoNameIndex[c.iso] = c.name; });
    return isoNameIndex;
  }

  function memberNames(group) {
    var names = isoNames();
    return group.isos.map(function (iso) { return names[iso] || iso; }).sort();
  }

  function renderGroupList() {
    var host = el.groupList;
    host.textContent = "";
    state.customGroups.forEach(function (group) {
      var chip = document.createElement("div");
      chip.className = "group-chip";

      var head = document.createElement("div");
      head.className = "group-chip-head";

      var toggle = document.createElement("button");
      toggle.type = "button";
      toggle.className = "group-view";
      toggle.setAttribute("aria-expanded", group.open ? "true" : "false");
      toggle.textContent = (group.open ? "▾ " : "▸ ") + group.name + " — " +
        group.isos.length + (group.isos.length === 1 ? " country" : " countries");
      toggle.addEventListener("click", function () {
        group.open = !group.open;
        track(group.open ? "detail_open" : "detail_close", "country_group_members");
        renderGroupList();
      });

      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "group-remove";
      rm.setAttribute("aria-label", "Remove " + group.name);
      rm.title = "Remove this group";
      rm.textContent = "×";
      rm.addEventListener("click", function () {
        state.customGroups = state.customGroups.filter(function (g) { return g.id !== group.id; });
        track("filter", "remove_country_group");
        render();
      });

      head.appendChild(toggle);
      head.appendChild(rm);
      chip.appendChild(head);

      if (group.open) {
        var mem = document.createElement("div");
        mem.className = "group-members";
        mem.textContent = memberNames(group).join(", ");
        chip.appendChild(mem);
      }
      host.appendChild(chip);
    });
  }

  // ---- render -------------------------------------------------------------------

  function fmt(v) { return v === null ? "—" : v.toFixed(1) + "%"; }

  function periodText() { return state.start + "–" + state.end; }

  function renderTitle() {
    var who = state.customGroups.length ? "Country-group shares" :
              "Annex I and Annex II shares";
    // Colonial attribution is deliberately not named here: the extra clause pushed
    // the title onto a second line and shifted the chart down on every toggle.
    // The footnote states it instead (see renderFootnote).
    var t = who + " of world cumulative " +
            MEASURE_TITLES[state.measure] + ", " + periodText();
    el.title.textContent = t;
  }

  function renderLegend(rows) {
    var activeLegends = {};
    rows.forEach(function (r) {
      r.points.forEach(function (p) { activeLegends[SOURCE_META[p.source].legend] = true; });
    });
    el.legend.textContent = "";
    LEGEND_ORDER.forEach(function (name) {
      var color = null;
      Object.keys(SOURCE_META).forEach(function (k) {
        if (SOURCE_META[k].legend === name) color = SOURCE_META[k].color;
      });
      var item = document.createElement("span");
      item.className = "legend-item" + (activeLegends[name] ? "" : " inactive");
      if (!activeLegends[name]) item.title = "No data for the selected year range";
      var sw = document.createElement("span");
      sw.className = "legend-swatch";
      sw.style.background = color;
      item.appendChild(sw);
      item.appendChild(document.createTextNode(name));
      el.legend.appendChild(item);
    });
    // GMST entry
    var gmstItem = document.createElement("span");
    gmstItem.className = "legend-item" + (gmstActive() ? "" : " inactive");
    if (!gmstActive()) gmstItem.title = "No data for the selected year range";
    var gsw = document.createElement("span");
    gsw.className = "legend-swatch gmst";
    gmstItem.appendChild(gsw);
    gmstItem.appendChild(document.createTextNode("Contribution to GMST rise"));
    el.legend.appendChild(gmstItem);
  }

  // The viewBox is set to the measured container width so the SVG renders 1:1 and
  // its text keeps its true pixel size. Drawing at a fixed 860 and letting the SVG
  // scale down shrank the labels to about 5px at 320px wide.
  var lastChartWidth = null;
  function chartWidth() {
    return Math.max(320, Math.round((el.wrap && el.wrap.clientWidth) || 860));
  }

  function renderChart(rows) {
    var W = chartWidth();
    lastChartWidth = W;

    // Below ~560px the left gutter and row pitch have to give way to the plot.
    var narrow = W < 560;
    var margin = {
      top: 16,
      right: narrow ? 16 : 24,
      bottom: narrow ? 46 : 52,
      left: narrow ? 84 : 128
    };
    var rowH = narrow ? 76 : 96;
    var tickStep = narrow ? 25 : 10;
    var labelMax = narrow ? 10 : 16;
    var H = margin.top + rowH * rows.length + margin.bottom;
    var plotW = W - margin.left - margin.right;

    var svg = el.svg;
    svg.setAttribute("viewBox", "0 0 " + W + " " + H);
    svg.textContent = "";

    function xPos(v) { return margin.left + v / 100 * plotW; }

    // gridlines + x labels
    for (var g = 0; g <= 100; g += tickStep) {
      var gx = xPos(g);
      svg.appendChild(svgEl("line", {
        class: "grid-line", x1: gx, x2: gx,
        y1: margin.top, y2: margin.top + rowH * rows.length
      }));
      var lbl = svgEl("text", {
        class: "axis-label", x: gx, y: margin.top + rowH * rows.length + 18,
        "text-anchor": "middle"
      });
      lbl.textContent = g;
      svg.appendChild(lbl);
    }
    var axisTitle = svgEl("text", {
      class: "axis-label", x: margin.left + plotW / 2,
      y: margin.top + rowH * rows.length + 40, "text-anchor": "middle"
    });
    axisTitle.textContent = "Share of world total over selected period (%)";
    svg.appendChild(axisTitle);

    rows.forEach(function (row, i) {
      var cy = margin.top + rowH * i + rowH / 2;

      var labelText = row.label.length > labelMax ?
        row.label.slice(0, labelMax - 1) + "…" : row.label;
      var rowLbl = svgEl("text", {
        class: "row-label", x: margin.left - 14, y: cy + 5, "text-anchor": "end"
      });
      rowLbl.textContent = labelText;
      if (labelText !== row.label) {
        var tEl = svgEl("title", {});
        tEl.textContent = row.label;
        rowLbl.appendChild(tEl);
      }
      svg.appendChild(rowLbl);

      if (row.custom) {
        var rm = svgEl("text", {
          class: "row-remove", x: margin.left - 14, y: cy + 22,
          "text-anchor": "end", role: "button", tabindex: 0,
          "aria-label": "Remove " + row.label
        });
        rm.textContent = "remove ×";
        function removeRow() {
          state.customGroups = state.customGroups.filter(function (g) { return g.id !== row.id; });
          track("filter", "remove_country_group");
          render();
        }
        rm.addEventListener("click", removeRow);
        rm.addEventListener("keydown", function (e) {
          if (e.key === "Enter" || e.key === " ") { e.preventDefault(); removeRow(); }
        });
        svg.appendChild(rm);
      }

      if (row.loading) {
        var loadLbl = svgEl("text", {
          class: "axis-label", x: margin.left + plotW / 2, y: cy + 4, "text-anchor": "middle"
        });
        loadLbl.textContent = "Loading country data…";
        svg.appendChild(loadLbl);
        return;
      }

      // connector spans the source dots only (GMST is a reference mark)
      var vals = row.points.map(function (p) { return p.value; });
      if (vals.length > 1) {
        svg.appendChild(svgEl("line", {
          class: "connector",
          x1: xPos(Math.min.apply(null, vals)), x2: xPos(Math.max.apply(null, vals)),
          y1: cy, y2: cy
        }));
      }

      // All marks on this row, for cluster-aware tooltips when dots overlap
      var marks = row.points.map(function (p) {
        return {
          value: p.value,
          label: SOURCE_META[p.source].label,
          color: SOURCE_META[p.source].color,
          notes: p.notes,
          isGmst: false
        };
      });
      if (row.gmst !== null) {
        marks.push({
          value: row.gmst,
          label: "Contribution to GMST rise",
          color: GMST_COLOR,
          notes: row.gmstNotes,
          isGmst: true
        });
      }

      // dots first, then the GMST cross on top so it stays visible when values coincide
      marks.forEach(function (m) {
        if (m.isGmst) return;
        svg.appendChild(svgEl("circle", {
          class: "dot", cx: xPos(m.value), cy: cy, r: 6.5, fill: m.color
        }));
      });
      marks.forEach(function (m) {
        if (!m.isGmst) return;
        var gx2 = xPos(m.value), arm = 7;
        svg.appendChild(svgEl("line", { class: "gmst-mark", x1: gx2 - arm, x2: gx2 + arm, y1: cy, y2: cy }));
        svg.appendChild(svgEl("line", { class: "gmst-mark", x1: gx2, x2: gx2, y1: cy - arm, y2: cy + arm }));
      });

      // Hit targets on top; hovering one mark reads out every mark within 0.75pp.
      // Each target is a tall band reaching halfway to its neighbours (capped),
      // so the pointer only has to be near a mark rather than on it. Bands abut
      // rather than overlap, so bands around near-coincident marks are thin —
      // harmless, because those marks are one cluster with one shared tooltip
      // and the pair still covers the full width between them.
      var CLUSTER_PP = 0.75;
      var HIT_H = Math.min(rowH - 12, 64);
      var HIT_MAX_HALF = 30;
      var ordered = marks.slice().sort(function (a, b) { return a.value - b.value; });
      ordered.forEach(function (m, idx) {
        var x = xPos(m.value);
        function half(neighbour) {
          if (!neighbour) return HIT_MAX_HALF;
          return Math.min(HIT_MAX_HALF, Math.abs(xPos(neighbour.value) - x) / 2);
        }
        var left = half(ordered[idx - 1]), right = half(ordered[idx + 1]);
        var hit = svgEl("rect", {
          class: "dot-hit", x: x - left, y: cy - HIT_H / 2,
          width: left + right, height: HIT_H, tabindex: 0
        });
        hit.setAttribute("aria-label", m.label + ", " + row.label + ": " + fmt(m.value));
        function over() {
          var cluster = marks.filter(function (o) {
            return Math.abs(o.value - m.value) <= CLUSTER_PP;
          });
          var entries = cluster.map(function (o) {
            return { valueText: fmt(o.value), labelText: o.label, color: o.color };
          });
          var notes = [row.label];
          cluster.forEach(function (o) {
            o.notes.forEach(function (n) {
              var line = (cluster.length > 1 ? o.label + ": " : "") + n;
              if (notes.indexOf(line) === -1) notes.push(line);
            });
          });
          // Anchor on the mark itself, not the (possibly clipped) hit band.
          var box = svg.getBoundingClientRect();
          var scale = box.width / W;
          showTooltip(box.left + x * scale, box.top + cy * scale, entries, notes);
        }
        hit.addEventListener("pointerenter", over);
        hit.addEventListener("focus", over);
        hit.addEventListener("pointerleave", hideTooltip);
        hit.addEventListener("blur", hideTooltip);
        svg.appendChild(hit);
      });
    });
  }

  function renderFootnote(rows) {
    var notes = [];
    notes.push("Each dot is one dataset's estimate of the group's share of world cumulative emissions over " +
      periodText() + "; the + marks the group's estimated contribution to the rise in global mean surface temperature.");
    notes.push("Annex I: industrialised countries and economies in transition under the UNFCCC. Annex II: the OECD-member subset of Annex I with climate-finance obligations.");
    if (state.attribution === "colonial") {
      notes.push("Colonial attribution reassigns emissions of territories under colonial rule (1850–2023, Carbon Brief territorial-rule database) to the ruling power.");
    }
    var truncated = {};
    rows.forEach(function (r) {
      r.points.forEach(function (p) {
        p.notes.forEach(function (n) {
          if (n.indexOf("selected countries") === -1) {
            truncated[SOURCE_META[p.source].label + ": " + n] = true;
          }
        });
      });
    });
    Object.keys(truncated).forEach(function (t) { notes.push(t); });
    notes.push("Datasets that do not cover the selected year range are hidden automatically.");
    if (state.customGroups.length) {
      notes.push("Custom groups sum only the countries each dataset reports, so datasets with " +
        "incomplete coverage of the selection understate its share — hover a dot to see how many " +
        "of the selected countries it includes. Dissolved states and small territories are the " +
        "usual gaps, and datasets can also differ in how they assign the history of former " +
        "states (e.g. the USSR) to today's countries.");
    }
    el.footnote.textContent = notes.join(" ");
  }

  function renderTable(rows) {
    rows = rows.filter(function (r) { return !r.loading; });
    var tbl = el.table;
    tbl.textContent = "";
    var thead = document.createElement("thead");
    var hr = document.createElement("tr");
    ["Dataset"].concat(rows.map(function (r) { return r.label; })).forEach(function (h) {
      var th = document.createElement("th");
      th.textContent = h;
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    tbl.appendChild(thead);

    var tbody = document.createElement("tbody");
    var names = activeSources();
    names.forEach(function (name) {
      var tr = document.createElement("tr");
      var td0 = document.createElement("td");
      td0.textContent = SOURCE_META[name].label;
      tr.appendChild(td0);
      rows.forEach(function (r) {
        var p = null;
        r.points.forEach(function (q) { if (q.source === name) p = q; });
        var td = document.createElement("td");
        td.textContent = p ? fmt(p.value) : "—";
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    if (gmstActive()) {
      var tr = document.createElement("tr");
      var td0 = document.createElement("td");
      td0.textContent = "Contribution to GMST rise";
      tr.appendChild(td0);
      rows.forEach(function (r) {
        var td = document.createElement("td");
        td.textContent = fmt(r.gmst);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);
  }

  function render() {
    updateSliderVisuals();
    picker.btn.disabled = state.customGroups.length >= MAX_CUSTOM_GROUPS;
    picker.btn.title = picker.btn.disabled ?
      "Maximum of " + MAX_CUSTOM_GROUPS + " custom groups — remove one to add another" : "";
    var rows = currentValues();
    renderGroupList();
    renderTitle();
    renderLegend(rows);
    renderChart(rows);
    renderFootnote(rows);
    renderTable(rows);
  }

  // ---- init -----------------------------------------------------------------

  // Optional URL params set the initial view, e.g. ?measure=ghg_incLUC&colonial=1&start=1900&end=2000
  // (?custom=0 hides the country picker, for blog embeds). A page can also define
  // window.EMBED_DEFAULTS = {measure, colonial, start, end, custom, customGroups: [{name, isos}]}
  // before app.js loads.
  (function applyParams() {
    var d = window.EMBED_DEFAULTS || {};
    if (MEASURE_TITLES[d.measure]) state.measure = d.measure;
    if (d.colonial) state.attribution = "colonial";
    if (typeof d.start === "number") state.start = d.start;
    if (typeof d.end === "number") state.end = d.end;
    if (Array.isArray(d.customGroups)) {
      d.customGroups.slice(0, MAX_CUSTOM_GROUPS).forEach(function (g) {
        if (g && Array.isArray(g.isos) && g.isos.length) {
          state.customGroups.push({
            id: customIdSeq++,
            name: String(g.name || "Custom group " + customIdSeq),
            isos: g.isos.map(String)
          });
        }
      });
    }
    var q = new URLSearchParams(window.location.search);
    if (d.custom === false || q.get("custom") === "0") {
      picker.section.hidden = true;
    }
    if (MEASURE_TITLES[q.get("measure")]) state.measure = q.get("measure");
    if (q.get("colonial") === "1") state.attribution = "colonial";
    if (!isNaN(parseInt(q.get("start"), 10))) state.start = parseInt(q.get("start"), 10);
    if (!isNaN(parseInt(q.get("end"), 10))) state.end = parseInt(q.get("end"), 10);
    if (state.end < state.start) state.end = state.start;
    el.measure.value = state.measure;
    el.colonial.checked = state.attribution === "colonial";
  })();

  el.tableDetails.addEventListener("toggle", function () {
    track(el.tableDetails.open ? "detail_open" : "detail_close", "data_table");
  });

  setSliderBounds(state.attribution);
  render();

  // Redraw when the container width changes (host page resize, or the iframe being
  // re-sized by the parent). Width only: renderChart changes the SVG's height, so
  // reacting to height would feed back into itself.
  if (window.ResizeObserver && el.wrap) {
    var widthTimer = null;
    new ResizeObserver(function () {
      if (chartWidth() === lastChartWidth) return;
      clearTimeout(widthTimer);
      widthTimer = setTimeout(render, 80);
    }).observe(el.wrap);
  }

  if (window.CGDTracking) window.CGDTracking.trackView();
})();
