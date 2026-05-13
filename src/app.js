import dashboardData from "./dashboard-data.js";
import { setupLiveExcel } from "./live-excel.js";

let data = dashboardData;

const state = {
  campus: "All Campuses",
  metric: "attendance",
  showEvents: true,
  bigFiveEvent: null,
};

const metricOrder = [
  "attendance",
  "kids",
  "growthTrack",
  "salvations",
  "firstTimers",
  "dreamTeam",
];

let metricLabels = getMetricLabels();

const els = {
  latestDate: document.querySelector("#latestDate"),
  sourceName: document.querySelector("#sourceName"),
  liveStatus: document.querySelector("#liveStatus"),
  connectExcelButton: document.querySelector("#connectExcelButton"),
  refreshExcelButton: document.querySelector("#refreshExcelButton"),
  campusSelect: document.querySelector("#campusSelect"),
  metricTabs: document.querySelector("#metricTabs"),
  eventToggle: document.querySelector("#eventToggle"),
  kpis: document.querySelector("#kpis"),
  trendTitle: document.querySelector("#trendTitle"),
  trendMeta: document.querySelector("#trendMeta"),
  lineChart: document.querySelector("#lineChart"),
  barTitle: document.querySelector("#barTitle"),
  barMeta: document.querySelector("#barMeta"),
  barChart: document.querySelector("#barChart"),
  volatilityChart: document.querySelector("#volatilityChart"),
  insights: document.querySelector("#insights"),
  campusTable: document.querySelector("#campusTable"),
  bigFiveEventSelect: document.querySelector("#bigFiveEventSelect"),
  bigFiveKpis: document.querySelector("#bigFiveKpis"),
  bigFiveTrendTitle: document.querySelector("#bigFiveTrendTitle"),
  bigFiveTrendMeta: document.querySelector("#bigFiveTrendMeta"),
  bigFiveYearChart: document.querySelector("#bigFiveYearChart"),
  bigFivePhaseTitle: document.querySelector("#bigFivePhaseTitle"),
  bigFivePhaseMeta: document.querySelector("#bigFivePhaseMeta"),
  bigFivePhaseChart: document.querySelector("#bigFivePhaseChart"),
  bigFiveInsights: document.querySelector("#bigFiveInsights"),
  bigFiveTable: document.querySelector("#bigFiveTable"),
};

let liveExcel;

function getMetricLabels() {
  return Object.fromEntries(
    metricOrder.map((key) => [key, data.metrics[key]?.label || key]),
  );
}

function parseIsoDate(iso) {
  const [year, month, day] = iso.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function formatDate(iso) {
  return parseIsoDate(iso).toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  });
}

function shortDate(iso) {
  return parseIsoDate(iso).toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
  });
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  return Math.round(value).toLocaleString("en-US");
}

function formatPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  const sign = value > 0 ? "+" : "";
  return `${sign}${Number(value).toFixed(1)}%`;
}

function toneClass(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "neutral";
  if (value > 0) return "positive";
  if (value < 0) return "negative";
  return "neutral";
}

function pctChange(current, previous) {
  if (!previous || current === null || current === undefined) return null;
  return ((current - previous) / previous) * 100;
}

function phaseOrder(phase) {
  const parsed = Number(String(phase || "").split(" ")[0]);
  return Number.isFinite(parsed) ? parsed : 99;
}

function defaultBigFiveEvent() {
  const latest = data.bigFive.eventYears
    .filter((row) => row.campaignTotal > 0)
    .sort((a, b) => a.endDate.localeCompare(b.endDate))
    .at(-1);
  return latest?.event || data.bigFive.events[0]?.event;
}

function setupControls() {
  state.bigFiveEvent = defaultBigFiveEvent();
  updateSourceLabels();
  if (!hasDashboardData()) {
    setLiveStatus("Sign in with Microsoft to load live dashboard data.");
  }

  syncCampusOptions();

  for (const key of metricOrder) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.metric = key;
    button.textContent = metricLabels[key];
    button.setAttribute("role", "tab");
    button.addEventListener("click", () => {
      state.metric = key;
      updateDashboard();
    });
    els.metricTabs.append(button);
  }

  els.campusSelect.addEventListener("change", (event) => {
    state.campus = event.target.value;
    updateDashboard();
  });

  els.eventToggle.addEventListener("change", (event) => {
    state.showEvents = event.target.checked;
    updateDashboard();
  });

  syncBigFiveOptions();
  els.bigFiveEventSelect.addEventListener("change", (event) => {
    state.bigFiveEvent = event.target.value;
    updateDashboard();
  });

  liveExcel = setupLiveExcel({
    onData: applyDashboardData,
    onStatus: setLiveStatus,
    onReady: (ready) => {
      els.refreshExcelButton.disabled = !ready;
    },
  });
  if (!liveExcel) {
    els.connectExcelButton.disabled = true;
    els.refreshExcelButton.disabled = true;
  }

  els.connectExcelButton.addEventListener("click", async () => {
    await runLiveAction(() => liveExcel?.connect());
  });
  els.refreshExcelButton.addEventListener("click", async () => {
    await runLiveAction(() => liveExcel?.refresh());
  });
}

function syncCampusOptions() {
  const current = state.campus;
  els.campusSelect.innerHTML = "";
  const options = ["All Campuses", ...data.campuses];
  for (const campus of options) {
    const option = document.createElement("option");
    option.value = campus;
    option.textContent = campus;
    els.campusSelect.append(option);
  }
  state.campus = options.includes(current) ? current : "All Campuses";
  els.campusSelect.value = state.campus;
}

function syncBigFiveOptions() {
  const current = state.bigFiveEvent;
  els.bigFiveEventSelect.innerHTML = "";
  for (const event of data.bigFive.events) {
    const option = document.createElement("option");
    option.value = event.event;
    option.textContent = event.label;
    els.bigFiveEventSelect.append(option);
  }
  const eventKeys = data.bigFive.events.map((event) => event.event);
  state.bigFiveEvent = eventKeys.includes(current) ? current : defaultBigFiveEvent();
  els.bigFiveEventSelect.value = state.bigFiveEvent;
}

function updateSourceLabels() {
  els.latestDate.textContent = data.source.latestAttendanceDate
    ? `Latest: ${formatDate(data.source.latestAttendanceDate)}`
    : "Latest: --";
  els.sourceName.textContent = data.source.workbook || "Source workbook";
}

function hasDashboardData() {
  return data.campuses.length > 0 || data.totals.length > 0 || data.bigFive.rows.length > 0;
}

function setLiveStatus(message, tone = "info") {
  els.liveStatus.textContent = message;
  els.liveStatus.className = tone === "error" ? "negative" : "neutral";
}

async function runLiveAction(action) {
  if (!action) return;
  els.connectExcelButton.disabled = true;
  els.refreshExcelButton.disabled = true;
  try {
    await action();
    els.refreshExcelButton.disabled = false;
  } catch {
    els.connectExcelButton.disabled = false;
  } finally {
    els.connectExcelButton.disabled = false;
  }
}

function applyDashboardData(nextData) {
  data = nextData;
  metricLabels = getMetricLabels();
  updateSourceLabels();
  syncCampusOptions();
  syncBigFiveOptions();
  updateDashboard();
}

function getMetricPoints(metricKey, campus) {
  const metric = data.metrics[metricKey];
  if (!metric) return [];

  if (campus !== "All Campuses") {
    return (metric.series[campus] || [])
      .filter((point) => point.isSunday && point.value > 0)
      .filter((point) => state.showEvents || point.eventType === "normal");
  }

  const byDate = new Map();
  for (const points of Object.values(metric.series)) {
    for (const point of points) {
      if (!point.isSunday || point.value <= 0) continue;
      if (!state.showEvents && point.eventType !== "normal") continue;
      if (!byDate.has(point.date)) {
        byDate.set(point.date, {
          date: point.date,
          value: 0,
          event: point.event,
          eventType: point.eventType,
          isSunday: true,
        });
      }
      const bucket = byDate.get(point.date);
      bucket.value += point.value;
      if (!bucket.event && point.event) bucket.event = point.event;
      if (point.eventType !== "normal") bucket.eventType = point.eventType;
    }
  }

  return Array.from(byDate.values()).sort((a, b) => a.date.localeCompare(b.date));
}

function latestByCampus(metricKey) {
  const metric = data.metrics[metricKey];
  const rows = [];
  for (const campus of data.campuses) {
    const points = (metric.series[campus] || [])
      .filter((point) => point.isSunday && point.value > 0)
      .filter((point) => state.showEvents || point.eventType === "normal");
    const latest = points.at(-1);
    if (latest) {
      rows.push({ campus, value: latest.value, date: latest.date });
    }
  }
  return rows.sort((a, b) => b.value - a.value);
}

function statForCampus(campus) {
  return data.campusStats.find((item) => item.campus === campus);
}

function totalsDiagnostics(points) {
  const latest = points.at(-1);
  const previous = points.at(-2);
  const weekChange = latest && previous ? pctChange(latest.value, previous.value) : null;
  const previousFour = points
    .slice(0, -1)
    .filter((point) => point.eventType !== "holiday" && point.eventType !== "weather")
    .slice(-4);
  const baseline =
    previousFour.length > 0
      ? previousFour.reduce((sum, point) => sum + point.value, 0) / previousFour.length
      : null;
  const vsBaseline = latest && baseline ? pctChange(latest.value, baseline) : null;
  return { latest, previous, weekChange, baseline, vsBaseline };
}

function renderKpis(points) {
  const latest = points.at(-1);
  const previous = points.at(-2);
  const weekChange = latest && previous ? pctChange(latest.value, previous.value) : null;
  const focusedStat = state.campus !== "All Campuses" ? statForCampus(state.campus) : null;
  const totalDiag = totalsDiagnostics(points);

  const swingCount =
    state.campus === "All Campuses"
      ? data.volatility.reduce((sum, campus) => sum + campus.swing10Count, 0)
      : focusedStat?.swing10Count;

  const baselineDelta = totalDiag.vsBaseline;

  const yoy =
    state.metric !== "attendance"
      ? null
      : state.campus === "All Campuses"
        ? pctChange(latest?.value, data.yoy?.total)
        : pctChange(latest?.value, data.yoy?.byCampus?.[state.campus]);

  const latestLabel =
    state.campus === "All Campuses"
      ? `${metricLabels[state.metric]} total`
      : `${state.campus} ${metricLabels[state.metric]}`;

  const kpis = [
    {
      label: latestLabel,
      value: formatNumber(latest?.value),
      note: latest ? `Latest reported Sunday, ${shortDate(latest.date)}` : "No data",
    },
    {
      label: "Week over week",
      value: formatPct(weekChange),
      valueClass: toneClass(weekChange),
      note: previous ? `Compared with ${shortDate(previous.date)}` : "No comparison",
    },
    {
      label: "Attendance 10%+ swings",
      value: formatNumber(swingCount),
      note:
        state.campus === "All Campuses"
          ? "Campus-level Sunday swings YTD"
          : "Sunday swings YTD",
    },
    {
      label: "Vs recent baseline",
      value: formatPct(baselineDelta),
      valueClass: toneClass(baselineDelta),
      note: "Excludes holiday and weather weeks",
    },
    {
      label: "Comparable YoY",
      value: formatPct(yoy),
      valueClass: toneClass(yoy),
      note:
        state.metric === "attendance" && data.yoy
          ? `Against ${shortDate(data.yoy.matchedDate)}`
          : "Attendance only",
    },
  ];

  els.kpis.innerHTML = kpis
    .map(
      (kpi) => `
        <article class="kpi">
          <div class="kpi-label">${kpi.label}</div>
          <div class="kpi-value ${kpi.valueClass || ""}">${kpi.value}</div>
          <div class="kpi-note">${kpi.note}</div>
        </article>
      `,
    )
    .join("");
}

function renderLineChart(points) {
  if (!points.length) {
    els.lineChart.innerHTML = `<div class="empty">No points available.</div>`;
    return;
  }

  const width = 900;
  const height = 320;
  const margin = { top: 22, right: 28, bottom: 48, left: 58 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const values = points.map((point) => point.value);
  const maxValue = Math.max(...values);
  const minValue = Math.min(...values);
  const yMin = Math.max(0, minValue - (maxValue - minValue) * 0.25);
  const yMax = maxValue + Math.max(10, (maxValue - minValue) * 0.18);
  const x = (index) =>
    margin.left + (points.length === 1 ? innerWidth / 2 : (index / (points.length - 1)) * innerWidth);
  const y = (value) => margin.top + ((yMax - value) / (yMax - yMin)) * innerHeight;
  const linePath = points.map((point, index) => `${index === 0 ? "M" : "L"} ${x(index)} ${y(point.value)}`).join(" ");
  const areaPath = `${linePath} L ${x(points.length - 1)} ${margin.top + innerHeight} L ${x(0)} ${
    margin.top + innerHeight
  } Z`;
  const gridValues = Array.from({ length: 5 }, (_, index) => yMin + ((yMax - yMin) / 4) * index);
  const labelIndexes = new Set([0, Math.floor((points.length - 1) / 2), points.length - 1]);

  const svg = `
    <svg viewBox="0 0 ${width} ${height}" role="img">
      ${gridValues
        .map((value) => {
          const yy = y(value);
          return `
            <line class="grid-line" x1="${margin.left}" y1="${yy}" x2="${width - margin.right}" y2="${yy}"></line>
            <text class="axis-label" x="${margin.left - 10}" y="${yy + 4}" text-anchor="end">${formatNumber(value)}</text>
          `;
        })
        .join("")}
      <path class="series-area" d="${areaPath}"></path>
      <path class="series-line" d="${linePath}"></path>
      ${points
        .map((point, index) => {
          const isEvent = point.eventType !== "normal";
          return `
            <circle class="point ${isEvent ? "event" : ""}" cx="${x(index)}" cy="${y(point.value)}" r="${
              isEvent ? 5 : 4
            }">
              <title>${shortDate(point.date)}: ${formatNumber(point.value)}${
                point.event ? `, ${point.event}` : ""
              }</title>
            </circle>
          `;
        })
        .join("")}
      ${points
        .map((point, index) =>
          labelIndexes.has(index)
            ? `<text class="axis-label" x="${x(index)}" y="${height - 14}" text-anchor="${
                index === 0 ? "start" : index === points.length - 1 ? "end" : "middle"
              }">${shortDate(point.date)}</text>`
            : "",
        )
        .join("")}
    </svg>
  `;

  els.lineChart.innerHTML = svg;
}

function renderBars() {
  const rows = latestByCampus(state.metric);
  const max = Math.max(...rows.map((row) => row.value), 1);
  const latestDate = rows[0]?.date;

  els.barTitle.textContent = `${metricLabels[state.metric]} by Campus`;
  els.barMeta.textContent = latestDate ? shortDate(latestDate) : "";
  els.barChart.innerHTML = rows
    .map((row) => {
      const width = Math.max(3, (row.value / max) * 100);
      const selected = state.campus === row.campus;
      return `
        <div class="bar-row ${selected ? "active" : ""}">
          <div class="bar-label" title="${row.campus}">${row.campus}</div>
          <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
          <div class="bar-value">${formatNumber(row.value)}</div>
        </div>
      `;
    })
    .join("");
}

function renderVolatility() {
  const max = Math.max(...data.volatility.map((row) => row.swing10Count), 1);
  els.volatilityChart.innerHTML = data.volatility
    .map((row) => {
      const width = Math.max(4, (row.swing10Count / max) * 100);
      return `
        <div class="bar-row">
          <div class="bar-label" title="${row.campus}">${row.campus}</div>
          <div class="bar-track"><div class="bar-fill vol-fill" style="width:${width}%"></div></div>
          <div class="bar-value">${row.swing10Count} / ${row.avgAbsSwing10Pct}%</div>
        </div>
      `;
    })
    .join("");
}

function selectedCampusInsights() {
  const stat = statForCampus(state.campus);
  if (!stat) return [];
  const maxSwing = stat.maxSwing;
  const weekTone = Math.abs(stat.weekChangePct || 0) >= 10 ? "warning" : "info";
  const baselineTone = (stat.vsPrevious4Pct || 0) < -5 ? "warning" : "info";

  return [
    {
      title: `${state.campus} latest movement`,
      body: `${state.campus} reported ${formatNumber(stat.latest)} on ${shortDate(
        stat.latestDate,
      )}, ${formatPct(stat.weekChangePct)} versus ${shortDate(stat.previousDate)}.`,
      severity: weekTone,
    },
    {
      title: `${state.campus} has ${stat.swing10Count} flagged swings`,
      body: `Its 10%+ Sunday-to-Sunday swings average ${stat.avgAbsSwing10Pct}% across flagged weeks.`,
      severity: stat.swing10Count >= 7 ? "warning" : "info",
    },
    {
      title: "Recent baseline check",
      body: `${state.campus} is ${formatPct(stat.vsPrevious4Pct)} against its previous four normal Sundays.`,
      severity: baselineTone,
    },
    {
      title: "Largest movement",
      body: maxSwing
        ? `${formatPct(maxSwing.pct)} from ${shortDate(maxSwing.fromDate)} to ${shortDate(maxSwing.toDate)}${
            maxSwing.event ? ` during ${maxSwing.event}` : ""
          }.`
        : "No comparable Sunday swings yet.",
      severity: "info",
    },
  ];
}

function renderInsights() {
  const insights =
    state.campus === "All Campuses"
      ? data.insights.map((insight) => ({
          title: insight.title,
          body: insight.body,
          severity: insight.severity,
        }))
      : selectedCampusInsights();

  els.insights.innerHTML = insights
    .map(
      (insight) => `
        <article class="insight ${insight.severity || "info"}">
          <h3 class="insight-title">${insight.title}</h3>
          <p class="insight-body">${insight.body}</p>
        </article>
      `,
    )
    .join("");
}

function renderTable() {
  els.campusTable.innerHTML = data.campusStats
    .map(
      (row) => `
        <tr data-campus="${row.campus}" class="${row.campus === state.campus ? "active" : ""}">
          <td>${row.campus}</td>
          <td>${formatNumber(row.latest)}</td>
          <td class="${toneClass(row.weekChangePct)}">${formatPct(row.weekChangePct)}</td>
          <td class="${toneClass(row.vsPrevious4Pct)}">${formatPct(row.vsPrevious4Pct)}</td>
          <td>${row.swing10Count}</td>
          <td>${row.avgAbsSwingPct}%</td>
        </tr>
      `,
    )
    .join("");

  for (const row of els.campusTable.querySelectorAll("tr")) {
    row.addEventListener("click", () => {
      state.campus = row.dataset.campus;
      els.campusSelect.value = state.campus;
      updateDashboard();
    });
  }
}

function summarizeBigFiveEvent() {
  const rows = data.bigFive.rows.filter(
    (row) =>
      row.event === state.bigFiveEvent &&
      (state.campus === "All Campuses" || row.campus === state.campus),
  );
  const byYear = new Map();

  for (const row of rows) {
    if (!byYear.has(row.year)) {
      byYear.set(row.year, {
        event: row.event,
        eventLabel: row.eventLabel,
        year: row.year,
        startDate: row.date,
        endDate: row.date,
        campaignTotal: 0,
        phaseTotals: new Map(),
        campusTotals: new Map(),
      });
    }
    const bucket = byYear.get(row.year);
    bucket.startDate = row.date < bucket.startDate ? row.date : bucket.startDate;
    bucket.endDate = row.date > bucket.endDate ? row.date : bucket.endDate;
    bucket.campaignTotal += row.attendance;
    bucket.phaseTotals.set(row.phase, {
      phase: row.phase,
      phaseLabel: row.phaseLabel,
      total: (bucket.phaseTotals.get(row.phase)?.total || 0) + row.attendance,
    });
    bucket.campusTotals.set(row.campus, {
      campus: row.campus,
      total: (bucket.campusTotals.get(row.campus)?.total || 0) + row.attendance,
    });
  }

  const records = Array.from(byYear.values())
    .map((record) => {
      const phaseTotals = Array.from(record.phaseTotals.values()).sort(
        (a, b) => phaseOrder(a.phase) - phaseOrder(b.phase),
      );
      const campusTotals = Array.from(record.campusTotals.values()).sort((a, b) => b.total - a.total);
      const totalPhase = phaseTotals.find((phase) => phase.phase === "07 Total" && phase.total > 0);
      const featuredPhase =
        totalPhase ||
        phaseTotals
          .filter(
            (phase) =>
              phase.total > 0 &&
              phase.phase !== "01 Pre Event" &&
              !phase.phaseLabel.toLowerCase().startsWith("post"),
          )
          .sort((a, b) => b.total - a.total)[0];
      const pre = phaseTotals.find((phase) => phase.phase === "01 Pre Event");
      const postWeek1 = phaseTotals.find((phase) => phase.phase === "08 Post Week 1");
      const liftPct = featuredPhase && pre?.total ? pctChange(featuredPhase.total, pre.total) : null;
      const postRetentionPct =
        featuredPhase && postWeek1?.total ? (postWeek1.total / featuredPhase.total) * 100 : null;

      return {
        ...record,
        phaseTotals,
        campusTotals,
        preTotal: pre?.total ?? null,
        featuredPhase: featuredPhase?.phase ?? null,
        featuredPhaseLabel: featuredPhase?.phaseLabel ?? null,
        featuredTotal: featuredPhase?.total ?? null,
        postWeek1: postWeek1?.total ?? null,
        liftPct,
        postRetentionPct,
      };
    })
    .sort((a, b) => a.year - b.year);

  const lookup = new Map(records.map((record) => [record.year, record]));
  for (const record of records) {
    const previous = lookup.get(record.year - 1);
    record.yoyCampaignPct = previous ? pctChange(record.campaignTotal, previous.campaignTotal) : null;
    record.yoyFeaturedPct =
      previous?.featuredTotal && record.featuredTotal
        ? pctChange(record.featuredTotal, previous.featuredTotal)
        : null;
  }

  return records;
}

function renderBarList(container, rows, options = {}) {
  const max = Math.max(...rows.map((row) => row.value), 1);
  container.innerHTML = rows
    .map((row) => {
      const width = Math.max(3, (row.value / max) * 100);
      return `
        <div class="bar-row">
          <div class="bar-label" title="${row.label}">${row.label}</div>
          <div class="bar-track"><div class="bar-fill ${options.fillClass || ""}" style="width:${width}%"></div></div>
          <div class="bar-value">${options.format ? options.format(row.value) : formatNumber(row.value)}</div>
        </div>
      `;
    })
    .join("");
}

function bigFiveEventLabel() {
  return data.bigFive.events.find((event) => event.event === state.bigFiveEvent)?.label || "Big 5";
}

function renderBigFiveKpis(records) {
  const latest = records.filter((record) => record.campaignTotal > 0).at(-1);
  const kpis = [
    {
      label: "Campaign total",
      value: formatNumber(latest?.campaignTotal),
      note: latest ? `${latest.year} ${bigFiveEventLabel()}` : "No event data",
    },
    {
      label: "Featured total",
      value: formatNumber(latest?.featuredTotal),
      note: latest?.featuredPhaseLabel || "No featured phase",
    },
    {
      label: "Campaign YoY",
      value: formatPct(latest?.yoyCampaignPct),
      valueClass: toneClass(latest?.yoyCampaignPct),
      note: "Versus prior year",
    },
    {
      label: "Event lift",
      value: formatPct(latest?.liftPct),
      valueClass: toneClass(latest?.liftPct),
      note: "Featured total vs pre-event",
    },
    {
      label: "Post retention",
      value: formatPct(latest?.postRetentionPct),
      valueClass:
        latest?.postRetentionPct !== null && latest?.postRetentionPct !== undefined
          ? latest.postRetentionPct < 50
            ? "negative"
            : "positive"
          : "neutral",
      note: "Post Week 1 vs featured total",
    },
  ];

  els.bigFiveKpis.innerHTML = kpis
    .map(
      (kpi) => `
        <article class="kpi">
          <div class="kpi-label">${kpi.label}</div>
          <div class="kpi-value ${kpi.valueClass || ""}">${kpi.value}</div>
          <div class="kpi-note">${kpi.note}</div>
        </article>
      `,
    )
    .join("");
}

function renderBigFiveCharts(records) {
  const usable = records.filter((record) => record.campaignTotal > 0);
  const latest = usable.at(-1);
  const eventLabel = bigFiveEventLabel();
  els.bigFiveTrendTitle.textContent = `${eventLabel} Campaign Attendance`;
  els.bigFiveTrendMeta.textContent =
    state.campus === "All Campuses" ? "All campuses" : state.campus;

  renderBarList(
    els.bigFiveYearChart,
    usable.map((record) => ({
      label: String(record.year),
      value: record.campaignTotal,
    })),
  );

  els.bigFivePhaseTitle.textContent = latest ? `${latest.year} Phase Breakdown` : "Phase Breakdown";
  els.bigFivePhaseMeta.textContent = latest
    ? `${shortDate(latest.startDate)} - ${shortDate(latest.endDate)}`
    : "";
  renderBarList(
    els.bigFivePhaseChart,
    (latest?.phaseTotals || []).map((phase) => ({
      label: phase.phaseLabel,
      value: phase.total,
    })),
    { fillClass: "vol-fill" },
  );
}

function renderBigFiveInsights(records) {
  const latest = records.filter((record) => record.campaignTotal > 0).at(-1);
  if (!latest) {
    els.bigFiveInsights.innerHTML = "";
    return;
  }

  const bestYear = records.reduce(
    (best, record) => (record.campaignTotal > (best?.campaignTotal || 0) ? record : best),
    null,
  );
  const topCampus = latest.campusTotals[0];
  const biggestPhase = latest.phaseTotals
    .filter((phase) => phase.total > 0)
    .sort((a, b) => b.total - a.total)[0];
  const insights = [
    {
      title: `${latest.year} ${latest.eventLabel} is ${formatPct(latest.yoyCampaignPct)} YoY`,
      body: `${formatNumber(latest.campaignTotal)} campaign attendance across ${latest.phaseTotals.length} tracked phases.`,
      severity: latest.yoyCampaignPct < -10 ? "critical" : "info",
    },
    {
      title: latest.liftPct === null ? "Lift needs a pre-event baseline" : `Event lift is ${formatPct(latest.liftPct)}`,
      body:
        latest.liftPct === null
          ? "This event does not have a usable pre-event comparison for the selected campus/view."
          : `${latest.featuredPhaseLabel} reached ${formatNumber(latest.featuredTotal)} versus ${formatNumber(latest.preTotal)} pre-event.`,
      severity: "info",
    },
    {
      title:
        latest.postRetentionPct === null
          ? "No post-week retention yet"
          : `Post retention is ${formatPct(latest.postRetentionPct)}`,
      body:
        latest.postRetentionPct === null
          ? "This event does not have Post Week 1 data in the selected view."
          : `Post Week 1 retained ${formatNumber(latest.postWeek1)} of ${formatNumber(latest.featuredTotal)} from the featured total.`,
      severity: latest.postRetentionPct !== null && latest.postRetentionPct < 50 ? "warning" : "info",
    },
    {
      title:
        state.campus === "All Campuses" && topCampus
          ? `${topCampus.campus} is the top contributor`
          : `${biggestPhase?.phaseLabel || "Latest phase"} leads the phase mix`,
      body:
        state.campus === "All Campuses" && topCampus
          ? `${topCampus.campus} contributed ${formatNumber(topCampus.total)} to the ${latest.year} campaign total.`
          : `${biggestPhase?.phaseLabel || "The leading phase"} contributed ${formatNumber(biggestPhase?.total)} in ${latest.year}.`,
      severity: "info",
    },
    {
      title: `Best historical year: ${bestYear?.year}`,
      body: `${bestYear?.eventLabel} peaked at ${formatNumber(bestYear?.campaignTotal)} campaign attendance in this dataset.`,
      severity: latest.year === bestYear?.year ? "info" : "warning",
    },
  ];

  els.bigFiveInsights.innerHTML = insights
    .map(
      (insight) => `
        <article class="insight ${insight.severity || "info"}">
          <h3 class="insight-title">${insight.title}</h3>
          <p class="insight-body">${insight.body}</p>
        </article>
      `,
    )
    .join("");
}

function renderBigFiveTable(records) {
  els.bigFiveTable.innerHTML = records
    .slice()
    .reverse()
    .map(
      (record) => `
        <tr>
          <td>${record.eventLabel}</td>
          <td>${record.year}</td>
          <td>${formatNumber(record.campaignTotal)}</td>
          <td>${formatNumber(record.featuredTotal)}</td>
          <td class="${toneClass(record.yoyCampaignPct)}">${formatPct(record.yoyCampaignPct)}</td>
          <td class="${toneClass(record.liftPct)}">${formatPct(record.liftPct)}</td>
          <td class="${
            record.postRetentionPct !== null && record.postRetentionPct !== undefined
              ? record.postRetentionPct < 50
                ? "negative"
                : "positive"
              : "neutral"
          }">${formatPct(record.postRetentionPct)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderBigFive() {
  const records = summarizeBigFiveEvent();
  els.bigFiveEventSelect.value = state.bigFiveEvent;
  renderBigFiveKpis(records);
  renderBigFiveCharts(records);
  renderBigFiveInsights(records);
  renderBigFiveTable(records);
}

function updateDashboard() {
  for (const button of els.metricTabs.querySelectorAll("button")) {
    const active = button.dataset.metric === state.metric;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  }

  const points = getMetricPoints(state.metric, state.campus);
  const latest = points.at(-1);
  const previous = points.at(-2);
  const weekChange = latest && previous ? pctChange(latest.value, previous.value) : null;

  els.trendTitle.textContent =
    state.campus === "All Campuses"
      ? `${metricLabels[state.metric]} Trend`
      : `${state.campus} ${metricLabels[state.metric]}`;
  els.trendMeta.textContent = latest
    ? `${formatNumber(latest.value)} latest, ${formatPct(weekChange)} WoW`
    : "";

  renderKpis(points);
  renderLineChart(points);
  renderBars();
  renderVolatility();
  renderInsights();
  renderBigFive();
  renderTable();
}

setupControls();
updateDashboard();
