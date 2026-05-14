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
  "bigFive",
  "healthReport",
];

let metricLabels = getMetricLabels();

const healthTargetDefaults = {
  attendance: { label: "Attendance", unit: "count" },
  attendanceYoy: { label: "Attendance vs Last Year", unit: "%", optimalMin: 0, direction: "higher" },
  kidsPct: { label: "Kids % of Attendance", unit: "%", optimalMin: 13, direction: "higher" },
  growthTrackPct: {
    label: "Growth Track %",
    unit: "%",
    optimalMin: 2,
    optimalMax: 4,
    direction: "range",
  },
  baptismPct: { label: "Baptisms %", unit: "%", optimalMin: 0.25, direction: "higher" },
  salvationsPct: { label: "Salvations %", unit: "%", optimalMin: 0.75, direction: "higher" },
  firstTimersPct: {
    label: "First Time Guests %",
    unit: "%",
    optimalMin: 2,
    optimalMax: 4,
    direction: "range",
  },
  dreamTeamPct: { label: "Dream Team %", unit: "%", optimalMin: 33, direction: "higher" },
  groupsPct: {
    label: "Total Groups %",
    unit: "%",
    optimalMin: 8,
    optimalMax: 10,
    direction: "range",
  },
  groupGoalPct: { label: "Group Goal Hit", unit: "%", optimalMin: 100, direction: "higher" },
  groupMembersPct: { label: "Group Members % of Attendance", unit: "%", optimalMin: 100, direction: "higher" },
  groupAttendancePct: { label: "Group Attendance / Signups", unit: "%" },
  heartSoulDirectorPct: { label: "Heart & Soul Directors", unit: "%", optimalMin: 100, direction: "higher" },
  heartSoulCoordinatorPct: {
    label: "Heart & Soul Coordinators",
    unit: "%",
    optimalMin: 100,
    direction: "higher",
  },
  heartSoulTeamLeadPct: { label: "Heart & Soul Team Leads", unit: "%", optimalMin: 100, direction: "higher" },
  leadershipFillPct: { label: "Leadership Fill Rate", unit: "%", optimalMin: 100, direction: "higher" },
  leadershipVacancies: { label: "Leadership Vacancies", unit: "count", optimalMax: 0, direction: "lower" },
};

const healthOptimalLabels = {
  attendance: "Avg",
  attendanceYoy: "Positive",
  groupGoalPct: ">=100%",
  groupAttendancePct: "Higher",
  heartSoulTeamLeadPct: "If included",
  leadershipVacancies: "0",
};

const metricRateTargets = {
  kids: {
    label: "Kids",
    lower: 13,
    type: "higher",
    ministry: "kids ministry",
    emphasis: "Execute 52 Excellent Sundays",
    key: "Excellent Sunday Experience",
  },
  growthTrack: {
    label: "Growth Track",
    lower: 2,
    upper: 4,
    type: "range",
    ministry: "next steps",
    emphasis: "Maximize Next Steps",
    key: "Effective Assimilation Systems",
  },
  salvations: {
    label: "Salvations",
    lower: 0.75,
    type: "higher",
    ministry: "salvation response",
    emphasis: "Maximize Next Steps",
    key: "Growth and invite vision",
  },
  firstTimers: {
    label: "First Time Guests",
    lower: 2,
    upper: 4,
    type: "range",
    ministry: "guest connection",
    emphasis: "Marketing and compelling communication",
    key: "Growth and invite vision",
  },
  dreamTeam: {
    label: "Dream Team",
    lower: 33,
    type: "higher",
    ministry: "serving team",
    emphasis: "Fully Staffed Dream Team",
    key: "Mature and galvanizing leaders",
  },
};

const growthKeys = [
  "Excellent Sunday Experience",
  "Effective Assimilation Systems",
  "Mature and galvanizing leaders",
  "Electric experience in every corner",
  "Growth and invite vision",
];

const growthBarriers = [
  {
    max: 500,
    label: "Under 500",
    issue: "Sundays are not yet excellent across kids, preaching, worship, experience, and location.",
    emphasis: "Execute 52 Excellent Sundays",
  },
  {
    max: 1000,
    label: "Under 1,000",
    issue: "Systems and leadership capacity usually become the growth lid.",
    emphasis: "Fully staffed leaders, Dream Team, groups, and Growth Track",
  },
  {
    max: 1500,
    label: "Under 1,500",
    issue: "Invite culture and compelling communication become the next barrier.",
    emphasis: "Marketing and compelling communication for all events",
  },
  {
    max: 2500,
    label: "Under 2,500",
    issue: "Sundays need to move from excellent to memorable.",
    emphasis: "Electric Rally Nights and excellent every-room experiences",
  },
  {
    max: 5000,
    label: "Under 5,000",
    issue: "Stewardship and high-capacity leaders become the scaling need.",
    emphasis: "Identify and develop high-capacity leaders",
  },
  {
    max: 10000,
    label: "Under 10,000",
    issue: "Preference can supersede growth when people, systems, or culture hit capacity.",
    emphasis: "Build scalable systems that keep growth ahead of preference",
  },
];

const bigFiveGrowthGoals = [
  { campus: "BWI", aliases: ["BWI"], preEasterAvg: 3756, easterTotal: 10374, growthGoal: 5741, barrier: "5K" },
  {
    campus: "Columbia",
    aliases: ["COL", "Columbia"],
    preEasterAvg: 1358,
    easterTotal: 3419,
    growthGoal: 1976,
    barrier: "1,500 - 2K",
  },
  { campus: "UBC", aliases: ["UBC"], preEasterAvg: 693, easterTotal: 1516, growthGoal: 940, barrier: "1K" },
  {
    campus: "Flowers",
    aliases: ["FLO", "Flowers"],
    preEasterAvg: 929,
    easterTotal: 1702,
    growthGoal: 1161,
    barrier: "1K - 1,500",
  },
  {
    campus: "Falls Church",
    aliases: ["FC", "Falls Church"],
    preEasterAvg: 554,
    easterTotal: 1252,
    growthGoal: 763,
    barrier: "1K",
  },
  {
    campus: "Silver Spring",
    aliases: ["SS", "Silver Spring"],
    preEasterAvg: 494,
    easterTotal: 887,
    growthGoal: 612,
    barrier: "500 - 1K",
  },
];

const els = {
  latestDate: document.querySelector("#latestDate"),
  sourceName: document.querySelector("#sourceName"),
  liveStatus: document.querySelector("#liveStatus"),
  connectExcelButton: document.querySelector("#connectExcelButton"),
  refreshExcelButton: document.querySelector("#refreshExcelButton"),
  metricPanels: document.querySelectorAll("[data-metric-panel]"),
  attendanceOnlyPanels: document.querySelectorAll("[data-attendance-only]"),
  campusSelect: document.querySelector("#campusSelect"),
  metricTabs: document.querySelector("#metricTabs"),
  eventToggle: document.querySelector("#eventToggle"),
  eventToggleWrapper: document.querySelector("#eventToggle")?.closest(".toggle"),
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
  bigFiveTableHead: document.querySelector("#bigFiveTableHead"),
  bigFiveTable: document.querySelector("#bigFiveTable"),
  healthMonthMeta: document.querySelector("#healthMonthMeta"),
  healthTableHead: document.querySelector("#healthTableHead"),
  healthTableBody: document.querySelector("#healthTableBody"),
  leadershipVacancyChart: document.querySelector("#leadershipVacancyChart"),
  leadershipVacancyMeta: document.querySelector("#leadershipVacancyMeta"),
};

let liveExcel;

function getMetricLabels() {
  return Object.fromEntries(
    metricOrder.map((key) => [
      key,
      key === "healthReport"
        ? "Monthly Health Report"
        : key === "bigFive"
          ? "Big 5"
          : data.metrics[key]?.label || key,
    ]),
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

function formatMonth(monthKey) {
  if (!monthKey) return "--";
  return parseIsoDate(`${monthKey}-01`).toLocaleDateString("en-US", {
    month: "long",
    year: "numeric",
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

function formatHealthPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  const decimals = Math.abs(value) > 0 && Math.abs(value) < 1 ? 2 : 1;
  return `${Number(value).toFixed(decimals)}%`;
}

function formatTargetPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  return `${Number(value).toLocaleString("en-US", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 0,
  })}%`;
}

function formatSignedHealthPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  const sign = value > 0 ? "+" : "";
  return `${sign}${Number(value).toFixed(1)}%`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function isMultiWeekBigFiveEvent() {
  const eventText = normalizeText(`${state.bigFiveEvent} ${bigFiveEventLabel()}`);
  return eventText.includes("relationshipseries") || eventText.includes("atthemovies");
}

function isGoalBigFiveEvent() {
  const eventText = normalizeText(`${state.bigFiveEvent} ${bigFiveEventLabel()}`);
  return eventText.includes("welcomehomesunday") || eventText.includes("christmas") || eventText.includes("eoy");
}

function growthBarrierForAttendance(value) {
  if (!isFiniteNumber(value)) return null;
  return growthBarriers.find((barrier) => value < barrier.max) || growthBarriers.at(-1);
}

function goalRowsForCampus(campus) {
  if (campus === "All Campuses") return bigFiveGrowthGoals;
  const normalized = normalizeText(campus);
  return bigFiveGrowthGoals.filter((goal) => goal.aliases.some((alias) => normalizeText(alias) === normalized));
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

function isFiniteNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function latestAttendancePoint(campus) {
  return (data.metrics.attendance?.series?.[campus] || [])
    .filter((point) => point.isSunday && typeof point.value === "number" && Number.isFinite(point.value) && point.value > 0)
    .filter((point) => state.showEvents || point.eventType === "normal")
    .at(-1);
}

function pointForCampusDate(metricKey, campus, date) {
  if (!date) return null;
  const point = (data.metrics[metricKey]?.series?.[campus] || []).find(
    (item) => item.date === date && item.isSunday,
  );
  if (!point || point.value === null || point.value === undefined || Number.isNaN(point.value)) return null;
  if (metricKey === "attendance" && point.value <= 0) return null;
  return point;
}

function targetText(target) {
  if (!target) return "";
  return target.upper ? `${target.lower}-${target.upper}%` : `${target.lower}%+`;
}

function rateRowsForMetric(metricKey, campusScope = state.campus) {
  const campuses = campusScope === "All Campuses" ? data.campuses : [campusScope];
  return campuses
    .map((campus) => {
      const attendance = latestAttendancePoint(campus);
      const metric = pointForCampusDate(metricKey, campus, attendance?.date);
      const ratio = attendance?.value > 0 && metric ? (metric.value / attendance.value) * 100 : null;
      return {
        campus,
        date: attendance?.date ?? null,
        attendance: attendance?.value ?? null,
        value: metric?.value ?? null,
        ratio,
      };
    })
    .filter((row) => row.attendance > 0);
}

function totalRateForMetric(metricKey, campusScope = state.campus) {
  const rows = rateRowsForMetric(metricKey, campusScope).filter((row) => row.value !== null && row.value !== undefined);
  const attendance = rows.reduce((sum, row) => sum + row.attendance, 0);
  const value = rows.reduce((sum, row) => sum + row.value, 0);
  return attendance > 0 ? { attendance, value, ratio: (value / attendance) * 100, rows } : null;
}

function shortCampusList(rows, limit = 3) {
  return rows
    .slice(0, limit)
    .map((row) => row.campus)
    .join(", ");
}

function rateSeverity(ratio, target) {
  if (!target || !isFiniteNumber(ratio)) return "info";
  if (ratio >= target.lower) return "info";
  return ratio < target.lower * 0.75 ? "critical" : "warning";
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

function attendanceSundayCount(campus) {
  return (data.metrics.attendance?.series?.[campus] || []).filter(
    (point) => point.isSunday && typeof point.value === "number" && Number.isFinite(point.value) && point.value > 0,
  ).length;
}

function movementStats(row) {
  const cleanedSwings = (row.swings || []).filter(
    (swing) =>
      Number(swing.fromValue || 0) > 0 &&
      Number(swing.toValue || 0) > 0 &&
      Math.abs(Number(swing.pct || 0)) >= 10,
  );
  const hasSwingDetails = Array.isArray(row.swings);
  const qualifyingSwings = hasSwingDetails ? cleanedSwings : [];
  const bigChangeCount = hasSwingDetails ? qualifyingSwings.length : row.swing10Count ?? 0;
  const upCount =
    hasSwingDetails
      ? qualifyingSwings.filter((swing) => Number(swing.pct || 0) >= 10).length
      : row.positiveSwing10Count ?? 0;
  const downCount =
    hasSwingDetails
      ? qualifyingSwings.filter((swing) => Number(swing.pct || 0) <= -10).length
      : row.negativeSwing10Count ?? 0;
  const seriesSundayCount = attendanceSundayCount(row.campus);
  const sundayCount = seriesSundayCount || row.sundayCount || qualifyingSwings.length + 1;

  return {
    ...row,
    bigChangeCount,
    upCount,
    downCount,
    sundayCount,
  };
}

function renderKpis(points) {
  const latest = points.at(-1);
  const previous = points.at(-2);
  const weekChange = latest && previous ? pctChange(latest.value, previous.value) : null;
  const totalDiag = totalsDiagnostics(points);

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
      note: latest ? `Most recent Sunday, ${shortDate(latest.date)}` : "No data",
    },
    {
      label: "Since last Sunday",
      value: formatPct(weekChange),
      valueClass: toneClass(weekChange),
      note: previous ? `Compared with ${shortDate(previous.date)}` : "No comparison",
    },
    {
      label: "Compared with recent Sundays",
      value: formatPct(baselineDelta),
      valueClass: toneClass(baselineDelta),
      note: "Against the last four regular Sundays",
    },
    {
      label: "Compared with last year",
      value: formatPct(yoy),
      valueClass: toneClass(yoy),
      note:
        state.metric === "attendance" && data.yoy
          ? `Same season last year: ${shortDate(data.yoy.matchedDate)}`
          : "Shown on attendance only",
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
  const tooltipWidth = 156;

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
          const cx = x(index);
          const cy = y(point.value);
          const tooltipX = Math.min(Math.max(cx - tooltipWidth / 2, margin.left - 24), width - margin.right - tooltipWidth + 24);
          const tooltipY = Math.max(cy - 68, margin.top - 10);
          return `
            <g class="chart-point-tooltip" data-tooltip-index="${index}" transform="translate(${tooltipX} ${tooltipY})">
              <rect width="${tooltipWidth}" height="50" rx="7"></rect>
              <text x="12" y="20">${escapeHtml(shortDate(point.date))}</text>
              <text class="tooltip-number" x="12" y="38">${escapeHtml(formatNumber(point.value))}</text>
            </g>
            <circle class="point ${isEvent ? "event" : ""}" data-point-index="${index}" tabindex="0" aria-label="${escapeHtml(
              `${shortDate(point.date)} ${formatNumber(point.value)}`,
            )}" cx="${cx}" cy="${cy}" r="${
              isEvent ? 5 : 4
            }">
              <title>${escapeHtml(shortDate(point.date))}: ${escapeHtml(formatNumber(point.value))}${
                point.event ? `, ${escapeHtml(point.event)}` : ""
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
  wireLineChartTooltips();
}

function wireLineChartTooltips() {
  let pinnedIndex = null;
  const points = els.lineChart.querySelectorAll("[data-point-index]");
  const tooltips = els.lineChart.querySelectorAll("[data-tooltip-index]");

  const showTooltip = (index, pinned = false) => {
    if (pinned) pinnedIndex = index;
    for (const tooltip of tooltips) {
      tooltip.classList.toggle("active", tooltip.dataset.tooltipIndex === index);
    }
  };

  const hideTooltip = () => {
    if (pinnedIndex !== null) return;
    for (const tooltip of tooltips) tooltip.classList.remove("active");
  };

  for (const point of points) {
    point.addEventListener("mouseenter", () => showTooltip(point.dataset.pointIndex));
    point.addEventListener("focus", () => showTooltip(point.dataset.pointIndex));
    point.addEventListener("mouseleave", hideTooltip);
    point.addEventListener("blur", hideTooltip);
    point.addEventListener("click", () => {
      pinnedIndex = pinnedIndex === point.dataset.pointIndex ? null : point.dataset.pointIndex;
      if (pinnedIndex === null) {
        for (const tooltip of tooltips) tooltip.classList.remove("active");
      } else {
        showTooltip(point.dataset.pointIndex, true);
      }
    });
  }
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
  const rows = data.volatility.map(movementStats);
  if (!rows.length) {
    els.volatilityChart.innerHTML = `<div class="empty">Connect Excel to load attendance movement.</div>`;
    return;
  }

  const totals = rows.reduce(
    (sum, row) => ({
      bigChangeCount: sum.bigChangeCount + row.bigChangeCount,
      upCount: sum.upCount + row.upCount,
      downCount: sum.downCount + row.downCount,
      sundayCount: sum.sundayCount + row.sundayCount,
    }),
    { bigChangeCount: 0, upCount: 0, downCount: 0, sundayCount: 0 },
  );
  const max = Math.max(...rows.map((row) => row.bigChangeCount), 1);

  els.volatilityChart.innerHTML = `
    <div class="movement-summary">
      <div class="movement-stat">
        <span>Total 10%+ Sundays</span>
        <strong>${formatNumber(totals.bigChangeCount)}</strong>
      </div>
      <div class="movement-stat">
        <span>Up Sundays</span>
        <strong class="positive">${formatNumber(totals.upCount)}</strong>
      </div>
      <div class="movement-stat">
        <span>Down Sundays</span>
        <strong class="negative">${formatNumber(totals.downCount)}</strong>
      </div>
      <div class="movement-stat">
        <span>Tracked Sundays</span>
        <strong>${formatNumber(totals.bigChangeCount)} of ${formatNumber(totals.sundayCount)}</strong>
      </div>
    </div>
    <div class="movement-list">
      ${rows
        .map((row) => {
          const width = row.bigChangeCount ? Math.max(4, (row.bigChangeCount / max) * 100) : 0;
          const avg = row.avgAbsSwing10Pct ? `${row.avgAbsSwing10Pct}% avg` : "No average yet";
          return `
        <div class="bar-row movement-row ${row.campus === state.campus ? "active" : ""}">
          <div class="bar-label" title="${row.campus}">${row.campus}</div>
          <div class="bar-track"><div class="bar-fill vol-fill" style="width:${width}%"></div></div>
          <div class="movement-value">
            <strong>${formatNumber(row.bigChangeCount)} of ${formatNumber(row.sundayCount)}</strong>
            <span>${formatNumber(row.upCount)} up / ${formatNumber(row.downCount)} down · ${avg}</span>
          </div>
        </div>
      `;
        })
        .join("")}
    </div>
  `;
}

function attendanceOpportunityInsights() {
  const stats = data.campusStats.filter((row) => isFiniteNumber(row.latest));
  if (!stats.length) return [];

  if (state.campus !== "All Campuses") {
    const stat = statForCampus(state.campus);
    if (!stat) return [];
    const gap = stat.previous4Avg && stat.latest < stat.previous4Avg ? stat.previous4Avg - stat.latest : 0;
    const insights = [];
    if (gap > 0) {
      insights.push({
        title: "Growth gap to close",
        body: `${state.campus} is about ${formatNumber(gap)} people below its recent four-Sunday pace. That is the clearest short-term attendance gap to recover.`,
        severity: gap > stat.previous4Avg * 0.1 ? "critical" : "warning",
      });
    } else if (isFiniteNumber(stat.vsPrevious4Pct)) {
      insights.push({
        title: "Momentum to protect",
        body: `${state.campus} is ${formatPct(stat.vsPrevious4Pct)} compared with its recent four-Sunday pace. Look for what is working and make it repeatable.`,
        severity: "info",
      });
    }
    const barrier = growthBarrierForAttendance(stat.previous4Avg || stat.latest);
    if (barrier) {
      insights.push({
        title: `Growth barrier lens: ${barrier.label}`,
        body: `${barrier.issue} Emphasis: ${barrier.emphasis}.`,
        severity: barrier.max <= 1000 ? "warning" : "info",
      });
    }
    return insights;
  }

  const belowRecent = stats
    .filter((row) => isFiniteNumber(row.vsPrevious4Pct) && row.vsPrevious4Pct < -5)
    .sort((a, b) => a.vsPrevious4Pct - b.vsPrevious4Pct);
  const aboveRecent = stats
    .filter((row) => isFiniteNumber(row.vsPrevious4Pct) && row.vsPrevious4Pct > 5)
    .sort((a, b) => b.vsPrevious4Pct - a.vsPrevious4Pct);
  const peopleGap = belowRecent.reduce(
    (sum, row) => sum + Math.max(0, (row.previous4Avg || 0) - (row.latest || 0)),
    0,
  );
  const insights = [];

  if (belowRecent.length) {
    insights.push({
      title: `Growth gap: ${belowRecent.length} campuses below recent pace`,
      body: `Closing the gap for ${shortCampusList(belowRecent)} would recover about ${formatNumber(peopleGap)} people toward the recent four-Sunday pace.`,
      severity: belowRecent.length >= Math.ceil(stats.length / 3) ? "critical" : "warning",
    });
  }

  if (aboveRecent.length) {
    const leader = aboveRecent[0];
    insights.push({
      title: `${leader.campus} has momentum worth studying`,
      body: `${leader.campus} is ${formatPct(leader.vsPrevious4Pct)} above its recent four-Sunday pace. Compare invite, follow-up, serving, and local outreach rhythms with campuses that are trailing.`,
      severity: "info",
    });
  }

  const latestTotal = stats.reduce((sum, row) => sum + (row.latest || 0), 0);
  const topCampus = stats.slice().sort((a, b) => (b.latest || 0) - (a.latest || 0))[0];
  if (latestTotal > 0 && topCampus) {
    const share = (topCampus.latest / latestTotal) * 100;
    insights.push({
      title: `${topCampus.campus} carries ${formatHealthPct(share)} of attendance`,
      body:
        share > 30
          ? "Overall growth is meaningfully tied to one campus. Protect that strength while building repeatable growth plays for the next tier of campuses."
          : "Attendance is fairly distributed, which means system-wide gains can come from small improvements across several campuses.",
      severity: share > 30 ? "warning" : "info",
    });
  }

  const barrierCounts = new Map();
  for (const row of stats) {
    const barrier = growthBarrierForAttendance(row.previous4Avg || row.latest);
    if (!barrier) continue;
    const current = barrierCounts.get(barrier.label) || { ...barrier, count: 0, campuses: [] };
    current.count += 1;
    current.campuses.push(row.campus);
    barrierCounts.set(barrier.label, current);
  }
  const topBarrier = Array.from(barrierCounts.values()).sort((a, b) => b.count - a.count)[0];
  if (topBarrier) {
    insights.push({
      title: `Most common growth barrier: ${topBarrier.label}`,
      body: `${topBarrier.campuses.slice(0, 3).join(", ")} point to this barrier. ${topBarrier.issue} Emphasis: ${topBarrier.emphasis}.`,
      severity: topBarrier.max <= 1000 ? "warning" : "info",
    });
  }

  insights.push({
    title: "5 Keys diagnosis",
    body: `When attendance is down or flat, read it through: ${growthKeys.join(", ")}.`,
    severity: "info",
  });

  return insights;
}

function metricOpportunityInsights() {
  const target = metricRateTargets[state.metric];
  if (!target) return [];

  const total = totalRateForMetric(state.metric);
  if (!total || !isFiniteNumber(total.ratio)) return [];

  const label = metricLabels[state.metric];
  const insights = [];
  const targetLabel = targetText(target);
  const statusText =
    total.ratio < target.lower
      ? `below the ${targetLabel} ministry health target`
      : target.upper && total.ratio > target.upper
        ? `above the ${targetLabel} target range`
        : `inside the ${targetLabel} target range`;

  insights.push({
    title: `${label} is ${formatHealthPct(total.ratio)} of attendance`,
    body: `${label} is ${statusText}. This connects the attendance number to the ${target.ministry} lane instead of reading the raw count alone. Emphasis: ${target.emphasis}.`,
    severity: rateSeverity(total.ratio, target),
  });

  const rows = total.rows.filter((row) => isFiniteNumber(row.ratio));
  if (!rows.length) return insights;

  const lowRows = rows.slice().sort((a, b) => a.ratio - b.ratio);
  const highRows = rows.slice().sort((a, b) => b.ratio - a.ratio);
  const lowest = lowRows[0];
  const highest = highRows[0];

  if (state.campus === "All Campuses") {
    if (lowest && lowest.ratio < target.lower) {
      const peopleGap = Math.max(0, (target.lower / 100) * lowest.attendance - (lowest.value || 0));
      insights.push({
        title: `${lowest.campus} is the clearest ${label} opportunity`,
        body: `${lowest.campus} is at ${formatHealthPct(lowest.ratio)}. Reaching ${targetLabel} would mean roughly ${formatNumber(peopleGap)} more people in this lane on a comparable Sunday.`,
        severity: rateSeverity(lowest.ratio, target),
      });
    }

    if (highest && highest.campus !== lowest?.campus) {
      insights.push({
        title: `Learn from ${highest.campus}`,
        body: `${highest.campus} is at ${formatHealthPct(highest.ratio)} for ${label}. That campus may have a practice the others can borrow.`,
        severity: "info",
      });
    }

    if (highest && lowest && highest.campus !== lowest.campus) {
      const spread = highest.ratio - lowest.ratio;
      insights.push({
        title: `Campus spread is ${formatHealthPct(spread)}`,
        body: `${label} ranges from ${formatHealthPct(lowest.ratio)} at ${lowest.campus} to ${formatHealthPct(highest.ratio)} at ${highest.campus}. Large spreads usually point to a training or process opportunity.`,
        severity: spread > Math.max(4, target.lower * 0.75) ? "warning" : "info",
      });
    }
  } else {
    const campusRow = rows[0];
    if (campusRow.ratio < target.lower) {
      const peopleGap = Math.max(0, (target.lower / 100) * campusRow.attendance - (campusRow.value || 0));
      insights.push({
        title: `${state.campus} has a ${label} gap`,
        body: `Reaching ${targetLabel} would mean roughly ${formatNumber(peopleGap)} more people in ${target.ministry} on a comparable Sunday.`,
        severity: rateSeverity(campusRow.ratio, target),
      });
    }
  }

  return insights;
}

function growthLeverInsights() {
  const rows = Object.keys(metricRateTargets)
    .map((metricKey) => {
      const target = metricRateTargets[metricKey];
      const total = totalRateForMetric(metricKey);
      if (!total || !isFiniteNumber(total.ratio)) return null;
      return {
        metricKey,
        label: metricLabels[metricKey],
        target,
        ratio: total.ratio,
        gap: target.lower - total.ratio,
      };
    })
    .filter(Boolean)
    .filter((row) => row.gap > 0)
    .sort((a, b) => b.gap / b.target.lower - a.gap / a.target.lower);

  const top = rows[0];
  if (!top) return [];
  return [
    {
      title: `Growth lever to watch: ${top.label}`,
      body: `${top.label} is ${formatHealthPct(top.ratio)} of attendance against a ${targetText(top.target)} target. This ties to ${top.target.key} and the emphasis area: ${top.target.emphasis}.`,
      severity: top.ratio < top.target.lower * 0.75 ? "critical" : "warning",
    },
  ];
}

function selectedCampusInsights() {
  const stat = statForCampus(state.campus);
  if (!stat) return [];
  const weekTone = Math.abs(stat.weekChangePct || 0) >= 10 ? "warning" : "info";
  const baselineTone = (stat.vsPrevious4Pct || 0) < -5 ? "warning" : "info";

  return [
    {
      title: `${state.campus} this Sunday`,
      body: `${state.campus} reported ${formatNumber(stat.latest)} on ${shortDate(
        stat.latestDate,
      )}, ${formatPct(stat.weekChangePct)} compared with ${shortDate(stat.previousDate)}.`,
      severity: weekTone,
    },
    {
      title: "Compared with recent regular Sundays",
      body: `${state.campus} is ${formatPct(stat.vsPrevious4Pct)} compared with its previous four regular Sundays.`,
      severity: baselineTone,
    },
  ];
}

function renderInsights() {
  const baseInsights =
    state.metric !== "attendance"
      ? []
      : state.campus === "All Campuses"
      ? data.insights
          .filter((insight) => !["volatility", "event"].includes(insight.type))
          .map((insight) => ({
            title: insight.title,
            body: insight.body,
            severity: insight.severity,
          }))
      : selectedCampusInsights();
  const insights = [
    ...baseInsights,
    ...(state.metric === "attendance" ? attendanceOpportunityInsights() : []),
    ...metricOpportunityInsights(),
    ...(state.metric === "attendance" ? growthLeverInsights() : []),
  ];

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

function monthKeyFromIso(iso) {
  return iso ? iso.slice(0, 7) : null;
}

function healthMonths() {
  const months = new Set(data.health?.months || []);
  for (const points of Object.values(data.metrics.attendance?.series || {})) {
    for (const point of points) {
      if (point.isSunday) months.add(monthKeyFromIso(point.date));
    }
  }
  return Array.from(months).filter(Boolean).sort();
}

function selectedHealthCampuses() {
  if (state.campus === "All Campuses") return data.campuses;
  return data.campuses.includes(state.campus) ? [state.campus] : [];
}

function currentMonthKey() {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
}

function latestCompletedHealthMonth() {
  const months = healthMonths();
  if (!months.length) return null;
  const completed = months.filter((month) => month < currentMonthKey());
  return completed.at(-1) || months.at(-1);
}

function sameMonth(iso, month) {
  return Boolean(iso && month && iso.startsWith(month));
}

function averageNumbers(values) {
  const usable = values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (!usable.length) return null;
  return usable.reduce((sum, value) => sum + value, 0) / usable.length;
}

function sumNumbers(values) {
  const usable = values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (!usable.length) return null;
  return usable.reduce((sum, value) => sum + value, 0);
}

function metricValueOnDate(metricKey, campuses, date) {
  const series = data.metrics[metricKey]?.series || {};
  let total = 0;
  let found = false;
  for (const campus of campuses) {
    const point = (series[campus] || []).find((item) => item.date === date);
    if (point && typeof point.value === "number" && Number.isFinite(point.value)) {
      total += point.value;
      found = true;
    }
  }
  return found ? total : null;
}

function attendanceDatesForMonth(campuses, month) {
  const dates = new Set();
  for (const campus of campuses) {
    for (const point of data.metrics.attendance?.series?.[campus] || []) {
      if (point.isSunday && point.value > 0 && sameMonth(point.date, month)) dates.add(point.date);
    }
  }
  return Array.from(dates).sort();
}

function historyValueOnDate(series, campuses, date) {
  let total = 0;
  let found = false;
  for (const campus of campuses) {
    const point = (series[campus] || []).find((item) => item.date === date);
    if (point && typeof point.value === "number" && Number.isFinite(point.value)) {
      total += point.value;
      found = true;
    }
  }
  return found ? total : null;
}

function nearestPriorAttendance(date, campuses) {
  const history = data.history?.attendance2025 || {};
  const target = parseIsoDate(date);
  target.setFullYear(target.getFullYear() - 1);
  const maxDistance = 9 * 24 * 60 * 60 * 1000;
  let bestDate = null;
  let bestDistance = Infinity;

  for (const campus of campuses) {
    for (const point of history[campus] || []) {
      if (!point.isSunday || point.value <= 0) continue;
      const distance = Math.abs(parseIsoDate(point.date) - target);
      if (distance < bestDistance) {
        bestDate = point.date;
        bestDistance = distance;
      }
    }
  }

  if (!bestDate || bestDistance > maxDistance) return null;
  const value = historyValueOnDate(history, campuses, bestDate);
  return value === null ? null : { date: bestDate, value };
}

function weeklyPercentOfAttendance(metricKey, date, campuses) {
  const attendance = metricValueOnDate("attendance", campuses, date);
  const numerator = metricValueOnDate(metricKey, campuses, date);
  return attendance && numerator !== null ? (numerator / attendance) * 100 : null;
}

function monthlyPercentOfAttendance(metricKey, dates, campuses) {
  return averageNumbers(dates.map((date) => weeklyPercentOfAttendance(metricKey, date, campuses)));
}

function healthTarget(key) {
  const normalized = key.toLowerCase().replace(/[^a-z0-9]/g, "");
  const workbookTarget = (data.health?.targets || []).find(
    (target) => String(target.key || "").toLowerCase().replace(/[^a-z0-9]/g, "") === normalized,
  );
  const override = Object.fromEntries(
    Object.entries(workbookTarget || {}).filter(([, value]) => value !== null && value !== undefined && value !== ""),
  );
  return { ...(healthTargetDefaults[key] || {}), ...override };
}

function targetLabel(key) {
  if (healthOptimalLabels[key]) return healthOptimalLabels[key];
  const target = healthTarget(key);
  const hasMin = target.optimalMin !== undefined && target.optimalMin !== null;
  const hasMax = target.optimalMax !== undefined && target.optimalMax !== null;
  if (!hasMin && !hasMax) return "No target";
  if (target.direction === "lower") {
    return target.unit === "%" ? `<= ${formatTargetPct(target.optimalMax)}` : `<= ${formatNumber(target.optimalMax)}`;
  }
  if (hasMin && hasMax) {
    return `${formatTargetPct(target.optimalMin)}-${formatTargetPct(target.optimalMax)}`;
  }
  if (target.unit === "%") return `${formatTargetPct(target.optimalMin)}`;
  return `${formatNumber(target.optimalMin)}+`;
}

function severityFromGrade(grade) {
  if (grade === null || grade === undefined) return { label: "Info", tone: "neutral" };
  if (grade <= 1) return { label: "On Track", tone: "positive" };
  if (grade <= 4) return { label: "Watch", tone: "watch" };
  if (grade <= 7) return { label: "Urgent", tone: "urgent" };
  return { label: "Critical", tone: "critical" };
}

function gradeFromDeficit(deficitRatio) {
  if (deficitRatio <= 0) return 1;
  if (deficitRatio <= 0.05) return 3;
  if (deficitRatio <= 0.15) return 5;
  if (deficitRatio <= 0.3) return 7;
  if (deficitRatio <= 0.5) return 9;
  return 10;
}

function gradeAboveRange(excessRatio) {
  if (excessRatio <= 0) return 1;
  if (excessRatio <= 0.1) return 2;
  if (excessRatio <= 0.25) return 3;
  return 5;
}

function gradeForLowerTarget(value, max) {
  if (value <= max) return 1;
  if (max && max > 0) return gradeFromDeficit((value - max) / max);
  if (value <= 1) return 4;
  if (value <= 3) return 6;
  if (value <= 5) return 8;
  return 10;
}

function healthStatus(key, value) {
  if (key === "heartSoulTeamLeadPct" && (value === null || value === undefined || Number.isNaN(value))) {
    return { label: "Not Graded", tone: "not-graded", grade: null };
  }
  if (value === null || value === undefined || Number.isNaN(value)) {
    return { label: "Info", tone: "neutral", grade: null };
  }
  const target = healthTarget(key);
  const hasMin = target.optimalMin !== undefined && target.optimalMin !== null;
  const hasMax = target.optimalMax !== undefined && target.optimalMax !== null;
  if (!hasMin && !hasMax) {
    return { label: "Info", tone: "neutral", grade: null };
  }
  if (target.direction === "lower") {
    const grade = gradeForLowerTarget(value, target.optimalMax);
    return { ...severityFromGrade(grade), grade };
  }
  if (hasMin && value < target.optimalMin) {
    const grade = gradeFromDeficit((target.optimalMin - value) / target.optimalMin);
    return { ...severityFromGrade(grade), grade };
  }
  if (hasMax && value > target.optimalMax) {
    const grade = gradeAboveRange((value - target.optimalMax) / target.optimalMax);
    return { label: grade <= 2 ? "Above Target" : "Review", tone: grade <= 2 ? "neutral" : "watch", grade };
  }
  return { label: "On Track", tone: "positive", grade: 1 };
}

function configCoversMonth(row, month) {
  if (!row.startDate && !row.endDate) return true;
  const start = row.startDate ? row.startDate.slice(0, 7) : month;
  const end = row.endDate ? row.endDate.slice(0, 7) : month;
  return month >= start && month <= end;
}

function sumConfig(rows, field) {
  return sumNumbers(rows.map((row) => row[field]));
}

function groupConfigFor(campuses, month) {
  const rows = (data.health?.groupConfig || []).filter(
    (row) => campuses.includes(row.campus) && configCoversMonth(row, month),
  );
  if (!rows.length) return null;
  return {
    groupGoal: sumConfig(rows, "groupGoal"),
    activeGroups: sumConfig(rows, "activeGroups"),
    groupSignups: sumConfig(rows, "groupSignups"),
    totalGroupMembers: sumConfig(rows, "totalGroupMembers"),
  };
}

function groupAttendanceByWeek(campuses, month) {
  const byWeek = new Map();
  for (const row of data.health?.groupAttendance || []) {
    if (!campuses.includes(row.campus) || !sameMonth(row.weekStart, month)) continue;
    if (!byWeek.has(row.weekStart)) byWeek.set(row.weekStart, { weekStart: row.weekStart, attendance: 0 });
    byWeek.get(row.weekStart).attendance += row.groupAttendance || 0;
  }
  return Array.from(byWeek.values()).sort((a, b) => a.weekStart.localeCompare(b.weekStart));
}

function heartSoulSummary(campuses, month, role) {
  const normalizedRole = role.toLowerCase();
  const rows = (data.health?.heartSoulRows || []).filter(
    (row) =>
      campuses.includes(row.campus) &&
      row.month === month &&
      String(row.role || "").toLowerCase() === normalizedRole &&
      row.included !== false,
  );
  const invited = sumNumbers(rows.map((row) => row.invitedCount));
  const eligible = sumNumbers(rows.map((row) => row.eligibleCount));
  const attended = sumNumbers(rows.map((row) => row.attendedCount));
  const denominator = invited || eligible;
  return {
    attended,
    denominator,
    pct: denominator && attended !== null ? (attended / denominator) * 100 : null,
  };
}

function leadershipSummary(campuses, month) {
  const sourceRows = (data.health?.leadershipRows || []).filter(
    (row) => campuses.includes(row.campus) && row.month === month,
  );
  const roles = ["Director", "Coordinator", "Team Lead"].map((role) => {
    const rows = sourceRows.filter((row) => String(row.role || "").toLowerCase() === role.toLowerCase());
    const target = sumNumbers(rows.map((row) => row.targetPositions));
    const filled = sumNumbers(rows.map((row) => row.filledPositions));
    const vacancies = sumNumbers(rows.map((row) => row.vacancies));
    return {
      role,
      target,
      filled,
      vacancies,
      fillPct: target && filled !== null ? (filled / target) * 100 : null,
    };
  });
  const target = sumNumbers(roles.map((row) => row.target));
  const filled = sumNumbers(roles.map((row) => row.filled));
  const vacancies = sumNumbers(roles.map((row) => row.vacancies));
  return {
    roles,
    target,
    filled,
    vacancies,
    fillPct: target && filled !== null ? (filled / target) * 100 : null,
  };
}

function buildHealthReport() {
  const campuses = selectedHealthCampuses();
  const month = latestCompletedHealthMonth();
  const dates = month ? attendanceDatesForMonth(campuses, month) : [];
  const attendanceValues = dates.map((date) => metricValueOnDate("attendance", campuses, date));
  const attendanceAvg = averageNumbers(attendanceValues);
  const attendanceYoyValues = dates.map((date) => {
    const current = metricValueOnDate("attendance", campuses, date);
    const prior = nearestPriorAttendance(date, campuses)?.value;
    return pctChange(current, prior);
  });
  const attendanceYoy = averageNumbers(attendanceYoyValues);
  const config = groupConfigFor(campuses, month);
  const groupWeeks = groupAttendanceByWeek(campuses, month);
  const leadership = leadershipSummary(campuses, month);

  const row = (key, weeklyValues, monthValue, format, source, statusValue = monthValue) => ({
    key,
    label: healthTarget(key).label || key,
    optimal: targetLabel(key),
    weeklyValues,
    monthValue,
    format,
    source,
    status: healthStatus(key, statusValue),
  });

  const ratioRow = (key, metricKey, source = "Weekly workbook") =>
    row(
      key,
      dates.map((date) => weeklyPercentOfAttendance(metricKey, date, campuses)),
      monthlyPercentOfAttendance(metricKey, dates, campuses),
      "pct",
      source,
    );

  const groupAttendanceValues = dates.map((date) => {
    const week = groupWeeks.find((item) => item.weekStart === date);
    const attended = week?.attendance ?? null;
    const signups = config?.groupSignups ?? null;
    return {
      attended,
      signups,
      pct: signups && attended !== null ? (attended / signups) * 100 : null,
    };
  });
  const groupAverage = averageNumbers(groupWeeks.map((week) => week.attendance));
  const groupMonthValue = {
    attended: groupAverage,
    signups: config?.groupSignups ?? null,
    pct: averageNumbers(groupAttendanceValues.map((value) => value.pct)),
  };

  const heartRoles = [
    ["heartSoulDirectorPct", "Director"],
    ["heartSoulCoordinatorPct", "Coordinator"],
    ["heartSoulTeamLeadPct", "Team Lead"],
  ];

  const rows = [
    row("attendance", attendanceValues, attendanceAvg, "count", "Attendance sheet"),
    row(
      "attendanceYoy",
      attendanceYoyValues,
      attendanceYoy,
      "signedPct",
      "Attendance + 2025 sheet",
    ),
    ratioRow("kidsPct", "kids"),
    ratioRow("growthTrackPct", "growthTrack"),
    ratioRow("baptismPct", "baptism", "Baptism sheet"),
    ratioRow("salvationsPct", "salvations"),
    ratioRow("firstTimersPct", "firstTimers"),
    ratioRow("dreamTeamPct", "dreamTeam"),
    row(
      "groupsPct",
      dates.map(() => null),
      config?.activeGroups !== null && config?.activeGroups !== undefined && attendanceAvg
        ? (config.activeGroups / attendanceAvg) * 100
        : null,
      "pct",
      "Group Semester Config",
    ),
    row(
      "groupGoalPct",
      dates.map(() => null),
      config?.activeGroups !== null && config?.activeGroups !== undefined && config?.groupGoal
        ? (config.activeGroups / config.groupGoal) * 100
        : null,
      "pct",
      "Group Semester Config",
    ),
    row(
      "groupMembersPct",
      dates.map(() => null),
      config?.totalGroupMembers !== null && config?.totalGroupMembers !== undefined && attendanceAvg
        ? (config.totalGroupMembers / attendanceAvg) * 100
        : null,
      "pct",
      "Group Semester Config",
    ),
    row(
      "groupAttendancePct",
      groupAttendanceValues,
      groupMonthValue,
      "attendanceVsSignups",
      "Group Attendance",
      groupMonthValue.pct,
    ),
    ...heartRoles.map(([key, roleName]) => {
      const summary = heartSoulSummary(campuses, month, roleName);
      return row(key, dates.map(() => null), summary.pct, "pct", "Heart Soul Input");
    }),
    row("leadershipFillPct", dates.map(() => null), leadership.fillPct, "pct", "Leadership Input"),
    row("leadershipVacancies", dates.map(() => null), leadership.vacancies, "count", "Leadership Input"),
  ];

  return { campuses, month, dates, rows, groupWeeks, leadership };
}

function formatHealthCell(format, value) {
  if (format === "count") return formatNumber(value);
  if (format === "pct") return formatHealthPct(value);
  if (format === "signedPct") return formatSignedHealthPct(value);
  if (format === "attendanceVsSignups") {
    return value?.pct === null || value?.pct === undefined ? "--" : formatHealthPct(value.pct);
  }
  return value ?? "--";
}

function renderHealthTable(report) {
  const headerCell = (label) => `
    <th>${label}</th>
  `;
  const weekHeaders = report.dates.map((date) => headerCell(shortDate(date))).join("");
  els.healthTableHead.innerHTML = `
    <tr>
      ${headerCell("Metric")}
      ${headerCell("Optimal")}
      ${weekHeaders}
      ${headerCell("Month")}
      ${headerCell("Grade")}
      ${headerCell("Status")}
    </tr>
  `;

  if (!report.month) {
    els.healthTableBody.innerHTML = `
      <tr>
        <td colspan="${report.dates.length + 5}">Connect Excel to load the monthly health report.</td>
      </tr>
    `;
    return;
  }

  els.healthTableBody.innerHTML = report.rows
    .map(
      (row) => `
        <tr>
          <td>${row.label}</td>
          <td>${row.optimal}</td>
          ${row.weeklyValues.map((value) => `<td>${formatHealthCell(row.format, value)}</td>`).join("")}
          <td>${formatHealthCell(row.format, row.monthValue)}</td>
          <td class="grade-cell ${row.status.tone}">${row.status.grade ?? "--"}</td>
          <td class="status-cell ${row.status.tone}">${row.status.label}</td>
        </tr>
      `,
    )
    .join("");
}

function renderLeadershipVacancyChart(report) {
  const rows = report.leadership.roles.map((role) => ({
    label: role.role,
    value: role.vacancies ?? null,
    target: role.target,
    filled: role.filled,
  }));
  const hasData = rows.some((row) => row.value !== null);
  const totalVacancies = sumNumbers(rows.map((row) => row.value));

  els.leadershipVacancyMeta.textContent =
    totalVacancies !== null ? `${formatNumber(totalVacancies)} total vacancies` : "Needs data";

  if (!hasData) {
    els.leadershipVacancyChart.innerHTML = `
      <div class="empty">Leadership target and filled counts will appear here.</div>
    `;
    return;
  }

  const max = Math.max(...rows.map((row) => row.value || 0), 1);
  els.leadershipVacancyChart.innerHTML = rows
    .map((row) => {
      const width = row.value ? Math.max(6, (row.value / max) * 100) : 0;
      return `
        <div class="leadership-row">
          <div>
            <div class="leadership-role">${row.label}</div>
            <div class="leadership-subtext">
              ${formatNumber(row.filled)} filled of ${formatNumber(row.target)} needed
            </div>
          </div>
          <div class="bar-track">
            <div class="bar-fill vacancy-fill" style="width:${width}%"></div>
          </div>
          <div class="leadership-value">${formatNumber(row.value)}</div>
        </div>
      `;
    })
    .join("");
}

function renderHealth() {
  const report = buildHealthReport();
  els.healthMonthMeta.textContent = report.month
    ? `${formatMonth(report.month)} · ${state.campus}`
    : "Latest completed month";
  renderHealthTable(report);
  renderLeadershipVacancyChart(report);
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
        campusPhaseTotals: new Map(),
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
    if (!bucket.campusPhaseTotals.has(row.campus)) {
      bucket.campusPhaseTotals.set(row.campus, {
        campus: row.campus,
        total: 0,
        phaseTotals: new Map(),
      });
    }
    const campusBucket = bucket.campusPhaseTotals.get(row.campus);
    campusBucket.total += row.attendance;
    campusBucket.phaseTotals.set(row.phase, {
      phase: row.phase,
      phaseLabel: row.phaseLabel,
      total: (campusBucket.phaseTotals.get(row.phase)?.total || 0) + row.attendance,
    });
  }

  const records = Array.from(byYear.values())
    .map((record) => {
      const phaseTotals = Array.from(record.phaseTotals.values()).sort(
        (a, b) => phaseOrder(a.phase) - phaseOrder(b.phase),
      );
      const campusTotals = Array.from(record.campusTotals.values()).sort((a, b) => b.total - a.total);
      const campusSummaries = Array.from(record.campusPhaseTotals.values())
        .map((campus) => {
          const campusPhases = Array.from(campus.phaseTotals.values()).sort(
            (a, b) => phaseOrder(a.phase) - phaseOrder(b.phase),
          );
          const campusTotalPhase = campusPhases.find((phase) => phase.phase === "07 Total" && phase.total > 0);
          const campusFeatured =
            campusTotalPhase ||
            campusPhases
              .filter(
                (phase) =>
                  phase.total > 0 &&
                  phase.phase !== "01 Pre Event" &&
                  !phase.phaseLabel.toLowerCase().startsWith("post"),
              )
              .sort((a, b) => b.total - a.total)[0];
          const campusPre = campusPhases.find((phase) => phase.phase === "01 Pre Event");
          const campusPostWeeks = campusPhases
            .filter((phase) => phase.total > 0 && phase.phaseLabel.toLowerCase().startsWith("post"))
            .slice(0, 4);
          const campusPostWeek1 = campusPostWeeks[0] || campusPhases.find((phase) => phase.phase === "08 Post Week 1");
          const campusPostFourWeekAvg = averageNumbers(campusPostWeeks.map((phase) => phase.total));
          return {
            campus: campus.campus,
            total: campus.total,
            preTotal: campusPre?.total ?? null,
            featuredTotal: campusFeatured?.total ?? null,
            postWeek1: campusPostWeek1?.total ?? null,
            liftPct: campusFeatured && campusPre?.total ? pctChange(campusFeatured.total, campusPre.total) : null,
            postRetentionPct:
              campusFeatured && campusPostWeek1?.total ? (campusPostWeek1.total / campusFeatured.total) * 100 : null,
            postFourWeekAvg: campusPostFourWeekAvg,
            postFourWeekRetentionPct:
              campusFeatured && campusPostFourWeekAvg ? (campusPostFourWeekAvg / campusFeatured.total) * 100 : null,
            postFourWeekGrowthPct:
              campusPostFourWeekAvg && campusPre?.total ? pctChange(campusPostFourWeekAvg, campusPre.total) : null,
          };
        })
        .sort((a, b) => b.total - a.total);
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
      const postWeeks = phaseTotals
        .filter((phase) => phase.total > 0 && phase.phaseLabel.toLowerCase().startsWith("post"))
        .slice(0, 4);
      const postWeek1 = postWeeks[0] || phaseTotals.find((phase) => phase.phase === "08 Post Week 1");
      const postFourWeekAvg = averageNumbers(postWeeks.map((phase) => phase.total));
      const liftPct = featuredPhase && pre?.total ? pctChange(featuredPhase.total, pre.total) : null;
      const postRetentionPct =
        featuredPhase && postWeek1?.total ? (postWeek1.total / featuredPhase.total) * 100 : null;

      return {
        ...record,
        phaseTotals,
        campusTotals,
        campusSummaries,
        preTotal: pre?.total ?? null,
        featuredPhase: featuredPhase?.phase ?? null,
        featuredPhaseLabel: featuredPhase?.phaseLabel ?? null,
        featuredTotal: featuredPhase?.total ?? null,
        postWeek1: postWeek1?.total ?? null,
        postWeeks,
        postFourWeekAvg,
        postFourWeekRetentionPct:
          featuredPhase && postFourWeekAvg ? (postFourWeekAvg / featuredPhase.total) * 100 : null,
        postFourWeekGrowthPct: postFourWeekAvg && pre?.total ? pctChange(postFourWeekAvg, pre.total) : null,
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
  const multiWeekEvent = isMultiWeekBigFiveEvent();
  const hasPostWeeks = (latest?.postWeeks || []).length > 0;
  const comparisonValue = multiWeekEvent ? latest?.yoyCampaignPct : latest?.yoyFeaturedPct;
  const kpis = [];

  if (multiWeekEvent) {
    kpis.push({
      label: "Total attendance",
      value: formatNumber(latest?.campaignTotal),
      note: latest ? `${latest.year} ${bigFiveEventLabel()}` : "No event data",
    });
  }

  kpis.push(
    {
      label: "Main Sunday",
      value: formatNumber(latest?.featuredTotal),
      note: latest?.featuredPhaseLabel || "No main Sunday yet",
    },
    {
      label: "Compared with last year",
      value: formatPct(comparisonValue),
      valueClass: toneClass(comparisonValue),
      note: multiWeekEvent ? "Total attendance vs prior year" : "Main Sunday vs prior year",
    },
    {
      label: "Vs previous Sunday",
      value: formatPct(latest?.liftPct),
      valueClass: toneClass(latest?.liftPct),
      note: "Main Sunday vs previous Sunday",
    },
  );

  if (hasPostWeeks) {
    kpis.push(
    {
        label: "Post Week 1 Hold",
      value: formatPct(latest?.postRetentionPct),
      valueClass:
        latest?.postRetentionPct !== null && latest?.postRetentionPct !== undefined
          ? latest.postRetentionPct < 50
            ? "negative"
            : "positive"
          : "neutral",
        note: "Post Week 1 vs main Sunday",
    },
      {
        label: "4-Week Hold",
        value: formatPct(latest?.postFourWeekRetentionPct),
        valueClass:
          latest?.postFourWeekRetentionPct !== null && latest?.postFourWeekRetentionPct !== undefined
            ? latest.postFourWeekRetentionPct < 70
              ? "warning"
              : "positive"
            : "neutral",
        note: `${formatPct(latest?.postFourWeekGrowthPct)} vs pre-event Sundays`,
      },
    );
  }

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
  const multiWeekEvent = isMultiWeekBigFiveEvent();
  els.bigFiveTrendTitle.textContent = multiWeekEvent
    ? `${eventLabel} Total Attendance`
    : `${eventLabel} Main Sunday Attendance`;
  els.bigFiveTrendMeta.textContent =
    state.campus === "All Campuses" ? "All campuses" : state.campus;

  renderBarList(
    els.bigFiveYearChart,
    usable.map((record) => ({
      label: String(record.year),
      value: multiWeekEvent ? record.campaignTotal : record.featuredTotal,
    })),
  );

  els.bigFivePhaseTitle.textContent = latest ? `${latest.year} Event Weeks` : "Event Weeks";
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

  const multiWeekEvent = isMultiWeekBigFiveEvent();
  const hasPostWeeks = (latest.postWeeks || []).length > 0;
  const comparisonValue = multiWeekEvent ? latest.yoyCampaignPct : latest.yoyFeaturedPct;
  const comparisonSubject = multiWeekEvent ? "total attendance" : "main Sunday";
  const bestYear = records.reduce(
    (best, record) => (record.campaignTotal > (best?.campaignTotal || 0) ? record : best),
    null,
  );
  const topCampus = latest.campusTotals[0];
  const biggestPhase = latest.phaseTotals
    .filter((phase) => phase.total > 0)
    .sort((a, b) => b.total - a.total)[0];
  const recentRecords = records.filter((record) => record.campaignTotal > 0).slice(-3);
  const avgLift = averageNumbers(recentRecords.map((record) => record.liftPct));
  const avgRetention = averageNumbers(recentRecords.map((record) => record.postRetentionPct));
  const campusRetentionRows = (latest.campusSummaries || []).filter((row) => isFiniteNumber(row.postRetentionPct));
  const lowRetentionCampus = campusRetentionRows.slice().sort((a, b) => a.postRetentionPct - b.postRetentionPct)[0];
  const strongRetentionCampus = campusRetentionRows.slice().sort((a, b) => b.postRetentionPct - a.postRetentionPct)[0];
  const bestBooster = records
    .filter((record) => isFiniteNumber(record.liftPct) && isFiniteNumber(record.postRetentionPct))
    .map((record) => ({
      ...record,
      boosterScore: Math.max(0, record.liftPct) * 0.35 + record.postRetentionPct,
    }))
    .sort((a, b) => b.boosterScore - a.boosterScore)[0];
  const insights = [
    {
      title: `${latest.year} ${latest.eventLabel} ${comparisonSubject} is ${formatPct(comparisonValue)} compared with last year`,
      body: multiWeekEvent
        ? `${formatNumber(latest.campaignTotal)} total attendance across ${latest.phaseTotals.length} tracked event weeks.`
        : `${formatNumber(latest.featuredTotal)} on the main Sunday. Total attendance is not used as the lead measure for one-Sunday Big 5 events.`,
      severity: comparisonValue < -10 ? "critical" : "info",
    },
    {
      title:
        latest.liftPct === null
          ? "Needs the previous Sunday for comparison"
          : `Main Sunday was ${formatPct(latest.liftPct)} vs the previous Sunday`,
      body:
        latest.liftPct === null
          ? "This event does not have a previous-Sunday comparison for the selected campus view."
          : `${latest.featuredPhaseLabel} reached ${formatNumber(latest.featuredTotal)} compared with ${formatNumber(latest.preTotal)} the previous Sunday.`,
      severity: "info",
    },
    {
      title:
        latest.postRetentionPct === null
          ? "No following Sunday yet"
          : `Post Week 1 held ${formatPct(latest.postRetentionPct)} of the main Sunday`,
      body:
        latest.postRetentionPct === null
          ? "This event does not have following-Sunday data in the selected view."
          : `Post Week 1 was ${formatNumber(latest.postWeek1)} compared with ${formatNumber(latest.featuredTotal)} on the main Sunday.`,
      severity:
        latest.postRetentionPct !== null && latest.postRetentionPct < 50
          ? "critical"
          : latest.postRetentionPct !== null && latest.postRetentionPct < 70
            ? "warning"
            : "info",
    },
    {
      title:
        latest.liftPct !== null && latest.postRetentionPct !== null && latest.liftPct > 0 && latest.postRetentionPct < 70
          ? "Turn event lift into next-Sunday growth"
          : "Event growth handoff",
      body:
        latest.liftPct !== null && latest.postRetentionPct !== null
          ? `This event lifted attendance ${formatPct(latest.liftPct)} from the previous Sunday and held ${formatHealthPct(latest.postRetentionPct)} into Post Week 1. The opportunity is to make every guest follow-up and next-Sunday invitation measurable.`
          : "Once both previous-Sunday and following-Sunday data are available, this card will show whether the event created lasting growth momentum.",
      severity:
        latest.liftPct !== null && latest.postRetentionPct !== null && latest.liftPct > 0 && latest.postRetentionPct < 70
          ? "warning"
          : "info",
    },
    {
      title:
        state.campus === "All Campuses" && topCampus
          ? `${topCampus.campus} had the largest share`
          : `${biggestPhase?.phaseLabel || "Latest week"} was the strongest week`,
      body:
        state.campus === "All Campuses" && topCampus
          ? `${topCampus.campus} contributed ${formatNumber(topCampus.total)} to the ${latest.year} total attendance.`
          : `${biggestPhase?.phaseLabel || "The leading week"} contributed ${formatNumber(biggestPhase?.total)} in ${latest.year}.`,
      severity: "info",
    },
    {
      title: `Highest year so far: ${bestYear?.year}`,
      body: `${bestYear?.eventLabel} reached ${formatNumber(bestYear?.campaignTotal)} total attendance in this dashboard.`,
      severity: latest.year === bestYear?.year ? "info" : "warning",
    },
  ];

  if (hasPostWeeks) {
    insights.push({
      title: `Four-week hold is ${formatPct(latest.postFourWeekRetentionPct)}`,
      body: `The next four weeks averaged ${formatNumber(latest.postFourWeekAvg)}, which is ${formatPct(latest.postFourWeekGrowthPct)} compared with the pre-event Sunday. This shows whether the Big 5 moment became sustained growth, not just one strong Sunday.`,
      severity:
        latest.postFourWeekRetentionPct !== null && latest.postFourWeekRetentionPct < 70
          ? "warning"
          : "info",
    });
  }

  if (isGoalBigFiveEvent()) {
    const goalRows = goalRowsForCampus(state.campus);
    const goalPre = sumNumbers(goalRows.map((row) => row.preEasterAvg));
    const goalEaster = sumNumbers(goalRows.map((row) => row.easterTotal));
    const goalTarget = sumNumbers(goalRows.map((row) => row.growthGoal));
    const goalIncrease = goalEaster !== null && goalPre !== null ? goalEaster - goalPre : null;
    const currentMainSunday = latest.year === 2026 ? latest.featuredTotal : null;
    const gap = currentMainSunday !== null && goalTarget !== null ? goalTarget - currentMainSunday : null;
    const selectedGoal = state.campus === "All Campuses" ? null : goalRows[0];

    insights.push({
      title:
        state.campus === "All Campuses"
          ? `2026 ${bigFiveEventLabel()} goal: ${formatNumber(goalTarget)}`
          : `${state.campus} 2026 goal: ${formatNumber(selectedGoal?.growthGoal)}`,
      body:
        state.campus === "All Campuses"
          ? `The goal keeps 30% of the Easter increase: ${formatNumber(goalPre)} pre-Easter average to ${formatNumber(goalEaster)} Easter total, a ${formatNumber(goalIncrease)} increase. ${gap !== null ? `Current gap is ${formatNumber(Math.max(0, gap))}.` : "Use this as the target before the event lands in the data."}`
          : `${state.campus} is aiming for ${formatNumber(selectedGoal?.growthGoal)} after averaging ${formatNumber(selectedGoal?.preEasterAvg)} pre-Easter and seeing ${formatNumber(selectedGoal?.easterTotal)} at Easter. Growth barrier: ${selectedGoal?.barrier}.`,
      severity: gap !== null && gap > 0 ? "warning" : "info",
    });

    insights.push({
      title: "Growth plan lens",
      body: "Tie the goal to the 5 keys and campus emphasis areas: excellent Sundays, fully staffed leaders and Dream Team, compelling communication, Maximizing Big 5 Sundays, and clear next steps after the event.",
      severity: "info",
    });
  }

  if (recentRecords.length >= 2 && (avgLift !== null || avgRetention !== null)) {
    insights.push({
      title: "Recent Big 5 pattern",
      body: `Across the last ${recentRecords.length} years shown, this event averaged ${formatPct(avgLift)} lift from the previous Sunday and held ${formatHealthPct(avgRetention)} into Post Week 1.`,
      severity: avgRetention !== null && avgRetention < 70 ? "warning" : "info",
    });
  }

  if (state.campus === "All Campuses" && lowRetentionCampus) {
    insights.push({
      title: `${lowRetentionCampus.campus} needs the clearest follow-up plan`,
      body: `${lowRetentionCampus.campus} held ${formatHealthPct(lowRetentionCampus.postRetentionPct)} of its main Sunday into Post Week 1. That points to a guest follow-up, invite-back, and team readiness opportunity.`,
      severity: lowRetentionCampus.postRetentionPct < 50 ? "critical" : "warning",
    });
  }

  if (state.campus === "All Campuses" && strongRetentionCampus && strongRetentionCampus.campus !== lowRetentionCampus?.campus) {
    insights.push({
      title: `Learn from ${strongRetentionCampus.campus}`,
      body: `${strongRetentionCampus.campus} held ${formatHealthPct(strongRetentionCampus.postRetentionPct)} into Post Week 1. Compare its follow-up rhythm and weekend experience with campuses that lost more momentum.`,
      severity: "info",
    });
  }

  if (bestBooster) {
    insights.push({
      title: `Best growth-booster year: ${bestBooster.year}`,
      body: `${bestBooster.eventLabel} paired ${formatPct(bestBooster.liftPct)} lift with ${formatHealthPct(bestBooster.postRetentionPct)} Post Week 1 hold. Use that year as the playbook to compare planning, invites, and follow-up.`,
      severity: bestBooster.year === latest.year ? "info" : "warning",
    });
  }

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
  const multiWeekEvent = isMultiWeekBigFiveEvent();
  const hasPostWeeks = records.some((record) => (record.postWeeks || []).length > 0);
  const headers = [
    "Event",
    "Year",
    ...(multiWeekEvent ? ["Total Attendance"] : []),
    "Main Sunday",
    multiWeekEvent ? "Last Year Total" : "Last Year Main",
    "Vs Previous Sunday",
    ...(hasPostWeeks ? ["Post Week 1", "4-Week Hold", "4-Week Growth"] : []),
  ];

  els.bigFiveTableHead.innerHTML = headers.map((header) => `<th>${header}</th>`).join("");
  els.bigFiveTable.innerHTML = records
    .slice()
    .reverse()
    .map(
      (record) => `
        <tr>
          <td>${record.eventLabel}</td>
          <td>${record.year}</td>
          ${multiWeekEvent ? `<td>${formatNumber(record.campaignTotal)}</td>` : ""}
          <td>${formatNumber(record.featuredTotal)}</td>
          <td class="${toneClass(multiWeekEvent ? record.yoyCampaignPct : record.yoyFeaturedPct)}">${formatPct(
            multiWeekEvent ? record.yoyCampaignPct : record.yoyFeaturedPct,
          )}</td>
          <td class="${toneClass(record.liftPct)}">${formatPct(record.liftPct)}</td>
          ${
            hasPostWeeks
              ? `
                <td class="${record.postRetentionPct !== null && record.postRetentionPct < 50 ? "negative" : "neutral"}">${formatPct(record.postRetentionPct)}</td>
                <td>${formatPct(record.postFourWeekRetentionPct)}</td>
                <td class="${toneClass(record.postFourWeekGrowthPct)}">${formatPct(record.postFourWeekGrowthPct)}</td>
              `
              : ""
          }
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

function isHealthReportSelected() {
  return state.metric === "healthReport";
}

function isBigFiveSelected() {
  return state.metric === "bigFive";
}

function isStandardMetricSelected() {
  return !isHealthReportSelected() && !isBigFiveSelected();
}

function renderMetricPanels() {
  const showHealth = isHealthReportSelected();
  const showBigFive = isBigFiveSelected();
  const showStandard = isStandardMetricSelected();
  for (const panel of els.metricPanels) {
    const mode = panel.dataset.metricPanel;
    const visible =
      (mode === "health" && showHealth) ||
      (mode === "big-five" && showBigFive) ||
      (mode === "standard" && showStandard);
    panel.classList.toggle("is-hidden", !visible);
  }
  for (const panel of els.attendanceOnlyPanels) {
    panel.classList.toggle("is-hidden", state.metric !== "attendance");
  }
  els.eventToggleWrapper?.classList.toggle("is-hidden", !showStandard);
}

function updateDashboard() {
  for (const button of els.metricTabs.querySelectorAll("button")) {
    const active = button.dataset.metric === state.metric;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  }

  renderMetricPanels();
  if (isHealthReportSelected()) {
    renderHealth();
    return;
  }
  if (isBigFiveSelected()) {
    renderBigFive();
    return;
  }

  const points = getMetricPoints(state.metric, state.campus);
  const latest = points.at(-1);
  const previous = points.at(-2);
  const weekChange = latest && previous ? pctChange(latest.value, previous.value) : null;

  els.trendTitle.textContent =
    state.campus === "All Campuses"
      ? `${metricLabels[state.metric]} Over Time`
      : `${state.campus} ${metricLabels[state.metric]}`;
  els.trendMeta.textContent = latest
    ? `${formatNumber(latest.value)} most recent, ${formatPct(weekChange)} since last Sunday`
    : "";

  renderKpis(points);
  renderLineChart(points);
  renderBars();
  renderVolatility();
  renderInsights();
  renderHealth();
  if (state.metric === "attendance") renderTable();
}

setupControls();
updateDashboard();
