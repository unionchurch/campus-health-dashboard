import dashboardData from "./dashboard-data.js?v=20260525-vacancy-role-order";
import { setupLiveExcel } from "./live-excel.js?v=20260525-vacancy-role-order";

let data = dashboardData;

const state = {
  campus: "All Campuses",
  metric: "attendance",
  showEvents: true,
  bigFiveEvent: null,
  bigFiveYear: null,
  healthMonth: null,
  selectedDate: null,
};

const metricOrder = [
  "attendance",
  "kids",
  "growthTrack",
  "salvations",
  "firstTimers",
  "dreamTeam",
  "healthReport",
  "bigFive",
  "campusGrowthHistory",
];

let metricLabels = getMetricLabels();

const healthTargetDefaults = {
  attendance: { label: "Attendance", unit: "count" },
  attendanceYoy: {
    label: "Attendance vs Last Year",
    unit: "%",
    optimalMin: 7,
    optimalMax: 10,
    direction: "range",
  },
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
  dreamTeamPct: {
    label: "Active Dream Team % of Attendance",
    unit: "%",
    optimalMin: 33,
    direction: "higher",
  },
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
  trendInsightLayout: document.querySelector("#trendInsightLayout"),
  chartPanel: document.querySelector(".chart-panel"),
  lineChart: document.querySelector("#lineChart"),
  insightsPanel: document.querySelector(".insights-panel"),
  barKicker: document.querySelector("#barKicker"),
  barTitle: document.querySelector("#barTitle"),
  barMeta: document.querySelector("#barMeta"),
  barChart: document.querySelector("#barChart"),
  dreamTeamCampusPanel: document.querySelector("#dreamTeamCampusPanel"),
  dreamTeamCampusMeta: document.querySelector("#dreamTeamCampusMeta"),
  dreamTeamCampusChart: document.querySelector("#dreamTeamCampusChart"),
  comparisonGrid: document.querySelector("#comparisonGrid"),
  volatilityMeta: document.querySelector("#volatilityMeta"),
  volatilityChart: document.querySelector("#volatilityChart"),
  insights: document.querySelector("#insights"),
  campusTable: document.querySelector("#campusTable"),
  bigFiveEventSelect: document.querySelector("#bigFiveEventSelect"),
  bigFiveYearSelect: document.querySelector("#bigFiveYearSelect"),
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
  healthMonthSelect: document.querySelector("#healthMonthSelect"),
  healthMonthMeta: document.querySelector("#healthMonthMeta"),
  healthReportMeta: document.querySelector("#healthReportMeta"),
  healthInsights: document.querySelector("#healthInsights"),
  healthTableHead: document.querySelector("#healthTableHead"),
  healthTableBody: document.querySelector("#healthTableBody"),
  leadershipVacancyChart: document.querySelector("#leadershipVacancyChart"),
  leadershipVacancyMeta: document.querySelector("#leadershipVacancyMeta"),
  leadershipCampusSnapshotWrap: document.querySelector("#leadershipCampusSnapshotWrap"),
  leadershipCampusMeta: document.querySelector("#leadershipCampusMeta"),
  leadershipCampusSnapshot: document.querySelector("#leadershipCampusSnapshot"),
  growthHistoryMeta: document.querySelector("#growthHistoryMeta"),
  growthHistoryKpis: document.querySelector("#growthHistoryKpis"),
  growthHistoryTrendTitle: document.querySelector("#growthHistoryTrendTitle"),
  growthHistoryTrendMeta: document.querySelector("#growthHistoryTrendMeta"),
  growthHistoryTrendChart: document.querySelector("#growthHistoryTrendChart"),
  growthHistoryShareMeta: document.querySelector("#growthHistoryShareMeta"),
  growthHistoryShareChart: document.querySelector("#growthHistoryShareChart"),
  growthHistoryInsights: document.querySelector("#growthHistoryInsights"),
  growthHistoryTableHead: document.querySelector("#growthHistoryTableHead"),
  growthHistoryTable: document.querySelector("#growthHistoryTable"),
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
          : key === "campusGrowthHistory"
            ? "Campus Growth History"
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

function formatSignedNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "--";
  const rounded = Math.round(value);
  return `${rounded > 0 ? "+" : ""}${rounded.toLocaleString("en-US")}`;
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
    state.bigFiveYear = null;
    updateDashboard();
  });
  els.bigFiveYearSelect.addEventListener("change", (event) => {
    state.bigFiveYear = Number(event.target.value) || null;
    updateDashboard();
  });
  syncHealthMonthOptions();
  els.healthMonthSelect.addEventListener("change", (event) => {
    state.healthMonth = event.target.value;
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

  window.addEventListener("resize", syncTakeawayLayout);
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

function bigFiveArchiveRecords(records) {
  return records.filter((record) => record.campaignTotal > 0);
}

function latestBigFiveYear(records) {
  return bigFiveArchiveRecords(records).at(-1)?.year ?? null;
}

function selectedBigFiveYear(records) {
  const years = bigFiveArchiveRecords(records).map((record) => record.year);
  if (!years.length) return null;
  if (!years.includes(state.bigFiveYear)) {
    state.bigFiveYear = latestBigFiveYear(records);
  }
  return state.bigFiveYear;
}

function selectedBigFiveRecord(records) {
  const year = selectedBigFiveYear(records);
  return bigFiveArchiveRecords(records).find((record) => record.year === year) || null;
}

function syncBigFiveYearOptions(records) {
  const usable = bigFiveArchiveRecords(records);
  const latestYear = latestBigFiveYear(records);
  const selectedYear = selectedBigFiveYear(records);
  els.bigFiveYearSelect.innerHTML = "";
  if (!usable.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No years available";
    els.bigFiveYearSelect.append(option);
    els.bigFiveYearSelect.disabled = true;
    return;
  }
  for (const record of usable.slice().reverse()) {
    const option = document.createElement("option");
    option.value = String(record.year);
    option.textContent = record.year === latestYear ? `${record.year} (Latest)` : String(record.year);
    els.bigFiveYearSelect.append(option);
  }
  if (selectedYear) els.bigFiveYearSelect.value = String(selectedYear);
  els.bigFiveYearSelect.disabled = usable.length === 0;
}

function syncHealthMonthOptions() {
  const months = healthArchiveMonths();
  const defaultMonth = defaultHealthMonth();
  const currentMonth = currentMonthKey();
  const currentSelection = state.healthMonth;
  els.healthMonthSelect.innerHTML = "";
  for (const month of months.slice().reverse()) {
    const option = document.createElement("option");
    option.value = month;
    option.textContent =
      month === currentMonth
        ? `${formatMonth(month)} (Current)`
        : month === defaultMonth
          ? `${formatMonth(month)} (Latest)`
          : formatMonth(month);
    els.healthMonthSelect.append(option);
  }
  state.healthMonth = months.includes(currentSelection) ? currentSelection : defaultMonth;
  if (state.healthMonth) els.healthMonthSelect.value = state.healthMonth;
  els.healthMonthSelect.disabled = months.length === 0;
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
  syncHealthMonthOptions();
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

function selectedPointContext(points) {
  if (!points.length) {
    return {
      index: -1,
      current: null,
      previous: null,
      previousFour: [],
      baseline: null,
      vsBaseline: null,
      weekChange: null,
      isArchive: false,
    };
  }

  const selectedIndex = state.selectedDate
    ? points.findIndex((point) => point.date === state.selectedDate)
    : -1;
  const index = selectedIndex >= 0 ? selectedIndex : points.length - 1;
  const current = points[index] || null;
  const previous = index > 0 ? points[index - 1] : null;
  const previousFour = points
    .slice(0, index)
    .filter((point) => point.eventType !== "holiday" && point.eventType !== "weather")
    .slice(-4);
  const baseline =
    previousFour.length > 0
      ? previousFour.reduce((sum, point) => sum + point.value, 0) / previousFour.length
      : null;

  return {
    index,
    current,
    previous,
    previousFour,
    baseline,
    vsBaseline: current && baseline ? pctChange(current.value, baseline) : null,
    weekChange: current && previous ? pctChange(current.value, previous.value) : null,
    isArchive: Boolean(state.selectedDate && current?.date === state.selectedDate),
  };
}

function metricPointContext(metricKey, campus = state.campus) {
  return selectedPointContext(getMetricPoints(metricKey, campus));
}

function currentMetricDate(metricKey = state.metric, campus = state.campus) {
  return metricPointContext(metricKey, campus).current?.date || null;
}

function activeArchiveDate(metricKey = state.metric, campus = state.campus) {
  const context = metricPointContext(metricKey, campus);
  return context.isArchive ? context.current.date : null;
}

function pointContextForCampusDate(metricKey, campus, date) {
  const points = getMetricPoints(metricKey, campus);
  const index = points.findIndex((point) => point.date === date);
  if (index < 0) return { current: null, previous: null, weekChange: null };
  const current = points[index];
  const previous = index > 0 ? points[index - 1] : null;
  return {
    current,
    previous,
    weekChange: current && previous ? pctChange(current.value, previous.value) : null,
  };
}

function valuesByCampusOnDate(metricKey, date) {
  if (!date) return [];
  return data.campuses
    .map((campus) => {
      const context = pointContextForCampusDate(metricKey, campus, date);
      return {
        campus,
        value: context.current?.value ?? null,
        date,
        previous: context.previous?.value ?? null,
        weekChangePct: context.weekChange,
      };
    })
    .filter((row) => isFiniteNumber(row.value))
    .sort((a, b) => b.value - a.value);
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

function comparisonRowsByCampus(metricKey, date = currentMetricDate(metricKey)) {
  return date ? valuesByCampusOnDate(metricKey, date) : latestByCampus(metricKey);
}

function isFiniteNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function attendancePointForCampus(campus, preferredDate = state.selectedDate) {
  const points = (data.metrics.attendance?.series?.[campus] || [])
    .filter((point) => point.isSunday && typeof point.value === "number" && Number.isFinite(point.value) && point.value > 0)
    .filter((point) => state.showEvents || point.eventType === "normal");
  if (preferredDate) {
    return points.find((point) => point.date === preferredDate) || null;
  }
  return points.at(-1) || null;
}

function latestAttendancePoint(campus) {
  return attendancePointForCampus(campus, null);
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
      const attendance = attendancePointForCampus(campus);
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
  return attendanceCampusRowsForContext().find((item) => item.campus === campus);
}

function totalsDiagnostics(points) {
  const context = selectedPointContext(points);
  return {
    latest: context.current,
    previous: context.previous,
    weekChange: context.weekChange,
    baseline: context.baseline,
    vsBaseline: context.vsBaseline,
  };
}

function attendanceSundayCount(campus, throughDate = null) {
  return (data.metrics.attendance?.series?.[campus] || []).filter(
    (point) =>
      point.isSunday &&
      typeof point.value === "number" &&
      Number.isFinite(point.value) &&
      point.value > 0 &&
      (!throughDate || point.date <= throughDate),
  ).length;
}

function attendanceCampusRowsForContext() {
  return data.campuses
    .map((campus) => {
      const points = getMetricPoints("attendance", campus);
      const context = selectedPointContext(points);
      const current = context.current;
      if (!current) return null;
      return {
        campus,
        sundayCount: points.length,
        latest: current.value,
        latestDate: current.date,
        previous: context.previous?.value ?? null,
        previousDate: context.previous?.date ?? null,
        weekChangePct: context.weekChange,
        previous4Avg: context.baseline ? Math.round(context.baseline * 10) / 10 : null,
        vsPrevious4Pct: context.vsBaseline,
      };
    })
    .filter(Boolean);
}

function movementStats(row) {
  const throughDate = activeArchiveDate("attendance") || null;
  const cleanedSwings = (row.swings || []).filter(
    (swing) =>
      Number(swing.fromValue || 0) > 0 &&
      Number(swing.toValue || 0) > 0 &&
      Math.abs(Number(swing.pct || 0)) >= 10 &&
      (!throughDate || swing.toDate <= throughDate),
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
  const seriesSundayCount = attendanceSundayCount(row.campus, throughDate);
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
  const context = selectedPointContext(points);
  const latest = context.current;
  const previous = context.previous;
  const weekChange = context.weekChange;
  const totalDiag = totalsDiagnostics(points);

  const baselineDelta = totalDiag.vsBaseline;
  const campuses = state.campus === "All Campuses" ? data.campuses : [state.campus];
  const priorYear = state.metric === "attendance" && latest ? nearestPriorAttendance(latest.date, campuses) : null;

  const yoy =
    state.metric !== "attendance"
      ? null
      : priorYear
        ? pctChange(latest?.value, priorYear.value)
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
      note: latest
        ? `${context.isArchive ? "Sunday archive" : "Most recent Sunday"}, ${shortDate(latest.date)}`
        : "No data",
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
        state.metric === "attendance" && (priorYear || data.yoy)
          ? `Same season last year: ${shortDate(priorYear?.date || data.yoy.matchedDate)}`
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

function chartNoteLines(note, maxChars = 28, maxLines = 2) {
  const words = String(note || "")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter(Boolean);
  const lines = [];

  for (const word of words) {
    const current = lines.at(-1) || "";
    if (!current) {
      lines.push(word);
    } else if (`${current} ${word}`.length <= maxChars) {
      lines[lines.length - 1] = `${current} ${word}`;
    } else if (lines.length < maxLines) {
      lines.push(word);
    } else {
      lines[lines.length - 1] = `${current}...`;
      break;
    }
  }

  return lines.slice(0, maxLines);
}

const movementContextMetrics = ["kids", "dreamTeam", "growthTrack", "firstTimers", "salvations", "baptism"];

function scopedMetricValue(metricKey, date) {
  const campuses = state.campus === "All Campuses" ? data.campuses : [state.campus];
  return metricValueOnDate(metricKey, campuses, date);
}

function metricChangeOnDates(metricKey, date, previousDate) {
  const current = scopedMetricValue(metricKey, date);
  const previous = scopedMetricValue(metricKey, previousDate);
  const delta = current !== null && previous !== null ? current - previous : null;
  return {
    metricKey,
    label: metricLabels[metricKey] || data.metrics[metricKey]?.label || metricKey,
    current,
    previous,
    delta,
    changePct: pctChange(current, previous),
  };
}

function campusNameMatches(value, campus) {
  if (campus === "All Campuses") return true;
  const normalizedValue = normalizeText(value);
  const terms = [campus, ...campusAliasTerms(campus)].map(normalizeText);
  return terms.includes(normalizedValue);
}

function dreamTeamDetailRows(date, campusScope = state.campus) {
  return (data.dreamTeamDetail || []).filter(
    (row) => row.date === date && campusNameMatches(row.campus, campusScope),
  );
}

function dreamTeamDetailDate(campusScope = state.campus) {
  if (state.selectedDate && dreamTeamDetailRows(state.selectedDate, campusScope).length) {
    return state.selectedDate;
  }
  return (
    (data.dreamTeamDetail || [])
      .filter((row) => campusNameMatches(row.campus, campusScope))
      .map((row) => row.date)
      .filter(Boolean)
      .sort()
      .at(-1) || currentMetricDate("dreamTeam", campusScope)
  );
}

function dreamTeamDetailStats(campusScope = state.campus) {
  const rows = (data.dreamTeamDetail || []).filter((row) => campusNameMatches(row.campus, campusScope));
  const dates = Array.from(new Set(rows.map((row) => row.date).filter(Boolean))).sort();
  const campuses = Array.from(new Set(rows.map((row) => row.campus).filter(Boolean))).sort();
  return { rows, rowCount: rows.length, dates, campuses };
}

function dreamTeamDetailDiagnostic(date, campusScope = state.campus) {
  const stats = dreamTeamDetailStats(campusScope);
  if (!stats.rowCount) return "DT detail loaded 0 rows from Sunday - DT Detail.";
  const dateList = stats.dates.slice(-3).map(shortDate).join(", ");
  return `DT detail loaded ${formatNumber(stats.rowCount)} rows, ${stats.campuses.length} campuses, dates: ${dateList || "--"}.`;
}

function dreamTeamTeamTotals(date, campusScope = state.campus) {
  const rows = dreamTeamDetailRows(date, campusScope);
  const byTeam = new Map();
  for (const row of rows) {
    const current = byTeam.get(row.team) || { team: row.team, served: 0, target: 0, hasTarget: false };
    current.served += row.served || 0;
    if (row.target !== null && row.target !== undefined) {
      current.target += row.target;
      current.hasTarget = true;
    }
    byTeam.set(row.team, current);
  }
  return Array.from(byTeam.values()).sort((a, b) => b.served - a.served);
}

function dreamTeamCampusRowsFromDetail(date) {
  const byCampus = new Map();
  for (const row of dreamTeamDetailRows(date, "All Campuses")) {
    const current = byCampus.get(row.campus) || 0;
    byCampus.set(row.campus, current + (row.served || 0));
  }
  return Array.from(byCampus.entries())
    .map(([campus, value]) => ({ campus, value, date }))
    .sort((a, b) => b.value - a.value);
}

function dreamTeamTeamChanges(date, previousDate, campusScope = state.campus) {
  if (!dreamTeamDetailRows(date, campusScope).length || !dreamTeamDetailRows(previousDate, campusScope).length) {
    return [];
  }
  const current = new Map(dreamTeamTeamTotals(date, campusScope).map((row) => [row.team, row]));
  const previous = new Map(dreamTeamTeamTotals(previousDate, campusScope).map((row) => [row.team, row]));
  const teams = new Set([...current.keys(), ...previous.keys()]);
  return Array.from(teams)
    .map((team) => {
      const currentServed = current.get(team)?.served ?? 0;
      const previousServed = previous.get(team)?.served ?? 0;
      return {
        team,
        current: currentServed,
        previous: previousServed,
        delta: currentServed - previousServed,
        changePct: pctChange(currentServed, previousServed),
        target: current.get(team)?.hasTarget ? current.get(team).target : null,
      };
    })
    .filter((row) => row.current || row.previous)
    .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
}

function chartContextLines(point, index, points) {
  const previous = points[index - 1];
  if (!previous) return [];

  if (state.metric === "attendance") {
    const attendanceChange = pctChange(point.value, previous.value);
    const direction = attendanceChange < 0 ? -1 : attendanceChange > 0 ? 1 : 0;
    return movementContextMetrics
      .map((metricKey) => metricChangeOnDates(metricKey, point.date, previous.date))
      .filter((row) => row.delta !== null && row.delta !== 0)
      .filter((row) => !direction || Math.sign(row.delta) === direction)
      .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta))
      .slice(0, 3)
      .map((row) => `${row.label}: ${formatSignedNumber(row.delta)} (${formatPct(row.changePct)})`);
  }

  if (state.metric === "dreamTeam") {
    return dreamTeamTeamChanges(point.date, previous.date)
      .slice(0, 3)
      .map((row) => `${row.team}: ${formatSignedNumber(row.delta)} (${formatPct(row.changePct)})`);
  }

  return [];
}

function pointChangeSummary(point, index, points) {
  const previous = points[index - 1];
  if (!previous) return null;
  const delta = point.value - previous.value;
  const changePct = pctChange(point.value, previous.value);
  if (!isFiniteNumber(delta) || !isFiniteNumber(changePct)) return null;
  return {
    label: `${formatSignedNumber(delta)} (${formatPct(changePct)})`,
    tone: toneClass(changePct),
  };
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
  const tooltipWidth = 250;

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
          const noteLines = [...chartNoteLines(point.event), ...chartContextLines(point, index, points)].slice(0, 4);
          const pointChange = pointChangeSummary(point, index, points);
          const hasNote = noteLines.length > 0;
          const isEvent = hasNote || point.eventType !== "normal";
          const cx = x(index);
          const cy = y(point.value);
          const tooltipHeight = hasNote ? 50 + noteLines.length * 14 : 50;
          const tooltipX = Math.min(Math.max(cx - tooltipWidth / 2, margin.left - 24), width - margin.right - tooltipWidth + 24);
          const tooltipY = Math.max(cy - tooltipHeight - 18, margin.top - 10);
          return `
            <g class="chart-point-tooltip" data-tooltip-index="${index}" transform="translate(${tooltipX} ${tooltipY})">
              <rect width="${tooltipWidth}" height="${tooltipHeight}" rx="7"></rect>
              <text x="12" y="20">${escapeHtml(shortDate(point.date))}</text>
              <text class="tooltip-number" x="12" y="38">
                <tspan>${escapeHtml(formatNumber(point.value))}</tspan>
                ${
                  pointChange
                    ? `<tspan class="tooltip-change ${pointChange.tone}" dx="8">${escapeHtml(pointChange.label)}</tspan>`
                    : ""
                }
              </text>
              ${noteLines
                .map(
                  (line, lineIndex) =>
                    `<text class="tooltip-note" x="12" y="${58 + lineIndex * 14}">${escapeHtml(line)}</text>`,
                )
                .join("")}
            </g>
            <circle class="point ${isEvent ? "event" : ""} ${
              state.selectedDate === point.date ? "selected" : ""
            }" data-point-index="${index}" data-point-date="${point.date}" tabindex="0" aria-label="${escapeHtml(
              `${shortDate(point.date)} ${formatNumber(point.value)}${point.event ? ` ${point.event}` : ""}`,
            )}" cx="${cx}" cy="${cy}" r="${
              state.selectedDate === point.date ? 6 : isEvent ? 5 : 4
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
      state.selectedDate = state.selectedDate === point.dataset.pointDate ? null : point.dataset.pointDate;
      updateDashboard();
    });
  }
}

function renderCampusBarChart(container, rows, { emptyMessage = "No campus detail available for this Sunday." } = {}) {
  if (!container) return;
  if (!rows.length) {
    container.innerHTML = `<div class="empty">${emptyMessage}</div>`;
    return;
  }

  const max = Math.max(...rows.map((row) => row.value), 1);
  container.innerHTML = rows
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

function renderDreamTeamCampusBreakdown(selectedDate) {
  const show = state.metric === "dreamTeam" && state.campus === "All Campuses";
  els.dreamTeamCampusPanel?.classList.toggle("is-hidden", !show);
  if (!show) return;

  const detailRows = dreamTeamCampusRowsFromDetail(selectedDate);
  const rows = detailRows.length ? detailRows : comparisonRowsByCampus("dreamTeam", selectedDate);
  if (els.dreamTeamCampusMeta) {
    els.dreamTeamCampusMeta.textContent = selectedDate
      ? `${rows.length} ${rows.length === 1 ? "campus" : "campuses"} · ${shortDate(selectedDate)}`
      : "";
  }
  renderCampusBarChart(els.dreamTeamCampusChart, rows);
}

function renderBars() {
  const selectedDate = state.metric === "dreamTeam" ? dreamTeamDetailDate() : currentMetricDate(state.metric);
  const isArchive = Boolean(state.selectedDate && selectedDate === state.selectedDate);
  renderDreamTeamCampusBreakdown(selectedDate);
  if (els.barKicker) {
    els.barKicker.textContent =
      state.metric === "dreamTeam" ? "Dream Team Detail" : isArchive ? "Sunday Archive" : "Most Recent Sunday";
  }

  if (state.metric === "dreamTeam") {
    const latestDate = selectedDate;
    const teamRows = latestDate ? dreamTeamTeamTotals(latestDate) : [];
    const campusCount = latestDate
      ? new Set(dreamTeamDetailRows(latestDate, state.campus).map((row) => row.campus)).size
      : 0;
    els.barTitle.textContent =
      state.campus === "All Campuses" ? "Dream Team by Team" : `${state.campus} Dream Team by Team`;
    els.barMeta.textContent = latestDate
      ? `${teamRows.length} ${teamRows.length === 1 ? "team" : "teams"} · ${campusCount} ${
          campusCount === 1 ? "campus" : "campuses"
        } · ${shortDate(latestDate)}`
      : "";
    if (teamRows.length) {
      const max = Math.max(...teamRows.map((row) => row.served), 1);
      els.barChart.innerHTML = teamRows
        .map((row) => {
          const width = Math.max(3, (row.served / max) * 100);
          const targetGap = row.hasTarget && row.target ? row.served - row.target : null;
          return `
            <div class="bar-row">
              <div class="bar-label" title="${row.team}">${row.team}</div>
              <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
              <div class="bar-value" title="${targetGap === null ? "" : `${formatSignedNumber(targetGap)} vs target`}">${formatNumber(row.served)}</div>
            </div>
          `;
        })
        .join("");
      return;
    }

    els.barChart.innerHTML = `
      <div class="empty">No Dream Team team detail found for ${
        latestDate ? shortDate(latestDate) : "this Sunday"
      }. ${dreamTeamDetailDiagnostic(latestDate)}</div>
    `;
    return;
  }

  const rows = comparisonRowsByCampus(state.metric, selectedDate);
  if (!rows.length) {
    els.barTitle.textContent = `${metricLabels[state.metric]} by Campus`;
    els.barMeta.textContent = selectedDate ? shortDate(selectedDate) : "";
    renderCampusBarChart(els.barChart, rows);
    return;
  }
  const latestDate = selectedDate || rows[0]?.date;

  els.barTitle.textContent = `${metricLabels[state.metric]} by Campus`;
  els.barMeta.textContent = latestDate ? shortDate(latestDate) : "";
  renderCampusBarChart(els.barChart, rows);
}

function renderVolatility() {
  const allRows = data.volatility.map(movementStats);
  const rows =
    state.campus === "All Campuses"
      ? allRows
      : allRows.filter((row) => row.campus === state.campus);
  if (!rows.length) {
    els.volatilityChart.innerHTML = `<div class="empty">Connect Excel to load attendance movement.</div>`;
    if (els.volatilityMeta) {
      const archiveDate = activeArchiveDate("attendance");
      const archiveLabel = archiveDate ? `through ${shortDate(archiveDate)}` : "YTD";
      els.volatilityMeta.textContent =
        state.campus === "All Campuses"
          ? `${archiveLabel}, 10%+ from previous Sunday`
          : `${state.campus} · ${archiveLabel}`;
    }
    return;
  }

  if (els.volatilityMeta) {
    const archiveDate = activeArchiveDate("attendance");
    const archiveLabel = archiveDate ? `through ${shortDate(archiveDate)}` : "YTD";
    els.volatilityMeta.textContent =
      state.campus === "All Campuses"
        ? `All campuses · ${archiveLabel}, 10%+ from previous Sunday`
        : `${state.campus} · ${archiveLabel}, 10%+ from previous Sunday`;
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
  const stats = attendanceCampusRowsForContext().filter((row) => isFiniteNumber(row.latest));
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
        body: `This attendance range usually needs a clear focus on Sunday quality, team strength, and next-step follow-up.`,
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
      body: `${topBarrier.campuses.slice(0, 3).join(", ")} are in this attendance range. These campuses may need the same kind of support plan.`,
      severity: topBarrier.max <= 1000 ? "warning" : "info",
    });
  }

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
    body: `${label} is ${statusText}. Watch whether weekend attendance is turning into healthy ${target.ministry} movement.`,
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
        body: `${label} ranges from ${formatHealthPct(lowest.ratio)} at ${lowest.campus} to ${formatHealthPct(highest.ratio)} at ${highest.campus}. Compare campus practices where the gap is widest.`,
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
      body: `${top.label} is ${formatHealthPct(top.ratio)} of attendance against a ${targetText(top.target)} target. This is the clearest ministry lane to strengthen next.`,
      severity: top.ratio < top.target.lower * 0.75 ? "critical" : "warning",
    },
  ];
}

function metricCampusMovementRows(metricKey) {
  const campuses = state.campus === "All Campuses" ? data.campuses : [state.campus];
  return campuses
    .map((campus) => {
      const points = getMetricPoints(metricKey, campus);
      const context = selectedPointContext(points);
      const latest = context.current;
      const previous = context.previous;
      const changePct = context.weekChange;
      return {
        campus,
        latest,
        previous,
        delta: latest && previous ? latest.value - previous.value : null,
        changePct,
      };
    })
    .filter((row) => row.latest && row.previous && isFiniteNumber(row.changePct));
}

function movementThreshold(metricKey) {
  if (metricKey === "attendance" || metricKey === "kids" || metricKey === "dreamTeam") return 10;
  return 25;
}

function metricMovementInsights() {
  if (state.metric === "attendance") return [];
  const rows = metricCampusMovementRows(state.metric);
  if (!rows.length) return [];

  const label = metricLabels[state.metric];
  const threshold = movementThreshold(state.metric);
  const surged = rows
    .filter((row) => row.changePct >= threshold && Math.abs(row.delta) >= 2)
    .sort((a, b) => b.changePct - a.changePct);
  const dropped = rows
    .filter((row) => row.changePct <= -threshold && Math.abs(row.delta) >= 2)
    .sort((a, b) => a.changePct - b.changePct);
  const insights = [];

  const writeCampusPrompt = (row, direction) =>
    state.campus === "All Campuses"
      ? `${row.campus} moved ${formatPct(row.changePct)} (${formatSignedNumber(row.delta)}) from the previous Sunday. Ask what changed and whether it should be repeated or corrected.`
      : `${state.campus} moved ${formatPct(row.changePct)} (${formatSignedNumber(row.delta)}) from the previous Sunday. Review what changed in the weekend experience and reporting.`;

  if (surged.length) {
    const top = surged[0];
    insights.push({
      title:
        state.campus === "All Campuses"
          ? `${top.campus} had a ${label} surge`
          : `${label} surged at ${state.campus}`,
      body: writeCampusPrompt(top, "surge"),
      severity: "info",
    });
  }

  if (dropped.length) {
    const low = dropped[0];
    insights.push({
      title:
        state.campus === "All Campuses"
          ? `${low.campus} had a large ${label} drop`
          : `${label} dropped sharply at ${state.campus}`,
      body: writeCampusPrompt(low, "drop"),
      severity: Math.abs(low.changePct) >= threshold * 1.75 ? "critical" : "warning",
    });
  }

  if (state.campus === "All Campuses") {
    const upCount = rows.filter((row) => row.changePct > 0).length;
    const downCount = rows.filter((row) => row.changePct < 0).length;
    if (upCount || downCount) {
      insights.push({
        title: `${label} moved up at ${upCount} campuses and down at ${downCount}`,
        body: `Strongest up: ${movementCampusList(
          rows.filter((row) => row.changePct > 0).sort((a, b) => b.changePct - a.changePct),
          2,
        ) || "none"}. Largest drop: ${
          movementCampusList(rows.filter((row) => row.changePct < 0).sort((a, b) => a.changePct - b.changePct), 2) ||
          "none"
        }. Use the contrast to find repeatable practices.`,
        severity: downCount > upCount ? "warning" : "info",
      });
    }
  }

  return insights;
}

function attendanceCategoryContributionInsights() {
  const points = getMetricPoints("attendance", state.campus);
  const context = selectedPointContext(points);
  const latest = context.current;
  const previous = context.previous;
  if (!latest || !previous) return [];

  const attendanceChange = pctChange(latest.value, previous.value);
  if (!isFiniteNumber(attendanceChange) || attendanceChange === 0) return [];

  const direction = attendanceChange < 0 ? -1 : 1;
  const movers = movementContextMetrics
    .map((metricKey) => metricChangeOnDates(metricKey, latest.date, previous.date))
    .filter((row) => row.delta !== null && row.delta !== 0 && Math.sign(row.delta) === direction)
    .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta))
    .slice(0, 3);

  if (!movers.length) return [];

  return [
    {
      title: attendanceChange < 0 ? "What likely contributed to the attendance dip" : "What rose with attendance",
      body: `${movers
        .map((row) => `${row.label} ${formatSignedNumber(row.delta)} (${formatPct(row.changePct)})`)
        .join(", ")}. These areas moved in the same direction as attendance and are worth reviewing first.`,
      severity: attendanceChange < 0 ? "warning" : "info",
      category: "focus",
    },
  ];
}

function dreamTeamDetailInsights() {
  if (state.metric !== "dreamTeam") return [];
  const points = getMetricPoints("dreamTeam", state.campus);
  const context = selectedPointContext(points);
  const latest = context.current;
  const previous = context.previous;
  if (!latest) return [];

  const currentTeams = dreamTeamTeamTotals(latest.date);
  if (!currentTeams.length) return [];
  if (!previous) {
    return [
      {
        title: `Serving team detail for ${shortDate(latest.date)}`,
        body: `Largest teams: ${currentTeams
          .slice(0, 3)
          .map((row) => `${row.team} (${formatNumber(row.served)})`)
          .join(", ")}.`,
        severity: "info",
      },
    ];
  }

  const changes = dreamTeamTeamChanges(latest.date, previous.date);
  if (!changes.length) {
    return [
      {
        title: `Serving team detail for ${shortDate(latest.date)}`,
        body: `Largest teams: ${currentTeams
          .slice(0, 3)
          .map((row) => `${row.team} (${formatNumber(row.served)})`)
          .join(", ")}. Add the prior Sunday by team to see which teams drove the change.`,
        severity: "info",
      },
    ];
  }
  const down = changes.filter((row) => row.delta < 0).sort((a, b) => a.delta - b.delta)[0];
  const up = changes.filter((row) => row.delta > 0).sort((a, b) => b.delta - a.delta)[0];
  const targetMiss = changes
    .filter((row) => row.target !== null && row.current < row.target)
    .sort((a, b) => a.current - a.target - (b.current - b.target))[0];
  const insights = [];

  if (down) {
    insights.push({
      title: `${down.team} drove the largest Dream Team drop`,
      body: `${down.team} was ${formatSignedNumber(down.delta)} from the previous Sunday (${formatPct(
        down.changePct,
      )}). Review scheduling, call-outs, and reporting for that team.`,
      severity: Math.abs(down.changePct || 0) >= 25 ? "critical" : "warning",
    });
  }

  if (up) {
    insights.push({
      title: `${up.team} had the strongest serving lift`,
      body: `${up.team} was ${formatSignedNumber(up.delta)} from the previous Sunday (${formatPct(
        up.changePct,
      )}). Find out what helped that team and whether it can be repeated.`,
      severity: "info",
    });
  }

  if (targetMiss) {
    insights.push({
      title: `${targetMiss.team} is below serving target`,
      body: `${targetMiss.team} had ${formatNumber(targetMiss.current)} serving against a target of ${formatNumber(
        targetMiss.target,
      )}. Gap: ${formatSignedNumber(targetMiss.current - targetMiss.target)}.`,
      severity: "warning",
    });
  }

  return insights;
}

function gradeAttendanceTrend(change) {
  if (!isFiniteNumber(change)) return "info";
  if (change <= -10) return "critical";
  if (change < 0) return "warning";
  return "info";
}

function attendanceTrendVerb(change) {
  if (!isFiniteNumber(change)) return "changed";
  if (change < 0) return "dipped";
  if (change > 0) return "grew";
  return "was flat";
}

function formatTrendMove(change) {
  if (!isFiniteNumber(change)) return "--";
  if (change < 0) return `${Math.abs(change).toFixed(1)}%`;
  return formatPct(change);
}

function campusMovementRows() {
  return attendanceCampusRowsForContext()
    .map((row) => ({
      campus: row.campus,
      latest: row.latestDate ? { date: row.latestDate, value: row.latest } : null,
      previous: row.previousDate ? { date: row.previousDate, value: row.previous } : null,
      changePct: row.weekChangePct,
    }))
    .filter((row) => row.latest && row.previous && isFiniteNumber(row.changePct));
}

function movementCampusList(rows, limit = 3) {
  return rows
    .slice(0, limit)
    .map((row) => `${row.campus} (${formatPct(row.changePct)})`)
    .join(", ");
}

function allCampusAttendanceTrendInsights() {
  const points = getMetricPoints("attendance", "All Campuses");
  const context = selectedPointContext(points);
  const latest = context.current;
  const previous = context.previous;
  if (!latest || !previous) return [];

  const change = context.weekChange;
  const verb = attendanceTrendVerb(change);
  const trendSeverity = gradeAttendanceTrend(change);
  const movementRows = campusMovementRows();
  const downRows = movementRows.filter((row) => row.changePct < 0).sort((a, b) => a.changePct - b.changePct);
  const upRows = movementRows.filter((row) => row.changePct > 0).sort((a, b) => b.changePct - a.changePct);
  const flatRows = movementRows.filter((row) => row.changePct === 0);
  const totalCampusCount = movementRows.length;

  const insights = [
    {
      title:
        change === 0
          ? "Attendance was flat from the previous Sunday"
          : `Attendance ${verb} ${formatTrendMove(change)} from the previous Sunday`,
      body:
        change < 0
          ? `Total in-person attendance was ${formatNumber(latest.value)} on ${shortDate(
              latest.date,
            )}, down from ${formatNumber(previous.value)} on ${shortDate(
              previous.date,
            )}. Even a modest negative move is worth naming so it does not read as steady.`
          : `Total in-person attendance was ${formatNumber(latest.value)} on ${shortDate(
              latest.date,
            )}, compared with ${formatNumber(previous.value)} on ${shortDate(previous.date)}.`,
      severity: trendSeverity,
      category: "pulse",
    },
  ];

  if (!totalCampusCount) return insights;

  if (downRows.length === totalCampusCount) {
    insights.push({
      title: "Every campus dipped from the previous Sunday",
      body: `${downRows.length} of ${totalCampusCount} campuses were down. The largest dips were ${movementCampusList(
        downRows,
      )}. Treat this as a system-wide weekend pattern before isolating one campus.`,
      severity: downRows.some((row) => row.changePct <= -10) ? "critical" : "warning",
      category: "pulse",
    });
  } else if (downRows.length >= 2 && upRows.length === 1 && flatRows.length === 0) {
    const exception = upRows[0];
    insights.push({
      title: `Most campuses dipped; ${exception.campus} was the exception`,
      body: `${downRows.length} of ${totalCampusCount} campuses were down from the previous Sunday. ${exception.campus} was up ${formatPct(
        exception.changePct,
      )}, so compare what was different there before treating the dip as unavoidable.`,
      severity: downRows.some((row) => row.changePct <= -10) ? "critical" : "warning",
      category: "pulse",
    });
  } else if (downRows.length > upRows.length) {
    insights.push({
      title: `${downRows.length} of ${totalCampusCount} campuses dipped`,
      body: `Down campuses: ${movementCampusList(downRows)}. Up campuses: ${
        upRows.length ? movementCampusList(upRows) : "none"
      }. This points to a broad weekend pattern with a few campus-specific exceptions.`,
      severity: downRows.length >= Math.ceil(totalCampusCount * 0.75) ? "critical" : "warning",
      category: "pulse",
    });
  } else if (upRows.length === totalCampusCount) {
    insights.push({
      title: "Every campus grew from the previous Sunday",
      body: `${upRows.length} of ${totalCampusCount} campuses were up. The strongest gains were ${movementCampusList(
        upRows,
      )}. Capture what worked while the weekend is still fresh.`,
      severity: "info",
      category: "pulse",
    });
  } else if (upRows.length > downRows.length) {
    insights.push({
      title: `${upRows.length} of ${totalCampusCount} campuses grew`,
      body: `Up campuses: ${movementCampusList(upRows)}. Down campuses: ${
        downRows.length ? movementCampusList(downRows) : "none"
      }. Look for shared drivers among the campuses that moved up.`,
      severity: "info",
      category: "pulse",
    });
  } else {
    insights.push({
      title: "Campus movement was mixed",
      body: `${upRows.length} campuses were up, ${downRows.length} were down, and ${flatRows.length} were flat. Up campuses: ${
        upRows.length ? movementCampusList(upRows) : "none"
      }. Down campuses: ${downRows.length ? movementCampusList(downRows) : "none"}.`,
      severity: downRows.length ? "warning" : "info",
      category: "pulse",
    });
  }

  return insights;
}

function selectedCampusInsights() {
  const points = getMetricPoints("attendance", state.campus);
  const context = selectedPointContext(points);
  const latest = context.current;
  const previous = context.previous;
  const stat = statForCampus(state.campus);
  if (!stat || !latest) return [];
  const weekChange = context.weekChange ?? stat.weekChangePct;
  const weekTone = gradeAttendanceTrend(weekChange);
  const baselineTone =
    (stat.vsPrevious4Pct || 0) <= -10 ? "critical" : (stat.vsPrevious4Pct || 0) < 0 ? "warning" : "info";

  return [
    {
      title: `${state.campus} attendance ${attendanceTrendVerb(weekChange)} this Sunday`,
      body: previous
        ? `${state.campus} reported ${formatNumber(latest.value)} on ${shortDate(
            latest.date,
          )}, ${formatPct(weekChange)} compared with ${formatNumber(previous.value)} on ${shortDate(previous.date)}.`
        : `${state.campus} reported ${formatNumber(latest.value)} on ${shortDate(latest.date)}.`,
      severity: weekTone,
      category: "pulse",
    },
    {
      title:
        (stat.vsPrevious4Pct || 0) < 0
          ? "Below recent regular Sundays"
          : "Compared with recent regular Sundays",
      body: `${state.campus} is ${formatPct(stat.vsPrevious4Pct)} compared with its previous four regular Sundays.`,
      severity: baselineTone,
      category: "pulse",
    },
  ];
}

const attendanceInsightColumns = [
  {
    key: "pulse",
    kicker: "Weekend Pulse",
    title: "What changed",
  },
  {
    key: "focus",
    kicker: "Campus Focus",
    title: "Where to lean in",
  },
  {
    key: "levers",
    kicker: "Growth Levers",
    title: "What to strengthen",
  },
];

function attendanceInsightColumnKey(insight) {
  if (insight.category) return insight.category;
  const title = normalizeText(insight.title);
  if (
    title.includes("growthlever") ||
    title.includes("5keys") ||
    title.includes("barrier") ||
    title.includes("target") ||
    title.includes("emphasis")
  ) {
    return "levers";
  }
  if (
    title.includes("gap") ||
    title.includes("below") ||
    title.includes("campus") ||
    title.includes("opportunity") ||
    title.includes("carries") ||
    title.includes("clearest")
  ) {
    return "focus";
  }
  return "pulse";
}

function renderAttendanceInsights(insights) {
  const grouped = new Map(attendanceInsightColumns.map((column) => [column.key, []]));
  for (const insight of insights) {
    const key = grouped.has(attendanceInsightColumnKey(insight))
      ? attendanceInsightColumnKey(insight)
      : "pulse";
    grouped.get(key).push(insight);
  }

  els.insights.className = "attendance-insight-columns";
  els.insights.innerHTML = attendanceInsightColumns
    .filter((column) => grouped.get(column.key)?.length)
    .map((column) => {
      const columnInsights = grouped.get(column.key);
      return `
        <section class="attendance-insight-column">
          <div class="attendance-insight-column-heading">
            <span class="panel-kicker">${column.kicker}</span>
            <strong>${column.title}</strong>
            <span>${columnInsights.length} ${columnInsights.length === 1 ? "takeaway" : "takeaways"}</span>
          </div>
          <div class="attendance-insight-list">
            ${columnInsights
              .map(
                (insight) => `
                  <article class="insight ${insight.severity || "info"}">
                    <h3 class="insight-title">${insight.title}</h3>
                    <p class="insight-body">${insight.body}</p>
                  </article>
                `,
              )
              .join("")}
          </div>
        </section>
      `;
    })
    .join("");
}

function renderInsights() {
  const baseInsights =
    state.metric !== "attendance"
      ? []
      : state.campus === "All Campuses"
        ? state.selectedDate
          ? []
          : data.insights
              .filter((insight) => !["trend", "volatility", "event"].includes(insight.type))
              .map((insight) => ({
                title: insight.title,
                body: insight.body,
                severity: insight.severity,
              }))
        : selectedCampusInsights();
  const insights = [
    ...(state.metric === "attendance" && state.campus === "All Campuses" ? allCampusAttendanceTrendInsights() : []),
    ...baseInsights,
    ...(state.metric === "attendance" ? attendanceOpportunityInsights() : []),
    ...(state.metric === "attendance" ? attendanceCategoryContributionInsights() : []),
    ...metricMovementInsights(),
    ...dreamTeamDetailInsights(),
    ...metricOpportunityInsights(),
    ...(state.metric === "attendance" ? growthLeverInsights() : []),
  ];

  if (state.metric === "attendance") {
    renderAttendanceInsights(insights);
    return;
  }

  els.insights.className = "insight-list";
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
  els.campusTable.innerHTML = attendanceCampusRowsForContext()
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

function defaultHealthMonth() {
  const months = healthArchiveMonths();
  if (!months.length) return null;
  const current = currentMonthKey();
  if (months.includes(current)) return current;
  return months.at(-1);
}

function selectedHealthMonth() {
  const months = healthArchiveMonths();
  if (!months.length) return null;
  if (!months.includes(state.healthMonth)) {
    state.healthMonth = defaultHealthMonth();
  }
  return state.healthMonth;
}

function previousHealthMonth(month) {
  if (!month) return null;
  return healthArchiveMonths().filter((candidate) => candidate < month).at(-1) || null;
}

function healthArchiveMonths() {
  const months = healthMonths();
  const current = currentMonthKey();
  const available = months.filter((month) => month <= current);
  return available.length ? available : months;
}

function isCurrentHealthMonth(month) {
  return Boolean(month && month === currentMonthKey());
}

function healthMonthDisplay(month) {
  return isCurrentHealthMonth(month) ? `${formatMonth(month)} month-to-date` : formatMonth(month);
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

function normalizeHealthTarget(target) {
  const output = { ...target };
  const unitText = String(output.unit || "").toLowerCase();
  const isPercentTarget = output.unit === "%" || unitText.includes("percent");
  for (const key of ["optimalMin", "optimalMax"]) {
    if (isPercentTarget && output[key] > 0 && output[key] <= 1) {
      output[key] *= 100;
    }
  }
  if (unitText.includes("percent")) output.unit = "%";
  if (output.direction === "in_range") output.direction = "range";
  return output;
}

function healthTarget(key) {
  const normalized = key.toLowerCase().replace(/[^a-z0-9]/g, "");
  const workbookTarget = (data.health?.targets || []).find(
    (target) => String(target.key || "").toLowerCase().replace(/[^a-z0-9]/g, "") === normalized,
  );
  const override = Object.fromEntries(
    Object.entries(workbookTarget || {}).filter(([, value]) => value !== null && value !== undefined && value !== ""),
  );
  return normalizeHealthTarget({ ...(healthTargetDefaults[key] || {}), ...override });
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

function activeDreamTeamForMonth(campuses, month, attendanceAvg) {
  const rows = (data.health?.activeDreamTeam || []).filter(
    (row) => campuses.includes(row.campus) && row.month === month,
  );
  const total = sumNumbers(rows.map((row) => row.activeDreamTeam));
  return {
    total,
    pct: total !== null && attendanceAvg ? (total / attendanceAvg) * 100 : null,
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

const leadershipRoleOrder = ["Director", "Coordinator", "Team Lead"];

function leadershipRoleRank(role) {
  const rank = leadershipRoleOrder.findIndex((item) => item.toLowerCase() === String(role || "").toLowerCase());
  return rank === -1 ? leadershipRoleOrder.length : rank;
}

function leadershipRoleLabel(role) {
  if (String(role).toLowerCase() === "team lead") return "Team Leads";
  return `${role}s`;
}

function campusLeadershipVacancyRows(month) {
  if (!month) return [];
  return data.campuses
    .map((campus) => {
      const leadership = leadershipSummary([campus], month);
      const roleVacancies = leadership.roles
        .filter((role) => isFiniteNumber(role.vacancies) && role.vacancies > 0)
        .sort((a, b) => leadershipRoleRank(a.role) - leadershipRoleRank(b.role));
      const roleSummary = roleVacancies.length
        ? roleVacancies
            .map((role) => `${formatNumber(role.vacancies)} ${leadershipRoleLabel(role.role)}`)
            .join(", ")
        : isFiniteNumber(leadership.vacancies)
          ? "No vacancies reported"
          : "Needs leadership data";
      return {
        campus,
        vacancies: leadership.vacancies,
        filled: leadership.filled,
        target: leadership.target,
        fillPct: leadership.fillPct,
        priorityRank: roleVacancies[0] ? leadershipRoleRank(roleVacancies[0].role) : leadershipRoleOrder.length,
        roleDetails: roleVacancies.map((role) => ({
          role: role.role,
          label: leadershipRoleLabel(role.role),
          vacancies: role.vacancies,
        })),
        roleSummary,
      };
    })
    .filter(
      (row) => isFiniteNumber(row.vacancies) || isFiniteNumber(row.filled) || isFiniteNumber(row.target),
    )
    .sort(
      (a, b) =>
        a.priorityRank - b.priorityRank || (b.vacancies || 0) - (a.vacancies || 0) || a.campus.localeCompare(b.campus),
    );
}

function buildHealthReport(selectedMonth = selectedHealthMonth()) {
  const campuses = selectedHealthCampuses();
  const month = selectedMonth;
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
  const activeDreamTeam = activeDreamTeamForMonth(campuses, month, attendanceAvg);
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
    row("dreamTeamPct", dates.map(() => null), activeDreamTeam.pct, "pct", "Active Dream Team"),
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

function healthRowMap(report) {
  return new Map((report?.rows || []).map((row) => [row.key, row]));
}

function healthNumericValue(row) {
  if (!row) return null;
  if (row.format === "attendanceVsSignups") return row.monthValue?.pct ?? null;
  return typeof row.monthValue === "number" && Number.isFinite(row.monthValue) ? row.monthValue : null;
}

function healthInsightSeverity(tone) {
  if (tone === "critical" || tone === "urgent") return "critical";
  if (tone === "watch" || tone === "negative") return "warning";
  return "info";
}

function formatHealthValueForInsight(row) {
  return row ? formatHealthCell(row.format, row.monthValue) : "--";
}

function healthStatusCounts(report) {
  return report.rows
    .filter((row) => row.status.grade !== null)
    .reduce(
      (counts, row) => {
        counts.total += 1;
        if (row.status.tone === "positive") counts.onTrack += 1;
        if (row.status.tone === "watch") counts.watch += 1;
        if (row.status.tone === "urgent") counts.urgent += 1;
        if (row.status.tone === "critical") counts.critical += 1;
        return counts;
      },
      { total: 0, onTrack: 0, watch: 0, urgent: 0, critical: 0 },
    );
}

function statusChangeInsights(report, previousReport) {
  if (!previousReport) return [];
  const previousRows = healthRowMap(previousReport);
  const changes = report.rows
    .map((row) => {
      const previous = previousRows.get(row.key);
      if (row.status.grade === null || previous?.status.grade === null || previous?.status.grade === undefined) {
        return null;
      }
      return {
        row,
        previous,
        gradeChange: row.status.grade - previous.status.grade,
      };
    })
    .filter(Boolean)
    .filter((item) => item.gradeChange !== 0);

  const worse = changes.filter((item) => item.gradeChange > 0).sort((a, b) => b.gradeChange - a.gradeChange)[0];
  const better = changes.filter((item) => item.gradeChange < 0).sort((a, b) => a.gradeChange - b.gradeChange)[0];

  if (worse) {
    return [
      {
        title: `${worse.row.label} needs attention since last month`,
        body: `The grade moved from ${worse.previous.status.grade} to ${worse.row.status.grade}. ${formatMonth(
          report.month,
        )} is ${formatHealthValueForInsight(worse.row)} against an optimal of ${worse.row.optimal}.`,
        severity: healthInsightSeverity(worse.row.status.tone),
      },
    ];
  }

  if (better) {
    return [
      {
        title: `${better.row.label} improved since last month`,
        body: `The grade moved from ${better.previous.status.grade} to ${better.row.status.grade}. ${formatMonth(
          report.month,
        )} is ${formatHealthValueForInsight(better.row)} against an optimal of ${better.row.optimal}.`,
        severity: "info",
      },
    ];
  }

  return [
    {
      title: "Health grades held steady from last month",
      body: `No graded metric changed status from ${formatMonth(previousReport.month)} to ${formatMonth(report.month)}.`,
      severity: "info",
    },
  ];
}

function healthTakeaways(report) {
  if (!report.month) return [];
  const reportMonthLabel = healthMonthDisplay(report.month);
  const previousMonth = previousHealthMonth(report.month);
  const previousReport = previousMonth ? buildHealthReport(previousMonth) : null;
  const counts = healthStatusCounts(report);
  const rowMap = healthRowMap(report);
  const previousRows = healthRowMap(previousReport);
  const attendance = rowMap.get("attendance");
  const previousAttendance = previousRows.get("attendance");
  const attendanceChange = pctChange(healthNumericValue(attendance), healthNumericValue(previousAttendance));
  const urgentRows = report.rows
    .filter((row) => row.status.grade !== null)
    .sort((a, b) => (b.status.grade || 0) - (a.status.grade || 0));
  const mostUrgent = urgentRows[0];
  const vacancies = rowMap.get("leadershipVacancies");
  const previousVacancies = previousRows.get("leadershipVacancies");
  const vacancyValue = healthNumericValue(vacancies);
  const previousVacancyValue = healthNumericValue(previousVacancies);
  const vacancyChange =
    vacancyValue !== null && previousVacancyValue !== null ? vacancyValue - previousVacancyValue : null;
  const takeaways = [
    {
      title: `${reportMonthLabel} health snapshot`,
      body: counts.total
        ? `${counts.onTrack} on track, ${counts.watch} watch, ${counts.urgent} urgent, and ${counts.critical} critical across ${counts.total} graded health metrics.`
        : "Connect the health input sheets to grade this month.",
      severity: counts.critical || counts.urgent ? "critical" : counts.watch ? "warning" : "info",
    },
  ];

  if (previousReport && attendance) {
    takeaways.push({
      title:
        attendanceChange === null
          ? "Attendance needs last month for comparison"
          : `Average attendance ${attendanceChange < 0 ? "dipped" : attendanceChange > 0 ? "grew" : "held flat"} ${formatPct(
              attendanceChange,
            )} vs last month`,
      body:
        attendanceChange === null
          ? `No clean ${formatMonth(previousReport.month)} attendance average was available for comparison.`
          : `${reportMonthLabel} averaged ${formatNumber(healthNumericValue(attendance))}, compared with ${formatNumber(
              healthNumericValue(previousAttendance),
            )} in ${formatMonth(previousReport.month)}.`,
      severity: attendanceChange !== null && attendanceChange < 0 ? "warning" : "info",
    });
  } else {
    takeaways.push({
      title: "Month archive starts here",
      body: "Once another month is available, this panel will compare the selected month with the prior month.",
      severity: "info",
    });
  }

  if (mostUrgent) {
    takeaways.push({
      title:
        mostUrgent.status.grade <= 1
          ? "No urgent graded health gaps"
          : `Highest attention metric: ${mostUrgent.label}`,
      body:
        mostUrgent.status.grade <= 1
          ? "Every graded metric is currently on track for the selected month."
          : `${mostUrgent.label} is ${formatHealthValueForInsight(mostUrgent)} against an optimal of ${
              mostUrgent.optimal
            }. Current grade: ${mostUrgent.status.grade}, ${mostUrgent.status.label}.`,
      severity: healthInsightSeverity(mostUrgent.status.tone),
    });
  }

  takeaways.push(...statusChangeInsights(report, previousReport));

  if (vacancies) {
    takeaways.push({
      title: `Leadership vacancies: ${formatNumber(vacancyValue)}`,
      body:
        vacancyChange === null
          ? "Vacancy comparison will appear once the previous month has leadership data."
          : vacancyChange === 0
            ? `Leadership vacancies were unchanged from ${formatMonth(previousReport.month)}.`
            : `${Math.abs(vacancyChange)} ${Math.abs(vacancyChange) === 1 ? "vacancy" : "vacancies"} ${
                vacancyChange > 0 ? "added" : "closed"
              } compared with ${formatMonth(previousReport.month)}.`,
      severity: vacancyValue > 0 ? (vacancyValue >= 5 ? "critical" : "warning") : "info",
    });
  }

  return takeaways;
}

function renderHealthInsights(report) {
  const takeaways = healthTakeaways(report);
  if (!takeaways.length) {
    els.healthInsights.innerHTML = `
      <article class="insight">
        <h3 class="insight-title">Connect Excel to load health takeaways</h3>
        <p class="insight-body">The current month, month archive, and month-to-month comparison will appear here once live workbook data is loaded.</p>
      </article>
    `;
    return;
  }
  els.healthInsights.innerHTML = takeaways
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
  const monthSummaryLabel = isCurrentHealthMonth(report.month) ? "Month to Date" : "Month";
  els.healthTableHead.innerHTML = `
    <tr>
      ${headerCell("Metric")}
      ${headerCell("Optimal")}
      ${weekHeaders}
      ${headerCell(monthSummaryLabel)}
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

function renderCampusVacancySnapshot(report) {
  if (!els.leadershipCampusSnapshotWrap || !els.leadershipCampusSnapshot) return;
  const showSnapshot = state.campus === "All Campuses";
  els.leadershipCampusSnapshotWrap.classList.toggle("is-hidden", !showSnapshot);
  if (!showSnapshot) return;

  const rows = campusLeadershipVacancyRows(report.month);
  const totalVacancies = sumNumbers(rows.map((row) => row.vacancies));
  const campusesWithVacancies = rows.filter((row) => (row.vacancies || 0) > 0).length;
  els.leadershipCampusMeta.textContent =
    totalVacancies !== null
      ? `${formatNumber(totalVacancies)} total · ${formatNumber(campusesWithVacancies)} campuses`
      : "Needs data";

  if (!rows.length) {
    els.leadershipCampusSnapshot.innerHTML = `
      <div class="empty">Campus vacancy totals will appear here once leadership data is available.</div>
    `;
    return;
  }

  const maxVacancies = Math.max(...rows.map((row) => row.vacancies || 0), 1);
  els.leadershipCampusSnapshot.innerHTML = rows
    .map((row) => {
      const vacancies = isFiniteNumber(row.vacancies) ? row.vacancies : null;
      const width = vacancies ? Math.max(8, (vacancies / maxVacancies) * 100) : 0;
      const statusClass =
        vacancies === null ? "needs-data" : vacancies === 0 ? "clear" : vacancies >= 5 ? "critical" : "watch";
      const roleDetails = row.roleDetails?.length
        ? `
          <div class="campus-vacancy-roles">
            ${row.roleDetails
              .map(
                (role) => `
                  <span>
                    <strong>${formatNumber(role.vacancies)}</strong>
                    ${escapeHtml(role.label)}
                  </span>
                `,
              )
              .join("")}
          </div>
        `
        : `<div class="campus-vacancy-detail">${escapeHtml(row.roleSummary)}</div>`;
      return `
        <article class="campus-vacancy-card ${statusClass}">
          <div>
            <div class="campus-vacancy-name">${escapeHtml(row.campus)}</div>
            ${roleDetails}
          </div>
          <div class="campus-vacancy-number">${formatNumber(vacancies)}</div>
          <div class="campus-vacancy-track" aria-hidden="true">
            <span style="width:${width}%"></span>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderHealth() {
  syncHealthMonthOptions();
  const report = buildHealthReport();
  els.healthMonthMeta.textContent = report.month
    ? `${state.campus}`
    : "Current month";
  els.healthReportMeta.textContent = report.month
    ? `${healthMonthDisplay(report.month)} · ${state.campus}`
    : "Current month";
  renderHealthInsights(report);
  renderHealthTable(report);
  renderLeadershipVacancyChart(report);
  renderCampusVacancySnapshot(report);
}

function growthHistoryRows() {
  return data.campusGrowthHistory?.rows || [];
}

function growthHistoryYears() {
  return data.campusGrowthHistory?.years || [];
}

function campusAliasTerms(campus) {
  const aliases = {
    Columbia: ["col", "columbia"],
    Flowers: ["flo", "flowers"],
    "Falls Church": ["fc", "fallschurch", "falls church"],
    "Silver Spring": ["ss", "silverspring", "silver spring"],
    "North Meck": ["northmeck", "north meck", "nm", "clt", "concord"],
    "Mint Hill": ["minthill", "mint hill", "mh"],
    BWI: ["bwi"],
    UBC: ["ubc"],
  };
  return aliases[campus] || [campus];
}

function growthHistoryRowForCampus() {
  const rows = growthHistoryRows();
  if (state.campus === "All Campuses") {
    return rows.find((row) => normalizeText(row.campus) === "total") || null;
  }
  return rows.find((row) => normalizeText(row.campus) === normalizeText(state.campus)) || null;
}

function relevantGrowthHistoryNote(year, campus = state.campus) {
  const note = (data.campusGrowthHistory?.notes || []).find((row) => row.year === Number(year))?.note;
  if (!note) return "";
  const lines = String(note)
    .split(/\n+/)
    .map((line) => line.replace(/^-\s*/, "").trim())
    .filter(Boolean);
  if (campus === "All Campuses") return lines.join(" ");
  const terms = campusAliasTerms(campus).map(normalizeText);
  const relevant = lines.filter((line) => {
    const normalized = normalizeText(line);
    return terms.some((term) => normalized.includes(term)) || normalized.includes("allcampus");
  });
  return (relevant.length ? relevant : lines.filter((line) => /snow|closure|closed|broadcast/i.test(line))).join(" ");
}

function renderGrowthHistoryKpis(row, years) {
  const latestYear = years.at(-1);
  const firstYear = years.find((year) => isFiniteNumber(row?.yearly?.[year]));
  const latestValue = latestYear ? row?.yearly?.[latestYear] : null;
  const firstValue = firstYear ? row?.yearly?.[firstYear] : null;
  const hasRangeValues = isFiniteNumber(firstValue) && isFiniteNumber(latestValue);
  const kpis = [
    {
      label: latestYear ? `${latestYear} average` : "Latest average",
      value: formatNumber(latestValue),
      note: state.campus === "All Campuses" ? "All campuses" : state.campus,
    },
    {
      label: "YoY Growth",
      value: formatPct(row?.yoyPct),
      valueClass: toneClass(row?.yoyPct),
      note: "Latest year vs prior year",
    },
    {
      label: "Growth Since Launch",
      value: row?.growthSinceLaunch || "--",
      note: firstYear ? `Compared with ${firstYear}` : "Needs history",
    },
    {
      label: "Net Growth",
      value: formatSignedNumber(row?.netGrowth),
      valueClass: toneClass(row?.netGrowth),
      note: hasRangeValues ? `From ${formatNumber(firstValue)} to ${formatNumber(latestValue)}` : "Needs history",
    },
  ];

  els.growthHistoryKpis.innerHTML = kpis
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

function renderGrowthHistoryCharts(row, years) {
  const latestYear = years.at(-1);
  els.growthHistoryTrendTitle.textContent =
    state.campus === "All Campuses" ? "All Campuses Attendance by Year" : `${state.campus} Attendance by Year`;
  els.growthHistoryTrendMeta.textContent = latestYear ? `Through ${latestYear}` : "";
  renderBarList(
    els.growthHistoryTrendChart,
    years
      .filter((year) => isFiniteNumber(row?.yearly?.[year]))
      .map((year) => ({
        label: String(year),
        value: row.yearly[year],
        note: relevantGrowthHistoryNote(year),
      })),
  );

  const campusRows = growthHistoryRows().filter(
    (item) => normalizeText(item.campus) !== "total" && isFiniteNumber(item.yearly?.[latestYear]),
  );
  els.growthHistoryShareMeta.textContent = latestYear ? String(latestYear) : "";
  renderBarList(
    els.growthHistoryShareChart,
    campusRows
      .sort((a, b) => (b.yearly[latestYear] || 0) - (a.yearly[latestYear] || 0))
      .map((item) => ({
        label: item.campus,
        value: item.yearly[latestYear],
      })),
    { activeLabel: state.campus !== "All Campuses" ? state.campus : undefined },
  );
}

function renderGrowthHistoryInsights(row, years) {
  const latestYear = years.at(-1);
  const priorYear = years.at(-2);
  if (!row || !latestYear) {
    els.growthHistoryInsights.innerHTML = `
      <article class="insight">
        <h3 class="insight-title">Connect Excel to load growth history</h3>
        <p class="insight-body">Campus growth history will appear once the live workbook data is available.</p>
      </article>
    `;
    return;
  }

  const latestValue = row.yearly?.[latestYear];
  const priorValue = priorYear ? row.yearly?.[priorYear] : null;
  const change = pctChange(latestValue, priorValue);
  const note = relevantGrowthHistoryNote(latestYear);
  const rows = growthHistoryRows().filter((item) => normalizeText(item.campus) !== "total");
  const fastest = rows
    .filter((item) => isFiniteNumber(item.yoyPct))
    .sort((a, b) => b.yoyPct - a.yoyPct)[0];
  const slowest = rows
    .filter((item) => isFiniteNumber(item.yoyPct))
    .sort((a, b) => a.yoyPct - b.yoyPct)[0];
  const insights = [
    {
      title: `${latestYear} growth is ${formatPct(change)}`,
      body: `${row.campus} moved from ${formatNumber(priorValue)} in ${priorYear} to ${formatNumber(
        latestValue,
      )} in ${latestYear}.`,
      severity: change < 0 ? "warning" : "info",
    },
  ];

  if (note) {
    insights.push({
      title: `${latestYear} context`,
      body: note,
      severity: "info",
    });
  }

  if (state.campus === "All Campuses" && fastest && slowest && fastest.campus !== slowest.campus) {
    insights.push({
      title: `${fastest.campus} is growing fastest; ${slowest.campus} needs review`,
      body: `${fastest.campus} is ${formatPct(fastest.yoyPct)} YoY while ${slowest.campus} is ${formatPct(
        slowest.yoyPct,
      )}. Ask what is repeatable from the fastest campus and what changed at the campus losing ground.`,
      severity: slowest.yoyPct < 0 ? "warning" : "info",
    });
  }

  if (row.marketSharePct !== null && row.marketSharePct !== undefined) {
    insights.push({
      title: `Current market share: ${formatHealthPct(row.marketSharePct)}`,
      body:
        state.campus === "All Campuses"
          ? "All campuses equals the full in-person attendance base."
          : `${state.campus} represents ${formatHealthPct(row.marketSharePct)} of current in-person attendance.`,
      severity: "info",
    });
  }

  els.growthHistoryInsights.innerHTML = insights
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

function renderGrowthHistoryTable(years) {
  const headers = [
    "Campus",
    ...years.map(String),
    "YoY",
    "Growth Since Launch",
    "CAGR",
    "Net Growth",
    "Market Share",
  ];
  els.growthHistoryTableHead.innerHTML = headers.map((header) => `<th>${header}</th>`).join("");
  els.growthHistoryTable.innerHTML = growthHistoryRows()
    .map(
      (row) => `
        <tr class="${normalizeText(row.campus) === normalizeText(state.campus) || (state.campus === "All Campuses" && normalizeText(row.campus) === "total") ? "active" : ""}">
          <td>${row.campus}</td>
          ${years.map((year) => `<td>${formatNumber(row.yearly?.[year])}</td>`).join("")}
          <td class="${toneClass(row.yoyPct)}">${formatPct(row.yoyPct)}</td>
          <td>${row.growthSinceLaunch || "--"}</td>
          <td class="${toneClass(row.cagrPct)}">${formatPct(row.cagrPct)}</td>
          <td class="${toneClass(row.netGrowth)}">${formatSignedNumber(row.netGrowth)}</td>
          <td>${formatHealthPct(row.marketSharePct)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderGrowthHistory() {
  const years = growthHistoryYears();
  const row = growthHistoryRowForCampus();
  els.growthHistoryMeta.textContent = years.length ? `${years[0]}-${years.at(-1)} · ${state.campus}` : state.campus;
  renderGrowthHistoryKpis(row, years);
  renderGrowthHistoryCharts(row, years);
  renderGrowthHistoryInsights(row, years);
  renderGrowthHistoryTable(years);
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
      const active = options.activeLabel !== undefined && String(row.label) === String(options.activeLabel);
      return `
        <div class="bar-row ${active ? "active" : ""} ${row.note ? "has-note" : ""}" tabindex="${row.note ? "0" : "-1"}">
          <div class="bar-label" title="${row.label}">${row.label}</div>
          <div class="bar-track"><div class="bar-fill ${options.fillClass || ""}" style="width:${width}%"></div></div>
          <div class="bar-value">${options.format ? options.format(row.value) : formatNumber(row.value)}</div>
          ${row.note ? `<div class="bar-note">${escapeHtml(row.note)}</div>` : ""}
        </div>
      `;
    })
    .join("");
}

function bigFiveEventLabel() {
  return data.bigFive.events.find((event) => event.event === state.bigFiveEvent)?.label || "Big 5";
}

function renderBigFiveKpis(records) {
  const selected = selectedBigFiveRecord(records);
  const multiWeekEvent = isMultiWeekBigFiveEvent();
  const hasPostWeeks = (selected?.postWeeks || []).length > 0;
  const comparisonValue = multiWeekEvent ? selected?.yoyCampaignPct : selected?.yoyFeaturedPct;
  const kpis = [];

  if (multiWeekEvent) {
    kpis.push({
      label: "Total attendance",
      value: formatNumber(selected?.campaignTotal),
      note: selected ? `${selected.year} ${bigFiveEventLabel()}` : "No event data",
    });
  }

  kpis.push(
    {
      label: "Main Sunday",
      value: formatNumber(selected?.featuredTotal),
      note: selected?.featuredPhaseLabel || "No main Sunday yet",
    },
    {
      label: "Compared with last year",
      value: formatPct(comparisonValue),
      valueClass: toneClass(comparisonValue),
      note: multiWeekEvent ? "Total attendance vs prior year" : "Main Sunday vs prior year",
    },
    {
      label: "Vs previous Sunday",
      value: formatPct(selected?.liftPct),
      valueClass: toneClass(selected?.liftPct),
      note: "Main Sunday vs previous Sunday",
    },
  );

  if (hasPostWeeks) {
    kpis.push(
    {
        label: "Post Week 1 Hold",
      value: formatPct(selected?.postRetentionPct),
      valueClass:
        selected?.postRetentionPct !== null && selected?.postRetentionPct !== undefined
          ? selected.postRetentionPct < 50
            ? "negative"
            : "positive"
          : "neutral",
        note: "Post Week 1 vs main Sunday",
    },
      {
        label: "4-Week Hold",
        value: formatPct(selected?.postFourWeekRetentionPct),
        valueClass:
          selected?.postFourWeekRetentionPct !== null && selected?.postFourWeekRetentionPct !== undefined
            ? selected.postFourWeekRetentionPct < 70
              ? "warning"
              : "positive"
            : "neutral",
        note: `${formatPct(selected?.postFourWeekGrowthPct)} vs pre-event Sundays`,
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
  const selected = selectedBigFiveRecord(records);
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
      note: relevantGrowthHistoryNote(record.year),
    })),
    { activeLabel: selected?.year },
  );

  els.bigFivePhaseTitle.textContent = selected ? `${selected.year} Event Weeks` : "Event Weeks";
  els.bigFivePhaseMeta.textContent = selected
    ? `${shortDate(selected.startDate)} - ${shortDate(selected.endDate)}`
    : "";
  renderBarList(
    els.bigFivePhaseChart,
    (selected?.phaseTotals || []).map((phase) => ({
      label: phase.phaseLabel,
      value: phase.total,
    })),
    { fillClass: "vol-fill" },
  );
}

const bigFiveInsightSections = [
  {
    key: "results",
    kicker: "Event Scorecard",
    title: "What happened",
  },
  {
    key: "return",
    kicker: "Return Momentum",
    title: "Who came back after the event",
  },
  {
    key: "next",
    kicker: "Next Big 5 Focus",
    title: "What to aim for next",
  },
  {
    key: "campus",
    kicker: "Campus Playbook",
    title: "Where to learn or lean in",
  },
];

function bigFiveInsightSectionKey(insight) {
  if (insight.category) return insight.category;
  const title = normalizeText(insight.title);
  if (
    title.includes("postweek") ||
    title.includes("following") ||
    title.includes("handoff") ||
    title.includes("fourweek") ||
    title.includes("held")
  ) {
    return "return";
  }
  if (
    title.includes("goal") ||
    title.includes("growthplan") ||
    title.includes("pattern") ||
    title.includes("growthbooster")
  ) {
    return "next";
  }
  if (
    title.includes("campus") ||
    title.includes("learnfrom") ||
    title.includes("share") ||
    title.includes("support") ||
    title.includes("needs")
  ) {
    return "campus";
  }
  return "results";
}

function renderBigFiveInsightSections(insights) {
  const grouped = new Map(bigFiveInsightSections.map((section) => [section.key, []]));
  for (const insight of insights) {
    const key = grouped.has(bigFiveInsightSectionKey(insight)) ? bigFiveInsightSectionKey(insight) : "results";
    grouped.get(key).push(insight);
  }

  els.bigFiveInsights.innerHTML = bigFiveInsightSections
    .filter((section) => grouped.get(section.key)?.length)
    .map((section) => {
      const sectionInsights = grouped.get(section.key);
      return `
        <details class="big-five-insight-section" open>
          <summary>
            <span class="big-five-insight-heading">
              <span class="panel-kicker">${section.kicker}</span>
              <strong>${section.title}</strong>
              <span>${sectionInsights.length} ${sectionInsights.length === 1 ? "insight" : "insights"}</span>
            </span>
          </summary>
          <div class="insight-list insight-grid">
            ${sectionInsights
              .map(
                (insight) => `
                  <article class="insight ${insight.severity || "info"}">
                    <h3 class="insight-title">${insight.title}</h3>
                    <p class="insight-body">${insight.body}</p>
                  </article>
                `,
              )
              .join("")}
          </div>
        </details>
      `;
    })
    .join("");
}

function renderBigFiveInsights(records) {
  const latest = selectedBigFiveRecord(records);
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
  const recentRecords = records
    .filter((record) => record.campaignTotal > 0 && record.year <= latest.year)
    .slice(-3);
  const avgLift = averageNumbers(recentRecords.map((record) => record.liftPct));
  const avgRetention = averageNumbers(recentRecords.map((record) => record.postRetentionPct));
  const campusRetentionRows = (latest.campusSummaries || []).filter((row) => isFiniteNumber(row.postRetentionPct));
  const lowRetentionCampus = campusRetentionRows.slice().sort((a, b) => a.postRetentionPct - b.postRetentionPct)[0];
  const strongRetentionCampus = campusRetentionRows.slice().sort((a, b) => b.postRetentionPct - a.postRetentionPct)[0];
  const bestBooster = records
    .filter(
      (record) =>
        record.year <= latest.year &&
        isFiniteNumber(record.liftPct) &&
        isFiniteNumber(record.postRetentionPct),
    )
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
      title: "Event readiness focus",
      body: "Use the goal to check Sunday quality, staffing, communication, and next-step follow-up before the event.",
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

  renderBigFiveInsightSections(insights);
}

function renderBigFiveTable(records) {
  const multiWeekEvent = isMultiWeekBigFiveEvent();
  const hasPostWeeks = records.some((record) => (record.postWeeks || []).length > 0);
  const selectedYear = selectedBigFiveYear(records);
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
        <tr data-big-five-year="${record.year}" class="${record.year === selectedYear ? "active" : ""}">
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

  for (const row of els.bigFiveTable.querySelectorAll("[data-big-five-year]")) {
    row.addEventListener("click", () => {
      state.bigFiveYear = Number(row.dataset.bigFiveYear) || null;
      updateDashboard();
    });
  }
}

function renderBigFive() {
  const records = summarizeBigFiveEvent();
  els.bigFiveEventSelect.value = state.bigFiveEvent;
  syncBigFiveYearOptions(records);
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

function isGrowthHistorySelected() {
  return state.metric === "campusGrowthHistory";
}

function isStandardMetricSelected() {
  return !isHealthReportSelected() && !isBigFiveSelected() && !isGrowthHistorySelected();
}

function renderMetricPanels() {
  const showHealth = isHealthReportSelected();
  const showBigFive = isBigFiveSelected();
  const showGrowthHistory = isGrowthHistorySelected();
  const showStandard = isStandardMetricSelected();
  for (const panel of els.metricPanels) {
    const mode = panel.dataset.metricPanel;
    const visible =
      (mode === "health" && showHealth) ||
      (mode === "big-five" && showBigFive) ||
      (mode === "growth-history" && showGrowthHistory) ||
      (mode === "standard" && showStandard);
    panel.classList.toggle("is-hidden", !visible);
  }
  for (const panel of els.attendanceOnlyPanels) {
    panel.classList.toggle("is-hidden", state.metric !== "attendance");
  }
  els.dreamTeamCampusPanel?.classList.add("is-hidden");
  els.trendInsightLayout?.classList.toggle("attendance-stack", state.metric === "attendance");
  els.comparisonGrid?.classList.toggle("single-panel", state.metric !== "attendance");
  els.eventToggleWrapper?.classList.toggle("is-hidden", !showStandard);
}

function syncTakeawayLayout() {
  if (!els.trendInsightLayout || !els.chartPanel || !els.insightsPanel) return;
  els.trendInsightLayout.classList.remove("takeaways-below");
  if (!isStandardMetricSelected() || state.metric === "attendance") return;

  const figcaption = els.chartPanel.querySelector("figcaption");
  const naturalChartHeight =
    (figcaption?.scrollHeight || 0) + (els.lineChart?.scrollHeight || 0) + 28;
  const takeawayHeight = els.insightsPanel.scrollHeight;
  els.trendInsightLayout.classList.toggle("takeaways-below", takeawayHeight > naturalChartHeight + 24);
}

function renderTrendMeta(context) {
  const current = context.current;
  if (!current) {
    els.trendMeta.textContent = "";
    return;
  }

  if (context.isArchive) {
    els.trendMeta.innerHTML = `
      <span class="archive-meta">
        <span>Sunday Archive: ${formatDate(current.date)}</span>
        <button class="archive-reset" type="button">Back to latest</button>
      </span>
    `;
    els.trendMeta.querySelector(".archive-reset")?.addEventListener("click", () => {
      state.selectedDate = null;
      updateDashboard();
    });
    return;
  }

  els.trendMeta.textContent = `${formatNumber(current.value)} most recent, ${formatPct(
    context.weekChange,
  )} since last Sunday`;
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
  if (isGrowthHistorySelected()) {
    renderGrowthHistory();
    return;
  }

  const points = getMetricPoints(state.metric, state.campus);
  const context = selectedPointContext(points);

  els.trendTitle.textContent =
    state.campus === "All Campuses"
      ? `${metricLabels[state.metric]} Over Time`
      : `${state.campus} ${metricLabels[state.metric]}`;
  renderTrendMeta(context);

  renderKpis(points);
  renderLineChart(points);
  renderBars();
  if (state.metric === "attendance") renderVolatility();
  renderInsights();
  syncTakeawayLayout();
  renderHealth();
  if (state.metric === "attendance") renderTable();
}

setupControls();
updateDashboard();
