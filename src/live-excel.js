const SHAREPOINT_WORKBOOK_URL =
  "https://theunionchurch.sharepoint.com/:x:/s/UnionChurch/IQC8762dS2Q1RaHRsGED0SxPAcLDsAzuNmjaz9K7vzNg9NU?e=iAChfN";

const MSAL_CONFIG = {
  auth: {
    clientId: "278dcce1-6a0a-4793-b0be-a444d5faed03",
    authority: "https://login.microsoftonline.com/b98a50a2-9b8b-4900-81ad-a1c0d95cf816",
    redirectUri: getRedirectUri(),
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

const LOGIN_REQUEST = {
  scopes: ["User.Read", "Files.ReadWrite"],
};

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

function getRedirectUri() {
  if (window.location.hostname === "localhost") {
    return window.location.origin;
  }
  return `${window.location.origin}${window.location.pathname}`;
}

const CORE_METRIC_SHEETS = {
  Attendance: "attendance",
  Kids: "kids",
  "Growth Track": "growthTrack",
  Salvations: "salvations",
  "First Timers": "firstTimers",
  "Sunday - Dream Team": "dreamTeam",
};

const OPTIONAL_METRIC_SHEETS = {
  Baptism: "baptism",
};

const METRIC_SHEETS = {
  ...CORE_METRIC_SHEETS,
  ...OPTIONAL_METRIC_SHEETS,
};

const DISPLAY_NAMES = {
  attendance: "Attendance",
  kids: "Kids",
  growthTrack: "Growth Track",
  baptism: "Baptisms",
  salvations: "Salvations",
  firstTimers: "First Timers",
  dreamTeam: "Dream Team",
};

let msalClient;
let activeAccount;

export function setupLiveExcel({ onData, onStatus, onReady }) {
  if (window.location.protocol === "file:") {
    onStatus("Open this from http://localhost:4173 to connect Excel. File links can only show static data.", "error");
    onReady(false);
    return;
  }

  if (["127.0.0.1", "::1"].includes(window.location.hostname)) {
    onStatus("Use http://localhost:4173 instead of 127.0.0.1 so Microsoft sign-in matches the setup.", "error");
    onReady(false);
    return;
  }

  if (!window.msal) {
    onStatus("Microsoft sign-in library did not load. Check internet access.", "error");
    return;
  }

  msalClient = new window.msal.PublicClientApplication(MSAL_CONFIG);
  const accounts = msalClient.getAllAccounts();
  activeAccount = accounts[0] || null;
  onReady(Boolean(activeAccount));

  return {
    async connect() {
      await loadLiveWorkbook({ onData, onStatus, forceLogin: true });
      onReady(Boolean(activeAccount));
    },
    async refresh() {
      await loadLiveWorkbook({ onData, onStatus, forceLogin: false });
      onReady(Boolean(activeAccount));
    },
  };
}

async function loadLiveWorkbook({ onData, onStatus, forceLogin }) {
  try {
    onStatus(forceLogin ? "Opening Microsoft sign-in..." : "Refreshing Excel data...");
    const token = await getAccessToken(forceLogin);
    onStatus("Resolving SharePoint workbook...");
    const driveItem = await resolveSharedWorkbook(token);
    onStatus(`Reading ${driveItem.name || "workbook"}...`);
    const liveData = await readWorkbookData(token, driveItem);
    onData(liveData);
    onStatus(`Live Excel connected: ${driveItem.name || "SharePoint workbook"}`);
  } catch (error) {
    console.error(error);
    onStatus(readableError(error), "error");
    throw error;
  }
}

async function getAccessToken(forceLogin) {
  if (!activeAccount || forceLogin) {
    const login = await msalClient.loginPopup(LOGIN_REQUEST);
    activeAccount = login.account;
    msalClient.setActiveAccount(activeAccount);
  }

  try {
    const token = await msalClient.acquireTokenSilent({
      ...LOGIN_REQUEST,
      account: activeAccount,
    });
    return token.accessToken;
  } catch (error) {
    const token = await msalClient.acquireTokenPopup(LOGIN_REQUEST);
    activeAccount = token.account;
    msalClient.setActiveAccount(activeAccount);
    return token.accessToken;
  }
}

async function resolveSharedWorkbook(token) {
  const shareId = toSharingToken(SHAREPOINT_WORKBOOK_URL);
  return graphJson(token, `/shares/${shareId}/driveItem`);
}

function toSharingToken(url) {
  const encoded = btoa(url).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  return `u!${encoded}`;
}

async function readWorkbookData(token, driveItem) {
  const driveId = driveItem.parentReference?.driveId;
  const itemId = driveItem.id;
  if (!driveId || !itemId) {
    throw new Error("Graph did not return a driveId/itemId for the shared workbook.");
  }

  const sessionId = await createWorkbookSession(token, driveId, itemId);
  const readRange = (sheetName, address) =>
    getWorkbookRange(token, driveId, itemId, sheetName, address, sessionId);
  const tryReadRange = async (sheetName, address) => {
    try {
      return await readRange(sheetName, address);
    } catch (error) {
      console.warn(`Optional sheet not loaded: ${sheetName}`, error);
      return null;
    }
  };
  const tryReadFirstRange = async (sheetNames, address) => {
    for (const sheetName of sheetNames) {
      const range = await tryReadRange(sheetName, address);
      if (range) return range;
    }
    return null;
  };

  const sheets = {};
  for (const sheetName of Object.keys(CORE_METRIC_SHEETS)) {
    sheets[sheetName] = await readRange(sheetName, "A1:BW70");
  }
  sheets.Baptism = await tryReadFirstRange(["Baptism", "Baptisms"], "A1:BW70");
  sheets["2025"] = await readRange("2025", "A1:BX71");
  sheets["Big 5 Historical Raw Data"] = await readRange("Big 5 Historical Raw Data", "A1:F700");
  sheets["Big 5 Historical Data"] = await readRange("Big 5 Historical Data", "A1:J40");
  sheets["Campus Growth History"] = await tryReadRange("Campus Growth History", "A1:Z150");
  sheets["Sunday - DT Detail"] = await tryReadFirstRange(
    ["Sunday - DT Detail", "Sunday DT Detail", "Dream Team Detail"],
    "A1:BW1200",
  );
  sheets["Active Dream Team"] = await tryReadRange("Active Dream Team", "A1:Z80");
  sheets["Health Targets"] = await tryReadRange("Health Targets", "A1:H150");
  sheets["Group Semester Config"] = await tryReadRange("Group Semester Config", "A1:K500");
  sheets["Group Health Input"] = await tryReadRange("Group Health Input", "A1:L1200");
  sheets["Group Attendance"] = await tryReadFirstRange(
    ["Group Attendance", "Groups Attendance", "Group Attendance Input", "Groups Health Input"],
    "A1:BW120",
  );
  sheets["Heart Soul Input"] = await tryReadFirstRange(
    ["Heart Soul Input", "Heart & Soul Input", "Heart and Soul Input"],
    "A1:J700",
  );
  sheets["Leadership Input"] = await tryReadFirstRange(
    ["Leadership Input", "Leadership Health Input", "Leadership Health"],
    "A1:I700",
  );

  return buildDashboardData(sheets, driveItem);
}

async function createWorkbookSession(token, driveId, itemId) {
  try {
    const response = await graphJson(
      token,
      `/drives/${driveId}/items/${itemId}/workbook/createSession`,
      {
        method: "POST",
        body: JSON.stringify({ persistChanges: false }),
        headers: { "Content-Type": "application/json" },
      },
    );
    return response.id;
  } catch (error) {
    console.warn("Continuing without workbook session", error);
    return null;
  }
}

async function getWorkbookRange(token, driveId, itemId, sheetName, address, sessionId) {
  const headers = sessionId ? { "Workbook-Session-Id": sessionId } : {};
  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${encodeURIComponent(
    sheetName,
  )}/range(address='${address}')`;
  return graphJson(token, path, { headers });
}

async function graphJson(token, path, options = {}) {
  const response = await fetch(`${GRAPH_ROOT}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    let detail = "";
    try {
      const body = await response.json();
      detail = body.error?.message || JSON.stringify(body.error || body);
    } catch {
      detail = await response.text();
    }
    throw new Error(`Microsoft Graph ${response.status}: ${detail}`);
  }

  return response.status === 204 ? null : response.json();
}

function buildDashboardData(sheets, driveItem) {
  const metrics = {};
  for (const [sheetName, key] of Object.entries(METRIC_SHEETS)) {
    metrics[key] = extractMetricSheet(sheets[sheetName]);
  }

  const attendanceSeries = metrics.attendance.series;
  const campuses = Object.keys(attendanceSeries);
  const latestDate = latestCommonDate(attendanceSeries);
  const campusStats = campusStatistics(attendanceSeries, latestDate);
  const totals = totalSeries(attendanceSeries);
  const yoy = extract2025Yoy(sheets["2025"], campuses, latestDate);
  const attendance2025 = extractMetricSheet(sheets["2025"]).series;
  const bigFive = extractBigFive(
    sheets["Big 5 Historical Raw Data"],
    sheets["Big 5 Historical Data"],
  );
  const dreamTeamDetail = extractDreamTeamDetail(sheets["Sunday - DT Detail"]);
  const campusGrowthHistory = extractCampusGrowthHistory(sheets["Campus Growth History"]);
  const health = extractHealthData(sheets, metrics, latestDate);

  return {
    source: {
      workbook: driveItem.name || "SharePoint workbook",
      generatedFrom: SHAREPOINT_WORKBOOK_URL,
      latestAttendanceDate: latestDate,
    },
    campuses,
    metrics: Object.fromEntries(
      Object.entries(metrics).map(([key, payload]) => [
        key,
        { label: DISPLAY_NAMES[key], series: payload.series },
      ]),
    ),
    history: {
      attendance2025,
    },
    latest: {
      date: latestDate,
      attendanceByCampus: valuesOnDate(attendanceSeries, latestDate),
      metricSnapshots: latestMetricSnapshots(metrics, latestDate),
    },
    totals,
    campusStats: campusStats.sort((a, b) => (b.latest || 0) - (a.latest || 0)),
    volatility: campusStats
      .slice()
      .sort((a, b) => b.swing10Count - a.swing10Count || b.avgAbsSwingPct - a.avgAbsSwingPct),
    yoy,
    bigFive,
    dreamTeamDetail,
    campusGrowthHistory,
    health,
    insights: buildInsights(campusStats, totals, yoy),
  };
}

function matrix(range, key) {
  return range?.[key] || [];
}

function cell(range, key, row, col) {
  return matrix(range, key)?.[row]?.[col];
}

function cleanText(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).replace(/\u00a0/g, " ").trim().replace(/\s+/g, " ");
  return text || null;
}

function displayLabel(value) {
  const text = cleanText(value) || "";
  const match = text.match(/^\d+\s+(.*)$/);
  return match ? match[1] : text;
}

function asNumber(value, textValue) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = cleanText(textValue ?? value);
  if (!text || text === "-") return null;
  const parsed = Number(text.replace(/,/g, "").replace("%", ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function asDate(value, textValue) {
  if (typeof value === "number" && Number.isFinite(value) && value > 20000) {
    return excelSerialToIso(value);
  }

  const raw = cleanText(textValue ?? value);
  if (!raw) return null;
  const text = raw.includes("-") && raw.includes("/") ? raw.split("-").at(-1).trim() : raw;
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;
  return toIsoDate(parsed);
}

function excelSerialToIso(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400 * 1000;
  return toIsoDate(new Date(utcValue));
}

function toIsoDate(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseIsoDate(iso) {
  const [year, month, day] = iso.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function isSunday(iso) {
  return parseIsoDate(iso).getDay() === 0;
}

function normalizeValue(value) {
  if (Number.isInteger(value)) return value;
  return Math.round(value * 100) / 100;
}

function pctChange(current, previous) {
  if (current === null || current === undefined || !previous) return null;
  return (current - previous) / previous;
}

function pct(value) {
  return value === null || value === undefined || Number.isNaN(value)
    ? null
    : Math.round(value * 1000) / 10;
}

function percentValue(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return null;
  const normalized = Math.abs(value) > 1 ? value : value * 100;
  return Math.round(normalized * 10) / 10;
}

function average(values) {
  const usable = values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (!usable.length) return null;
  return usable.reduce((sum, value) => sum + value, 0) / usable.length;
}

function classifyEvent(label) {
  const text = (label || "").toLowerCase();
  if (["easter", "mother", "father", "christmas"].some((token) => text.includes(token))) {
    return "holiday";
  }
  if (["snow", "closed", "rain"].some((token) => text.includes(token))) return "weather";
  if (["daylight", "memorial", "4th", "superbowl"].some((token) => text.includes(token))) {
    return "calendar";
  }
  return text ? "program" : "normal";
}

function extractMetricSheet(range) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  const headers = [];

  for (let col = 1; col < (values[1]?.length || 0); col += 1) {
    const dateIso = asDate(values[1]?.[col], text[1]?.[col]);
    if (!dateIso) continue;
    const event = cleanText(text[0]?.[col] || values[0]?.[col]);
    headers.push({
      col,
      date: dateIso,
      event,
      eventType: classifyEvent(event),
      isSunday: isSunday(dateIso),
    });
  }

  const series = {};
  for (let row = 2; row < values.length; row += 1) {
    const campus = cleanText(text[row]?.[0] || values[row]?.[0]);
    if (!campus) continue;
    const lower = campus.toLowerCase();
    if (
      lower.startsWith("total") ||
      lower.startsWith("% swing") ||
      lower.startsWith("monthly") ||
      lower.startsWith("quarter")
    ) {
      break;
    }

    series[campus] = headers
      .map((header) => {
        const value = asNumber(values[row]?.[header.col], text[row]?.[header.col]);
        if (value === null) return null;
        return {
          date: header.date,
          value: normalizeValue(value),
          event: header.event,
          eventType: header.eventType,
          isSunday: header.isSunday,
        };
      })
      .filter(Boolean);
  }

  return { headers, series };
}

function latestCommonDate(seriesByCampus) {
  const dates = new Set();
  for (const points of Object.values(seriesByCampus)) {
    for (const point of points) {
      if (point.isSunday && typeof point.value === "number" && Number.isFinite(point.value) && point.value > 0) {
        dates.add(point.date);
      }
    }
  }
  return Array.from(dates).sort().at(-1) || null;
}

function valuesOnDate(seriesByCampus, selectedDate) {
  const output = {};
  for (const [campus, points] of Object.entries(seriesByCampus)) {
    const point = points.find((item) => item.date === selectedDate);
    if (point) output[campus] = point.value;
  }
  return output;
}

function positiveSundays(points) {
  return points.filter((point) => point.isSunday && typeof point.value === "number" && Number.isFinite(point.value) && point.value > 0);
}

function campusStatistics(attendanceSeries, latestDate) {
  return Object.entries(attendanceSeries).map(([campus, points]) => {
    const sundays = positiveSundays(points);
    const latestIndex = Math.max(
      0,
      sundays.findIndex((point) => point.date === latestDate),
    );
    const latest = sundays[latestIndex] || sundays.at(-1);
    const previous = latestIndex > 0 ? sundays[latestIndex - 1] : null;
    const previousFour = sundays
      .slice(0, latestIndex || sundays.length)
      .filter((point) => !["holiday", "weather"].includes(point.eventType))
      .slice(-4);

    const swings = [];
    for (let i = 1; i < sundays.length; i += 1) {
      const prev = sundays[i - 1];
      const cur = sundays[i];
      const change = pctChange(cur.value, prev.value);
      if (change === null) continue;
      swings.push({
        fromDate: prev.date,
        toDate: cur.date,
        fromValue: prev.value,
        toValue: cur.value,
        pct: pct(change),
        absPct: Math.round(Math.abs(change) * 1000) / 10,
        event: cur.event,
        eventType: cur.eventType,
      });
    }

    const swing10 = swings.filter((swing) => swing.absPct >= 10);
    const positiveSwing10 = swing10.filter((swing) => swing.pct >= 10);
    const negativeSwing10 = swing10.filter((swing) => swing.pct <= -10);
    const maxSwing = swings.slice().sort((a, b) => b.absPct - a.absPct)[0] || null;
    const prev4Avg = average(previousFour.map((point) => point.value));
    return {
      campus,
      sundayCount: sundays.length,
      latest: latest?.value ?? null,
      latestDate: latest?.date ?? null,
      previous: previous?.value ?? null,
      previousDate: previous?.date ?? null,
      weekChangePct: pct(pctChange(latest?.value, previous?.value)),
      previous4Avg: prev4Avg ? Math.round(prev4Avg * 10) / 10 : null,
      vsPrevious4Pct: pct(pctChange(latest?.value, prev4Avg)),
      swing10Count: swing10.length,
      positiveSwing10Count: positiveSwing10.length,
      negativeSwing10Count: negativeSwing10.length,
      avgAbsSwingPct: Math.round((average(swings.map((swing) => swing.absPct)) || 0) * 10) / 10,
      avgAbsSwing10Pct:
        Math.round((average(swing10.map((swing) => swing.absPct)) || 0) * 10) / 10,
      maxSwing,
      swings,
    };
  });
}

function totalSeries(attendanceSeries) {
  const totals = new Map();
  const metadata = new Map();
  for (const points of Object.values(attendanceSeries)) {
    for (const point of points) {
      if (!point.isSunday) continue;
      totals.set(point.date, (totals.get(point.date) || 0) + point.value);
      metadata.set(point.date, {
        event: point.event,
        eventType: point.eventType,
        isSunday: point.isSunday,
      });
    }
  }

  return Array.from(totals.entries())
    .filter(([, value]) => value > 0)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, value]) => ({
      date,
      value: Math.round(value),
      event: metadata.get(date)?.event,
      eventType: metadata.get(date)?.eventType || "normal",
    }));
}

function latestMetricSnapshots(metrics, latestDate) {
  return Object.entries(metrics).map(([key, payload]) => {
    const byCampus = valuesOnDate(payload.series, latestDate);
    const total = Object.values(byCampus).reduce((sum, value) => sum + value, 0);
    return {
      key,
      label: DISPLAY_NAMES[key],
      total: normalizeValue(total),
      byCampus,
    };
  });
}

function extract2025Yoy(range, campuses, latestDate) {
  if (!latestDate) return null;
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  const target = parseIsoDate(latestDate);
  target.setFullYear(2025);

  let best = null;
  for (let col = 1; col < (values[1]?.length || 0); col += 1) {
    const dateIso = asDate(values[1]?.[col], text[1]?.[col]);
    if (!dateIso || !isSunday(dateIso)) continue;
    const date = parseIsoDate(dateIso);
    const distance = Math.abs(date - target);
    if (!best || distance < best.distance) best = { col, dateIso, distance };
  }
  if (!best) return null;

  const byCampus = {};
  for (let row = 2; row < values.length; row += 1) {
    const campus = cleanText(text[row]?.[0] || values[row]?.[0]);
    if (!campus || campus === "Total" || !campuses.includes(campus)) continue;
    const value = asNumber(values[row]?.[best.col], text[row]?.[best.col]);
    if (value) byCampus[campus] = Math.round(value);
  }

  return {
    matchedDate: best.dateIso,
    byCampus,
    total: Object.values(byCampus).reduce((sum, value) => sum + value, 0),
  };
}

function extractHealthData(sheets, metrics, latestDate) {
  const groupConfig = extractGroupConfig(sheets["Group Semester Config"]);
  const groupAttendance = extractGroupAttendance(
    sheets["Group Health Input"],
    sheets["Group Attendance"],
  );
  const heartSoulRows = extractHeartSoulRows(sheets["Heart Soul Input"]);
  const leadershipRows = extractLeadershipRows(sheets["Leadership Input"]);
  const activeDreamTeam = extractActiveDreamTeam(sheets["Active Dream Team"], latestDate);
  const targets = extractHealthTargets(sheets["Health Targets"]);

  return {
    targets,
    groupConfig,
    groupAttendance,
    heartSoulRows,
    leadershipRows,
    activeDreamTeam,
    months: extractHealthMonths(metrics, groupConfig, groupAttendance, heartSoulRows, leadershipRows, activeDreamTeam),
  };
}

function normalizeHeader(value) {
  const text = cleanText(value);
  if (!text) return null;
  return text
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/%/g, " percent ")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function tableRows(range, requiredHeaders = []) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  if (!values.length) return [];

  const required = requiredHeaders.map(normalizeHeader).filter(Boolean);
  let headerRow = -1;
  let headerKeys = [];

  for (let row = 0; row < values.length; row += 1) {
    const keys = (values[row] || []).map((_, col) =>
      normalizeHeader(text[row]?.[col] || values[row]?.[col]),
    );
    const usable = keys.filter(Boolean);
    if (!required.length && usable.length >= 2) {
      headerRow = row;
      headerKeys = keys;
      break;
    }
    if (required.length && required.every((key) => usable.includes(key))) {
      headerRow = row;
      headerKeys = keys;
      break;
    }
  }

  if (headerRow < 0) return [];

  const rows = [];
  for (let row = headerRow + 1; row < values.length; row += 1) {
    const record = { $text: {} };
    let hasValue = false;
    for (let col = 0; col < headerKeys.length; col += 1) {
      const key = headerKeys[col];
      if (!key) continue;
      const value = values[row]?.[col];
      const textValue = text[row]?.[col];
      if (cleanText(value) || cleanText(textValue)) hasValue = true;
      record[key] = value;
      record.$text[key] = textValue;
    }
    if (hasValue) rows.push(record);
  }
  return rows;
}

function rowFieldKey(row, names) {
  const keys = names.map(normalizeHeader).filter(Boolean);
  return (
    keys.find((key) => key in row && cleanText(row.$text[key] ?? row[key])) ||
    keys.find((key) => key in row) ||
    null
  );
}

function rowText(row, names) {
  const key = rowFieldKey(row, names);
  return key ? cleanText(row.$text[key] ?? row[key]) : null;
}

function rowNumber(row, names) {
  const key = rowFieldKey(row, names);
  return key ? asNumber(row[key], row.$text[key]) : null;
}

function rowDate(row, names) {
  const key = rowFieldKey(row, names);
  return key ? asDate(row[key], row.$text[key]) : null;
}

function rowBool(row, names) {
  const value = rowText(row, names);
  if (!value) return null;
  const normalized = value.toLowerCase();
  if (["yes", "y", "true", "1", "included"].includes(normalized)) return true;
  if (["no", "n", "false", "0", "not included"].includes(normalized)) return false;
  return null;
}

function monthKeyFromIso(iso) {
  return iso ? iso.slice(0, 7) : null;
}

function rowMonth(row, names) {
  const key = rowFieldKey(row, names);
  if (!key) return null;
  const iso = asDate(row[key], row.$text[key]);
  if (iso) return monthKeyFromIso(iso);

  const text = cleanText(row.$text[key] ?? row[key]);
  if (!text) return null;
  const isoLike = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (isoLike) return `${isoLike[1]}-${String(Number(isoLike[2])).padStart(2, "0")}`;
  const usLike = text.match(/^(\d{1,2})[-/](\d{4})$/);
  if (usLike) return `${usLike[2]}-${String(Number(usLike[1])).padStart(2, "0")}`;
  const parsed = new Date(`${text} 1`);
  return Number.isNaN(parsed.getTime()) ? null : toIsoDate(parsed).slice(0, 7);
}

function monthNumberFromHeader(value) {
  const normalized = normalizeHeader(value);
  if (!normalized) return null;
  const monthLookup = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12,
  };
  const direct = monthLookup[normalized];
  if (direct) return direct;
  const firstToken = normalized.split("_")[0];
  return monthLookup[firstToken] || null;
}

function extractActiveDreamTeam(range, latestDate) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  if (!values.length) return [];

  const year = Number(latestDate?.slice(0, 4)) || new Date().getFullYear();
  let headerRow = -1;
  let monthColumns = [];

  for (let row = 0; row < values.length; row += 1) {
    const candidates = [];
    for (let col = 1; col < (values[row]?.length || 0); col += 1) {
      const monthNumber = monthNumberFromHeader(text[row]?.[col] || values[row]?.[col]);
      if (!monthNumber) continue;
      candidates.push({
        col,
        month: `${year}-${String(monthNumber).padStart(2, "0")}`,
      });
    }
    if (candidates.length >= 3) {
      headerRow = row;
      monthColumns = candidates;
      break;
    }
  }

  if (headerRow < 0) return [];

  const rows = [];
  for (let row = headerRow + 1; row < values.length; row += 1) {
    const campus = cleanText(text[row]?.[0] || values[row]?.[0]);
    if (!campus) continue;
    if (campus.toLowerCase().startsWith("total")) break;
    for (const { col, month } of monthColumns) {
      const activeDreamTeam = asNumber(values[row]?.[col], text[row]?.[col]);
      if (activeDreamTeam === null) continue;
      rows.push({
        month,
        campus,
        activeDreamTeam: normalizeValue(activeDreamTeam),
      });
    }
  }

  return rows;
}

function extractLongDreamTeamDetail(range) {
  return tableRows(range, ["campus"])
    .map((row) => ({
      date: rowDate(row, ["date", "sunday", "week", "service_date", "served_date"]),
      campus: rowText(row, ["campus"]),
      team: rowText(row, ["team", "serving_team", "dream_team", "role", "area", "ministry"]),
      served: rowNumber(row, ["served", "serving", "count", "team_count", "attendance", "volunteers"]),
      target: rowNumber(row, ["target", "goal", "needed", "health_target"]),
      notes: rowText(row, ["notes"]),
    }))
    .filter((row) => row.date && row.campus && row.team && row.served !== null);
}

function extractWideDreamTeamDetail(range) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  if (!values.length) return [];

  let headerRow = -1;
  let campusCol = -1;
  let teamCol = -1;
  let dateColumns = [];

  for (let row = 0; row < values.length; row += 1) {
    const keys = (values[row] || []).map((_, col) => normalizeHeader(text[row]?.[col] || values[row]?.[col]));
    campusCol = keys.findIndex((key) => key === "campus");
    teamCol = keys.findIndex((key) => ["team", "serving_team", "dream_team", "role", "area", "ministry"].includes(key));
    dateColumns = [];
    for (let col = 0; col < (values[row]?.length || 0); col += 1) {
      const date = asDate(values[row]?.[col], text[row]?.[col]);
      if (date) dateColumns.push({ col, date });
    }
    if (campusCol >= 0 && teamCol >= 0 && dateColumns.length) {
      headerRow = row;
      break;
    }
  }

  if (headerRow < 0) return [];

  const rows = [];
  for (let row = headerRow + 1; row < values.length; row += 1) {
    const campus = cleanText(text[row]?.[campusCol] || values[row]?.[campusCol]);
    const team = cleanText(text[row]?.[teamCol] || values[row]?.[teamCol]);
    if (!campus || !team) continue;
    for (const { col, date } of dateColumns) {
      const served = asNumber(values[row]?.[col], text[row]?.[col]);
      if (served === null) continue;
      rows.push({
        date,
        campus,
        team,
        served: normalizeValue(served),
        target: null,
        notes: null,
      });
    }
  }
  return rows;
}

function extractDreamTeamDetail(range) {
  const rows = extractLongDreamTeamDetail(range);
  const wideRows = rows.length ? [] : extractWideDreamTeamDetail(range);
  return [...rows, ...wideRows].sort(
    (a, b) => a.date.localeCompare(b.date) || a.campus.localeCompare(b.campus) || a.team.localeCompare(b.team),
  );
}

function extractCampusGrowthHistory(range) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  if (!values.length) return { years: [], rows: [], notes: [] };

  const header = values[0] || [];
  const yearCols = [];
  const metricCols = {};
  for (let col = 1; col < header.length; col += 1) {
    const year = asNumber(values[0]?.[col], text[0]?.[col]);
    const key = normalizeHeader(text[0]?.[col] || values[0]?.[col]);
    if (year && year >= 1900 && year <= 2100) {
      yearCols.push({ col, year: Math.round(year) });
    } else if (key) {
      metricCols[key] = col;
    }
  }

  const rows = [];
  const notes = [];
  for (let row = 1; row < values.length; row += 1) {
    const campus = cleanText(text[row]?.[0] || values[row]?.[0]);
    if (!campus) continue;
    if (campus.toLowerCase().startsWith("notes")) {
      for (const { col, year } of yearCols) {
        const note = cleanText(text[row]?.[col] || values[row]?.[col]);
        if (note) notes.push({ year, note });
      }
      continue;
    }

    const yearly = {};
    for (const { col, year } of yearCols) {
      const value = asNumber(values[row]?.[col], text[row]?.[col]);
      if (value !== null) yearly[year] = normalizeValue(value);
    }

    const keyFrom = (patterns) =>
      Object.keys(metricCols).find((candidate) => patterns.some((pattern) => candidate.includes(pattern)));
    const valueFrom = (patterns) => {
      const key = keyFrom(patterns);
      return key ? asNumber(values[row]?.[metricCols[key]], text[row]?.[metricCols[key]]) : null;
    };
    const textFrom = (patterns) => {
      const key = keyFrom(patterns);
      return key ? cleanText(text[row]?.[metricCols[key]] || values[row]?.[metricCols[key]]) : null;
    };

    rows.push({
      campus,
      yearly,
      yoyPct: percentValue(valueFrom(["yoy"])),
      growthSinceLaunch: textFrom(["growth_since_launch", "growth_since", "since_launch"]),
      cagrPct: percentValue(valueFrom(["cagr", "compound"])),
      netGrowth: valueFrom(["net_growth"]),
      marketSharePct: percentValue(valueFrom(["market_share"])),
    });
  }

  return {
    years: yearCols.map((item) => item.year),
    rows,
    notes,
  };
}

function extractHealthTargets(range) {
  return tableRows(range, ["metric_key", "metric_label"])
    .map((row) => ({
      key: rowText(row, ["metric_key", "key"]),
      label: rowText(row, ["metric_label", "label", "metric"]),
      targetType: rowText(row, ["target_type", "type"]),
      optimalMin: rowNumber(row, ["optimal_min", "min"]),
      optimalMax: rowNumber(row, ["optimal_max", "max"]),
      unit: rowText(row, ["unit", "format"]),
      direction: rowText(row, ["direction"]),
      notes: rowText(row, ["notes"]),
    }))
    .filter((row) => row.key || row.label);
}

function extractGroupConfig(range) {
  return tableRows(range, ["campus", "group_goal"])
    .map((row) => ({
      semester: rowText(row, ["semester"]),
      year: rowNumber(row, ["year"]),
      startDate: rowDate(row, ["start_date", "semester_start", "start"]),
      endDate: rowDate(row, ["end_date", "semester_end", "end"]),
      campus: rowText(row, ["campus"]),
      groupGoal: rowNumber(row, ["group_goal", "groups_goal", "goal"]),
      activeGroups: rowNumber(row, ["active_groups", "total_groups", "groups"]),
      groupSignups: rowNumber(row, ["group_signups", "signups", "people_signed_up"]),
      totalGroupMembers: rowNumber(row, ["total_group_members", "group_members", "members"]),
      notes: rowText(row, ["notes"]),
    }))
    .filter((row) => row.campus);
}

function extractLongGroupAttendance(range) {
  return tableRows(range, ["campus"])
    .map((row) => ({
      weekStart: rowDate(row, ["week_start", "week", "date", "sunday", "sunday_date"]),
      campus: rowText(row, ["campus"]),
      groupAttendance: rowNumber(row, [
        "group_attendance",
        "weekly_attendance",
        "attendance",
        "attended_count",
        "attended",
      ]),
      groupSignups: rowNumber(row, ["group_signups", "signups", "people_signed_up"]),
      totalGroupMembers: rowNumber(row, ["total_group_members", "group_members", "members"]),
      activeGroups: rowNumber(row, ["active_groups", "total_groups", "groups"]),
      notes: rowText(row, ["notes"]),
    }))
    .filter((row) => row.weekStart && row.campus && row.groupAttendance !== null);
}

function extractWideGroupAttendance(range) {
  const extracted = extractMetricSheet(range);
  return Object.entries(extracted.series).flatMap(([campus, points]) =>
    points.map((point) => ({
      weekStart: point.date,
      campus,
      groupAttendance: point.value,
      event: point.event,
    })),
  );
}

function extractGroupAttendance(longRange, wideRange) {
  const rows = [...extractLongGroupAttendance(longRange)];
  const wideRows = [
    ...extractWideGroupAttendance(wideRange),
    ...(rows.length ? [] : extractWideGroupAttendance(longRange)),
  ];
  const seen = new Set(rows.map((row) => `${row.weekStart}::${row.campus}`));
  for (const row of wideRows) {
    const key = `${row.weekStart}::${row.campus}`;
    if (!seen.has(key)) {
      seen.add(key);
      rows.push(row);
    }
  }
  return rows.sort((a, b) => a.weekStart.localeCompare(b.weekStart) || a.campus.localeCompare(b.campus));
}

function extractHeartSoulRows(range) {
  return tableRows(range, ["event_date", "campus", "role"])
    .map((row) => {
      const eventDate = rowDate(row, ["event_date", "date"]);
      return {
        month: rowMonth(row, ["month"]) || monthKeyFromIso(eventDate),
        eventDate,
        campus: rowText(row, ["campus"]),
        role: rowText(row, ["role"]),
        eligibleCount: rowNumber(row, ["eligible_count", "eligible"]),
        invitedCount: rowNumber(row, ["invited_count", "invited"]),
        attendedCount: rowNumber(row, ["attended_count", "attended", "attendance"]),
        included: rowBool(row, ["role_included_this_month", "included", "invited_this_month"]),
        source: rowText(row, ["source"]),
        notes: rowText(row, ["notes"]),
      };
    })
    .filter((row) => row.month && row.campus && row.role);
}

function extractLeadershipRows(range) {
  return tableRows(range, ["month", "campus", "role"])
    .map((row) => {
      const targetPositions = rowNumber(row, ["target_positions", "target", "needed_positions"]);
      const filledPositions = rowNumber(row, ["filled_positions", "filled", "current_positions"]);
      const enteredVacancies = rowNumber(row, ["vacancies", "vacancy_count"]);
      const vacancies =
        enteredVacancies ?? (targetPositions !== null && filledPositions !== null
          ? Math.max(0, targetPositions - filledPositions)
          : null);
      return {
        month: rowMonth(row, ["month"]),
        dataDueDate: rowDate(row, ["data_due_date", "due_date"]),
        campus: rowText(row, ["campus"]),
        role: rowText(row, ["role"]),
        targetPositions,
        filledPositions,
        vacancies,
        source: rowText(row, ["source"]),
        notes: rowText(row, ["notes"]),
      };
    })
    .filter((row) => row.month && row.campus && row.role);
}

function monthsBetween(startIso, endIso) {
  if (!startIso || !endIso) return [];
  const start = parseIsoDate(`${startIso.slice(0, 7)}-01`);
  const end = parseIsoDate(`${endIso.slice(0, 7)}-01`);
  const months = [];
  for (
    let cursor = new Date(start.getTime());
    cursor <= end;
    cursor.setMonth(cursor.getMonth() + 1)
  ) {
    const year = cursor.getFullYear();
    const month = String(cursor.getMonth() + 1).padStart(2, "0");
    months.push(`${year}-${month}`);
  }
  return months;
}

function extractHealthMonths(metrics, groupConfig, groupAttendance, heartSoulRows, leadershipRows, activeDreamTeam) {
  const months = new Set();
  for (const points of Object.values(metrics.attendance?.series || {})) {
    for (const point of points) {
      if (point.isSunday) months.add(monthKeyFromIso(point.date));
    }
  }
  for (const row of groupConfig) {
    for (const month of monthsBetween(row.startDate, row.endDate)) months.add(month);
  }
  for (const row of groupAttendance) months.add(monthKeyFromIso(row.weekStart));
  for (const row of heartSoulRows) months.add(row.month);
  for (const row of leadershipRows) months.add(row.month);
  for (const row of activeDreamTeam) months.add(row.month);
  return Array.from(months).filter(Boolean).sort();
}

function phaseSortKey(phase) {
  const parsed = Number(String(phase || "").split(" ")[0]);
  return Number.isFinite(parsed) ? parsed : 99;
}

function extractBigFive(rawRange, pivotRange) {
  const values = matrix(rawRange, "values");
  const text = matrix(rawRange, "text");
  const rows = [];
  const events = new Map();
  const phases = new Map();
  const years = new Set();

  for (let row = 1; row < values.length; row += 1) {
    const event = cleanText(text[row]?.[0] || values[row]?.[0]);
    const year = asNumber(values[row]?.[1], text[row]?.[1]);
    const date = asDate(values[row]?.[2], text[row]?.[2]);
    const campus = cleanText(text[row]?.[3] || values[row]?.[3]);
    const phase = cleanText(text[row]?.[4] || values[row]?.[4]);
    const attendance = asNumber(values[row]?.[5], text[row]?.[5]);

    if (!event || !year || !date || !campus || !phase || attendance === null) continue;

    events.set(event, displayLabel(event));
    phases.set(phase, displayLabel(phase));
    years.add(year);
    rows.push({
      event,
      eventLabel: displayLabel(event),
      year,
      date,
      campus,
      phase,
      phaseLabel: displayLabel(phase),
      attendance: normalizeValue(attendance),
    });
  }

  const grouped = new Map();
  for (const row of rows) {
    const key = `${row.event}::${row.year}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  }

  const eventYears = Array.from(grouped.values())
    .map((eventRows) => summarizeBigFiveRows(eventRows))
    .sort((a, b) => a.event.localeCompare(b.event) || a.year - b.year);

  const lookup = new Map(eventYears.map((row) => [`${row.event}::${row.year}`, row]));
  for (const row of eventYears) {
    const previous = lookup.get(`${row.event}::${row.year - 1}`);
    row.yoyCampaignPct = previous ? pct(pctChange(row.campaignTotal, previous.campaignTotal)) : null;
    row.yoyFeaturedPct =
      previous?.featuredTotal && row.featuredTotal
        ? pct(pctChange(row.featuredTotal, previous.featuredTotal))
        : null;
  }

  return {
    events: Array.from(events.entries()).map(([event, label]) => ({ event, label })),
    years: Array.from(years).sort(),
    phases: Array.from(phases.entries())
      .sort(([a], [b]) => phaseSortKey(a) - phaseSortKey(b))
      .map(([phase, label]) => ({ phase, label })),
    rows,
    eventYears,
    pivot: extractBigFivePivot(pivotRange),
    insights: buildBigFiveInsights(eventYears),
  };
}

function summarizeBigFiveRows(eventRows) {
  const first = eventRows[0];
  const phaseMap = new Map();
  const campusMap = new Map();
  for (const row of eventRows) {
    phaseMap.set(row.phase, {
      phase: row.phase,
      phaseLabel: row.phaseLabel,
      total: (phaseMap.get(row.phase)?.total || 0) + row.attendance,
    });
    campusMap.set(row.campus, {
      campus: row.campus,
      total: (campusMap.get(row.campus)?.total || 0) + row.attendance,
    });
  }

  const phaseTotals = Array.from(phaseMap.values()).sort(
    (a, b) => phaseSortKey(a.phase) - phaseSortKey(b.phase),
  );
  const campusTotals = Array.from(campusMap.values()).sort((a, b) => b.total - a.total);
  const eventPhase = phaseTotals.find((phase) => phase.phase === "07 Total" && phase.total > 0);
  const featured =
    eventPhase ||
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
  const postFourWeekAvg = average(postWeeks.map((phase) => phase.total));
  const featuredTotal = featured?.total ?? null;

  return {
    event: first.event,
    eventLabel: first.eventLabel,
    year: first.year,
    startDate: eventRows.map((row) => row.date).sort()[0],
    endDate: eventRows.map((row) => row.date).sort().at(-1),
    campaignTotal: Math.round(eventRows.reduce((sum, row) => sum + row.attendance, 0)),
    preTotal: pre?.total ? Math.round(pre.total) : null,
    featuredPhase: featured?.phase || null,
    featuredPhaseLabel: featured?.phaseLabel || null,
    featuredTotal: featuredTotal ? Math.round(featuredTotal) : null,
    postWeek1: postWeek1?.total ? Math.round(postWeek1.total) : null,
    postWeeks: postWeeks.map((phase) => ({ ...phase, total: Math.round(phase.total) })),
    postFourWeekAvg: postFourWeekAvg ? Math.round(postFourWeekAvg) : null,
    postFourWeekRetentionPct: featuredTotal && postFourWeekAvg ? pct(postFourWeekAvg / featuredTotal) : null,
    postFourWeekGrowthPct: postFourWeekAvg && pre?.total ? pct(pctChange(postFourWeekAvg, pre.total)) : null,
    liftPct: featuredTotal && pre?.total ? pct(pctChange(featuredTotal, pre.total)) : null,
    postRetentionPct: featuredTotal && postWeek1?.total ? pct(postWeek1.total / featuredTotal) : null,
    phaseTotals: phaseTotals.map((phase) => ({ ...phase, total: Math.round(phase.total) })),
    campusTotals: campusTotals.map((campus) => ({ ...campus, total: Math.round(campus.total) })),
  };
}

function extractBigFivePivot(range) {
  const values = matrix(range, "values");
  const text = matrix(range, "text");
  const yearCols = [];
  for (let col = 2; col < (values[3]?.length || 0); col += 1) {
    const year = asNumber(values[3]?.[col], text[3]?.[col]);
    if (year) yearCols.push({ col, year });
  }

  const output = [];
  let currentEvent = null;
  for (let row = 4; row < values.length; row += 1) {
    const label = cleanText(text[row]?.[0] || values[row]?.[0]);
    const phase = cleanText(text[row]?.[1] || values[row]?.[1]);
    if (label && !label.endsWith("Total") && label !== "Grand Total") currentEvent = label;
    if (label?.endsWith("Total") && label !== "Grand Total") {
      const event = label.replace(/ Total$/, "");
      for (const { col, year } of yearCols) {
        const value = asNumber(values[row]?.[col], text[row]?.[col]);
        if (value !== null) output.push({ event, year, campaignTotal: Math.round(value) });
      }
    } else if (phase && currentEvent) {
      for (const { col, year } of yearCols) {
        const value = asNumber(values[row]?.[col], text[row]?.[col]);
        if (value !== null) {
          output.push({ event: currentEvent, year, phase, phaseTotal: Math.round(value) });
        }
      }
    }
  }
  return output;
}

function buildInsights(campusStats, totals, yoy) {
  const insights = [];
  const latestTotal = totals.at(-1);
  const previousTotal = totals.at(-2);
  if (latestTotal && previousTotal) {
    const change = pct(pctChange(latestTotal.value, previousTotal.value));
    insights.push({
      type: "trend",
      title: "Latest Sunday is steady overall",
      body: `Total in-person attendance was ${latestTotal.value.toLocaleString()} on ${latestTotal.date}, ${change > 0 ? "+" : ""}${change}% compared with the Sunday before.`,
      severity: Math.abs(change || 0) < 10 ? "info" : "warning",
    });
  }

  const volatile = campusStats
    .slice()
    .sort((a, b) => b.swing10Count - a.swing10Count || b.avgAbsSwingPct - a.avgAbsSwingPct)[0];
  if (volatile) {
    insights.push({
      type: "volatility",
      title: `${volatile.campus} has the most big Sunday changes`,
      body: `${volatile.campus} has ${volatile.swing10Count} Sundays where attendance moved 10%+ from the Sunday before (${volatile.positiveSwing10Count} up, ${volatile.negativeSwing10Count} down).`,
      severity: "warning",
    });
  }

  const low = campusStats
    .filter((row) => typeof row.vsPrevious4Pct === "number")
    .sort((a, b) => a.vsPrevious4Pct - b.vsPrevious4Pct)[0];
  if (low && low.vsPrevious4Pct < -5) {
    insights.push({
      type: "watch",
      title: `${low.campus} is below recent Sundays`,
      body: `${low.campus} is ${low.vsPrevious4Pct}% below its previous four regular Sundays (${low.previous4Avg?.toLocaleString()}).`,
      severity: "warning",
    });
  }

  const biggestSwing = campusStats
    .flatMap((row) => (row.maxSwing ? [{ ...row.maxSwing, campus: row.campus }] : []))
    .sort((a, b) => b.absPct - a.absPct)[0];
  if (biggestSwing) {
    insights.push({
      type: "event",
      title: `${biggestSwing.campus} had the biggest Sunday-to-Sunday change`,
      body: `${biggestSwing.campus} moved ${biggestSwing.pct > 0 ? "+" : ""}${biggestSwing.pct}% from ${biggestSwing.fromDate} to ${biggestSwing.toDate}. Special Sundays should be read separately from regular attendance patterns.`,
      severity: "info",
    });
  }

  if (yoy?.total && latestTotal) {
    const change = pct(pctChange(latestTotal.value, yoy.total));
    insights.push({
      type: "yoy",
      title: "Compared with last year",
      body: `${latestTotal.date} is ${change > 0 ? "+" : ""}${change}% compared with ${yoy.matchedDate} across matching campuses (${latestTotal.value.toLocaleString()} vs ${yoy.total.toLocaleString()}).`,
      severity: change < -10 ? "critical" : "info",
    });
  }
  return insights;
}

function buildBigFiveInsights(eventYears) {
  const active = eventYears.filter((row) => row.campaignTotal > 0);
  if (!active.length) return [];
  const latestYear = Math.max(...active.map((row) => row.year));
  const latestRows = active.filter((row) => row.year === latestYear).sort((a, b) => b.campaignTotal - a.campaignTotal);
  const leader = latestRows[0];
  const insights = [
    {
      type: "leader",
      title: `${leader.eventLabel} had the largest ${latestYear} Big 5 turnout`,
      body: `${leader.eventLabel} has ${leader.campaignTotal.toLocaleString()} total attendance in ${latestYear}; its main Sunday was ${leader.featuredTotal?.toLocaleString()}.`,
      severity: "info",
    },
  ];

  const yoyRows = latestRows.filter((row) => typeof row.yoyCampaignPct === "number");
  if (yoyRows.length) {
    const target = yoyRows
      .slice()
      .sort((a, b) => Math.abs(b.yoyCampaignPct) - Math.abs(a.yoyCampaignPct))[0];
    insights.push({
      type: "yoy",
      title: `${target.eventLabel} is ${target.yoyCampaignPct > 0 ? "+" : ""}${target.yoyCampaignPct}% compared with last year`,
      body: `${target.eventLabel} total attendance is ${target.campaignTotal.toLocaleString()} in ${target.year} compared with the prior year.`,
      severity: target.yoyCampaignPct < -10 ? "critical" : "info",
    });
  }

  const lift = latestRows.filter((row) => typeof row.liftPct === "number").sort((a, b) => b.liftPct - a.liftPct)[0];
  if (lift) {
    insights.push({
      type: "lift",
      title: `${lift.eventLabel} had the strongest jump from the previous Sunday`,
      body: `${lift.featuredPhaseLabel} was ${lift.liftPct > 0 ? "+" : ""}${lift.liftPct}% above the previous Sunday (${lift.featuredTotal.toLocaleString()} vs ${lift.preTotal.toLocaleString()}).`,
      severity: "info",
    });
  }

  const retention = latestRows
    .filter((row) => typeof row.postRetentionPct === "number")
    .sort((a, b) => a.postRetentionPct - b.postRetentionPct)[0];
  if (retention) {
    insights.push({
      type: "retention",
      title: `${retention.eventLabel} following Sunday is ${retention.postRetentionPct}% of the main Sunday`,
      body: `The following Sunday had ${retention.postWeek1.toLocaleString()} compared with ${retention.featuredTotal.toLocaleString()} on the main Sunday.`,
      severity: retention.postRetentionPct < 50 ? "warning" : "info",
    });
  }

  return insights;
}

function readableError(error) {
  const message = error?.message || String(error);
  if (message.includes("AADSTS65001") || message.includes("consent")) {
    return "Microsoft needs permission approval for User.Read and Files.ReadWrite.";
  }
  if (message.includes("403")) {
    return "Microsoft Graph denied access. Confirm Files.ReadWrite permission and workbook access.";
  }
  if (message.includes("404")) {
    return "The SharePoint workbook link could not be resolved by Graph.";
  }
  return message;
}
