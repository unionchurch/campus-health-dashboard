export const dashboardData = {
  source: {
    workbook: "Sign in required",
    generatedFrom: "",
    latestAttendanceDate: null,
  },
  campuses: [],
  metrics: {
    attendance: { label: "Attendance", series: {} },
    kids: { label: "Kids", series: {} },
    growthTrack: { label: "Growth Track", series: {} },
    baptism: { label: "Baptisms", series: {} },
    salvations: { label: "Salvations", series: {} },
    firstTimers: { label: "First Timers", series: {} },
    dreamTeam: { label: "Dream Team", series: {} },
  },
  history: {
    attendance2025: {},
  },
  latest: {
    date: null,
    attendanceByCampus: {},
    metricSnapshots: [],
  },
  totals: [],
  campusStats: [],
  volatility: [],
  yoy: null,
  bigFive: {
    events: [],
    years: [],
    phases: [],
    rows: [],
    eventYears: [],
    pivot: [],
    insights: [],
  },
  health: {
    targets: [],
    groupConfig: [],
    groupAttendance: [],
    heartSoulRows: [],
    leadershipRows: [],
    months: [],
  },
  insights: [],
};

export default dashboardData;
