export const dashboardData = {
  source: {
    workbook: "Sign in required",
    generatedFrom: "",
    latestAttendanceDate: null,
  },
  campuses: [],
  metrics: {
    attendance: { label: "Attendance", series: {}, totalSeries: [] },
    kids: { label: "Kids", series: {}, totalSeries: [] },
    growthTrack: { label: "Growth Track", series: {}, totalSeries: [] },
    baptism: { label: "Baptisms", series: {}, totalSeries: [] },
    salvations: { label: "Salvations", series: {}, totalSeries: [] },
    firstTimers: { label: "First Timers", series: {}, totalSeries: [] },
    dreamTeam: { label: "Dream Team", series: {}, totalSeries: [] },
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
  dreamTeamDetail: [],
  campusGrowthHistory: {
    years: [],
    rows: [],
    notes: [],
  },
  health: {
    targets: [],
    groupConfig: [],
    groupAttendance: [],
    heartSoulRows: [],
    leadershipRows: [],
    activeDreamTeam: [],
    directorRoster: [],
    months: [],
  },
  insights: [],
};

export default dashboardData;
