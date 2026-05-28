// ============================================================
// QualityService.gs — Aggregations and logic for Quality Audits
// Source: 'PLX Raw data' tab
// ============================================================

var QUALITY_SHEET_NAME = 'PLX Raw data';

// ── SCHEMA MAPPING ────────────────────────────────────────────────────────

var Q_COLS = {
  CASE_ID: 0,            // A
  ENTITY_GROUP: 1,       // B
  AGENT_LDAP: 2,         // C
  OPENING_CHANNEL: 3,    // D
  REVIEW_DATE: 4,        // E
  REVIEW_WEEK: 5,        // F
  REVIEW_MONTH: 6,       // G
  CASE_DATE: 7,          // H
  CASE_WEEK: 8,          // I
  CASE_MONTH: 9,         // J
  CUSTOMER_CRITICAL: 10, // K
  BUSINESS_CRITICAL: 11, // L
  COMPLIANCE_CRITICAL: 12,// M
  REVIEWER_COMMENTS: 13, // N

  // Customer Critical Parameters (O-W)
  LISTENING: 14,
  PROBING: 15,
  COMPLETE_RESOLUTION: 16,
  TROUBLESHOOTING: 17,
  USER_EXPECTATIONS: 18,
  EMPATHY: 19,
  OWNERSHIP: 20,
  REFUNDS: 21,
  RESPONSIVENESS: 22,

  // Business Critical Parameters (X-AD)
  CONSULTS_ESCALATIONS: 23,
  CASE_DETAILS: 24,
  CATEGORIZATION: 25,
  CSAT_REMINDER: 26,
  CASE_STATE: 27,
  OPENING_CLOSING: 28,
  LANGUAGE_PROFICIENCY: 29,

  // Compliance Critical Parameters (AE-AH)
  AUTHENTICATION: 30,
  GOOGLE_ONLY_INFO: 31,
  PROFESSIONAL_CONDUCT: 32,
  PAYMENT_COMPLAINTS: 33,

  TEAM: 36,              // AK
  AGENT_WORKFLOW: 37,    // AL

  SUPERVISOR: 66,        // BO
  MANAGER: 67            // BP
};

var Q_TARGETS = {
  CUSTOMER: 95,
  BUSINESS: 90,
  COMPLIANCE: 99.50
};

// ── APPS SCRIPT WEB APP ───────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('QualityView')
    .evaluate()
    .setTitle('Quality Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── DATA LOADING ──────────────────────────────────────────────────────────

function getRawQualityData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  if (!sheet) return [];

  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return [];

  return raw.slice(1);
}

function getAvailableQualityMonths() {
  var rows = getRawQualityData();
  var seen = {};
  rows.forEach(function(r) {
    var month = r[Q_COLS.REVIEW_MONTH];
    if (month) {
      if (month instanceof Date) {
        month = Utilities.formatDate(month, Session.getScriptTimeZone(), 'yyyy-MM');
      }
      seen[month] = true;
    }
  });
  return Object.keys(seen).sort().reverse();
}

function normalizeQualityMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(val).trim();
}

function normalizeLdap(val) {
  if (!val) return '';
  return String(val).trim().toLowerCase().split('@')[0];
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

// ── AGGREGATION ───────────────────────────────────────────────────────────

function aggregateQualityRows(rows) {
  if (rows.length === 0) return null;

  var customerSum = 0, businessSum = 0, complianceSum = 0;

  var params = {};
  var paramGroups = {
    customer: ['LISTENING', 'PROBING', 'COMPLETE_RESOLUTION', 'TROUBLESHOOTING', 'USER_EXPECTATIONS', 'EMPATHY', 'OWNERSHIP', 'REFUNDS', 'RESPONSIVENESS'],
    business: ['CONSULTS_ESCALATIONS', 'CASE_DETAILS', 'CATEGORIZATION', 'CSAT_REMINDER', 'CASE_STATE', 'OPENING_CLOSING', 'LANGUAGE_PROFICIENCY'],
    compliance: ['AUTHENTICATION', 'GOOGLE_ONLY_INFO', 'PROFESSIONAL_CONDUCT', 'PAYMENT_COMPLAINTS']
  };

  var paramCols = [].concat(paramGroups.customer, paramGroups.business, paramGroups.compliance);
  paramCols.forEach(function(p) { params[p] = { yes: 0, total: 0 }; });

  rows.forEach(function(r) {
    customerSum += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
    businessSum += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
    complianceSum += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);

    paramCols.forEach(function(p) {
      var val = String(r[Q_COLS[p]]).trim().toLowerCase();
      if (val === 'yes' || val === 'no' || val === '1' || val === '0') {
        params[p].total++;
        if (val === 'yes' || val === '1') params[p].yes++;
      }
    });
  });

  var count = rows.length;
  var paramScores = {};
  paramCols.forEach(function(p) {
    paramScores[p] = params[p].total > 0 ? (params[p].yes / params[p].total) * 100 : null;
  });

  var groupedParams = {
    customer: paramGroups.customer.map(p => ({ name: p, score: paramScores[p] })),
    business: paramGroups.business.map(p => ({ name: p, score: paramScores[p] })),
    compliance: paramGroups.compliance.map(p => ({ name: p, score: paramScores[p] }))
  };

  return {
    customer: (customerSum / count) * 100,
    business: (businessSum / count) * 100,
    compliance: (complianceSum / count) * 100,
    count: count,
    params: paramScores,
    groupedParams: groupedParams,
    targets: Q_TARGETS
  };
}

function aggregateTrends(rows) {
  var daily = {};
  var weekly = {};

  rows.forEach(function(r) {
    var dateRaw = r[Q_COLS.REVIEW_DATE];
    var date = formatDate(dateRaw);
    var week = r[Q_COLS.REVIEW_WEEK];

    [ {obj: daily, key: date}, {obj: weekly, key: week} ].forEach(function(t) {
      if (!t.key) return;
      if (!t.obj[t.key]) t.obj[t.key] = { customer: 0, business: 0, compliance: 0, count: 0 };
      t.obj[t.key].customer += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
      t.obj[t.key].business += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
      t.obj[t.key].compliance += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);
      t.obj[t.key].count++;
    });
  });

  var formatTrend = function(obj) {
    return Object.keys(obj).sort().map(function(k) {
      var d = obj[k];
      return {
        label: k,
        customer: (d.customer / d.count) * 100,
        business: (d.business / d.count) * 100,
        compliance: (d.compliance / d.count) * 100,
        avg: ((d.customer + d.business + d.compliance) / (d.count * 3)) * 100
      };
    });
  };

  return { daily: formatTrend(daily), weekly: formatTrend(weekly) };
}

// ── CLIENT WRAPPERS ───────────────────────────────────────────────────────

function clientGetAvailableQualityMonths() {
  return getAvailableQualityMonths();
}

function clientGetMyQuality(ldap, month) {
  if (!ldap) ldap = Session.getActiveUser().getEmail().split('@')[0];

  var allRows = getRawQualityData();
  var filtered = allRows.filter(function(r) {
    return normalizeLdap(r[Q_COLS.AGENT_LDAP]) === normalizeLdap(ldap) &&
           normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var caseLog = filtered.map(function(r) {
    return {
      caseId: r[Q_COLS.CASE_ID],
      reviewDate: formatDate(r[Q_COLS.REVIEW_DATE]),
      customer: r[Q_COLS.CUSTOMER_CRITICAL],
      business: r[Q_COLS.BUSINESS_CRITICAL],
      compliance: r[Q_COLS.COMPLIANCE_CRITICAL],
      comments: r[Q_COLS.REVIEWER_COMMENTS],
      details: {
        customer: {
          LISTENING: r[Q_COLS.LISTENING],
          PROBING: r[Q_COLS.PROBING],
          COMPLETE_RESOLUTION: r[Q_COLS.COMPLETE_RESOLUTION],
          TROUBLESHOOTING: r[Q_COLS.TROUBLESHOOTING],
          USER_EXPECTATIONS: r[Q_COLS.USER_EXPECTATIONS],
          EMPATHY: r[Q_COLS.EMPATHY],
          OWNERSHIP: r[Q_COLS.OWNERSHIP],
          REFUNDS: r[Q_COLS.REFUNDS],
          RESPONSIVENESS: r[Q_COLS.RESPONSIVENESS]
        },
        business: {
          CONSULTS_ESCALATIONS: r[Q_COLS.CONSULTS_ESCALATIONS],
          CASE_DETAILS: r[Q_COLS.CASE_DETAILS],
          CATEGORIZATION: r[Q_COLS.CATEGORIZATION],
          CSAT_REMINDER: r[Q_COLS.CSAT_REMINDER],
          CASE_STATE: r[Q_COLS.CASE_STATE],
          OPENING_CLOSING: r[Q_COLS.OPENING_CLOSING],
          LANGUAGE_PROFICIENCY: r[Q_COLS.LANGUAGE_PROFICIENCY]
        },
        compliance: {
          AUTHENTICATION: r[Q_COLS.AUTHENTICATION],
          GOOGLE_ONLY_INFO: r[Q_COLS.GOOGLE_ONLY_INFO],
          PROFESSIONAL_CONDUCT: r[Q_COLS.PROFESSIONAL_CONDUCT],
          PAYMENT_COMPLAINTS: r[Q_COLS.PAYMENT_COMPLAINTS]
        }
      }
    };
  });

  return {
    ldap: ldap,
    month: month,
    stats: stats,
    trends: trends,
    caseLog: caseLog,
    hasData: filtered.length > 0
  };
}

function clientGetTeamQuality(supervisor, month) {
  var allRows = getRawQualityData();
  var filtered = allRows.filter(function(r) {
    return String(r[Q_COLS.SUPERVISOR]).trim() === String(supervisor).trim() &&
           normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var agentStats = {};
  var uniqueLdaps = [];
  filtered.forEach(function(r) {
    var ldap = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
    if (!agentStats[ldap]) {
      agentStats[ldap] = [];
      uniqueLdaps.push(ldap);
    }
    agentStats[ldap].push(r);
  });

  var agents = uniqueLdaps.map(function(ldap) {
    return {
      ldap: ldap,
      stats: aggregateQualityRows(agentStats[ldap])
    };
  }).sort((a,b) => (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance));

  return {
    supervisor: supervisor,
    month: month,
    stats: stats,
    trends: trends,
    agents: agents,
    hasData: filtered.length > 0
  };
}

function clientGetClusterQuality(manager, month) {
  var allRows = getRawQualityData();
  var filtered = allRows.filter(function(r) {
    return String(r[Q_COLS.MANAGER]).trim() === String(manager).trim() &&
           normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var supervisorStats = {};
  var uniqueSupervisors = [];
  filtered.forEach(function(r) {
    var sup = String(r[Q_COLS.SUPERVISOR]).trim();
    if (!supervisorStats[sup]) {
      supervisorStats[sup] = [];
      uniqueSupervisors.push(sup);
    }
    supervisorStats[sup].push(r);
  });

  var supervisors = uniqueSupervisors.map(function(sup) {
    return {
      name: sup,
      stats: aggregateQualityRows(supervisorStats[sup])
    };
  }).sort((a,b) => (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance));

  return {
    manager: manager,
    month: month,
    stats: stats,
    trends: trends,
    supervisors: supervisors,
    hasData: filtered.length > 0
  };
}

function clientGetAllAgents() {
  var rows = getRawQualityData();
  var seen = {};
  rows.forEach(function(r) {
    var ldap = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
    if (ldap) seen[ldap] = true;
  });
  return Object.keys(seen).sort().map(function(ldap) {
    return { ldap: ldap };
  });
}

function clientGetAllSupervisors() {
  var rows = getRawQualityData();
  var seen = {};
  rows.forEach(function(r) {
    var sup = String(r[Q_COLS.SUPERVISOR]).trim();
    if (sup) seen[sup] = true;
  });
  return Object.keys(seen).sort();
}

function clientGetAllManagers() {
  var rows = getRawQualityData();
  var seen = {};
  rows.forEach(function(r) {
    var mgr = String(r[Q_COLS.MANAGER]).trim();
    if (mgr) seen[mgr] = true;
  });
  return Object.keys(seen).sort();
}

function clientGetSession() {
  var email = Session.getActiveUser().getEmail();
  return {
    ldap: email.split('@')[0],
    email: email
  };
}
