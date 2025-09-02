/**
 * Instantly Campaign Report v2: Using API v2 See https://developer.instantly.ai/api/v2;
 * Config sheet:
 *   B1 = start date (YYYY-MM-DD or Date)
 *   B2 = end date   (YYYY-MM-DD or Date)
 *   B3 = v2 API key (string)
 */
function generateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Instantly Campaigns Data");
  var config = ss.getSheetByName("Config");

  var startRaw = config.getRange("B1").getValue();
  var endRaw   = config.getRange("B2").getValue();
  var apiKey   = String(config.getRange("B3").getValue()).trim();
  if (!apiKey) throw new Error("Config!B3 (API key) is empty.");

  var tz = ss.getSpreadsheetTimeZone();
  var startDate = normalizeYMD_(startRaw, tz);
  var endDate   = normalizeYMD_(endRaw, tz);

  // 1) Fetch analytics (may be empty for new campaigns)
  var analytics = fetchAnalytics_(apiKey, startDate, endDate);

  // 2) Fetch all campaigns
  var campaigns = fetchAllCampaigns_(apiKey);

  // 3) Build a map campaign_id -> analytics object
  var aMap = {};
  analytics.forEach(function(a) {
    var id = a.campaign_id || a.id;
    if (id) aMap[id] = a;
  });

  // 4) Prepare sheet
  dashboard.clear();
  var headers = [
    "Sl. No",
    "Campaign ID",
    "Campaign Name",
    "Number of Contacts",     // leads_count
    "Emails Sent",            // emails_sent_count
    "Emails Read",            // open_count (total)
    "Contacts Opened Email",  // open_count_unique
    "Contacts Replied",       // reply_count
    "Completed",              // completed_count
    "Bounced Leads",          // bounced_count
    "Unsubscribed"            // unsubscribed_count
  ];
  dashboard.getRange(1, 1, 1, headers.length).setValues([headers]);
  dashboard.setFrozenRows(1);

  // 5) Left-join: every campaign gets a row; analytics filled if present, else zeros
  var rows = campaigns.map(function(c, i) {
    var id = c.id || c.campaign_id || "";
    var name = c.name || c.campaign_name || "";

    var a = aMap[id] || {};
    return [
      i + 1,
      id,
      name,
      num(a.leads_count),
      num(a.emails_sent_count),
      num(a.open_count),
      num(a.open_count_unique),
      num(a.reply_count),
      num(a.completed_count),
      num(a.bounced_count),
      num(a.unsubscribed_count)
    ];
  });

  if (rows.length) {
    dashboard.getRange(2, 1, rows.length, headers.length).setValues(rows);
    dashboard.autoResizeColumns(1, headers.length);
  } else {
    // Extremely rare: no campaigns in workspace
    SpreadsheetApp.getUi().alert(
      'No Campaigns',
      'No campaigns found in this workspace yet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/** Fetch analytics for date range (paged by page/per_page). */
function fetchAnalytics_(apiKey, startYMD, endYMD) {
  var base = 'https://api.instantly.ai/api/v2/campaigns/analytics';
  var perPage = 200, page = 1, all = [];

  while (true) {
    var url = base + '?start_date=' + encodeURIComponent(startYMD)
                   + '&end_date='   + encodeURIComponent(endYMD)
                   + '&per_page='   + perPage
                   + '&page='       + page;

    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code !== 200) throw new Error('Analytics fetch HTTP ' + code + ': ' + res.getContentText());

    var json = JSON.parse(res.getContentText());
    var batch = Array.isArray(json) ? json : (json && Array.isArray(json.data) ? json.data : []);
    if (!batch.length) break;

    all = all.concat(batch);
    if (batch.length < perPage) break;
    page += 1;
  }
  return all;
}

/** Fetch ALL campaigns via cursor pagination: next_starting_after. */
function fetchAllCampaigns_(apiKey) {
  var base = 'https://api.instantly.ai/api/v2/campaigns';
  var all = [];
  var startingAfter = null;

  while (true) {
    var url = base + (startingAfter ? ('?starting_after=' + encodeURIComponent(startingAfter)) : '');
    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code !== 200) throw new Error('Campaigns fetch HTTP ' + code + ': ' + res.getContentText());

    var json = JSON.parse(res.getContentText());
    var items = (json && Array.isArray(json.items)) ? json.items : [];
    all = all.concat(items);

    if (json && json.next_starting_after) {
      startingAfter = json.next_starting_after;
    } else {
      break;
    }
  }

  return all;
}

/** Helpers */
function normalizeYMD_(val, tz) {
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val.trim())) return val.trim();
  var d = new Date(val);
  if (isNaN(d.getTime())) throw new Error('Invalid date in Config');
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}
function num(v){ v = Number(v); return isFinite(v) ? v : 0; }

