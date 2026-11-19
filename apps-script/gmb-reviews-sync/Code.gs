/**
 * CARS24 GMB Sell dashboard sync
 *
 * Reads the review-by-region/day summary and filtered SI_DATA rows, then
 * publishes the dashboard snapshot to a GitHub repository. Negative review
 * rows are reduced to date/region counts; review text, reviewer identity,
 * business location, and other free-text fields are never uploaded.
 *
 * Configure Script Properties before running installSyncTriggers().
 */
const SYNC_CONFIG = {
  TIME_ZONE: 'Asia/Kolkata',
  REVIEW_TAB_RE: /^[A-Za-z]{3}['’]\d{2}$/,
  STOCK_TAB: 'SI_DATA',
  NEGATIVE_TAB: 'Negative Reviews',
  CONDITION: 'Non Scrap & Perfect Working Condition Car',
  HANDLER: 'syncFromSheets',
  MAIL_HANDLER: 'pollLatestMailReport',
  MAIL_QUERY: 'filename:xlsx newer_than:120d',
  MAIL_DATA_PATH: 'mail-mtd-data.json'
};

function installSyncTriggers() {
  const cfg = getConfig_();
  const sourceIds = [cfg.REVIEWS_SHEET_ID, cfg.STOCK_SHEET_ID];
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(trigger) {
    if ([SYNC_CONFIG.HANDLER, SYNC_CONFIG.MAIL_HANDLER].indexOf(trigger.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  sourceIds.forEach(function(id) {
    ScriptApp.newTrigger(SYNC_CONFIG.HANDLER).forSpreadsheet(id).onEdit().create();
    ScriptApp.newTrigger(SYNC_CONFIG.HANDLER).forSpreadsheet(id).onChange().create();
    ScriptApp.newTrigger(SYNC_CONFIG.HANDLER).forSpreadsheet(id).onFormSubmit().create();
  });
  ScriptApp.newTrigger(SYNC_CONFIG.HANDLER).timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger(SYNC_CONFIG.MAIL_HANDLER).timeBased().everyMinutes(5).create();

  return syncFromSheets();
}

function removeSyncTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if ([SYNC_CONFIG.HANDLER, SYNC_CONFIG.MAIL_HANDLER].indexOf(trigger.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/** Poll Gmail every five minutes and publish only aggregate MTD counts. */
function pollLatestMailReport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { status: 'busy' };
  let tempId = '';
  try {
    const cfg = getConfig_();
    const props = PropertiesService.getScriptProperties();
    const list = Gmail.Users.Messages.list('me', { q: cfg.MAIL_QUERY || SYNC_CONFIG.MAIL_QUERY, maxResults: 30 });
    const messages = list.messages || [];
    if (!messages.length) return { status: 'no-matching-mail' };
    const lastDate = props.getProperty('LAST_EMAIL_REPORT_DATE') || '';
    let selected = null, attachment = null;
    for (let i = 0; i < messages.length; i++) {
      const msg = Gmail.Users.Messages.get('me', messages[i].id, { format: 'full' });
      const found = findXlsxAttachment_(msg.payload);
      if (found && reportDateFromFilename_(found.filename) >= lastDate) {
        const received = Number(msg.internalDate) || 0;
        if (!selected || received > selected.internalDate) {
          selected = { id: messages[i].id, internalDate: received };
          attachment = found;
        }
      }
    }
    if (!selected || !attachment) return { status: 'no-new-report' };
    if (selected.id === props.getProperty('LAST_EMAIL_MESSAGE_ID')) return { status: 'unchanged' };
    const part = Gmail.Users.Messages.Attachments.get('me', selected.id, attachment.attachmentId);
    const bytes = Utilities.base64DecodeWebSafe(part.data);
    const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment.filename);
    const imported = Drive.Files.create({ name: 'Temporary CARS24 MTD report import', mimeType: 'application/vnd.google-apps.spreadsheet' }, blob, { fields: 'id' });
    tempId = imported.id;
    const report = SpreadsheetApp.openById(tempId);
    const snapshot = buildMailSnapshot_(report, attachment.filename, selected.internalDate);
    const json = JSON.stringify(snapshot);
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json, Utilities.Charset.UTF_8));
    if (hash !== props.getProperty('LAST_EMAIL_HASH')) {
      publishJsonToGitHub_(cfg, SYNC_CONFIG.MAIL_DATA_PATH, json, 'Update partner review MTD through ' + snapshot.meta.asOf);
      props.setProperty('LAST_EMAIL_HASH', hash);
    }
    props.setProperty('LAST_EMAIL_MESSAGE_ID', selected.id);
    props.setProperty('LAST_EMAIL_REPORT_DATE', snapshot.meta.asOf);
    props.setProperty('LAST_EMAIL_SYNC', new Date().toISOString());
    return { status: 'published', asOf: snapshot.meta.asOf, total: snapshot.summary.total };
  } finally {
    if (tempId) {
      try { Drive.Files.update({ trashed: true }, tempId); } catch (cleanupError) { console.log('Temporary import cleanup failed: ' + cleanupError); }
    }
    lock.releaseLock();
  }
}

function findXlsxAttachment_(part) {
  if (!part) return null;
  const filename = String((part.filename || (part.body && part.body.filename)) || '');
  const body = part.body || {};
  if (/^CARS24_Reviews_MTD_.*\.xlsx$/i.test(filename) && body.attachmentId) {
    return { filename: filename, attachmentId: body.attachmentId };
  }
  const children = part.parts || [];
  for (let i = 0; i < children.length; i++) {
    const found = findXlsxAttachment_(children[i]);
    if (found) return found;
  }
  return null;
}

function reportDateFromFilename_(filename) {
  const m = String(filename).match(/(\d{1,2})[-_ ]([A-Za-z]{3})[-_ ](\d{4})/);
  if (!m) return '';
  const months = { jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12' };
  const mm = months[m[2].toLowerCase()];
  return mm ? m[3] + '-' + mm + '-' + ('0' + m[1]).slice(-2) : '';
}

function buildMailSnapshot_(book, filename, receivedAt) {
  const sheet = book.getSheetByName('Reviews Data');
  const summarySheet = book.getSheetByName('Summary');
  if (!sheet || !summarySheet) throw new Error('Report must contain Reviews Data and Summary sheets.');
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) throw new Error('Reviews Data has no rows.');
  const headers = rows[0].map(function(v) { return String(v || '').trim().toLowerCase(); });
  const col = function(name) { const i = headers.indexOf(name.toLowerCase()); if (i < 0) throw new Error('Missing column: ' + name); return i; };
  const dateCol = col('Review Date'), businessCol = col('Business Name'), locationCol = col('Location'), ratingCol = col('Review Rating');
  const daily = {}, regions = {}, totals = { buy:0, sell:0, other:0, buyNegative:0, sellNegative:0, otherNegative:0, total:0, negative:0 };
  let minDay = '', maxDay = '';
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r], date = flexibleDateKey_(row[dateCol]);
    if (!date) continue;
    const team = classifyBusiness_(row[businessCol]), region = regionFromReportLocation_(row[locationCol]);
    const rating = Number(row[ratingCol]), negative = rating > 0 && rating <= 2;
    totals[team]++; totals.total++; if (negative) { totals[team + 'Negative']++; totals.negative++; }
    if (!daily[date]) daily[date] = { date: date, buy:0, sell:0, other:0, buyNegative:0, sellNegative:0, otherNegative:0 };
    if (!regions[region]) regions[region] = { region:region, buy:0, sell:0, other:0, buyNegative:0, sellNegative:0, otherNegative:0 };
    daily[date][team]++; regions[region][team]++;
    if (negative) { daily[date][team + 'Negative']++; regions[region][team + 'Negative']++; }
    if (!minDay || date < minDay) minDay = date;
    if (date > maxDay) maxDay = date;
  }
  if (!totals.total) throw new Error('No dated review rows found.');
  const reported = reportSummaryCounts_(summarySheet);
  if (reported.total !== totals.total || reported.negative !== totals.negative) {
    throw new Error('Report reconciliation failed: rows ' + totals.total + '/' + totals.negative + ' negatives; summary ' + reported.total + '/' + reported.negative + '.');
  }
  totals.buyPositive = totals.buy - totals.buyNegative; totals.sellPositive = totals.sell - totals.sellNegative; totals.otherPositive = totals.other - totals.otherNegative; totals.positive = totals.total - totals.negative;
  const complete = function(item) {
    item.total = item.buy + item.sell + item.other;
    item.negative = item.buyNegative + item.sellNegative + item.otherNegative;
    item.buyPositive = item.buy - item.buyNegative;
    item.sellPositive = item.sell - item.sellNegative;
    item.otherPositive = item.other - item.otherNegative;
    item.positive = item.total - item.negative;
    return item;
  };
  const dayRows = Object.keys(daily).sort().map(function(k) { return complete(daily[k]); });
  const regionRows = Object.keys(regions).sort(function(a,b) { return regions[b].buy + regions[b].sell - regions[a].buy - regions[a].sell; }).map(function(k) { return complete(regions[k]); });
  return { meta:{ asOf:maxDay, periodStart:minDay, periodEnd:maxDay, source:'Email report ' + filename, receivedAt:new Date(receivedAt).toISOString(), dateRange:{ start:minDay, end:maxDay }, notes:['Buy/Sell team classification uses Business Name: names containing Buy are Buy; otherwise names containing Sell are Sell; remaining names are Other.','Negative means a 1–2 star rating. Positive is total minus negative (neutral included), per dashboard convention.','Other listings are retained for reconciliation and excluded from Buy/Sell comparisons.','Report may lag actual collection by 1–2 days.'] }, summary:totals, byRegion:regionRows, byDay:dayRows };
}

function reportSummaryCounts_(sheet) {
  const values = sheet.getDataRange().getValues(); let total = null, negative = null;
  values.forEach(function(row) {
    const label = String(row[0] || '').toLowerCase(), value = Number(row[1]);
    if (/^total reviews?$/.test(label)) total = value;
    if (label.indexOf('negative') >= 0 || /^1.?2.? star/.test(label)) negative = value;
  });
  if (total === null || negative === null) throw new Error('Cannot locate total and negative counts in Summary sheet for reconciliation.');
  return { total:total, negative:negative };
}

function classifyBusiness_(value) {
  const name = String(value || '').toLowerCase();
  if (name.indexOf('buy') >= 0) return 'buy';
  if (name.indexOf('sell') >= 0) return 'sell';
  return 'other';
}

function flexibleDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, SYNC_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  if (typeof value === 'number' && value >= 20000 && value <= 90000) return dateKey_(value);
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})[ -]([A-Za-z]{3})[ -](\d{4})$/);
  if (m) return reportDateFromFilename_(m[1] + '-' + m[2] + '-' + m[3]);
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return iso ? iso[1] + '-' + iso[2] + '-' + iso[3] : '';
}

function regionFromReportLocation_(value) {
  const s = String(value || '').toLowerCase();
  if (/dubai|sharjah|abu dhabi|uae/.test(s)) return 'UAE';
  const regions = ['Andhra Pradesh','Bihar','Chandigarh','Chhattisgarh','Delhi','Goa','Gujarat','Haryana','Himachal Pradesh','Jharkhand','Karnataka','Kerala','Madhya Pradesh','Maharashtra','Odisha','Punjab','Rajasthan','Tamil Nadu','Telangana','Uttar Pradesh','Uttarakhand','West Bengal'];
  for (let i=0;i<regions.length;i++) if (s.indexOf(regions[i].toLowerCase()) >= 0) return regions[i];
  const cityToState = { mumbai:'Maharashtra',pune:'Maharashtra',nagpur:'Maharashtra',nashik:'Maharashtra',bengaluru:'Karnataka',bangalore:'Karnataka',hyderabad:'Telangana',chennai:'Tamil Nadu',kolkata:'West Bengal',delhi:'Delhi',gurugram:'Haryana',gurgaon:'Haryana',noida:'Uttar Pradesh',lucknow:'Uttar Pradesh',jaipur:'Rajasthan',ahmedabad:'Gujarat',surat:'Gujarat',vadodara:'Gujarat',kochi:'Kerala',patna:'Bihar',bhopal:'Madhya Pradesh',dehradun:'Uttarakhand',chandigarh:'Chandigarh' };
  for (const city in cityToState) if (s.indexOf(city) >= 0) return cityToState[city];
  return 'Unmapped';
}

function syncFromSheets() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('Sync skipped: another run is in progress.');
    return;
  }
  try {
    const cfg = getConfig_();
    const snapshot = buildSnapshot_(cfg);
    const json = JSON.stringify(snapshot);
    const hash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json, Utilities.Charset.UTF_8)
    );
    const props = PropertiesService.getScriptProperties();
    if (hash === props.getProperty('LAST_SUCCESSFUL_HASH')) {
      console.log('No source changes; GitHub was not updated.');
      return { status: 'unchanged', asOf: snapshot.meta.asOf };
    }

    publishToGitHub_(cfg, json);
    props.setProperty('LAST_SUCCESSFUL_HASH', hash);
    props.setProperty('LAST_SUCCESSFUL_SYNC', new Date().toISOString());
    console.log('Published dashboard data through ' + snapshot.meta.asOf);
    return { status: 'published', asOf: snapshot.meta.asOf };
  } finally {
    lock.releaseLock();
  }
}

function getConfig_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const cfg = {
    REVIEWS_SHEET_ID: p.REVIEWS_SHEET_ID,
    STOCK_SHEET_ID: p.STOCK_SHEET_ID,
    GITHUB_REPOSITORY: p.GITHUB_REPOSITORY,
    GITHUB_BRANCH: p.GITHUB_BRANCH || 'main',
    GITHUB_FILE_PATH: p.GITHUB_FILE_PATH || 'data.json',
    GITHUB_TOKEN: p.GITHUB_TOKEN,
    MAIL_QUERY: p.MAIL_QUERY || SYNC_CONFIG.MAIL_QUERY
  };
  const required = ['REVIEWS_SHEET_ID', 'STOCK_SHEET_ID', 'GITHUB_REPOSITORY', 'GITHUB_TOKEN'];
  const missing = required.filter(function(k) { return !cfg[k]; });
  if (missing.length) {
    throw new Error('Set these Script Properties first: ' + missing.join(', '));
  }
  if (cfg.GITHUB_REPOSITORY.indexOf('/') < 1) {
    throw new Error('GITHUB_REPOSITORY must use owner/repository format.');
  }
  return cfg;
}

function buildSnapshot_(cfg) {
  const cutoff = yesterdayKey_();
  const reviewBook = SpreadsheetApp.openById(cfg.REVIEWS_SHEET_ID);
  const stockBook = SpreadsheetApp.openById(cfg.STOCK_SHEET_ID);
  const reviewResult = readReviewSummary_(reviewBook, cutoff);
  const stockResult = readStockIn_(stockBook, cutoff);
  const negatives = readNegativeCounts_(reviewBook, cutoff);
  const regionSet = new Set();
  Object.keys(reviewResult.byMonth).forEach(function(month) {
    Object.keys(reviewResult.byMonth[month]).forEach(function(region) { regionSet.add(region); });
  });
  Object.keys(stockResult.byRegionDay).forEach(function(region) { regionSet.add(region); });
  negatives.forEach(function(row) { regionSet.add(row.region); });
  const regions = Array.from(regionSet).filter(Boolean).sort();

  const latestReview = reviewResult.latestDate || cutoff;
  const latestStock = stockResult.latestDate || cutoff;
  return {
    meta: {
      asOf: cutoff,
      sources: {
        reviews: 'GMB Sell Reviews monthly region/day summaries',
        stockIn: 'Customer Review Form (Responses): SI_DATA'
      },
      notes: [
        'Review counts are read from the current monthly region-by-day tabs through ' + latestReview + '.',
        'Positive count = total reviews minus listed negative review records, floored at zero.',
        'Qualified stock-ins require RETAIL_REGION, column AA date, nonblank DRIVER_UPLOAD_START, and VEHICLE_CONDITION exactly equal to ' + SYNC_CONFIG.CONDITION + '. Stock-in rows are read through ' + latestStock + '.',
        'Automatic public sync includes negative-review date and region counts only; it omits review text, reviewer identity, business location, and other row-level details.',
        'Review / Stock-In % is a comparison ratio, not a customer-level conversion rate.'
      ]
    },
    regions: regions,
    reviewsByRegionDay: reviewResult.byMonth,
    stockInByRegionDay: stockResult.byRegionDay,
    negativeReviews: negatives
  };
}

function readReviewSummary_(book, cutoff) {
  const byMonth = {};
  let latestDate = '';
  let matchedTabs = 0;
  book.getSheets().forEach(function(sheet) {
    const month = monthKeyFromTab_(sheet.getName());
    if (!month) return;
    const values = sheet.getDataRange().getValues();
    if (values.length < 3) return;
    matchedTabs++;
    const headers = values[1] || [];
    const dateColumns = [];
    headers.forEach(function(value, col) {
      const key = dateKey_(value);
      if (key && key.slice(0, 7) === month && key <= cutoff) {
        dateColumns.push({ col: col, date: key });
        if (key > latestDate) latestDate = key;
      }
    });
    for (let r = 2; r < values.length; r++) {
      const region = String(values[r][0] || '').trim();
      if (!/^[A-Z]{3}$/.test(region)) continue;
      byMonth[month] = byMonth[month] || {};
      byMonth[month][region] = byMonth[month][region] || {};
      dateColumns.forEach(function(item) {
        byMonth[month][region][item.date] = Number(values[r][item.col]) || 0;
      });
    }
  });
  if (!matchedTabs) {
    throw new Error("No monthly review tabs found. Expected names like Oct'26 or Nov'26.");
  }
  return { byMonth: byMonth, latestDate: latestDate };
}

function readStockIn_(book, cutoff) {
  const sheet = book.getSheetByName(SYNC_CONFIG.STOCK_TAB);
  if (!sheet) throw new Error('Stock-in tab not found: ' + SYNC_CONFIG.STOCK_TAB);
  const lastRow = sheet.getLastRow();
  const byRegionDay = {};
  let latestDate = '';
  if (lastRow < 2) return { byRegionDay: byRegionDay, latestDate: latestDate };
  const rowCount = lastRow - 1;
  const regions = sheet.getRange(2, 6, rowCount, 1).getValues();      // F RETAIL_REGION
  const uploads = sheet.getRange(2, 14, rowCount, 1).getValues();     // N DRIVER_UPLOAD_START
  const conditions = sheet.getRange(2, 23, rowCount, 1).getValues();  // W VEHICLE_CONDITION
  const dates = sheet.getRange(2, 27, rowCount, 1).getValues();       // AA date
  for (let i = 0; i < rowCount; i++) {
    const region = String(regions[i][0] || '').trim();
    const upload = uploads[i][0];
    const condition = String(conditions[i][0] || '').trim();
    const date = dateKey_(dates[i][0]);
    if (!region || upload === '' || upload === null || upload === undefined ||
        condition !== SYNC_CONFIG.CONDITION || !date || date > cutoff) continue;
    byRegionDay[region] = byRegionDay[region] || {};
    byRegionDay[region][date] = (byRegionDay[region][date] || 0) + 1;
    if (date > latestDate) latestDate = date;
  }
  return { byRegionDay: byRegionDay, latestDate: latestDate };
}

function readNegativeCounts_(book, cutoff) {
  const sheet = book.getSheetByName(SYNC_CONFIG.NEGATIVE_TAB);
  if (!sheet) throw new Error('Negative-review tab not found: ' + SYNC_CONFIG.NEGATIVE_TAB);
  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[9] || '').trim().toLowerCase() !== 'negative') continue;
    const date = dateKey_(row[13]);
    if (!date || date > cutoff) continue;
    const suppliedRegion = String(row[18] || '').trim();
    const region = suppliedRegion || regionForLocation_(String(row[3] || ''));
    out.push({
      store: region,
      name: 'Negative review · details not synced',
      location: region,
      rating: '',
      sentiment: 'Negative',
      text: 'Review-level text and identifying details are excluded from automatic public sync.',
      date: displayDate_(date),
      region: region
    });
  }
  return out;
}

function monthKeyFromTab_(name) {
  const m = String(name).match(/^([A-Za-z]{3})['’](\d{2})$/);
  if (!m) return '';
  const months = { Jan:'01', Feb:'02', Mar:'03', Apr:'04', May:'05', Jun:'06',
    Jul:'07', Aug:'08', Sep:'09', Oct:'10', Nov:'11', Dec:'12' };
  const key = months[m[1].slice(0, 1).toUpperCase() + m[1].slice(1, 3).toLowerCase()];
  return key ? '20' + m[2] + '-' + key : '';
}

function dateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, SYNC_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  }
  if (typeof value === 'number' && isFinite(value) && value >= 20000 && value <= 90000) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return Utilities.formatDate(date, SYNC_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  }
  return '';
}

function yesterdayKey_() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, SYNC_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}

function displayDate_(key) {
  const parts = key.split('-');
  const month = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][Number(parts[1]) - 1];
  return parts[2] + '-' + month + '-' + parts[0].slice(2);
}

function regionForLocation_(location) {
  const s = String(location || '').toLowerCase();
  const map = {
    NHR:['gurgaon','gurugram','rohtak','गुरुग्राम'], NRJ:['jaipur','जयपुर'],
    BLR:['bangalore','bengaluru','karnataka','ಬेंಗಳೂರು','ಬೆಂಗಳೂರು','बेंगलुरु'],
    SKL:['lucknow','लखनऊ'], STN:['hyderabad','telangana','हैदराबाद'],
    EWB:['kolkata','howrah','siliguri','west bengal','कोलकाता'],
    NPB:['noida','ghaziabad','नोएडा','गाजियाबाद'], WMP:['pune','पुणे'],
    GJN:['vadodara','વડોદરા','वडोदरा'], EJH:['guwahati','assam','गुवाहाटी'],
    STS:['hyderabad','telangana','हैदराबाद'], EBR:['west bengal','kolkata','howrah'],
    CHN:['chennai','tamil nadu','चेन्नई'], NDL:['delhi','दिल्ली'],
    MUM:['mumbai','मुम्बई','मुंबई'], AMD:['ahmedabad','અમદાવાદ','अहमदाबाद'],
    WMH:['maharashtra','नासिक','नागपुर','नाशिक'], UPE:['uttar pradesh'],
    PUN:['punjab','chandigarh','चंडीगढ़'], GJS:['south gujarat','surat','સુરત'],
    UPW:['west up','meerut','uttarakhand'], NUK:['uttarakhand'],
    EOR:['odisha','orissa','bhubaneswar'], EAS:['assam'],
    SKA:['karnataka'], SAP:['andhra']
  };
  for (const region in map) {
    if (map[region].some(function(term) { return s.indexOf(term) >= 0; })) return region;
  }
  return 'Unmapped';
}

function publishToGitHub_(cfg, json) {
  publishJsonToGitHub_(cfg, cfg.GITHUB_FILE_PATH, json, 'Sync GMB dashboard snapshot through ' + yesterdayKey_());
}

function publishJsonToGitHub_(cfg, filePath, json, commitMessage) {
  const base = 'https://api.github.com/repos/' + cfg.GITHUB_REPOSITORY + '/contents/';
  const path = filePath.split('/').map(encodeURIComponent).join('/');
  const url = base + path + '?ref=' + encodeURIComponent(cfg.GITHUB_BRANCH);
  const headers = {
    Authorization: 'Bearer ' + cfg.GITHUB_TOKEN,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
  const getResponse = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
  const getCode = getResponse.getResponseCode();
  let sha = '';
  if (getCode === 200) {
    sha = JSON.parse(getResponse.getContentText()).sha || '';
  } else if (getCode !== 404) {
    throw new Error('GitHub read failed (' + getCode + '): ' + getResponse.getContentText());
  }

  const payload = {
    message: commitMessage,
    content: Utilities.base64Encode(json, Utilities.Charset.UTF_8),
    branch: cfg.GITHUB_BRANCH
  };
  if (sha) payload.sha = sha;
  const putResponse = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const putCode = putResponse.getResponseCode();
  if (putCode !== 200 && putCode !== 201) {
    throw new Error('GitHub write failed (' + putCode + '): ' + putResponse.getContentText());
  }
}

function getSyncStatus() {
  const p = PropertiesService.getScriptProperties();
  return {
    lastSync: p.getProperty('LAST_SUCCESSFUL_SYNC') || 'not yet synced',
    lastHash: p.getProperty('LAST_SUCCESSFUL_HASH') ? 'saved' : 'not set',
    triggers: ScriptApp.getProjectTriggers().filter(function(t) {
      return [SYNC_CONFIG.HANDLER, SYNC_CONFIG.MAIL_HANDLER].indexOf(t.getHandlerFunction()) >= 0;
    }).length,
    lastEmailSync: p.getProperty('LAST_EMAIL_SYNC') || 'not yet synced',
    mailQuery: p.getProperty('MAIL_QUERY') || SYNC_CONFIG.MAIL_QUERY
  };
}

