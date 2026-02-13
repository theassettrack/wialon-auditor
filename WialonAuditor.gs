// Wialon Auditor — pull fleet config data from Wialon into Google Sheets.
// Single-file GAS script. Paste into Apps Script, deploy as Web App, done.

// --- Config ---

var WIALON_HOST = 'https://hst-api.wialon.com/wialon/ajax.html';
var TOKEN_MAX_AGE_DAYS = 30;
var _sharedSession = null;
var _batchMode = false;

function getConfig() {
  return {
    get: function(key, defaultValue) {
      var defaults = {
        'BATCH_SIZE': '50',
        'PARALLEL_BATCHES': '5',
        'DEBUG_MODE': 'false',
        'WIALON_HOST': WIALON_HOST
      };
      return defaults[key] || defaultValue;
    },
    getBoolean: function(key, defaultValue) {
      var value = this.get(key, String(defaultValue));
      return value === 'true' || value === true;
    },
    getNumber: function(key, defaultValue) {
      var value = this.get(key, String(defaultValue));
      return parseInt(value) || defaultValue;
    }
  };
}

// --- Menu ---

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Wialon Tools')
      .addItem('Login', 'startOAuthLogin')
      .addItem('Login Status', 'showLoginStatus')
      .addSeparator()
      .addItem('Fetch Units', 'fetchUnits')
      .addItem('Fetch Hardware', 'fetchHardware')
      .addItem('Fetch Sensors', 'fetchSensors')
      .addItem('Fetch Commands', 'fetchCommands')
      .addItem('Fetch Profiles', 'fetchProfiles')
      .addItem('Fetch Custom Fields', 'fetchCustomFields')
      .addItem('Fetch Drive Rank', 'fetchDriveRank')
      .addItem('Fetch All', 'fetchAll')
      .addSeparator()
      .addItem('Setup Filter Sheet', 'setupFilterSheet')
      .addItem('Export as JSON', 'exportAsJson')
      .addSeparator()
      .addItem('Save Snapshot', 'saveSnapshot')
      .addItem('Compare with Snapshot', 'compareWithSnapshot')
      .addSeparator()
      .addItem('Get Web App URL', 'getWebAppUrl')
      .addItem('Logout', 'logoutOAuth')
      .addToUi();
  } catch (error) {
    console.log('Menu creation error:', error.message);
  }
}

// --- Auth ---

function startOAuthLogin() {
  var ui = SpreadsheetApp.getUi();

  // !! CHANGE THIS to your deployed Web App URL !!
  var redirectUri = 'PASTE WEB APP URL HERE';

  if (!redirectUri || redirectUri === 'PASTE WEB APP URL HERE') {
    ui.alert('Setup Required',
      'Please update the redirect URL in the code with your Web App URL:\n\n' +
      '1. Go to Deploy > Manage Deployments\n' +
      '2. Copy the Web App URL\n' +
      '3. Update the redirectUri variable in the startOAuthLogin function',
      ui.ButtonSet.OK);
    return;
  }

  var oauthUrl = 'https://hosting.wialon.com/login.html' +
    '?client_id=auditor' +
    '&access_type=-1' +
    '&activation_time=0' +
    '&duration=2592000' +
    '&flags=1' +
    '&redirect_uri=' + encodeURIComponent(redirectUri);

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial, sans-serif; padding: 30px; text-align: center;">' +
      '<h2 style="color: #2c3e50;">Wialon Authentication</h2>' +
      '<div style="margin: 30px 0;">' +
        '<p style="color: #7f8c8d; font-size: 16px; line-height: 1.6;">' +
          'Click the button below to log in to Wialon.<br>' +
          'After login, you\'ll be automatically redirected back.' +
        '</p>' +
      '</div>' +
      '<button onclick="loginToWialon()"' +
        ' style="background-color: #3498db; color: white; border: none;' +
        ' padding: 14px 40px; font-size: 18px; border-radius: 5px;' +
        ' cursor: pointer; margin: 20px;">' +
        'Login to Wialon' +
      '</button>' +
      '<div id="status" style="margin-top: 30px; padding: 15px;"></div>' +
    '</div>' +
    '<script>' +
      'var checkInterval;' +
      'function loginToWialon() {' +
        'window.open("' + oauthUrl + '", "_blank");' +
        'document.getElementById("status").innerHTML =' +
          '"<span style=\\"color: #3498db; font-size: 16px;\\">Waiting for login...<br><br>' +
          'Complete the login in the new window.</span>";' +
        'checkInterval = setInterval(checkLoginStatus, 2000);' +
      '}' +
      'function checkLoginStatus() {' +
        'google.script.run' +
          '.withSuccessHandler(function(status) {' +
            'if (status.loggedIn) {' +
              'clearInterval(checkInterval);' +
              'document.getElementById("status").innerHTML =' +
                '"<span style=\\"color: #27ae60; font-size: 18px;\\">Login successful!<br><br>' +
                'Logged in as: <strong>" + status.user + "</strong></span>";' +
              'setTimeout(function() { google.script.host.close(); }, 2000);' +
            '}' +
          '})' +
          '.withFailureHandler(function(error) {' +
            'console.error("Status check failed:", error);' +
          '})' +
          '.checkLoginStatus();' +
      '}' +
    '</script>'
  )
  .setWidth(450)
  .setHeight(350);

  ui.showModalDialog(html, 'Wialon Login');
}

function getStoredToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var storedToken = scriptProperties.getProperty('WIALON_TOKEN');

  if (storedToken) {
    try {
      var tokenData = JSON.parse(storedToken);
      var tokenAge = (new Date().getTime() - tokenData.timestamp) / 1000 / 60 / 60 / 24;

      if (tokenAge < TOKEN_MAX_AGE_DAYS) {
        return tokenData.token;
      } else {
        scriptProperties.deleteProperty('WIALON_TOKEN');
        return null;
      }
    } catch (e) {
      console.error('Error parsing stored token:', e);
      return null;
    }
  }

  return null;
}

function authenticateWialon() {
  if (_sharedSession) return _sharedSession;

  var token = getStoredToken();
  if (!token) throw new Error('Not logged in. Please use Wialon Tools > Login to authenticate.');

  var session = createSessionFromToken(token);
  if (!session) throw new Error('Authentication failed. Token may be expired. Please login again.');

  return session;
}

function checkLoginStatus() {
  var token = getStoredToken();
  if (!token) return { loggedIn: false, message: 'Not logged in' };

  try {
    var session = createSessionFromToken(token);
    if (session) {
      return { loggedIn: true, user: session.user, message: 'Logged in as: ' + session.user };
    }
    return { loggedIn: false, message: 'Token invalid or expired' };
  } catch (error) {
    return { loggedIn: false, message: 'Authentication check failed' };
  }
}

function showLoginStatus() {
  var ui = SpreadsheetApp.getUi();
  var status = checkLoginStatus();

  if (status.loggedIn) {
    ui.alert('Login Status', status.message, ui.ButtonSet.OK);
  } else {
    var result = ui.alert('Login Status', 'Not logged in. Would you like to login now?', ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) startOAuthLogin();
  }
}

function logoutOAuth() {
  PropertiesService.getScriptProperties().deleteProperty('WIALON_TOKEN');
  SpreadsheetApp.getUi().alert('Logged out successfully');
}

function createSessionFromToken(token) {
  try {
    var response = UrlFetchApp.fetch(WIALON_HOST, {
      method: 'post',
      payload: { svc: 'token/login', params: JSON.stringify({ token: token, fl: 1, operateAs: '' }) },
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.error) { console.error('Session creation error:', result.error); return null; }
    return { eid: result.eid, user: result.user, au: result.au };
  } catch (error) {
    console.error('Session creation error:', error);
    return null;
  }
}

function logoutWialon(sessionId) {
  try {
    UrlFetchApp.fetch(WIALON_HOST, {
      method: 'post',
      payload: { svc: 'core/logout', params: JSON.stringify({}), sid: sessionId },
      muteHttpExceptions: true
    });
    return true;
  } catch (error) {
    console.error('Logout error:', error);
    return false;
  }
}

function makeWialonRequest(sessionId, service, requestBody) {
  try {
    var response = UrlFetchApp.fetch(WIALON_HOST, {
      method: 'post',
      payload: { svc: service, params: JSON.stringify(requestBody), sid: sessionId },
      muteHttpExceptions: true
    });

    var responseText = response.getContentText();
    if (!responseText || responseText.trim() === '') {
      throw new Error('Empty response from ' + service);
    }

    var result = JSON.parse(responseText);
    if (result.error) throw new Error('API Error: ' + result.error);
    return result;
  } catch (error) {
    console.error('Request failed:', service, error.message);
    throw error;
  }
}

// --- Web App (OAuth callback) ---

function doGet(e) {
  var params = e.parameter;

  if (params.access_token) {
    PropertiesService.getScriptProperties().setProperty('WIALON_TOKEN', JSON.stringify({
      token: params.access_token,
      timestamp: new Date().getTime(),
      user: params.user_name || 'Unknown'
    }));

    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<title>Login Successful</title></head>' +
      '<body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">' +
      '<div style="max-width: 500px; margin: 0 auto; padding: 30px; border: 2px solid #27ae60; border-radius: 10px;">' +
      '<h1 style="color: #27ae60;">Login Successful!</h1>' +
      '<p style="color: #34495e; font-size: 16px;">Your token has been saved. You can close this window and return to Google Sheets.</p>' +
      '<button onclick="window.close()" style="margin-top: 20px; padding: 10px 30px; background-color: #3498db; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px;">Close Window</button>' +
      '</div>' +
      '<script>setTimeout(function() { window.close(); }, 3000);</script>' +
      '</body></html>'
    );
  }

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><title>Login Failed</title></head>' +
    '<body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">' +
    '<div style="max-width: 500px; margin: 0 auto; padding: 30px; border: 2px solid #e74c3c; border-radius: 10px;">' +
    '<h1 style="color: #e74c3c;">Login Failed</h1>' +
    '<p style="color: #34495e;">No access token received. Please try logging in again.</p>' +
    '</div></body></html>'
  );
}

function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();

  if (url) {
    ui.alert('Web App URL', 'Your callback URL is:\n\n' + url, ui.ButtonSet.OK);
    return url;
  }

  ui.alert('Not Deployed',
    'Deploy this script as a Web App first:\n\n' +
    '1. Deploy > New Deployment\n' +
    '2. Choose: Web app\n' +
    '3. Access: Anyone\n' +
    '4. Deploy and copy the URL',
    ui.ButtonSet.OK);
  return null;
}

// --- Sheet helpers ---

function getOrCreateSheet(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

// Writes rows in 500-row chunks, falls back to 100 then row-by-row if GAS chokes
function appendToSheet(sheet, data, includeHeader) {
  if (!sheet || !data || data.length === 0) return;

  var columns = Object.keys(data[0]);
  var rows = [];

  if (includeHeader) rows.push(columns);

  data.forEach(function(item) {
    rows.push(columns.map(function(col) {
      var value = item[col];
      if (value === null || value === undefined) return '';
      if (typeof value === 'object') return JSON.stringify(value);
      return value;
    }));
  });

  var startRow = includeHeader ? 1 : sheet.getLastRow() + 1;

  var CHUNK = 500;
  for (var i = 0; i < rows.length; i += CHUNK) {
    var chunk = rows.slice(i, Math.min(i + CHUNK, rows.length));
    var targetRow = startRow + i;

    try {
      sheet.getRange(targetRow, 1, chunk.length, columns.length).setValues(chunk);
      if (i > 0 && i % (CHUNK * 5) === 0) Utilities.sleep(100);
    } catch (e) {
      // Fallback: smaller chunks
      for (var j = 0; j < chunk.length; j += 100) {
        var small = chunk.slice(j, Math.min(j + 100, chunk.length));
        try {
          sheet.getRange(targetRow + j, 1, small.length, columns.length).setValues(small);
          Utilities.sleep(50);
        } catch (e2) {
          // Last resort: row by row
          for (var k = 0; k < small.length; k++) {
            try { sheet.getRange(targetRow + j + k, 1, 1, columns.length).setValues([small[k]]); }
            catch (e3) { console.error('Failed row ' + (targetRow + j + k)); }
          }
        }
      }
    }
  }

  if (includeHeader) {
    try {
      sheet.getRange(1, 1, 1, columns.length)
        .setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    } catch (e) { }
  }
}

// --- Internal helpers ---

function buildSearchSpec_() {
  return {
    itemsType: 'avl_unit',
    propName: 'sys_name',
    propValueMask: '*',
    sortType: 'sys_name',
    propType: 'property',
    or_logic: 0
  };
}

function getTotalUnitCount_(sessionId) {
  var result = makeWialonRequest(sessionId, 'core/search_items', {
    spec: buildSearchSpec_(), force: 1, flags: 1, from: 0, to: 0
  });
  return result.totalItemsCount || (result.items ? result.items.length : 0);
}

// Fetches all units in batches of `batchSize`, retries with smaller chunks on failure
function fetchUnitsBatched_(sessionId, flags, batchSize) {
  batchSize = batchSize || 1000;
  var spec = buildSearchSpec_();
  var total = getTotalUnitCount_(sessionId);
  var units = [];

  for (var from = 0; from < total; from += batchSize) {
    var to = Math.min(from + batchSize - 1, total - 1);

    try {
      var res = makeWialonRequest(sessionId, 'core/search_items', {
        spec: spec, force: 1, flags: flags, from: from, to: to
      });
      if (res.items) units.push.apply(units, res.items);
    } catch (error) {
      console.error('Batch ' + from + '-' + to + ' failed, retrying smaller:', error);
      var smaller = Math.min(200, batchSize);
      for (var i = from; i <= to; i += smaller) {
        var smallTo = Math.min(i + smaller - 1, to);
        try {
          var r = makeWialonRequest(sessionId, 'core/search_items', {
            spec: spec, force: 1, flags: flags, from: i, to: smallTo
          });
          if (r.items) units.push.apply(units, r.items);
        } catch (e) { console.error('Sub-batch ' + i + '-' + smallTo + ' also failed'); }
      }
    }

    if (from + batchSize < total) Utilities.sleep(100);
  }

  return units;
}

// Runs many API calls in parallel using core/batch + UrlFetchApp.fetchAll
// paramBuilder(unit) should return { svc: '...', params: {...} }
function parallelBatchFetch_(sessionId, units, paramBuilder, resultKey) {
  var url = getConfig().get('WIALON_HOST', WIALON_HOST);
  var BATCH = 100, PARALLEL = 10, DELAY = 100;
  var data = {};

  for (var i = 0; i < units.length; i += BATCH * PARALLEL) {
    var requests = [];
    var groups = [];

    for (var p = 0; p < PARALLEL && i + p * BATCH < units.length; p++) {
      var start = i + p * BATCH;
      var end = Math.min(start + BATCH, units.length);
      var batch = units.slice(start, end);
      if (batch.length === 0) break;

      var params = [];
      batch.forEach(function(u) { params.push(paramBuilder(u)); });

      requests.push({
        url: url, method: 'post',
        payload: { svc: 'core/batch', params: JSON.stringify({ params: params, flags: 0 }), sid: sessionId },
        muteHttpExceptions: true
      });
      groups.push(batch);
    }

    try {
      var pct = Math.round(Math.min(i + BATCH * PARALLEL, units.length) / units.length * 100);
      console.log('Fetching ' + resultKey + ': ' + pct + '%');
      var responses = UrlFetchApp.fetchAll(requests);

      responses.forEach(function(resp, gi) {
        try {
          var text = resp.getContentText();
          if (!text || !text.trim()) return;
          var parsed = JSON.parse(text);
          var idx = 0;
          groups[gi].forEach(function(u) {
            var entry = {};
            entry[resultKey] = parsed[idx++];
            data[u.id] = entry;
          });
        } catch (e) { console.error('Batch group ' + gi + ' parse error:', e); }
      });
    } catch (error) { console.error('Batch request error:', error); }

    if (i + BATCH * PARALLEL < units.length) Utilities.sleep(DELAY);
  }

  return data;
}

function notify_(title, message) {
  if (_batchMode) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
  } else {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function safeLogout_(sessionId) {
  if (!_batchMode) logoutWialon(sessionId);
}

// Shared workflow: auth -> fetch units -> process -> write sheet -> logout -> alert
function standardFetchWorkflow_(opts) {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();

  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Authentication failed.'); return; }

    var units = fetchUnitsBatched_(session.eid, opts.flags);
    if (units.length === 0) {
      ui.alert('Error', 'No units found.', ui.ButtonSet.OK);
      safeLogout_(session.eid);
      return;
    }

    var sheet = getOrCreateSheet(spreadsheet, opts.sheetName);
    sheet.clear();

    var result = opts.processUnits(units, session);
    var rows = result.rows || result;
    var extras = result.extras || {};

    for (var i = 0; i < rows.length; i += 1000) {
      appendToSheet(sheet, rows.slice(i, Math.min(i + 1000, rows.length)), i === 0);
      if (i + 1000 < rows.length) Utilities.sleep(100);
    }

    safeLogout_(session.eid);
    extras.totalTime = Math.round((new Date().getTime() - startTime) / 1000);
    extras.totalProcessed = units.length;
    notify_(opts.successTitle, opts.summaryMessage(extras));
  } catch (error) {
    console.error(opts.sheetName + ' fetch error:', error);
    ui.alert('Error', opts.sheetName + ' fetch failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// --- Hardware types ---

function fetchHardwareTypes(sessionId, hwIds) {
  var url = getConfig().get('WIALON_HOST', WIALON_HOST);

  var svcParams;
  if (hwIds && (hwIds.size > 0 || hwIds.length > 0)) {
    var idArray = hwIds.size ? Array.from(hwIds) : hwIds;
    svcParams = JSON.stringify({ filterType: 'id', filterValue: idArray });
  } else {
    svcParams = JSON.stringify({});
  }

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: { svc: 'core/get_hw_types', params: svcParams, sid: sessionId },
      muteHttpExceptions: true
    });

    var text = response.getContentText();
    if (!text || !text.trim()) return {};

    var result = JSON.parse(text);
    if (result.error) return {};

    var types = {};
    if (Array.isArray(result)) {
      result.forEach(function(hw) {
        types[hw.id] = {
          id: hw.id, uid2: hw.uid2, name: hw.name,
          hw_category: hw.hw_category || '', hw_features: hw.hw_features || '',
          tp: hw.tp || '', up: hw.up || '', type: hw.type || ''
        };
      });
    }
    return types;
  } catch (error) {
    console.error('Error fetching hardware types:', error);
    return {};
  }
}

function batchFetchHardwareNames(sessionId, units) {
  var uniqueHwIds = new Set();
  units.forEach(function(u) { if (u.hw) uniqueHwIds.add(u.hw); });

  var hwTypes = fetchHardwareTypes(sessionId, uniqueHwIds);

  var hwData = {};
  units.forEach(function(u) {
    hwData[u.id] = {
      hw_id: u.hw || 0,
      hw_name: (u.hw && hwTypes[u.hw]) ? (hwTypes[u.hw].name || '') : ''
    };
  });
  return hwData;
}

// --- Data fetchers ---

// Flags: 1(base) + 2(custom props) + 4(billing) + 256(advanced) + 8192(counters) + 131072(config) = 139527
function fetchUnits() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();

  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Authentication failed'); return; }

    var spec = buildSearchSpec_();
    var totalCount = getTotalUnitCount_(session.eid);

    var unitsSheet = getOrCreateSheet(spreadsheet, 'units');
    unitsSheet.clear();

    var FLAGS = 139527;
    var BATCH = 1000;
    var isFirst = true;
    var processed = 0;

    for (var from = 0; from < totalCount; from += BATCH) {
      var to = Math.min(from + BATCH - 1, totalCount - 1);

      var batchUnits = [];
      try {
        var res = makeWialonRequest(session.eid, 'core/search_items', {
          spec: spec, force: 1, flags: FLAGS, from: from, to: to
        });
        if (res.items) batchUnits = res.items;
      } catch (error) {
        for (var i = from; i <= to; i += 200) {
          var st = Math.min(i + 199, to);
          try {
            var r = makeWialonRequest(session.eid, 'core/search_items', {
              spec: spec, force: 1, flags: FLAGS, from: i, to: st
            });
            if (r.items) batchUnits.push.apply(batchUnits, r.items);
          } catch (e) { }
        }
      }

      if (batchUnits.length === 0) continue;

      var hwData = batchFetchHardwareNames(session.eid, batchUnits);

      var rows = [];
      batchUnits.forEach(function(u) {
        rows.push({
          id: u.id || '', nm: u.nm || '', cls: u.cls || 0, mu: u.mu || 0,
          uacl: u.uacl || '', ct: u.ct || 0,
          prp: JSON.stringify(u.prp || {}),
          crt: u.crt || 0, bact: u.bact || 0,
          hw: u.hw || 0, hw_name: (hwData[u.id] && hwData[u.id].hw_name) || '',
          ph: u.ph || '', ph2: u.ph2 || '', psw: u.psw || '',
          uid: u.uid || '', uid2: u.uid2 || '',
          act: u.act || 0, act_reason: u.act_reason || 0, dactt: u.dactt || 0,
          ftp: JSON.stringify(u.ftp || {}),
          cfl: u.cfl || 0, cnm: u.cnm || 0, cnm_km: u.cnm_km || 0,
          cneh: u.cneh || 0, cnkb: u.cnkb || 0,
          hch: JSON.stringify(u.hch || {}),
          // trip detector
          rtd_type: (u.rtd && u.rtd.type) || 0,
          rtd_gps_correction: (u.rtd && u.rtd.gpsCorrection) || 0,
          rtd_min_sat: (u.rtd && u.rtd.minSat) || 0,
          rtd_min_moving_speed: (u.rtd && u.rtd.minMovingSpeed) || 0,
          rtd_min_stay_time: (u.rtd && u.rtd.minStayTime) || 0,
          rtd_max_msg_distance: (u.rtd && u.rtd.maxMessagesDistance) || 0,
          rtd_min_trip_time: (u.rtd && u.rtd.minTripTime) || 0,
          rtd_min_trip_distance: (u.rtd && u.rtd.minTripDistance) || 0,
          // fuel consumption
          rfc_calc_types: (u.rfc && u.rfc.calcTypes) || 0,
          rfc_fll_flags: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.flags) || 0,
          rfc_fll_ignore_stay: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.ignoreStayTimeout) || 0,
          rfc_fll_min_fill_volume: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.minFillingVolume) || 0,
          rfc_fll_min_theft_timeout: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.minTheftTimeout) || 0,
          rfc_fll_min_theft_volume: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.minTheftVolume) || 0,
          rfc_fll_filter_quality: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.filterQuality) || 0,
          rfc_fll_fill_join_interval: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.fillingsJoinInterval) || 0,
          rfc_fll_theft_join_interval: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.theftsJoinInterval) || 0,
          rfc_fll_extra_fill_timeout: (u.rfc && u.rfc.fuelLevelParams && u.rfc.fuelLevelParams.extraFillingTimeout) || 0,
          rfc_math_idling: (u.rfc && u.rfc.fuelConsMath && u.rfc.fuelConsMath.idling) || 0,
          rfc_math_urban: (u.rfc && u.rfc.fuelConsMath && u.rfc.fuelConsMath.urban) || 0,
          rfc_math_suburban: (u.rfc && u.rfc.fuelConsMath && u.rfc.fuelConsMath.suburban) || 0,
          rfc_rates_summer: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.consSummer) || 0,
          rfc_rates_winter: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.consWinter) || 0,
          rfc_rates_winter_month_from: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.winterMonthFrom) || 0,
          rfc_rates_winter_day_from: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.winterDayFrom) || 0,
          rfc_rates_winter_month_to: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.winterMonthTo) || 0,
          rfc_rates_winter_day_to: (u.rfc && u.rfc.fuelConsRates && u.rfc.fuelConsRates.winterDayTo) || 0,
          rfc_impulse_max: (u.rfc && u.rfc.fuelConsImpulse && u.rfc.fuelConsImpulse.maxImpulses) || 0,
          rfc_impulse_skip_zero: (u.rfc && u.rfc.fuelConsImpulse && u.rfc.fuelConsImpulse.skipZero) || 0
        });
      });

      appendToSheet(unitsSheet, rows, isFirst);
      isFirst = false;
      processed += rows.length;

      batchUnits = null;
      rows.length = 0;
      if (from + BATCH < totalCount) Utilities.sleep(100);
    }

    safeLogout_(session.eid);

    var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
    notify_('Units Fetch Complete', 'Fetched ' + processed + ' units\nTime: ' + elapsed + 's');
  } catch (error) {
    console.error('Error:', error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

function fetchHardware() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();

  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Authentication failed.'); return; }

    // flag 257 = 1(base) + 256(advanced) — gives us the hw field
    var units = fetchUnitsBatched_(session.eid, 257);
    if (units.length === 0) {
      ui.alert('No units found.'); safeLogout_(session.eid); return;
    }

    var uniqueHwIds = new Set();
    units.forEach(function(u) { if (u.hw) uniqueHwIds.add(u.hw); });
    var hwTypes = fetchHardwareTypes(session.eid, uniqueHwIds);

    var sheet = getOrCreateSheet(spreadsheet, 'hardware');
    sheet.clear();

    var hwData = parallelBatchFetch_(session.eid, units, function(u) {
      return { svc: 'unit/update_hw_params', params: { itemId: u.id, hwId: u.hw || 0, fullData: 1, action: 'get' } };
    }, 'hw_params');

    var results = [];
    var withHw = 0;

    units.forEach(function(u) {
      if (!hwData || !hwData[u.id] || !hwData[u.id].hw_params) return;
      var params = hwData[u.id].hw_params;
      if (!params || params.error || !Array.isArray(params)) return;

      withHw++;
      var info = (hwTypes && u.hw && hwTypes[u.hw]) ? hwTypes[u.hw] : {};

      var allParams = {};
      var modified = 0;
      params.forEach(function(p) {
        if (!p || !p.name) return;
        allParams[p.name] = p.value || '';
        if (p.value !== p.default) modified++;
      });

      results.push({
        id: u.id, nm: u.nm || '',
        hw_id: u.hw || '', hw_name: info.name || 'Unknown',
        param_count: params.length, modified_count: modified,
        hw_info: JSON.stringify({ id: u.hw || '', name: info.name || 'Unknown', uid2: info.uid2 || '',
          category: info.hw_category || '', features: info.hw_features || '',
          type: info.type || '', tcp_port: info.tp || '', udp_port: info.up || '' }),
        params: JSON.stringify(allParams)
      });
    });

    for (var i = 0; i < results.length; i += 1000) {
      appendToSheet(sheet, results.slice(i, Math.min(i + 1000, results.length)), i === 0);
    }

    safeLogout_(session.eid);
    var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
    notify_('Hardware Fetch Complete',
      units.length + ' units, ' + withHw + ' with hardware, ' +
      uniqueHwIds.size + ' types\nTime: ' + elapsed + 's');
  } catch (error) {
    console.error('Hardware fetch error:', error);
    ui.alert('Error', 'Hardware fetch failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function fetchDriveRank() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Authentication failed.'); return; }

    // flag 4097 = 1(base) + 4096(sensors) — need sensors to resolve validator names
    var units = fetchUnitsBatched_(session.eid, 4097);
    if (units.length === 0) {
      ui.alert('No units found.'); safeLogout_(session.eid); return;
    }

    var unitSensors = {};
    units.forEach(function(u) { unitSensors[u.id] = u.sens || {}; });

    var drData = parallelBatchFetch_(session.eid, units, function(u) {
      return { svc: 'unit/get_drive_rank_settings', params: { itemId: u.id } };
    }, 'drive_rank_settings');

    var sheet = getOrCreateSheet(spreadsheet, 'drive_rank');
    sheet.clear();

    var results = [];
    var withDR = 0;
    var DR_TYPES = ['acceleration', 'brake', 'turn', 'speeding', 'harsh'];

    function resolveValidator(unitId, validatorId) {
      if (!validatorId) return '';
      var key = String(validatorId);
      return (unitSensors[unitId] && unitSensors[unitId][key]) ? (unitSensors[unitId][key].n || '') : '';
    }

    function buildSettings(item) {
      return JSON.stringify({
        flags: item.flags || 0, min_value: item.min_value || '', max_value: item.max_value || '',
        min_speed: item.min_speed || '', max_speed: item.max_speed || '',
        min_duration: item.min_duration || '', max_duration: item.max_duration || '',
        penalties: item.penalties || 0
      });
    }

    units.forEach(function(u) {
      if (!drData || !drData[u.id] || !drData[u.id].drive_rank_settings) return;
      var settings = drData[u.id].drive_rank_settings;
      if (!settings) return;
      withDR++;

      DR_TYPES.forEach(function(type) {
        if (!settings[type]) return;
        settings[type].forEach(function(item) {
          results.push({
            id: u.id, nm: u.nm || '', type: type, name: item.name || '',
            validator_id: item.validator_id || '', validator_name: resolveValidator(u.id, item.validator_id),
            sensor_id: '', sensor_name: '',
            settings: buildSettings(item)
          });
        });
      });

      // Custom sensor-based criteria
      if (settings.sensor) {
        settings.sensor.forEach(function(item) {
          var sKey = item.sensor_id ? String(item.sensor_id) : '';
          var sName = (sKey && unitSensors[u.id] && unitSensors[u.id][sKey]) ? (unitSensors[u.id][sKey].n || '') : '';

          results.push({
            id: u.id, nm: u.nm || '', type: 'sensor', name: item.name || '',
            validator_id: item.validator_id || '', validator_name: resolveValidator(u.id, item.validator_id),
            sensor_id: item.sensor_id || '', sensor_name: sName,
            settings: buildSettings(item)
          });
        });
      }
    });

    for (var i = 0; i < results.length; i += 1000) {
      appendToSheet(sheet, results.slice(i, Math.min(i + 1000, results.length)), i === 0);
      if (i + 1000 < results.length) Utilities.sleep(100);
    }

    safeLogout_(session.eid);
    notify_('Drive Rank Complete',
      units.length + ' units, ' + withDR + ' with drive rank, ' + results.length + ' settings total');
  } catch (error) {
    console.error('Drive rank error:', error);
    ui.alert('Error', 'Drive rank fetch failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// flag 4097 = 1(base) + 4096(sensors)
function fetchSensors() {
  standardFetchWorkflow_({
    sheetName: 'sensors', flags: 4097, successTitle: 'Sensors Fetch Complete',
    processUnits: function(units) {
      var total = 0;
      var rows = [];
      units.forEach(function(u) {
        var ids = {}, names = {}, types = {}, measurements = {}, params = {};
        var flags = {}, configs = {}, valTypes = {}, valSensors = {}, tables = {};
        var count = 0;

        if (u.sens && typeof u.sens === 'object') {
          Object.keys(u.sens).forEach(function(k) {
            var s = u.sens[k];
            ids[k] = s.id || 0; names[k] = s.n || ''; types[k] = s.t || '';
            measurements[k] = s.m || ''; params[k] = s.p || ''; flags[k] = s.f || 0;
            configs[k] = s.c || ''; valTypes[k] = s.vt || 0; valSensors[k] = s.vs || 0;
            tables[k] = s.tbl || [];
            count++;
          });
          total += count;
        }

        rows.push({
          id: u.id || '', nm: u.nm || '', sensors_count: count,
          sensor_ids: JSON.stringify(ids), sensor_names: JSON.stringify(names),
          sensor_types: JSON.stringify(types), sensor_measurements: JSON.stringify(measurements),
          sensor_parameters: JSON.stringify(params), sensor_flags: JSON.stringify(flags),
          sensor_configs: JSON.stringify(configs), sensor_validation_types: JSON.stringify(valTypes),
          sensor_validation_sensors: JSON.stringify(valSensors), sensor_tables: JSON.stringify(tables)
        });
      });
      return { rows: rows, extras: { totalSensors: total } };
    },
    summaryMessage: function(x) {
      return x.totalProcessed + ' units, ' + (x.totalSensors || 0) + ' sensors\nTime: ' + x.totalTime + 's';
    }
  });
}

// flag 524289 = 1(base) + 524288(commands)
function fetchCommands() {
  standardFetchWorkflow_({
    sheetName: 'commands', flags: 524289, successTitle: 'Commands Fetch Complete',
    processUnits: function(units) {
      var total = 0;
      var rows = [];
      units.forEach(function(u) {
        var ids = {}, names = {}, types = {}, links = {}, params = {}, access = {}, phoneFlags = {};
        var count = 0;

        if (u.cml && typeof u.cml === 'object') {
          Object.keys(u.cml).forEach(function(k) {
            var c = u.cml[k];
            ids[k] = c.id || 0; names[k] = c.n || ''; types[k] = c.c || '';
            links[k] = c.l || ''; params[k] = c.p || ''; access[k] = c.a || 0;
            phoneFlags[k] = c.f || 0;
            count++;
          });
          total += count;
        }

        rows.push({
          id: u.id || '', nm: u.nm || '', commands_count: count,
          command_ids: JSON.stringify(ids), command_names: JSON.stringify(names),
          command_types: JSON.stringify(types), command_link_types: JSON.stringify(links),
          command_parameters: JSON.stringify(params), command_access_levels: JSON.stringify(access),
          command_phone_flags: JSON.stringify(phoneFlags)
        });
      });
      return { rows: rows, extras: { totalCommands: total } };
    },
    summaryMessage: function(x) {
      return x.totalProcessed + ' units, ' + (x.totalCommands || 0) + ' commands\nTime: ' + x.totalTime + 's';
    }
  });
}

// flag 137 = 1(base) + 8(custom fields) + 128(admin fields)
function fetchCustomFields() {
  standardFetchWorkflow_({
    sheetName: 'custom_fields', flags: 137, successTitle: 'Custom Fields Fetch Complete',
    processUnits: function(units) {
      var rows = [];
      units.forEach(function(u) {
        rows.push({
          id: u.id || '', nm: u.nm || '',
          flds: JSON.stringify(u.flds || {}), flds_count: u.flds ? Object.keys(u.flds).length : 0,
          aflds: JSON.stringify(u.aflds || {}), aflds_count: u.aflds ? Object.keys(u.aflds).length : 0
        });
      });
      return rows;
    },
    summaryMessage: function(x) {
      return x.totalProcessed + ' units fetched\nTime: ' + x.totalTime + 's';
    }
  });
}

// flag 8388609 = 1(base) + 8388608(profiles)
function fetchProfiles() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var startTime = new Date().getTime();
    var session = authenticateWialon();
    if (!session || !session.eid) throw new Error('Authentication failed');

    // to: 0 returns all items at once
    var res = makeWialonRequest(session.eid, 'core/search_items', {
      spec: buildSearchSpec_(), force: 1, flags: 8388609, from: 0, to: 0
    });
    if (!res || !res.items) throw new Error('No units found');

    var units = res.items;
    var fields = [
      'vehicle_type', 'vehicle_class', 'vin', 'registration_plate',
      'brand', 'model', 'year', 'color', 'engine_model', 'engine_power',
      'engine_displacement', 'primary_fuel_type', 'co2_emission',
      'cargo_type', 'carrying_capacity', 'width', 'height', 'depth',
      'effective_capacity', 'gross_vehicle_weight', 'axles'
    ];

    var sheet = getOrCreateSheet(spreadsheet, 'profiles');
    sheet.clear();

    var headers = ['id', 'nm'].concat(fields);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold');

    var withProfiles = 0;

    for (var c = 0; c < units.length; c += 1000) {
      var chunk = units.slice(c, Math.min(c + 1000, units.length));
      var data = [];

      for (var i = 0; i < chunk.length; i++) {
        var u = chunk[i];
        var row = [u.id, u.nm];
        var hasData = false;

        for (var f = 0; f < fields.length; f++) {
          var val = '';
          if (u.pflds) {
            for (var fid in u.pflds) {
              if (u.pflds[fid].n === fields[f] && u.pflds[fid].v) {
                val = u.pflds[fid].v; hasData = true; break;
              }
            }
          }
          row.push(val);
        }
        if (hasData) withProfiles++;
        data.push(row);
      }

      if (data.length > 0) {
        sheet.getRange(2 + c, 1, data.length, headers.length).setValues(data);
      }
    }

    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    safeLogout_(session.eid);

    var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
    notify_('Profiles Complete',
      units.length + ' units, ' + withProfiles + ' with profiles\nTime: ' + elapsed + 's');
  } catch (error) {
    console.error('Profiles error:', error);
    ui.alert('Error', 'Profiles fetch failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// --- Fetch All ---

function fetchAll() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('Fetch All',
    'This will fetch all 7 data categories sequentially.\nMay take several minutes for large fleets.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();

  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Authentication failed.'); return; }

    _sharedSession = session;
    _batchMode = true;

    var steps = [
      { name: 'Units', fn: fetchUnits },
      { name: 'Hardware', fn: fetchHardware },
      { name: 'Sensors', fn: fetchSensors },
      { name: 'Commands', fn: fetchCommands },
      { name: 'Profiles', fn: fetchProfiles },
      { name: 'Custom Fields', fn: fetchCustomFields },
      { name: 'Drive Rank', fn: fetchDriveRank }
    ];

    for (var i = 0; i < steps.length; i++) {
      ss.toast(steps[i].name + '... (' + (i + 1) + '/' + steps.length + ')', 'Fetch All', 120);
      try { steps[i].fn(); } catch (e) { console.error(steps[i].name + ' failed:', e); }
    }

    _batchMode = false;
    _sharedSession = null;
    logoutWialon(session.eid);

    var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
    ui.alert('Fetch All Complete', 'All 7 categories fetched in ' + elapsed + 's', ui.ButtonSet.OK);
  } catch (error) {
    _batchMode = false;
    _sharedSession = null;
    console.error('Fetch All error:', error);
    ui.alert('Error', 'Fetch All failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// --- Export ---

function exportAsJson() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ['units', 'hardware', 'sensors', 'commands', 'profiles', 'custom_fields', 'drive_rank'];
  var output = {};

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        if (headers[j]) row[headers[j]] = data[i][j];
      }
      rows.push(row);
    }

    output[name] = rows;
  });

  if (Object.keys(output).length === 0) {
    ui.alert('Nothing to export. Fetch data first.');
    return;
  }

  var json = JSON.stringify(output, null, 2);
  var fileName = ss.getName() + '_export_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.json';
  var file = DriveApp.createFile(fileName, json, MimeType.PLAIN_TEXT);

  ui.alert('Export Complete', 'Saved to Google Drive:\n' + fileName + '\n\n' + file.getUrl(), ui.ButtonSet.OK);
}

// --- Snapshot & Diff ---

var DATA_SHEETS_ = ['units', 'hardware', 'sensors', 'commands', 'profiles', 'custom_fields', 'drive_rank'];

function saveSnapshot() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var saved = 0;

  DATA_SHEETS_.forEach(function(name) {
    var src = ss.getSheetByName(name);
    if (!src || src.getLastRow() < 2) return;

    var snapName = '_snap_' + name;
    var snap = ss.getSheetByName(snapName);
    if (snap) ss.deleteSheet(snap);

    src.copyTo(ss).setName(snapName).hideSheet();
    saved++;
  });

  if (saved === 0) {
    ui.alert('Nothing to snapshot. Fetch data first.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('SNAPSHOT_TIME', new Date().toISOString());
  ui.alert('Snapshot Saved', saved + ' sheets saved.\nFetch new data, then use "Compare with Snapshot" to see changes.', ui.ButtonSet.OK);
}

function getRowKey_(headers, row, sheetName) {
  var idCol = headers.indexOf('id');
  var key = String(row[idCol >= 0 ? idCol : 0]);
  // drive_rank has multiple rows per unit — use composite key
  if (sheetName === 'drive_rank') {
    var typeCol = headers.indexOf('type');
    var nameCol = headers.indexOf('name');
    if (typeCol >= 0) key += '|' + row[typeCol];
    if (nameCol >= 0) key += '|' + row[nameCol];
  }
  return key;
}

function compareWithSnapshot() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var snapTime = PropertiesService.getScriptProperties().getProperty('SNAPSHOT_TIME');
  if (!snapTime) {
    ui.alert('No snapshot found. Use "Save Snapshot" first.');
    return;
  }

  var diffs = [];

  DATA_SHEETS_.forEach(function(name) {
    var current = ss.getSheetByName(name);
    var snap = ss.getSheetByName('_snap_' + name);
    if (!current && !snap) return;

    var curData = current ? current.getDataRange().getValues() : [[]];
    var snapData = snap ? snap.getDataRange().getValues() : [[]];
    if (curData.length < 2 && snapData.length < 2) return;

    var headers = curData.length > 0 ? curData[0] : snapData[0];
    var snapHeaders = snapData.length > 0 ? snapData[0] : [];
    var nmCol = headers.indexOf('nm');
    if (nmCol === -1) nmCol = 1;

    // Index both by row key
    var snapMap = {};
    for (var i = 1; i < snapData.length; i++) {
      var sk = getRowKey_(snapHeaders, snapData[i], name);
      snapMap[sk] = i;
    }

    var curMap = {};
    for (var i = 1; i < curData.length; i++) {
      var ck = getRowKey_(headers, curData[i], name);
      curMap[ck] = i;
    }

    // Find added and modified
    for (var key in curMap) {
      var cr = curMap[key];
      if (!snapMap[key]) {
        diffs.push({ sheet: name, unit: curData[cr][nmCol] || key, field: '--', change: 'added', old_value: '', new_value: '(new row)' });
        continue;
      }

      var sr = snapMap[key];
      for (var j = 0; j < headers.length; j++) {
        var cv = String(curData[cr][j] || '');
        var sv = (j < snapData[sr].length) ? String(snapData[sr][j] || '') : '';
        if (cv !== sv) {
          diffs.push({ sheet: name, unit: curData[cr][nmCol] || key, field: headers[j] || 'col_' + j, change: 'modified', old_value: sv, new_value: cv });
        }
      }
      delete snapMap[key];
    }

    // Remaining in snapMap are removed
    var snapNmCol = snapHeaders.indexOf('nm');
    if (snapNmCol === -1) snapNmCol = 1;
    for (var key in snapMap) {
      var sr = snapMap[key];
      diffs.push({ sheet: name, unit: snapData[sr][snapNmCol] || key, field: '--', change: 'removed', old_value: '(deleted row)', new_value: '' });
    }
  });

  if (diffs.length === 0) {
    ui.alert('No Changes', 'Current data matches the snapshot from\n' + snapTime, ui.ButtonSet.OK);
    return;
  }

  var diffSheet = getOrCreateSheet(ss, 'diff');
  diffSheet.clear();
  appendToSheet(diffSheet, diffs, true);

  // Color-code rows
  var lastRow = diffSheet.getLastRow();
  if (lastRow > 1) {
    var changeCol = Object.keys(diffs[0]).indexOf('change') + 1; // 1-indexed
    var changes = diffSheet.getRange(2, changeCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < changes.length; i++) {
      var color = '#ffffff';
      if (changes[i][0] === 'added') color = '#d4edda';
      else if (changes[i][0] === 'removed') color = '#f8d7da';
      else if (changes[i][0] === 'modified') color = '#fff3cd';
      diffSheet.getRange(i + 2, 1, 1, Object.keys(diffs[0]).length).setBackground(color);
    }
  }

  ss.setActiveSheet(diffSheet);
  ui.alert('Comparison Complete', diffs.length + ' changes found vs snapshot from\n' + snapTime, ui.ButtonSet.OK);
}

// --- QUERYUNIT custom formula ---

// Use in cells: =QUERYUNIT("Unit Name") or =QUERYUNIT("Unit Name", "sensors")
// Sections: unit, sensors, commands, profiles, custom_fields, hardware, drive_rank
// @customfunction
function QUERYUNIT(unitNameOrId, section) {
  if (!unitNameOrId) return [['Please provide a unit name or ID']];

  var cache = CacheService.getDocumentCache();
  var cacheKey = 'unit_' + unitNameOrId + '_' + (section || 'all');
  try {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { }

  if (section) section = String(section).toLowerCase().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  // Resolve name -> ID via reference sheet
  var unitId = unitNameOrId;
  var refSheet = ss.getSheetByName('_unit_reference');
  if (refSheet) {
    var ref = refSheet.getDataRange().getValues();
    for (var i = 0; i < ref.length; i++) {
      if (ref[i][0] === unitNameOrId || String(ref[i][1]) === String(unitNameOrId)) {
        unitId = ref[i][1]; break;
      }
    }
  }

  var showAll = !section;

  function addSection(sheetName, label, typeLabel, idCol) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    var headers = data[0];
    var col = headers.indexOf(idCol || 'id');
    if (col === -1) col = headers.indexOf('unit_id');
    if (col === -1) col = 0;

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col]) === String(unitId)) rows.push(i);
    }
    if (rows.length === 0) return;

    if (results.length > 0) results.push(['', '', '']);
    results.push(['', label, '']);

    if (idCol === 'id' && label === 'units') {
      // Key-value for the units sheet
      var r = rows[0];
      headers.forEach(function(h, j) {
        if (h && data[r][j] !== undefined && data[r][j] !== '') {
          results.push(['Unit', h, String(data[r][j])]);
        }
      });
    } else {
      // Table for multi-row sheets
      var th = ['Type', 'Item'];
      headers.forEach(function(h) { if (h !== 'unit_id' && h !== 'unit_name') th.push(h); });
      results.push(th);

      var nameIdx = headers.indexOf('name');
      var typeIdx = headers.indexOf('type');

      rows.forEach(function(r) {
        var item;
        if (label === 'eco-driving / drive_rank') {
          item = (data[r][typeIdx] || '') + ': ' + (data[r][nameIdx] || 'Setting');
        } else {
          item = (nameIdx >= 0 ? data[r][nameIdx] : '') || typeLabel + ' ' + r;
        }
        var row = [typeLabel, item];
        headers.forEach(function(h, j) {
          if (h !== 'unit_id' && h !== 'unit_name') row.push(data[r][j] || '');
        });
        results.push(row);
      });
    }
  }

  if (showAll || section === 'unit' || section === 'units') addSection('units', 'units', 'Unit', 'id');
  if (showAll || section === 'sensor' || section === 'sensors') addSection('sensors', 'sensors', 'Sensor', 'unit_id');
  if (showAll || section === 'command' || section === 'commands') addSection('commands', 'commands', 'Command', 'unit_id');
  if (showAll || section === 'profile' || section === 'profiles') addSection('profiles', 'profiles', 'Profile', 'unit_id');
  if (showAll || section === 'custom' || section === 'custom_fields') addSection('custom_fields', 'custom_fields', 'Custom', 'unit_id');
  if (showAll || section === 'service' || section === 'service_intervals') addSection('service_intervals', 'service_intervals', 'Service', 'unit_id');
  if (showAll || section === 'hardware' || section === 'hw') addSection('hardware', 'hardware', 'Hardware', 'unit_id');
  if (showAll || section === 'drive_rank' || section === 'drive' || section === 'eco-driving' || section === 'eco' || section === 'rank')
    addSection('drive_rank', 'eco-driving / drive_rank', 'DriveRank', 'unit_id');

  var out = results.length > 0 ? results : [['No data found for this unit']];
  try { cache.put(cacheKey, JSON.stringify(out), 3600); } catch (e) { }
  return out;
}

// --- Filter sheet ---

// Flattens nested objects into [key, value] rows for display
function formatDataForSheet(data) {
  var result = [];

  function flatten(obj, prefix) {
    prefix = prefix || '';
    for (var key in obj) {
      if (!obj.hasOwnProperty(key)) continue;
      var path = prefix ? prefix + '.' + key : key;
      if (prefix && (key === 'unit_id' || key === 'unit_name')) continue;

      if (typeof obj[key] === 'string' && (obj[key].charAt(0) === '{' || obj[key].charAt(0) === '[')) {
        try {
          var parsed = JSON.parse(obj[key]);
          if (typeof parsed === 'object' && !Array.isArray(parsed)) { flatten(parsed, path); continue; }
          else { result.push([path, JSON.stringify(parsed)]); continue; }
        } catch (e) { }
      }

      if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
        flatten(obj[key], path);
      } else {
        var v = obj[key];
        if (v === null || v === undefined || v === '') continue;
        if (Array.isArray(v)) v = v.join(', ');
        else if (typeof v === 'object') v = JSON.stringify(v);
        else if (typeof v === 'boolean') v = v ? 'Yes' : 'No';

        var label = path.replace(/_/g, ' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
        if (label.indexOf('Settings ') >= 0) label = label.replace('Settings ', '');
        label = label.replace(/^(Unit |Hw |Sensor |Command |Profile )/, '');
        result.push([label, String(v)]);
      }
    }
  }

  flatten(data);
  return result.length > 0 ? result : [['No data', '']];
}

function setupFilterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName('Filter');
  if (!sheet) { sheet = ss.insertSheet('Filter'); } else { sheet.clear(); }

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  sheet.getRange(1, 1, 3, 4).setValues([
    ['UNIT LOOKUP', '', '', ''],
    ['', '', '', ''],
    ['Unit Name:', '', 'Status:', '']
  ]);

  sheet.getRange(1, 1, 1, 20).merge()
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#2c3e50').setFontColor('#ffffff').setHorizontalAlignment('center');

  sheet.getRange(3, 1).setFontWeight('bold').setBackground('#ecf0f1');
  sheet.getRange(3, 2).setBackground('#fff3cd')
    .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3, 3).setFontWeight('bold').setBackground('#ecf0f1');
  sheet.getRange(3, 4).setFormula(
    '=IF(B3="","Enter unit name",IF(ISERROR(QUERYUNIT(B3,"unit")),"Not found","Found"))');

  var sections = [
    { col: 1, title: 'BASIC INFO', q: 'unit' },
    { col: 4, title: 'HARDWARE', q: 'hardware' },
    { col: 7, title: 'SENSORS', q: 'sensors' },
    { col: 10, title: 'PROFILES', q: 'profiles' },
    { col: 13, title: 'DRIVE RANK', q: 'drive_rank' },
    { col: 16, title: 'CUSTOM FIELDS', q: 'custom_fields' },
    { col: 19, title: 'COMMANDS', q: 'commands' }
  ];

  sections.forEach(function(s) {
    sheet.getRange(5, s.col, 1, 2).merge().setValue(s.title)
      .setFontWeight('bold').setBackground('#4a90e2').setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#2c3e50', SpreadsheetApp.BorderStyle.SOLID_THICK);

    sheet.getRange(6, s.col).setValue('Property').setFontWeight('bold').setBackground('#e8f1f5')
      .setBorder(true, true, true, false, false, false, '#95a5a6', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, s.col + 1).setValue('Value').setFontWeight('bold').setBackground('#e8f1f5')
      .setBorder(true, false, true, true, false, false, '#95a5a6', SpreadsheetApp.BorderStyle.SOLID);

    for (var r = 7; r <= 40; r++) {
      sheet.getRange(r, s.col, 1, 2).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
    }
    sheet.getRange(7, s.col, 40, 2)
      .setBorder(false, true, true, true, true, true, '#dee2e6', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(7, s.col).setFormula(
      '=IF($B$3="","",IFERROR(QUERYUNIT($B$3,"' + s.q + '"),"No data"))');
  });

  for (var i = 0; i < 7; i++) {
    sheet.setColumnWidth(1 + i * 3, 180);
    sheet.setColumnWidth(2 + i * 3, 250);
    if (i < 6) sheet.setColumnWidth(3 + i * 3, 15);
  }

  sheet.getRange(3, 2).setNote(
    'Enter exact unit name to see all data.\nData appears in the tables below.');
  sheet.setFrozenRows(6);

  SpreadsheetApp.getUi().alert('Filter sheet ready. Enter a unit name in B3.');
}

// --- Test ---

function testSimple() {
  var ui = SpreadsheetApp.getUi();
  try {
    var session = authenticateWialon();
    if (!session) { ui.alert('Auth failed'); return; }

    var result = makeWialonRequest(session.eid, 'core/search_items', {
      spec: { itemsType: 'avl_unit', propName: 'sys_name', propValueMask: '*', sortType: 'sys_name' },
      force: 1, flags: 1, from: 0, to: 1
    });
    logoutWialon(session.eid);

    if (result && result.items && result.items.length > 0) {
      ui.alert('Connected! Found unit: ' + result.items[0].nm);
    } else {
      ui.alert('Connected but no units found');
    }
  } catch (error) {
    ui.alert('Error: ' + error.toString());
  }
}
