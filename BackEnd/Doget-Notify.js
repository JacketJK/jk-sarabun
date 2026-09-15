const doGet = () => {
  var page = HtmlService.createTemplateFromFile('index').evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(nameSystem)
    .setFaviconUrl(logoSystem)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return page;
}

function _loadSystemSettings() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('SYSTEM_SETTINGS_ARRAY');
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length >= 8 && (parsed[0] || parsed[6] || parsed[7])) {
        return parsed;
      }
    }
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const sheet = ss.getSheetByName('Setting');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 1) {
          const numRows = Math.min(Math.max(lastRow, 1), 36);
          const values = sheet.getRange(1, 2, numRows, 1).getDisplayValues();
          const flat = values.map(function(r) { return r[0]; });
          while (flat.length < 36) flat.push('');
          if (flat[0] || flat[6] || flat[7]) {
            try {
              CacheService.getScriptCache().put('SYSTEM_SETTINGS_ARRAY', JSON.stringify(flat), 1800);
            } catch (e) {}
          }
          return flat;
        }
      }
    }
  } catch (err) {
    Logger.log('Error loading system settings: ' + err);
  }
  return new Array(36).fill('');
}

function clearSystemSettingsCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('SYSTEM_SETTINGS_ARRAY');
    cache.remove('SYSTEM_SETTINGS_MAP');
  } catch (e) {}
  _refreshGlobals();
}

var _systemSettings = _loadSystemSettings();
var nameSystem = _systemSettings[0] || '';
var logoSystem = _systemSettings[1] || '';
var userFolder = _systemSettings[2] || '';
var baordCastFolder = _systemSettings[3] || '';
var documentFolder = _systemSettings[4] || '';
var carReqFolder = _systemSettings[5] || '';
var sheetSetting = _systemSettings[6] || '';
var sheetDocuMent = _systemSettings[7] || '';
var sheetCarReq = _systemSettings[8] || '';
var sheetCalendar = _systemSettings[15] || '';
var sheetLeave = _systemSettings[16] || '';
var sheetReport = _systemSettings[17] || '';
var sheetProject = _systemSettings[18] || '';
var sheetBroadCast = _systemSettings[19] || '';
var sheetNotify = _systemSettings[20] || '';
var sheetMeet = _systemSettings[21] || '';
var meetFolder = _systemSettings[22] || '';
var sheetHrm = _systemSettings[23] || '';
var hrmFolder = _systemSettings[24] || '';
var sheetLand = _systemSettings[25] || '';
var landFolder = _systemSettings[26] || '';
var sheetBin = _systemSettings[27] || '';
var binFolder = _systemSettings[28] || '';
var sheetEquip = _systemSettings[29] || '';
var equipFolder = _systemSettings[30] || '';
var sheetUseds = _systemSettings[31] || '';
var sheetChat = _systemSettings[32] || '';
var ChannelAccessToken = _systemSettings[33] || '';
var LineMesAPI = _systemSettings[34] || '';
var LinkWebApp = _systemSettings[35] || '';

function _refreshGlobals() {
  _systemSettings = _loadSystemSettings();
  nameSystem = _systemSettings[0] || '';
  logoSystem = _systemSettings[1] || '';
  userFolder = _systemSettings[2] || '';
  baordCastFolder = _systemSettings[3] || '';
  documentFolder = _systemSettings[4] || '';
  carReqFolder = _systemSettings[5] || '';
  sheetSetting = _systemSettings[6] || '';
  sheetDocuMent = _systemSettings[7] || '';
  sheetCarReq = _systemSettings[8] || '';
  sheetCalendar = _systemSettings[15] || '';
  sheetLeave = _systemSettings[16] || '';
  sheetReport = _systemSettings[17] || '';
  sheetProject = _systemSettings[18] || '';
  sheetBroadCast = _systemSettings[19] || '';
  sheetNotify = _systemSettings[20] || '';
  sheetMeet = _systemSettings[21] || '';
  meetFolder = _systemSettings[22] || '';
  sheetHrm = _systemSettings[23] || '';
  hrmFolder = _systemSettings[24] || '';
  sheetLand = _systemSettings[25] || '';
  landFolder = _systemSettings[26] || '';
  sheetBin = _systemSettings[27] || '';
  binFolder = _systemSettings[28] || '';
  sheetEquip = _systemSettings[29] || '';
  equipFolder = _systemSettings[30] || '';
  sheetUseds = _systemSettings[31] || '';
  sheetChat = _systemSettings[32] || '';
  ChannelAccessToken = _systemSettings[33] || '';
  LineMesAPI = _systemSettings[34] || '';
  LinkWebApp = _systemSettings[35] || '';
}

const getLineNotifyToken = (code, state) => {
  try {
    const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName('User'); 
    const data = sheet.getDataRange().getValues();

    var requestBody = {
      grant_type: 'authorization_code',
      code: code,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: redirectUri
    };

    var options = {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: requestBody
    };

    var response = UrlFetchApp.fetch('https://notify-bot.line.me/oauth/token', options);
    var tokenData = JSON.parse(response.getContentText());

    if (tokenData.access_token) {
      for (let i = 1; i < data.length; i++) {
        const id = data[i][0];
        if (state === id) {
          sheet.getRange(i + 1, 15).setValue(tokenData.access_token);
          return tokenData;
        }
      }
    }

  } catch (error) {
    Logger.log('Error fetching LINE Notify token: ' + error);
    return { error: 'Error fetching token' };
  }
};

const addToken_linenotify = (obj) => {
    const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName('User'); 
    const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (obj.uuid === id) {
      sheet.getRange(i + 1, 15).setValue(obj.userLine);
      return obj.userLine;
    }
  }
}

const getdataSetting = () => {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('SYSTEM_SETTINGS_MAP');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
    const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();

    const settings = {};

    data.forEach(row => {
      const key = row[0];
      const value = row[1];
      settings[key] = value;
    });

    try {
      CacheService.getScriptCache().put('SYSTEM_SETTINGS_MAP', JSON.stringify(settings), 21600);
    } catch (e) {}

    return settings;
  } catch (err) {
    Logger.log('Error in getdataSetting: ' + err);
    return {};
  }
};

function sendClosingAnnouncement(userId, message, altText = "แจ้งเตือนสารบรรณอิเล็กทรอนิกส์") {
  const token = ChannelAccessToken;
  const url = 'https://api.line.me/v2/bot/message/push';

  let flexMessage = {
    "type": "flex",
    "altText": altText,
    "contents": {
      "type": "bubble",
      "direction": "ltr",
      "body": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [message]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "button",
            "action": {
              "type": "uri",
              "label": "เข้าสู่ระบบสารบรรณ",
              "uri": LinkWebApp
            },
            "style": "primary",
            "height": "sm"
          }
        ]
      },
      "styles": {
        "footer": {
          "separator": true
        }
      }
    }
  };

  const payload = {
    to: userId,
    messages: [flexMessage],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token,
    },
    payload: JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Announcement sent successfully: ' + response.getContentText());
  } catch (error) {
    Logger.log('Error sending announcement: ' + error.message);
  }
}

function sendImageOnly(userId, imageUrl) {
  const token = ChannelAccessToken;
  const url = 'https://api.line.me/v2/bot/message/push';

  if (!imageUrl) {
    Logger.log('No image URL provided.');
    return;
  }

  const payload = {
    to: userId,
    messages: [
      {
        type: 'image',
        originalContentUrl: imageUrl,
        previewImageUrl: imageUrl
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Image sent successfully: ' + response.getContentText());
  } catch (error) {
    Logger.log('Error sending image: ' + error.message);
  }
}