const doGet = () => {
  var page = HtmlService.createTemplateFromFile('index').evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(nameSystem)
    .setFaviconUrl(logoSystem)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return page;
}

const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
const nameSystem = settingSheet.getRange('B1').getDisplayValue();
const logoSystem = settingSheet.getRange('B2').getDisplayValue();
const userFolder = settingSheet.getRange('B3').getDisplayValue();
const baordCastFolder = settingSheet.getRange('B4').getDisplayValue();
const documentFolder = settingSheet.getRange('B5').getDisplayValue();
const carReqFolder = settingSheet.getRange('B6').getDisplayValue();
const sheetSetting = settingSheet.getRange('B7').getDisplayValue();
const sheetDocuMent = settingSheet.getRange('B8').getDisplayValue();
const sheetCarReq = settingSheet.getRange('B9').getDisplayValue();
const sheetCalendar = settingSheet.getRange('B16').getDisplayValue();
const sheetLeave = settingSheet.getRange('B17').getDisplayValue();
const sheetReport = settingSheet.getRange('B18').getDisplayValue();
const sheetProject = settingSheet.getRange('B19').getDisplayValue();
const sheetBroadCast = settingSheet.getRange('B20').getDisplayValue();
const sheetNotify = settingSheet.getRange('B21').getDisplayValue();
const sheetMeet = settingSheet.getRange('B22').getDisplayValue();
const meetFolder = settingSheet.getRange('B23').getDisplayValue();
const sheetHrm = settingSheet.getRange('B24').getDisplayValue();
const hrmFolder = settingSheet.getRange('B25').getDisplayValue();
const sheetLand = settingSheet.getRange('B26').getDisplayValue();
const landFolder = settingSheet.getRange('B27').getDisplayValue();
const sheetBin = settingSheet.getRange('B28').getDisplayValue();
const binFolder = settingSheet.getRange('B29').getDisplayValue();
const sheetEquip = settingSheet.getRange('B30').getDisplayValue();
const equipFolder = settingSheet.getRange('B31').getDisplayValue();
const sheetUseds = settingSheet.getRange('B32').getDisplayValue();
const sheetChat = settingSheet.getRange('B33').getDisplayValue();
const ChannelAccessToken = settingSheet.getRange('B34').getDisplayValue();
const LineMesAPI = settingSheet.getRange('B35').getDisplayValue();
const LinkWebApp = settingSheet.getRange('B36').getDisplayValue();

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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();

  const settings = {};

  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    settings[key] = value;
  });

  return settings;
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