function formatToDateThai(date) {
  var thaiMonths = [
    'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
    'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
  ];
  var thaiMonth = thaiMonths[date.getMonth()];
  var thaiYear = date.getFullYear() + 543;

  var hours = date.getHours();
  var minutes = date.getMinutes();

  var thaiDate = date.getDate() + ' ' + thaiMonth + ' ' + thaiYear + ' เวลา ' + hours + ':' + (minutes < 10 ? '0' : '') + minutes + ' น.';

  return thaiDate;
}
function formatDateSave(date) {
  var formatYear = date.getFullYear();
  var formatMonth = (date.getMonth() + 1 < 10 ? '0' : '') + (date.getMonth() + 1);
  var formatDay = (date.getDate() < 10 ? '0' : '') + date.getDate();
  var formatHours = (date.getHours() < 10 ? '0' : '') + date.getHours();
  var formatMinutes = (date.getMinutes() < 10 ? '0' : '') + date.getMinutes();
  var formatSeconds = (date.getSeconds() < 10 ? '0' : '') + date.getSeconds();

  var formatDate = formatYear + '-' + formatMonth + '-' + formatDay;
  var formatTime = formatHours + ':' + formatMinutes + ':' + formatSeconds;
  
  return formatDate + ' ' + formatTime;
}
function getSetting(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting'); 
  var data = sheet.getRange("B1:B" + sheet.getLastRow()).getValues();
  return data;
} 
function selectSender() {
  var sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  var data = sheet.getDataRange().getValues().slice(1);

  data = data.filter(function(row) {
    return row[0] !== '';
  });

  data.sort(function(a, b) {
    return a[8] - b[8];
  });

  var selectedData = data.map(function(row) {
    return [row[0], row[4], row[13], row[14], row[6]];
  });

  return selectedData;
}

function getDataCheckboxUser() {
  var ss = SpreadsheetApp.openById(sheetUseds).getSheetByName('User'); 
  var data = ss.getDataRange().getDisplayValues().slice(1);
  return data
}

function checkLoginSystem(username, password, userIpAddress, userAgent) {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const data = sheet.getDataRange().getValues();

  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
      
  for (var i = 1; i < data.length; i++) { 
    if (data[i][1].toLowerCase() === username.toLowerCase() && data[i][2] === password) {
      if (data[i][15] === true) {
        var user = {
          publicKey: data[i][0],
          username: data[i][1],
          password: data[i][2],
          name_1: data[i][3],
          name_2: data[i][4],
          name_3: data[i][5],
          position: data[i][6],
          agency: data[i][7],
          level: data[i][8],
          tel: data[i][9],
          email: data[i][10],
          address: data[i][11],
          sig: data[i][12],
          profile: data[i][13],
          token: data[i][14],
          status: data[i][15],
          uidChat: data[i][16],
        };

        const logSheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("Log");
        logSheet.appendRow(["'" + user.username, userIpAddress, browserInfo + " " + osInfo, new Date(), "Login"]);

        return user;
      } else if (data[i][10] === false) {
        return '⚠️ ชื่อผู้ใช้งานนี้ถูกระงับการใช้งาน';
      }
    } 
  } 
  return '⚠️ ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง';
} 
function checkLogoutSystem(username, userIpAddress, userAgent) {
  const logSheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("Log");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  
  logSheet.appendRow(["'" + username, userIpAddress, browserInfo + " " + osInfo, new Date(), "Logout"]);
}
function checkActive(user) {
  var ss = SpreadsheetApp.openById(sheetUseds);
  var sheet = ss.getSheetByName('Log');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var fetchRows = Math.min(lastRow - 1, 300);
  var values = sheet.getRange(lastRow - fetchRows + 1, 1, fetchRows, sheet.getLastColumn()).getDisplayValues();
  var filteredData = [];

  for (var i = values.length - 1; i >= 0; i--) { 
    var dateValue = values[i][3]; 
    var userValue = values[i][0]; 
    if (dateValue !== "" && userValue === user) {
      filteredData.push(values[i]);
      if (filteredData.length >= 10) {
        break;
      }
    }
  }

  return filteredData;
}
function checkHistory(user) {
  var ss = SpreadsheetApp.openById(sheetUseds);
  var sheet = ss.getSheetByName('History');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var fetchRows = Math.min(lastRow - 1, 300);
  var values = sheet.getRange(lastRow - fetchRows + 1, 1, fetchRows, sheet.getLastColumn()).getDisplayValues();
  var filteredData = [];

  for (var i = values.length - 1; i >= 0; i--) { 
    var userValue = values[i][0]; 
    if (userValue === user) {
      filteredData.push(values[i]);
      if (filteredData.length >= 20) {
        break;
      }
    }
  }

  return filteredData;
}
function checkPosition(user) {
  var ss = SpreadsheetApp.openById(sheetUseds);
  var sheet = ss.getSheetByName('User');
  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  var filteredData = [];

  for (var i = values.length - 1; i >= 1; i--) { 
    var userValue = values[i][0]; 
    if (userValue === user) {
      filteredData.push(values[i]);
    }
  }

  return filteredData;
}
const getDataUser = () => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data;
}

const forgetsendOTP = (email) => {
  var otp = Math.floor(100000 + Math.random() * 900000);

  MailApp.sendEmail({
    to: email,
    subject: 'OTP Verification Reset Password I-OFFICE',
    body: 'Your OTP is: ' + otp,
  });

  return otp;
}
const confirmResetPassword = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) { 
    if (data[i][1] === obj.user) {
      sheet.getRange(i + 1, 3).setValue("'" + obj.password);
      break;
    }
  }
}

function getDataBaordcast() {
  const sheet = SpreadsheetApp.openById(sheetBroadCast).getSheetByName("Broadcast");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data;  
}
function changePw(obj) {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) { 
    var userValue = data[i][1];
    if (userValue === obj.userData1) {
      sheet.getRange(i + 1, 3).setValue("'" + obj.userData2);
      break; 
    }
  }
}

function checkedUser(obj) {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1 ; i < data.length; i++){
    const userValue = data[i][0];
    if (userValue === obj.userId) {
      sheet.getRange(i + 1, 16).setValue(obj.status);
      break; 
    }
  }
}
function deleteUser(obj) {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getDisplayValues();

  let rowID = data.findIndex(r => r[1] == obj.user) + 1;
  if (rowID > 0) {
    sheet.deleteRow(rowID);
  }
}
function addUserRegistor(obj) {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase");

  const publicKey = generateRandomKey();
  const partCode = send_otp_registor();

  var dataObj = ["'" + publicKey, "'" + obj.regictorData1, "'" + obj.regictorData2, obj.regictorData3, obj.regictorData4, obj.regictorData5, obj.regictorPosition, obj.regictorAgency, "'" + obj.regictorLevel, "'" + obj.regictorData6, obj.regictorData7, obj.regictorData8 + " หมู่ที่ " + obj.regictorData9 + " " + obj.regictorData10, , , , true, partCode]

  sheet.appendRow(dataObj);

  var dataHRM = ["'" + publicKey, "", obj.regictorData3, obj.regictorData4, obj.regictorData5, obj.regictorTypeposition, obj.regictorPosition, "", "", "", "ปฏิบัติราชการอยู่", formatToDateThai(new Date()), "", "", "", "", obj.regictorAgency]  

  sheetHRM.appendRow(dataHRM);

  return publicKey;
}
function addProfile(obj, ipAddress, userAgent) {
  const folder = DriveApp.getFolderById(userFolder);
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase");
  const userData = sheet.getDataRange().getValues();
  const hrmData = sheetHRM.getDataRange().getValues();
  var profileUrl = "";
  var number = "";

  if (obj.check !== "") {
    var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, obj.code);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    profileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
  } else {
    profileUrl = obj.profile;
  }

  for (let i = userData.length - 1; i >= 0; i--) {
    const publicKey = userData[i][0];
    if (publicKey === obj.code) {
      number = i;
      var dataHistory = [
        "'" + userData[i][1],
        userData[i][3] + userData[i][4] + " " + userData[i][5],
        ipAddress,
        browserInfo + " " + osInfo,
        "ลงทะเบียน",
        new Date()
      ];
      sheet.getRange(i + 1, 14).setValue(profileUrl);
      history.appendRow(dataHistory);
    }
  }

  for (let j = hrmData.length - 1; j >= 0; j--) {
    const publicKeyHRM = hrmData[j][0];
    if (publicKeyHRM === obj.code) {
      sheetHRM.getRange(j + 1, 10).setValue(profileUrl); 
      break; 
    }
  }

  return { data: userData[number], profileUrl: profileUrl };
}

function genCodeBaordcast(currentNumber) {
  var prefix = 'BC-';
  var paddingSize = 4;
  var number = currentNumber.toString();
  while (number.length < paddingSize) {
    number = '0' + number;
  }
  return prefix + number;
}
function notifyBaordcast(obj) {
  const folder = DriveApp.getFolderById(baordCastFolder);
  const sheet = SpreadsheetApp.openById(sheetBroadCast).getSheetByName("Broadcast");
  var lastRow = sheet.getLastRow();
  var codeBC = genCodeBaordcast(lastRow);
  var baordcast_Url = "";

  if (obj.check !== "") {
    var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, codeBC);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    baordcast_Url = "https://lh3.googleusercontent.com/d/" + fileId;
  } 

  var rowData = [codeBC, obj.text, baordcast_Url, , new Date()];
  sheet.appendRow(rowData);

  var alt = "ประกาศแจ้งเตือนข่าวประชาสัมพันธ์";

  if (obj.token.length > 0) {
    var row = [obj.token.join(', ')];
    sheet.getRange(lastRow + 1, 4, 1, row.length).setValues([row]);
  
    for (let i = 0; i < obj.token.length; i++) {
      const token = obj.token[i];
      if (baordcast_Url !== "") {
         var message = {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "text",
                "text": "แจ้งเตือนประกาศข่าวประชาสัมพันธ์",
                "weight": "bold",
                "color": "#1DB446",
                "size": "lg"
              },
              {
                "type": "text",
                "text": codeBC,
                "size": "xs",
                "color": "#aaaaaa",
                "wrap": true
              },
              {
                "type": "separator",
                "margin": "xxl"
              },
              {
                "type": "box",
                "layout": "vertical",
                "margin": "xxl",
                "spacing": "sm",
                "contents": [
                  {
                    "type": "box",
                    "layout": "horizontal",
                    "contents": [
                      {
                        "type": "text",
                        "text": obj.text,
                        "size": "sm",
                        "color": "#555555",
                        "flex": 0
                      }
                    ]
                  },
                ]
              }
            ]
          }

        sendClosingAnnouncement(token, message, alt);
        sendImageOnly(token, baordcast_Url)
      } else {
        var message = {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "text",
                "text": "แจ้งเตือนประกาศข่าวประชาสัมพันธ์",
                "weight": "bold",
                "color": "#1DB446",
                "size": "lg"
              },
              {
                "type": "text",
                "text": codeBC,
                "size": "xs",
                "color": "#aaaaaa",
                "wrap": true
              },
              {
                "type": "separator",
                "margin": "xxl"
              },
              {
                "type": "box",
                "layout": "vertical",
                "margin": "xxl",
                "spacing": "sm",
                "contents": [
                  {
                    "type": "box",
                    "layout": "horizontal",
                    "contents": [
                      {
                        "type": "text",
                        "text": obj.text,
                        "size": "sm",
                        "color": "#555555",
                        "flex": 0
                      }
                    ]
                  },
                ]
              }
            ]
          }

        sendClosingAnnouncement(token, message, alt);
      }
    }
  }
}
function editReportcast(obj) {
  const folder = DriveApp.getFolderById(baordCastFolder);
  const sheet = SpreadsheetApp.openById(sheetBroadCast).getSheetByName("Broadcast");
  const data = sheet.getDataRange().getValues();
  var baordcast_Url = "";

  for (let j = 1; j <= data.length; j++) {
    const codeValue = data[j][0];
    if (codeValue === obj.code) {
      sheet.getRange(j + 1, 2).setValue(obj.text);
      if (obj.check !== "") {
        var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
        var blob = Utilities.newBlob(datafile, obj.filetype, obj.code);
        var file = folder.createFile(blob);
        var fileId = file.getId();
        baordcast_Url = "https://lh3.googleusercontent.com/d/" + fileId;
        sheet.getRange(j + 1, 3).setValue(baordcast_Url);
      } 
      break;
    }
  }
}
function reSendNotify(obj) {
  var alt = "ประกาศแจ้งเตือนข่าวประชาสัมพันธ์";

  if (obj.token.length > 0) { 
    for (let i = 0; i < obj.token.length; i++) {
      const token = obj.token[i];

      var message = {
          "type": "box",
          "layout": "vertical",
          "contents": [
            {
              "type": "text",
              "text": "แจ้งเตือนประกาศข่าวประชาสัมพันธ์",
              "weight": "bold",
              "color": "#1DB446",
              "size": "lg"
            },
            {
              "type": "text",
              "text": codeBC,
              "size": "xs",
              "color": "#aaaaaa",
              "wrap": true
            },
            {
              "type": "separator",
              "margin": "xxl"
            },
            {
              "type": "box",
              "layout": "vertical",
              "margin": "xxl",
              "spacing": "sm",
              "contents": [
                {
                  "type": "box",
                  "layout": "horizontal",
                  "contents": [
                    {
                      "type": "text",
                      "text": obj.text,
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    }
                  ]
                },
              ]
            }
          ]
        }

        sendClosingAnnouncement(token, message, alt);
        sendImageOnly(token, obj.image)
    }
  }
}
function deleteSendNotify(obj) {
  const sheet = SpreadsheetApp.openById(sheetBroadCast).getSheetByName("Broadcast");
  const data = sheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(baordCastFolder);

  const files = folder.getFilesByName(obj.code);

    while (files.hasNext()) {
      var file = files.next();
      file.setTrashed(true);
    }

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === obj.code) {
      sheet.deleteRow(i + 1);
    }
  }  
}
function testSendNotify(token) {
  if (token !== "") {
    var alt = "ทดสอบประกาศแจ้งเตือนข่าวประชาสัมพันธ์";
    var message = {
          "type": "box",
          "layout": "vertical",
          "contents": [
            {
              "type": "text",
              "text": "ทดสอบแจ้งเตือนประกาศข่าวประชาสัมพันธ์",
              "weight": "bold",
              "color": "#1DB446",
              "size": "lg"
            },
            {
              "type": "text",
              "text": "สำหรับทดสอบ",
              "size": "xs",
              "color": "#aaaaaa",
              "wrap": true
            },
            {
              "type": "separator",
              "margin": "xxl"
            },
            {
              "type": "box",
              "layout": "vertical",
              "margin": "xxl",
              "spacing": "sm",
              "contents": [
                {
                  "type": "box",
                  "layout": "horizontal",
                  "contents": [
                    {
                      "type": "text",
                      "text": "🔔สำหรับทดสอบการส่งแจ้งเตือน🔔",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    }
                  ]
                },
              ]
            }
          ]
        }
        
    sendClosingAnnouncement(token, message, alt);

    return '🔔 ทดสอบส่งข้อความแจ้งเตือนสำเร็จ';
  } else {
    return '⚠️ ทดสอบส่งข้อความแจ้งเตือนไม่สำเร็จ';
  }
}
function clearDropdownCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_DROPDOWN_DATA');
  } catch (e) {}
}

function getAllDropdownData() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('ALL_DROPDOWN_DATA');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.openById(sheetSetting);

    function getSheetData(sheetName, colLimit) {
      try {
        const sh = ss.getSheetByName(sheetName);
        if (!sh) return [];
        const lastRow = sh.getLastRow();
        if (lastRow < 3) return [];
        const data = sh.getRange('A3:' + colLimit + lastRow).getValues();
        return data.filter(function(row) {
          return row[0] !== '';
        });
      } catch (err) {
        return [];
      }
    }

    const result = {
      agency: getSheetData("Agency", "C"),
      institute: getSheetData("Institute", "C"),
      department: getSheetData("DepartMent", "F"),
      position: getSheetData("Position", "C"),
      classSpeed: getSheetData("ClassSpeed", "C"),
      classSecret: getSheetData("ClassSecret", "C"),
      docCategory: getSheetData("DocCategory", "C"),
      objective: getSheetData("Objective", "C"),
      responseStatus: getSheetData("Response", "C"),
      jobWord: getSheetData("JobWord", "C"),
      typeEmployee: getSheetData("TypeEmployee", "C"),
      mooBan: getSheetData("MooBan", "C"),
      equipCategory: getSheetData("EquipCategory", "C"),
      equipType: getSheetData("EquipType", "C"),
      equipStatus: getSheetData("EquipStatus", "C"),
      equipAcquire: getSheetData("EquipAcquire", "C"),
      equipExpense: getSheetData("EquipExpense", "C"),
      equipVendor: getSheetData("EquipVendor", "C"),
      matCategory: getSheetData("MatCategory", "C"),
      matType: getSheetData("MatType", "C"),
      equipUnit: getSheetData("EquipUnit", "C"),
      equipLocation: getSheetData("EquipLocation", "C")
    };

    try {
      const jsonStr = JSON.stringify(result);
      if (jsonStr.length < 100000) {
        CacheService.getScriptCache().put('ALL_DROPDOWN_DATA', jsonStr, 21600); // 6 hours
      }
    } catch (e) {}

    return result;
  } catch (err) {
    Logger.log('Error in getAllDropdownData: ' + err);
    return {};
  }
}

function selectAgency() {
  const all = getAllDropdownData();
  if (all && all.agency && all.agency.length > 0) return all.agency;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Agency"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function getInstitute() {
  const all = getAllDropdownData();
  if (all && all.institute && all.institute.length > 0) return all.institute;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Institute"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectDepartment() {
  const all = getAllDropdownData();
  if (all && all.department && all.department.length > 0) return all.department;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DepartMent"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:F' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectPosition() {
  const all = getAllDropdownData();
  if (all && all.position && all.position.length > 0) return all.position;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Position"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectClassSpeed() {
  const all = getAllDropdownData();
  if (all && all.classSpeed && all.classSpeed.length > 0) return all.classSpeed;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSpeed"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectClassSecret() {
  const all = getAllDropdownData();
  if (all && all.classSecret && all.classSecret.length > 0) return all.classSecret;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSecret"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectDocCategory() {
  const all = getAllDropdownData();
  if (all && all.docCategory && all.docCategory.length > 0) return all.docCategory;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DocCategory"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectObjective() {
  const all = getAllDropdownData();
  if (all && all.objective && all.objective.length > 0) return all.objective;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Objective"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectResponseStatus() {
  const all = getAllDropdownData();
  if (all && all.responseStatus && all.responseStatus.length > 0) return all.responseStatus;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Response"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectJobWord() {
  const all = getAllDropdownData();
  if (all && all.jobWord && all.jobWord.length > 0) return all.jobWord;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("JobWord"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectTypeEmployee() {
  const all = getAllDropdownData();
  if (all && all.typeEmployee && all.typeEmployee.length > 0) return all.typeEmployee;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("TypeEmployee"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function selectMooBan() {
  const all = getAllDropdownData();
  if (all && all.mooBan && all.mooBan.length > 0) return all.mooBan;
  var sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("MooBan"); 
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange('A3:C' + lastRow).getValues().filter(function(row) { return row[0] !== ''; });
}
function formatDate(date) {
  const year = date.getFullYear().toString().slice(-2); 
  const month = ('0' + (date.getMonth() + 1)).slice(-2); 
  const day = ('0' + date.getDate()).slice(-2); 
  const hours = ('0' + date.getHours()).slice(-2); 
  const minutes = ('0' + date.getMinutes()).slice(-2); 
  const seconds = ('0' + date.getSeconds()).slice(-2); 
  return year + month + day + hours + minutes + seconds; 
}
function generateRandomKey() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp; 
  for (let i = 0; i < 15; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key; 
}
function gsAddnewdoc(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const documentKey = generateRandomKey();

  const rowData = [documentKey, obj.addNewDoc1, obj.addNewDoc2, obj.addNewDoc3, obj.addNewDoc4, obj.addNewDoc5, obj.addNewDoc6, obj.addNewDoc7, obj.addNewDoc8, "'" + obj.addNewDoc9, obj.addNewDoc10, obj.addNewDoc11, obj.addNewDoc12, obj.addNewDoc13, obj.addNewDoc14, obj.addNewDoc15, obj.addNewDoc16, ];

  sheet.appendRow(rowData);

  const setSheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DepartMent"); 
  const data = setSheet.getRange('C3:F' + setSheet.getLastRow()).getValues();
  if (obj.addNewDoc2 !== "-") {
    for (let i = 0; i < data.length; i++) {
      const category = data[i][0];
      const dataNumber = parseInt(data[i][3]) + 1;
      if (obj.addNewDoc4 === category) {
        setSheet.getRange(i + 3, 6).setValue(dataNumber);
      }
      break;
    }
  }

  return documentKey;
}
const editGsAddnewdoc = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 1, 4).setValue(obj.addNewDoc3);
      sheet.getRange(i + 1, 5).setValue(obj.addNewDoc4);
      sheet.getRange(i + 1, 6).setValue(obj.addNewDoc5);
      sheet.getRange(i + 1, 7).setValue(obj.addNewDoc6);
      sheet.getRange(i + 1, 8).setValue(obj.addNewDoc7);
      sheet.getRange(i + 1, 9).setValue(obj.addNewDoc8);
      sheet.getRange(i + 1, 10).setValue("'" + obj.addNewDoc9);
      sheet.getRange(i + 1, 11).setValue(obj.addNewDoc10);
      sheet.getRange(i + 1, 12).setValue(obj.addNewDoc11);
      sheet.getRange(i + 1, 13).setValue(obj.addNewDoc12);
      sheet.getRange(i + 1, 14).setValue(obj.addNewDoc13);
      sheet.getRange(i + 1, 15).setValue(obj.addNewDoc14);
      sheet.getRange(i + 1, 16).setValue(obj.addNewDoc15);
      sheet.getRange(i + 1, 17).setValue(obj.addNewDoc16);
      sheet.getRange(i + 1, 26).setValue(formatToDateThai(new Date()));

      const updatedData = sheet.getRange(i + 1, 1, 1, 23).getValues();
      
      return updatedData[0];
    }
  }
}
function gsAddnewdocOut(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const documentKey = generateRandomKey();
  const lastRow = sheet.getLastRow() + 1;

  const rowData = [documentKey, obj.addNewDocOut1, obj.addNewDocOut2, "-", obj.addNewDocOut3, "-", obj.addNewDocOut4, obj.addNewDocOut5, obj.addNewDocOut6, "'" + obj.addNewDocOut7, obj.addNewDocOut8, obj.addNewDocOut9, obj.addNewDocOut10, obj.addNewDocOut11, obj.addNewDocOut12, obj.addNewDocOut13, obj.addNewDocOut14];

  sheet.appendRow(rowData);
  sheet.getRange(lastRow, 35).setValue(obj.addNewDocOut15);

  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting"); 
  const dataOut = setSheet.getRange('B10').getValue();
  const dataIn = setSheet.getRange('B11').getValue();
  if (obj.addNewDocOut1 === "ลงรับภายนอก") {
    setSheet.getRange('B10').setValue(parseInt(dataOut) + 1);
  } else {
    setSheet.getRange('B11').setValue(parseInt(dataIn) + 1);
  }

  return documentKey;
}
function gsAddnewDocument(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const documentKey = generateRandomKey();
  const lastRow = sheet.getLastRow() + 1;

  const rowData = [documentKey, obj.addNewDocument1, obj.addNewDocument2, "-", obj.addNewDocument3, "-", obj.addNewDocument4, obj.addNewDocument5, obj.addNewDocument6, "'" + obj.addNewDocument7, obj.addNewDocument8, obj.addNewDocument9, obj.addNewDocument10, obj.addNewDocument11, obj.addNewDocument12, obj.addNewDocument13, obj.addNewDocument14];

  sheet.appendRow(rowData);
  sheet.getRange(lastRow, 35).setValue("-");

  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting"); 
  const data = setSheet.getRange('A12:B15').getValues();
    for (let i = 0; i <= data.length; i++) {
      if (obj.addNewDocument1 === data[i][0]) {
        setSheet.getRange(i + 12, 2).setValue(parseInt(data[i][1]) + 1);
        break;
      }
    }

  return documentKey;
}
function _sanitizeValues(rows) {
  if (!rows || !Array.isArray(rows)) return [];
  return rows.map(function(row) {
    if (!Array.isArray(row)) return [];
    return row.map(function(cell) {
      if (cell === null || cell === undefined) return '';
      if (cell instanceof Date) {
        if (isNaN(cell.getTime())) return '';
        try {
          return formatToDateThai(cell);
        } catch (e) {
          return cell.toISOString();
        }
      }
      return String(cell);
    });
  });
}

function dataDocAlldetail() {
  try {
    const ss = _getDocuMentSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName("DataDocument"); 
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 2) return [];
    return _sanitizeValues(data.slice(2));
  } catch (err) {
    Logger.log("Error in dataDocAlldetail: " + err);
    return [];
  }
}
function dataDocPending(key) {
  try {
    const ss = _getDocuMentSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName("DataDocument"); 
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues().slice(2);
    const cleanKey = String(key || '').replace(/^'/, '').trim();

    data = data.filter((row) => {
      const rowKey = String(row[21] || '').replace(/^'/, '').trim();
      const status = String(row[22] || '').trim();
      return (rowKey === cleanKey && (status === '1' || status === '2' || status === '')) || status === '2';
    });
    return _sanitizeValues(data);
  } catch (err) {
    Logger.log("Error in dataDocPending: " + err);
    return [];
  }
}
function dataDocFollow(key) {
  try {
    const ss = _getDocuMentSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName("DataDocument"); 
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues().slice(2);
    const cleanKey = String(key || '').replace(/^'/, '').trim();

    data = data.filter((row) => {
      const rowKey = String(row[21] || '').replace(/^'/, '').trim();
      const follow = String(row[15] || '').trim();
      return rowKey === cleanKey && follow === '1';
    });
    return _sanitizeValues(data);
  } catch (err) {
    Logger.log("Error in dataDocFollow: " + err);
    return [];
  }
}
function dataDocReturn(key) {
  try {
    const ss = _getDocuMentSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName("DataDocument"); 
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues().slice(2);
    const cleanKey = String(key || '').replace(/^'/, '').trim();

    data = data.filter((row) => {
      const rowKey = String(row[21] || '').replace(/^'/, '').trim();
      const status = String(row[22] || '').trim();
      return rowKey === cleanKey && status === '0';
    });
    return _sanitizeValues(data);
  } catch (err) {
    Logger.log("Error in dataDocReturn: " + err);
    return [];
  }
}
function _getDocuMentSS() {
  if (!sheetDocuMent) _refreshGlobals();
  if (sheetDocuMent) {
    try { return SpreadsheetApp.openById(sheetDocuMent); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _getUserSS() {
  if (!sheetUseds) _refreshGlobals();
  if (sheetUseds) {
    try { return SpreadsheetApp.openById(sheetUseds); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _getSettingSS() {
  if (!sheetSetting) _refreshGlobals();
  if (sheetSetting) {
    try { return SpreadsheetApp.openById(sheetSetting); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function dataDocInside(key) {
  try {
    const ss = _getDocuMentSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName("DataDocument"); 
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues().slice(2);
    const cleanKey = String(key || '').replace(/^'/, '').trim();

    data = data.filter((row) => {
      const rowKey = String(row[21] || '').replace(/^'/, '').trim();
      return !cleanKey || rowKey === cleanKey;
    });
    return _sanitizeValues(data);
  } catch (err) {
    Logger.log("Error in dataDocInside: " + err);
    return [];
  }
}
function userActionSendDoc(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const userValue = data[i][0];
    if (userValue === obj.key) {
      sheet.getRange(i + 1, 18).setValue(obj.name);
      sheet.getRange(i + 1, 19).setValue(obj.position);
      sheet.getRange(i + 1, 20).setValue(obj.agency);
      sheet.getRange(i + 1, 21).setValue(obj.token);
      sheet.getRange(i + 1, 22).setValue(obj.publicKey);
      sheet.getRange(i + 1, 23).setValue('1');
      sheet.getRange(i + 1, 24).setValue(true);
      sheet.getRange(i + 1, 25).setValue(formatToDateThai(new Date()));
      sheet.getRange(i + 1, 26).setValue(formatToDateThai(new Date()));

      const updatedData = sheet.getRange(i + 1, 1, 1, 23).getValues();
      
      return updatedData[0];
    }
  }

  return null;
}
function saveFile(e, fileNumber) {
  let folder = DriveApp.getFolderById(documentFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  fileUrl = folder.createFile(blob).getId();
  var url = "https://drive.google.com/open?id=" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === e.code) {
      sheet.getRange(i + 1, fileNumber).setValue(url);
      sheet.getRange(i + 1, 23).setValue("2");
      sheet.getRange(i + 1, 26).setValue(formatToDateThai(new Date()));
      break;
    }
  }
  
  return 'อัปโหลดไฟล์เรียบร้อย';
}
function deleteFileAtDrive(code, number) {
  const folder = DriveApp.getFolderById(documentFolder); 
  const files = folder.getFilesByName(code + "_" + number);

    while (files.hasNext()) {
      var file = files.next();
      file.setTrashed(true);
    }

  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  var inumber = "";
  if (number === 1) {
    inumber = 27;
  } else if (number === 2) {
    inumber = 28;
  } else if (number === 3) {
    inumber = 29;
  } else if (number === 4) {
    inumber = 30;
  } else if (number === 5) {
    inumber = 31;
  }

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === code) {
      sheet.getRange(i + 1, inumber).clearContent();
    }
  }

  return "🚫 ลบไฟล์สำเร็จ";
}
function sendDataToUser(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === obj.code) {
      sheet.getRange(i + 1, 32).setValue(obj.link1);
      sheet.getRange(i + 1, 32).setValue(obj.link1);
      sheet.getRange(i + 1, 33).setValue(obj.link2);
      sheet.getRange(i + 1, 23).setValue("2");
      sheet.getRange(i + 1, 26).setValue(formatToDateThai(new Date()));
      return data[i][21];
    }
  }
}
function toggleFollowDoc(key) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === key) {
      if (data[i][15] == 1) {
        sheet.getRange(i + 1, 16).setValue("");
        return 0;
      } else {
        sheet.getRange(i + 1, 16).setValue("1");
        return 1;
      }
    }
  }
}
function toggleFollowDocUser(key) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === key) {
      if (data[i][9] == 1) {
        sheet.getRange(i + 1, 10).setValue("");
        return 0;
      } else {
        sheet.getRange(i + 1, 10).setValue("1");
        return 1;
      }
    }
  }
}
function sendSuccessDoc(obj) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
  const sheetSend = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const data = sheet.getDataRange().getValues();
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const dataUser = sheetUser.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === obj.code) {
      obj.keys.forEach(function(key) {
        const docSendKey = generateRandomKey();
        sheetSend.appendRow([docSendKey, obj.code, key, '0', formatToDateThai(new Date()), 'รอตอบรับ']);
        var genCode = generateRandomKey();
        notiSheet.appendRow([genCode, key, data[i][21], "ส่งหนังสือเลขที่ " + data[i][2], "เรื่อง " + data[i][7], "docdetailfile", obj.code, "0", "'" + formatDateSave(new Date())]);
      });

      const tokensRow = obj.tokens.length > 1 ? [obj.tokens.join(', ')] : [obj.tokens];
      sheet.getRange(i + 1, 36).setValues([tokensRow]);
      sheet.getRange(i + 1, 23).setValue("3");
      sheet.getRange(i + 1, 26).setValue(formatToDateThai(new Date()));

      var alt = "แจ้งเตือนมีหนังสือส่งถึงท่าน";

      var Message = {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "แจ้งเตือนการส่งหนังสือ",
            "weight": "bold",
            "color": "#1DB446",
            "size": "lg"
          },
          {
            "type": "text",
            "text": obj.code,
            "size": "xs",
            "color": "#aaaaaa",
            "wrap": true
          },
          {
            "type": "separator",
            "margin": "xxl"
          },
          {
            "type": "box",
            "layout": "vertical",
            "margin": "xxl",
            "spacing": "sm",
            "contents": [
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  { "type": "text","text": "ประเภทหนังสือ","size": "sm","color": "#555555","flex": 0 },
                  { "type": "text","text": data[i][1],"size": "sm","color": "#111111","align": "end" }
                ]
              },
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  { "type": "text","text": "เลขที่หนังสือ","size": "sm","color": "#555555","flex": 0 },
                  { "type": "text","text": data[i][2],"size": "sm","color": "#111111","align": "end" }
                ]
              },
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  { "type": "text","text": "เรื่อง","size": "sm","color": "#555555","flex": 0 },
                  { "type": "text","text": data[i][7],"size": "sm","color": "#111111","align": "end" }
                ]
              }
            ]
          }
        ]
      };

      let allTokens = [];
      for (let i = 0; i < tokensRow.length; i++) {
        const tokenSet = String(tokensRow[i]).split(",");
        for (let j = 0; j < tokenSet.length; j++) {
          const trimmed = String(tokenSet[j]).trim();
          if (trimmed !== "") {
            allTokens.push(trimmed);
          }
        }
      }

      const uniqueTokens = [...new Set(allTokens)];

      for (let k = 0; k < uniqueTokens.length; k++) {
        const token = uniqueTokens[k];
        sendClosingAnnouncement(token, Message, alt);
        Utilities.sleep(1000); 
      }

      obj.keys.forEach(function(key) {
        for (let u = 1; u < dataUser.length; u++) {
          if (dataUser[u][0] === key) {   // คอลัมน์ที่ 1 ของ User
            const email = dataUser[u][10]; // คอลัมน์ K (index 10)
            if (email && email !== "") {
              const subject = "แจ้งเตือนการส่งหนังสือ เลขที่ " + data[i][2];
              
              let fileLinks = "";
              for (let col = 26; col <= 30; col++) {
                const fileLink = data[i][col];
                if (fileLink) {
                  fileLinks += `<li><a href="${fileLink}" target="_blank">เอกสารแนบ ${col - 25}</a></li>`;
                }
              }

              const bodyHtml = `
                <p>เรียนคุณ ${data[i][17]}</p>
                <p>มีหนังสือส่งถึงท่าน</p>
                <ul>
                  <li><b>เลขที่:</b> ${data[i][2]}</li>
                  <li><b>เรื่อง:</b> ${data[i][7]}</li>
                  <li><b>วันที่:</b> ${new Date(data[i][9]).toLocaleDateString('th-TH', {
                      year: 'numeric',
                      month: 'long',
                      day: 'numeric',
                    })}</li>
                </ul>
                <p><b>ไฟล์แนบ:</b></p>
                <ul>
                  ${fileLinks || "<li>ไม่มีไฟล์แนบ</li>"}
                </ul>
              `;

              MailApp.sendEmail({
                to: email,
                subject: subject,
                htmlBody: bodyHtml
              });
            }
          }
        }
      });



      return data[i][21];
    }
  }
}


function openDocument(key) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const userValue = data[i][0];
    if (userValue === key) {
      if (data[i][3] < 1) {
        sheet.getRange(i + 1, 4).setValue("1");
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
      }
      return data[i][2];
    }
  }
}
function dataDocInprogress(key) {
  try {
    const ssDoc = _getDocuMentSS();
    if (!ssDoc) return [[], []];
    const sheet = ssDoc.getSheetByName("SendDocument");
    const sheetDoc = ssDoc.getSheetByName("DataDocument");
    if (!sheet || !sheetDoc) return [[], []];

    const data = sheet.getDataRange().getValues().slice(1);
    const dataDoc = sheetDoc.getDataRange().getValues().slice(2);
    const cleanKey = String(key || '').replace(/^'/, '').trim();

    const filteredSend = [];
    const filteredDoc = [];

    // Map docId (column A, index 0) -> docRow
    const docMap = new Map();
    for (let d = 0; d < dataDoc.length; d++) {
      const docRow = dataDoc[d];
      if (docRow[0]) {
        docMap.set(String(docRow[0]).replace(/^'/, '').trim(), docRow);
      }
    }

    for (let s = 0; s < data.length; s++) {
      const sendRow = data[s];
      const sendKey = String(sendRow[2] || '').replace(/^'/, '').trim();
      if (!cleanKey || sendKey === cleanKey) {
        const docId = String(sendRow[1] || '').replace(/^'/, '').trim();
        const docRow = docMap.get(docId);
        if (docRow) {
          const docStatus = Number(docRow[22]);
          if (isNaN(docStatus) || docStatus < 5) {
            filteredSend.push(sendRow);
            filteredDoc.push(docRow);
          }
        }
      }
    }

    filteredSend.reverse();
    filteredDoc.reverse();

    return [_sanitizeValues(filteredSend), _sanitizeValues(filteredDoc)];
  } catch (err) {
    Logger.log("Error in dataDocInprogress: " + err);
    return [[], []];
  }
}
function dataDocInprogressOut(key) {
  try {
    const ssDoc = _getDocuMentSS();
    const ssUser = _getUserSS();
    if (!ssDoc || !ssUser) return [[], [], []];
    const sheet = ssDoc.getSheetByName("SendDocument");
    const sheetDoc = ssDoc.getSheetByName("DataDocument");
    const sheetUser = ssUser.getSheetByName("User");
    if (!sheet || !sheetDoc || !sheetUser) return [[], [], []];

    const lastRowSend = sheet.getLastRow();
    const lastRowDoc = sheetDoc.getLastRow();
    const lastRowUser = sheetUser.getLastRow();
    if (lastRowDoc < 2) return [[], [], []];

    const data = lastRowSend >= 2 ? sheet.getRange(2, 1, lastRowSend - 1, sheet.getLastColumn()).getDisplayValues() : [];
    const dataDoc = sheetDoc.getRange(2, 1, lastRowDoc - 1, sheetDoc.getLastColumn()).getDisplayValues();
    const dataUser = lastRowUser >= 2 ? sheetUser.getRange(2, 1, lastRowUser - 1, sheetUser.getLastColumn()).getDisplayValues() : [];

    const userMap = new Map();
    for (let u = 0; u < dataUser.length; u++) {
      userMap.set(dataUser[u][0], dataUser[u]);
    }

    const filteredData = dataDoc.filter(row => row[21] === key && (row[1] === "ลงรับภายนอก" || row[1] === "ลงรับภายใน"));
    const filteredDataDoc = [];
    const filteredDataUser = [];

    for (const row of filteredData) {
      const keyCode = row[0];
      const matchingRows = data.filter(docRow => docRow[1] === keyCode);
      filteredDataDoc.push(...matchingRows);
      for (const rowUser of matchingRows) {
        const userCode = rowUser[2];
        const matchingUser = userMap.get(userCode);
        if (matchingUser) {
          filteredDataUser.push(matchingUser);
        }
      }
    }
    
    return [filteredData, filteredDataDoc, filteredDataUser];
  } catch (err) {
    Logger.log("Error in dataDocInprogressOut: " + err);
    return [[], [], []];
  }
}
function dataDocuMentSave(key) {
  try {
    const ssDoc = SpreadsheetApp.openById(sheetDocuMent);
    const sheet = ssDoc.getSheetByName("SendDocument");
    const sheetDoc = ssDoc.getSheetByName("DataDocument");
    const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
    if (!sheet || !sheetDoc || !sheetUser) return [[], [], []];

    const lastRowSend = sheet.getLastRow();
    const lastRowDoc = sheetDoc.getLastRow();
    const lastRowUser = sheetUser.getLastRow();
    if (lastRowDoc < 2) return [[], [], []];

    const data = lastRowSend >= 2 ? sheet.getRange(2, 1, lastRowSend - 1, sheet.getLastColumn()).getDisplayValues() : [];
    const dataDoc = sheetDoc.getRange(2, 1, lastRowDoc - 1, sheetDoc.getLastColumn()).getDisplayValues();
    const dataUser = lastRowUser >= 2 ? sheetUser.getRange(2, 1, lastRowUser - 1, sheetUser.getLastColumn()).getDisplayValues() : [];

    const userMap = new Map();
    for (let u = 0; u < dataUser.length; u++) {
      userMap.set(dataUser[u][0], dataUser[u]);
    }

    const filteredDataDoc = dataDoc.filter(row => row[0] === key);
    const filteredData = [];
    const filteredStatus = [];

    for (const row of filteredDataDoc) {
      const keyCode = row[0];
      const matchingRows = data.filter(docRow => docRow[1] === keyCode);
      filteredStatus.push(...matchingRows);
      for (const rowUser of matchingRows) {
        const userCode = rowUser[2];
        const matchingUser = userMap.get(userCode) || [];
        filteredData.push(matchingUser);
      }
    }
    
    return [filteredDataDoc, filteredStatus, filteredData];
  } catch (err) {
    Logger.log("Error in dataDocuMentSave: " + err);
    return [[], [], []];
  }
}
function saveWorkingData(obj) {
  try {
    const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
    const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
    if (!sheet || !sheetDoc) return null;

    const data = sheet.getDataRange().getValues();
    const lastRowDoc = sheetDoc.getLastRow();
    const dataDoc = lastRowDoc >= 3 ? sheetDoc.getRange(3, 1, lastRowDoc - 2, 1).getValues() : [];

    for (let i = data.length - 1; i >= 0; i--) {
      const keyValue = data[i][0];
      const keyDoc = data[i][1];

      if (keyValue === obj.key) {
        sheet.getRange(i + 1, 4).setValue("3");
        sheet.getRange(i + 1, 6).setValue(obj.response);
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 8).setValue(obj.jobWord);
        sheet.getRange(i + 1, 9).setValue(obj.replace);

        for (let j = dataDoc.length - 1; j >= 0; j--) {
          const keyValueDoc = dataDoc[j][0];
          if (keyValueDoc === keyDoc) {
            sheetDoc.getRange(j + 3, 23).setValue("4");
            sheetDoc.getRange(j + 3, 26).setValue(formatToDateThai(new Date()));
            break;
          }
        }
        return data[i][2];
      }
    }
  } catch (err) {
    Logger.log("Error in saveWorkingData: " + err);
    return null;
  }
}
function closeJobDocument(key) {
  try {
    const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
    if (!sheetDoc) return;
    const lastRowDoc = sheetDoc.getLastRow();
    const dataDoc = lastRowDoc >= 3 ? sheetDoc.getRange(3, 1, lastRowDoc - 2, 1).getValues() : [];

    for (let i = dataDoc.length - 1; i >= 0; i--) {
      const keyDoc = dataDoc[i][0];
      if (keyDoc === key) {
        sheetDoc.getRange(i + 3, 23).setValue("5");
        sheetDoc.getRange(i + 3, 26).setValue(formatToDateThai(new Date()));
        break;
      }
    }
  } catch (err) {
    Logger.log("Error in closeJobDocument: " + err);
  }
}
function returnJobDocument(obj) {
  try {
    const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
    const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
    if (!sheetDoc || !sheet) return null;

    const lastRowDoc = sheetDoc.getLastRow();
    const dataDoc = lastRowDoc >= 3 ? sheetDoc.getRange(3, 1, lastRowDoc - 2, 1).getValues() : [];
    const data = sheet.getDataRange().getValues();

    for (let j = dataDoc.length - 1; j >= 0; j--) {
      const keyDoc = dataDoc[j][0];
      if (keyDoc === obj.keyDocument) {
        sheetDoc.getRange(j + 3, 23).setValue("0");
        sheetDoc.getRange(j + 3, 26).setValue(formatToDateThai(new Date()));
      }
    }
    for (let i = data.length - 1; i >= 0; i--) {
      const keyValue = data[i][0];
      if (keyValue === obj.codeDocument) {
        sheet.getRange(i + 1, 4).setValue('2');
        sheet.getRange(i + 1, 6).setValue(obj.noteDocument);
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        return data[i][2];
      }
    }
  } catch (err) {
    Logger.log("Error in returnJobDocument: " + err);
    return null;
  }
}
function cancelDocument(obj) {
  try {
    const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
    if (!sheetDoc) return null;
    const lastRowDoc = sheetDoc.getLastRow();
    const dataDoc = lastRowDoc >= 3 ? sheetDoc.getRange(3, 1, lastRowDoc - 2, 1).getValues() : [];
    const targetKey = obj.key || obj.code || obj.keyDocument;
    let docNumber = "";

    for (let i = dataDoc.length - 1; i >= 0; i--) {
      const keyDoc = dataDoc[i][0];
      if (keyDoc === targetKey) {
        const rowIndex = i + 3;
        sheetDoc.getRange(rowIndex, 23).setValue("-1"); // สถานะยกเลิก
        sheetDoc.getRange(rowIndex, 26).setValue(formatToDateThai(new Date()));
        
        if (obj.note) {
          sheetDoc.getRange(rowIndex, 35).setValue(obj.note);
        }
        
        const curNote = String(sheetDoc.getRange(rowIndex, 15).getValue() || "").trim();
        const cancelTag = obj.note ? `[ยกเลิก: ${obj.note}]` : "[ยกเลิกหนังสือ]";
        if (!curNote.includes("[ยกเลิก")) {
          sheetDoc.getRange(rowIndex, 15).setValue(curNote ? `${curNote} ${cancelTag}` : cancelTag);
        }

        docNumber = sheetDoc.getRange(rowIndex, 3).getValue();
        break;
      }
    }

    // อัปเดตในตาราง SendDocument (ถ้ามี)
    try {
      const sheetSend = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
      if (sheetSend) {
        const dataSend = sheetSend.getDataRange().getValues();
        for (let s = 1; s < dataSend.length; s++) {
          if (dataSend[s][1] === targetKey) {
            sheetSend.getRange(s + 1, 4).setValue("-1");
            if (obj.note) {
              sheetSend.getRange(s + 1, 6).setValue("ยกเลิก: " + obj.note);
            }
            sheetSend.getRange(s + 1, 7).setValue(formatToDateThai(new Date()));
          }
        }
      }
    } catch (e) {
      Logger.log("Error updating SendDocument on cancel: " + e);
    }

    return docNumber;
  } catch (err) {
    Logger.log("Error in cancelDocument: " + err);
    return null;
  }
}

// ฟังก์ชันดึงรายการหนังสือที่ถูกยกเลิก (สถานะ -1)
function dataDocCanceled(key, userLevel) {
  const sheet = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument"); 
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var data = sheet.getDataRange().getValues().slice(2);

  data = data.filter((row) => {
    const isCanceled = String(row[22]) === "-1" || row[22] === -1;
    if (!isCanceled) return false;

    // ถ้าเป็นสารบรรณกลาง (ระดับ 3) หรือ ผู้ดูแลระบบ (ระดับ 1) ให้เห็นรายการหนังสือที่ยกเลิกทั้งหมด
    if (String(userLevel) === '1' || String(userLevel) === '3') {
      return true;
    }
    // ผู้ใช้งานทั่วไปให้เห็นเฉพาะที่ตนเองเป็นเจ้าของ
    return row[21] === key;
  });

  return data;
}

// ฟังก์ชันดึงรายการหนังสือรับที่ถูกยกเลิก (สำหรับเมนู รับหนังสือ)
function dataDocPickupCanceled(key, userLevel) {
  const ssDoc = SpreadsheetApp.openById(sheetDocuMent);
  const sheetSend = ssDoc.getSheetByName("SendDocument");
  const sheetDoc = ssDoc.getSheetByName("DataDocument");
  if (!sheetSend || !sheetDoc) return [];

  const lastRowDoc = sheetDoc.getLastRow();
  if (lastRowDoc < 3) return [];

  const dataDoc = sheetDoc.getDataRange().getValues().slice(2);
  const dataSend = sheetSend.getDataRange().getValues().slice(1);

  const isSarabun = String(userLevel) === '1' || String(userLevel) === '3';

  // เก็บแผนที่ของเอกสารที่ถูกยกเลิก (status == -1) ใน DataDocument
  const canceledDocMap = new Map();
  for (let i = 0; i < dataDoc.length; i++) {
    const isCanceled = String(dataDoc[i][22]) === "-1" || dataDoc[i][22] === -1;
    if (isCanceled) {
      canceledDocMap.set(dataDoc[i][0], dataDoc[i]);
    }
  }

  const result = [];
  const addedDocIds = new Set();

  for (let j = 0; j < dataSend.length; j++) {
    const sendRow = dataSend[j];
    const docId = sendRow[1];
    const receiverKey = sendRow[2];

    if (isSarabun || receiverKey === key) {
      if (canceledDocMap.has(docId) && !addedDocIds.has(docId)) {
        addedDocIds.add(docId);
        result.push(canceledDocMap.get(docId));
      }
    }
  }

  if (isSarabun) {
    canceledDocMap.forEach((docRow, docId) => {
      if (!addedDocIds.has(docId)) {
        addedDocIds.add(docId);
        result.push(docRow);
      }
    });
  }

  return result;
}

// ฟังก์ชันลบหนังสือออกจากระบบอย่างถาวร (เฉพาะสารบรรณกลางระดับ 3 หรือแอดมินระดับ 1 เท่านั้น)
function deleteCanceledDocumentPermanent(key, userLevel) {
  if (String(userLevel) !== '1' && String(userLevel) !== '3') {
    return { success: false, message: 'ไม่มีสิทธิ์ในการลบหนังสือถาวร (เฉพาะสารบรรณกลางหรือผู้ดูแลระบบเท่านั้น)' };
  }

  const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
  if (!sheetDoc) return { success: false, message: 'ไม่พบชีต DataDocument' };
  
  const lastRow = sheetDoc.getLastRow();
  if (lastRow < 3) return { success: false, message: 'ไม่มีข้อมูลเอกสาร' };

  const dataDoc = sheetDoc.getRange('A3:A' + lastRow).getValues();
  let deletedDocNumber = "";
  let found = false;

  for (let i = dataDoc.length - 1; i >= 0; i--) {
    if (dataDoc[i][0] === key) {
      const rowIndex = i + 3;
      deletedDocNumber = String(sheetDoc.getRange(rowIndex, 3).getValue() || "");
      sheetDoc.deleteRow(rowIndex);
      found = true;
      break;
    }
  }

  // ลบใน SendDocument ด้วยหากมี
  try {
    const sheetSend = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("SendDocument");
    if (sheetSend) {
      const dataSend = sheetSend.getDataRange().getValues();
      for (let s = dataSend.length - 1; s >= 1; s--) {
        if (dataSend[s][1] === key) {
          sheetSend.deleteRow(s + 1);
        }
      }
    }
  } catch (e) {
    Logger.log("Error deleting SendDocument: " + e);
  }

  return { success: found, docNumber: deletedDocNumber };
}

// ฟังก์ชันกู้คืนหนังสือที่ยกเลิกกลับมาเป็นสถานะปกติ (รอดำเนินการ)
function restoreCanceledDocument(key) {
  const sheetDoc = SpreadsheetApp.openById(sheetDocuMent).getSheetByName("DataDocument");
  if (!sheetDoc) return { success: false, message: 'ไม่พบชีต DataDocument' };

  const lastRow = sheetDoc.getLastRow();
  if (lastRow < 3) return { success: false, message: 'ไม่มีข้อมูลเอกสาร' };

  const dataDoc = sheetDoc.getRange('A3:A' + lastRow).getValues();
  let docNumber = "";
  let found = false;

  for (let i = dataDoc.length - 1; i >= 0; i--) {
    if (dataDoc[i][0] === key) {
      const rowIndex = i + 3;
      sheetDoc.getRange(rowIndex, 23).setValue("1"); // ปรับกลับเป็นรอดำเนินการ (1)
      sheetDoc.getRange(rowIndex, 26).setValue(formatToDateThai(new Date()));
      docNumber = String(sheetDoc.getRange(rowIndex, 3).getValue() || "");
      
      // เอาแท็ก [ยกเลิก: ...] ออกจากหมายเหตุ (คอลัมน์ 15)
      let curNote = String(sheetDoc.getRange(rowIndex, 15).getValue() || "");
      curNote = curNote.replace(/\[ยกเลิก:[^\]]*\]/g, "").replace(/\[ยกเลิกหนังสือ\]/g, "").trim();
      sheetDoc.getRange(rowIndex, 15).setValue(curNote);

      found = true;
      break;
    }
  }

  return { success: found, docNumber: docNumber };
}

///////////////////////////////ProJect System////////////////////////////////

const generateProjectNumber = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'POJ';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}

const gsNewProject = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("ProjectData");
  const sheetSend = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const projectKey = generateProjectNumber();  
  const lastRow = sheet.getLastRow();

  const rowData = [projectKey, formatToDateThai(new Date()), obj.name, obj.publicKey, obj.token, obj.data1, obj.data2, obj.data3, obj.data4, , obj.data6, , obj.data8, obj.data9, obj.data10, , obj.data12]
  sheet.appendRow(rowData)
  const data5 = obj.data5.length > 1 ? [obj.data5.join(', ')] : [obj.data5];
  const data7 = obj.data7.length > 1 ? [obj.data7.join(', ')] : [obj.data7];
  const data11 = obj.data11.length > 1 ? [obj.data11.join(', ')] : [obj.data11];
  sheet.getRange(lastRow + 1, 10).setValues([data5]);
  sheet.getRange(lastRow + 1, 12).setValues([data7]);
  sheet.getRange(lastRow + 1, 16).setValues([data11]);
  sheet.getRange(lastRow + 1, 18).setValue(true);
  
  obj.keyusers.forEach(function(key) {
    const projectSendKey = generateRandomKey();
    sheetSend.appendRow([projectSendKey, projectKey, key, "0", formatToDateThai(new Date()), 'รอตอบรับ']);
    var genCode = generateRandomKey();
    notiSheet.appendRow([genCode, key, obj.publicKey, "เสนอโครงการ " + obj.data1, "หน่วยงานที่เสนอ " + obj.data2, "changePageProject", "2", "0", "'" + formatDateSave(new Date())]);
  });

  return data = sheet.getDataRange().getValues().slice(1);
}

const getDataProject = () => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("ProjectData");
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
}
const dataProjectSave = (key) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const sheetDoc = SpreadsheetApp.openById(sheetProject).getSheetByName("ProjectData");
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues().slice(1);
  const dataDoc = sheetDoc.getDataRange().getValues().slice(1);
  const dataUser = sheetUser.getDataRange().getValues().slice(1);

  const filteredDataDoc = dataDoc.filter(row => row[0] === key);
  const filteredData = [];
  const filteredStatus = [];

  for (const row of filteredDataDoc) {
    const keyCode = row[0];
    const matchingRows = data.filter(docRow => docRow[1] === keyCode);
    filteredStatus.push(...matchingRows);
    for (const rowUser of matchingRows) {
      const userCode = rowUser[2];
      const matchingUser = dataUser.filter(userRow => userRow[0] === userCode);
      filteredData.push(...matchingUser);
    }
  }
  
  return [filteredDataDoc, filteredStatus, filteredData];
}
const dataProjectInprogress = (key) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const sheetData = SpreadsheetApp.openById(sheetProject).getSheetByName("ProjectData");
  const data = sheet.getDataRange().getValues().slice(1);
  const dataDoc = sheetData.getDataRange().getValues().slice(1);

  const filteredData = data.filter(row => row[2] === key);
  const filteredDataDoc = [];

  for (const row of filteredData) {
    const userCode = row[1];
    const matchingRows = dataDoc.filter(docRow => docRow[0] === userCode && docRow[17] === true);
    filteredDataDoc.push(...matchingRows);
  }

  filteredDataDoc.reverse();
  filteredData.reverse();
  
  return [filteredData, filteredDataDoc];
}
const openProject = (key) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const userValue = data[i][0];
    if (userValue === key) {
      if (data[i][3] < 1) {
        sheet.getRange(i + 1, 4).setValue("1");
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
      }
      return data[i][2];
    }
  }
}
const saveProjectingData = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
        sheet.getRange(i + 1, 4).setValue(obj.status);
        sheet.getRange(i + 1, 6).setValue(obj.response);
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 8).setValue(obj.jobWord);
        sheet.getRange(i + 1, 9).setValue(obj.replace);
      return data[i][2];
    }
  }
}
const returnProject = (key, text) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("SendProject");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === key) {
        sheet.getRange(i + 1, 4).setValue("2");
        sheet.getRange(i + 1, 6).setValue("ถูกตีกลับ");
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 8).setValue(text);
      return data[i][2];
    }
  }
}
const closeProjectDoc = (key) => {
  const sheet = SpreadsheetApp.openById(sheetProject).getSheetByName("ProjectData");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === key) {
        sheet.getRange(i + 1, 18).setValue(false);
      return data[i][3];
    }
  }
}
const noticationShow = (key) => {
  const sheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  const user = sheetUser.getDataRange().getValues();

  const filteredData = data.filter(row => row[1] === key);
  const filteredDataUser = [];

  for (const row of filteredData) {
    const userCode = row[2];
    const matchingRows = user.filter(user => user[0] === userCode);
    filteredDataUser.push(...matchingRows);
  }

  filteredDataUser.reverse();
  filteredData.reverse();
  
  return [filteredData, filteredDataUser];
}
const openNotify = (key) => {
  const sheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    const publicKey = data[i][0];
    if (publicKey === key) {
      sheet.getRange(i + 1, 8).setValue('1');
      return data[i][1];
    }
  }
}
const saveFileProfile = (e, ipAddress, userAgent) => {
  let folder = DriveApp.getFolderById(userFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === e.code) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "แก้ไขรูปประจำตัว",new Date()];
      sheet.getRange(i + 1, 14).setValue(url);
      history.appendRow(dataHistory);
      break;
    }
  }
  
  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  const dataHRM = sheetHRM.getDataRange().getValues();
  
  for (let i = 1; i < dataHRM.length; i++) {
    const userValue = dataHRM[i][0];  
    if (userValue === e.code) {
      sheetHRM.getRange(i + 1, 10).setValue(url);
      break;
    }
  }

  return url;
}

const changePasswordProflie = (obj, ipAddress, userAgent) => {
  const usersheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const data = usersheet.getDataRange().getValues();
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);

  for (let i = 0; i < data.length; i++) {
    const keyValue = data[i][0];
    if (obj.code === keyValue) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "เปลี่ยนรหัสผ่าน", new Date()];
      usersheet.getRange(i + 1, 3).setValue("'" + obj.password);
      history.appendRow(dataHistory);
      break;
    }
  }
  return obj.code;
}

const edit_Userchangedata = (obj, ipAddress, userAgent) => {
  const usersheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const data = usersheet.getDataRange().getValues();
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);

  for (let i = 0; i < data.length; i++) {
    const keyValue = data[i][0];
    if (obj.key === keyValue) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "แก้ไขข้อมูลส่วนตัว", new Date()];
      usersheet.getRange(i + 1, 4).setValue(obj.data1);
      usersheet.getRange(i + 1, 5).setValue(obj.data2);
      usersheet.getRange(i + 1, 6).setValue(obj.data3);
      usersheet.getRange(i + 1, 10).setValue("'" + obj.data4);
      usersheet.getRange(i + 1, 12).setValue(obj.data5);
      history.appendRow(dataHistory);
      break;
    }
  }
  return obj;
}

const saveFileSignetur = (e, ipAddress, userAgent) => {
  let folder = DriveApp.getFolderById(userFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === e.code) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "แก้ไขลายเซ็นต์",new Date()];
      sheet.getRange(i + 1, 13).setValue(url);
      history.appendRow(dataHistory);
      break;
    }
  }

  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  const dataHRM = sheetHRM.getDataRange().getValues();
  
  for (let i = 1; i < dataHRM.length; i++) {
    const userValue = dataHRM[i][0];  
    if (userValue === e.code) {
      sheetHRM.getRange(i + 1, 16).setValue(url);
      break;
    }
  }

  return url;
}
const getUsernameList = () => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); 
  return data;
}
const getEmailList = () => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 11, lastRow - 1, 1).getValues(); 
  return data;
}
const send_otp_registor = () => {
  var otp = Math.floor(100000 + Math.random() * 900000);
  return otp;
}
const addNewRegictor = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase");
  const publicKey = generateRandomKey();

  var rowData = ["'" + publicKey, "'" + obj.data1, "'" + obj.data2, obj.data3, obj.data4, obj.data5, obj.data13, obj.data14, "'5", "'" + obj.data6, obj.data11, obj.data7 + " หมู่ที่ " + obj.data8 + " " + obj.data10, , , , true, obj.data12]

  sheet.appendRow(rowData);

  var dataHRM = ["'" + publicKey, "", obj.data3, obj.data4, obj.data5, obj.data15, obj.data13, "", "", "", "ปฏิบัติราชการอยู่", formatToDateThai(new Date()), "", "", "", "", obj.data14]  

  sheetHRM.appendRow(dataHRM);

  return { publicKey, username: obj.data1, password: obj.data2 };
}
const saveuploadimageprofile = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const sheetHRM = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase");
  const data = sheet.getDataRange().getValues();
  const dataHRM = sheetHRM.getDataRange().getValues();
  var folder = DriveApp.getFolderById(userFolder);

  var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
  var blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
  var file = folder.createFile(blob);
  var fileId = file.getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileId;

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === obj.code) {
      sheet.getRange(i + 1, 14).setValue(url);
      break;
    }
  }

  for (let i = 1; i < dataHRM.length; i++) {
    const userValue = dataHRM[i][0];  
    if (userValue === obj.code) {
      sheetHRM.getRange(i + 1, 10).setValue(url);
      break;
    }
  }

  return url;
}

const saveFileEditProfileImage = (e, ipAddress, userAgent) => {
  let folder = DriveApp.getFolderById(userFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === e.code) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "แก้ไขรูปประจำโดย Admin", new Date()];
      sheet.getRange(i + 1, 14).setValue(url);
      history.appendRow(dataHistory);
      break;
    }
  }
  
  return url;
}

const editUserDataAll = (obj, ipAddress, userAgent) => {
  const history = SpreadsheetApp.openById(sheetUseds).getSheetByName("History");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);

  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName("User"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === obj.editUserDataAll1) {
      var dataHistory = ["'" + data[i][1], data[i][3] + data[i][4] + " " + data[i][5], ipAddress, browserInfo + " " + osInfo, "แก้ไขข้อมูลโดย Admin", new Date()];

      sheet.getRange(i + 1, 4).setValue(obj.editUserDataAll7);
      sheet.getRange(i + 1, 5).setValue(obj.editUserDataAll8);
      sheet.getRange(i + 1, 6).setValue(obj.editUserDataAll9);
      sheet.getRange(i + 1, 7).setValue(obj.editUserDataAll3);
      sheet.getRange(i + 1, 8).setValue(obj.editUserDataAll2);
      sheet.getRange(i + 1, 9).setValue(obj.editUserDataAll4);
      sheet.getRange(i + 1, 10).setValue("'" + obj.editUserDataAll10);
      sheet.getRange(i + 1, 11).setValue(obj.editUserDataAll11);
      sheet.getRange(i + 1, 12).setValue(obj.editUserDataAll12);

      history.appendRow(dataHistory);
      break;
    }
  }
  
  return;
}