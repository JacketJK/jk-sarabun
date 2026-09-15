function generateCarNumber() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'CAR';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}
function formatToDateThaiFull(date) {
  var thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  var thaiMonth = thaiMonths[date.getMonth()];
  var thaiYear = date.getFullYear() + 543;

  var thaiDate = 'วันที่ ' + date.getDate() + ' ' + thaiMonth + ' ' + 'พ.ศ. ' + thaiYear;
  return thaiDate;
}
function getDataForCar() {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName('Car');
  const dataRange = sheet.getDataRange();
  var data = dataRange.getValues().slice(1);

  data = data.filter(function(row) {
    return row[11] == true;
  });
  return data; 
}
function addCarRequrst(obj) {
  const sheetData = SpreadsheetApp.openById(sheetCarReq).getSheetByName("CarData");
  const sheetCar = SpreadsheetApp.openById(sheetCarReq).getSheetByName("Car"); 
  const data = sheetCar.getDataRange().getDisplayValues();
  var lastRow = sheetData.getLastRow();

  var codeCarNumber = generateCarNumber();
  var rowData = [codeCarNumber, formatToDateThai(new Date()), obj.data1, obj.user, obj.position, obj.data2, obj.data3, obj.data8, obj.data4, formatToDateThaiFull(new Date(obj.data5)), formatToDateThaiFull(new Date(obj.data6)), obj.data7, obj.data9];

  sheetData.appendRow(rowData);
  sheetData.getRange(lastRow + 1, 22).setValue(obj.sig);
  sheetData.getRange(lastRow + 1, 25).setValue(obj.token);
  sheetData.getRange(lastRow + 1, 26).setValue(obj.userSender);
  sheetData.getRange(lastRow + 1, 28).setValue(obj.image);
  sheetData.getRange(lastRow + 1, 29).setValue(obj.key);
  sheetData.getRange(lastRow + 1, 30).setValue(obj.keySender);

  for (let i = 1 ; i<=data.length; i++){
    let carRegistration = data[i][2];
    if (carRegistration === obj.data1) {
      sheetCar.getRange(i + 1, 6).setValue('');
      sheetCar.getRange(i + 1, 7).setValue('');
      sheetCar.getRange(i + 1, 8).setValue(obj.user);
      sheetCar.getRange(i + 1, 9).setValue(obj.data2 + ' เพื่อ' + obj.data3);
      sheetCar.getRange(i + 1, 10).setValue('ยังไม่ได้รับการอนุญาต');
      sheetCar.getRange(i + 1, 11).setValue(codeCarNumber);
      break;
    }
  }

  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  obj.keySender.forEach(function(key) {
    var genCode = generateRandomKey();
    notiSheet.appendRow([genCode, key, obj.key, "ขออนุญาตใช้รถ " + obj.data1, "ไปที่" + obj.data2 + " เพื่อ"+ obj.data3, "changePageCarReserve", "3", "0", "'" + formatDateSave(new Date())]);
  });

  var token = obj.userSender[0];

  var alt = "มีผู้ขอใช้รถ " + obj.data1;
  var msg = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "แจ้งเตือนการขอใช้รถส่วนกลาง",
        "weight": "bold",
        "color": "#1DB446",
        "size": "lg"
      },
      {
        "type": "text",
        "text": codeCarNumber,
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
                "text": "เลขทะเบียน",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้ขอใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.user,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ตำแหน่ง",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.position,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          }
        ]
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
                "text": "ไปที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data2,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เพื่อใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data3,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เขตพื้นที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data8,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ใน" + formatToDateThaiFull(new Date(obj.data5)),
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ถึง" + formatToDateThaiFull(new Date(obj.data6)),
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "พลขับ",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data7,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "มีคนนั่งไปด้วย",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data4 + " ท่าน",
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
      }
    ]
  };

  sendClosingAnnouncement(token, msg, alt); 
}

const getDataDrive = () => {
  const sheet = SpreadsheetApp.openById(sheetUseds).getSheetByName('User');
  const data = sheet.getDataRange().getValues().slice(1);

  const fullName = data.map(function(row) {
    return `${row[3]}${row[4]} ${row[5]}`
  });

  return fullName;
};

function considerCarRequrst(obj) {
  const sheetData = SpreadsheetApp.openById(sheetCarReq).getSheetByName("CarData");
  const dataCar = sheetData.getDataRange().getValues();

  for (let i = 1 ; i<=dataCar.length; i++){
    let code = dataCar[i][0];
    if (code === obj.consider4) {
      sheetData.getRange(i + 1, 17).setValue(obj.consider5);
      sheetData.getRange(i + 1, 18).setValue(obj.consider1);
      sheetData.getRange(i + 1, 23).setValue(obj.consider2);
      sheetData.getRange(i + 1, 26).setValue(obj.consider3);
      sheetData.getRange(i + 1, 27).setValue(obj.userSender);
      sheetData.getRange(i + 1, 31).setValue(obj.keySender);
      sheetData.getRange(i + 1, 33).setValue(obj.consider6);

      usercar2  = sheetData.getRange(i + 1, 4).getValue();
      usercar3  = sheetData.getRange(i + 1, 5).getValue();
      usercar1  = sheetData.getRange(i + 1, 3).getValue();
      usercar4  = sheetData.getRange(i + 1, 6).getValue();
      usercar5  = sheetData.getRange(i + 1, 7).getValue();
      usercar8  = sheetData.getRange(i + 1, 10).getValue();
      usercar9  = sheetData.getRange(i + 1, 11).getValue();
      usercar6  = sheetData.getRange(i + 1, 8).getValue();
      usercar7  = sheetData.getRange(i + 1, 9).getValue();
      usercar10  = sheetData.getRange(i + 1, 12).getValue();

      break;
    }
  }
  
  // var token = [obj.userSender];
  // var msg = '📣มีผู้ขอใช้รถ🚗\n' + '\n⚙️รหัสคำขอ: ' + obj.consider4 + '\n🧑✈️ผู้ขอใช้: ' + usercar2 + '\n🏢ตำแหน่ง: ' + usercar3 + '\n🚘รถที่ขอใช้: ' + usercar1 + '\n---------------------------------\n                 📎รายละเอียด\n' +  '\n📌ไปที่: ' + usercar4 + '\n✨เพื่อใช้: ' + usercar5 + '\n📅ใน' + usercar8 + '\n📅ถึง' + usercar9 + '\n🗾เขตพื้นที่: ' + usercar6 + '\n🪑มีผู้นั่งไปด้วย จำนวน ' + usercar7 + ' คน' +'\n🧑✈️พลขับ: ' + usercar10 + '\n\n---------------------------------\n' + '🪧ข้อพิจารณา: ' + obj.consider5 + '\n🤫ผู้เสนอ: ' + obj.consider1;

  var token = obj.userSender;
  var alt = "การเสนอพิจารณาขอใช้รถ " + usercar1;

  var msg = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "แจ้งเตือนการขอใช้รถส่วนกลาง",
        "weight": "bold",
        "color": "#1DB446",
        "size": "lg"
      },
      {
        "type": "text",
        "text": obj.consider4,
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
                "text": "เลขทะเบียน",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้ขอใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar2,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ตำแหน่ง",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar3,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          }
        ]
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
                "text": "ไปที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar4,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เพื่อใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar5,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เขตพื้นที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar6,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ใน" + usercar8,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ถึง" + usercar9,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "พลขับ",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar10,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "มีคนนั่งไปด้วย",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar7 + " ท่าน",
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
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
                "text": "ข้อพิจารณา",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.consider5,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้เสนอ",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.consider1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
      }
    ]
  };

  sendClosingAnnouncement(token, msg, alt); 
}
function approveCarRequrst(obj) {
  const sheetData = SpreadsheetApp.openById(sheetCarReq).getSheetByName("CarData");
  const dataCar = sheetData.getDataRange().getValues();

  const sheetOffice = SpreadsheetApp.openById(sheetCarReq).getSheetByName("Car"); 
  const data = sheetOffice.getDataRange().getValues();

  for (let i = 1; i <= dataCar.length; i++){
    let code = dataCar[i][0];
    if (code === obj.approve4) {
      sheetData.getRange(i + 1, 19).setValue(obj.approve5);
      sheetData.getRange(i + 1, 20).setValue(obj.approve1);
      sheetData.getRange(i + 1, 21).setValue(formatToDateThaiFull(new Date()));
      sheetData.getRange(i + 1, 24).setValue(obj.approve2);
      sheetData.getRange(i + 1, 27).setValue(obj.approve3);
      sheetData.getRange(i + 1, 34).setValue(obj.approve6);

      usercar2  = sheetData.getRange(i + 1, 4).getValue();
      usercar3  = sheetData.getRange(i + 1, 5).getValue();
      usercar1  = sheetData.getRange(i + 1, 3).getValue();
      usercar4  = sheetData.getRange(i + 1, 6).getValue();
      usercar5  = sheetData.getRange(i + 1, 7).getValue();
      usercar8  = sheetData.getRange(i + 1, 10).getValue();
      usercar9  = sheetData.getRange(i + 1, 11).getValue();
      usercar6  = sheetData.getRange(i + 1, 8).getValue();
      usercar7  = sheetData.getRange(i + 1, 9).getValue();
      usercar10  = sheetData.getRange(i + 1, 12).getValue();
      consider5  = sheetData.getRange(i + 1, 17).getValue();
      consider1  = sheetData.getRange(i + 1, 18).getValue();
      tokens = sheetData.getRange(i + 1, 25).getValue();

      break;
    }
  }

  for (let i = 1 ; i <=data.length; i++){
    let code = data[i][10];
    if (code === obj.approve4) {
      sheetOffice.getRange(i + 1, 10).setValue("อยู่ระหว่างใช้งาน");
      break;
    }
  }

  // var token = tokens;
  // var msg = '📣มีผู้ขอใช้รถ🚗\n' + '\n⚙️รหัสคำขอ: ' + obj.approve4 + '\n🧑✈️ผู้ขอใช้: ' + usercar2 + '\n🏢ตำแหน่ง: ' + usercar3 + '\n🚘รถที่ขอใช้: ' + usercar1 + '\n---------------------------------\n                 📎รายละเอียด\n' +  '\n📌ไปที่: ' + usercar4 + '\n✨เพื่อใช้: ' + usercar5 + '\n📅ใน' + usercar8 + '\n📅ถึง' + usercar9 + '\n🗾เขตพื้นที่: ' + usercar6 + '\n🪑มีผู้นั่งไปด้วย จำนวน ' + usercar7 + ' คน' +'\n🧑✈️พลขับ: ' + usercar10 + '\n\n---------------------------------\n' + '🪧ข้อพิจารณา: ' + consider5 + '\n🤫ผู้เสนอ: ' + consider1 + '\n🎀ผลการพิจารณา: ' + 
  // obj.approve5 + '\n🏡ผู้พิจารณา: ' + obj.approve1;

  var token = tokens;
  var alt = "การอนุญาตขอใช้รถ " + usercar1;

  var msg = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "แจ้งเตือนการขอใช้รถส่วนกลาง",
        "weight": "bold",
        "color": "#1DB446",
        "size": "lg"
      },
      {
        "type": "text",
        "text": obj.approve4,
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
                "text": "เลขทะเบียน",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้ขอใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar2,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ตำแหน่ง",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar3,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          }
        ]
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
                "text": "ไปที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar4,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เพื่อใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar5,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เขตพื้นที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar6,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ใน" + usercar8,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ถึง" + usercar9,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "พลขับ",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar10,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "มีคนนั่งไปด้วย",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar7 + " ท่าน",
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
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
                "text": "ผลการพิจารณา",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.approve5,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้พิจารณา",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.approve1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
      }
    ]
  };

  sendClosingAnnouncement(token, msg, alt); 
}

function returnCarRequrst(obj) {
  const sheetData = SpreadsheetApp.openById(sheetCarReq).getSheetByName("CarData");
  const dataCar = sheetData.getDataRange().getValues();

  const sheetOffice = SpreadsheetApp.openById(sheetCarReq).getSheetByName("Car"); 
  const data = sheetOffice.getDataRange().getValues();

  for (let i = 1 ;i <= dataCar.length; i++){
    if (sheetData.getRange(i + 1, 1).getValue() === obj.returnData1) {
      sheetData.getRange(i + 1, 14).setValue(obj.returnData2);
      sheetData.getRange(i + 1, 15).setValue(obj.returnData3);
      sheetData.getRange(i + 1, 16).setValue(obj.returnData4);
      sheetData.getRange(i + 1, 32).setValue(formatToDateThai(new Date()));

      usercar2  = sheetData.getRange(i + 1, 4).getValue();
      usercar3  = sheetData.getRange(i + 1, 5).getValue();
      usercar1  = sheetData.getRange(i + 1, 3).getValue();
      usercar4  = sheetData.getRange(i + 1, 6).getValue();
      usercar5  = sheetData.getRange(i + 1, 7).getValue();
      usercar8  = sheetData.getRange(i + 1, 10).getValue();
      usercar9  = sheetData.getRange(i + 1, 11).getValue();
      usercar6  = sheetData.getRange(i + 1, 8).getValue();
      usercar7  = sheetData.getRange(i + 1, 9).getValue();
      usercar10  = sheetData.getRange(i + 1, 12).getValue();
      milebefor  = sheetData.getRange(i + 1, 13).getDisplayValue();

      tokens1 = sheetData.getRange(i + 1, 25).getValue();
      tokens2 = sheetData.getRange(i + 1, 26).getValue();
      tokens3 = sheetData.getRange(i + 1, 27).getValue();

      break;
    }
  }

  for (let i = 1 ; i <= data.length; i++){   
    if (sheetOffice.getRange(i + 1, 11).getValue() === obj.returnData1) {
      sheetOffice.getRange(i + 1, 10).setValue("พร้อมใช้งาน");
      sheetOffice.getRange(i + 1, 5).setValue(obj.returnData2);
      sheetOffice.getRange(i + 1, 6).setValue(obj.returnData3);
      sheetOffice.getRange(i + 1, 7).setValue(obj.returnData4);

      break;
    }
  }
  // var msg = '📣ขอคืนรถ🚗\n' + '\n⚙️รหัสคำขอ: ' + obj.returnData1 + '\n🧑✈️ผู้ขอใช้: ' + usercar2 + '\n🏢ตำแหน่ง: ' + usercar3 + '\n🚘รถที่ขอใช้: ' + usercar1 + '\n---------------------------------\n                 📎รายละเอียด\n' +  '\n📌ไปที่: ' + usercar4 + '\n✨เพื่อใช้: ' + usercar5 + '\n📅ใน' + usercar8 + '\n📅ถึง' + usercar9 + '\n🗾เขตพื้นที่: ' + usercar6 + '\n🪑มีผู้นั่งไปด้วย จำนวน ' + usercar7 + ' คน' +'\n🧑✈️พลขับ: ' + usercar10 + '\n\n---------------------------------\n' + '🕛เลขไมล์ก่อนไป: ' + milebefor + '\n🕛เลขไมล์ปัจจุบัน: ' + obj.returnData2 + '\n🧭ระยะทาง: ' + obj.returnData4 + ' กิโลเมตร\n⛽น้ำมันคงเหลือ: ' + obj.returnData3 + ' %';

  var mile = "" + obj.returnData2; 
  var tokens = [tokens1, tokens2, tokens3];
  var alt = "ขอคืนรถ " + usercar1;
  var msg = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "แจ้งเตือนการขอใช้รถส่วนกลาง",
        "weight": "bold",
        "color": "#1DB446",
        "size": "lg"
      },
      {
        "type": "text",
        "text": obj.returnData1,
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
                "text": "เลขทะเบียน",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar1,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ผู้ขอใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar2,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ตำแหน่ง",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar3,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          }
        ]
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
                "text": "ไปที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar4,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เพื่อใช้",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar5,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เขตพื้นที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": usercar6,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ใน" + usercar8,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "📅",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": "ถึง" + usercar9,
                "size": "sm",
                "color": "#111111",
                "align": "end",
              }
            ]
          },
        ]
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
                "text": "เลขไมล์ก่อนไป",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": milebefor,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เลขไมล์ปัจจุบัน",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": mile,
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "ระยะทางรวม",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.returnData4 + " กิโลเมตร",
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "เชื้อเพลิงคงเหลือ",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.returnData3 + " %",
                "size": "sm",
                "color": "#111111",
                "align": "end"
              }
            ]
          },
        ]
      }
    ]
  };

  for (var i = 0; i < tokens.length; i++) {
      var currentToken = tokens[i];

      try {
          sendClosingAnnouncement(currentToken, msg, alt); 
      } catch (error) {
        
      }
  }
}

function getDataAllowCarGs() {
  var sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName('CarData');
  var dataRange = sheet.getDataRange(); 
  var data = dataRange.getValues().slice(1);

  data = data.filter(function(row) {
    return row[16] === '';
  });

  data.reverse();

  return data;
}
function getDataConsideCarGs() {
  var sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName('CarData');
  var dataRange = sheet.getDataRange(); 
  var data = dataRange.getValues().slice(1);

  data = data.filter(function(row) {
    return row[18] === '' && row[16] !== '';
  });

  data.reverse();

  return data;
}
function getDataAllHistory() {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName('CarData');
  var data = sheet.getDataRange().getValues().slice(1);
  data = data.filter(function(row) {
    return row[31] !== '';
  });
  return data;
}

function selectNumberCar() {
  var sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName('Car');
  var data = sheet.getRange('C2:C' + sheet.getLastRow()).getValues();
  
  data = data.filter(function(row) {
    return row[0] !== '';
  });
  return data;
}
///////////////////////////////Calendar System////////////////////////////////

function generateCalendarNumber() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'CLD';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}
const gsActivityNew = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetCalendar).getSheetByName('Calendar');
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");

  obj.keyusers.forEach(function(key, index) {
    const codeCldNumber = generateCalendarNumber();
    var genCode = generateRandomKey();
    sheet.appendRow([codeCldNumber, obj.data2, "'" + obj.data3, "'" + obj.data4, obj.data5, obj.data6, obj.data7, obj.data1, key, obj.tokens[index], obj.name, obj.textColor, obj.backgroundColor, obj.borderColor, obj.publicKey]);
    notiSheet.appendRow([genCode, key, obj.publicKey, "เพิ่มกิจกรรมบนปฏิทินของท่าน", "กิจจกรรม" + obj.data2 + " สถานที่"+ obj.data5, "changePageCalendar", "1", "0", "'" + formatDateSave(new Date())]);
  });
}

function getCalendar(key) {
  var sheet = SpreadsheetApp.openById(sheetCalendar).getSheetByName('Calendar');
  var data = sheet.getDataRange().getDisplayValues().slice(1);

  data = data.filter(function(row) {
    return row[8] === key;
  });

    var sheetData = [];

    for (var row of data) {
      var newRow = [];
      for (var cell of row) {
        newRow.push(cell);
      }
      sheetData.push(newRow);
    }
    
  var events = [];

  for (var i = 0; i < sheetData.length; i++) {
    var eventData = sheetData[i];
    var event = {
      codekey: eventData[0],
      title: eventData[1],
      start: eventData[2],
      end: eventData[3],
      location: eventData[4],
      activity: eventData[5],
      note: eventData[6],
      agency: eventData[7],
      userKey: eventData[8],
      requestedBy: eventData[10],
      textColor: eventData[11],
      backgroundColor: eventData[12],
      borderColor: eventData[13],
    };
    events.push(event);
  }
  return events;
}

function getDataAllCalendar(key) {
  const sheet = SpreadsheetApp.openById(sheetCalendar).getSheetByName('Calendar');
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues().slice(1);
  const dataUser = sheetUser.getDataRange().getValues().slice(1);

  const userMap = new Map();
  for (let u = 0; u < dataUser.length; u++) {
    userMap.set(dataUser[u][0], dataUser[u]);
  }

  const filteredData = data.filter(row => row[14] === key);
  const filteredUser = [];

  for (const row of filteredData) {
    const userCode = row[8];
    const matched = userMap.get(userCode);
    if (matched) {
      filteredUser.push(matched);
    } else {
      filteredUser.push(["", "", "", "-", "", "", "-", "-"]);
    }
  }
  filteredData.reverse();
  filteredUser.reverse();

  return [filteredData, filteredUser];
}

function deleteActivity(key) {
  const sheet = SpreadsheetApp.openById(sheetCalendar).getSheetByName('Calendar');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) { 
    const keyValue = data[i][0]; 
    if (keyValue === key) {
      sheet.deleteRow(i + 1); 
      break;
    }
  }
  
  return true;
}
const editActivityCalendar = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetCalendar).getSheetByName('Calendar');
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) { 
    const keyValue = data[i][0]; 
    if (keyValue === obj.key) {
      sheet.getRange(i + 1, 2).setValue(obj.data1);
      sheet.getRange(i + 1, 3).setValue("'" + obj.data4);
      sheet.getRange(i + 1, 4).setValue("'" + obj.data5);
      sheet.getRange(i + 1, 5).setValue(obj.data3);
      sheet.getRange(i + 1, 6).setValue(obj.data2);
      sheet.getRange(i + 1, 7).setValue(obj.data6);
      return obj;
    }
  }
}