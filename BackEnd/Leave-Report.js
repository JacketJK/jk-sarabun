function generateLaeveNumber() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'LEV';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}

const gsLeaveNew = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const lastRow = sheet.getLastRow();
  const codeLEVNumber = generateLaeveNumber();

  const rowData = [codeLEVNumber, formatToDateThai(new Date()), obj.name, obj.position, obj.agency, obj.data1, obj.data2, formatToDateThaiFull(new Date(obj.data3)), formatToDateThaiFull(new Date(obj.data4)), obj.data5, obj.data6, obj.data7, obj.data8, obj.data9, obj.data10, "1",]

  sheet.appendRow(rowData);
  sheet.getRange(lastRow + 1, 17).setValue(obj.publicKey);
  sheet.getRange(lastRow + 1, 18).setValue(obj.token);
  sheet.getRange(lastRow + 1, 19).setValue(obj.tokenuser);

  const sheetSend = SpreadsheetApp.openById(sheetLeave).getSheetByName('StepLeave');
  obj.keyusers.forEach(function(key) {
    const leaveSendKey = generateRandomKey();
    const genCode = generateRandomKey();
    sheetSend.appendRow([leaveSendKey, codeLEVNumber, key, "0", formatToDateThai(new Date()), 'รอตอบรับ']);
    notiSheet.appendRow([genCode, key, obj.publicKey, "ขออนุญาตลา " + obj.data1, "เนื่องจาก" + obj.data2 + " ในวันที่" + formatToDateThaiFull(new Date(obj.data3)), "changePageLeave", "4", "0", "'" + formatDateSave(new Date())]);
  });

  var token = obj.tokenuser;
  console.log("DEBUG :" + obj);
  var alt = "การอนุญาตลา " + obj.data1;

  var msg = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "แจ้งเตือนการขออนุญาตลา",
        "weight": "bold",
        "color": "#1DB446",
        "size": "lg"
      },
      {
        "type": "text",
        "text": codeLEVNumber,
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
                "text": "ผู้ขอลา",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.name,
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
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": "สังกัด",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.agency,
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
                "text": "ประเภทการลา",
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
                "text": "เหตุผล",
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
                "text": "ตั้งแต่วันที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": formatToDateThaiFull(new Date(obj.data3)),
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
                "text": "ถึงวันที่",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": formatToDateThaiFull(new Date(obj.data4)),
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
                "text": "เป็นเวลา",
                "size": "sm",
                "color": "#555555",
                "flex": 0
              },
              {
                "type": "text",
                "text": obj.data5 + " วัน",
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
                "text": "ที่อยู่ที่สามารถติดต่อได้ : " + obj.data10,
                "size": "sm",
                "color": "#111111",
              },
            ]
          },
        ]
      }
    ]
  };

  sendClosingAnnouncement(token, msg, alt); 

  return;
}

const getDataLeaveAll = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  var data = sheet.getDataRange().getValues().slice(1);
  data = data.filter(function(row) {
    return row[16] === key;
  });
  var selectedData = data.map(function(row) {
    return [row[0], row[5], row[6], row[7], row[8], row[9], row[15], row[14]];
  });
  return selectedData;
}

const getDataLoadLeave = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  var data = sheet.getDataRange().getValues();
  var filteredData = [];

  for (var i = data.length - 1; i >= 1; i--) { 
    var keyValue = data[i][16]; 
    if (keyValue === key) {
      filteredData.push(data[i]);
      if (filteredData.length >= 1) {
        break;
      }
    }
  }

  var selectedData = filteredData.map(function(row) {
    return [row[5], row[7], row[8], row[9], row[14]];
  });

  return selectedData;
}

const editLeaveGs = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1;i > 0; i--) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
        sheet.getRange(i + 1, 6).setValue(obj.data1);
        sheet.getRange(i + 1, 7).setValue(obj.data2);
        sheet.getRange(i + 1, 8).setValue(formatToDateThaiFull(new Date(obj.data3)));
        sheet.getRange(i + 1, 9).setValue(formatToDateThaiFull(new Date(obj.data4)));
        sheet.getRange(i + 1, 10).setValue(obj.data5);
        sheet.getRange(i + 1, 15).setValue(obj.data6);
      return data[i][0];
    }
  }
}
const cancelLeaveGs = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1;i > 0; i--) {
    var keyValue = data[i][0];
    if (keyValue === key) {
        sheet.getRange(i + 1, 16).setValue("0");
      return data[i][16];
    }
  }
}

const getDataLeaveCheck = (status) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName('LeaveData');
  var data = sheet.getDataRange().getValues();

  data = data.filter(function(row) {
    return row[15] == status;
  });
  var selectedData = data.map(function(row) {
    return [row[0], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[14], row[15]];
  });
  return selectedData;
}

const dataLeaveInprogress = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("LeaveData");
  const sheetStep = SpreadsheetApp.openById(sheetLeave).getSheetByName("StepLeave");
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues().slice(1);
  const dataStep = sheetStep.getDataRange().getValues().slice(1);
  const dataUser = sheetUser.getDataRange().getValues().slice(1);

  const filteredData = data.filter(row => row[0] === key);
  const filteredStep = [];
  const filteredUser = [];

  for (const row of filteredData) {
    const userCode = row[0];
    const matchingRows = dataStep.filter(stepRow => stepRow[1] === userCode);
    filteredStep.push(...matchingRows);
    for (const rowUser of matchingRows) {
      const userCode = rowUser[2];
      const matchingUser = dataUser.filter(userRow => userRow[0] === userCode);
      filteredUser.push(...matchingUser);
    }
  }

  filteredStep.reverse();
  filteredData.reverse();
  filteredUser.reverse();
  
  return [filteredData, filteredStep, filteredUser];
}

const openDocLeave = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("StepLeave");
  const sheetData = SpreadsheetApp.openById(sheetLeave).getSheetByName("LeaveData");
  const data = sheet.getDataRange().getValues();
  const dataLeave = sheetData.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    const leaveValue = data[i][1];
    if (keyValue === key) {
      if (data[i][3] < 1) {
        sheet.getRange(i + 1, 4).setValue("1");
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        for (let j = dataLeave.length - 1; j >= 0; j--) {
          if (leaveValue === dataLeave[j][0] && dataLeave[j][15] < 3) {
            sheetData.getRange(j + 1, 16).setValue("2");
            break;
          }
        }
      }
      return data[i][2];
    }
  }
}

const saveLeavingData = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("StepLeave");
  const sheetData = SpreadsheetApp.openById(sheetLeave).getSheetByName("LeaveData");
  const data = sheet.getDataRange().getValues();
  const dataLeave = sheetData.getDataRange().getValues();
  const leaveSendKey = generateRandomKey();
  var alt = "";
  var msg = "";

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    const leaveValue = data[i][1];
    if (keyValue === obj.key) {
        sheet.getRange(i + 1, 4).setValue("3");
        sheet.getRange(i + 1, 6).setValue(obj.response);
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 8).setValue(obj.jobWord);
        sheet.getRange(i + 1, 9).setValue(obj.replace);
        for (let j = dataLeave.length - 1; j >= 0; j--) {
          if (leaveValue === dataLeave[j][0]) {
            sheetData.getRange(j + 1, 16).setValue(obj.status);
            sheetData.getRange(j + 1, 20).setValue(obj.tokenuser);

            alt = "พิจารณาการอนุญาตลา ของ" + dataLeave[j][2];
            msg = {
              "type": "box",
              "layout": "vertical",
              "contents": [
                {
                  "type": "text",
                  "text": "แจ้งเตือนการขออนุญาตลา",
                  "weight": "bold",
                  "color": "#1DB446",
                  "size": "lg"
                },
                {
                  "type": "text",
                  "text": dataLeave[j][0],
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
                          "text": "ผู้ขอลา",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][2],
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
                          "text": dataLeave[j][3],
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
                          "text": "สังกัด",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][4],
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
                          "text": "ประเภทการลา",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][5],
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
                          "text": "เหตุผล",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][6],
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
                          "text": "ตั้งแต่วันที่",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][7],
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
                          "text": "ถึงวันที่",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][8],
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
                          "text": "เป็นเวลา",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": dataLeave[j][9] + " วัน",
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
                          "text": "เกษียน",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": obj.jobWord,
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
                          "text": "การพิจารณา",
                          "size": "sm",
                          "color": "#555555",
                          "flex": 0
                        },
                        {
                          "type": "text",
                          "text": obj.response,
                          "size": "sm",
                          "color": "#111111",
                          "align": "end",
                        }
                      ]
                    },
                  ]
                }
              ]
            };

            break;
          }
        }

        sheet.appendRow([leaveSendKey, leaveValue, obj.keyuser, "0", formatToDateThai(new Date()), 'รอตอบรับ']);
        var token = obj.tokenuser;
        sendClosingAnnouncement(token, msg, alt); 
      return data[i][2];
    }
  }
}

const saveLeavingDataFinal = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("StepLeave");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
        sheet.getRange(i + 1, 4).setValue("3");
        sheet.getRange(i + 1, 6).setValue(obj.response);
        sheet.getRange(i + 1, 7).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 8).setValue(obj.jobWord);
        sheet.getRange(i + 1, 9).setValue(obj.replace);
      return data[i][2];
    }
  }
}
const returnLeave = (key, text) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("StepLeave");
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
const closeLeaveDoc = (key) => {
  const sheet = SpreadsheetApp.openById(sheetLeave).getSheetByName("LeaveData");
  const data = sheet.getDataRange().getValues();

  var alt = "";
  var msg = "";

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === key) {
        sheet.getRange(i + 1, 16).setValue("6");

        alt = "แจ้งผลการพิจารณาการอนุญาตลา" + data[i][5];
        msg = {
          "type": "box",
          "layout": "vertical",
          "contents": [
            {
              "type": "text",
              "text": "แจ้งเตือนการขออนุญาตลา",
              "weight": "bold",
              "color": "#1DB446",
              "size": "lg"
            },
            {
              "type": "text",
              "text": data[i][0],
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
                      "text": "ผู้ขอลา",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][2],
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
                      "text": data[i][3],
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
                      "text": "สังกัด",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][4],
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
                      "text": "ประเภทการลา",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][5],
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
                      "text": "เหตุผล",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][6],
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
                      "text": "ตั้งแต่วันที่",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][7],
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
                      "text": "ถึงวันที่",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][8],
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
                      "text": "เป็นเวลา",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": data[i][9] + " วัน",
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
                      "text": "การพิจารณา",
                      "size": "sm",
                      "color": "#555555",
                      "flex": 0
                    },
                    {
                      "type": "text",
                      "text": "อนุญาต",
                      "size": "sm",
                      "color": "#111111",
                      "align": "end",
                    }
                  ]
                },
              ]
            }
          ]
        };

        var token = data[i][17];
        sendClosingAnnouncement(token, msg, alt); 
      return data[i][16];
    }
  }
}

///////////////////////////////Report System////////////////////////////////

function generateReportNumber() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'REP';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}

const gsaddReport = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName('Report');
  const sheetSend = SpreadsheetApp.openById(sheetReport).getSheetByName('SendReport');
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const codeREPNumber = generateLaeveNumber();

  const rowData = [codeREPNumber, formatToDateThai(new Date()), obj.publicKey, obj.name, obj.position, obj.agency, obj.sig, obj.data1, obj.data2, formatToDateThaiFull(new Date(obj.data3)), obj.data4, obj.data5, obj.data6, obj.data7, true]
  sheet.appendRow(rowData);

  obj.keyusers.forEach(function(key) {
  const reportSendKey = generateRandomKey();
    sheetSend.appendRow([reportSendKey, codeREPNumber, key, "0", formatToDateThai(new Date()), 'รอตอบรับ']);
    var genCode = generateRandomKey();
    notiSheet.appendRow([genCode, key, obj.publicKey, "รายงานประจำเดือน หมู่ที่ " + obj.data1 + " " + obj.data2, "ประชุม" + formatToDateThaiFull(new Date(obj.data3)) + " จำนวนผู้เข้าร่วมประชุม "+ obj.data4 + " ราย", "changePageReport", "1", "0", "'" + formatDateSave(new Date())]);
  });

  return;
}

const getDataReport = () => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName('Report');
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
}
const getDataAllReport = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName('Report');
  var data = sheet.getDataRange().getValues().slice(1);
  data = data.filter(function(row) {
    return row[2] === key;
  });
  return data;
}

const dataReportInprogress = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
  const sheetData = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
  const data = sheet.getDataRange().getValues().slice(1);
  const dataDoc = sheetData.getDataRange().getValues().slice(1);

  const filteredData = data.filter(row => row[2] === key);
  const filteredDataDoc = [];

  for (const row of filteredData) {
    const userCode = row[1];
    const matchingRows = dataDoc.filter(docRow => docRow[0] === userCode && docRow[14] == true);
    filteredDataDoc.push(...matchingRows);
  }

  filteredDataDoc.reverse();
  filteredData.reverse();
  
  return [filteredData, filteredDataDoc];
}

const getDataShowReport = (year, month) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
  const data = sheet.getDataRange().getValues();
  var filteredData = []

  for (let i = data.length - 1; i >= 1 ; i--) {
    var filter = data[i][9].split(' ');
    if (filter[4] == year && filter[2] === month) {
      filteredData.push(data[i]);
    }
  }

  return filteredData;
}

const openReport = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
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
const dataReportSave = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
  const sheetDoc = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
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

const returnReport = (key, text) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
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
const closeReturnDoc = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const keyValue = data[i][0];
    if (keyValue === key) {
        sheet.getRange(i + 1, 15).setValue(false);
      return data[i][2];
    }
  }
}
const saveReportingData = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
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
const deleteReport = (key) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
  const reportData = sheet.getDataRange().getValues();

  const sheetSent = SpreadsheetApp.openById(sheetReport).getSheetByName("SendReport");
  const sentData = sheetSent.getDataRange().getValues();

  const reportRowsToDelete = [];
  const sentRowsToDelete = [];

  reportData.forEach((row, index) => {
    if (row[0] === key) {
      reportRowsToDelete.push(index + 1);
    }
  });

  sentData.forEach((row, index) => {
    if (row[1] === key) {
      sentRowsToDelete.push(index + 1);
    }
  });

  reportRowsToDelete.sort((a, b) => b - a).forEach((rowIndex) => {
    sheet.deleteRow(rowIndex);
  });

  sentRowsToDelete.sort((a, b) => b - a).forEach((rowIndex) => {
    sheetSent.deleteRow(rowIndex);
  });
}

const editingReport = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetReport).getSheetByName("Report");
  const reportData = sheet.getDataRange().getValues();

  for (let i = reportData.length - 1; i >= 0; i--) {
    const keyValue = reportData[i][0];
    if (keyValue === obj.key) {
        sheet.getRange(i + 1, 2).setValue(formatToDateThai(new Date()));
        sheet.getRange(i + 1, 10).setValue(formatToDateThaiFull(new Date(obj.data1)));
        sheet.getRange(i + 1, 11).setValue(obj.data2);
        sheet.getRange(i + 1, 12).setValue(obj.data3);
        sheet.getRange(i + 1, 13).setValue(obj.data4);
        sheet.getRange(i + 1, 14).setValue(obj.data5);
      return reportData[i][2];
    }
  }
}