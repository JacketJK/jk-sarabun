const setEditSetAgency = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Institute"); 

  var valuesToSet = [
    [obj.dataSetAgency1],
    [obj.dataSetAgency2],
    [obj.dataSetAgency3],
    [obj.dataSetAgency4],
    [obj.dataSetAgency5],
    [obj.dataSetAgency6],
    [obj.dataSetAgency7],
    [obj.dataSetAgency8],
    [obj.dataSetAgency9],
    [obj.dataSetAgency10],
  ];

  var range = sheet.getRange('C3:C12');
  range.setValues(valuesToSet);
}
const getCarSet = () => {
  const ss = SpreadsheetApp.openById(sheetCarReq);
  const sheet = ss.getSheetByName("Car"); 
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
}
const saveOrderCar = (order) => {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName("Car"); 
  const data = sheet.getDataRange().getValues().slice(1);
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
}
const getMeetingRoomSet = () => {
  const ss = SpreadsheetApp.openById(sheetMeet);
  const sheet = ss.getSheetByName("MeetingRoom"); 
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
}
const saveOrderMeetingRoom = (order) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName("MeetingRoom"); 
  const data = sheet.getDataRange().getValues().slice(1);
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
}
const getDepartMentSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("DepartMent"); 
  const data = sheet.getRange("B3:F" + sheet.getLastRow()).getValues();
  return data;
}
const saveOrderDepartMent = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DepartMent"); 
  const data = sheet.getRange("B3:F" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const getPositionSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Position"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const saveOrderPosition = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Position"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const addDepartMentSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  var lastRow = sheet.getLastRow();
  var genCode = genCodeData(obj.sheetname, lastRow)
  const rowData = [parseInt(lastRow) - 1, genCode, obj.depart, obj.doccer, obj.color, obj.number];
  sheet.appendRow(rowData);
  return sheet.getRange("B3:F" + sheet.getLastRow()).getValues();
}
const editDepartMentSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("B3:F" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, 3).setValue(obj.depart);
      sheet.getRange(i + 3, 4).setValue(obj.doccer);
      sheet.getRange(i + 3, 5).setValue(obj.color);
      sheet.getRange(i + 3, 6).setValue(obj.number);
      break;
    }
  }

  return sheet.getRange("B3:F" + sheet.getLastRow()).getValues();
}
const deleteDepartMentSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("B3:F" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.deleteRow(i + 3);
      break;
    }
  }

  return sheet.getRange("B3:F" + sheet.getLastRow()).getValues();  
}
/////////////////////////// DataSet ////////////////////////////

const getAngecySet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Agency"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
  return data;
}

const saveOrderAgency = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Agency"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
}
const addAgencySet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Agency"); 
  const rowData = [obj.key, obj.name, obj.boss];
  sheet.appendRow(rowData);
  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
}
const editAgencySet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Agency"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, 2).setValue(obj.name);
      sheet.getRange(i + 3, 3).setValue(obj.boss);
      break;
    }
  }

  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
}
const deleteAgencySet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Agency"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.deleteRow(i + 3);
      break;
    }
  }

  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();  
}

const getCategorySet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("DocCategory"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getObjectiveSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Objective"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getClassSpeedSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("ClassSpeed"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getClassSecretSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("ClassSecret"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getResponseSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Response"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getJobWordSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("JobWord"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getTypeEmployeeSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("TypeEmployee"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getRevenueSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Revenue"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getExpensesSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("Expenses"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  return data;
}
const getMooBanSet = () => {
  const ss = SpreadsheetApp.openById(sheetSetting);
  const sheet = ss.getSheetByName("MooBan"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
  return data;
}
const saveOrderJobWord = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("JobWord"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderTypeEmployee = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("TypeEmployee"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}

const saveOrderRevenue = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Revenue"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}

const saveOrderExpenses = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Expenses"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}

const saveOrderCategory = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DocCategory"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderClassSpeed = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSpeed"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderClassSecret = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSecret"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderObjective = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Objective"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderResponse = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Response"); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
}
const saveOrderMooBan = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("MooBan"); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  sheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
}
const genCodeData = (...data) => {
  if (data.length < 2) {
    throw new Error('Insufficient data provided');
  }

  const sheetName = data[0];
  const rowNumber = parseInt(data[1]) - 1;
  const paddingSize = 4;

  const prefixMap = {
    DocCategory: 'DC',
    Objective: 'OJ',
    ClassSpeed: 'CS',
    ClassSecret: 'SC',
    Response: 'RS',
    JobWord: 'JW',
    TypeEmployee: 'TE',
    Revenue: 'RV',
    Expenses: 'EP',
    MooBan: 'MB',
    DepartMent: 'DP',
    Position: 'PT',
    Car: 'CAR',
    MeetingRoom: 'MR',
  };

  const prefix = prefixMap[sheetName] || '';
  const number = rowNumber.toString().padStart(paddingSize, '0');

  return prefix + number;
}
const addMeetingRoomSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName(obj.sheetname); 
  const meetingFolder = DriveApp.getFolderById(meetFolder);
  const lastRow = sheet.getLastRow() + 1;
  var genCode = genCodeData(obj.sheetname, lastRow)

  var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
  var blob = Utilities.newBlob(datafile, obj.filetype, genCode);
  var file = meetingFolder.createFile(blob);
  var fileId = file.getId();
  var urlroom = "https://lh3.googleusercontent.com/d/" + fileId;

  const rowData = [genCode, obj.nameroom, obj.address, obj.number, urlroom, true];
  sheet.appendRow(rowData);
  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
}
const editMeetingRoomSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();
  const meetingFolder = DriveApp.getFolderById(meetFolder);

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 2, 2).setValue(obj.nameroom);
      sheet.getRange(i + 2, 3).setValue(obj.address);
      sheet.getRange(i + 2, 4).setValue(obj.number);
      if (obj.check !== "") {
        const files = carFolder.getFilesByName(obj.key);
        while (files.hasNext()) {
          var file = files.next();
          file.setTrashed(true);
        }
        var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
        var blob = Utilities.newBlob(datafile, obj.filetype, obj.key);
        var file = meetingFolder.createFile(blob);
        var fileId = file.getId();
        var urlroom = "https://lh3.googleusercontent.com/d/" + fileId;
        sheet.getRange(i + 2, 5).setValue(urlroom);
      }
      break;
    }
  }

  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
}
const statusMeetingRoomSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName(obj.sheetname); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1 ; i < data.length; i++){
    var userValue = data[i][0];
    if (userValue === obj.key) {
      sheet.getRange(i + 1, 6).setValue(obj.status);
      break; 
    }
  }
  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
}

const addCarSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName(obj.sheetname); 
  const carFolder = DriveApp.getFolderById(carReqFolder);
  const lastRow = sheet.getLastRow() + 1;
  var genCode = genCodeData(obj.sheetname, lastRow)

  var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
  var blob = Utilities.newBlob(datafile, obj.filetype, genCode);
  var file = carFolder.createFile(blob);
  var fileId = file.getId();
  var urlcar = "https://lh3.googleusercontent.com/d/" + fileId;

  const rowData = [genCode, urlcar, obj.car, obj.depart, obj.mile, "50"];
  sheet.appendRow(rowData);
  sheet.getRange(lastRow, 10).setValue('พร้อมใช้งาน');
  sheet.getRange(lastRow, 12).setValue(true);
  return sheet.getRange("A2:L" + sheet.getLastRow()).getValues();
}
const editCarSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();
  const carFolder = DriveApp.getFolderById(carReqFolder);

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 2, 3).setValue(obj.car);
      sheet.getRange(i + 2, 4).setValue(obj.depart);
      sheet.getRange(i + 2, 5).setValue(obj.mile);
      if (obj.check !== "") {
        const files = carFolder.getFilesByName(obj.key);
        while (files.hasNext()) {
          var file = files.next();
          file.setTrashed(true);
        }
        var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
        var blob = Utilities.newBlob(datafile, obj.filetype, obj.key);
        var file = carFolder.createFile(blob);
        var fileId = file.getId();
        var urlcar = "https://lh3.googleusercontent.com/d/" + fileId;
        sheet.getRange(i + 2, 2).setValue(urlcar);
      }
      break;
    }
  }

  return sheet.getRange("A2:L" + sheet.getLastRow()).getValues();
}
const statusCarSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetCarReq).getSheetByName(obj.sheetname); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1 ; i < data.length; i++){
    var userValue = data[i][0];
    if (userValue === obj.key) {
      sheet.getRange(i + 1, 12).setValue(obj.status);
      break; 
    }
  }
  return sheet.getRange("A2:L" + sheet.getLastRow()).getValues();
}
const addMooBanSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  var lastRow = sheet.getLastRow();
  var genCode = genCodeData(obj.sheetname, lastRow)
  const rowData = [genCode, obj.number, obj.ban];
  sheet.appendRow(rowData);
  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
}
const editMooBanSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, 2).setValue(obj.number);
      sheet.getRange(i + 3, 3).setValue(obj.ban);
      break;
    }
  }

  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();
}
const deleteMooBanSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("A3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.deleteRow(i + 3);
      break;
    }
  }

  return sheet.getRange("A3:C" + sheet.getLastRow()).getValues();  
}
const addDataSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  var lastRow = sheet.getLastRow();
  var genCode = genCodeData(obj.sheetname, lastRow)
  const rowData = [parseInt(lastRow) - 1, genCode, obj.list];
  sheet.appendRow(rowData);
  return sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
}
const editDataSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, 3).setValue(obj.list);
      break;
    }
  }

  return sheet.getRange("B3:C" + sheet.getLastRow()).getValues();
}
const deleteDataSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName(obj.sheetname); 
  const data = sheet.getRange("B3:C" + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.deleteRow(i + 3);
      break;
    }
  }

  return sheet.getRange("B3:C" + sheet.getLastRow()).getValues();  
}

//////////////////////////// Setting //////////////////////////////

const getSettingShow = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting"); 
  const data = sheet.getRange("B1:B" + sheet.getLastRow()).getValues();
  return data;
}
const setEditSetting1 = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
  const ranges = ['B1', 'B2'];
  const values = [obj.settingData9, obj.settingData10];

  ranges.forEach((range, index) => {
    sheet.getRange(range).setValue(values[index]);
  });
  clearSystemSettingsCache();
}
const setEditSetting2 = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
  const ranges = ['B7', 'B8', 'B9', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B24', 'B26', 'B28', 'B30', 'B32', 'B33'];
  const values = [obj.settingData1, obj.settingData2, obj.settingData3, obj.settingData4, obj.settingData5, obj.settingData6, obj.settingData7, obj.settingData8, obj.settingData21, obj.settingData22, obj.settingData28, obj.settingData29, obj.settingData30, obj.settingData31, obj.settingData32, obj.settingData33];

  ranges.forEach((range, index) => {
    sheet.getRange(range).setValue(values[index]);
  });
  clearSystemSettingsCache();
}
const setEditSetting3 = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
  const ranges = ['B3', 'B4', 'B5', 'B6', 'B23', 'B25', 'B27', 'B29', 'B31'];
  const values = [obj.settingData11, obj.settingData12, obj.settingData13, obj.settingData14, obj.settingData23, obj.settingData24, obj.settingData25, obj.settingData26, obj.settingData27];

  ranges.forEach((range, index) => {
    sheet.getRange(range).setValue(values[index]);
  });
  clearSystemSettingsCache();
}
const setEditSetting4 = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
  const ranges = ['B10', 'B11', 'B12', 'B13', 'B14', 'B15'];
  const values = [obj.settingData15, obj.settingData16, obj.settingData17, obj.settingData18, obj.settingData19, obj.settingData20];

  ranges.forEach((range, index) => {
    sheet.getRange(range).setValue(values[index]);
  });
  clearSystemSettingsCache();
}

////////////////////// Meeting Room /////////////////////////
const generateMeetKey = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'MTR';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}
const getMeetingRoomAll = () => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName('MeetingRoom');
  const data = sheet.getDataRange().getValues().slice(1);
  return [data];
}
const getMeetingRoom = () => {
  const sheet = SpreadsheetApp.openById(sheetMeet)
  const values = sheet.getSheetByName("MeetingRoom");
  const data = values.getDataRange().getValues().slice(1);

  return [data];
}

const addMeetingSubmit = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName("DataMeeting");
  const code = generateMeetKey();
  var userSender = obj.addMeetingData5.split(', ');
  const rowData = [code, obj.addMeetingData1, "'" + obj.addMeetingData2, obj.addMeetingData3, obj.addMeetingData4, userSender[1], userSender[0]]
  sheet.appendRow(rowData);
  return code;
}

const addUserMeeting = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName("DataMeeting");
  const notiSheet = SpreadsheetApp.openById(sheetNotify).getSheetByName("Notify");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    const keyCode = data[i][0];
    if (keyCode === obj.booking) {
      sheet.getRange(i + 1, 8).setValue(obj.publicKey);
      sheet.getRange(i + 1, 9).setValue(obj.fullname);
      sheet.getRange(i + 1, 10).setValue(obj.token);
      sheet.getRange(i + 1, 11).setValue("0");
      sheet.getRange(i + 1, 12).setValue(obj.image);
      const genCode = generateRandomKey();
      notiSheet.appendRow([genCode, data[i][6], obj.publicKey, "ขอจอง" + data[i][1], "ในวันที่" + formatToDateThaiFull(new Date(data[i][2])), "changePageMeeting", "1", "0", "'" + formatDateSave(new Date())]);
      break;
    }
  }
}
const getMeetingDataAll = () => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName('DataMeeting');
  const data = sheet.getDataRange().getValues().slice(1);
  const dateNow = new Date().toISOString().split('T')[0]; 
  var filter = data.filter(function(row) {
    const meetingDate = new Date(row[2]);
    return meetingDate.toISOString().split('T')[0] >= dateNow;
  });
  return filter;
}
const allowMeetChoose = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetMeet).getSheetByName("DataMeeting");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    const keyCode = data[i][0];
    if (keyCode === obj.dataMeetingroom1) {
      sheet.getRange(i + 1, 11).setValue(obj.dataMeetingroom2);
      sheet.getRange(i + 1, 13).setValue(obj.dataMeetingroom3);
      break;
    }
  }
}


/////////////////////////////////// License ////////////////////////////////
const getDataLicense = () => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("License");
  const data = sheet.getDataRange().getValues().slice(2);
  return data;
}

const getLicenseData = (key) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("SideBar");
  const data = sheet.getDataRange().getValues().slice(2);
  const column = parseInt(key) + 2

  var filter = data.map(function(row) {
    return [row[0], row[1], row[2], row[column]];
  });

  return filter;
}
const saveLicenseStatus = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("SideBar");
  const data = sheet.getDataRange().getValues().slice(2);
  const column = parseInt(obj.column) + 3;

  for (let i = 0; i <= data.length; i++) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, column).setValue(obj.status);
      break;
    }
  }
}
const saveFileLicense = (e, fileNumber) => {
  let folder = DriveApp.getFolderById(userFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("License"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === e.code) {
      sheet.getRange(i + 1, fileNumber).setValue(url);
      break;
    }
  }
  
  return 'อัปโหลดไฟล์เรียบร้อย';
}
const licenseEdit = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("License"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i <= data.length; i++) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 1, 3).setValue(obj.namePosition);
      sheet.getRange(i + 1, 4).setValue(obj.workPosition);
      break;
    }
  }
};


const setting_notify = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
  const ranges = ['B34', 'B35', 'B36'];
  const values = [obj.setting_system_7, obj.setting_system_8];

  ranges.forEach((range, index) => {
    sheet.getRange(range).setValue(values[index]);
  });
};