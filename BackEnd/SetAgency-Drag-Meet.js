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
  try {
    CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA');
  } catch (e) {}
}

const updateAgencyBackgroundUrl = (bgUrl) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Institute");
  sheet.getRange('C12').setValue(bgUrl);
  try {
    CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA');
  } catch (e) {}
  return bgUrl;
}
function getCarSet() {
  try {
    const ss = SpreadsheetApp.openById(sheetCarReq);
    if (!ss) return [];
    const sheet = ss.getSheetByName("Car"); 
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch (e) {
    return [];
  }
}
const saveOrderCar = (order) => {
  try {
    const ss = SpreadsheetApp.openById(sheetCarReq);
    if (!ss) return;
    const sheet = ss.getSheetByName("Car"); 
    if (!sheet) return;
    const data = sheet.getDataRange().getValues().slice(1);
    var newData = [];

    order.forEach(function(id) {
      data.forEach(function(row) {
        if (row[0] === id) {
          newData.push(row);
        }
      });
    });

    if (newData.length > 0) {
      sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
  } catch (e) {}
}
function getMeetingRoomSet() {
  try {
    const ss = SpreadsheetApp.openById(sheetMeet);
    if (!ss) return [];
    const sheet = ss.getSheetByName("MeetingRoom"); 
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch (e) {
    return [];
  }
}
const saveOrderMeetingRoom = (order) => {
  try {
    const ss = SpreadsheetApp.openById(sheetMeet);
    if (!ss) return;
    const sheet = ss.getSheetByName("MeetingRoom"); 
    if (!sheet) return;
    const data = sheet.getDataRange().getValues().slice(1);
    var newData = [];

    order.forEach(function(id) {
      data.forEach(function(row) {
        if (row[0] === id) {
          newData.push(row);
        }
      });
    });

    if (newData.length > 0) {
      sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
  } catch (e) {}
}
function getDepartMentSet() {
  return _safeGetRangeValues("DepartMent", "B", "F");
}
const saveOrderDepartMent = (order) => {
  try {
    const ss = _getSettingSS();
    if (!ss) return;
    const sheet = ss.getSheetByName("DepartMent"); 
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const data = sheet.getRange("B3:F" + lastRow).getValues();
    var newData = [];

    order.forEach(function(id) {
      data.forEach(function(row) {
        if (row[0] === id) {
          newData.push(row);
        }
      });
    });

    if (newData.length > 0) {
      sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
    }
  } catch (e) {}
}
function getPositionSet() {
  return _safeGetRangeValues("Position", "B", "C");
}
const saveOrderPosition = (order) => {
  try {
    const ss = _getSettingSS();
    if (!ss) return;
    const sheet = ss.getSheetByName("Position"); 
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const data = sheet.getRange("B3:C" + lastRow).getValues();
    var newData = [];

    order.forEach(function(id) {
      data.forEach(function(row) {
        if (row[0] === id) {
          newData.push(row);
        }
      });
    });

    if (newData.length > 0) {
      sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
    }
  } catch (e) {}
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

function _safeGetRangeValues(sheetName, startCol, endCol) {
  try {
    const ss = _getSettingSS();
    if (!ss) return [];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    return sheet.getRange(startCol + "3:" + endCol + lastRow).getValues();
  } catch (err) {
    Logger.log("Error reading sheet " + sheetName + ": " + err);
    return [];
  }
}

function getAngecySet() {
  return _safeGetRangeValues("Agency", "A", "C");
}

const saveOrderAgency = (order) => {
  const ss = _getSettingSS();
  const sheet = ss ? ss.getSheetByName("Agency") : null; 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("A3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
  }
}
const addAgencySet = (obj) => {
  const ss = _getSettingSS();
  const sheet = ss ? ss.getSheetByName("Agency") : null; 
  if (!sheet) return [];
  const rowData = [obj.key, obj.name, obj.boss];
  sheet.appendRow(rowData);
  return _safeGetRangeValues("Agency", "A", "C");
}
const editAgencySet = (obj) => {
  const ss = _getSettingSS();
  const sheet = ss ? ss.getSheetByName("Agency") : null; 
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  const data = sheet.getRange("A3:C" + lastRow).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 3, 2).setValue(obj.name);
      sheet.getRange(i + 3, 3).setValue(obj.boss);
      break;
    }
  }

  return _safeGetRangeValues("Agency", "A", "C");
}
const deleteAgencySet = (obj) => {
  const ss = _getSettingSS();
  const sheet = ss ? ss.getSheetByName("Agency") : null; 
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  const data = sheet.getRange("A3:C" + lastRow).getValues();

  for (let i = 0; i < data.length; i++) {
    var keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.deleteRow(i + 3);
      break;
    }
  }

  return _safeGetRangeValues("Agency", "A", "C");  
}

function getCategorySet() {
  return _safeGetRangeValues("DocCategory", "B", "C");
}
function getObjectiveSet() {
  return _safeGetRangeValues("Objective", "B", "C");
}
function getClassSpeedSet() {
  return _safeGetRangeValues("ClassSpeed", "B", "C");
}
function getClassSecretSet() {
  return _safeGetRangeValues("ClassSecret", "B", "C");
}
function getResponseSet() {
  return _safeGetRangeValues("Response", "B", "C");
}
function getJobWordSet() {
  return _safeGetRangeValues("JobWord", "B", "C");
}
function getTypeEmployeeSet() {
  return _safeGetRangeValues("TypeEmployee", "B", "C");
}
function getRevenueSet() {
  return _safeGetRangeValues("Revenue", "B", "C");
}
function getExpensesSet() {
  return _safeGetRangeValues("Expenses", "B", "C");
}
function getMooBanSet() {
  return _safeGetRangeValues("MooBan", "A", "C");
}

const _ensureAndGetDataSet = (sheetName, defaultItems) => {
  const ss = _getSettingSS();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['#', 'รหัส', 'รายการ']);
    sheet.appendRow(['', '', '']);
    if (Array.isArray(defaultItems) && defaultItems.length > 0) {
      const rows = defaultItems.map((item, idx) => {
        const rowNum = idx + 1;
        const code = genCodeData(sheetName, rowNum + 1);
        return [rowNum, code, item];
      });
      sheet.getRange(3, 1, rows.length, 3).setValues(rows);
    }
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange("B3:C" + lastRow).getValues();
};

function getEquipCategorySet() {
  return _ensureAndGetDataSet("EquipCategory", [
    'ครุภัณฑ์สำนักงาน', 'ครุภัณฑ์คอมพิวเตอร์', 'ครุภัณฑ์ยานพาหนะและขนส่ง', 
    'ครุภัณฑ์การเกษตร', 'ครุภัณฑ์โฆษณาและเผยแพร่', 'ครุภัณฑ์งานบ้านงานครัว', 
    'ครุภัณฑ์วิทยาศาสตร์และการแพทย์', 'ครุภัณฑ์สำรวจ', 'ครุภัณฑ์ดนตรีและนาฏศิลป์', 'ครุภัณฑ์อื่น ๆ'
  ]);
}

function getEquipTypeSet() {
  return _ensureAndGetDataSet("EquipType", [
    'เครื่องปรับอากาศ', 'คอมพิวเตอร์และอุปกรณ์', 'เครื่องพิมพ์/สแกนเนอร์', 'โต๊ะ/เก้าอี้สำนักงาน', 
    'ตู้เก็บเอกสาร', 'กล้องถ่ายภาพ/วิดีโอ', 'โทรทัศน์/จอภาพ', 'ยานพาหนะ', 'เครื่องสำรองไฟฟ้า', 'อุปกรณ์เครือข่าย'
  ]);
}

function getEquipStatusSet() {
  return _ensureAndGetDataSet("EquipStatus", [
    'พร้อมใช้งาน', 'กำลังใช้งาน', 'ชำรุด', 'รอซ่อม', 'รอจำหน่าย', 'ตัดจำหน่าย'
  ]);
}

function getEquipAcquireSet() {
  return _ensureAndGetDataSet("EquipAcquire", [
    'ตกลงราคา', 'สอบราคา', 'ประกวดราคา', 'วิธีเฉพาะเจาะจง', 'คัดเลือก', 'รับบริจาค'
  ]);
}

function getEquipExpenseSet() {
  return _ensureAndGetDataSet("EquipExpense", [
    'งบลงทุน', 'งบดำเนินงาน', 'เงินอุดหนุน', 'เงินนอกงบประมาณ'
  ]);
}

function getEquipVendorSet() {
  return _ensureAndGetDataSet("EquipVendor", [
    'ห้างหุ้นส่วนจำกัด/บริษัททั่วไป', 'ร้านค้าท้องถิ่น', 'ผู้จัดจำหน่ายส่วนกลาง', 'ผู้บริจาค'
  ]);
}

function getMatCategorySet() {
  return _ensureAndGetDataSet("MatCategory", [
    'วัสดุสำนักงาน', 'วัสดุคอมพิวเตอร์', 'วัสดุไฟฟ้าและวิทยุ', 'วัสดุงานบ้านงานครัว',
    'วัสดุยานพาหนะและขนส่ง', 'วัสดุการเกษตร', 'วัสดุก่อสร้าง', 'วัสดุอื่น ๆ'
  ]);
}

function getMatTypeSet() {
  return _ensureAndGetDataSet("MatType", [
    'กระดาษและสิ่งพิมพ์', 'เครื่องเขียนและอุปกรณ์', 'หมึกพิมพ์และโทนเนอร์', 'อุปกรณ์บันทึกข้อมูล',
    'อุปกรณ์ทำความสะอาด', 'หลอดไฟและอุปกรณ์ไฟฟ้า', 'อะไหล่และอุปกรณ์ซ่อมบำรุง'
  ]);
}

function getEquipUnitSet() {
  return _ensureAndGetDataSet("EquipUnit", [
    'เครื่อง', 'ชุด', 'คัน', 'หลัง', 'ตัว', 'ชิ้น', 'อัน', 'รีม', 'กล่อง', 'เล่ม', 'แพ็ค', 'ม้วน', 'แผ่น'
  ]);
}

function getEquipLocationSet() {
  return _ensureAndGetDataSet("EquipLocation", [
    'ห้องพัสดุกลาง', 'ห้องสำนักงานปลัด', 'ห้องกองคลัง', 'ห้องกองช่าง', 'ห้องประชุมสภา', 'ห้องศูนย์ข้อมูล/เซิร์ฟเวอร์', 'อาคารจอดรถ'
  ]);
}

const saveOrderJobWord = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("JobWord"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderTypeEmployee = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("TypeEmployee"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}

const saveOrderRevenue = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Revenue"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}

const saveOrderExpenses = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Expenses"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}

const saveOrderCategory = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("DocCategory"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderClassSpeed = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSpeed"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderClassSecret = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("ClassSecret"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderObjective = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Objective"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderResponse = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("Response"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
  }
}
const saveOrderMooBan = (order) => {
  const sheet = SpreadsheetApp.openById(sheetSetting).getSheetByName("MooBan"); 
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("A3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
  }
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
    EquipCategory: 'EQC',
    EquipType: 'EQT',
    EquipStatus: 'EQS',
    EquipAcquire: 'EQA',
    EquipExpense: 'EQE',
    EquipVendor: 'EQV',
    MatCategory: 'MTC',
    MatType: 'MTT',
    EquipUnit: 'UNT',
    EquipLocation: 'LOC',
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
  const ss = _getSettingSS();
  let sheet = ss.getSheetByName(obj.sheetname);
  if (!sheet) {
    sheet = ss.insertSheet(obj.sheetname);
    sheet.appendRow(['#', 'รหัส', 'รายการ']);
    sheet.appendRow(['', '', '']);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.appendRow(['#', 'รหัส', 'รายการ']);
    sheet.appendRow(['', '', '']);
    lastRow = 2;
  }
  var genCode = genCodeData(obj.sheetname, lastRow + 1);
  const rowData = [parseInt(lastRow) - 1, genCode, obj.list];
  sheet.appendRow(rowData);
  try { CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA'); } catch (e) {}
  return _safeGetRangeValues(obj.sheetname, 2, 3);
}

const editDataSet = (obj) => {
  const sheet = _getSettingSS().getSheetByName(obj.sheetname); 
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    const data = sheet.getRange("B3:C" + lastRow).getValues();

    for (let i = 0; i < data.length; i++) {
      var keyValue = data[i][0];
      if (keyValue === obj.key) {
        sheet.getRange(i + 3, 3).setValue(obj.list);
        break;
      }
    }
  }
  try { CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA'); } catch (e) {}
  return _safeGetRangeValues(obj.sheetname, 2, 3);
}

const deleteDataSet = (obj) => {
  const sheet = _getSettingSS().getSheetByName(obj.sheetname); 
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    const data = sheet.getRange("B3:C" + lastRow).getValues();

    for (let i = 0; i < data.length; i++) {
      var keyValue = data[i][0];
      if (keyValue === obj.key) {
        sheet.deleteRow(i + 3);
        break;
      }
    }
  }
  try { CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA'); } catch (e) {}
  return _safeGetRangeValues(obj.sheetname, 2, 3);
}

const saveOrderDataSet = (sheetName, order) => {
  const sheet = _getSettingSS().getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const data = sheet.getRange("B3:C" + lastRow).getValues();
  var newData = [];

  order.forEach(function(id) {
    data.forEach(function(row) {
      if (row[0] === id) {
        newData.push(row);
      }
    });
  });

  if (newData.length > 0) {
    sheet.getRange(3, 2, newData.length, newData[0].length).setValues(newData);
    try { CacheService.getScriptCache().remove('ALL_DROPDOWN_DATA'); } catch (e) {}
  }
};

function saveOrderEquipCategory(order) { return saveOrderDataSet('EquipCategory', order); }
function saveOrderEquipType(order) { return saveOrderDataSet('EquipType', order); }
function saveOrderEquipStatus(order) { return saveOrderDataSet('EquipStatus', order); }
function saveOrderEquipAcquire(order) { return saveOrderDataSet('EquipAcquire', order); }
function saveOrderEquipExpense(order) { return saveOrderDataSet('EquipExpense', order); }
function saveOrderEquipVendor(order) { return saveOrderDataSet('EquipVendor', order); }
function saveOrderMatCategory(order) { return saveOrderDataSet('MatCategory', order); }
function saveOrderMatType(order) { return saveOrderDataSet('MatType', order); }
function saveOrderEquipUnit(order) { return saveOrderDataSet('EquipUnit', order); }
function saveOrderEquipLocation(order) { return saveOrderDataSet('EquipLocation', order); }


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
  const ss = SpreadsheetApp.openById(sheetMeet);
  const roomSheet = ss.getSheetByName("MeetingRoom");
  const meetSheet = ss.getSheetByName("DataMeeting");

  const rooms = roomSheet ? roomSheet.getDataRange().getValues().slice(1) : [];
  const bookings = meetSheet ? meetSheet.getDataRange().getValues().slice(1) : [];

  return [rooms, bookings];
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