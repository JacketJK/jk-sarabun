const getAlldataHrm = () => {
  const ss = SpreadsheetApp.openById(sheetHrm);
  const sheet = ss.getSheetByName("hrmBase"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  var filteredData = data.filter(function(row) {
    return row[10] === "ปฏิบัติราชการอยู่";
  });

  return filteredData;
}
const getAlldataHrmStatus = () => {
  const ss = SpreadsheetApp.openById(sheetHrm);
  const sheet = ss.getSheetByName("hrmBase"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data;
}
const editDataEmployee = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName('hrmBase');
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 1, 2).setValue(obj.data1);
      sheet.getRange(i + 1, 3).setValue(obj.data2);
      sheet.getRange(i + 1, 4).setValue(obj.data3);
      sheet.getRange(i + 1, 5).setValue(obj.data4);
      sheet.getRange(i + 1, 6).setValue(obj.data5);
      sheet.getRange(i + 1, 7).setValue(obj.data6);
      sheet.getRange(i + 1, 8).setValue("'" + obj.data7);
      sheet.getRange(i + 1, 9).setValue("'" + obj.data8);
      sheet.getRange(i + 1, 12).setValue(formatToDateThai(new Date()));
      break;
    }
  }
}
const editFileProfile = (e) => {
  let folder = DriveApp.getFolderById(hrmFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][1];  
    if (userValue === e.code) {
      sheet.getRange(i + 1, 10).setValue(url);
      sheet.getRange(i + 1, 12).setValue(formatToDateThai(new Date()));
      break;
    }
  }
  
  return url;
}
const statusHrm  = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === obj.key) {
      sheet.getRange(i + 1, 11).setValue(obj.status);
      sheet.getRange(i + 1, 12).setValue(formatToDateThai(new Date()));
      break;
    }
  }
}
const updataSetCardSave = (obj) => {
  const sheetSet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SetCard"); 
  const folder = DriveApp.getFolderById(hrmFolder);

  if (obj.setCard_data3 !== "") {
    const fileSig = folder.createFile(obj.setCard_data3);
    const sigID = fileSig.getId();
    const sigName = fileSig.getName();
    const sigUrl = "https://lh3.googleusercontent.com/d/" + sigID;
    if (sigName !== "Untitled") {
      sheetSet.getRange('B3').setValue(sigUrl);
    }
  } else {
    Logger.log("setCard_data3 is empty");
  }

  if (obj.setCard_data5 !== "") {
    const fileSael = folder.createFile(obj.setCard_data5);
    const saelID = fileSael.getId();
    const saelName = fileSael.getName();
    const saelUrl = "https://lh3.googleusercontent.com/d/" + saelID;
    if (saelName !== "Untitled") {
      sheetSet.getRange('B5').setValue(saelUrl);
    }
  } else {
    Logger.log("setCard_data5 is empty");
  }

  sheetSet.getRange('B1').setValue(obj.setCard_data1 || "");
  sheetSet.getRange('B2').setValue(obj.setCard_data2 || "");
  sheetSet.getRange('B4').setValue(obj.setCard_data4 || "");
};

const edit_CardEmployeeSet = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  const sheetSet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SetCard"); 
  const data = sheet.getDataRange().getValues();
  const dataSet = sheetSet.getRange('B4').getValue(); 
  const runNumber = parseInt(dataSet) + 1;

  for (let i = 0; i < data.length; i++) {
    const userValue = data[i][0];  
    if (userValue === obj.key) {
      sheet.getRange(i + 1, 3).setValue(obj.data1);
      sheet.getRange(i + 1, 4).setValue(obj.data2);
      sheet.getRange(i + 1, 5).setValue(obj.data3);
      sheet.getRange(i + 1, 7).setValue(obj.data4);
      sheet.getRange(i + 1, 13).setValue(obj.data5);
      sheet.getRange(i + 1, 14).setValue(obj.data6);
      sheet.getRange(i + 1, 15).setValue(obj.data7);
      sheet.getRange(i + 1, 12).setValue(formatToDateThai(new Date())); 
      sheetSet.getRange('B4').setValue(runNumber);
      break;
    }
  }  
}
const generateHrm = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'HRM';
  let key = prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}
const add_newEmployee = (obj) => {
  let folder = DriveApp.getFolderById(hrmFolder);
  var datafile1 = Utilities.base64Decode(obj.imageDataUrl1.split(',')[1]);
  var datafile2 = Utilities.base64Decode(obj.imageDataUrl2.split(',')[1]);
  var blob1 = Utilities.newBlob(datafile1, obj.filetype, "IMG_" + obj.data1);
  var blob2 = Utilities.newBlob(datafile2, obj.filetype, "SIG_" + obj.data1);
  var file1 = folder.createFile(blob1);
  var file2 = folder.createFile(blob2);
  var fileId1 = file1.getId();
  var fileId2 = file2.getId();
  var profile_Url = "https://lh3.googleusercontent.com/d/" + fileId1;
  var signeture_Url = "https://lh3.googleusercontent.com/d/" + fileId2;

  const code = generateHrm();

  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase"); 
  var rowData = [code, obj.data1, obj.data2, obj.data3, obj.data4, obj.data5, obj.data6, "'" + obj.data7, "'" + obj.data8, profile_Url, "ปฏิบัติราชการอยู่", formatToDateThai(new Date()), , , , signeture_Url];

  sheet.appendRow(rowData); 
}
const setCardEmployee = () => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SetCard"); 
  var data = sheet.getRange('B1:B' + sheet.getLastRow()).getValues();
  return data;
}

//////////////////////////////////// SaveLand /////////////////////////////////

const getDataSaveLandDB = () => {
  var sheet = SpreadsheetApp.openById(sheetLand).getSheetByName("Project"); 
  var data = sheet.getRange("B2:C" + sheet.getLastRow()).getValues(); 
  var obj = {};

  data.forEach(function([colA, colB]) {
    if (!obj[colA]) {
      obj[colA] = [];
    }

    if (obj[colA].indexOf(colB) === -1) {
      obj[colA].push(colB);
    }
  });

  var sortedKeys = Object.keys(obj).sort((a, b) => b - a);
  var sortedObj = {};
  sortedKeys.forEach(key => {
    sortedObj[key] = obj[key];
  });

  return sortedObj;
}

const getDataTableSaveland = (obj)=> {
  var dataArray = SpreadsheetApp.openById(sheetLand).getSheetByName("Project").getDataRange().getDisplayValues();
  dataArray.shift();

  var selectedYear = obj.dropdown1;
  var selectedProject = obj.dropdown2;
  var landNumber = obj.dropdown3;

  var output = dataArray.filter(function (row) {
    if (selectedProject === "" && landNumber === "") {
      return (row[1] == selectedYear);
    } else if (selectedProject !== "" && landNumber === "") {
      return (row[1] == selectedYear &&
              row[2] == selectedProject );
    } else if (selectedProject === "" && selectedYear === "") {
      return ( row[3] == landNumber );
    } else {
      return (
        row[1] == selectedYear &&
        row[2] == selectedProject &&
        row[3] == landNumber
      );
    }
  });

  return output;
}
const generateSLand = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'SL';
  let key = prefix; 
  for (let i = 0; i < 8; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }

  return key;
}
const addSavelandgs = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetLand).getSheetByName("Project"); 
  const folder = DriveApp.getFolderById(landFolder);
  const pdf_Id = folder.createFile(obj.addNewSaveLand12).getId()
  const pdf_url = "https://drive.google.com/open?id=" + pdf_Id;
  const codeKey = generateSLand();
  const address = `บ้านเลขที่ ${obj.addNewSaveLand6} หมู่ที่ ${obj.addNewSaveLand7} ตำบล${obj.addNewSaveLand8} อำเภอ${obj.addNewSaveLand9} จังหวัด${obj.addNewSaveLand10} รหัสไปรษณีย์ ${obj.addNewSaveLand11}`

  const rowData = [codeKey, obj.addNewSaveLand1, obj.addNewSaveLand2, obj.addNewSaveLand3, obj.addNewSaveLand4, obj.addNewSaveLand5, address, pdf_url, formatToDateThai(new Date())];
  sheet.appendRow(rowData);
}

////////////////////// Check Bin //////////////////////
const getDataCheckBin = () => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Maps");
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
}
const getDataTrashCompian = () => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Compian");
  const data = sheet.getDataRange().getValues().slice(1);
  var dataFilter = data.filter(function(row) {
    return row[8] == false;
  });
  
  return dataFilter;
}
const editTrashPosition = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Maps");
  const data = sheet.getDataRange().getValues().slice(1);

  const updatedData = data.map((row) => {
    const keyValue = row[0];
    if (keyValue === obj.key) {
      row[1] = obj.lat;
      row[2] = obj.lon;
      row[11] = obj.lat + "," + obj.lon;
    }
    return row;
  });

  sheet.getRange(2, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
  
  return updatedData;
}
const editTrashData = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Maps"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const keyValue = data[i][0];
    const userValue = data[i][4];
    if (keyValue === obj.key && userValue === obj.data1) {
      sheet.getRange(i + 1, 6).setValue(obj.data2);
      sheet.getRange(i + 1, 7).setValue(obj.data3);
      sheet.getRange(i + 1, 13).setValue(obj.data4);
      break;
    }
  }
}
const saveFileTrash = (e, fileNumber) => {
  let folder = DriveApp.getFolderById(binFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Maps"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const keyValue = data[i][0];
    const userValue = data[i][4];
    if (keyValue === e.code && userValue === e.owner) {
      sheet.getRange(i + 1, fileNumber).setValue(url);
      break;
    }
  }
  
  return 'อัปโหลดไฟล์เรียบร้อย';
}
const getfeeTrashData = (...filter) => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Maps");
  const sheetfee = SpreadsheetApp.openById(sheetBin).getSheetByName("FeeTrash");
  const data = sheet.getDataRange().getValues().slice(1);
  const datafee = sheetfee.getDataRange().getValues().slice(1);
  const year = filter[0];
  const month = filter[1];
  const uniqueValues = new Set();
  const duplicateCount = [];

  data.forEach(row => {
    uniqueValues.add(row[4]);
    const index = duplicateCount.findIndex(entry => entry[0] === row[4]);
    if (index === -1) {
      duplicateCount.push([row[4], row[0], 1]);
    } else {
      duplicateCount[index][1] += `- ${row[0]}`;
      duplicateCount[index][2]++;
    }
  });

  const uniqueData = [];
  uniqueValues.forEach(value => {
    const filteredRow = data.find(row => row[4] === value);
    uniqueData.push(filteredRow);
  });

  datafee.forEach(row => {
    const index = duplicateCount.findIndex(entry => entry[0] === row[2]);
    if (row[0] == year && row[1] === month) {
      duplicateCount[index].push(row[4], row[5], row[6], row[7], row[8]);
    }
  });

  return [uniqueData, duplicateCount];
}

const feeTrashData = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("FeeTrash");
  const rowData = [obj.data3, obj.data4, obj.data1, obj.data2, obj.data5, obj.data6, obj.data7, formatToDateThaiFull(new Date()), obj.user];
  sheet.appendRow(rowData);
  const runConut = parseInt(obj.data6) + 1;

  return runConut;
}
const saveFileTrashCompian = (e) => {
  let folder = DriveApp.getFolderById(binFolder);
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  var fileUrl = folder.createFile(blob).getId();
  var url = "https://lh3.googleusercontent.com/d/" + fileUrl;

  const sheet = SpreadsheetApp.openById(sheetBin).getSheetByName("Compian"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const keyValue = data[i][0];
    if (keyValue === e.code) {
      sheet.getRange(i + 1, 10).setValue(url);
      break;
    }
  }
  
  return 'อัปโหลดไฟล์เรียบร้อย';
};
const compian_success = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetEquip).getSheetByName("Compian"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const keyValue = data[i][0];
    if (keyValue === obj.key) {
      sheet.getRange(i + 1, 9).setValue(true);
      sheet.getRange(i + 1, 11).setValue(obj.user);
      sheet.getRange(i + 1, 12).setValue(formatToDateThaiFull(new Date()));
      break;
    }
  }
};

const getContentCategoty = () => {
  if (typeof getEquipmentSummary === 'function') {
    return getEquipmentSummary();
  }
  const sheet = SpreadsheetApp.openById(sheetEquip).getSheetByName("Category"); 
  const data = sheet.getDataRange().getValues().slice(1);
  return data;
};

const cloundChat = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetChat).getSheetByName("Chat"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    const chatRoomValue = data[i][0];
    if (chatRoomValue == obj.roomId) {
      sheet.getRange(i + 1, 2).setValue(obj.msg);
      sheet.getRange(i + 1, 3).setValue("0");
      sheet.getRange(i + 1, 4).setValue(new Date()) ;
      break;
    }
  }
  return;
};

const createPrivateChat = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetChat).getSheetByName("Chat");
  const data = sheet.getDataRange().getValues();
  let checkDuplicate = false;

  const roomId = parseInt(obj.uidoutChat) + parseInt(obj.uidinChat);
  const member = obj.keyOutChat + ", " + obj.keyInChat;

  for (let i = 0; i < data.length; i++) {
    const chatRoomValue = data[i][0];
    if (chatRoomValue == roomId) {     
      checkDuplicate = true;
      break; 
    }
  }

  if (!checkDuplicate) {
    sheet.appendRow([roomId, "", 0, new Date(), member]);
  }
};

const getDataChat = (publicKey) => {
  const sheet = SpreadsheetApp.openById(sheetChat).getSheetByName("Chat");
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues().slice(2);
  const dataUser = sheetUser.getDataRange().getValues().slice(1); 

  let chatRooms = [];
  let users = new Set();

  data.forEach(row => {
    const members = row[4].split(", "); 
    if (members.includes(publicKey)) {
      chatRooms.push({
        roomId: row[0],
        lastMessage: row[1],
        status: row[2],
        timestamp: row[3].toString(),
      });

      members.forEach(member => users.add(member));
    }
  });

  let userDetails = [];
  users.forEach(memberKey => {
    if (memberKey !== publicKey) { 
      const user = dataUser.find(userRow => userRow[0] === memberKey); 
      if (user) {
        userDetails.push({
          publicKey: user[0],
          username: user[1], 
          fullname: user[3] + user[4] + " " + user[5], 
          profileUrl: user[13] 
        });
      }
    }
  });

  return {
    chatRooms: chatRooms,
    users: userDetails
  };
};

// system esalary
const getAllEmployee = () => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("hrmBase");
  const data = sheet.getDataRange().getValues();

  const employees = data.slice(1).map(row => ({
    uid: row[0],
    name: row[2] + row[3] + " " + row[4],
    type: row[5],
    position: row[6], 
    agency: row[16], 
  }));

  return employees;
};

function saveSalaryData(data) {
  try {
    const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SlipBase");
    const key = generateRandomKey();

    const dataRange = sheet.getDataRange().getValues();
    const uidIndex = 1; 
    const roundIndex = 2; 
    const keyIndex = 0; 

    if (data.key === "") { // กรณีเพิ่มข้อมูลใหม่
      const isDuplicate = dataRange.some(row => row[uidIndex] === data.uid && row[roundIndex] === data.round);
      if (isDuplicate) {
        return { success: false, error: "มีการบันทึกข้อมูลนี้ไปแล้ว" };
      }

      const revenuesJson = JSON.stringify(data.revenues);
      const expensessJson = JSON.stringify(data.expensess);

      const row = [
        key,
        data.uid,
        "'" + data.round,
        revenuesJson,
        expensessJson,
        data.revenuesCount,
        data.expensessCount,
        data.total,
        "'" + new Date()
      ];
      sheet.appendRow(row);

      return { success: true };
    } else { // กรณีอัปเดตข้อมูลเดิม
      const rowIndex = dataRange.findIndex(row => row[keyIndex] === data.key);

      if (rowIndex === -1) {
        return { success: false, error: "ไม่พบข้อมูลที่ต้องการอัปเดต" };
      }

      const revenuesJson = JSON.stringify(data.revenues);
      const expensessJson = JSON.stringify(data.expensess);

      sheet.getRange(rowIndex + 1, 2).setValue(data.uid); 
      sheet.getRange(rowIndex + 1, 3).setValue("'" + data.round); 
      sheet.getRange(rowIndex + 1, 4).setValue(revenuesJson); 
      sheet.getRange(rowIndex + 1, 5).setValue(expensessJson); 
      sheet.getRange(rowIndex + 1, 6).setValue(data.revenuesCount); 
      sheet.getRange(rowIndex + 1, 7).setValue(data.expensessCount); 
      sheet.getRange(rowIndex + 1, 8).setValue(data.total); 
      sheet.getRange(rowIndex + 1, 9).setValue("'" + new Date()); 

      return { success: true };
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
}

const getDataSlipBase = (uid) => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SlipBase");
  const data = sheet.getDataRange().getValues();

  const filteredData = data.filter(row => row[1] === uid);

  return filteredData;
};

const getDataAllSlipBase = () => {
  const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SlipBase");
  const sheetUser = SpreadsheetApp.openById(sheetUseds).getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  const dataUser = sheetUser.getDataRange().getValues().slice(1);

  const userMap = new Map(dataUser.map(row => [row[0], row])); 
  
  const result = data.map(row => {
    const key = row[1]; 
    const userRow = userMap.get(key); 
    return userRow ? [...row, ...userRow] : [...row, 'ไม่พบข้อมูล']; 
  });

  return result; 
};

const addDataSalaryFileIncome = (resultIncome, resultExpenses) => {
  try {
    const sheet = SpreadsheetApp.openById(sheetHrm).getSheetByName("SlipBase");

    resultIncome.forEach((element, index) => {
      const key = generateRandomKey(); 
      const detailsIncome = JSON.stringify(element.details || []);
      const detailsExpense = JSON.stringify(resultExpenses[index]?.details || []);
      const income = resultExpenses[index]?.income || "";
      const expenses = resultExpenses[index]?.expenses || "";
      const total = resultExpenses[index]?.total || "";
      const now = new Date();

      sheet.appendRow([
        key,
        element.uid || "",
        "'" + element.dateround,
        detailsIncome, 
        detailsExpense, 
        income || "0", 
        expenses || "0", 
        total || "0", 
        "'" + now 
      ]);
    });

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message }; 
  }
};
