/**
 * BackEnd/Equipment.js
 * ระบบบริหารจัดการวัสดุและครุภัณฑ์ (Equipment, Supplies & Borrow-Return System)
 */

function _getEquipSS() {
  if (typeof sheetEquip !== 'undefined' && sheetEquip && sheetEquip.trim() !== '') {
    try {
      return SpreadsheetApp.openById(sheetEquip);
    } catch (e) {
      Logger.log('Could not open sheetEquip, fallback to active spreadsheet: ' + e);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}


function initEquipmentDatabase() {
  const ss = _getEquipSS();
  
  // 1. Sheet Equipment
  let eqSheet = ss.getSheetByName('Equipment');
  if (!eqSheet) {
    eqSheet = ss.insertSheet('Equipment');
    eqSheet.appendRow([
      'Key', 'หมายเลขครุภัณฑ์', 'รหัส GFMIS', 'หน่วยงาน', 'หมวดหมู่ครุภัณฑ์',
      'ประเภทครุภัณฑ์', 'สถานที่ตั้ง', 'รหัสที่ตั้งทรัพย์สิน', 'สถานะ', 'หมายเหตุ',
      'ปีงบประมาณ', 'วิธีได้มา', 'โครงการ', 'กิจกรรม', 'จัดซื้อจัดจ้าง',
      'เลขที่สัญญา', 'PO Code', 'วันที่ตรวจรับ', 'แหล่งเงิน', 'ประเภทรายจ่าย',
      'ผู้ขาย', 'รายการ', 'S/N', 'รุ่น', 'ยี่ห้อ',
      'มูลค่า', 'อายุการใช้งาน', 'หน่วยนับ', 'ทะเบียนรถ', 'รายละเอียดเพิ่มเติม',
      'อายุการใช้งานจริง', 'ทรัพย์สินย่อย', 'วันที่บันทึก', 'ผู้บันทึก'
    ]);
  }

  // 2. Sheet Material
  let matSheet = ss.getSheetByName('Material');
  if (!matSheet) {
    matSheet = ss.insertSheet('Material');
    matSheet.appendRow([
      'Key', 'หมายเลขวัสดุ', 'รายการ', 'หมวดหมู่วัสดุ', 'ประเภทวัสดุ',
      'หน่วยงาน', 'จำนวนคงเหลือ', 'หน่วยนับ', 'ราคาต่อหน่วย', 'จุดสั่งซื้อ',
      'สถานที่เก็บ', 'สถานะ', 'วันที่บันทึก', 'ประวัติการเบิกจ่าย'
    ]);
  }

  // 3. Sheet Borrow
  let brwSheet = ss.getSheetByName('Borrow');
  if (!brwSheet) {
    brwSheet = ss.insertSheet('Borrow');
    brwSheet.appendRow([
      'Key', 'ประเภท', 'รหัสทรัพย์สิน', 'รายการ', 'ผู้ยืม',
      'หน่วยงานผู้ยืม', 'วันที่ยืม', 'กำหนดคืน', 'วันที่คืนจริง', 'วัตถุประสงค์',
      'ผู้บันทึก', 'สถานะ', 'สภาพเมื่อคืน', 'วันที่บันทึก'
    ]);
  }

  // 4. Sheet Category
  let catSheet = ss.getSheetByName('Category');
  if (!catSheet) {
    catSheet = ss.insertSheet('Category');
    catSheet.appendRow(['Key', 'ชื่อ', 'หมวดหมู่', 'URL รูปภาพ', 'จำนวน']);
    catSheet.appendRow(['CAT-01', 'คอมพิวเตอร์และอุปกรณ์', 'ครุภัณฑ์', 'https://cdn-icons-png.flaticon.com/512/3067/3067260.png', 0]);
    catSheet.appendRow(['CAT-02', 'โต๊ะ-เก้าอี้สำนักงาน', 'ครุภัณฑ์', 'https://cdn-icons-png.flaticon.com/512/2603/2603741.png', 0]);
    catSheet.appendRow(['CAT-03', 'ยานพาหนะและขนส่ง', 'ครุภัณฑ์', 'https://cdn-icons-png.flaticon.com/512/3202/3202926.png', 0]);
    catSheet.appendRow(['CAT-04', 'กระดาษและสิ่งพิมพ์', 'วัสดุ', 'https://cdn-icons-png.flaticon.com/512/2541/2541988.png', 0]);
    catSheet.appendRow(['CAT-05', 'เครื่องเขียนและอุปกรณ์', 'วัสดุ', 'https://cdn-icons-png.flaticon.com/512/2666/2666505.png', 0]);
  }

  return true;
}

function _generateEquipKey(prefix) {
  const d = new Date();
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  const rnd = Math.floor(1000 + Math.random() * 9000);
  return `${prefix}-${y}${m}${day}-${rnd}`;
}

// Helper to normalize category names (strip prefix 'ครุภัณฑ์' or 'วัสดุ')
function _normalizeCat(cat) {
  if (!cat) return '';
  return String(cat).trim().replace(/^(ครุภัณฑ์|วัสดุ)\s*/, '').trim();
}

function seedSampleEquipmentDataIfNeeded(ss) {
  try {
    const eqSheet = ss.getSheetByName('Equipment');
    const matSheet = ss.getSheetByName('Material');
    if (!eqSheet || !matSheet) return;

    const nowStr = typeof formatToDateThaiFull !== 'undefined' ? formatToDateThaiFull(new Date()) : new Date().toLocaleString('th-TH');

    // If Equipment sheet is empty, seed realistic sample items
    if (eqSheet.getLastRow() <= 1) {
      const sampleEquips = [
        ['EQP-2026-001', 'EQP-6701-001', 'GF-67001', 'สำนักปลัด', 'คอมพิวเตอร์', 'คอมพิวเตอร์และอุปกรณ์', 'ห้องสำนักงานปลัด', 'LOC-01', 'พร้อมใช้งาน', '-', '2567', 'เฉพาะเจาะจง', 'โครงการพัฒนาระบบดิจิทัล', 'จัดซื้ออุปกรณ์', 'จัดซื้อจัดจ้าง', 'CN-67/01', 'PO-6701', '01/10/2566', 'งบประมาณประจำปี', 'งบลงทุน', 'บจก. ไอทีโซลูชั่น', 'เครื่องคอมพิวเตอร์ All-in-One Dell OptiPlex', 'SN-DELL-9981', 'OptiPlex 7400', 'Dell', 28500, '5 ปี', 'เครื่อง', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-002', 'EQP-6701-002', 'GF-67002', 'กองคลัง', 'สำนักงาน', 'เครื่องปรับอากาศ', 'ห้องกองคลัง', 'LOC-02', 'กำลังใช้งาน', '-', '2567', 'ตกลงราคา', 'โครงการปรับปรุงสำนักงาน', 'จัดซื้อเครื่องปรับอากาศ', 'จัดซื้อจัดจ้าง', 'CN-67/02', 'PO-6702', '15/10/2566', 'งบประมาณประจำปี', 'งบลงทุน', 'บจก. แอร์เซอร์วิส', 'เครื่องปรับอากาศ Daikin Inverter 24000 BTU', 'SN-DK-2401', 'FTKQ24UV2S', 'Daikin', 32000, '8 ปี', 'เครื่อง', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-003', 'EQP-6701-003', 'GF-67003', 'กองช่าง', 'ยานพาหนะและขนส่ง', 'ยานพาหนะ', 'อาคารจอดรถ', 'LOC-03', 'กำลังใช้งาน', '-', '2566', 'ประกวดราคา', 'โครงการยานพาหนะส่วนกลาง', 'จัดซื้อรถส่วนกลาง', 'จัดซื้อจัดจ้าง', 'CN-66/09', 'PO-6609', '20/08/2566', 'งบประมาณประจำปี', 'งบลงทุน', 'บจก. โตโยต้า มอเตอร์', 'รถกระบะบรรทุก Double Cab 4WD', 'SN-TY-4421', 'Hilux Revo', 'Toyota', 750000, '10 ปี', 'คัน', 'กข-1234 กทม', '', '2 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-004', 'EQP-6701-004', 'GF-67004', 'กองช่าง', 'สำรวจ', 'กล้องถ่ายภาพ/วิดีโอ', 'ห้องกองช่าง', 'LOC-04', 'พร้อมใช้งาน', '-', '2567', 'เฉพาะเจาะจง', 'โครงการสำรวจแผนที่', 'จัดซื้อโดรน', 'จัดซื้อจัดจ้าง', 'CN-67/05', 'PO-6705', '05/11/2566', 'งบประมาณประจำปี', 'งบลงทุน', 'บจก. โดรนเอเชีย', 'โดรนสำรวจภาพถ่ายทางอากาศ DJI Mavic 3 Enterprise', 'SN-DJI-7712', 'Mavic 3E', 'DJI', 125000, '5 ปี', 'ลำ', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-005', 'EQP-6701-005', 'GF-67005', 'กองสาธารณสุขและสิ่งแวดล้อม', 'การเกษตร', 'อุปกรณ์เครื่องมือการเกษตร', 'โรงเก็บพัสดุ', 'LOC-05', 'พร้อมใช้งาน', '-', '2567', 'ตกลงราคา', 'โครงการดูแลภูมิทัศน์', 'จัดซื้อเครื่องมือ', 'จัดซื้อจัดจ้าง', 'CN-67/06', 'PO-6706', '12/11/2566', 'งบประมาณประจำปี', 'งบดำเนินงาน', 'ร้านการเกษตรภัณฑ์', 'เครื่องตัดหญ้าสะพายข้อแข็ง 4 จังหวะ', 'SN-HD-5541', 'GX35', 'Honda', 9500, '5 ปี', 'เครื่อง', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-006', 'EQP-6701-006', 'GF-67006', 'กองช่าง', 'ก่อสร้าง', 'เครื่องมือก่อสร้าง', 'ห้องกองช่าง', 'LOC-06', 'พร้อมใช้งาน', '-', '2567', 'เฉพาะเจาะจง', 'โครงการซ่อมบำรุงทาง', 'จัดซื้อสว่าน', 'จัดซื้อจัดจ้าง', 'CN-67/07', 'PO-6707', '18/11/2566', 'งบประมาณประจำปี', 'งบดำเนินงาน', 'บจก. ฮาร์ดแวร์ทูลส์', 'สว่านโรตารี่เจาะกระแทกคอนกรีต Bosch GBH 2-28', 'SN-BS-1109', 'GBH 2-28', 'Bosch', 7800, '5 ปี', 'เครื่อง', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ'],
        ['EQP-2026-007', 'EQP-6701-007', 'GF-67007', 'สำนักปลัด', 'ไฟฟ้าและวิทยุ', 'อุปกรณ์เครือข่าย', 'ห้องวิทยุสื่อสาร', 'LOC-07', 'พร้อมใช้งาน', '-', '2566', 'เฉพาะเจาะจง', 'โครงการระบบสื่อสารฉุกเฉิน', 'จัดซื้อวิทยุสื่อสาร', 'จัดซื้อจัดจ้าง', 'CN-66/12', 'PO-6612', '22/12/2566', 'งบประมาณประจำปี', 'งบดำเนินงาน', 'บจก. สื่อสารโทรคมนาคม', 'วิทยุสื่อสารประจำที่ VHF ICOM IC-2300H', 'SN-ICOM-8832', 'IC-2300H', 'ICOM', 8900, '7 ปี', 'เครื่อง', '', '', '1 ปี', '[]', nowStr, 'ผู้ดูแลระบบ']
      ];
      sampleEquips.forEach(row => eqSheet.appendRow(row));
    }

    // If Material sheet is empty, seed realistic sample items
    if (matSheet.getLastRow() <= 1) {
      const sampleMaterials = [
        ['MAT-2026-001', 'MAT-6701-001', 'กระดาษถ่ายเอกสาร A4 80 แกรม Idea Work', 'สำนักงาน', 'กระดาษและสิ่งพิมพ์', 'สำนักปลัด', 85, 'รีม', 125, 20, 'ห้องพัสดุกลาง', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-002', 'MAT-6701-002', 'ตลับหมึกพิมพ์ HP LaserJet 85A', 'คอมพิวเตอร์', 'หมึกพิมพ์และโทนเนอร์', 'กองคลัง', 12, 'กล่อง', 2450, 5, 'ห้องศูนย์ข้อมูล/เซิร์ฟเวอร์', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-003', 'MAT-6701-003', 'หลอดไฟ LED Tube T8 18W Philips', 'ไฟฟ้าและวิทยุ', 'หลอดไฟและอุปกรณ์ไฟฟ้า', 'กองช่าง', 45, 'หลอด', 110, 15, 'ห้องพัสดุกลาง', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-004', 'MAT-6701-004', 'แบตเตอรี่รถยนต์ 12V 75Ah GS', 'ยานพาหนะและขนส่ง', 'อะไหล่และอุปกรณ์ซ่อมบำรุง', 'กองช่าง', 6, 'ลูก', 2800, 2, 'อาคารจอดรถ', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-005', 'MAT-6701-005', 'ชุดประแจรวม 14 ชิ้น และเครื่องมือช่าง', 'อะไหล่ทุกชนิด', 'อะไหล่และอุปกรณ์ซ่อมบำรุง', 'กองช่าง', 18, 'ชุด', 650, 5, 'ห้องกองช่าง', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-006', 'MAT-6701-006', 'น้ำยาทำความสะอาดพื้นอเนกประสงค์ 3.8L', 'งานบ้านงานครัว', 'อุปกรณ์ทำความสะอาด', 'สำนักปลัด', 24, 'แกลลอน', 180, 8, 'ห้องพัสดุกลาง', 'พร้อมใช้งาน', nowStr, '[]'],
        ['MAT-2026-007', 'MAT-6701-007', 'ปูนซีเมนต์ปอร์ตแลนด์ ผสมเสร็จ 50 กก.', 'ก่อสร้าง', 'วัสดุก่อสร้าง', 'กองช่าง', 30, 'ถุง', 145, 10, 'กองช่าง', 'พร้อมใช้งาน', nowStr, '[]']
      ];
      sampleMaterials.forEach(row => matSheet.appendRow(row));
    }
  } catch (e) {
    Logger.log('seedSampleEquipmentDataIfNeeded error: ' + e);
  }
}

// -------------------------------------------------------------
// 1. Dashboard & Statistics
// -------------------------------------------------------------
function getEquipmentSummary() {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    seedSampleEquipmentDataIfNeeded(ss);

    // Equipment stats
    const eqSheet = ss.getSheetByName('Equipment');
    const eqData = eqSheet ? eqSheet.getDataRange().getValues().slice(1) : [];
    
    let totalEquip = 0;
    let totalValue = 0;
    let totalDisposed = 0;
    const equipCatMap = {};

    eqData.forEach(row => {
      if (!row[0]) return;
      totalEquip++;
      const val = parseFloat(row[25]) || 0;
      const status = String(row[8] || '').trim();
      const rawCat = String(row[4] || 'อื่นๆ').trim();
      const normCat = _normalizeCat(rawCat) || rawCat;

      if (status === 'ตัดจำหน่าย' || status === 'จำหน่ายแล้ว') {
        totalDisposed += val;
      } else {
        totalValue += val;
      }

      equipCatMap[normCat] = (equipCatMap[normCat] || 0) + 1;
    });

    // Material stats
    const matSheet = ss.getSheetByName('Material');
    const matData = matSheet ? matSheet.getDataRange().getValues().slice(1) : [];
    
    let totalMaterial = 0;
    const matCatMap = {};

    matData.forEach(row => {
      if (!row[0]) return;
      totalMaterial++;
      const rawCat = String(row[3] || 'อื่นๆ').trim();
      const normCat = _normalizeCat(rawCat) || rawCat;
      matCatMap[normCat] = (matCatMap[normCat] || 0) + 1;
    });

    // Borrow stats
    const brwSheet = ss.getSheetByName('Borrow');
    const brwData = brwSheet ? brwSheet.getDataRange().getValues().slice(1) : [];
    let borrowingCount = 0;
    let returnedCount = 0;
    let overdueCount = 0;
    const todayStr = new Date().toISOString().split('T')[0];

    brwData.forEach(row => {
      if (!row[0]) return;
      const status = String(row[11] || '').trim();
      const returnDueDate = String(row[7] || '').trim();
      
      if (status === 'คืนแล้ว') {
        returnedCount++;
      } else {
        borrowingCount++;
        if (returnDueDate && returnDueDate < todayStr) {
          overdueCount++;
        }
      }
    });

    // Swiper Categories
    const catSheet = ss.getSheetByName('Category');
    const catData = catSheet ? catSheet.getDataRange().getValues().slice(1) : [];
    const swiperCategories = catData.map(r => {
      const code = r[0];
      const catName = String(r[1] || '').trim();
      const type = String(r[2] || '').trim();
      const icon = r[3];
      const norm = _normalizeCat(catName);

      let count = 0;
      if (type === 'ครุภัณฑ์') {
        count = equipCatMap[norm] || equipCatMap[catName] || 0;
      } else if (type === 'วัสดุ') {
        count = matCatMap[norm] || matCatMap[catName] || 0;
      } else {
        count = (equipCatMap[norm] || 0) + (matCatMap[norm] || 0);
      }

      if (count === 0 && r[4] && !isNaN(r[4])) {
        count = parseInt(r[4]) || 0;
      }

      return [code, catName, type, icon, count];
    });

    // Chart-ready series and labels
    const matLabels = Object.keys(matCatMap);
    const matSeries = matLabels.map(k => matCatMap[k]);

    const equipLabels = Object.keys(equipCatMap);
    const equipSeries = equipLabels.map(k => equipCatMap[k]);

    return {
      totalEquip: totalEquip,
      totalMaterial: totalMaterial,
      equipCount: totalEquip,
      matCount: totalMaterial,
      totalValue: totalValue,
      totalDisposed: totalDisposed,
      disposedValue: totalDisposed,
      borrowingCount: borrowingCount,
      returnedCount: returnedCount,
      overdueCount: overdueCount,
      // Both map and chart formats for compatibility
      equipCategories: {
        labels: equipLabels,
        series: equipSeries,
        total: totalEquip,
        map: equipCatMap
      },
      matCategories: {
        labels: matLabels,
        series: matSeries,
        total: totalMaterial,
        map: matCatMap
      },
      materialCategories: {
        labels: matLabels,
        series: matSeries,
        total: totalMaterial,
        map: matCatMap
      },
      swiperCategories: swiperCategories,
      categories: swiperCategories
    };
  } catch (err) {
    Logger.log('Error in getEquipmentSummary: ' + err);
    return {
      totalEquip: 0,
      totalMaterial: 0,
      equipCount: 0,
      matCount: 0,
      totalValue: 0,
      totalDisposed: 0,
      disposedValue: 0,
      borrowingCount: 0,
      returnedCount: 0,
      overdueCount: 0,
      equipCategories: { labels: [], series: [], total: 0, map: {} },
      matCategories: { labels: [], series: [], total: 0, map: {} },
      materialCategories: { labels: [], series: [], total: 0, map: {} },
      swiperCategories: [],
      categories: []
    };
  }
}

// -------------------------------------------------------------
// 2. Equipment Management (CRUD, GFMIS, Dispose)
// -------------------------------------------------------------
function getEquipmentList(filter) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues().slice(1);
    if (!filter) return data.reverse();

    const filtered = data.filter(row => {
      if (!row[0]) return false;
      if (filter.agency && row[3] && !String(row[3]).includes(filter.agency)) return false;
      if (filter.category && row[4] && !String(row[4]).includes(filter.category)) return false;
      if (filter.type && row[5] && !String(row[5]).includes(filter.type)) return false;
      if (filter.status && row[8] && !String(row[8]).includes(filter.status)) return false;
      if (filter.gfmis && row[2] && !String(row[2]).toLowerCase().includes(filter.gfmis.toLowerCase())) return false;
      if (filter.number && row[1] && !String(row[1]).toLowerCase().includes(filter.number.toLowerCase())) return false;
      if (filter.keyword) {
        const kw = filter.keyword.toLowerCase();
        const fullText = (row[1] + ' ' + row[2] + ' ' + row[21] + ' ' + row[23] + ' ' + row[24]).toLowerCase();
        if (!fullText.includes(kw)) return false;
      }
      return true;
    });

    return filtered.reverse();
  } catch (err) {
    Logger.log('Error in getEquipmentList: ' + err);
    return [];
  }
}

function saveEquipment(obj) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    const nowStr = typeof formatToDateThaiFull !== 'undefined' ? formatToDateThaiFull(new Date()) : new Date().toLocaleString('th-TH');

    const key = obj.key && obj.key.trim() !== '' ? obj.key : _generateEquipKey('EQP');
    const rowValues = [
      key,
      obj.equipAddData1 || '',  // หมายเลขครุภัณฑ์
      obj.equipAddData2 || '',  // รหัส GFMIS
      obj.equipAddData3 || '',  // หน่วยงาน
      obj.equipAddData4 || '',  // หมวดหมู่ครุภัณฑ์
      obj.equipAddData5 || '',  // ประเภทครุภัณฑ์
      obj.equipAddData6 || '',  // สถานที่ตั้ง
      obj.equipAddData7 || '',  // รหัสที่ตั้งทรัพย์สิน
      obj.equipAddData8 || 'พร้อมใช้งาน', // สถานะ
      obj.equipAddData9 || '',  // หมายเหตุ
      obj.equipAddData10 || '', // ปีงบประมาณ
      obj.equipAddData11 || '', // วิธีได้มา
      obj.equipAddData12 || '', // โครงการ
      obj.equipAddData13 || '', // กิจกรรม
      obj.equipAddData14 || '', // จัดซื้อจัดจ้าง
      obj.equipAddData15 || '', // เลขที่สัญญา
      obj.equipAddData16 || '', // PO Code
      obj.equipAddData17 || '', // วันที่ตรวจรับ
      obj.equipAddData18 || '', // แหล่งเงิน
      obj.equipAddData19 || '', // ประเภทรายจ่าย
      obj.equipAddData20 || '', // ผู้ขาย
      obj.equipAddData21 || '', // รายการ (ชื่อ)
      obj.equipAddData22 || '', // S/N
      obj.equipAddData23 || '', // รุ่น
      obj.equipAddData24 || '', // ยี่ห้อ
      parseFloat(obj.equipAddData25) || 0, // มูลค่า
      obj.equipAddData26 || '', // อายุการใช้งาน
      obj.equipAddData27 || '', // หน่วยนับ
      obj.equipAddData28 || '', // ทะเบียนรถ
      obj.equipAddData29 || '', // รายละเอียดเพิ่มเติม
      obj.equipAddData30 || '', // อายุการใช้งานจริง
      JSON.stringify(obj.subEquipments || []), // ทรัพย์สินย่อย
      nowStr,
      obj.user || 'Admin'
    ];

    if (obj.key && obj.key.trim() !== '') {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === obj.key) {
          sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
          return { success: true, action: 'update', key: key, message: 'แก้ไขข้อมูลครุภัณฑ์เรียบร้อยแล้ว' };
        }
      }
    }

    sheet.appendRow(rowValues);
    return { success: true, action: 'insert', key: key, message: 'บันทึกข้อมูลครุภัณฑ์สำเร็จ' };
  } catch (err) {
    Logger.log('Error in saveEquipment: ' + err);
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + err.toString() };
  }
}

function updateEquipmentGFMIS(key, gfmisCode) {
  try {
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 3).setValue("'" + gfmisCode);
        return { success: true, message: 'อัปเดตรหัส GFMIS เรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลครุภัณฑ์ที่ระบุ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function disposeEquipment(key, reason) {
  try {
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 9).setValue('ตัดจำหน่าย');
        const oldNote = data[i][9] || '';
        const dateNote = new Date().toLocaleDateString('th-TH');
        sheet.getRange(i + 1, 10).setValue(oldNote + ` [ตัดจำหน่ายเมื่อ ${dateNote}: ${reason || ''}]`);
        return { success: true, message: 'บันทึกการตัดจำหน่ายเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบครุภัณฑ์' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteEquipment(key) {
  try {
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข้อมูลครุภัณฑ์เรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบครุภัณฑ์' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 3. Materials / Supplies Management
// -------------------------------------------------------------
function getMaterialList(filter) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Material');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues().slice(1);
    if (!filter) return data.reverse();

    const filtered = data.filter(row => {
      if (!row[0]) return false;
      if (filter.agency && row[5] && !String(row[5]).includes(filter.agency)) return false;
      if (filter.category && row[3] && !String(row[3]).includes(filter.category)) return false;
      if (filter.type && row[4] && !String(row[4]).includes(filter.type)) return false;
      if (filter.number && row[1] && !String(row[1]).toLowerCase().includes(filter.number.toLowerCase())) return false;
      if (filter.name && row[2] && !String(row[2]).toLowerCase().includes(filter.name.toLowerCase())) return false;
      return true;
    });

    return filtered.reverse();
  } catch (err) {
    Logger.log('Error in getMaterialList: ' + err);
    return [];
  }
}

function saveMaterial(obj) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Material');
    const nowStr = typeof formatToDateThaiFull !== 'undefined' ? formatToDateThaiFull(new Date()) : new Date().toLocaleString('th-TH');

    const key = obj.key && obj.key.trim() !== '' ? obj.key : _generateEquipKey('MAT');
    const rowValues = [
      key,
      obj.number || '',         // หมายเลขวัสดุ
      obj.name || '',           // รายการ
      obj.category || '',       // หมวดหมู่วัสดุ
      obj.type || '',           // ประเภทวัสดุ
      obj.agency || '',         // หน่วยงาน
      parseInt(obj.quantity) || 0, // จำนวนคงเหลือ
      obj.unit || 'ชิ้น',       // หน่วยนับ
      parseFloat(obj.price) || 0, // ราคาต่อหน่วย
      parseInt(obj.minStock) || 5, // จุดสั่งซื้อขั้นต่ำ
      obj.location || '',       // สถานที่เก็บ
      obj.status || 'พร้อมใช้งาน', // สถานะ
      nowStr,                   // วันที่บันทึก
      JSON.stringify([])        // ประวัติการเบิกจ่าย
    ];

    if (obj.key && obj.key.trim() !== '') {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === obj.key) {
          rowValues[6] = obj.quantity !== undefined ? parseInt(obj.quantity) : data[i][6];
          rowValues[13] = data[i][13] || JSON.stringify([]);
          sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
          return { success: true, message: 'แก้ไขข้อมูลวัสดุเรียบร้อยแล้ว' };
        }
      }
    }

    sheet.appendRow(rowValues);
    return { success: true, message: 'เพิ่มรายการวัสดุเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function stockMaterial(obj) {
  try {
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Material');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === obj.key) {
        let currentStock = parseInt(data[i][6]) || 0;
        const qty = parseInt(obj.quantity) || 0;
        const nowStr = typeof formatToDateThaiFull !== 'undefined' ? formatToDateThaiFull(new Date()) : new Date().toLocaleString('th-TH');

        let history = [];
        try {
          history = JSON.parse(data[i][13] || '[]');
        } catch (e) {
          history = [];
        }

        if (obj.action === 'out') {
          if (currentStock < qty) {
            return { success: false, message: `ยอดคงเหลือไม่พอ (คงเหลือ ${currentStock} ${data[i][7]})` };
          }
          currentStock -= qty;
          history.push({
            type: 'เบิกจ่าย',
            quantity: qty,
            user: obj.user || 'ผู้เบิก',
            purpose: obj.purpose || '',
            date: nowStr
          });
        } else {
          currentStock += qty;
          history.push({
            type: 'รับเข้า',
            quantity: qty,
            user: obj.user || 'ผู้รับ',
            purpose: obj.note || '',
            date: nowStr
          });
        }

        sheet.getRange(i + 1, 7).setValue(currentStock);
        sheet.getRange(i + 1, 14).setValue(JSON.stringify(history));

        return {
          success: true,
          newStock: currentStock,
          message: obj.action === 'out' ? `เบิกจ่ายสำเร็จ คงเหลือ ${currentStock} ${data[i][7]}` : `รับเข้าสำเร็จ ยอดรวม ${currentStock} ${data[i][7]}`
        };
      }
    }
    return { success: false, message: 'ไม่พบรายการวัสดุ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteMaterial(key) {
  try {
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Material');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบรายการวัสดุเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบรายการวัสดุ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 4. Borrow & Return Workflow
// -------------------------------------------------------------
function getBorrowList(filter) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Borrow');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues().slice(1);
    const todayStr = new Date().toISOString().split('T')[0];

    const mapped = data.map(row => {
      if (row[11] === 'กำลังยืม' && row[7] && String(row[7]) < todayStr) {
        row[11] = 'เกินกำหนด';
      }
      return row;
    });

    if (!filter) return mapped.reverse();

    const filtered = mapped.filter(row => {
      if (!row[0]) return false;
      if (filter.status && row[11] !== filter.status) return false;
      if (filter.borrower && !String(row[4]).toLowerCase().includes(filter.borrower.toLowerCase())) return false;
      if (filter.keyword) {
        const kw = filter.keyword.toLowerCase();
        const str = (row[0] + ' ' + row[2] + ' ' + row[3] + ' ' + row[4] + ' ' + row[5]).toLowerCase();
        if (!str.includes(kw)) return false;
      }
      return true;
    });

    return filtered.reverse();
  } catch (err) {
    Logger.log('Error in getBorrowList: ' + err);
    return [];
  }
}

function saveBorrow(obj) {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const brwSheet = ss.getSheetByName('Borrow');
    const nowStr = typeof formatToDateThaiFull !== 'undefined' ? formatToDateThaiFull(new Date()) : new Date().toLocaleString('th-TH');
    const key = _generateEquipKey('BRW');

    const row = [
      key,
      obj.itemType || 'ครุภัณฑ์',
      obj.itemCode || '',
      obj.itemName || '',
      obj.borrowerName || '',
      obj.borrowerAgency || '',
      obj.borrowDate || '',
      obj.returnDueDate || '',
      '',
      obj.purpose || '',
      obj.recordedBy || 'Admin',
      'กำลังยืม',
      '',
      nowStr
    ];

    brwSheet.appendRow(row);

    if (obj.itemType === 'ครุภัณฑ์' && obj.itemCode) {
      const eqSheet = ss.getSheetByName('Equipment');
      if (eqSheet) {
        const eqData = eqSheet.getDataRange().getValues();
        for (let i = 1; i < eqData.length; i++) {
          if (eqData[i][0] === obj.itemCode || eqData[i][1] === obj.itemCode) {
            eqSheet.getRange(i + 1, 9).setValue('กำลังใช้งาน');
            break;
          }
        }
      }
    }

    return { success: true, key: key, message: 'บันทึกการยืมเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function returnBorrow(obj) {
  try {
    const ss = _getEquipSS();
    const brwSheet = ss.getSheetByName('Borrow');
    const data = brwSheet.getDataRange().getValues();
    const returnDate = obj.returnDate || (new Date().toISOString().split('T')[0]);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === obj.key) {
        brwSheet.getRange(i + 1, 9).setValue(returnDate);
        brwSheet.getRange(i + 1, 12).setValue('คืนแล้ว');
        brwSheet.getRange(i + 1, 13).setValue(obj.condition || 'สภาพปกติพร้อมใช้งาน');

        const itemType = data[i][1];
        const itemCode = data[i][2];
        if (itemType === 'ครุภัณฑ์' && itemCode) {
          const eqSheet = ss.getSheetByName('Equipment');
          if (eqSheet) {
            const eqData = eqSheet.getDataRange().getValues();
            for (let j = 1; j < eqData.length; j++) {
              if (eqData[j][0] === itemCode || eqData[j][1] === itemCode) {
                eqSheet.getRange(j + 1, 9).setValue('พร้อมใช้งาน');
                break;
              }
            }
          }
        }

        return { success: true, message: 'บันทึกรับคืนเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบรายการยืม' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getEquipmentAvailableForBorrow() {
  try {
    initEquipmentDatabase();
    const ss = _getEquipSS();
    const sheet = ss.getSheetByName('Equipment');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues().slice(1);
    return data
      .filter(r => r[0] && (r[8] === 'พร้อมใช้งาน' || !r[8]))
      .map(r => ({
        key: r[0],
        number: r[1],
        name: r[21] || r[1],
        category: r[4],
        brand: r[24]
      }));
  } catch (e) {
    return [];
  }
}
