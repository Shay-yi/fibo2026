/************ 配置：销售人员默认密码 ************/
const DEFAULT_SALES_PASSWORDS = {
  'Jessica': 'jessica2026',
  'Gaby': 'gaby2026',
  'Peter': 'peter2026',
  'Donna': 'donna2026',
  'Henry': 'henry2026',
  'Richard': 'richard2026',
  'Cecile': 'cecile2026',
  'Alex': 'alex2026',
  'Shay': 'shay2026',
  'Victor': 'victor2026',
  'Kelly': 'kelly2026',
  'Josie': 'josie2026',
  'Vera': 'vera2026',
  'Patrick': 'patrick2026',
  'Zack': 'zack2026',
  'Owen': 'owen2026',
  'Kat': 'kat2026'
};

const SALES_NAMES = Object.keys(DEFAULT_SALES_PASSWORDS);

/************ 目标文件夹：FIBO client collection ************/
const SALES_FOLDER_ID = '1MknVDt8-dxWIgU3Iymw_qnh3UzH5CkUa';

/************ 每个销售文件里的三个 sheet 名称 ************/
const FIBO_SHEET = 'FIBO';
const ASH_SHEET = 'ASH';
const BOOTH_SHEET = 'BOOTH';

/************ 表头：view/index.html 依赖列顺序 ************/
const FIBO_HEADER = [
  'Timestamp',
  'Event Type',
  'Full Name',
  'Company',
  'Product Category',
  'Country',
  'Meeting Date',
  'Time Slot',
  'Notes'
];

const ASH_BOOTH_HEADER = [
  'Timestamp',
  'Event Type',
  'Full Name',
  'Company',
  'Country',
  'Notes'
];

/************ 可选：全量汇总表 ************/
const MASTER_ENABLED = false;
const MASTER_SHEET_NAME = 'All Submissions';

/************ ScriptProperties 前缀 ************/
const PROP_PREFIX = 'SALES_SPREADSHEET_ID_';
const PASSWORD_PROP_PREFIX = 'SALES_PASSWORD_';

/************ SpreadsheetId 读写 ************/
function getSalesSpreadsheetId_(salesName) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(PROP_PREFIX + salesName);
}

function setSalesSpreadsheetId_(salesName, spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_PREFIX + salesName, spreadsheetId);
}

/************ 密码读写（简单版：明文存储） ************/
function getStoredPassword_(salesName) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(PASSWORD_PROP_PREFIX + salesName);
}

function setStoredPassword_(salesName, password) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PASSWORD_PROP_PREFIX + salesName, password);
}

function getEffectivePassword_(salesName) {
  return getStoredPassword_(salesName) || DEFAULT_SALES_PASSWORDS[salesName] || '';
}

function isUsingDefaultPassword_(salesName, effectivePassword) {
  const def = DEFAULT_SALES_PASSWORDS[salesName];
  if (!def) return false;
  return effectivePassword === def;
}

/************ 基础工具 ************/
function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSalesFolder_() {
  return DriveApp.getFolderById(SALES_FOLDER_ID);
}

/************ Sheet 结构辅助 ************/
function ensureHeader_(sheet, header) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.appendRow(header);
    return;
  }

  const firstRow = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const looksEmpty = !firstRow || firstRow.every(v => v === '' || v === null);
  if (looksEmpty) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function ensureSheet_(ss, sheetName, header) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(header);
    return sh;
  }
  ensureHeader_(sh, header);
  return sh;
}

function ensureSalesSpreadsheetStructure_(ss) {
  const allowed = {};
  allowed[FIBO_SHEET] = true;
  allowed[ASH_SHEET] = true;
  allowed[BOOTH_SHEET] = true;

  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    const nm = sh.getName();
    if (!allowed[nm]) {
      try { ss.deleteSheet(sh); } catch (err) {}
    }
  });

  ensureSheet_(ss, FIBO_SHEET, FIBO_HEADER);
  ensureSheet_(ss, ASH_SHEET, ASH_BOOTH_HEADER);
  ensureSheet_(ss, BOOTH_SHEET, ASH_BOOTH_HEADER);
}

function moveSpreadsheetToFolder_(spreadsheetId) {
  const folder = getSalesFolder_();
  const file = DriveApp.getFileById(spreadsheetId);
  file.moveTo(folder);
}

/************ 前端数据解析 ************/
function parseIncomingData_(e) {
  const paramData = e && e.parameter && e.parameter.data ? e.parameter.data : null;

  if (paramData) {
    const parsed = JSON.parse(paramData);
    if (parsed && parsed.data) {
      return (typeof parsed.data === 'string') ? JSON.parse(parsed.data) : parsed.data;
    }
    return parsed;
  }

  const contents = e && e.postData && e.postData.contents ? e.postData.contents : null;
  if (!contents) throw new Error('Missing request body');

  if (contents.startsWith('data=')) {
    const raw = decodeURIComponent(contents.substring(5));
    const parsed = JSON.parse(raw);
    if (parsed && parsed.data) {
      return (typeof parsed.data === 'string') ? JSON.parse(parsed.data) : parsed.data;
    }
    return parsed;
  }

  const parsed = JSON.parse(contents);
  if (parsed && parsed.data) {
    return (typeof parsed.data === 'string') ? JSON.parse(parsed.data) : parsed.data;
  }
  return parsed;
}

function toNotesWithAttendees_(notes, attendees) {
  const base = notes ? String(notes) : '';
  const att = (attendees === undefined || attendees === null) ? '' : String(attendees);

  if (!att) return base;
  if (!base) return 'Attendees: ' + att;
  return base + ' | Attendees: ' + att;
}

/************ 初始化/重建（已有表格时不要运行 rebuild） ************/
function initSalesSpreadsheets() {
  const created = [];

  SALES_NAMES.forEach(name => {
    const existingId = getSalesSpreadsheetId_(name);
    if (existingId) return;

    const ss = SpreadsheetApp.create('MBH ' + name + ' Client Data');
    ensureSalesSpreadsheetStructure_(ss);
    moveSpreadsheetToFolder_(ss.getId());

    setSalesSpreadsheetId_(name, ss.getId());
    created.push(name);
  });

  return 'Initialized: ' + created.join(', ');
}

function resetSalesSpreadsheetIds() {
  const props = PropertiesService.getScriptProperties();
  SALES_NAMES.forEach(name => props.deleteProperty(PROP_PREFIX + name));
  return 'reset spreadsheet ids done';
}

function rebuildSalesSheets() {
  resetSalesSpreadsheetIds();
  return initSalesSpreadsheets();
}

function dumpSalesSpreadsheetIds() {
  const props = PropertiesService.getScriptProperties();
  const out = {};
  SALES_NAMES.forEach(name => {
    out[name] = props.getProperty(PROP_PREFIX + name) || '';
  });
  return out;
}

/************ 写入（表单提交 / 修改密码） ************/
function doPost(e) {
  try {
    const data = parseIncomingData_(e);
    const action = String(data.action || '').trim();

    if (action === 'changePassword') {
      const salesName = String(data.salesName || '').trim();
      const oldPassword = String(data.oldPassword || '');
      const newPassword = String(data.newPassword || '');

      if (!salesName || !oldPassword || !newPassword) {
        throw new Error('Missing password change fields');
      }
      if (!DEFAULT_SALES_PASSWORDS[salesName]) {
        throw new Error('Unknown sales name');
      }
      if (newPassword.length < 6) {
        throw new Error('New password must be at least 6 characters');
      }

      const currentPassword = getEffectivePassword_(salesName);
      if (!currentPassword || currentPassword !== oldPassword) {
        throw new Error('Current password is incorrect');
      }

      setStoredPassword_(salesName, newPassword);

      return jsonOutput_({
        success: true,
        message: 'Password updated successfully'
      });
    }

    const eventType = String(data.eventType || '').toUpperCase().trim();
    const salesContact = String(data.salesContact || '').trim();

    if (!eventType) throw new Error('Missing eventType');
    if (!salesContact) throw new Error('Missing salesContact');

    const spreadsheetId = getSalesSpreadsheetId_(salesContact);
    if (!spreadsheetId) {
      throw new Error('SpreadsheetId not found for ' + salesContact + '. Run rebuildSalesSheets().');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    ensureSalesSpreadsheetStructure_(ss);

    if (MASTER_ENABLED) {
      let master = ss.getSheetByName(MASTER_SHEET_NAME);
      if (!master) {
        master = ss.insertSheet(MASTER_SHEET_NAME);
        master.appendRow([
          'Timestamp', 'Event Type', 'Sales Contact', 'Full Name', 'Company',
          'Product Category', 'Country', 'Meeting Date', 'Time Slot', 'Notes'
        ]);
      }
      master.appendRow([
        new Date().toLocaleString(),
        eventType,
        salesContact,
        data.fullName || '',
        data.company || '',
        data.productCategory || '',
        data.country || '',
        data.meetingDate || '',
        data.timeSlot || '',
        data.notes || ''
      ]);
    }

    const ts = new Date().toLocaleString();

    if (eventType === 'FIBO') {
      const sh = ensureSheet_(ss, FIBO_SHEET, FIBO_HEADER);
      sh.appendRow([
        ts,
        eventType,
        data.fullName || '',
        data.company || '',
        data.productCategory || '',
        data.country || '',
        data.meetingDate || '',
        data.timeSlot || '',
        data.notes || ''
      ]);
      return jsonOutput_({ success: true });
    }

    if (eventType === 'ASH') {
      const sh = ensureSheet_(ss, ASH_SHEET, ASH_BOOTH_HEADER);
      sh.appendRow([
        ts,
        eventType,
        data.fullName || '',
        data.company || '',
        data.country || '',
        toNotesWithAttendees_(data.notes, data.attendees)
      ]);
      return jsonOutput_({ success: true });
    }

    if (eventType === 'BOOTH') {
      const sh = ensureSheet_(ss, BOOTH_SHEET, ASH_BOOTH_HEADER);
      sh.appendRow([
        ts,
        eventType,
        data.fullName || '',
        data.company || '',
        data.country || '',
        toNotesWithAttendees_(data.notes, data.attendees)
      ]);
      return jsonOutput_({ success: true });
    }

    throw new Error('Unknown eventType: ' + eventType);

  } catch (err) {
    return jsonOutput_({
      success: false,
      message: err.toString()
    });
  }
}

/************ 读取（登录 / 加载数据） + GET 表单（?data= 与前端 iframe 一致） ************/
function doGet(e) {
  try {
    const action = (e && e.parameter) ? String(e.parameter.action || '').trim() : '';
    const rawData = (e && e.parameter && e.parameter.data) ? e.parameter.data : '';

    // 公开表单：前端用隐藏 iframe 加载 GET ?data=<json>，避免 POST/302 丢 body
    if (rawData && action !== 'login' && action !== 'loadData') {
      try {
        const parsed = JSON.parse(rawData);
        const ev = String(parsed.eventType || '').trim();
        const sc = String(parsed.salesContact || '').trim();
        if (ev && sc) {
          return doPost({ parameter: { data: rawData } });
        }
      } catch (err) {
        return jsonOutput_({ success: false, message: 'Invalid form data: ' + err.toString() });
      }
    }

    if (action !== 'login' && action !== 'loadData') {
      return jsonOutput_({ status: 'ok' });
    }

    const salesName = String((e && e.parameter && e.parameter.sales) ? e.parameter.sales : '').trim();
    const password = String((e && e.parameter && e.parameter.password) ? e.parameter.password : '');

    if (!salesName || !password) {
      return jsonOutput_({ success: false, message: 'Missing credentials' });
    }

    const effectivePassword = getEffectivePassword_(salesName);
    if (!effectivePassword || effectivePassword !== password) {
      return jsonOutput_({ success: false, message: 'Invalid credentials' });
    }

    const spreadsheetId = getSalesSpreadsheetId_(salesName);
    if (!spreadsheetId) {
      throw new Error('SpreadsheetId not found for ' + salesName + '. Run initSalesSpreadsheets() or rebuildSalesSheets().');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    ensureSalesSpreadsheetStructure_(ss);

    const fiboSh = ss.getSheetByName(FIBO_SHEET);
    const ashSh = ss.getSheetByName(ASH_SHEET);
    const boothSh = ss.getSheetByName(BOOTH_SHEET);

    const fibo = fiboSh ? fiboSh.getDataRange().getValues().slice(1) : [];
    const ash = ashSh ? ashSh.getDataRange().getValues().slice(1) : [];
    const booth = boothSh ? boothSh.getDataRange().getValues().slice(1) : [];

    const targetSheet = fiboSh || ss.getSheets()[0];
    const sheetUrl = targetSheet ? (ss.getUrl() + '#gid=' + targetSheet.getSheetId()) : ss.getUrl();

    if (action === 'login') {
      return jsonOutput_({
        success: true,
        mustChangePassword: (
          !!DEFAULT_SALES_PASSWORDS[salesName] &&
          effectivePassword === DEFAULT_SALES_PASSWORDS[salesName]
        ),
        sheetUrl,
        spreadsheetId
      });
    }

    return jsonOutput_({
      success: true,
      sheetUrl,
      spreadsheetId,
      data: { fibo, ash, booth }
    });

  } catch (err) {
    return jsonOutput_({
      success: false,
      message: err.toString()
    });
  }
}

function resetOnePassword(salesName) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PASSWORD_PROP_PREFIX + salesName);
  return 'reset done: ' + salesName;
}
