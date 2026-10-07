/**
 * Hospital Directory — read-only Google Sheets adapter.
 * Open the EXISTING directory sheet -> Extensions -> Apps Script.
 * Review this file against your existing script before replacing it.
 * This script NEVER edits spreadsheet cells or their sharing settings.
 */
const DIRECTORY_CONFIG = {
  spreadsheetId: '1yTJtGeebZ_Btzmhy9i3pqbjhrUiJQxW8Ffh2WFOETO8',
  sheetId: 0 // Verified directory tab; stable even if its title is renamed.
};

function doGet(e) {
  const requested = e && e.parameter && e.parameter.callback;
  const callback = requested || 'callback';
  if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,100}$/.test(callback)) {
    return ContentService.createTextOutput(JSON.stringify({error:'Invalid callback'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  let payload;
  try { payload = readDirectory_(); }
  catch (err) {
    console.error(err);
    payload = {error:'Directory unavailable. Check the spreadsheet, tab ID, required headers and deployment permissions.'};
  }
  const json = JSON.stringify(payload).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function readDirectory_() {
  const book = DIRECTORY_CONFIG.spreadsheetId
    ? SpreadsheetApp.openById(DIRECTORY_CONFIG.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('No spreadsheet configured');
  const sheet = book.getSheets().find(tab => tab.getSheetId() === DIRECTORY_CONFIG.sheetId);
  if (!sheet) throw new Error('Configured directory tab not found');
  // Display values retain leading zeros and multiple extensions in one cell.
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(v=>String(v).trim().toLowerCase().replace(/[\s_-]+/g,' '));
  const aliases = {
    id:['id','record id','المعرف'],
    category:['category','الفئة','التصنيف'],
    service:['service','service name','department','name','الخدمة','اسم الخدمة'],
    extension:['extension','ext','ext.','phone','التحويلة','الرقم'],
    building:['building','location','المبنى'],
    aliases:['aliases','keywords','search terms','كلمات البحث']
  };
  const indexes = {};
  Object.keys(aliases).forEach(key=>{
    const matches=headers.map((h,i)=>aliases[key].includes(h)?i:-1).filter(i=>i>=0);
    if(matches.length>1)throw new Error('Ambiguous header: '+key);
    indexes[key]=matches.length?matches[0]:-1;
  });
  ['category','service','extension'].forEach(key=>{if(indexes[key]<0)throw new Error('Missing header: '+key);});
  const rows=[],ids=new Set();
  for(let i=1;i<values.length;i++){
    const row=values[i];
    const get=key=>indexes[key]<0?'':String(row[indexes[key]]||'').trim();
    const record={id:get('id'),category:get('category'),service:get('service'),extension:get('extension'),building:get('building'),aliases:get('aliases')};
    if(!record.service&&!record.extension)continue;
    if(!record.service||!record.extension)throw new Error('Missing service or extension at row '+(i+1));
    if(record.id){if(ids.has(record.id))throw new Error('Duplicate record id at row '+(i+1));ids.add(record.id);}
    record.category=record.category||'General';
    rows.push(record);
  }
  return rows;
}
