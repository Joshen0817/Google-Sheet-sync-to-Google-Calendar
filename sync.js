/***** 基本設定 *****/
const SHEET_NAME = '9月總表';            // 你的工作表名稱
const CALENDAR_ID = 'c_4195f6751d6b20bd95f32d8b52c4657f5a2cf8771ea2790fefbae8f4c637fff5@group.calendar.google.com'; // ← 改成你的團隊行事曆 ID

// 支援多種表頭別名（你可保持目前中文欄名）
const HEADER_NAMES = {
  project: ['專案'],
  participants: ['參與人'],
  task: ['負責項目'],
  start: ['開始日期','開始時間','開始'],
  end: ['結束日期','結束時間','結束'],
  noteToCal: ['備註（Calendar）','Calendar備註','行事曆備註','附註','備註'],
  eventId: ['Event ID','事件ID'],
  status: ['狀態','Status']
};

/***** 實用工具 *****/
function getSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`找不到工作表「${SHEET_NAME}」`);
  return sh;
}
function getCalendar_() {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error(`找不到行事曆：${CALENDAR_ID}`);
  return cal;
}
function findHeaderMap_(sh) {
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v => String(v||'').trim());
  const idx = {};
  Object.entries(HEADER_NAMES).forEach(([key, aliases]) => {
    const pos = header.findIndex(h => aliases.map(a => a.toLowerCase()).includes(h.toLowerCase()));
    if (pos === -1) throw new Error(`找不到表頭：${aliases.join(' / ')}`);
    idx[key] = pos + 1; // 1-based
  });
  return idx;
}
function normalizeId_(id){ return id ? String(id).split('@')[0] : ''; }
function val_(sh,r,c){ return sh.getRange(r,c).getValue(); }
function set_(sh,r,c,v){ sh.getRange(r,c).setValue(v); }
function toDateOrNull_(v){
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (v === '' || v == null) return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
// 標題格式：A欄(專案) + "-" + F欄(負責項目) + "：" + E欄(參與人)
function buildTitle_(project, task, participants){
  const a = String(project||'').trim();
  const f = String(task||'').trim();
  const e = String(participants||'').trim();
  let t = '';
  if (a) t += a;
  if (f) t += (t ? '-' : '') + f;
  if (e) t += (t ? '：' : '') + e;
  return t || '(未命名事件)';
}

function isMidnight_(d){
  if (!(d instanceof Date)) return false;
  return d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0;
}
function dateOnly_(d){
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}



/***** Sheet → Calendar：單列同步（建立/更新/刪除）*****/
function syncRowToCalendar_(rowIndex){
  const sh = getSheet_();
  const cols = findHeaderMap_(sh);
  const cal = getCalendar_();

  const project = val_(sh,rowIndex,cols.project);        // A
  const participants = val_(sh,rowIndex,cols.participants); // E
  const task = val_(sh,rowIndex,cols.task);                 // F
  const start = toDateOrNull_(val_(sh,rowIndex,cols.start));// H
  const end   = toDateOrNull_(val_(sh,rowIndex,cols.end));  // I
  const note  = val_(sh,rowIndex,cols.noteToCal);           // M（放入 Calendar 備註）
  const status= String(val_(sh,rowIndex,cols.status)||'').trim().toUpperCase();
  const existingId = normalizeId_(val_(sh,rowIndex,cols.eventId));

  // 刪除
  if (status === 'DELETE' && existingId){
    const ev = CalendarApp.getEventById(existingId);
    if (ev) ev.deleteEvent();
    set_(sh,rowIndex,cols.eventId,'');
    set_(sh,rowIndex,cols.status,''); // 清掉 DELETE
    return;
  }

  // 必要欄位檢查
  if (!start){
    // 沒開始時間就不建/不更
    return;
  }

  const title = buildTitle_(project, task, participants);
  const desc = String(note || '');

  // 判斷是否用「整天事件」：當 start、end 都是 00:00（或只有 start 有值、end 空白）→ 視為整天
  const useAllDay = isMidnight_(start) && (!end || isMidnight_(end));

  // 決定時間邏輯
  let ev = existingId ? CalendarApp.getEventById(existingId) : null;

  if (useAllDay) {
    // 整天事件：Sheet 的「結束日期」視為最後一天（人類直覺）
    const startDay = dateOnly_(start);
    // 若沒填 end，就當作單日；若有填 end，就顯示到該日，所以要 +1 天給 Calendar
    const endDayPlus1 = end ? new Date(end.getFullYear(), end.getMonth(), end.getDate() + 1)
                            : new Date(startDay.getFullYear(), startDay.getMonth(), startDay.getDate() + 1);

    if (!ev){
      ev = getCalendar_().createAllDayEvent(title, startDay, endDayPlus1, {
        description: desc
      });
      set_(sh,rowIndex,cols.eventId, normalizeId_(ev.getId()));
    }else{
      // 若既有事件是「非整天」，先改成整天時間段
      ev.setTitle(title);
      ev.setDescription(desc);
      ev.setAllDayDates(startDay, endDayPlus1);
    }

  } else {
    // 一般「有時間」的事件：照你填的時間，不做 +1 天
    const endTime = (end && end > start) ? end : new Date(start.getTime() + 60*60*1000);

    if (!ev){
      ev = getCalendar_().createEvent(title, start, endTime, { description: desc });
      set_(sh,rowIndex,cols.eventId, normalizeId_(ev.getId()));
    }else{
      ev.setTitle(title);
      ev.setTime(start, endTime);
      ev.setDescription(desc);
    }
  }

}

/***** onEdit：只同步被改動的那一列 *****/
function onEdit(e){
  try{
    // 若是手動執行，e 會是 undefined；直接結束避免報錯
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    const row = e.range.getRow();
    if (row === 1) return; // 跳過表頭

    const cols = findHeaderMap_(sh);
    const changedCol = e.range.getColumn();
    const watched = [cols.project, cols.participants, cols.task, cols.start, cols.end, cols.noteToCal, cols.status, cols.eventId];
    if (!watched.includes(changedCol)) return;

    syncRowToCalendar_(row);
    emphasizeDeleteCell_();
  }catch(err){
    console.error(err);
  }
}


/***** 手動重推全表（可在需要時執行一次）*****/
function syncAllRows(){
  const sh = getSheet_();
  const last = sh.getLastRow();
  for (let r=2; r<=last; r++){
    syncRowToCalendar_(r);
  }
  emphasizeDeleteCell_();
}

/***** Event ID 欄：隱藏＋保護；顯示＋解除保護 *****/
function hideAndProtectEventId_() {
  const sh = getSheet_();
  const cols = findHeaderMap_(sh);
  const col = cols.eventId;

  // 隱藏
  sh.hideColumns(col);

  // 保護（自第2列起）
  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription && p.getDescription() === 'System-managed Event ID');
  protections.forEach(p => p.remove());

  const rng = sh.getRange(2, col, sh.getMaxRows() - 1);
  const p = rng.protect();
  p.setDescription('System-managed Event ID');
  try {
    p.removeEditors(p.getEditors());
  } catch (err) {
    if (p.canDomainEdit()) p.setDomainEdit(false);
  }
}
function showAndUnprotectEventId_() {
  const sh = getSheet_();
  const cols = findHeaderMap_(sh);
  const col = cols.eventId;

  sh.showColumns(col);
  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription && p.getDescription() === 'System-managed Event ID');
  protections.forEach(p => p.remove());
}

/***** 輔助：欄號轉 A1 欄字母 *****/
function colToA1_(n){
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/***** 套用格式與驗證（DELETE 標紅、日期驗證、狀態下拉、表頭保護）*****/
function applyFormattingAndValidation_() {
  const sh = getSheet_();
  const cols = findHeaderMap_(sh);

  const lastRow = Math.max(2, sh.getMaxRows());
  const lastCol = sh.getLastColumn();

  // 日期/時間欄：驗證 + 顯示格式
  const dateRngs = [
    sh.getRange(2, cols.start, lastRow - 1, 1),
    sh.getRange(2, cols.end,   lastRow - 1, 1)
  ];
  const dvDate = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('請輸入有效的日期或日期時間（例：2025/08/11 14:30）')
    .build();
  dateRngs.forEach(rng => {
    rng.setDataValidation(dvDate);
    rng.setNumberFormat('yyyy/mm/dd hh:mm');
  });

  // 狀態欄：下拉（允許空白、建議用 DELETE）
  const statusRng = sh.getRange(2, cols.status, lastRow - 1, 1);
  const dvStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(['DELETE'], true)
    .setAllowInvalid(true)
    .setHelpText('刪除事件請選擇或輸入：DELETE（大小寫需一致）')
    .build();
  statusRng.setDataValidation(dvStatus);

  // 條件式格式：狀態=DELETE → 整列標紅
  const statusColLetter = colToA1_(cols.status);
  const dataRange = sh.getRange(2, 1, sh.getMaxRows() - 1, lastCol);
  const rules = sh.getConditionalFormatRules() || [];
  const deleteFormula = `=$${statusColLetter}2="DELETE"`;
  const cleanedRules = rules.filter(rule => {
    const condition = rule.getBooleanCondition();
    if (!condition) return true;
    if (condition.getType() !== SpreadsheetApp.BooleanConditionType.CUSTOM_FORMULA) return true;
    const values = condition.getValues() || [];
    return values[0] !== deleteFormula;
  });

  const delRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(deleteFormula)
    .setBackground('#FDECEC')
    .setFontColor('#B00020')
    .setRanges([dataRange])
    .build();
  cleanedRules.push(delRule);
  sh.setConditionalFormatRules(cleanedRules);

  // 表頭列保護
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription && p.getDescription() === 'Header protected')
    .forEach(p => p.remove());
  const headProt = sh.getRange(1, 1, 1, lastCol).protect();
  headProt.setDescription('Header protected');
  try {
    headProt.removeEditors(headProt.getEditors());
  } catch (e) {
    if (headProt.canDomainEdit()) headProt.setDomainEdit(false);
  }

  // Event ID 欄維持隱藏＋保護
  hideAndProtectEventId_();
}

/***** 視覺提醒：把 DELETE 的儲存格加粗（可選）*****/
function emphasizeDeleteCell_(){
  const sh = getSheet_();
  const cols = findHeaderMap_(sh);
  const rng = sh.getRange(2, cols.status, sh.getMaxRows() - 1, 1);
  const values = rng.getValues();
  rng.setFontWeight('normal');
  for (let i=0; i<values.length; i++){
    if (String(values[i][0]).trim().toUpperCase() === 'DELETE'){
      sh.getRange(2 + i, cols.status).setFontWeight('bold');
    }
  }
}

/***** 自訂選單（整合所有功能）*****/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Calendar 同步')
    .addItem('重推全表（Sheet → Calendar）','syncAllRows')
    .addSeparator()
    .addItem('初始化（建立系統欄位）','initCalendarSync_')
    .addSeparator()
    .addItem('隱藏並保護 Event ID 欄','hideAndProtectEventId_')
    .addItem('顯示並解除保護 Event ID 欄','showAndUnprotectEventId_')
    .addSeparator()
    .addItem('套用格式與驗證（DELETE 標紅、日期驗證、表頭保護）','applyFormattingAndValidation_')
    .addToUi();

  // 開啟時自動套用一次（確保新列也有規則）
  applyFormattingAndValidation_();
  emphasizeDeleteCell_();
}

/***** 若缺欄就自動建立：Event ID、狀態 *****/
function ensureSystemColumns_() {
  const sh = getSheet_();
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v => String(v||'').trim());
  const needEventId = header.findIndex(h => ['Event ID','事件ID'].map(x=>x.toLowerCase()).includes(h.toLowerCase())) === -1;
  const needStatus  = header.findIndex(h => ['狀態','Status'].map(x=>x.toLowerCase()).includes(h.toLowerCase())) === -1;

  let added = false;

  if (needEventId) {
    sh.insertColumnAfter(sh.getLastColumn());
    const col = sh.getLastColumn();
    sh.getRange(1, col).setValue('Event ID');
    added = true;
  }
  if (needStatus) {
    sh.insertColumnAfter(sh.getLastColumn());
    const col = sh.getLastColumn();
    sh.getRange(1, col).setValue('狀態');
    added = true;
  }

  if (added) {
    // 重新套格式與保護（包含隱藏 Event ID 欄）
    applyFormattingAndValidation_();
  }
}

/***** 一鍵初始化：建立系統欄位＋套格式＋隱藏保護 *****/
function initCalendarSync_(){
  ensureSystemColumns_();
  applyFormattingAndValidation_();  // 日期驗證/DELETE 標紅/表頭保護/Event ID 保護
  emphasizeDeleteCell_();
  SpreadsheetApp.getUi().alert('初始化完成：已建立 Event ID／狀態 欄並套用格式與保護。');
}

// 早期版本使用立即執行函式注入選單，會在觸發器環境中造成錯誤；已改為 onOpen 中建構選單。
