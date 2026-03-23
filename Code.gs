const SHEET_GROUPS = 'Groups';
const SHEET_ACCOUNTS = 'Accounts';
const SHEET_EXTRA = 'Extra_Sessions';
const SHEET_USERS = 'Users';

const ACCOUNTS_ORDER = [
  'FMR65','FMR111','FMR312','FMR610','FMR620','FMR789','FMR790',
  'FMR187','FMR759','FMR536','FMR107','FMR484','FMR357','FMR184',
  'FMR440','FMR586','FMR1544','FMR1545','FMR1554','FMR1555',
  'FMR1559','FMR1560','FMR651','FMR421'
];

function sendJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]).trim() : ''; });
      return obj;
    });
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr) return null;
  timeStr = String(timeStr).trim().toUpperCase();
  const match12 = timeStr.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)/);
  if (match12) {
    let h = parseInt(match12[1]);
    const m = parseInt(match12[2] || '0');
    const ampm = match12[3];
    if (ampm === 'PM' && h !== 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return h * 60 + m;
  }
  const match24 = timeStr.match(/(\d{1,2}):(\d{2})/);
  if (match24) return parseInt(match24[1]) * 60 + parseInt(match24[2]);
  return null;
}

function parseTimeRange(timeStr) {
  if (!timeStr) return null;
  timeStr = String(timeStr).replace('\u2013','-').replace('\u2014','-');
  const parts = timeStr.split('-');
  if (parts.length < 2) return null;
  const start = parseTimeToMinutes(parts[0].trim());
  const end = parseTimeToMinutes(parts.slice(1).join('-').trim());
  if (start === null || end === null) return null;
  return { start, end };
}

function timesOverlap(r1, r2) {
  if (!r1 || !r2) return false;
  return r1.start < r2.end && r2.start < r1.end;
}

function detectConflicts(groups, extraSessions) {
  const conflicts = [];
  const today = new Date();
  const activeSessions = (extraSessions || []).filter(s => {
    try { return today >= new Date(s['Start Date']) && today <= new Date(s['End Date']); }
    catch(e) { return false; }
  });
  for (let i = 0; i < groups.length; i++) {
    for (let j = i + 1; j < groups.length; j++) {
      const g1 = groups[i], g2 = groups[j];
      if (!g1['Zoom Account'] || !g2['Zoom Account']) continue;
      if (g1['Zoom Account'] !== g2['Zoom Account']) continue;
      if (g1['Type'] !== 'Online' || g2['Type'] !== 'Online') continue;
      const days1 = parseDays(g1['Days']);
      const days2 = parseDays(g2['Days']);
      const sharedDays = days1.filter(d => days2.includes(d));
      if (!sharedDays.length) continue;
      const t1 = parseTimeRange(g1['Time']);
      const t2 = parseTimeRange(g2['Time']);
      if (!timesOverlap(t1, t2)) continue;
      const sameSlot = groups.filter(g =>
        g['Zoom Account'] === g1['Zoom Account'] &&
        g['Type'] === 'Online' &&
        parseDays(g['Days']).some(d => sharedDays.includes(d)) &&
        timesOverlap(parseTimeRange(g['Time']), t1)
      );
      if (sameSlot.length > 2) {
        conflicts.push({
          type: 'Group-Group', severity: 'HIGH',
          account: g1['Zoom Account'],
          item1: { name: g1['Group ID'], days: g1['Days'], time: g1['Time'] },
          item2: { name: g2['Group ID'], days: g2['Days'], time: g2['Time'] },
          overlappingDays: sharedDays, overlappingTime: g1['Time']
        });
      }
    }
  }
  groups.forEach(g => {
    if (!g['Zoom Account'] || g['Type'] !== 'Online') return;
    const gDays = parseDays(g['Days']);
    const gTime = parseTimeRange(g['Time']);
    activeSessions.forEach(s => {
      if (s['Zoom Account'] !== g['Zoom Account']) return;
      const sDays = parseDays(s['Days']);
      const sharedDays = gDays.filter(d => sDays.includes(d));
      if (!sharedDays.length) return;
      const sTime = { start: parseTimeToMinutes(s['Start Time']), end: parseTimeToMinutes(s['End Time']) };
      if (!timesOverlap(gTime, sTime)) return;
      const count = groups.filter(g2 =>
        g2['Zoom Account'] === g['Zoom Account'] && g2['Type'] === 'Online' &&
        parseDays(g2['Days']).some(d => sharedDays.includes(d)) &&
        timesOverlap(parseTimeRange(g2['Time']), gTime)
      ).length;
      if (count + 1 > 2) {
        conflicts.push({
          type: 'Group-ExtraSession', severity: 'CRITICAL',
          account: g['Zoom Account'],
          item1: { name: g['Group ID'], days: g['Days'], time: g['Time'] },
          item2: { name: s['Session Name'], days: s['Days'], time: s['Start Time']+' - '+s['End Time'] },
          overlappingDays: sharedDays, overlappingTime: g['Time']
        });
      }
    });
  });
  return conflicts;
}

function canAssign(accountName, newGroup, currentGroups, activeSessions) {
  const newDays = parseDays(newGroup['Days']);
  const newTime = parseTimeRange(newGroup['Time']);
  if (!newTime) return true;
  const acctGroups = currentGroups.filter(g => g['Zoom Account'] === accountName && g['Type'] === 'Online');
  const acctSessions = activeSessions.filter(s => s['Zoom Account'] === accountName);
  for (const day of newDays) {
    const intervals = [];
    acctGroups.forEach(g => {
      if (parseDays(g['Days']).includes(day)) { const t = parseTimeRange(g['Time']); if (t) intervals.push(t); }
    });
    acctSessions.forEach(s => {
      if (parseDays(s['Days']).includes(day)) {
        const t = { start: parseTimeToMinutes(s['Start Time']), end: parseTimeToMinutes(s['End Time']) };
        if (t.start !== null && t.end !== null) intervals.push(t);
      }
    });
    intervals.push(newTime);
    const events = [];
    intervals.forEach(t => { events.push({ time: t.start, type: 1 }); events.push({ time: t.end, type: -1 }); });
    events.sort((a,b) => a.time - b.time || a.type - b.type);
    let cur = 0, max = 0;
    events.forEach(e => { cur += e.type; if (cur > max) max = cur; });
    if (max > 2) return false;
  }
  return true;
}

function runAutoAssign() {
  const groups = sheetToObjects(SHEET_GROUPS);
  const extra = sheetToObjects(SHEET_EXTRA);
  const today = new Date();
  const activeSessions = extra.filter(s => {
    try { return today >= new Date(s['Start Date']) && today <= new Date(s['End Date']); }
    catch(e) { return false; }
  });
  const sheet = getSheet(SHEET_GROUPS);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const accountCol = headers.indexOf('Zoom Account') + 1;
  let assigned = 0, unassigned = 0;
  const details = [];
  const working = groups.map(g => ({...g}));
  working.forEach((group, idx) => {
    if (group['Zoom Account'] && group['Zoom Account'].trim() !== '') return;
    if (group['Type'] === 'Offline') {
      const acct = ACCOUNTS_ORDER[idx % ACCOUNTS_ORDER.length];
      working[idx]['Zoom Account'] = acct;
      sheet.getRange(idx+2, accountCol).setValue(acct);
      assigned++;
      details.push({ group: group['Group ID'], account: acct, type: 'Offline' });
      return;
    }
    let found = false;
    for (const acct of ACCOUNTS_ORDER) {
      if (canAssign(acct, group, working, activeSessions)) {
        working[idx]['Zoom Account'] = acct;
        sheet.getRange(idx+2, accountCol).setValue(acct);
        assigned++;
        details.push({ group: group['Group ID'], account: acct, type: 'Online' });
        found = true;
        break;
      }
    }
    if (!found) { unassigned++; details.push({ group: group['Group ID'], account: 'UNASSIGNED', type: 'Online' }); }
  });
  return { success: true, assigned, unassigned, details };
}

function runAutoFix() {
  const sheet = getSheet(SHEET_GROUPS);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const accountCol = headers.indexOf('Zoom Account') + 1;
  const groups = sheetToObjects(SHEET_GROUPS);
  const extra = sheetToObjects(SHEET_EXTRA);
  const today = new Date();
  const activeSessions = extra.filter(s => {
    try { return today >= new Date(s['Start Date']) && today <= new Date(s['End Date']); }
    catch(e) { return false; }
  });
  const conflicts = detectConflicts(groups, extra);
  if (!conflicts.length) return { success: true, fixed: 0, remaining: 0 };
  const conflictIds = new Set(conflicts.map(c => c.item1.name));
  let fixed = 0;
  const working = groups.map(g => ({...g}));
  working.forEach((group, idx) => {
    if (!conflictIds.has(group['Group ID'])) return;
    working[idx]['Zoom Account'] = '';
    for (const acct of ACCOUNTS_ORDER) {
      if (canAssign(acct, group, working, activeSessions)) {
        working[idx]['Zoom Account'] = acct;
        sheet.getRange(idx+2, accountCol).setValue(acct);
        fixed++;
        break;
      }
    }
  });
  return { success: true, fixed, remaining: detectConflicts(working, extra).length };
}

function computeStats() {
  const groups = sheetToObjects(SHEET_GROUPS);
  const extra = sheetToObjects(SHEET_EXTRA);
  const conflicts = detectConflicts(groups, extra);
  const groupsByDay = {};
  DAY_NAMES.forEach(d => groupsByDay[d] = 0);
  const groupsByDiploma = {}, groupsByAccount = {};
  const usedAccounts = new Set();
  groups.forEach(g => {
    parseDays(g['Days']).forEach(d => { groupsByDay[d] = (groupsByDay[d]||0) + 1; });
    const dip = g['Diploma'] || 'Unknown';
    groupsByDiploma[dip] = (groupsByDiploma[dip]||0) + 1;
    if (g['Zoom Account']) { usedAccounts.add(g['Zoom Account']); groupsByAccount[g['Zoom Account']] = (groupsByAccount[g['Zoom Account']]||0) + 1; }
  });
  return {
    totalGroups: groups.length,
    onlineGroups: groups.filter(g => g['Type']==='Online').length,
    offlineGroups: groups.filter(g => g['Type']==='Offline').length,
    activeGroups: groups.filter(g => g['Status']==='Active').length,
    wipGroups: groups.filter(g => g['Status']==='WIP').length,
    completeGroups: groups.filter(g => g['Status']==='Complete').length,
    totalAccounts: 24, usedAccounts: usedAccounts.size,
    availableAccounts: 24 - usedAccounts.size,
    totalConflicts: conflicts.length,
    groupsByDay, groupsByDiploma, groupsByAccount
  };
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getGroups') return sendJSON(sheetToObjects(SHEET_GROUPS));
    if (action === 'getAccounts') {
      const accounts = sheetToObjects(SHEET_ACCOUNTS);
      const groups = sheetToObjects(SHEET_GROUPS);
      return sendJSON(accounts.map(a => ({
        ...a,
        groupCount: groups.filter(g => g['Zoom Account'] === a['Account Name']).length
      })));
    }
    if (action === 'getExtraSessions') {
      const sessions = sheetToObjects(SHEET_EXTRA);
      const today = new Date();
      return sendJSON(sessions.map(s => ({
        ...s,
        isExpired: s['End Date'] ? today > new Date(s['End Date']) : false,
        isActive: s['Start Date'] && s['End Date'] ? today >= new Date(s['Start Date']) && today <= new Date(s['End Date']) : false
      })));
    }
    if (action === 'login') {
      const users = sheetToObjects(SHEET_USERS);
      const user = users.find(u => u['Username'] === e.parameter.username && u['Password'] === e.parameter.password);
      if (user) return sendJSON({ success: true, role: user['Role'], name: user['Name'], email: user['Email'] });
      return sendJSON({ success: false, message: 'Invalid credentials' });
    }
    if (action === 'getConflicts') {
      return sendJSON(detectConflicts(sheetToObjects(SHEET_GROUPS), sheetToObjects(SHEET_EXTRA)));
    }
    if (action === 'getStats') return sendJSON(computeStats());
    return sendJSON({ error: 'Unknown action' });
  } catch(err) { return sendJSON({ error: err.message }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'addGroup') {
      const sheet = getSheet(SHEET_GROUPS);
      const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      sheet.appendRow(headers.map(h => body.rowData[h] || ''));
      return sendJSON({ success: true, message: 'Group added' });
    }
    if (action === 'updateGroup') {
      const sheet = getSheet(SHEET_GROUPS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rowIdx = data.findIndex((row,i) => i>0 && row[headers.indexOf('Group ID')] === body.groupId);
      if (rowIdx === -1) return sendJSON({ success: false, message: 'Not found' });
      Object.entries(body.updates).forEach(([k,v]) => { const col = headers.indexOf(k); if(col!==-1) sheet.getRange(rowIdx+1,col+1).setValue(v); });
      return sendJSON({ success: true });
    }
    if (action === 'deleteGroup') {
      const sheet = getSheet(SHEET_GROUPS);
      const data = sheet.getDataRange().getValues();
      const rowIdx = data.findIndex((row,i) => i>0 && row[data[0].indexOf('Group ID')] === body.groupId);
      if (rowIdx === -1) return sendJSON({ success: false });
      sheet.deleteRow(rowIdx+1);
      return sendJSON({ success: true });
    }
    if (action === 'updateAccount') {
      const sheet = getSheet(SHEET_ACCOUNTS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rowIdx = data.findIndex((row,i) => i>0 && row[headers.indexOf('Account Name')] === body.accountName);
      if (rowIdx === -1) return sendJSON({ success: false });
      Object.entries(body.updates).forEach(([k,v]) => { const col = headers.indexOf(k); if(col!==-1) sheet.getRange(rowIdx+1,col+1).setValue(v); });
      return sendJSON({ success: true });
    }
    if (action === 'addExtraSession') {
      const sheet = getSheet(SHEET_EXTRA);
      const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      sheet.appendRow(headers.map(h => body.sessionData[h] || ''));
      return sendJSON({ success: true });
    }
    if (action === 'deleteExtraSession') {
      const sheet = getSheet(SHEET_EXTRA);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rowIdx = data.findIndex((row,i) => i>0 && row[headers.indexOf('Session Name')] === body.sessionName && row[headers.indexOf('Zoom Account')] === body.zoomAccount);
      if (rowIdx === -1) return sendJSON({ success: false });
      sheet.deleteRow(rowIdx+1);
      return sendJSON({ success: true });
    }
    if (action === 'autoAssign') return sendJSON(runAutoAssign());
    if (action === 'autoFix') return sendJSON(runAutoFix());
    return sendJSON({ error: 'Unknown action' });
  } catch(err) { return sendJSON({ error: err.message }); }
}

const DAY_NAMES = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];

function parseDays(daysStr) {
  if (!daysStr) return [];
  return String(daysStr).split(',')
    .map(d => d.trim().toLowerCase())
    .filter(d => DAY_NAMES.includes(d));
}
