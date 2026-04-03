/**
 * [근무표 자동화 시스템 세팅 가이드]
 * 1. 환경설정 시트:
 * - '시작일', '종료일' 항목에 기간을 입력하세요 (예: 2026-02-01).
 * - E열(5열) '공휴일 목록' 아래에 공휴일 날짜들을 입력하세요.
 * 2. 근무표 시트 구성 (필수): 
 * - A열(1열): 직원 성함 / B열(2열): 그룹명
 * - C열(3열): +1 (여성휴가 등 추가휴무) / D열(4열): H풀 / E열(5열): D풀 / F열(6열): P풀
 * - G열(7열): 휴무 합계 (자동 계산) 
 * - H~K열(8~11열): 이전 달 마지막 4일치 데이터 (N간격 및 6일룰 계산용)
 * - L열(12열) ~ : 당월 날짜별 근무 시작
 */

const CRITICAL_ROLES = ['N', 'D', 'E', 'H', 'M1', 'M2', 'M3', 'P'];

function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getConfigValue(configData, label) {
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0].toString().trim() === label) {
      return configData[i][1];
    }
  }
  return null;
}

function getAppConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('환경설정');
  
  const defaultItems = [
    ['항목', '값', '설명', '', '공휴일 목록'],
    ['시작일', '2026-02-01', 'YYYY-MM-DD 형식', '', '2026-03-01'],
    ['종료일', '2026-02-28', 'YYYY-MM-DD 형식', '', '2026-05-05'],
    ['업무배제시작일', 15, '재조정 기준일', '', ''],
    ['이름시작행', 3, '이름 시작 행', '', ''],
    ['그룹열번호', 2, '그룹명 열(B=2)', '', '']
  ];

  if (!configSheet) {
    configSheet = ss.insertSheet('환경설정');
    applyConfigTemplate(configSheet, defaultItems);
  } else {
    const currentData = configSheet.getDataRange().getValues();
    const hasStartDate = currentData.some(row => row[0].toString().trim() === '시작일');
    if (!hasStartDate) {
      configSheet.clear(); 
      applyConfigTemplate(configSheet, defaultItems);
    }
  }

  const configData = configSheet.getDataRange().getValues();
  const holidayValues = configSheet.getRange(2, 5, Math.max(configSheet.getLastRow(), 2), 1).getValues();
  const holidayList = holidayValues.flat().filter(String).map(d => {
    try { 
      const dateObj = new Date(d);
      if (isNaN(dateObj.getTime())) return null;
      return Utilities.formatDate(dateObj, "GMT+9", "yyyy-MM-dd"); 
    } catch(e) { return null; }
  }).filter(Boolean);

  return {
    START_DATE: new Date(getConfigValue(configData, '시작일') || new Date()),
    END_DATE: new Date(getConfigValue(configData, '종료일') || new Date()),
    EXCLUSION_START_DAY: parseInt(getConfigValue(configData, '업무배제시작일')) || 15,
    START_ROW: parseInt(getConfigValue(configData, '이름시작행')) || 3,
    GROUP_COL: parseInt(getConfigValue(configData, '그룹열번호')) || 2, 
    NAME_COL: 1,
    PLUS_ONE_COL: 3,   // C열: +1 (추가 휴무 부여)
    H_POOL_COL: 4,     // D열: H(조근) 인력풀
    D_POOL_COL: 5,     // E열: D(데스크) 인력풀
    P_POOL_COL: 6,     // F열: P 인력풀
    SUMMARY_COL: 7,    // G열: 휴무 합계
    PAST_DAYS_COL: 8,  // H열: 과거 4일 데이터 시작 (H, I, J, K)
    DATE_START_COL: 12,// L열: 당월 근무 시작
    HOLIDAYS: holidayList,
    WEEKEND_EXCLUDE_ROLES: ['M2', 'E']
  };
}

function applyConfigTemplate(sheet, items) {
  sheet.getRange(1, 1, items.length, 5).setValues(items);
  sheet.getRange(1, 1, 1, 5).setBackground('#444444').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 120); sheet.setColumnWidth(3, 200); sheet.setColumnWidth(4, 30); sheet.setColumnWidth(5, 120);
  sheet.getRange(2, 2, items.length - 1, 1).setBorder(true, true, true, true, null, null, "#0000ff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function manualResetConfig() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('환경설정 시트 초기화', '현재 설정된 값이 모두 사라집니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = ss.getSheetByName('환경설정');
    if (!configSheet) configSheet = ss.insertSheet('환경설정');
    configSheet.clear();
    const defaultItems = [['항목', '값', '설명', '', '공휴일 목록'], ['시작일', '2026-02-01', '', '', ''], ['종료일', '2026-02-28', '', '', ''], ['업무배제시작일', 15, '', '', ''], ['이름시작행', 3, '', '', ''], ['그룹열번호', 2, '', '', '']];
    applyConfigTemplate(configSheet, defaultItems);
    ui.alert('환경설정 시트가 초기화되었습니다.');
  }
}

function getDateIndex(inputStr, startDate, totalDays) {
  const parts = inputStr.split('/');
  if (parts.length !== 2) return -1;
  const month = parseInt(parts[0]);
  const day = parseInt(parts[1]);
  const targetDate = new Date(startDate.getFullYear(), month - 1, day);
  const timeDiff = targetDate.getTime() - startDate.getTime();
  const diffDays = Math.round(timeDiff / (1000 * 3600 * 24));
  return (diffDays < 0 || diffDays >= totalDays) ? -1 : diffDays;
}

function isRedDay(date, holidays) {
  const dayOfWeek = date.getDay(); 
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return dayOfWeek === 0 || dayOfWeek === 6 || holidays.includes(dateStr);
}

function isSundayOrHoliday(date, holidays) {
  const dayOfWeek = date.getDay();
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return dayOfWeek === 0 || holidays.includes(dateStr);
}

function getActualEmployeeCount(sheet, startRow, nameCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 0;
  const names = sheet.getRange(startRow, nameCol, lastRow - startRow + 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < names.length; i++) {
    if (names[i][0].toString().trim() !== "") count++;
    else break; 
  }
  return count;
}

function getPastData(sheet, config, employeeCount) {
  if (employeeCount <= 0) return [];
  const pastRange = sheet.getRange(config.START_ROW, config.PAST_DAYS_COL, employeeCount, 4);
  return pastRange.getValues().map(row => row.map(v => v.toString().trim().toUpperCase()));
}

function updateHolidaySummary(sheet, config, employeeCount, diffDays) {
  if (employeeCount <= 0) return;
  const summaryFormulas = [];
  const startColLetter = getColumnLetter(config.DATE_START_COL);
  const endColLetter = getColumnLetter(config.DATE_START_COL + diffDays - 1);
  const nameColLetter = getColumnLetter(config.NAME_COL);

  for (let i = 0; i < employeeCount; i++) {
    const row = config.START_ROW + i;
    const formula = `=IF(${nameColLetter}${row}="", "", COUNTIF(${startColLetter}${row}:${endColLetter}${row}, "X") + COUNTIF(${startColLetter}${row}:${endColLetter}${row}, "R"))`;
    summaryFormulas.push([formula]);
  }

  const summaryRange = sheet.getRange(config.START_ROW, config.SUMMARY_COL, employeeCount, 1);
  summaryRange.setFormulas(summaryFormulas);
  summaryRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
}

function getRestCountsAndAverage(data, employeeCount, diffDays) {
  let restCounts = new Array(employeeCount).fill(0);
  for(let i=0; i<employeeCount; i++) {
    for(let d=0; d<diffDays; d++) {
      if(data[i][d] === 'X' || data[i][d] === 'R') restCounts[i]++;
    }
  }
  let avgRestDays = Math.round(restCounts.reduce((a,b) => a + b, 0) / employeeCount);
  return { restCounts, avgRestDays };
}

function enforce6DayRulePostAssignment(data, pastData, employeeCount, diffDays) {
  for (let i = 0; i < employeeCount; i++) {
    let streak = 0;
    let currentStreakDays = [];
    
    for (let p = 0; p < 4; p++) {
      const role = pastData[i][p];
      // A (Absence) 추가
      if (['X', 'R', 'A', ''].includes(role)) {
        streak = 0;
        currentStreakDays = [];
      } else {
        streak++;
        currentStreakDays.push(p - 4); 
      }
    }

    for (let d = 0; d < diffDays; d++) {
      const role = data[i][d];
      if (['X', 'R', 'A'].includes(role)) {
        streak = 0;
        currentStreakDays = [];
      } else {
        streak++;
        currentStreakDays.push(d);
        if (streak > 6) {
          let sIndex = -1;
          for (let k = 0; k < currentStreakDays.length; k++) {
            let dayIdx = currentStreakDays[k];
            if (dayIdx >= 0 && data[i][dayIdx] === 'S') {
              sIndex = dayIdx;
              break;
            }
          }
          if (sIndex !== -1) {
            data[i][sIndex] = 'R';
            d = -1; 
            streak = 0;
            currentStreakDays = [];
            for (let p = 0; p < 4; p++) {
              const rRole = pastData[i][p];
              if (['X', 'R', 'A', ''].includes(rRole)) { streak = 0; currentStreakDays = []; }
              else { streak++; currentStreakDays.push(p - 4); }
            }
          }
        }
      }
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 근무표 관리')
    .addItem('0. 환경설정 시트 초기화/정리', 'manualResetConfig')
    .addSeparator()
    .addItem('1. 날짜 및 서식 설정 (전체 초기화)', 'setupDateHeaders')
    .addItem('1-1. 수기입력 유지 및 로테이션 초기화', 'clearScheduleKeepManual')
    .addItem('1-2. 특정 인원 데스크(D) 사전 지정', 'showDeskAssignmentDialog')
    .addItem('1-3. 특정 인원 H(조근) 사전 지정', 'showHAssignmentDialog')
    .addSeparator()
    .addItem('2. 로테이션 원클릭 자동완성 (생성+리밸런싱)', 'generateBaseSchedule')
    .addSeparator()
    .addItem('3. 특정 인원 배제 및 보직 사수', 'showExclusionDialog')
    .addItem('4. 중복 및 인원 검증', 'validateSchedule')
    .addToUi();
}

function setupDateHeaders() {
  const config = getAppConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === '환경설정') { SpreadsheetApp.getUi().alert('근무표 시트를 선택하고 실행해주세요.'); return; }

  const employeeCount = getActualEmployeeCount(sheet, config.START_ROW, config.NAME_COL);
  const timeDiff = config.END_DATE.getTime() - config.START_DATE.getTime();
  const diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)) + 1;
  
  if (diffDays <= 0) { SpreadsheetApp.getUi().alert('날짜 설정을 확인해주세요.'); return; }

  const headerRow = config.START_ROW - 1;
  const currentMaxRows = sheet.getMaxRows();
  const currentMaxCols = sheet.getMaxColumns();
  
  const clearColsHeader = Math.max(35, currentMaxCols - config.PLUS_ONE_COL + 1);
  sheet.getRange(headerRow, config.PLUS_ONE_COL, 1, clearColsHeader).clearContent().setBackground(null);
  
  const maxDataRows = Math.max(employeeCount, currentMaxRows - config.START_ROW + 1);
  const clearColsData = Math.max(35, currentMaxCols - config.SUMMARY_COL + 1);
  sheet.getRange(config.START_ROW, config.SUMMARY_COL, maxDataRows, clearColsData).clearContent().setBackground(null).setFontWeight('normal');
  
  sheet.getRange(headerRow, config.PLUS_ONE_COL).setValue("+1").setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f4cccc');
  sheet.getRange(headerRow, config.H_POOL_COL).setValue("H풀").setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d9ead3');
  sheet.getRange(headerRow, config.D_POOL_COL).setValue("D풀").setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c9daf8');
  sheet.getRange(headerRow, config.P_POOL_COL).setValue("P풀").setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fff2cc');
  sheet.getRange(headerRow, config.SUMMARY_COL).setValue("휴무계").setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d9d9d9');

  const pastDaysArr = ['D-4', 'D-3', 'D-2', 'D-1'];
  for(let p = 0; p < 4; p++) {
    const cell = sheet.getRange(headerRow, config.PAST_DAYS_COL + p);
    cell.setValue(pastDaysArr[p]).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#eeeeee').setFontColor('#888888');
    sheet.setColumnWidth(config.PAST_DAYS_COL + p, 40);
  }

  const headerRowData = [];
  const daysArr = ['일', '월', '화', '수', '목', '금', '토'];
  for (let d = 0; d < diffDays; d++) {
    let curr = new Date(config.START_DATE);
    curr.setDate(curr.getDate() + d);
    headerRowData.push(`${curr.getMonth()+1}/${curr.getDate()}(${daysArr[curr.getDay()]})`);
  }
  sheet.getRange(headerRow, config.DATE_START_COL, 1, diffDays).setValues([headerRowData]);
  
  for (let d = 0; d < diffDays; d++) {
    let curr = new Date(config.START_DATE);
    curr.setDate(curr.getDate() + d);
    const cell = sheet.getRange(headerRow, config.DATE_START_COL + d);
    cell.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (isRedDay(curr, config.HOLIDAYS)) cell.setBackground('#e6b8af');
    else cell.setBackground('#f3f3f3');
  }

  sheet.setColumnWidth(config.PLUS_ONE_COL, 40); 
  sheet.setColumnWidth(config.H_POOL_COL, 50);
  sheet.setColumnWidth(config.D_POOL_COL, 50);
  sheet.setColumnWidth(config.P_POOL_COL, 50);
  sheet.setColumnWidth(config.SUMMARY_COL, 60);
  
  updateHolidaySummary(sheet, config, maxDataRows, diffDays);
  applyFormatting(sheet, diffDays, config, maxDataRows);
  SpreadsheetApp.getUi().alert('날짜 헤더가 생성되었습니다.\n\n⚠️ H~K열(D-4 ~ D-1)에 이전 달 마지막 4일 치 근무 데이터를 복사/붙여넣기 하신 후 로테이션을 생성해 주세요.');
}

function clearScheduleKeepManual() {
  const config = getAppConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getName() === '환경설정') {
    SpreadsheetApp.getUi().alert('근무표 시트를 선택하고 실행해주세요.');
    return;
  }

  const employeeCount = getActualEmployeeCount(sheet, config.START_ROW, config.NAME_COL);
  const timeDiff = config.END_DATE.getTime() - config.START_DATE.getTime();
  const diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)) + 1;
  
  if (diffDays <= 0 || employeeCount <= 0) return;

  const range = sheet.getRange(config.START_ROW, config.DATE_START_COL, employeeCount, diffDays);
  const data = range.getValues();

  for (let i = 0; i < employeeCount; i++) {
    for (let d = 0; d < diffDays; d++) {
      const val = data[i][d].toString().toUpperCase();
      // 'A' 코드 추가 보호
      if (['X', 'P', 'D', 'H', 'A'].includes(val) || (d === 0 && val === 'W')) { 
        data[i][d] = val; 
      } else {
        data[i][d] = ''; 
      }
    }
  }

  range.setValues(data);
  updateHolidaySummary(sheet, config, employeeCount, diffDays);
  applyFormatting(sheet, diffDays, config, employeeCount);
  SpreadsheetApp.getUi().alert('수기 입력을 제외한 자동 배정 근무가 모두 초기화되었습니다.\n(과거 4일 치 데이터는 안전하게 유지됩니다.)');
}

function generateBaseSchedule() {
  const config = getAppConfig();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  const promptResult = ui.prompt('목표 휴무일수', '인당 기본 목표 휴무일수(X+R)를 입력하세요.\n(+1 열에 O가 있는 인원은 자동으로 1일이 더해집니다.)', ui.ButtonSet.OK_CANCEL);
  if (promptResult.getSelectedButton() !== ui.Button.OK) return; 
  let targetOffDays = parseInt(promptResult.getResponseText().trim()) || null;

  const employeeCount = getActualEmployeeCount(activeSheet, config.START_ROW, config.NAME_COL);
  const timeDiff = config.END_DATE.getTime() - config.START_DATE.getTime();
  const diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)) + 1;

  const range = activeSheet.getRange(config.START_ROW, config.DATE_START_COL, employeeCount, diffDays);
  const existingValues = range.getValues();
  const pastData = getPastData(activeSheet, config, employeeCount);

  const groupData = activeSheet.getRange(config.START_ROW, config.GROUP_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim());
  const plusOneData = activeSheet.getRange(config.START_ROW, config.PLUS_ONE_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());
  const hPoolData = activeSheet.getRange(config.START_ROW, config.H_POOL_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());
  const dPoolData = activeSheet.getRange(config.START_ROW, config.D_POOL_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());
  const pPoolData = activeSheet.getRange(config.START_ROW, config.P_POOL_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());

  const poolStandard = []; const poolE_Disp = []; 
  const poolH = []; const poolD = []; const poolP = [];
  for (let i = 0; i < employeeCount; i++) {
    if (groupData[i] === '일반직' || groupData[i] === '전문직') { poolStandard.push(i); }
    if (groupData[i] === '파견직') { poolE_Disp.push(i); }
    if (hPoolData[i] === 'O' || hPoolData[i] === 'ㅇ') poolH.push(i);
    if (dPoolData[i] === 'O' || dPoolData[i] === 'ㅇ') poolD.push(i);
    if (pPoolData[i] === 'O' || pPoolData[i] === 'ㅇ') poolP.push(i);
  }

  let ptrNW = 0, ptrD = 0, ptrE1 = 0, ptrE2 = 0, ptrH = 0, ptrM1 = 0, ptrM2 = 0, ptrM3 = 0, ptrP = 0;
  const accumulatedRestCounts = new Array(employeeCount).fill(0);
  const totalWorkCounts = new Array(employeeCount).fill(0);
  const consecutiveWorkDays = new Array(employeeCount).fill(0);

  for (let i = 0; i < employeeCount; i++) {
    let streak = 0;
    for (let p = 0; p < 4; p++) {
       let r = pastData[i][p];
       if (['X', 'R', 'A', ''].includes(r)) streak = 0;
       else streak++;
    }
    consecutiveWorkDays[i] = streak;
  }

  for (let i = 0; i < employeeCount; i++) {
    for (let d = 0; d < diffDays; d++) {
      if (existingValues[i][d] === 'X' || existingValues[i][d] === 'x') {
        accumulatedRestCounts[i]++;
      }
    }
  }

  const fullData = Array.from({ length: employeeCount }, () => new Array(diffDays).fill(''));

  for (let d = 0; d < diffDays; d++) {
    let currDate = new Date(config.START_DATE.getTime() + d * 86400000);
    const dayOfWeek = currDate.getDay();
    const isWk = isRedDay(currDate, config.HOLIDAYS);
    const isSunday = (dayOfWeek === 0);
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6); 
    const isMonToThu = (dayOfWeek >= 1 && dayOfWeek <= 4);  
    
    const assignedToday = new Set();
    let hasManualP = false;

    if (d === 0) {
      for (let i = 0; i < employeeCount; i++) {
        if (!assignedToday.has(i) && pastData[i][3] === 'N') {
          fullData[i][d] = 'W';
          assignedToday.add(i);
        }
      }
    }

    for (let i = 0; i < employeeCount; i++) {
      const val = existingValues[i][d].toString().toUpperCase();
      // 'A' 추가 보호
      if (['X', 'P', 'D', 'H', 'A'].includes(val) || (d === 0 && val === 'W')) { 
        fullData[i][d] = val; 
        assignedToday.add(i); 
        if (val !== 'X' && val !== 'W' && val !== 'A') totalWorkCounts[i]++; 
        if (val === 'P') hasManualP = true;
      }
    }

    if (d > 0) {
      for (let i = 0; i < employeeCount; i++) {
        if (!assignedToday.has(i) && fullData[i][d-1] === 'N') { 
          fullData[i][d] = 'W'; assignedToday.add(i); 
        }
      }
    }

    const assignRole = (pool, role, skipCondition = false) => {
      if (skipCondition || !pool || pool.length === 0) return;
      
      let roleStr = (role === 'E1' || role === 'E2') ? 'E' : role;
      
      if (['N', 'D', 'H', 'M1', 'M2', 'M3'].includes(roleStr)) {
        let alreadyAssigned = false;
        for (let i = 0; i < employeeCount; i++) {
          if (fullData[i][d] === roleStr) { alreadyAssigned = true; break; }
        }
        if (alreadyAssigned) return;
      }

      let startPtr = (role==='N') ? ptrNW : (role==='D') ? ptrD : (role==='E1') ? ptrE1 : (role==='E2') ? ptrE2 : (role==='H') ? ptrH : (role==='M1') ? ptrM1 : (role==='M2') ? ptrM2 : (role==='M3') ? ptrM3 : (role==='P') ? ptrP : 0;
      let selectedIdx = -1;

      const checkRecentN = (idx, currentD, gap) => {
        for (let k = Math.max(-4, currentD - gap); k < currentD; k++) {
          if (k < 0) {
             if (pastData[idx][4 + k] === 'N') return true;
          } else {
             if (fullData[idx][k] === 'N') return true;
          }
        }
        return false;
      };

      const checkTomorrowX = (idx, currentD) => {
        if (currentD + 1 >= diffDays) return false;
        let tom = existingValues[idx][currentD+1].toString().toUpperCase();
        // A 코드도 N-W 배정 불가
        return tom === 'X' || tom === 'A';
      };

      const getDaysSinceLastN = (idx, currentD) => {
        for (let k = currentD - 1; k >= -4; k--) {
          if (k < 0) {
              if (pastData[idx][4 + k] === 'N') return currentD - k;
          } else {
              if (fullData[idx][k] === 'N') return currentD - k;
          }
        }
        return 999; 
      };

      const checkSameDayOfWeekN = (idx, currentD) => {
        for (let k = currentD - 7; k >= -4; k -= 7) {
          if (k < 0) {
              if (pastData[idx][4 + k] === 'N') return true;
          } else {
              if (fullData[idx][k] === 'N') return true;
          }
        }
        return false;
      };

      const levels = [
        (idx) => { 
            if (role === 'N' && consecutiveWorkDays[idx] >= 5) return false;
            if (role !== 'N' && consecutiveWorkDays[idx] >= 6) return false;
            if (role === 'N') {
                if (d > 0 && fullData[idx][d-1] === 'N') return false;
                if (checkRecentN(idx, d, 5)) return false;
                if (checkTomorrowX(idx, d)) return false;
                if (checkSameDayOfWeekN(idx, d)) return false; 
            }
            return true;
        },
        (idx) => { 
            if (role !== 'N') return false;
            if (d > 0 && fullData[idx][d-1] === 'N') return false;
            if (checkRecentN(idx, d, 5)) return false;
            if (checkTomorrowX(idx, d)) return false;
            if (checkSameDayOfWeekN(idx, d)) return false;
            return true;
        },
        (idx) => { 
            if (role !== 'N') return false;
            if (d > 0 && fullData[idx][d-1] === 'N') return false;
            if (checkRecentN(idx, d, 5)) return false;
            if (checkSameDayOfWeekN(idx, d)) return false;
            return true;
        },
        (idx) => { 
            if (role !== 'N') return false;
            if (d > 0 && fullData[idx][d-1] === 'N') return false;
            if (checkRecentN(idx, d, 3)) return false;
            if (checkSameDayOfWeekN(idx, d)) return false;
            return true;
        },
        (idx) => { 
            if (role === 'N' && d > 0 && fullData[idx][d-1] === 'N') return false;
            if (role === 'N' && checkRecentN(idx, d, 3)) return false;
            return true; 
        }
      ];

      let validCandidates = [];
      for (let lvl = 0; lvl < levels.length; lvl++) {
        for (let idx of pool) {
            if (assignedToday.has(idx)) continue;
            if (levels[lvl](idx)) validCandidates.push(idx);
        }
        if (validCandidates.length > 0) break; 
      }

      if (validCandidates.length === 0) {
         for (let idx of poolStandard) {
            if (assignedToday.has(idx)) continue;
            if (role === 'N' && d > 0 && fullData[idx][d-1] === 'N') continue;
            validCandidates.push(idx);
         }
      }

      if (validCandidates.length > 0) {
        validCandidates.sort((a, b) => {
            let cA = 0; let cB = 0;
            for(let k=0; k<=d; k++) {
              let rA = fullData[a][k] === 'E1' || fullData[a][k] === 'E2' ? 'E' : fullData[a][k];
              if (rA === (role === 'E1' || role === 'E2' ? 'E' : role)) cA++;
            }
            for(let k=0; k<=d; k++) {
              let rB = fullData[b][k] === 'E1' || fullData[b][k] === 'E2' ? 'E' : fullData[b][k];
              if (rB === (role === 'E1' || role === 'E2' ? 'E' : role)) cB++;
            }
            if (cA !== cB) return cA - cB;

            if (role === 'N') {
                let gapA = getDaysSinceLastN(a, d);
                let gapB = getDaysSinceLastN(b, d);
                if (gapA !== gapB) return gapB - gapA; 
            }

            let critA = 0; let critB = 0;
            const cr = ['N','D','E','E1','E2','H','M1','M2','M3','P'];
            for(let k=0; k<=d; k++) { if (cr.includes(fullData[a][k])) critA++; }
            for(let k=0; k<=d; k++) { if (cr.includes(fullData[b][k])) critB++; }
            if (critA !== critB) return critA - critB;

            if (consecutiveWorkDays[a] !== consecutiveWorkDays[b]) {
                return consecutiveWorkDays[a] - consecutiveWorkDays[b];
            }

            let distA = pool.indexOf(a) !== -1 ? (pool.indexOf(a) - startPtr + pool.length) % pool.length : 999;
            let distB = pool.indexOf(b) !== -1 ? (pool.indexOf(b) - startPtr + pool.length) % pool.length : 999;
            return distA - distB;
        });

        selectedIdx = validCandidates[0]; 
        
        fullData[selectedIdx][d] = (role === 'E1' || role === 'E2') ? 'E' : role;
        assignedToday.add(selectedIdx); 
        totalWorkCounts[selectedIdx]++;
        
        let offset = pool.indexOf(selectedIdx);
        if (offset !== -1) {
           let nextPtr = (offset + 1) % pool.length;
           if (role==='N') ptrNW = nextPtr; 
           else if (role==='D') ptrD = nextPtr; 
           else if (role==='E1') ptrE1 = nextPtr;
           else if (role==='E2') ptrE2 = nextPtr; 
           else if (role==='H') ptrH = nextPtr; 
           else if (role==='M1') ptrM1 = nextPtr; 
           else if (role==='M2') ptrM2 = nextPtr;
           else if (role==='M3') ptrM3 = nextPtr; 
           else if (role==='P') ptrP = nextPtr;
        }
      }
    };

    assignRole(poolStandard, 'N');
    assignRole(poolD.length > 0 ? poolD : poolStandard, 'D'); 
    assignRole(poolStandard, 'E1', isWk); 
    assignRole(poolE_Disp, 'E2', isWk);   
    
    // =========================================================================
    // [H(조근) 맞춤형 차출 및 전날 휴식(R/E/A) 스왑 엔진]
    // =========================================================================
    if (!isSunday) {
        let alreadyAssignedH = false;
        for (let i = 0; i < employeeCount; i++) {
            if (fullData[i][d] === 'H') { alreadyAssignedH = true; break; }
        }

        if (!alreadyAssignedH) {
            let pool = poolH.length > 0 ? poolH : poolStandard;
            let assignedH = false;

            let validCandidates = [];
            for (let idx of pool) {
                if (assignedToday.has(idx)) continue;
                if (consecutiveWorkDays[idx] >= 6) continue;
                // 내일이 X(휴가), A(부재), H(조근)인 사람 배제
                if (d + 1 < diffDays) {
                    let tom = existingValues[idx][d+1].toString().toUpperCase();
                    if (tom === 'X' || tom === 'H' || tom === 'A') continue;
                }
                validCandidates.push(idx);
            }

            const sortCandsForH = (cands) => {
                cands.sort((a, b) => {
                    let hA = 0, hB = 0;
                    for(let k=0; k<=d; k++) { if(fullData[a][k]==='H') hA++; if(fullData[b][k]==='H') hB++; }
                    if(hA !== hB) return hA - hB;

                    let critA = 0, critB = 0;
                    const cr = ['N','D','E','E1','E2','H','M1','M2','M3','P'];
                    for(let k=0; k<=d; k++) { if (cr.includes(fullData[a][k])) critA++; }
                    for(let k=0; k<=d; k++) { if (cr.includes(fullData[b][k])) critB++; }
                    if (critA !== critB) return critA - critB;

                    if (consecutiveWorkDays[a] !== consecutiveWorkDays[b]) {
                        return consecutiveWorkDays[a] - consecutiveWorkDays[b];
                    }

                    let distA = pool.indexOf(a) !== -1 ? (pool.indexOf(a) - ptrH + pool.length) % pool.length : 999;
                    let distB = pool.indexOf(b) !== -1 ? (pool.indexOf(b) - ptrH + pool.length) % pool.length : 999;
                    return distA - distB;
                });
            };

            // 1차: 자연산 탐색 (전날 이미 E, R, X, W, A 인 사람)
            let phase1Cands = validCandidates.filter(idx => {
                let prev = d === 0 ? pastData[idx][3] : fullData[idx][d-1];
                return ['E', 'E1', 'E2', 'R', 'X', 'W', 'A'].includes(prev);
            });

            if (phase1Cands.length > 0) {
                sortCandsForH(phase1Cands);
                let selectedIdx = phase1Cands[0];
                fullData[selectedIdx][d] = 'H';
                assignedToday.add(selectedIdx);
                totalWorkCounts[selectedIdx]++;
                let offset = pool.indexOf(selectedIdx);
                if (offset !== -1) ptrH = (offset + 1) % pool.length;
                assignedH = true;
            }

            // 2차: R 스왑 (전날 S를 R로 깎으면서 당겨오기)
            if (!assignedH && d > 0) {
                let phase2Cands = validCandidates.filter(idx => {
                    let prev = fullData[idx][d-1];
                    if (prev !== 'S') return false; 

                    let prevDate = new Date(config.START_DATE.getTime() + (d-1)*86400000);
                    if (isRedDay(prevDate, config.HOLIDAYS)) return false; 

                    let prevDow = prevDate.getDay();
                    if (prevDow === 5 && groupData[idx] !== '파견직') {
                        let dC = 0, sC = 0;
                        for (let k = 0; k < employeeCount; k++) {
                            if (groupData[k] !== '파견직') {
                                if (fullData[k][d-1] === 'D') dC++;
                                if (fullData[k][d-1] === 'S') sC++;
                            }
                        }
                        if (dC + sC <= 3) return false; 
                    }
                    return true;
                });

                if (phase2Cands.length > 0) {
                    sortCandsForH(phase2Cands);
                    let selectedIdx = phase2Cands[0];
                    
                    fullData[selectedIdx][d-1] = 'R';
                    accumulatedRestCounts[selectedIdx]++;
                    totalWorkCounts[selectedIdx]--;
                    consecutiveWorkDays[selectedIdx] = 0; 
                    
                    fullData[selectedIdx][d] = 'H';
                    assignedToday.add(selectedIdx);
                    totalWorkCounts[selectedIdx]++;
                    let offset = pool.indexOf(selectedIdx);
                    if (offset !== -1) ptrH = (offset + 1) % pool.length;
                    assignedH = true;
                }
            }

            // 3차: E 스왑
            if (!assignedH && d > 0) {
                let phase3Cands = validCandidates.filter(idx => {
                    let prev = fullData[idx][d-1];
                    if (prev !== 'S') return false;

                    let prevDate = new Date(config.START_DATE.getTime() + (d-1)*86400000);
                    if (isRedDay(prevDate, config.HOLIDAYS)) return false; 

                    let prevDow = prevDate.getDay();
                    if (prevDow === 5 && groupData[idx] !== '파견직') {
                        let dC = 0, sC = 0;
                        for (let k = 0; k < employeeCount; k++) {
                            if (groupData[k] !== '파견직') {
                                if (fullData[k][d-1] === 'D') dC++;
                                if (fullData[k][d-1] === 'S') sC++;
                            }
                        }
                        if (dC + sC <= 3) return false; 
                    }
                    return true;
                });

                if (phase3Cands.length > 0) {
                    sortCandsForH(phase3Cands);
                    let selectedIdx = phase3Cands[0];
                    
                    fullData[selectedIdx][d-1] = 'E';
                    fullData[selectedIdx][d] = 'H';
                    assignedToday.add(selectedIdx);
                    totalWorkCounts[selectedIdx]++;
                    let offset = pool.indexOf(selectedIdx);
                    if (offset !== -1) ptrH = (offset + 1) % pool.length;
                    assignedH = true;
                }
            }

            if (!assignedH && validCandidates.length > 0) {
                sortCandsForH(validCandidates);
                let selectedIdx = validCandidates[0];
                fullData[selectedIdx][d] = 'H';
                assignedToday.add(selectedIdx);
                totalWorkCounts[selectedIdx]++;
                let offset = pool.indexOf(selectedIdx);
                if (offset !== -1) ptrH = (offset + 1) % pool.length;
                assignedH = true;
            }
        }
    }

    if (d > 0) {
        for (let i = 0; i < employeeCount; i++) {
            if (fullData[i][d] === 'H' && fullData[i][d-1] === 'S') {
                let prevDate = new Date(config.START_DATE.getTime() + (d-1)*86400000);
                if (!isRedDay(prevDate, config.HOLIDAYS)) {
                    let prevDow = prevDate.getDay();
                    let canChange = true;
                    if (prevDow === 5 && groupData[i] !== '파견직') {
                        let dC = 0, sC = 0;
                        for (let k = 0; k < employeeCount; k++) {
                            if (groupData[k] !== '파견직') {
                                if (fullData[k][d-1] === 'D') dC++;
                                if (fullData[k][d-1] === 'S') sC++;
                            }
                        }
                        if (dC + sC <= 3) canChange = false;
                    }
                    if (canChange) {
                        fullData[i][d-1] = 'R';
                        accumulatedRestCounts[i]++;
                        totalWorkCounts[i]--;
                        consecutiveWorkDays[i] = 0; 
                    }
                }
            }
        }
    }
    // =========================================================================

    assignRole(poolStandard, 'M1', false); 
    assignRole(poolStandard, 'M2', isWk); 
    assignRole(poolStandard, 'M3', !isMonToThu); 
    assignRole(poolP.length > 0 ? poolP : poolStandard, 'P', hasManualP || isWeekend);

    if (isWk) {
      const numGenNeeded = 1; const numDispNeeded = isSunday ? 2 : 3;
      const getScore = (i) => (accumulatedRestCounts[i] * 1000) - totalWorkCounts[i] - (consecutiveWorkDays[i] >= 6 ? 100000 : 0);
      
      const candidatesGen = poolStandard.filter(i => !assignedToday.has(i)).map(i => ({idx:i, s:getScore(i)})).sort((a,b)=>b.s-a.s);
      for(let k=0; k<Math.min(numGenNeeded, candidatesGen.length); k++){
        fullData[candidatesGen[k].idx][d] = 'S'; assignedToday.add(candidatesGen[k].idx); totalWorkCounts[candidatesGen[k].idx]++;
      }
      
      const candidatesDisp = poolE_Disp.filter(i => !assignedToday.has(i)).map(i => ({idx:i, s:getScore(i)})).sort((a,b)=>b.s-a.s);
      for(let k=0; k<Math.min(numDispNeeded, candidatesDisp.length); k++){
        fullData[candidatesDisp[k].idx][d] = 'S'; assignedToday.add(candidatesDisp[k].idx); totalWorkCounts[candidatesDisp[k].idx]++;
      }
      
      for (let i = 0; i < employeeCount; i++) {
         if (!assignedToday.has(i)) { 
            fullData[i][d] = 'R'; 
            assignedToday.add(i);
            accumulatedRestCounts[i]++; 
         }
      }
    } else {
      for (let i = 0; i < employeeCount; i++) {
        if (!assignedToday.has(i) && consecutiveWorkDays[i] >= 6) {
          fullData[i][d] = 'R';
          assignedToday.add(i);
          accumulatedRestCounts[i]++;
        }
      }
      for (let i = 0; i < employeeCount; i++) {
        if (!assignedToday.has(i)) { 
            fullData[i][d] = 'S'; 
            assignedToday.add(i); 
            totalWorkCounts[i]++; 
        }
      }
    }

    for (let i = 0; i < employeeCount; i++) {
      const role = fullData[i][d];
      if (['X', 'R', 'A'].includes(role)) {
        consecutiveWorkDays[i] = 0;
      } else {
        consecutiveWorkDays[i]++;
      }
    }
  }

  // =========================================================================
  // [통합 백그라운드 리밸런싱 엔진 가동]
  // =========================================================================
  if (targetOffDays !== null) {
      const empTargets = new Array(employeeCount).fill(targetOffDays).map((val, i) => val + (plusOneData[i] === 'O' || plusOneData[i] === 'ㅇ' ? 1 : 0));
      const dailyActiveCounts = new Array(diffDays).fill(0);
      for(let d=0; d<diffDays; d++) for(let i=0; i<employeeCount; i++) if(['E','S','D','P'].includes(fullData[i][d])) dailyActiveCounts[d]++;

      for(let i=0; i<employeeCount; i++) {
        let curOff = fullData[i].filter(v => v==='X'||v==='R').length; // A는 목표 휴무계산에서 제외됨
        let diff = empTargets[i] - curOff;
        
        if(diff > 0) { 
          const sortedDays = [];
          for(let d=0; d<diffDays; d++) {
            if (fullData[i][d] !== 'S') continue; 
            let p = dailyActiveCounts[d] * 10; 
            if(d>0 && fullData[i][d-1]==='W') p += 2000; 
            
            let forward = 0;
            for(let k=d+1; k<diffDays; k++) { if (['X','R','A'].includes(fullData[i][k])) break; forward++; }
            let backward = 0;
            for(let k=d-1; k>=-4; k--) { 
               let roleK = k < 0 ? pastData[i][4+k] : fullData[i][k];
               if (['X','R','A',''].includes(roleK)) break; backward++; 
            }
            let streakLength = forward + backward + 1;
            p += streakLength * 150; 
            sortedDays.push({d:d, p:p});
          }
          sortedDays.sort((a,b)=>b.p-a.p);
          
          for(let item of sortedDays) {
            if(diff <= 0) break;
            let d = item.d;
            let isWk = isRedDay(new Date(config.START_DATE.getTime() + d*86400000), config.HOLIDAYS);
            if (isWk) continue; 
            if (!isWk && dailyActiveCounts[d] <= 10) continue; 

            let dow = new Date(config.START_DATE.getTime() + d*86400000).getDay();
            if (dow === 5 && groupData[i] !== '파견직') {
                let dC = 0, sC = 0;
                for(let k=0; k<employeeCount; k++) {
                   if (groupData[k] !== '파견직') {
                       if (fullData[k][d] === 'D') dC++;
                       if (fullData[k][d] === 'S') sC++;
                   }
                }
                if (dC + sC <= 3) continue; 
            }
            
            fullData[i][d] = 'R'; 
            dailyActiveCounts[d]--; 
            diff--;
          }
        } 
        else if (diff < 0) { 
          const sortedDays = [];
          for(let d=0; d<diffDays; d++) {
             if (fullData[i][d] === 'R') sortedDays.push({day:d, count:dailyActiveCounts[d]});
          }
          sortedDays.sort((a,b)=> a.count - b.count); 

          for (let item of sortedDays) {
             if (diff >= 0) break;
             const d = item.day;
             let isWk = isRedDay(new Date(config.START_DATE.getTime() + d*86400000), config.HOLIDAYS);
             if (isWk) continue; 
             
             let forward = 0;
             for(let k=d+1; k<diffDays; k++) {
                 if (['X','R','A'].includes(fullData[i][k])) break;
                 forward++;
             }
             let backward = 0;
             for(let k=d-1; k>=-4; k--) {
                 let roleK = k < 0 ? pastData[i][4+k] : fullData[i][k];
                 if (['X','R','A',''].includes(roleK)) break;
                 backward++;
             }
             if (forward + backward + 1 > 6) continue;

             fullData[i][d] = 'S';
             let currDate = new Date(config.START_DATE.getTime() + d*86400000);
             if(!isRedDay(currDate, config.HOLIDAYS)) dailyActiveCounts[d]++;
             diff++;
          }
        }
      }

      const wouldViolateNGap = (idx, day) => {
        let hasRecentN = false;
        for(let k = Math.max(-4, day - 4); k <= Math.min(diffDays - 1, day + 4); k++) {
          if (k === day) continue;
          if (k < 0) { if (pastData[idx][4 + k] === 'N') { hasRecentN = true; break; } } 
          else { if (fullData[idx][k] === 'N') { hasRecentN = true; break; } }
        }
        return hasRecentN;
      };

      const checkFridayDS = (checkDay) => {
        if (checkDay >= diffDays) return false;
        let dow = new Date(config.START_DATE.getTime() + checkDay*86400000).getDay();
        if (dow === 5) {
            let dC = 0, sC = 0;
            for (let k = 0; k < employeeCount; k++) {
                if (groupData[k] !== '파견직') {
                    if (fullData[k][checkDay] === 'D') dC++;
                    if (fullData[k][checkDay] === 'S') sC++;
                }
            }
            return (dC + sC < 3);
        }
        return false;
      };

      let nBalanced = false;
      let nPasses = 0;
      while (!nBalanced && nPasses < 50) {
        nPasses++;
        let nCounts = new Array(employeeCount).fill(0);
        for(let i=0; i<employeeCount; i++) {
          for(let d=0; d<diffDays; d++) { if(fullData[i][d] === 'N') nCounts[i]++; }
        }

        let nList = []; let minN = 999; let maxN = -1;
        for(let i=0; i<employeeCount; i++) {
          if(groupData[i] === '파견직') continue; 
          let xCount = 0;
          for(let d=0; d<diffDays; d++) if(fullData[i][d] === 'X') xCount++;
          if (xCount > diffDays / 2) continue;

          let count = nCounts[i];
          nList.push({idx: i, count: count});
          if (count < minN) minN = count;
          if (count > maxN) maxN = count;
        }

        if (nList.length === 0 || maxN - minN <= 1) { nBalanced = true; break; }
        nList.sort((a, b) => b.count - a.count);

        let swapped = false;
        for (let donorObj of nList) {
          if (donorObj.count <= minN) continue; 
          for (let victimObj of [...nList].reverse()) {
            if (victimObj.count >= donorObj.count - 1) continue; 

            let donor = donorObj.idx; let victim = victimObj.idx;
            for (let d = 0; d < diffDays; d++) {
              if (fullData[donor][d] === 'N') {
                let donRoleD = 'N'; let donRoleNext = (d + 1 < diffDays) ? fullData[donor][d+1] : null;
                let vicRoleD = fullData[victim][d]; let vicRoleNext = (d + 1 < diffDays) ? fullData[victim][d+1] : null;

                if (['X', 'P', 'N', 'D', 'H', 'A'].includes(vicRoleD)) continue;
                if (vicRoleNext !== null && ['X', 'P', 'N', 'D', 'H', 'A'].includes(vicRoleNext)) continue;
                if (wouldViolateNGap(victim, d)) continue;

                const checkElig = (empIdx, role) => {
                  if (!role || ['S', 'R', 'W'].includes(role)) return true;
                  if (role === 'E') return (groupData[victim] === '파견직') === (groupData[empIdx] === '파견직'); 
                  if (['M1', 'M2', 'M3'].includes(role) && groupData[empIdx] !== '파견직') return true;
                  return false; 
                };

                if (!checkElig(donor, vicRoleD)) continue;
                if (vicRoleNext !== null && !checkElig(donor, vicRoleNext)) continue;

                fullData[donor][d] = vicRoleD;
                if (donRoleNext !== null) fullData[donor][d+1] = vicRoleNext !== null ? vicRoleNext : 'S';
                fullData[victim][d] = 'N';
                if (vicRoleNext !== null) fullData[victim][d+1] = 'W';

                const check6 = (idx) => {
                  let streak = 0;
                  for (let p = 0; p < 4; p++) {
                     let r = pastData[idx][p];
                     if (['X', 'R', 'A', ''].includes(r)) streak = 0; else streak++;
                  }
                  for (let k = 0; k < diffDays; k++) {
                    if (['X', 'R', 'A'].includes(fullData[idx][k])) streak = 0;
                    else { streak++; if (streak > 6) return true; }
                  }
                  return false;
                };

                let fridayIssue = false;
                if (groupData[donor] !== '파견직' && (donRoleD === 'S' || donRoleNext === 'S')) {
                    if (donRoleD === 'S' && checkFridayDS(d)) fridayIssue = true;
                    if (donRoleNext === 'S' && checkFridayDS(d+1)) fridayIssue = true;
                }

                if (check6(donor) || check6(victim) || fridayIssue) {
                  fullData[donor][d] = donRoleD;
                  if (donRoleNext !== null) fullData[donor][d+1] = donRoleNext;
                  fullData[victim][d] = vicRoleD;
                  if (vicRoleNext !== null) fullData[victim][d+1] = vicRoleNext;
                  continue;
                }
                swapped = true; break; 
              }
            }
            if (swapped) break;
          }
          if (swapped) break;
        }
        if (!swapped) break; 
      }

      let changedAny = true; let passes = 0;
      const wouldViolate6DayRule = (idx, day, newRole) => {
        if (newRole === 'X' || newRole === 'R' || newRole === 'A') return false;
        let original = fullData[idx][day]; fullData[idx][day] = newRole;
        let streak = 0; let violated = false;
        for (let p = 0; p < 4; p++) {
           let r = pastData[idx][p];
           if (['X', 'R', 'A', ''].includes(r)) streak = 0; else streak++;
        }
        for (let d = 0; d < diffDays; d++) {
          if (['X', 'R', 'A'].includes(fullData[idx][d])) streak = 0;
          else { streak++; if (streak > 6) { violated = true; break; } }
        }
        fullData[idx][day] = original; return violated;
      };

      while (changedAny && passes < 30) {
        changedAny = false; passes++;
        
        const dailyActiveCounts = new Array(diffDays).fill(0);
        for(let d=0; d<diffDays; d++) for(let i=0; i<employeeCount; i++) if(['E','S','D','P'].includes(fullData[i][d])) dailyActiveCounts[d]++;

        let restInfo = getRestCountsAndAverage(fullData, employeeCount, diffDays);
        let victims = [];
        for (let i=0; i<employeeCount; i++) { if (restInfo.restCounts[i] < empTargets[i]) victims.push(i); }

        if (victims.length === 0) break; 
        victims.sort((a, b) => (restInfo.restCounts[a] - empTargets[a]) - (restInfo.restCounts[b] - empTargets[b]));

        for (let v of victims) {
          let deficit = empTargets[v] - restInfo.restCounts[v];
          let sortedDays = [];
          for(let d=0; d<diffDays; d++) sortedDays.push({d: d, act: dailyActiveCounts[d]});
          sortedDays.sort((a,b) => b.act - a.act);

          for (let item of sortedDays) {
            let d = item.d;
            if (deficit <= 0) break;
            if (fullData[v][d] === 'S') {
              let bestDonor = -1; let maxR = -1;
              for (let i=0; i<employeeCount; i++) {
                if (i === v) continue;
                let isWk = isRedDay(new Date(config.START_DATE.getTime() + d*86400000), config.HOLIDAYS);
                let donorSurplus = restInfo.restCounts[i] - empTargets[i];
                let victimSurplus = restInfo.restCounts[v] - empTargets[v]; 
                
                if (fullData[i][d] === 'R' && donorSurplus > victimSurplus + 1) {
                  if (isWk && (groupData[v] === '파견직') !== (groupData[i] === '파견직')) continue; 
                  if (!wouldViolate6DayRule(i, d, 'S')) {
                    if (d + 1 < diffDays && fullData[i][d+1] === 'H') continue;
                    
                    if (restInfo.restCounts[i] > maxR) { maxR = restInfo.restCounts[i]; bestDonor = i; }
                  }
                }
              }
              if (bestDonor !== -1) {
                fullData[bestDonor][d] = 'S'; fullData[v][d] = 'R';
                restInfo.restCounts[bestDonor]--; restInfo.restCounts[v]++;
                deficit--; changedAny = true;
              }
            }
          }

          if (deficit > 0) {
             for (let item of sortedDays) {
                let d = item.d;
                if (deficit <= 0) break;
                let role = fullData[v][d];
                
                if (['E', 'M1', 'M2', 'M3'].includes(role)) { 
                   let bestDonor = -1; let maxScore = -1; let isDonorS = false;
                   let isWk = isRedDay(new Date(config.START_DATE.getTime() + d*86400000), config.HOLIDAYS);

                   for (let i=0; i<employeeCount; i++) {
                     if (i === v) continue;
                     
                     let isEligible = true;
                     if (['M1', 'M2', 'M3'].includes(role) && groupData[i] === '파견직') isEligible = false;
                     if (role === 'E' && (groupData[v] === '파견직' !== (groupData[i] === '파견직'))) isEligible = false;
                     if (!isEligible || wouldViolate6DayRule(i, d, role)) continue;
                     
                     let donorSurplus = restInfo.restCounts[i] - empTargets[i];
                     let victimSurplus = restInfo.restCounts[v] - empTargets[v];

                     if (fullData[i][d] === 'S' && (!isWk && dailyActiveCounts[d] > 10)) {
                        let dow = new Date(config.START_DATE.getTime() + d*86400000).getDay();
                        if (dow === 5 && groupData[i] !== '파견직') {
                            let dC = 0, sC = 0;
                            for(let k=0; k<employeeCount; k++) {
                               if (groupData[k] !== '파견직') {
                                   if (fullData[k][d] === 'D') dC++;
                                   if (fullData[k][d] === 'S') sC++;
                               }
                            }
                            if (dC + sC <= 3) continue; 
                        }

                        let score = 1000 + donorSurplus; 
                        if (score > maxScore) { maxScore = score; bestDonor = i; isDonorS = true; }
                     }
                     else if (fullData[i][d] === 'R' && donorSurplus > victimSurplus + 1) {
                        if (role !== 'E' && d + 1 < diffDays && fullData[i][d+1] === 'H') continue;
                        
                        let score = donorSurplus;
                        if (!isDonorS && score > maxScore) { maxScore = score; bestDonor = i; isDonorS = false; }
                     }
                   }

                   if (bestDonor !== -1) {
                     fullData[bestDonor][d] = role; fullData[v][d] = 'R';
                     if (isDonorS) dailyActiveCounts[d]--; else restInfo.restCounts[bestDonor]--; 
                     restInfo.restCounts[v]++; deficit--; changedAny = true;
                   }
                }
             }
          }
          
          if (deficit > 0) {
             for (let item of sortedDays) {
                let d = item.d;
                if (deficit <= 0) break;
                if (fullData[v][d] === 'N') {
                   let bestDonor = -1; let maxScore = -1; let isDonorS = false;
                   let isWk = isRedDay(new Date(config.START_DATE.getTime() + d*86400000), config.HOLIDAYS);

                   for (let i=0; i<employeeCount; i++) {
                     if (i === v) continue;
                     if (groupData[i] === '파견직') continue; 
                     
                     let currentRole = fullData[i][d];
                     if (currentRole !== 'S' && currentRole !== 'R') continue;
                     
                     let donorSurplus = restInfo.restCounts[i] - empTargets[i];
                     let victimSurplus = restInfo.restCounts[v] - empTargets[v];

                     if (currentRole === 'S' && (isWk || dailyActiveCounts[d] <= 10)) continue;
                     if (currentRole === 'R' && donorSurplus <= victimSurplus + 1) continue;

                     if (currentRole === 'S' && groupData[i] !== '파견직') {
                        let dow = new Date(config.START_DATE.getTime() + d*86400000).getDay();
                        if (dow === 5) {
                            let dC = 0, sC = 0;
                            for(let k=0; k<employeeCount; k++) {
                               if (groupData[k] !== '파견직') {
                                   if (fullData[k][d] === 'D') dC++;
                                   if (fullData[k][d] === 'S') sC++;
                               }
                            }
                            if (dC + sC <= 3) continue; 
                        }
                     }

                     if (d + 1 < diffDays && !['S', 'R', 'A'].includes(fullData[i][d+1])) continue;
                     if (wouldViolateNGap(i, d)) continue;
                     
                     let origD = fullData[i][d]; let origD1 = (d+1 < diffDays) ? fullData[i][d+1] : null;
                     fullData[i][d] = 'N'; 
                     if(d+1 < diffDays) fullData[i][d+1] = 'W';
                     let streak = 0; let viol = false;
                     
                     for (let p = 0; p < 4; p++) {
                        let r = pastData[i][p];
                        if (['X', 'R', 'A', ''].includes(r)) streak = 0; else streak++;
                     }
                     for(let k=0; k<diffDays; k++) {
                         if(['X','R','A'].includes(fullData[i][k])) streak = 0;
                         else { streak++; if(streak>6){ viol=true; break; } }
                     }
                     fullData[i][d] = origD; 
                     if(d+1 < diffDays) fullData[i][d+1] = origD1;
                     
                     if (viol) continue;

                     if (currentRole === 'S') {
                         let score = 1000 + donorSurplus;
                         if (score > maxScore) { maxScore = score; bestDonor = i; isDonorS = true; }
                     } else {
                         let score = donorSurplus;
                         if (!isDonorS && score > maxScore) { maxScore = score; bestDonor = i; isDonorS = false; }
                     }
                   }

                   if (bestDonor !== -1) {
                     fullData[bestDonor][d] = 'N';
                     if (d + 1 < diffDays) {
                         fullData[bestDonor][d+1] = 'W';
                         if (!['X','P','A'].includes(fullData[v][d+1])) fullData[v][d+1] = 'S'; 
                     }
                     fullData[v][d] = 'R';
                     
                     if (isDonorS) dailyActiveCounts[d]--; else restInfo.restCounts[bestDonor]--;
                     restInfo.restCounts[v]++; deficit--; changedAny = true;
                   }
                }
             }
          }
        }
      }
      
      let finalRest = getRestCountsAndAverage(fullData, employeeCount, diffDays);
      let stillDeficit = 0;
      for(let i=0; i<employeeCount; i++) if(finalRest.restCounts[i] < empTargets[i]) stillDeficit++;

      if (stillDeficit === 0) {
         ui.alert(`✅ 로테이션 및 리밸런싱 원클릭 완료!\n\n과거 데이터를 연동하여 숙직 및 조근(H) 보호, 목표 휴무일 분배를 완벽하게 마쳤습니다.`);
      } else {
         ui.alert(`⚠️ 로테이션 완료 (일부 휴무 미달)\n\n사전 지정 보호 및 스왑 엔진을 풀가동했으나, 스케줄 구조 한계(금요일 하한선 등)로 인해 ${stillDeficit}명은 목표 휴무일에 도달하지 못했습니다.`);
      }
  } else {
      ui.alert('✅ 로테이션 생성이 완료되었습니다.\n(목표 휴무일을 입력하지 않아 기본 배분만 진행했습니다.)');
  }

  enforce6DayRulePostAssignment(fullData, pastData, employeeCount, diffDays);

  range.setValues(fullData);
  updateHolidaySummary(activeSheet, config, employeeCount, diffDays);
  applyFormatting(activeSheet, diffDays, config, employeeCount);
}

function showExclusionDialog() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('업무 배제', '행번호, 시작일(MM/DD), 종료일(MM/DD) 입력', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    const input = res.getResponseText().split(',');
    if (input.length !== 3) return;
    const config = getAppConfig();
    const totalDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
    const sIdx = getDateIndex(input[1].trim(), config.START_DATE, totalDays);
    const eIdx = getDateIndex(input[2].trim(), config.START_DATE, totalDays);
    if (sIdx !== -1 && eIdx !== -1) applyExclusionAndStrictShift(parseInt(input[0]), sIdx, eIdx);
  }
}

function applyExclusionAndStrictShift(targetRow, startIdx, endIdx) {
  const config = getAppConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const employeeCount = getActualEmployeeCount(sheet, config.START_ROW, config.NAME_COL);
  const diffDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
  const targetIdx = targetRow - config.START_ROW;
  const range = sheet.getRange(config.START_ROW, config.DATE_START_COL, employeeCount, diffDays);
  const data = range.getValues();
  const pastData = getPastData(sheet, config, employeeCount);
  const groupData = sheet.getRange(config.START_ROW, config.GROUP_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim());
  
  const getPool = (colIdx) => {
      const poolData = sheet.getRange(config.START_ROW, colIdx, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());
      const pool = [];
      for(let i=0; i<employeeCount; i++) if(poolData[i]==='O'||poolData[i]==='ㅇ') pool.push(i);
      return pool;
  };
  const poolH = getPool(config.H_POOL_COL);
  const poolD = getPool(config.D_POOL_COL);
  const poolP = getPool(config.P_POOL_COL);

  const restInfo = getRestCountsAndAverage(data, employeeCount, diffDays);
  let failLog = [];

  for (let d = startIdx; d <= endIdx; d++) {
    const orgRole = data[targetIdx][d];
    if (orgRole === 'X' || orgRole === 'A') continue;
    data[targetIdx][d] = 'X';

    if (CRITICAL_ROLES.includes(orgRole)) {
      const success = reassignRoleStrict(data, pastData, targetIdx, d, orgRole, groupData, config, diffDays, poolH, poolD, poolP, restInfo);
      if (!success) failLog.push(`${d+1}일 ${orgRole} 이관 불가 (휴무 보장 규칙 등 원인)`);
    }

    let currDate = new Date(config.START_DATE.getTime() + d*86400000);
    if (!isRedDay(currDate, config.HOLIDAYS)) {
      let esCount = 0; let esdCount = 0;
      for(let i=0; i<employeeCount; i++) {
        if(['E','S'].includes(data[i][d])) esCount++;
        if(['E','S','D','P'].includes(data[i][d])) esdCount++;
      }
      
      while (esCount < 9 || esdCount < 10) {
        let bestR = -1;
        let maxR = -1;
        for(let i=0; i<employeeCount; i++) {
          if(data[i][d] === 'R' && i !== targetIdx) { 
            if (restInfo.restCounts[i] > maxR) {
              maxR = restInfo.restCounts[i];
              bestR = i;
            }
          } 
        }
        
        if (bestR !== -1 && restInfo.restCounts[bestR] - 1 >= restInfo.avgRestDays - 1) {
          data[bestR][d] = 'S'; 
          restInfo.restCounts[bestR]--;
          esCount++; 
          esdCount++;
        } else {
          failLog.push(`${d+1}일 가동 인원 부족 보충 실패 (휴무 밸런스 규칙 충돌)`);
          break;
        }
      }
    }
  }

  enforce6DayRulePostAssignment(data, pastData, employeeCount, diffDays);

  range.setValues(data);
  updateHolidaySummary(sheet, config, employeeCount, diffDays);
  applyFormatting(sheet, diffDays, config, employeeCount);
  SpreadsheetApp.getUi().alert(failLog.length > 0 ? '⚠️ 알림:\n' + failLog.join('\n') : '✅ 업무 배제 및 보직 재배치 완료');
}

function showDeskAssignmentDialog() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('데스크 지정', '행번호, 시작일(MM/DD), 종료일(MM/DD) 입력', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    const input = res.getResponseText().split(',');
    if (input.length !== 3) return;
    const config = getAppConfig();
    const totalDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
    const sIdx = getDateIndex(input[1].trim(), config.START_DATE, totalDays);
    const eIdx = getDateIndex(input[2].trim(), config.START_DATE, totalDays);
    if (sIdx !== -1 && eIdx !== -1) applyFixedRolePoolStrict(parseInt(input[0]), sIdx, eIdx, 'D', config.D_POOL_COL);
  }
}

function showHAssignmentDialog() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('H(조근) 지정', '행번호, 시작일(MM/DD), 종료일(MM/DD) 입력', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    const input = res.getResponseText().split(',');
    if (input.length !== 3) return;
    const config = getAppConfig();
    const totalDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
    const sIdx = getDateIndex(input[1].trim(), config.START_DATE, totalDays);
    const eIdx = getDateIndex(input[2].trim(), config.START_DATE, totalDays);
    if (sIdx !== -1 && eIdx !== -1) applyFixedRolePoolStrict(parseInt(input[0]), sIdx, eIdx, 'H', config.H_POOL_COL);
  }
}

function applyFixedRolePoolStrict(targetRow, startIdx, endIdx, roleToAssign, poolColIdx) {
  const config = getAppConfig(); 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const employeeCount = getActualEmployeeCount(sheet, config.START_ROW, config.NAME_COL);
  const diffDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
  const targetIdx = targetRow - config.START_ROW;
  
  if (targetIdx < 0 || targetIdx >= employeeCount) {
    SpreadsheetApp.getUi().alert('유효하지 않은 행 번호입니다.'); return;
  }

  const range = sheet.getRange(config.START_ROW, config.DATE_START_COL, employeeCount, diffDays);
  const data = range.getValues();
  const pastData = getPastData(sheet, config, employeeCount);
  const groupData = sheet.getRange(config.START_ROW, config.GROUP_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim());
  
  const getPool = (colIdx) => {
      const poolData = sheet.getRange(config.START_ROW, colIdx, employeeCount, 1).getValues().map(r => r[0].toString().trim().toUpperCase());
      const pool = [];
      for(let i=0; i<employeeCount; i++) if(poolData[i]==='O'||poolData[i]==='ㅇ') pool.push(i);
      return pool;
  };
  
  const poolH = getPool(config.H_POOL_COL);
  const poolD = getPool(config.D_POOL_COL);
  const poolP = getPool(config.P_POOL_COL);
  
  const targetPool = getPool(poolColIdx);
  const restInfo = getRestCountsAndAverage(data, employeeCount, diffDays);
  let failLog = [];

  for (let d = startIdx; d <= endIdx; d++) {
    let currDate = new Date(config.START_DATE.getTime() + d * 86400000);
    if (roleToAssign === 'H' && currDate.getDay() === 0) continue; 

    const orgRole = data[targetIdx][d];
    if (orgRole === roleToAssign) continue;

    for (let i = 0; i < employeeCount; i++) {
      if (data[i][d] === roleToAssign && i !== targetIdx) { data[i][d] = 'S'; break; }
    }

    if (orgRole === 'X' || orgRole === 'R' || orgRole === 'A') {
      let found = false;
      if (targetPool.length > 0) {
        let candidates = targetPool.filter(i => i !== targetIdx);
        
        candidates.sort((a, b) => {
          let countA = 0; let countB = 0;
          for (let k = 0; k < diffDays; k++) {
            if (data[a][k] === roleToAssign) countA++;
            if (data[b][k] === roleToAssign) countB++;
          }
          if (countA !== countB) return countA - countB;

          const isAS = data[a][d] === 'S';
          const isBS = data[b][d] === 'S';
          
          const check6 = (idx) => {
            if (data[idx][d] !== 'R') return false;
            let f = 0; for(let k=d+1; k<diffDays; k++) { if(['X','R','A'].includes(data[idx][k])) break; f++; }
            let b_ = 0; 
            for(let k=d-1; k>=-4; k--) { 
                let roleK = k < 0 ? pastData[idx][4+k] : data[idx][k];
                if(['X','R','A',''].includes(roleK)) break; b_++; 
            }
            return (f + b_ + 1) > 6;
          };

          let violA = check6(a) ? 1 : 0;
          let violB = check6(b) ? 1 : 0;
          if (violA !== violB) return violA - violB;

          if (isAS && !isBS) return -1;
          if (!isAS && isBS) return 1;
          
          if (!isAS && !isBS && restInfo) {
            return restInfo.restCounts[b] - restInfo.restCounts[a];
          }
          return 0;
        });

        for (let subIdx of candidates) {
          let phase = data[subIdx][d];
          if (phase !== 'S' && phase !== 'R') continue;
          
          if (phase === 'R') {
            if (restInfo.restCounts[subIdx] - 1 < restInfo.avgRestDays - 1) continue;
          }

          data[subIdx][d] = roleToAssign;
          if (phase === 'R') restInfo.restCounts[subIdx]--;
          found = true;
          break;
        }
      }
      
      if (!found) {
        failLog.push(`${d+1}일: ${roleToAssign} 가용 인원 부족 (휴무일 보장 방어 작동)`);
      }
      continue; 
    }

    if (CRITICAL_ROLES.includes(orgRole)) {
      if (orgRole !== 'W') {
        let success = reassignRoleStrict(data, pastData, targetIdx, d, orgRole, groupData, config, diffDays, poolH, poolD, poolP, restInfo);
        if (!success) failLog.push(`${d+1}일: 기존 보직(${orgRole}) 대근자 부족 (N간격/휴무 보장 등 제약)`);
      }
    } 
    
    data[targetIdx][d] = roleToAssign;
  }

  if (targetPool.length > 1) {
    const membersToBalance = targetPool.filter(idx => idx !== targetIdx);
    if (membersToBalance.length > 1) {
      let improved = true;
      let loopLimit = 0;
      while (improved && loopLimit < 100) {
        improved = false;
        loopLimit++;

        const counts = membersToBalance.map(idx => {
          let c = 0;
          for (let d = 0; d < diffDays; d++) if (data[idx][d] === roleToAssign) c++;
          return { idx, count: c };
        });

        counts.sort((a, b) => a.count - b.count);
        
        let swapped = false;
        for (let i = counts.length - 1; i >= 0 && !swapped; i--) {
          for (let j = 0; j < counts.length && !swapped; j++) {
            if (i <= j) continue;
            if (counts[i].count - counts[j].count > 1) {
              let maxIdx = counts[i].idx;
              let minIdx = counts[j].idx;
              
              for (let d = 0; d < diffDays; d++) {
                if (data[maxIdx][d] === roleToAssign && data[minIdx][d] === 'S') {
                  data[maxIdx][d] = 'S';
                  data[minIdx][d] = roleToAssign;
                  swapped = true;
                  improved = true;
                  break;
                }
              }
            }
          }
        }
      }
    }
  }

  enforce6DayRulePostAssignment(data, pastData, employeeCount, diffDays);

  range.setValues(data); 
  updateHolidaySummary(sheet, config, employeeCount, diffDays); 
  applyFormatting(sheet, diffDays, config, employeeCount);

  if (failLog.length > 0) {
    SpreadsheetApp.getUi().alert(`⚠️ 일부 지정 실패 발생:\n${failLog.join('\n')}\n\n* 그 외 배정 및 밸런싱은 완료되었습니다.`);
  } else {
    SpreadsheetApp.getUi().alert(`${targetRow}행 인원의 ${roleToAssign} 고정 배치가 완료되었습니다.\n(풀 내 밸런싱 적용 완료)`);
  }
}

function reassignRoleStrict(data, pastData, targetIdx, d, roleToShift, groupData, config, diffDays, poolH = [], poolD = [], poolP = [], restInfo = null) {
  const employeeCount = data.length;
  let found = false;

  let candidateIndices = [];
  for (let step = 1; step < employeeCount; step++) {
    candidateIndices.push((targetIdx + step) % employeeCount);
  }

  const phases = ['S', 'R'];
  let relaxationLevel = 0;

  while (relaxationLevel <= 6) {
    for (let phaseRole of phases) {
      let validCandidates = [];

      for (let r of candidateIndices) {
        const currentRole = data[r][d];
        if (currentRole !== phaseRole) continue;

        const subGroup = groupData[r];
        
        if (relaxationLevel < 6) {
            if (roleToShift === 'D' && poolD.length > 0 && !poolD.includes(r)) continue;
            if (roleToShift === 'H' && poolH.length > 0 && !poolH.includes(r)) continue;
            if (roleToShift === 'P' && poolP.length > 0 && !poolP.includes(r)) continue;
            if (['N', 'M1', 'M2', 'M3'].includes(roleToShift) && subGroup === '파견직') continue;
            if (roleToShift === 'E') {
              const isOriginalDisp = (groupData[targetIdx] === '파견직');
              if (isOriginalDisp !== (subGroup === '파견직')) continue;
            }
        }
        
        if (roleToShift === 'N') {
          if (d > 0 && data[r][d-1] === 'N') continue; 
          
          let nextRole = (d + 1 < diffDays) ? data[r][d+1] : null;
          if (nextRole === 'X' || nextRole === 'N' || nextRole === 'P' || nextRole === 'A') continue;
          
          if (relaxationLevel <= 3 && nextRole !== 'S' && nextRole !== 'R') continue;
          
          let minGap = 5;
          if (relaxationLevel <= 1) minGap = 5;
          else if (relaxationLevel === 2) minGap = 4;
          else if (relaxationLevel === 3 || relaxationLevel === 4) minGap = 3;
          else minGap = 2; 
          
          let hasRecentN = false;
          for(let k = Math.max(-4, d - minGap); k <= Math.min(diffDays-1, d + minGap); k++) {
            if (k === d) continue;
            let roleK = k < 0 ? pastData[r][4+k] : data[r][k];
            if (roleK === 'N') { hasRecentN = true; break; }
          }
          if (hasRecentN) continue;
        }

        if (phaseRole === 'R' && restInfo && relaxationLevel === 0) {
          if (restInfo.restCounts[r] - 1 < restInfo.avgRestDays - 1) continue;
        }

        validCandidates.push(r);
      }

      if (validCandidates.length > 0) {
        validCandidates.sort((a, b) => {
          let countA = 0; let countB = 0;
          for (let k = 0; k < diffDays; k++) {
            if (data[a][k] === roleToShift) countA++;
            if (data[b][k] === roleToShift) countB++;
          }
          if (countA !== countB) return countA - countB; 

          if (phaseRole === 'R' && restInfo) {
            let fA = 0; for(let k=d+1; k<diffDays; k++) { if(['X','R','A'].includes(data[a][k])) break; fA++; }
            let bA = 0; 
            for(let k=d-1; k>=-4; k--) { 
                let roleK = k < 0 ? pastData[a][4+k] : data[a][k];
                if(['X','R','A',''].includes(roleK)) break; bA++; 
            }
            let violA = (fA + bA + 1 > 6) ? 1 : 0;

            let fB = 0; for(let k=d+1; k<diffDays; k++) { if(['X','R','A'].includes(data[b][k])) break; fB++; }
            let bB = 0; 
            for(let k=d-1; k>=-4; k--) { 
                let roleK = k < 0 ? pastData[b][4+k] : data[b][k];
                if(['X','R','A',''].includes(roleK)) break; bB++; 
            }
            let violB = (fB + bB + 1 > 6) ? 1 : 0;

            if (violA !== violB) return violA - violB;
            return restInfo.restCounts[b] - restInfo.restCounts[a]; 
          }
          return 0;
        });

        for (let r of validCandidates) {
           const substituteOrgRole = data[r][d];
           const isR = substituteOrgRole === 'R';

           if (substituteOrgRole === 'S' && groupData[r] !== '파견직') {
              let dow = new Date(config.START_DATE.getTime() + d*86400000).getDay();
              if (dow === 5) {
                  let dC = 0, sC = 0;
                  for(let k=0; k<employeeCount; k++) {
                      if (groupData[k] !== '파견직') {
                          if (data[k][d] === 'D') dC++;
                          if (data[k][d] === 'S') sC++;
                      }
                  }
                  if (dC + sC <= 3) continue; 
              }
           }

           data[r][d] = roleToShift;
           if (isR && restInfo) restInfo.restCounts[r]--; 
           
           if (roleToShift === 'N' && d + 1 < diffDays) {
             const substituteNextRole = data[r][d+1]; 
             data[r][d+1] = 'W'; 
             
             if (!['X', 'R', 'P', 'A'].includes(data[targetIdx][d+1])) {
               data[targetIdx][d+1] = 'S'; 
             }

             if (CRITICAL_ROLES.includes(substituteNextRole) && substituteNextRole !== 'W') {
                reassignRoleStrict(data, pastData, targetIdx, d+1, substituteNextRole, groupData, config, diffDays, poolH, poolD, poolP, restInfo);
             }
           }
           found = true; 
           break;
        }
        if (found) break;
      }
    }
    if (found) break;
    relaxationLevel++; 
  }
  return found;
}

function validateSchedule() {
  const config = getAppConfig(); const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const employeeCount = getActualEmployeeCount(sheet, config.START_ROW, config.NAME_COL);
  const diffDays = Math.ceil((config.END_DATE - config.START_DATE) / 86400000) + 1;
  const names = sheet.getRange(config.START_ROW, config.NAME_COL, employeeCount, 1).getValues();
  const data = sheet.getRange(config.START_ROW, config.DATE_START_COL, employeeCount, diffDays).getValues();
  const pastData = getPastData(sheet, config, employeeCount);
  const groupData = sheet.getRange(config.START_ROW, config.GROUP_COL, employeeCount, 1).getValues().map(r => r[0].toString().trim());

  let errors = [];

  for (let d = 0; d < diffDays; d++) {
    let counts = { 'H':0, 'M1':0, 'M2':0, 'M3':0, 'N':0, 'W':0, 'E':0, 'D':0, 'P':0, 'S':0 };
    for (let r = 0; r < employeeCount; r++) if (counts.hasOwnProperty(data[r][d])) counts[data[r][d]]++;

    let curr = new Date(config.START_DATE.getTime() + d*86400000);
    const dateStr = `${curr.getMonth()+1}/${curr.getDate()}`;
    const dayOfWeek = curr.getDay();
    const isWk = isRedDay(curr, config.HOLIDAYS);
    const isSunday = (dayOfWeek === 0);
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
    const isMonToThu = (dayOfWeek >= 1 && dayOfWeek <= 4);

    if (!isWk) {
      if (counts['E'] + counts['S'] + counts['D'] + counts['P'] < 10) errors.push(`${dateStr}일: 가동 인원 부족 (최소 10)`);
      if (counts['E'] + counts['S'] < 8) errors.push(`${dateStr}일: ES 합계 부족 (최소 8)`); 
    }

    if (dayOfWeek === 5) {
      let genD = 0, genS = 0;
      for (let r = 0; r < employeeCount; r++) {
        if (groupData[r] !== '파견직') {
          if (data[r][d] === 'D') genD++;
          if (data[r][d] === 'S') genS++;
        }
      }
      if (genD + genS < 3) errors.push(`${dateStr}일: 금요일 일반/전문직 D+S 합계 부족 (최소 3명 필요, 현재 ${genD + genS}명)`);
    }

    for (let role of ['H', 'M1', 'M2', 'M3', 'N', 'D', 'P']) {
      let exp = (isWk && config.WEEKEND_EXCLUDE_ROLES.includes(role)) ? 0 : 1;
      if (role === 'H' && isSunday) exp = 0;
      if (role === 'P' && isWeekend) exp = 0;
      if (role === 'M3' && !isMonToThu) exp = 0;
      if (role === 'M2' && isWk) exp = 0;

      if (counts[role] < exp) {
        errors.push(`${dateStr}일: ${role} 누락`);
      }
      if ((role === 'H' || role === 'M3') && counts[role] > exp) {
        errors.push(`${dateStr}일: ${role} 초과 배정 (현재 ${counts[role]}명)`);
      }
    }
  }

  for (let i = 0; i < employeeCount; i++) {
    let consec = 0;
    let lastN = -999;
    
    for (let p = 0; p < 4; p++) {
       let r = pastData[i][p];
       if (['X', 'R', 'A', ''].includes(r)) consec = 0;
       else consec++;
       if (r === 'N') lastN = p - 4; 
    }

    for (let d = 0; d < diffDays; d++) {
      const role = data[i][d];
      
      if (role === 'H') {
          let prev = d === 0 ? pastData[i][3] : data[i][d-1];
          if (!['E', 'E1', 'E2', 'R', 'X', 'W', 'A'].includes(prev)) {
              let curr = new Date(config.START_DATE.getTime() + d*86400000);
              const dateStr = `${curr.getMonth()+1}/${curr.getDate()}`;
              errors.push(`조근(H) 보호 위반: ${names[i][0]}님 ${dateStr}일 조근 전날 휴식(E/R/X/W/A) 미확보`);
          }
      }

      if (role === 'N') {
        if (lastN !== -999 && (d - lastN) <= 5) {
          let curr = new Date(config.START_DATE.getTime() + d*86400000);
          const dateStr = `${curr.getMonth()+1}/${curr.getDate()}`;
          errors.push(`N간격 위반: ${names[i][0]}님 ${dateStr}일 기준 최소 4일 휴식 미확보`);
        }
        lastN = d;
      }

      if (['X', 'R', 'A'].includes(role)) {
        consec = 0;
      } else {
        consec++;
        if (consec > 6) {
          let curr = new Date(config.START_DATE.getTime() + d*86400000);
          const dateStr = `${curr.getMonth()+1}/${curr.getDate()}`;
          errors.push(`연속근무 위반: ${names[i][0]}님 ${dateStr}일 기준 7일 연속 출근 (W 포함)`);
        }
      }
    }
  }

  if (errors.length === 0) {
    SpreadsheetApp.getUi().alert('✅ 모든 배치가 완벽합니다.');
  } else {
    SpreadsheetApp.getUi().alert('❌ 오류 발생 내역:\n\n' + errors.slice(0, 15).join('\n'));
  }
}

function applyFormatting(sheet, diffDays, config, employeeCount) {
  if (employeeCount <= 0) return;
  const range = sheet.getRange(config.START_ROW, config.PAST_DAYS_COL, employeeCount, diffDays + 4);
  sheet.clearConditionalFormatRules();
  
  const sCountFormulas = [];
  for (let d = 0; d < diffDays; d++) {
    let curr = new Date(config.START_DATE.getTime() + d*86400000);
    if (isRedDay(curr, config.HOLIDAYS)) sheet.getRange(config.START_ROW, config.DATE_START_COL + d, employeeCount, 1).setBackground('#fff2f2');
    
    const colLetter = getColumnLetter(config.DATE_START_COL + d);
    sCountFormulas.push(`=COUNTIF(${colLetter}${config.START_ROW}:${colLetter}200, "S") + COUNTIF(${colLetter}${config.START_ROW}:${colLetter}200, "E")`);
  }
  
  sheet.getRange(1, config.DATE_START_COL, 1, diffDays)
       .setFormulas([sCountFormulas])
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setBackground('#fff2cc')
       .setFontColor('#b45f06');
       
  sheet.getRange(1, config.SUMMARY_COL)
       .setValue("E+S합계")  
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setBackground('#fff2cc');

  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('N').setBackground('#cfe2f3').setFontColor('#0b5394').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('D').setBackground('#d0e0e3').setFontColor('#134f5c').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('W').setBackground('#f3f3f3').setFontColor('#999999').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('H').setBackground('#d9ead3').setFontColor('#38761d').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('E').setBackground('#fce5cd').setFontColor('#b45f06').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('M1').setBackground('#d9d2e9').setFontColor('#351c75').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('M2').setBackground('#ead1dc').setFontColor('#741b47').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('M3').setBackground('#d5a6bd').setFontColor('#4c1130').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('P').setBackground('#fff2cc').setFontColor('#b45f06').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('X').setBackground('#f4cccc').setFontColor('#a61c00').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('R').setBackground('#f4cccc').setFontColor('#a61c00').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('A').setBackground('#f4cccc').setFontColor('#a61c00').setBold(true).setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}
