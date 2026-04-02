// Code.gs

// --- [유틸리티 함수] ---
function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(), diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

function padTime(t) {
  if (!t) return "";
  if (t instanceof Date) return Utilities.formatDate(t, "Asia/Seoul", "HH:mm");
  const parts = t.toString().split(":");
  return parts.length >= 2 ? parts[0].padStart(2, '0') + ":" + parts[1].padStart(2, '0') : t.toString().trim();
}

function getMinutes(timeStr) {
  if(!timeStr) return -1;
  const [h, m] = padTime(timeStr).split(':').map(Number);
  return h * 60 + m;
}

function getSemesterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const semester = (now.getMonth() + 1 >= 3 && now.getMonth() + 1 <= 8) ? "1학기" : "2학기";
  const sheetName = `상벌점기록_${now.getFullYear()}_${semester}`;
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // 🔥 I열 끝에 '입력유무' 헤더 추가
    sheet.appendRow(["타임스탬프", "날짜", "일정종류", "대상학번", "대상이름", "벌점항목", "점수", "기록자", "입력유무"]);
  }
  return sheet;
}

// --- [메인 라우터] ---
function doPost(e) {
  const action = e.parameter.action;
  let result = { success: false, message: "알 수 없는 요청입니다." };

  try {
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    
    if (action === 'login') result = doLogin(e.parameter.id, e.parameter.pw);
    else if (action === 'getInitialData') result = getInitialData();
    else if (action === 'changePw') result = changePw(e.parameter.id, e.parameter.oldPw, e.parameter.newPw);
    else if (action === 'submitEntry') result = processEntry(payload);
    else if (action === 'checkAttendance') result = markAttendance(payload);
    else if (action === 'getSchedule') result = getSchedule(e.parameter.startDate);
    else if (action === 'uploadSchedule') result = uploadSchedule(payload);
    else if (action === 'updateSingleSchedule') result = updateSingleSchedule(payload);
    else if (action === 'getMyInfo') result = getMyInfo(payload);
    else if (action === 'getAdminData') result = getAdminData(e.parameter.dateStr);
    
    // 게시판 관련 라우터
    else if (action === 'getBoardData') result = getBoardData();
    else if (action === 'savePost') result = savePost(payload);
    else if (action === 'updatePost') result = updatePost(payload);
    else if (action === 'deletePost') result = deletePost(payload);
    else if (action === 'saveComment') result = saveComment(payload);
    else if (action === 'saveWeeklySchedule') result = saveWeeklySchedule(payload); // 🔥 추가된 줄
    else if (action === 'getUnattendedStats') result = getUnattendedStats(payload);
    else if (action === 'deleteEntry') result = deleteEntry(payload);
    else if (action === 'saveMissingNote') result = saveMissingNote(payload);
  } catch(error) { result.message = "서버 에러: " + error.toString(); }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- [인증 및 출석 연동] ---
function doLogin(id, pw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("계정관리");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id && data[i][1].toString() === pw) {
      sheet.getRange(i + 1, 6).setValue(new Date());
      const user = { id: id, hakbun: data[i][2].toString(), name: data[i][3].toString(), role: data[i][4].toString() };
      checkAutoAttendance(user);
      return { success: true, user: user };
    }
  }
  return { success: false, message: "아이디 또는 비밀번호가 일치하지 않습니다." };
}

function checkAutoAttendance(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const attendSheet = ss.getSheetByName("출석기록");
  if(!scheduleSheet || !attendSheet) return;

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const nowMin = now.getHours() * 60 + now.getMinutes();

  const schedules = scheduleSheet.getDataRange().getValues().slice(1);
  for(let row of schedules) {
    if(!row[0]) continue;
    let rDate = row[0] instanceof Date ? Utilities.formatDate(row[0], "Asia/Seoul", "yyyy-MM-dd") : row[0].toString().trim();
    if(rDate === todayStr && row[3].toString() === user.hakbun) {
      let startMin = getMinutes(row[5]);
      let endMin = getMinutes(row[6]);
      if(nowMin >= startMin - 5 && nowMin <= endMin + 5) {
        const attData = attendSheet.getDataRange().getValues();
        let already = attData.some(a => a[0] instanceof Date ? Utilities.formatDate(a[0],"Asia/Seoul","yyyy-MM-dd") === todayStr : a[0].toString() === todayStr && a[2].toString() === user.name && a[1].toString() === row[2].toString());
        if(!already) attendSheet.appendRow([todayStr, row[2].toString(), user.name, padTime(`${now.getHours()}:${now.getMinutes()}`), "로그인자동", "정상"]);
      }
    }
  }
}

function markAttendance(payload) {
  const { user, excuse } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const attendSheet = ss.getSheetByName("출석기록");
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const nowMin = now.getHours() * 60 + now.getMinutes();

  const schedules = scheduleSheet.getDataRange().getValues().slice(1).filter(r => {
    let d = r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Seoul", "yyyy-MM-dd") : r[0].toString().trim();
    return d === todayStr && r[3].toString() === user.hakbun;
  });

  if(schedules.length === 0) return { success: false, message: "오늘 배정된 일정이 없습니다." };

  let closestSchedule = schedules.reduce((prev, curr) => Math.abs(getMinutes(curr[6]) - nowMin) < Math.abs(getMinutes(prev[6]) - nowMin) ? curr : prev);

  const actName = closestSchedule[2].toString();
  const startMin = getMinutes(closestSchedule[5]);
  const endMin = getMinutes(closestSchedule[6]);

  const attData = attendSheet.getDataRange().getValues();
  let already = attData.some(a => {
    let ad = a[0] instanceof Date ? Utilities.formatDate(a[0],"Asia/Seoul","yyyy-MM-dd") : a[0].toString().trim();
    return ad === todayStr && a[2].toString() === user.name && a[1].toString() === actName;
  });

  if(already) return { success: true, message: "이미 출석체크 되었습니다." };

  const isValidTime = (nowMin >= startMin - 5 && nowMin <= endMin + 5);
  if(!isValidTime && !excuse) return { success: false, requireExcuse: true, message: "담당 시간이 아닙니다. 사유를 입력해야 합니다." };

  attendSheet.appendRow([todayStr, actName, user.name, padTime(`${now.getHours()}:${now.getMinutes()}`), excuse ? "지연(수동)" : "수동출석", excuse || "정상"]);
  return { success: true, message: excuse ? "사유와 함께 출석이 기록되었습니다." : "출석이 완료되었습니다." };
}

// --- [상벌점 기록 및 시간 제어] ---
function processEntry(payload) {
  const { user, type, rule, students } = payload;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const nowMin = now.getHours() * 60 + now.getMinutes();

  if(user.role === 'YP') {
    const scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("YP일정");
    const schedules = scheduleSheet.getDataRange().getValues().slice(1).filter(r => {
      let d = r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Seoul", "yyyy-MM-dd") : r[0].toString().trim();
      return d === todayStr && r[3].toString() === user.hakbun;
    });

    let canSubmit = false;
    for(let s of schedules) {
      let startMin = getMinutes(s[5]);
      let endMin = getMinutes(s[6]);
      if(nowMin >= startMin && nowMin <= endMin + 5) { canSubmit = true; break; }
    }
    if(!canSubmit) return { success: false, message: "본인의 담당 시간(+5분)에만 기록이 가능합니다." };
  }

  const recordSheet = getSemesterSheet();
  const [ruleName, ruleScore] = rule.split('|');
  const existingData = recordSheet.getDataRange().getValues();
  let addedCount = 0;

  for (let student of students) {
    let isDuplicate = existingData.some(r => {
      if(r[0] === "타임스탬프") return false;
      let recDate = r[1] instanceof Date ? Utilities.formatDate(r[1], "Asia/Seoul", "yyyy-MM-dd") : r[1].toString();
      return recDate === todayStr && r[5] === ruleName && r[3].toString() === student.hakbun;
    });

    if (!isDuplicate) {
      recordSheet.appendRow([now, todayStr, "지정일정", student.hakbun, student.name, ruleName, ruleScore, user.name, "미입력"]);
      addedCount++;
    }
  }

  return addedCount > 0 ? { success: true, message: `${addedCount}건이 기록되었습니다.` } : { success: false, message: "모두 이미 기록된 중복 데이터입니다." };
}

// 🔥 상벌점 기록 삭제
function deleteEntry(payload) {
  const { user, date, target, rule, author } = payload;
  
  if (user.role !== '관리자') {
    return { success: false, message: "삭제 권한이 없습니다." };
  }

  const recordSheet = getSemesterSheet();
  if(!recordSheet) return { success: false, message: "상벌점 시트가 없습니다." };

  const data = recordSheet.getDataRange().getValues();
  // 뒤에서부터 탐색하여 역순으로 지움
  for (let i = data.length - 1; i >= 1; i--) {
    let r = data[i];
    let d = r[1] instanceof Date ? Utilities.formatDate(r[1], "Asia/Seoul", "yyyy-MM-dd") : r[1].toString();
    
    // 공백 포함 정확한 매칭을 위해 조합
    let rTarget = `${r[3]} ${r[4]}`;
    let rRule = r[5].toString();
    let rAuthor = r[7].toString();

    if (d === date && rTarget === target && rRule === rule && rAuthor === author) {
      recordSheet.deleteRow(i + 1);
      return { success: true, message: "선택한 기록이 삭제되었습니다." };
    }
  }

  return { success: false, message: "해당 일치하는 기록을 시트에서 찾을 수 없습니다." };
}

// 🔥 관리자용 미출석 비고 저장
function saveMissingNote(payload) {
  const { user, date, activity, hakbun, note } = payload;
  if(user.role !== '관리자') return { success: false, message: "접근 권한이 없습니다." };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let noteSheet = ss.getSheetByName("미출석비고");
  if(!noteSheet) {
    noteSheet = ss.insertSheet("미출석비고");
    noteSheet.appendRow(["날짜", "활동", "학번", "비고", "마지막수정"]);
  }
  
  const data = noteSheet.getDataRange().getValues();
  let updated = false;
  
  for(let i = 1; i < data.length; i++) {
    let rDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "Asia/Seoul", "yyyy-MM-dd") : data[i][0].toString().trim();
    if(rDate === date && data[i][1].toString() === activity && data[i][2].toString() === hakbun) {
      if(!note || note.trim() === "") {
        noteSheet.deleteRow(i + 1); // 내용이 없으면 행 삭제 (불필요한 데이터 정리)
      } else {
        noteSheet.getRange(i + 1, 4).setValue(note);
        noteSheet.getRange(i + 1, 5).setValue(new Date());
      }
      updated = true;
      break;
    }
  }
  
  if(!updated && note && note.trim() !== "") {
    noteSheet.appendRow([date, activity, hakbun, note, new Date()]);
  }
  
  return { success: true, message: "비고가 저장되었습니다." };
}

// --- [기존 조회, 일정, 설정 기능들] ---
function getMyInfo(payload) {
  const { user } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = getSemesterSheet();
  const myRecords = recordSheet.getDataRange().getValues().slice(1).filter(r => r[7].toString() === user.name).map(r => ({
      date: r[1] instanceof Date ? Utilities.formatDate(r[1],"Asia/Seoul","yyyy-MM-dd") : r[1].toString(),
      target: `${r[3]} ${r[4]}`, rule: r[5], score: r[6]
    })).reverse();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const mySchedules = scheduleSheet.getDataRange().getValues().slice(1).filter(r => r[3].toString() === user.hakbun).map(r => ({
      date: r[0] instanceof Date ? Utilities.formatDate(r[0],"Asia/Seoul","yyyy-MM-dd") : r[0].toString().trim(),
      day: r[1], activity: r[2], time: `${padTime(r[5])}~${padTime(r[6])}`
    })).sort((a,b) => new Date(a.date) - new Date(b.date));
  return { success: true, records: myRecords, schedules: mySchedules };
}

function getAdminData(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetDate = dateStr || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  const scheduleSheet = ss.getSheetByName("YP일정");
  const attendSheet = ss.getSheetByName("출석기록");
  let attendanceList = [];
  if(scheduleSheet) {
    const dailySchedules = scheduleSheet.getDataRange().getValues().slice(1).filter(r => {
      let d = r[0] instanceof Date ? Utilities.formatDate(r[0],"Asia/Seoul","yyyy-MM-dd") : r[0].toString().trim();
      return d === targetDate;
    });
    const attData = attendSheet ? attendSheet.getDataRange().getValues().slice(1).filter(a => {
      let d = a[0] instanceof Date ? Utilities.formatDate(a[0],"Asia/Seoul","yyyy-MM-dd") : a[0].toString().trim();
      return d === targetDate;
    }) : [];
    attendanceList = dailySchedules.map(s => {
      const name = s[4].toString(); const act = s[2].toString();
      const match = attData.find(a => a[2].toString() === name && a[1].toString() === act);
      return { activity: act, time: `${padTime(s[5])}~${padTime(s[6])}`, name: name, status: match ? match[4] : "미출석", timeLogged: match ? padTime(match[3]) : "-", excuse: match ? match[5] : "" };
    });
  }
  const recordSheet = getSemesterSheet();
  const allRecords = recordSheet.getDataRange().getValues().slice(1).filter(r => {
    let d = r[1] instanceof Date ? Utilities.formatDate(r[1],"Asia/Seoul","yyyy-MM-dd") : r[1].toString();
    return d === targetDate;
  }).map(r => ({ target: `${r[3]} ${r[4]}`, rule: r[5], score: r[6], author: r[7] })).reverse();
  return { success: true, attendance: attendanceList, records: allRecords };
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("학생명단");
  let students = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => ({ hakbun: r[0].toString(), name: r[1].toString() })) : [];
  const ruleSheet = ss.getSheetByName("상벌점규정");
  let rules = ruleSheet ? ruleSheet.getDataRange().getValues().slice(1).map(r => ({ name: r[0], score: r[1] })) : [];
  return { success: true, students: students, rules: rules };
}
function changePw(id, oldPw, newPw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("계정관리");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id && data[i][1].toString() === oldPw) {
      sheet.getRange(i + 1, 2).setValue(newPw); return { success: true };
    }
  }
  return { success: false, message: "비밀번호가 일치하지 않습니다." };
}
function getSchedule(startDateStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("YP일정");
  if(!sheet) return { success: false, message: "YP일정 시트가 없습니다." };
  const data = sheet.getDataRange().getValues().slice(1);
  let schedule = [];
  data.forEach(r => {
    if (!r[0]) return;
    let rowDate = "";
    try { rowDate = r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Seoul", "yyyy-MM-dd") : Utilities.formatDate(new Date(r[0].toString().trim()), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) { rowDate = r[0].toString().trim(); }
    schedule.push({ date: rowDate, day: r[1] ? r[1].toString().trim() : "", activity: r[2] ? r[2].toString().trim() : "", hakbun: r[3] ? r[3].toString().trim() : "", studentName: r[4] ? r[4].toString().trim() : "", time: `${padTime(r[5])}~${padTime(r[6])}` });
  });
  return { success: true, schedule: schedule };
}
function updateSingleSchedule(payload) {
  const { user, date, day, activity, time, newName } = payload;
  if(user.role !== '관리자' && user.role !== '선도부장') return { success: false, message: "권한이 없습니다." };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  let newHakbun = "";
  if (newName) {
    const match = ss.getSheetByName("계정관리").getDataRange().getValues().slice(1).find(a => a[3] && a[3].toString().replace(/\s+/g, '') === newName);
    if (!match) return { success: false, message: `[${newName}] 학생을 찾을 수 없습니다.` };
    newHakbun = match[2] ? match[2].toString().replace(/\s+/g, '') : "";
  }
  const data = scheduleSheet.getDataRange().getValues();
  let targetRowIndex = -1;
  const [start, end] = time.split('~');
  for (let i = 1; i < data.length; i++) {
    if(!data[i][0]) continue;
    let rDate = "";
    try { rDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "Asia/Seoul", "yyyy-MM-dd") : Utilities.formatDate(new Date(data[i][0].toString().trim()), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) { rDate = data[i][0].toString().trim(); }
    let rStart = padTime(data[i][5]); let rEnd = padTime(data[i][6]);
    if (rDate === date && data[i][2].toString().trim() === activity && `${rStart}~${rEnd}` === time) { targetRowIndex = i + 1; break; }
  }
  if (targetRowIndex !== -1) {
    if (newName) { scheduleSheet.getRange(targetRowIndex, 4).setValue(newHakbun); scheduleSheet.getRange(targetRowIndex, 5).setValue(newName); } 
    else scheduleSheet.deleteRow(targetRowIndex);
  } else if (newName) scheduleSheet.appendRow([date, day, activity, newHakbun, newName, start, end]);
  return { success: true, message: "반영되었습니다." };
}
function uploadSchedule(payload) {
  const { user, scheduleData } = payload;
  if(user.role !== '관리자' && user.role !== '선도부장' && user.role !== '선도차장') return { success: false, message: "권한 없음" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const validAccounts = ss.getSheetByName("계정관리").getDataRange().getValues().slice(1).map(a => ({ hakbun: a[2] ? a[2].toString().replace(/\s+/g, '') : "", name: a[3] ? a[3].toString().replace(/\s+/g, '') : "" }));
  for (let item of scheduleData) {
    if(item.studentName) {
      const match = validAccounts.find(a => a.name === item.studentName.toString().replace(/\s+/g, ''));
      if (!match) return { success: false, message: `검증 실패: [${item.studentName}]` }; else item.hakbun = match.hakbun; 
    }
  }
  const uploadedDates = new Set(); scheduleData.forEach(i => { if(i.date) uploadedDates.add(i.date.toString().trim()); });
  const lastRow = scheduleSheet.getLastRow(); let keptData = [];
  if (lastRow > 1) {
    scheduleSheet.getRange(2, 1, lastRow - 1, 7).getValues().forEach(row => {
      if(!row[0]) return;
      let rowDateStr = "";
      try { rowDateStr = row[0] instanceof Date ? Utilities.formatDate(row[0], "Asia/Seoul", "yyyy-MM-dd") : Utilities.formatDate(new Date(row[0].toString().trim()), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) { rowDateStr = row[0].toString().trim(); }
      if (!uploadedDates.has(rowDateStr)) keptData.push(row);
    });
  }
  if (lastRow > 1) scheduleSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  let finalRows = [...keptData];
  scheduleData.forEach(item => { if(item.studentName) finalRows.push([item.date, item.day, item.activity, item.hakbun, item.studentName.toString().replace(/\s+/g, ''), item.time.split('~')[0], item.time.split('~')[1]]); });
  if (finalRows.length > 0) scheduleSheet.getRange(2, 1, finalRows.length, 7).setValues(finalRows);
  return { success: true, message: "성공" };
}

// --- [게시판 기능] ---
function getBoardSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  return sheet;
}

function getBoardData() {
  const boardSheet = getBoardSheet("게시판", ["게시글ID", "구분", "제목", "내용", "작성자", "작성자학번", "작성시간"]);
  const commentSheet = getBoardSheet("댓글", ["댓글ID", "게시글ID", "내용", "작성자", "작성자학번", "작성시간"]);
  
  const bData = boardSheet.getDataRange().getValues().slice(1);
  const cData = commentSheet.getDataRange().getValues().slice(1);

  let notices = [], normals = [];
  bData.forEach(r => {
    if(!r[0]) return;
    const post = { id: r[0].toString(), type: r[1], title: r[2], content: r[3], author: r[4], hakbun: r[5].toString(), time: r[6] instanceof Date ? Utilities.formatDate(r[6], "Asia/Seoul", "yyyy-MM-dd HH:mm") : r[6] };
    if(post.type === '공지') notices.push(post); else normals.push(post);
  });
  let comments = cData.map(r => ({ id: r[0].toString(), postId: r[1].toString(), content: r[2], author: r[3], hakbun: r[4].toString(), time: r[5] instanceof Date ? Utilities.formatDate(r[5], "Asia/Seoul", "yyyy-MM-dd HH:mm") : r[5] }));

  notices.reverse(); normals.reverse();
  return { success: true, notices, normals, comments };
}

function savePost(payload) {
  const { user, isNotice, title, content } = payload;
  if(isNotice && user.role === 'YP') return { success: false, message: "공지 작성 권한이 없습니다." };
  const boardSheet = getBoardSheet("게시판", ["게시글ID", "구분", "제목", "내용", "작성자", "작성자학번", "작성시간"]);
  boardSheet.appendRow([new Date().getTime().toString(), isNotice ? "공지" : "일반", title, content, user.name, user.hakbun, new Date()]);
  return { success: true, message: "등록되었습니다." };
}

// 🔥 추가: 게시글 수정
function updatePost(payload) {
  const { user, postId, isNotice, title, content } = payload;
  if(isNotice && user.role === 'YP') return { success: false, message: "공지 권한이 없습니다." };
  
  const boardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("게시판");
  const data = boardSheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === postId) {
      if(data[i][5].toString() !== user.hakbun && user.role !== '관리자' && user.role !== '선도부장') {
        return { success: false, message: "수정 권한이 없습니다." };
      }
      boardSheet.getRange(i+1, 2).setValue(isNotice ? "공지" : "일반");
      boardSheet.getRange(i+1, 3).setValue(title);
      boardSheet.getRange(i+1, 4).setValue(content);
      return { success: true, message: "수정되었습니다." };
    }
  }
  return { success: false, message: "게시글을 찾을 수 없습니다." };
}

// 🔥 추가: 게시글 삭제 (게시글 + 연관 댓글 동시 삭제)
function deletePost(payload) {
  const { user, postId } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const boardSheet = ss.getSheetByName("게시판");
  if(!boardSheet) return { success: false, message: "게시판 시트가 없습니다." };

  const data = boardSheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === postId) {
      if(data[i][5].toString() !== user.hakbun && user.role !== '관리자' && user.role !== '선도부장') {
        return { success: false, message: "삭제 권한이 없습니다." };
      }
      boardSheet.deleteRow(i + 1);
      
      // 연관된 댓글도 지우기
      const commentSheet = ss.getSheetByName("댓글");
      if(commentSheet) {
        const cData = commentSheet.getDataRange().getValues();
        // 뒤에서부터 지워야 인덱스가 안 꼬입니다.
        for(let j = cData.length - 1; j >= 1; j--) {
          if(cData[j][1].toString() === postId) commentSheet.deleteRow(j + 1);
        }
      }
      return { success: true, message: "삭제되었습니다." };
    }
  }
  return { success: false, message: "게시글을 찾을 수 없습니다." };
}

function saveComment(payload) {
  const { user, postId, content } = payload;
  const commentSheet = getBoardSheet("댓글", ["댓글ID", "게시글ID", "내용", "작성자", "작성자학번", "작성시간"]);
  commentSheet.appendRow([new Date().getTime().toString(), postId, content, user.name, user.hakbun, new Date()]);
  return { success: true, message: "등록되었습니다." };
}

// 🔥 이번 주 일정 일괄 저장 (화면의 데이터를 모아 해당 주차만 덮어쓰기)
function saveWeeklySchedule(payload) {
  const { user, targetDates, scheduleData } = payload;
  if(user.role !== '관리자' && user.role !== '선도부장') return { success: false, message: "권한이 없습니다." };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const validAccounts = ss.getSheetByName("계정관리").getDataRange().getValues().slice(1).map(a => ({ hakbun: a[2] ? a[2].toString().replace(/\s+/g, '') : "", name: a[3] ? a[3].toString().replace(/\s+/g, '') : "" }));
  
  // 학번 매칭
  for (let item of scheduleData) {
    if(item.studentName) {
      const reqName = item.studentName.toString().replace(/\s+/g, '');
      const match = validAccounts.find(a => a.name === reqName);
      if (!match) return { success: false, message: `[${reqName}] 학생을 찾을 수 없습니다.` }; 
      else item.hakbun = match.hakbun; 
    }
  }

  // 화면에 띄워진 5일(targetDates)에 해당하는 기존 데이터는 싹 지우고 나머지는 유지
  const datesToReplace = new Set(targetDates);
  const lastRow = scheduleSheet.getLastRow(); 
  let keptData = [];
  
  if (lastRow > 1) {
    scheduleSheet.getRange(2, 1, lastRow - 1, 7).getValues().forEach(row => {
      if(!row[0]) return;
      let rowDateStr = "";
      try { rowDateStr = row[0] instanceof Date ? Utilities.formatDate(row[0], "Asia/Seoul", "yyyy-MM-dd") : Utilities.formatDate(new Date(row[0].toString().trim()), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) { rowDateStr = row[0].toString().trim(); }
      if (!datesToReplace.has(rowDateStr)) keptData.push(row);
    });
  }
  
  if (lastRow > 1) scheduleSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  let finalRows = [...keptData];
  
  scheduleData.forEach(item => { 
      if(item.studentName) {
          const [start, end] = item.time.split('~');
          finalRows.push([item.date, item.day, item.activity, item.hakbun, item.studentName.toString().replace(/\s+/g, ''), start, end]); 
      }
  });
  
  if (finalRows.length > 0) scheduleSheet.getRange(2, 1, finalRows.length, 7).setValues(finalRows);
  return { success: true, message: "일정이 저장되었습니다." };
}

// --- [자동 이메일 리포트 발송 기능] ---

// --- [자동 이메일 리포트 발송 기능] ---

// --- [자동 이메일 리포트 발송 기능] ---

function checkAndSendDailyReport() {
  const now = new Date();
  const dayOfWeek = now.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return; // 주말 제외

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("환경설정");
  if (!configSheet) return;

  const config = configSheet.getDataRange().getValues()[1];
  if (!config || !config[0]) return;

  const email = config[0].toString().trim();
  // 환경설정 시트의 B열(2번째 열)에서 최근발송일 확인
  const lastSent = config[1] ? Utilities.formatDate(new Date(config[1]), "Asia/Seoul", "yyyy-MM-dd") : "";
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  // 오늘 이미 발송했다면 중복 발송 방지
  if (lastSent === todayStr) return;

  sendHtmlEmail(email, todayStr);
  configSheet.getRange(2, 2).setValue(todayStr); // B2 셀에 오늘 날짜 기록
}

function sendHtmlEmail(email, dateStr) {
  const now = new Date();

  // 데이터 가져오기
  const adminData = getAdminData(dateStr);
  const missingStats = getUnattendedStats({ startDate: '2026-03-25', endDate: dateStr });
  
  const allSchedulesRes = getSchedule(); 
  const allSchedules = allSchedulesRes.success ? allSchedulesRes.schedule : [];

  const thisMonday = getMonday(now);
  const nextMonday = new Date(thisMonday);
  nextMonday.setDate(nextMonday.getDate() + 7);

  let htmlBody = `<div style="font-family: sans-serif; color: #333;">`;
  htmlBody += `<h2 style="color:#0d6efd; font-size:20px;">YP 학생회 일일 종합 리포트 (${dateStr})</h2>`;
  
  // [1. 일일 출석 현황] - font-size 12px 적용
  htmlBody += `<h3 style="font-size:16px;">📊 일일 출석 현황</h3><table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align:center; font-size:12px; width:100%; max-width:700px;">
    <tr style="background-color:#f8f9fa;"><th>활동(시간)</th><th>담당자</th><th>상태</th><th>사유</th></tr>`;
  if(adminData.attendance.length === 0) htmlBody += `<tr><td colspan="4">오늘 일정이 없습니다.</td></tr>`;
  else adminData.attendance.forEach(a => {
    let color = a.status.includes('정상') ? 'green' : (a.status.includes('미출석') ? 'red' : 'orange');
    htmlBody += `<tr><td>${a.activity}<br><small style="color:#6c757d;">${a.time}</small></td><td>${a.name}</td><td style="color:${color}; font-weight:bold;">${a.status}</td><td>${a.excuse}</td></tr>`;
  });
  htmlBody += `</table><br>`;

  // [2. 일일 상벌점 내역] - font-size 12px 적용
  htmlBody += `<h3 style="font-size:16px;">📝 일일 상벌점 부여 내역</h3><table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align:center; font-size:12px; width:100%; max-width:700px;">
    <tr style="background-color:#f8f9fa;"><th>대상</th><th>항목(점수)</th><th>기록자(YP)</th></tr>`;
  if(adminData.records.length === 0) htmlBody += `<tr><td colspan="3">오늘 부여된 상벌점이 없습니다.</td></tr>`;
  else adminData.records.forEach(r => {
    htmlBody += `<tr><td>${r.target}</td><td>${r.rule}</td><td>${r.author}</td></tr>`;
  });
  htmlBody += `</table><br>`;

  // [3. 누적 미출석 현황] - font-size 12px 적용
  htmlBody += `<h3 style="font-size:16px;">🚨 누적 미출석 현황 (26.03.25 ~ 오늘)</h3><table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align:center; font-size:12px; width:100%; max-width:700px;">
    <tr style="background-color:#ffeeba;"><th>날짜(요일)</th><th>활동(시간)</th><th>담당자</th><th>결석 누적</th></tr>`;
  if(!missingStats.success || missingStats.missingList.length === 0) {
    htmlBody += `<tr><td colspan="4">해당 기간 결석 인원이 없습니다. 🎉</td></tr>`;
  } else {
    missingStats.missingList.forEach(m => {
      htmlBody += `<tr><td>${m.date}<br><small>(${m.day})</small></td><td>${m.activity}<br><small style="color:#6c757d;">${m.time}</small></td><td><b>${m.name}</b><br><small style="color:#6c757d;">${m.hakbun}</small></td><td style="color:red; font-weight:bold;">${missingStats.studentCounts[m.hakbun].count}회</td></tr>`;
    });
  }
  htmlBody += `</table><br>`;

  // [4. 이번 주 & 다음 주 달력 렌더링]
  htmlBody += buildWeeklyCalendarHtml("이번 주 전체 YP 일정", thisMonday, allSchedules);
  htmlBody += buildWeeklyCalendarHtml("다음 주 전체 YP 일정", nextMonday, allSchedules);

  htmlBody += `</div>`;

  // 메일 전송
  MailApp.sendEmail({
    to: email,
    subject: `[YP시스템] ${dateStr} 학생회 활동 일일 종합 리포트`,
    htmlBody: htmlBody
  });
}

// 🔥 메일 전용 주간 달력 생성 헬퍼 함수
function buildWeeklyCalendarHtml(title, mondayDate, allSchedules) {
  const DAYS_KR = ["일", "월", "화", "수", "목", "금", "토"];
  const SCHEDULE_TEMPLATE = [
    { activity: "아침선도", time: "07:45~08:15" }, { activity: "아침선도", time: "07:45~08:15" },
    { activity: "점심엘베", time: "12:25~12:40" }, { activity: "점심엘베", time: "12:40~13:00" }, { activity: "점심엘베", time: "13:00~13:15" },
    { activity: "점심순찰", time: "12:50~13:15" }, { activity: "점심순찰", time: "12:50~13:15" }
  ];

  let weekDates = [];
  let headerHtml = `<tr style="background-color:#e2e3e5;"><th style="width:16%;">활동</th>`;
  for(let i=0; i<5; i++) {
    let tempDate = new Date(mondayDate);
    tempDate.setDate(mondayDate.getDate() + i);
    let dStr = Utilities.formatDate(tempDate, "Asia/Seoul", "yyyy-MM-dd");
    weekDates.push(dStr);
    headerHtml += `<th style="width:16.8%;">${DAYS_KR[i+1]}<br><small>(${tempDate.getMonth()+1}/${tempDate.getDate()})</small></th>`;
  }
  headerHtml += `</tr>`;

  let availableData = JSON.parse(JSON.stringify(allSchedules.filter(s => weekDates.includes(s.date))));

  let bodyHtml = '';
  SCHEDULE_TEMPLATE.forEach(row => {
    // 폰트 크기 비율 소폭 조정
    bodyHtml += `<tr><td style="line-height: 1.2; text-align:center;"><b>${row.activity}</b><br><span style="color:#6c757d;">${row.time}</span></td>`;
    for(let i=0; i<5; i++) {
      let targetDate = weekDates[i];
      let index = availableData.findIndex(d => d.date === targetDate && d.activity === row.activity && d.time === row.time);
      let display = "";
      if (index !== -1) {
        let record = availableData[index];
        display = `<b>${record.studentName}</b><br><small style="color:#6c757d;">${record.hakbun}</small>`;
        availableData.splice(index, 1);
      }
      bodyHtml += `<td style="text-align:center;">${display}</td>`;
    }
    bodyHtml += `</tr>`;
  });

  // font-size 12px 적용
  return `<h3 style="font-size:16px;">📅 ${title}</h3><table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-size:12px; width:100%; max-width:800px; table-layout:fixed;">
    <thead>${headerHtml}</thead><tbody>${bodyHtml}</tbody></table><br>`;
}

// 🔥 기간별 미출석 현황 조회 (Left Anti Join 및 과거 시간 검증 로직)
function getUnattendedStats(payload) {
  const { startDate, endDate } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("YP일정");
  const attendSheet = ss.getSheetByName("출석기록");

  if(!scheduleSheet) return { success: false, message: "YP일정 시트가 없습니다." };

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const nowMin = now.getHours() * 60 + now.getMinutes();

  const start = new Date(startDate);
  const end = new Date(endDate);

  const allSchedules = scheduleSheet.getDataRange().getValues().slice(1);
  const allAttendances = attendSheet ? attendSheet.getDataRange().getValues().slice(1) : [];

  // 🔥 미출석비고 가져오기
  const noteSheet = ss.getSheetByName("미출석비고");
  const allNotes = noteSheet ? noteSheet.getDataRange().getValues().slice(1) : [];

  let missingList = [];
  let studentCounts = {};

  allSchedules.forEach(s => {
    if(!s[0]) return;
    let dStr = s[0] instanceof Date ? Utilities.formatDate(s[0], "Asia/Seoul", "yyyy-MM-dd") : s[0].toString().trim();
    let sDate = new Date(dStr);

    // 1. 기간 필터링
    if(sDate >= start && sDate <= end) {
      
      // 2. 시간 검증 (과거의 일정인지 확인)
      let isPast = false;
      if(dStr < todayStr) {
        isPast = true; // 어제 이전 일정은 무조건 과거
      } else if (dStr === todayStr) {
        let endMin = getMinutes(s[6]);
        // 종료 시간 + 5분 유예기간이 지난 일정만 '과거'로 취급하여 검사
        if (nowMin > endMin + 5) isPast = true;
      }

      if(isPast) {
        const actName = s[2].toString().trim();
        const hakbun = s[3].toString().trim();
        const studentName = s[4].toString().trim();

        // 3. 교집합 검증 (출석 기록이 있는지 확인)
        const attended = allAttendances.some(a => {
          let aDateStr = a[0] instanceof Date ? Utilities.formatDate(a[0], "Asia/Seoul", "yyyy-MM-dd") : a[0].toString().trim();
          return aDateStr === dStr && a[2].toString().trim() === studentName && a[1].toString().trim() === actName;
        });

        // 출석 기록이 없다면 미출석 배열에 추가
        if(!attended) {
           let matchingNote = "";
           const matchedNoteRow = allNotes.find(n => {
              let nDateStr = n[0] instanceof Date ? Utilities.formatDate(n[0], "Asia/Seoul", "yyyy-MM-dd") : n[0].toString().trim();
              return nDateStr === dStr && n[1].toString().trim() === actName && n[2].toString().trim() === hakbun;
           });
           if(matchedNoteRow) matchingNote = matchedNoteRow[3].toString();

           missingList.push({
             date: dStr, day: s[1], time: `${padTime(s[5])}~${padTime(s[6])}`,
             activity: actName, hakbun: hakbun, name: studentName,
             note: matchingNote // 🔥 비고 속성 추가
           });

           if(!studentCounts[hakbun]) studentCounts[hakbun] = { count: 0 };
           studentCounts[hakbun].count++;
        }
      }
    }
  });

  // 4. 결석 누적 횟수가 높은 순(내림차순)으로 자동 정렬, 동률이면 최신순 정렬
  missingList.sort((a, b) => studentCounts[b.hakbun].count - studentCounts[a.hakbun].count || new Date(b.date) - new Date(a.date));

  return { success: true, missingList, studentCounts };
}
