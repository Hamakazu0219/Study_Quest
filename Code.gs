/**
 * Study Quest - GAS Webアプリ（サーバー側）
 * 要件定義 request.md に準拠
 */

const SHEET_NAMES = {
  PROBLEMS: '問題',
  STUDENTS: '生徒データ',
  SETTINGS: '設定'
};

const COLS = {
  PROBLEM: { ID: 0, TEXT: 1, C1: 2, C2: 3, C3: 4, C4: 5, CORRECT: 6, TYPE: 7, LEVEL: 8 },
  STUDENT: { NAME: 0, PASSWORD: 1, EMAIL: 2, NICKNAME: 3, LEVEL: 4, CLEARED: 5, LEVEL_AT: 6, UPDATED: 7 }
};

const TEACHER_PASSWORD_KEY = 'TEACHER_PASSWORD';
const DEFAULT_TEACHER_PASSWORD = '1234';
const ERROR_MSG = 'エラーが発生しました。しばらく経ってからお試しください。';
const MAX_LEVEL = 20;

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Study Quest')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    console.error('doGet error:', e);
    throw e;
  }
}

/**
 * 指定のスプレッドシートをアプリで使うように切り替えます。
 * GASエディタで「実行」→ この関数を選んで1回だけ実行してください。
 * 引数にスプレッドシートID（URLの /d/ と /edit の間）を入れます。
 */
function setSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId) {
    spreadsheetId = '1Sem5TDTd9IjdiBs0V-lh0hzVXdyhKzzq1VuhbxLAv7w'; // ここを変更しても可
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  Logger.log('SPREADSHEET_ID を ' + spreadsheetId + ' に設定しました。');
}

function getOrCreateSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) {
    var ss = SpreadsheetApp.create('Study_Quest_データ');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    ensureSheets(ss);
  }
  return SpreadsheetApp.openById(ssId);
}

function ensureSheets(ss) {
  var problemsSheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
  if (!problemsSheet) {
    problemsSheet = ss.insertSheet(SHEET_NAMES.PROBLEMS);
    problemsSheet.getRange(1, 1, 1, 9).setValues([['問題ID', '問題文', '選択肢1', '選択肢2', '選択肢3', '選択肢4', '正解', '問題タイプ', 'レベル']]);
  }
  var studentsSheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet(SHEET_NAMES.STUDENTS);
    studentsSheet.getRange(1, 1, 1, 8).setValues([['名前', 'パスワード', 'メールアドレス', 'ニックネーム', '現在のレベル', 'クリアした課題数', 'レベル到達日時', '最終更新日時']]);
  }
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['項目', '値']]);
  }
}

function normalizeForCompare(str) {
  if (str == null || str === '') return '';
  var s = String(str).replace(/\s+/g, ' ').trim();
  return s.replace(/[０-９Ａ-Ｚａ-ｚ]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  }).toLowerCase();
}

function getStudentRow(ss, name) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS.STUDENT.NAME]).trim() === String(name).trim()) return i + 1;
  }
  return -1;
}

function getStudentByRow(ss, row) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  var data = sheet.getRange(row, 1, row, 8).getValues()[0];
  return {
    name: data[COLS.STUDENT.NAME],
    nickname: data[COLS.STUDENT.NICKNAME],
    level: parseInt(data[COLS.STUDENT.LEVEL], 10) || 1,
    cleared: parseInt(data[COLS.STUDENT.CLEARED], 10) || 0,
    levelAt: data[COLS.STUDENT.LEVEL_AT]
  };
}

// ----- 生徒用API -----

function getStudentList() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    if (!sheet) return { success: true, names: [] };
    var data = sheet.getDataRange().getValues();
    var names = [];
    for (var i = 1; i < data.length; i++) {
      var n = String(data[i][COLS.STUDENT.NAME]).trim();
      if (n) names.push(n);
    }
    names.sort();
    return { success: true, names: names };
  } catch (e) {
    console.error('getStudentList:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function verifyStudent(name, password) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '名前またはパスワードが違います' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var data = sheet.getRange(row, 1, row, 8).getValues()[0];
    var pwd = String(data[COLS.STUDENT.PASSWORD]).trim();
    if (pwd !== String(password).trim()) return { success: false, error: '名前またはパスワードが違います' };
    var nickname = String(data[COLS.STUDENT.NICKNAME]).trim();
    return { success: true, needsNickname: !nickname };
  } catch (e) {
    console.error('verifyStudent:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function setNickname(name, password, nickname) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '名前またはパスワードが違います' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var data = sheet.getRange(row, 1, row, 2).getValues()[0];
    if (String(data[COLS.STUDENT.PASSWORD]).trim() !== String(password).trim()) return { success: false, error: '名前またはパスワードが違います' };
    var nick = String(nickname).trim().slice(0, 20);
    sheet.getRange(row, COLS.STUDENT.NICKNAME + 1).setValue(nick);
    sheet.getRange(row, COLS.STUDENT.UPDATED + 1).setValue(new Date());
    return { success: true };
  } catch (e) {
    console.error('setNickname:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function getProblemCountByLevel(ss, level) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var lev = data[i][COLS.PROBLEM.LEVEL];
    if (lev !== '' && lev !== null && parseInt(lev, 10) === level) count++;
  }
  return count;
}

function getRankAndTop5(ss, currentName) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!sheet) return { rank: 1, totalCount: 0, top5: [] };
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][COLS.STUDENT.NAME]).trim();
    if (!name) continue;
    var level = parseInt(data[i][COLS.STUDENT.LEVEL], 10) || 1;
    var levelAt = data[i][COLS.STUDENT.LEVEL_AT];
    var t = levelAt instanceof Date ? levelAt.getTime() : (levelAt ? new Date(levelAt).getTime() : 0);
    list.push({
      name: name,
      nickname: String(data[i][COLS.STUDENT.NICKNAME]).trim() || '（ニックネームなし）',
      level: level,
      levelAt: t
    });
  }
  list.sort(function(a, b) {
    if (a.level !== b.level) return b.level - a.level;
    return a.levelAt - b.levelAt;
  });
  var totalCount = list.length;
  var rank = 1;
  var prevLevel = null;
  var prevAt = null;
  var currentRank = 1;
  for (var j = 0; j < list.length; j++) {
    if (j > 0 && (list[j].level !== prevLevel || list[j].levelAt !== prevAt)) rank = j + 1;
    prevLevel = list[j].level;
    prevAt = list[j].levelAt;
    if (list[j].name === currentName) currentRank = rank;
  }
  var top5 = list.slice(0, 5).map(function(item, idx) {
    var r = 1;
    var plevel = null, pat = null;
    for (var k = 0; k <= idx; k++) {
      if (plevel !== null && (list[k].level !== plevel || list[k].levelAt !== pat)) r = k + 1;
      plevel = list[k].level;
      pat = list[k].levelAt;
    }
    return { rank: r, nickname: list[idx].nickname };
  });
  return { rank: currentRank, totalCount: totalCount, top5: top5 };
}

function getMainData(name, password) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '名前またはパスワードが違います' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var data = sheet.getRange(row, 1, row, 8).getValues()[0];
    if (String(data[COLS.STUDENT.PASSWORD]).trim() !== String(password).trim()) return { success: false, error: '名前またはパスワードが違います' };
    var level = parseInt(data[COLS.STUDENT.LEVEL], 10) || 1;
    var rankData = getRankAndTop5(ss, name);
    var questionCount = level < MAX_LEVEL ? getProblemCountByLevel(ss, level) : 0;
    var canChallenge = level < MAX_LEVEL && questionCount > 0;
    var message = '';
    if (level < MAX_LEVEL && questionCount === 0) message = '問題がまだ登録されていません。先生に連絡してください。';
    return {
      success: true,
      level: level,
      rank: rankData.rank,
      totalCount: rankData.totalCount,
      top5: rankData.top5,
      canChallenge: canChallenge,
      message: message,
      gameClear: level >= MAX_LEVEL
    };
  } catch (e) {
    console.error('getMainData:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function shuffleArray(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i];
    a[i] = a[j];
    a[j] = t;
  }
  return a;
}

function getQuestionsForChallenge(name, password) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '名前またはパスワードが違います' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var data = sheet.getRange(row, 1, row, 2).getValues()[0];
    if (String(data[COLS.STUDENT.PASSWORD]).trim() !== String(password).trim()) return { success: false, error: '名前またはパスワードが違います' };
    var level = parseInt(data[COLS.STUDENT.LEVEL], 10) || 1;
    if (level >= MAX_LEVEL) return { success: false, error: 'このレベルはクリア済みです' };
    var probSheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    if (!probSheet) return { success: true, questions: [] };
    var allData = probSheet.getDataRange().getValues();
    var pool = [];
    for (var i = 1; i < allData.length; i++) {
      var lev = allData[i][COLS.PROBLEM.LEVEL];
      if (lev === '' || lev == null) continue;
      if (parseInt(lev, 10) === level) pool.push({ rowIndex: i + 1, row: allData[i] });
    }
    if (pool.length === 0) return { success: false, error: '問題がまだ登録されていません。先生に連絡してください。' };
    var selected = pool.length <= 5 ? pool : shuffleArray(pool).slice(0, 5);
    var questions = [];
    for (var s = 0; s < selected.length; s++) {
      var r = selected[s].row;
      var type = String(r[COLS.PROBLEM.TYPE]).trim();
      var correct = String(r[COLS.PROBLEM.CORRECT]).trim();
      var id = String(r[COLS.PROBLEM.ID]);
      if (type === '選択式') {
        var choices = [r[COLS.PROBLEM.C1], r[COLS.PROBLEM.C2], r[COLS.PROBLEM.C3], r[COLS.PROBLEM.C4]].filter(function(c) { return c != null && String(c).trim() !== ''; });
        choices = shuffleArray(choices);
        questions.push({
          id: id,
          rowIndex: selected[s].rowIndex,
          text: r[COLS.PROBLEM.TEXT],
          type: '選択式',
          choices: choices,
          correct: correct
        });
      } else {
        questions.push({
          id: id,
          rowIndex: selected[s].rowIndex,
          text: r[COLS.PROBLEM.TEXT],
          type: '記述式',
          correct: correct
        });
      }
    }
    return { success: true, questions: questions };
  } catch (e) {
    console.error('getQuestionsForChallenge:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function submitChallenge(name, password, answers) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '名前またはパスワードが違います' };
    var stSheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var stData = stSheet.getRange(row, 1, row, 8).getValues()[0];
    if (String(stData[COLS.STUDENT.PASSWORD]).trim() !== String(password).trim()) return { success: false, error: '名前またはパスワードが違います' };
    var probSheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    var results = [];
    var allCorrect = true;
    for (var i = 0; i < answers.length; i++) {
      var ans = answers[i];
      var probRow = null;
      var probData = probSheet.getDataRange().getValues();
      for (var r = 1; r < probData.length; r++) {
        if (String(probData[r][COLS.PROBLEM.ID]) === String(ans.id)) {
          probRow = probData[r];
          break;
        }
      }
      if (!probRow) {
        results.push({ correct: false, correctAnswer: '' });
        allCorrect = false;
        continue;
      }
      var correctVal = String(probRow[COLS.PROBLEM.CORRECT]).trim();
      var userVal = ans.answer != null ? String(ans.answer).trim() : '';
      var type = String(probRow[COLS.PROBLEM.TYPE]).trim();
      var isCorrect = false;
      if (type === '選択式') {
        isCorrect = normalizeForCompare(userVal) === normalizeForCompare(correctVal);
      } else {
        isCorrect = normalizeForCompare(userVal) === normalizeForCompare(correctVal);
      }
      results.push({ correct: isCorrect, correctAnswer: correctVal });
      if (!isCorrect) allCorrect = false;
    }
    var level = parseInt(stData[COLS.STUDENT.LEVEL], 10) || 1;
    if (allCorrect && level < MAX_LEVEL) {
      var now = new Date();
      stSheet.getRange(row, COLS.STUDENT.LEVEL + 1).setValue(level + 1);
      stSheet.getRange(row, COLS.STUDENT.CLEARED + 1).setValue((parseInt(stData[COLS.STUDENT.CLEARED], 10) || 0) + 1);
      stSheet.getRange(row, COLS.STUDENT.LEVEL_AT + 1).setValue(now);
      stSheet.getRange(row, COLS.STUDENT.UPDATED + 1).setValue(now);
    }
    return {
      success: true,
      passed: allCorrect,
      results: results,
      levelUpdated: allCorrect
    };
  } catch (e) {
    console.error('submitChallenge:', e);
    return { success: false, error: ERROR_MSG };
  }
}

// ----- 教師用API -----

function getTeacherPassword() {
  var p = PropertiesService.getScriptProperties().getProperty(TEACHER_PASSWORD_KEY);
  return p || DEFAULT_TEACHER_PASSWORD;
}

function verifyTeacher(password) {
  try {
    var correct = getTeacherPassword();
    return { success: correct === String(password) };
  } catch (e) {
    return { success: false };
  }
}

function changeTeacherPassword(currentPassword, newPassword) {
  try {
    var current = getTeacherPassword();
    if (current !== String(currentPassword)) return { success: false, error: '現在のパスワードが違います' };
    PropertiesService.getScriptProperties().setProperty(TEACHER_PASSWORD_KEY, String(newPassword));
    return { success: true };
  } catch (e) {
    console.error('changeTeacherPassword:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function getStudents() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    if (!sheet) return { success: true, students: [] };
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][COLS.STUDENT.NAME]).trim();
      if (!name) continue;
      list.push({
        name: name,
        nickname: String(data[i][COLS.STUDENT.NICKNAME]).trim(),
        level: parseInt(data[i][COLS.STUDENT.LEVEL], 10) || 1,
        cleared: parseInt(data[i][COLS.STUDENT.CLEARED], 10) || 0
      });
    }
    list.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return { success: true, students: list };
  } catch (e) {
    console.error('getStudents:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function addStudent(name, password) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var nameTrim = String(name).trim();
    if (!nameTrim) return { success: false, error: '名前を入力してください' };
    if (getStudentRow(ss, nameTrim) >= 0) return { success: false, error: 'この名前は既に登録されています' };
    var pwd = String(password).trim();
    if (!/^\d{4}$/.test(pwd)) return { success: false, error: 'パスワードは4桁の数字で入力してください' };
    var now = new Date();
    sheet.appendRow([nameTrim, pwd, '', '', 1, 0, now, now]);
    return { success: true };
  } catch (e) {
    console.error('addStudent:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function deleteStudent(name) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '生徒が見つかりません' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    console.error('deleteStudent:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function resetStudentPassword(name, newPassword) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '生徒が見つかりません' };
    var pwd = String(newPassword).trim();
    if (!/^\d{4}$/.test(pwd)) return { success: false, error: 'パスワードは4桁の数字で入力してください' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    sheet.getRange(row, COLS.STUDENT.PASSWORD + 1).setValue(pwd);
    sheet.getRange(row, COLS.STUDENT.UPDATED + 1).setValue(new Date());
    return { success: true };
  } catch (e) {
    console.error('resetStudentPassword:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function updateStudentNickname(name, nickname) {
  try {
    var ss = getOrCreateSpreadsheet();
    var row = getStudentRow(ss, name);
    if (row < 0) return { success: false, error: '生徒が見つかりません' };
    var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
    var nick = String(nickname).trim().slice(0, 20);
    sheet.getRange(row, COLS.STUDENT.NICKNAME + 1).setValue(nick);
    sheet.getRange(row, COLS.STUDENT.UPDATED + 1).setValue(new Date());
    return { success: true };
  } catch (e) {
    console.error('updateStudentNickname:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function problemIdExists(ss, problemId) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  var idStr = String(problemId).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS.PROBLEM.ID]).trim() === idStr) return true;
  }
  return false;
}

function getProblems() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    if (!sheet) return { success: true, problems: [] };
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({
        rowIndex: i + 1,
        id: data[i][COLS.PROBLEM.ID],
        text: data[i][COLS.PROBLEM.TEXT],
        c1: data[i][COLS.PROBLEM.C1],
        c2: data[i][COLS.PROBLEM.C2],
        c3: data[i][COLS.PROBLEM.C3],
        c4: data[i][COLS.PROBLEM.C4],
        correct: data[i][COLS.PROBLEM.CORRECT],
        type: data[i][COLS.PROBLEM.TYPE],
        level: data[i][COLS.PROBLEM.LEVEL]
      });
    }
    list.sort(function(a, b) { return String(a.id).localeCompare(String(b.id)); });
    return { success: true, problems: list };
  } catch (e) {
    console.error('getProblems:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function addProblem(problemId, questionText, choice1, choice2, choice3, choice4, correct, problemType, level) {
  try {
    var ss = getOrCreateSpreadsheet();
    var idStr = String(problemId).trim();
    if (!idStr) return { success: false, error: '問題IDを入力してください' };
    if (problemIdExists(ss, idStr)) return { success: false, error: 'この問題IDは既に使われています' };
    var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    sheet.appendRow([idStr, questionText || '', choice1 || '', choice2 || '', choice3 || '', choice4 || '', correct || '', problemType || '選択式', level != null && level !== '' ? level : '']);
    return { success: true };
  } catch (e) {
    console.error('addProblem:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function updateProblem(rowIndex, problemId, questionText, choice1, choice2, choice3, choice4, correct, problemType, level) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    var row = parseInt(rowIndex, 10);
    if (isNaN(row) || row < 2) return { success: false, error: '問題が見つかりません' };
    sheet.getRange(row, 1, row, 9).setValues([[problemId, questionText || '', choice1 || '', choice2 || '', choice3 || '', choice4 || '', correct || '', problemType || '選択式', level != null && level !== '' ? level : '']]);
    return { success: true };
  } catch (e) {
    console.error('updateProblem:', e);
    return { success: false, error: ERROR_MSG };
  }
}

function deleteProblem(rowIndex) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.PROBLEMS);
    var row = parseInt(rowIndex, 10);
    if (isNaN(row) || row < 2) return { success: false, error: '問題が見つかりません' };
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    console.error('deleteProblem:', e);
    return { success: false, error: ERROR_MSG };
  }
}
