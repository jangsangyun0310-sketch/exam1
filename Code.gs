// 게시판 백엔드 (구글 시트를 DB로 사용)
// 컬럼: ID, Title, Content, Author, Password, Views, CreatedAt, UpdatedAt
// 주의: Password는 시트에 평문으로 저장됩니다. 민감한 정보 보호용이 아니라
// "아무나 함부로 남의 글을 수정/삭제하지 못하게" 막는 최소한의 장치입니다.

const SHEET_NAME = 'Posts';
const HEADERS = ['ID', 'Title', 'Content', 'Author', 'Password', 'Views', 'CreatedAt', 'UpdatedAt'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function readAll() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(function (row, idx) {
    const obj = { _row: idx + 2 }; // 실제 시트 행 번호 (헤더가 1행이므로 +2)
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function stripPassword(post) {
  const copy = {};
  for (const k in post) {
    if (k !== 'Password' && k !== '_row') copy[k] = post[k];
  }
  return copy;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 조회 - GET 요청
// 파라미터 없음: 목록 조회 / ?id=xxx: 상세 조회(조회수 +1)
function doGet(e) {
  const params = e.parameter;
  const posts = readAll();

  if (params.id) {
    const post = posts.find(function (p) { return p.ID === params.id; });
    if (!post) {
      return jsonOutput({ success: false, error: '게시글을 찾을 수 없습니다.' });
    }
    const sheet = getSheet();
    const newViews = (Number(post.Views) || 0) + 1;
    sheet.getRange(post._row, HEADERS.indexOf('Views') + 1).setValue(newViews);
    post.Views = newViews;
    return jsonOutput({ success: true, post: stripPassword(post) });
  }

  const list = posts
    .map(stripPassword)
    .sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
  return jsonOutput({ success: true, posts: list });
}

// 생성/수정/삭제 - POST 요청
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const sheet = getSheet();

  if (action === 'create') {
    if (!body.title || !body.author || !body.password) {
      return jsonOutput({ success: false, error: '제목/작성자/비밀번호는 필수입니다.' });
    }
    const id = Utilities.getUuid();
    const now = new Date();
    sheet.appendRow([id, body.title, body.content || '', body.author, body.password, 0, now, now]);
    return jsonOutput({ success: true, id: id });
  }

  if (action === 'update' || action === 'delete') {
    const posts = readAll();
    const post = posts.find(function (p) { return p.ID === body.id; });
    if (!post) {
      return jsonOutput({ success: false, error: '게시글을 찾을 수 없습니다.' });
    }
    if (String(post.Password) !== String(body.password)) {
      return jsonOutput({ success: false, error: '비밀번호가 일치하지 않습니다.' });
    }

    if (action === 'delete') {
      sheet.deleteRow(post._row);
      return jsonOutput({ success: true });
    }

    // update
    if (!body.title) {
      return jsonOutput({ success: false, error: '제목은 필수입니다.' });
    }
    sheet.getRange(post._row, HEADERS.indexOf('Title') + 1).setValue(body.title);
    sheet.getRange(post._row, HEADERS.indexOf('Content') + 1).setValue(body.content || '');
    sheet.getRange(post._row, HEADERS.indexOf('UpdatedAt') + 1).setValue(new Date());
    return jsonOutput({ success: true });
  }

  return jsonOutput({ success: false, error: '알 수 없는 요청입니다.' });
}
