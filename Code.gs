// 게시판 백엔드 (구글 시트를 DB로, 구글 드라이브를 이미지 저장소로 사용)
// Posts 시트: ID, Title, Content, Author, Password, Category, ImageUrl, ImageFileId, Views, CreatedAt, UpdatedAt
// Comments 시트: ID, PostID, Author, Content, Password, CreatedAt
// 주의: Password는 시트에 평문으로 저장됩니다. 민감한 정보 보호용이 아니라
// "아무나 함부로 남의 글/댓글을 수정/삭제하지 못하게" 막는 최소한의 장치입니다.

const POSTS_SHEET = 'Posts';
const POSTS_HEADERS = ['ID', 'Title', 'Content', 'Author', 'Password', 'Category', 'ImageUrl', 'ImageFileId', 'Views', 'CreatedAt', 'UpdatedAt'];
const COMMENTS_SHEET = 'Comments';
const COMMENTS_HEADERS = ['ID', 'PostID', 'Author', 'Content', 'Password', 'CreatedAt'];
const IMAGE_FOLDER_NAME = 'board-images';

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}
function getPostsSheet() { return getSheet(POSTS_SHEET, POSTS_HEADERS); }
function getCommentsSheet() { return getSheet(COMMENTS_SHEET, COMMENTS_HEADERS); }

function readAll(sheet, headers) {
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(function (row, idx) {
    const obj = { _row: idx + 2 }; // 실제 시트 행 번호 (헤더가 1행이므로 +2)
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function stripFields(obj, fields) {
  const copy = {};
  for (const k in obj) {
    if (fields.indexOf(k) === -1 && k !== '_row') copy[k] = obj[k];
  }
  return copy;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- 이미지 (구글 드라이브) ----
function getImageFolder() {
  const it = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

function uploadImage(image) {
  const folder = getImageFolder();
  const bytes = Utilities.base64Decode(image.data);
  const blob = Utilities.newBlob(bytes, image.mimeType, image.filename || 'image');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return { id: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000' };
}

function deleteImageFile(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { /* 이미 삭제된 경우 무시 */ }
}

// ---- 조회 ----
// 파라미터 없음: 목록(검색/카테고리/페이징) / id: 상세 + 댓글
function doGet(e) {
  const params = e.parameter;
  const postsSheet = getPostsSheet();
  const posts = readAll(postsSheet, POSTS_HEADERS);

  if (params.id) {
    const post = posts.find(function (p) { return p.ID === params.id; });
    if (!post) return jsonOutput({ success: false, error: '게시글을 찾을 수 없습니다.' });

    const newViews = (Number(post.Views) || 0) + 1;
    postsSheet.getRange(post._row, POSTS_HEADERS.indexOf('Views') + 1).setValue(newViews);
    post.Views = newViews;

    const commentsSheet = getCommentsSheet();
    const comments = readAll(commentsSheet, COMMENTS_HEADERS)
      .filter(function (c) { return c.PostID === params.id; })
      .sort(function (a, b) { return new Date(a.CreatedAt) - new Date(b.CreatedAt); })
      .map(function (c) { return stripFields(c, ['Password']); });

    return jsonOutput({ success: true, post: stripFields(post, ['Password']), comments: comments });
  }

  let filtered = posts;
  if (params.category) {
    filtered = filtered.filter(function (p) { return p.Category === params.category; });
  }
  if (params.q) {
    const q = params.q.toLowerCase();
    filtered = filtered.filter(function (p) {
      return String(p.Title).toLowerCase().indexOf(q) !== -1 ||
             String(p.Content).toLowerCase().indexOf(q) !== -1;
    });
  }
  filtered = filtered.slice().sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });

  const total = filtered.length;
  const pageSize = Math.max(1, Number(params.pageSize) || 10);
  const page = Math.max(1, Number(params.page) || 1);
  const start = (page - 1) * pageSize;
  const pageItems = filtered.slice(start, start + pageSize).map(function (p) { return stripFields(p, ['Password']); });

  const categories = posts
    .map(function (p) { return p.Category; })
    .filter(function (c, idx, arr) { return c && arr.indexOf(c) === idx; });

  return jsonOutput({ success: true, posts: pageItems, total: total, page: page, pageSize: pageSize, categories: categories });
}

// ---- 생성/수정/삭제 ----
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  switch (body.action) {
    case 'create': return createPost(body);
    case 'update': return updatePost(body);
    case 'delete': return deletePost(body);
    case 'comment-create': return createComment(body);
    case 'comment-delete': return deleteComment(body);
    default: return jsonOutput({ success: false, error: '알 수 없는 요청입니다.' });
  }
}

function createPost(body) {
  if (!body.title || !body.author || !body.password) {
    return jsonOutput({ success: false, error: '제목/작성자/비밀번호는 필수입니다.' });
  }
  const sheet = getPostsSheet();
  const id = Utilities.getUuid();
  const now = new Date();
  let imageUrl = '', imageFileId = '';
  if (body.image && body.image.data) {
    const uploaded = uploadImage(body.image);
    imageUrl = uploaded.url;
    imageFileId = uploaded.id;
  }
  sheet.appendRow([id, body.title, body.content || '', body.author, body.password,
    body.category || '', imageUrl, imageFileId, 0, now, now]);
  return jsonOutput({ success: true, id: id });
}

function updatePost(body) {
  const sheet = getPostsSheet();
  const posts = readAll(sheet, POSTS_HEADERS);
  const post = posts.find(function (p) { return p.ID === body.id; });
  if (!post) return jsonOutput({ success: false, error: '게시글을 찾을 수 없습니다.' });
  if (String(post.Password) !== String(body.password)) {
    return jsonOutput({ success: false, error: '비밀번호가 일치하지 않습니다.' });
  }
  if (!body.title) return jsonOutput({ success: false, error: '제목은 필수입니다.' });

  let imageUrl = post.ImageUrl, imageFileId = post.ImageFileId;
  if (body.removeImage) {
    deleteImageFile(imageFileId);
    imageUrl = ''; imageFileId = '';
  } else if (body.image && body.image.data) {
    deleteImageFile(imageFileId);
    const uploaded = uploadImage(body.image);
    imageUrl = uploaded.url; imageFileId = uploaded.id;
  }

  const row = post._row;
  sheet.getRange(row, POSTS_HEADERS.indexOf('Title') + 1).setValue(body.title);
  sheet.getRange(row, POSTS_HEADERS.indexOf('Content') + 1).setValue(body.content || '');
  sheet.getRange(row, POSTS_HEADERS.indexOf('Category') + 1).setValue(body.category || '');
  sheet.getRange(row, POSTS_HEADERS.indexOf('ImageUrl') + 1).setValue(imageUrl);
  sheet.getRange(row, POSTS_HEADERS.indexOf('ImageFileId') + 1).setValue(imageFileId);
  sheet.getRange(row, POSTS_HEADERS.indexOf('UpdatedAt') + 1).setValue(new Date());
  return jsonOutput({ success: true });
}

function deletePost(body) {
  const sheet = getPostsSheet();
  const posts = readAll(sheet, POSTS_HEADERS);
  const post = posts.find(function (p) { return p.ID === body.id; });
  if (!post) return jsonOutput({ success: false, error: '게시글을 찾을 수 없습니다.' });
  if (String(post.Password) !== String(body.password)) {
    return jsonOutput({ success: false, error: '비밀번호가 일치하지 않습니다.' });
  }
  deleteImageFile(post.ImageFileId);

  const commentsSheet = getCommentsSheet();
  const relatedComments = readAll(commentsSheet, COMMENTS_HEADERS)
    .filter(function (c) { return c.PostID === body.id; })
    .sort(function (a, b) { return b._row - a._row; }); // 뒤에서부터 삭제해야 행 번호가 안 꼬임
  relatedComments.forEach(function (c) { commentsSheet.deleteRow(c._row); });

  sheet.deleteRow(post._row);
  return jsonOutput({ success: true });
}

function createComment(body) {
  if (!body.postId || !body.author || !body.content || !body.password) {
    return jsonOutput({ success: false, error: '작성자/내용/비밀번호는 필수입니다.' });
  }
  const sheet = getCommentsSheet();
  const id = Utilities.getUuid();
  sheet.appendRow([id, body.postId, body.author, body.content, body.password, new Date()]);
  return jsonOutput({ success: true, id: id });
}

function deleteComment(body) {
  const sheet = getCommentsSheet();
  const comments = readAll(sheet, COMMENTS_HEADERS);
  const comment = comments.find(function (c) { return c.ID === body.id; });
  if (!comment) return jsonOutput({ success: false, error: '댓글을 찾을 수 없습니다.' });
  if (String(comment.Password) !== String(body.password)) {
    return jsonOutput({ success: false, error: '비밀번호가 일치하지 않습니다.' });
  }
  sheet.deleteRow(comment._row);
  return jsonOutput({ success: true });
}
