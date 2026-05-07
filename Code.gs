// ==================== PROJECT MANAGEMENT SYSTEM - BACKEND ====================
// Google Apps Script | Version 1.0.0

const SHEETS = {
  USERS: 'Users',
  PROJECTS: 'Projects',
  TASKS: 'Tasks',
  NOTIFICATIONS: 'Notifications'
};

// ==================== ENTRY POINT ====================
function doGet(e) {
  initializeApp();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Hệ thống Quản lý Dự án')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ==================== ROUTER ====================
function processRequest(action, data) {
  try {
    switch (action) {
      case 'login':                 return login(data);
      case 'logout':                return logout(data);
      case 'validateSession':       return validateSession(data.token);
      case 'getDashboardData':      return getDashboardData(data);
      case 'getProjects':           return getProjects(data);
      case 'getProjectById':        return getProjectById(data);
      case 'createProject':         return createProject(data);
      case 'updateProject':         return updateProject(data);
      case 'deleteProject':         return deleteProject(data);
      case 'getTasks':              return getTasks(data);
      case 'getTaskById':           return getTaskById(data);
      case 'createTask':            return createTask(data);
      case 'updateTask':            return updateTask(data);
      case 'deleteTask':            return deleteTask(data);
      case 'getNotifications':      return getNotifications(data);
      case 'markNotificationRead':  return markNotificationRead(data);
      case 'markAllRead':           return markAllNotificationsRead(data);
      case 'getUsers':              return getUsers(data);
      case 'createUser':            return createUser(data);
      case 'updateUser':            return updateUser(data);
      case 'deleteUser':            return deleteUser(data);
      case 'updateProfile':         return updateProfile(data);
      case 'changePassword':        return changePassword(data);
      default: return { success: false, message: 'Hành động không xác định' };
    }
  } catch (err) {
    Logger.log('Error [' + action + ']: ' + err.message);
    return { success: false, message: 'Lỗi hệ thống: ' + err.message };
  }
}

// ==================== HELPERS ====================
function getSpreadsheet() {
  // 1. Script gắn với Sheet (bound script) — dùng trực tiếp
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  // 2. Standalone script — tìm ID đã lưu trong Script Properties
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      Logger.log('Saved spreadsheet ID invalid, creating new one...');
    }
  }

  // 3. Tự động tạo Google Sheet mới và lưu ID
  const ss = SpreadsheetApp.create('PM_Database - Quản lý Dự án');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  Logger.log('✅ Đã tạo Google Sheet mới: ' + ss.getUrl());
  Logger.log('📋 Spreadsheet ID: ' + ss.getId());
  return ss;
}

function getCache() { return CacheService.getScriptCache(); }
function generateId() { return Utilities.getUuid(); }
function nowIso() { return new Date().toISOString(); }

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); initSheet(name, sheet); }
  return sheet;
}

function initSheet(name, sheet) {
  const headers = {
    Users:         ['id','name','email','password','role','department','isActive','createdAt'],
    Projects:      ['id','code','name','description','managerId','managerName','startDate','endDate','status','budget','createdAt','updatedAt'],
    Tasks:         ['id','projectId','projectName','title','description','assigneeId','assigneeName','reviewerId','reviewerName','startDate','dueDate','status','priority','estimatedCost','actualCost','progress','createdAt','updatedAt','createdBy'],
    Notifications: ['id','userId','title','message','type','isRead','relatedId','relatedType','createdAt']
  };
  const h = headers[name];
  if (h) {
    const r = sheet.getRange(1, 1, 1, h.length);
    r.setValues([h]);
    r.setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
  }
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === id) return i + 1; }
  return -1;
}

function updateRow(sheet, rowIndex, fields) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach(function(h, i) { if (fields[h] !== undefined) row[i] = fields[h]; });
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function getSessionUser(token) {
  if (!token) return null;
  const d = getCache().get(token);
  return d ? JSON.parse(d) : null;
}

function refreshSession(token, user) {
  getCache().put(token, JSON.stringify(user), 28800);
}

// ==================== TIỆN ÍCH SETUP ====================
// Chạy hàm này để xem URL Google Sheet đang dùng
function getSpreadsheetUrl() {
  const ss = getSpreadsheet();
  const url = ss.getUrl();
  Logger.log('🔗 Google Sheet URL: ' + url);
  Logger.log('📋 Spreadsheet ID: ' + ss.getId());
  return url;
}

// Đặt lại Spreadsheet ID thủ công (nếu muốn dùng sheet có sẵn)
function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  Logger.log('✅ Đã lưu Spreadsheet ID: ' + id);
}

// ==================== INIT ====================
function initializeApp() {
  ['Users','Projects','Tasks','Notifications'].forEach(function(n) { getSheet(n); });
  const usersSheet = getSheet(SHEETS.USERS);
  if (sheetToObjects(usersSheet).length === 0) {
    const ts = nowIso();
    usersSheet.appendRow([generateId(),'Administrator','admin@company.com',hashPassword('Admin@123'),'admin','Ban Giám đốc',true,ts]);
    usersSheet.appendRow([generateId(),'Nguyễn Văn Manager','manager@company.com',hashPassword('Manager@123'),'manager','Phòng Kỹ thuật',true,ts]);
    usersSheet.appendRow([generateId(),'Trần Thị Employee','employee@company.com',hashPassword('Employee@123'),'employee','Phòng Kỹ thuật',true,ts]);
  }
}

// ==================== AUTH ====================
function login(data) {
  if (!data.email || !data.password) return { success: false, message: 'Vui lòng nhập đầy đủ thông tin' };
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const hashed = hashPassword(data.password);
  const user = users.find(function(u) { return u.email === data.email && u.password === hashed && u.isActive === true; });
  if (!user) return { success: false, message: 'Email hoặc mật khẩu không chính xác' };
  const token = generateId();
  const session = { userId: user.id, name: user.name, email: user.email, role: user.role, department: user.department };
  getCache().put(token, JSON.stringify(session), 28800);
  return { success: true, token: token, user: session };
}

function logout(data) {
  if (data && data.token) getCache().remove(data.token);
  return { success: true };
}

function validateSession(token) {
  const user = getSessionUser(token);
  if (!user) return { success: false };
  refreshSession(token, user);
  return { success: true, user: user };
}

// ==================== PROJECTS ====================
function getProjects(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const allTasks = sheetToObjects(getSheet(SHEETS.TASKS));
  let projects = sheetToObjects(getSheet(SHEETS.PROJECTS));
  if (user.role !== 'admin') {
    const myProjIds = new Set(allTasks.filter(function(t) { return t.assigneeId === user.userId || t.reviewerId === user.userId; }).map(function(t) { return t.projectId; }));
    projects = projects.filter(function(p) { return myProjIds.has(p.id) || p.managerId === user.userId; });
  }
  const today = new Date();
  projects = projects.map(function(p) {
    const pt = allTasks.filter(function(t) { return t.projectId === p.id; });
    return Object.assign({}, p, {
      taskCount: pt.length,
      completedTasks: pt.filter(function(t) { return t.status === 'completed'; }).length,
      overdueTasks: pt.filter(function(t) { return t.status !== 'completed' && t.dueDate && new Date(t.dueDate) < today; }).length
    });
  });
  return { success: true, data: projects };
}

function getProjectById(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const project = sheetToObjects(getSheet(SHEETS.PROJECTS)).find(function(p) { return p.id === data.id; });
  if (!project) return { success: false, message: 'Dự án không tồn tại' };
  return { success: true, data: project };
}

function createProject(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role === 'employee') return { success: false, message: 'Bạn không có quyền tạo dự án' };
  const sheet = getSheet(SHEETS.PROJECTS);
  if (data.code && sheetToObjects(sheet).find(function(p) { return p.code === data.code; })) return { success: false, message: 'Mã dự án đã tồn tại' };
  const id = generateId(); const ts = nowIso();
  sheet.appendRow([id, data.code||'', data.name, data.description||'', data.managerId||user.userId, data.managerName||user.name, data.startDate||'', data.endDate||'', data.status||'active', Number(data.budget)||0, ts, ts]);
  if (data.managerId && data.managerId !== user.userId) addNotification(data.managerId, 'Dự án mới được giao', 'Bạn quản lý dự án: ' + data.name, 'project', id, 'project');
  return { success: true, id: id, message: 'Tạo dự án thành công' };
}

function updateProject(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role === 'employee') return { success: false, message: 'Bạn không có quyền chỉnh sửa dự án' };
  const sheet = getSheet(SHEETS.PROJECTS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Dự án không tồn tại' };
  updateRow(sheet, rowIndex, { code: data.code, name: data.name, description: data.description, managerId: data.managerId, managerName: data.managerName, startDate: data.startDate, endDate: data.endDate, status: data.status, budget: Number(data.budget)||0, updatedAt: nowIso() });
  return { success: true, message: 'Cập nhật dự án thành công' };
}

function deleteProject(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role !== 'admin') return { success: false, message: 'Chỉ Admin mới có thể xóa dự án' };
  const sheet = getSheet(SHEETS.PROJECTS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Dự án không tồn tại' };
  sheet.deleteRow(rowIndex);
  const tasksSheet = getSheet(SHEETS.TASKS);
  sheetToObjects(tasksSheet).filter(function(t) { return t.projectId === data.id; }).reverse().forEach(function(t) { const r = findRowById(tasksSheet, t.id); if (r !== -1) tasksSheet.deleteRow(r); });
  return { success: true, message: 'Xóa dự án thành công' };
}

// ==================== TASKS ====================
function getTasks(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  let tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  if (data.projectId) tasks = tasks.filter(function(t) { return t.projectId === data.projectId; });
  if (user.role === 'employee') tasks = tasks.filter(function(t) { return t.assigneeId === user.userId || t.reviewerId === user.userId; });
  const today = new Date();
  return { success: true, data: tasks.map(function(t) { return Object.assign({}, t, { isOverdue: t.status !== 'completed' && t.dueDate && new Date(t.dueDate) < today }); }) };
}

function getTaskById(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const task = sheetToObjects(getSheet(SHEETS.TASKS)).find(function(t) { return t.id === data.id; });
  if (!task) return { success: false, message: 'Công việc không tồn tại' };
  return { success: true, data: task };
}

function createTask(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.TASKS);
  const id = generateId(); const ts = nowIso();
  sheet.appendRow([id, data.projectId, data.projectName, data.title, data.description||'', data.assigneeId||'', data.assigneeName||'', data.reviewerId||'', data.reviewerName||'', data.startDate||'', data.dueDate||'', data.status||'todo', data.priority||'medium', Number(data.estimatedCost)||0, Number(data.actualCost)||0, Number(data.progress)||0, ts, ts, user.userId]);
  if (data.assigneeId && data.assigneeId !== user.userId) addNotification(data.assigneeId, 'Công việc mới được giao', 'Bạn được giao: "' + data.title + '" trong dự án ' + data.projectName, 'task', id, 'task');
  if (data.reviewerId && data.reviewerId !== user.userId && data.reviewerId !== data.assigneeId) addNotification(data.reviewerId, 'Được phân công kiểm duyệt', 'Bạn kiểm duyệt: "' + data.title + '"', 'task', id, 'task');
  return { success: true, id: id, message: 'Tạo công việc thành công' };
}

function updateTask(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.TASKS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Công việc không tồn tại' };
  const oldTask = sheetToObjects(sheet).find(function(t) { return t.id === data.id; });
  if (user.role === 'employee' && oldTask && oldTask.assigneeId !== user.userId && oldTask.reviewerId !== user.userId) return { success: false, message: 'Bạn không có quyền chỉnh sửa công việc này' };
  updateRow(sheet, rowIndex, { projectId: data.projectId, projectName: data.projectName, title: data.title, description: data.description, assigneeId: data.assigneeId, assigneeName: data.assigneeName, reviewerId: data.reviewerId, reviewerName: data.reviewerName, startDate: data.startDate, dueDate: data.dueDate, status: data.status, priority: data.priority, estimatedCost: Number(data.estimatedCost)||0, actualCost: Number(data.actualCost)||0, progress: Number(data.progress)||0, updatedAt: nowIso() });
  if (data.status === 'review' && oldTask && oldTask.status !== 'review' && oldTask.reviewerId) addNotification(oldTask.reviewerId, 'Công việc cần kiểm duyệt', '"' + (data.title||oldTask.title) + '" đã hoàn thành và cần kiểm duyệt', 'task', data.id, 'task');
  return { success: true, message: 'Cập nhật công việc thành công' };
}

function deleteTask(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role === 'employee') return { success: false, message: 'Bạn không có quyền xóa công việc' };
  const sheet = getSheet(SHEETS.TASKS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Công việc không tồn tại' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Xóa công việc thành công' };
}

// ==================== NOTIFICATIONS ====================
function addNotification(userId, title, message, type, relatedId, relatedType) {
  getSheet(SHEETS.NOTIFICATIONS).appendRow([generateId(), userId, title, message, type||'info', false, relatedId||'', relatedType||'', nowIso()]);
}

function getNotifications(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const all = sheetToObjects(getSheet(SHEETS.NOTIFICATIONS)).filter(function(n) { return n.userId === user.userId; }).sort(function(a, b) { return new Date(b.createdAt) - new Date(a.createdAt); }).slice(0, 50);
  return { success: true, data: all, unreadCount: all.filter(function(n) { return !n.isRead; }).length };
}

function markNotificationRead(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Thông báo không tồn tại' };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(rowIndex, headers.indexOf('isRead') + 1).setValue(true);
  return { success: true };
}

function markAllNotificationsRead(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const all = sheet.getDataRange().getValues();
  if (all.length <= 1) return { success: true };
  const headers = all[0];
  const uIdx = headers.indexOf('userId'); const rIdx = headers.indexOf('isRead');
  for (let i = 1; i < all.length; i++) { if (all[i][uIdx] === user.userId && !all[i][rIdx]) sheet.getRange(i + 1, rIdx + 1).setValue(true); }
  return { success: true };
}

// ==================== DASHBOARD ====================
function getDashboardData(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  let projects = sheetToObjects(getSheet(SHEETS.PROJECTS));
  let tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  if (user.role === 'employee') {
    const myProjIds = new Set(tasks.filter(function(t) { return t.assigneeId === user.userId || t.reviewerId === user.userId; }).map(function(t) { return t.projectId; }));
    projects = projects.filter(function(p) { return myProjIds.has(p.id) || p.managerId === user.userId; });
    tasks = tasks.filter(function(t) { return t.assigneeId === user.userId || t.reviewerId === user.userId; });
  } else if (user.role === 'manager') {
    const myProjIds = new Set(tasks.filter(function(t) { return t.assigneeId === user.userId || t.reviewerId === user.userId; }).map(function(t) { return t.projectId; }));
    projects = projects.filter(function(p) { return myProjIds.has(p.id) || p.managerId === user.userId; });
  }
  const today = new Date();
  const pStats = { total: projects.length, active: projects.filter(function(p){return p.status==='active';}).length, completed: projects.filter(function(p){return p.status==='completed';}).length, paused: projects.filter(function(p){return p.status==='paused';}).length };
  const tStats = { total: tasks.length, todo: tasks.filter(function(t){return t.status==='todo';}).length, inProgress: tasks.filter(function(t){return t.status==='in_progress';}).length, review: tasks.filter(function(t){return t.status==='review';}).length, completed: tasks.filter(function(t){return t.status==='completed';}).length, overdue: tasks.filter(function(t){return t.status!=='completed'&&t.dueDate&&new Date(t.dueDate)<today;}).length };
  const cStats = { totalEstimated: tasks.reduce(function(s,t){return s+(Number(t.estimatedCost)||0);},0), totalActual: tasks.reduce(function(s,t){return s+(Number(t.actualCost)||0);},0) };
  const monthlyData = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(); d.setMonth(d.getMonth() - i);
    const label = d.toLocaleDateString('vi-VN', { month: 'short', year: 'numeric' });
    const count = tasks.filter(function(t) { if (!t.updatedAt || t.status !== 'completed') return false; const td = new Date(t.updatedAt); return td.getMonth() === d.getMonth() && td.getFullYear() === d.getFullYear(); }).length;
    monthlyData.push({ month: label, completed: count });
  }
  const urgentTasks = tasks.filter(function(t){return t.status!=='completed'&&t.dueDate;}).sort(function(a,b){return new Date(a.dueDate)-new Date(b.dueDate);}).slice(0,5).map(function(t){return Object.assign({},t,{isOverdue:new Date(t.dueDate)<today});});
  const recentProjects = projects.sort(function(a,b){return new Date(b.createdAt)-new Date(a.createdAt);}).slice(0,5);
  return { success: true, data: { projectStats: pStats, taskStats: tStats, costStats: cStats, monthlyData: monthlyData, urgentTasks: urgentTasks, recentProjects: recentProjects } };
}

// ==================== USERS ====================
function getUsers(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const users = sheetToObjects(getSheet(SHEETS.USERS)).map(function(u) { return { id: u.id, name: u.name, email: u.email, role: u.role, department: u.department, isActive: u.isActive, createdAt: u.createdAt }; });
  return { success: true, data: users };
}

function createUser(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role !== 'admin') return { success: false, message: 'Chỉ Admin mới có thể tạo người dùng' };
  const sheet = getSheet(SHEETS.USERS);
  if (sheetToObjects(sheet).find(function(u){return u.email===data.email;})) return { success: false, message: 'Email này đã được sử dụng' };
  const id = generateId();
  sheet.appendRow([id, data.name, data.email, hashPassword(data.password||'Password@123'), data.role||'employee', data.department||'', true, nowIso()]);
  return { success: true, id: id, message: 'Tạo người dùng thành công! Mật khẩu mặc định: Password@123' };
}

function updateUser(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role !== 'admin') return { success: false, message: 'Chỉ Admin mới có thể sửa người dùng' };
  const sheet = getSheet(SHEETS.USERS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Người dùng không tồn tại' };
  const fields = { name: data.name, email: data.email, role: data.role, department: data.department, isActive: data.isActive };
  if (data.resetPassword) fields.password = hashPassword('Password@123');
  updateRow(sheet, rowIndex, fields);
  return { success: true, message: 'Cập nhật người dùng thành công' };
}

function deleteUser(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  if (user.role !== 'admin') return { success: false, message: 'Chỉ Admin mới có thể xóa người dùng' };
  if (data.id === user.userId) return { success: false, message: 'Không thể xóa tài khoản đang đăng nhập' };
  const sheet = getSheet(SHEETS.USERS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Người dùng không tồn tại' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Xóa người dùng thành công' };
}

function updateProfile(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.USERS);
  const rowIndex = findRowById(sheet, user.userId);
  if (rowIndex === -1) return { success: false, message: 'Người dùng không tồn tại' };
  updateRow(sheet, rowIndex, { name: data.name, department: data.department });
  user.name = data.name; user.department = data.department;
  refreshSession(data.token, user);
  return { success: true, message: 'Cập nhật hồ sơ thành công', user: user };
}

function changePassword(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.USERS);
  const current = sheetToObjects(sheet).find(function(u) { return u.id === user.userId; });
  if (!current) return { success: false, message: 'Người dùng không tồn tại' };
  if (current.password !== hashPassword(data.currentPassword)) return { success: false, message: 'Mật khẩu hiện tại không đúng' };
  if (!data.newPassword || data.newPassword.length < 6) return { success: false, message: 'Mật khẩu mới phải có ít nhất 6 ký tự' };
  const rowIndex = findRowById(sheet, user.userId);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(rowIndex, headers.indexOf('password') + 1).setValue(hashPassword(data.newPassword));
  return { success: true, message: 'Đổi mật khẩu thành công' };
}
