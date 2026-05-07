// ==================== PROJECT MANAGEMENT SYSTEM - BACKEND ====================
// Google Apps Script | Version 2.0.0 | Project-level RBAC

const SHEETS = {
  USERS: 'Users',
  PROJECTS: 'Projects',
  TASKS: 'Tasks',
  NOTIFICATIONS: 'Notifications',
  PROJECT_MEMBERS: 'ProjectMembers',
  SPRINTS: 'Sprints'
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
      case 'getProjectMembers':     return getProjectMembers(data);
      case 'addProjectMember':      return addProjectMember(data);
      case 'updateProjectMember':   return updateProjectMember(data);
      case 'removeProjectMember':   return removeProjectMember(data);
      case 'getTasks':              return getTasks(data);
      case 'getTaskById':           return getTaskById(data);
      case 'createTask':            return createTask(data);
      case 'updateTask':            return updateTask(data);
      case 'deleteTask':            return deleteTask(data);
      case 'getSprints':            return getSprints(data);
      case 'createSprint':          return createSprint(data);
      case 'updateSprint':          return updateSprint(data);
      case 'startSprint':           return startSprint(data);
      case 'completeSprint':        return completeSprint(data);
      case 'deleteSprint':          return deleteSprint(data);
      case 'moveIssuesToSprint':    return moveIssuesToSprint(data);
      case 'updateStoryPoints':     return updateStoryPoints(data);
      case 'updateBacklogOrder':    return updateBacklogOrder(data);
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
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) { Logger.log('Saved spreadsheet ID invalid, creating new one...'); }
  }
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
    Users:          ['id','name','email','password','role','department','isActive','createdAt'],
    Projects:       ['id','code','name','description','managerId','managerName','startDate','endDate','status','budget','createdAt','updatedAt','projectType'],
    Tasks:          ['id','projectId','projectName','title','description','assigneeId','assigneeName','reviewerId','reviewerName','startDate','dueDate','status','priority','estimatedCost','actualCost','progress','createdAt','updatedAt','createdBy','issueType','parentId','sprintId','storyPoints','backlogOrder'],
    Sprints:        ['id','projectId','name','goal','startDate','endDate','status','capacity','totalSpSnapshot','createdAt','createdBy'],
    Notifications:  ['id','userId','title','message','type','isRead','relatedId','relatedType','createdAt'],
    ProjectMembers: ['id','projectId','userId','userName','projectRole','addedAt','addedBy']
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
  const tz = Session.getScriptTimeZone();
  return data.slice(1).filter(function(row) {
    return row[0] !== '' && row[0] !== null && row[0] !== undefined;
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[h] = val;
    });
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

// ==================== PERMISSION HELPERS ====================
function getProjectRoleFrom(members, userId, projectId) {
  const entry = members.find(function(m) { return m.projectId === projectId && m.userId === userId; });
  return entry ? entry.projectRole : null;
}

// Can user see this project at all?
function canAccessProject(user, projectId, members) {
  if (user.role === 'admin') return true;
  return getProjectRoleFrom(members, user.userId, projectId) !== null;
}

// Can user edit project info, manage members, delete tasks?
function canManageProject(user, projectId, members) {
  if (user.role === 'admin') return true;
  return getProjectRoleFrom(members, user.userId, projectId) === 'owner';
}

// Can user create or edit tasks (owner or member)?
function canContributeToProject(user, projectId, members) {
  if (user.role === 'admin') return true;
  const role = getProjectRoleFrom(members, user.userId, projectId);
  return role === 'owner' || role === 'member';
}

function addMemberEntry(projectId, userId, userName, projectRole, addedBy) {
  getSheet(SHEETS.PROJECT_MEMBERS).appendRow([generateId(), projectId, userId, userName, projectRole, nowIso(), addedBy]);
}

// ==================== TIỆN ÍCH SETUP ====================
function getSpreadsheetUrl() {
  const ss = getSpreadsheet();
  const url = ss.getUrl();
  Logger.log('🔗 Google Sheet URL: ' + url);
  Logger.log('📋 Spreadsheet ID: ' + ss.getId());
  return url;
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  Logger.log('✅ Đã lưu Spreadsheet ID: ' + id);
}

// ==================== INIT ====================
function initializeApp() {
  ['Users','Projects','Tasks','Notifications','ProjectMembers','Sprints'].forEach(function(n) { getSheet(n); });
  const usersSheet = getSheet(SHEETS.USERS);
  if (sheetToObjects(usersSheet).length === 0) {
    const ts = nowIso();
    usersSheet.appendRow([generateId(),'Administrator','admin@company.com',hashPassword('Admin@123'),'admin','Ban Giám đốc',true,ts]);
    usersSheet.appendRow([generateId(),'Nguyễn Văn Manager','manager@company.com',hashPassword('Manager@123'),'manager','Phòng Kỹ thuật',true,ts]);
    usersSheet.appendRow([generateId(),'Trần Thị Employee','employee@company.com',hashPassword('Employee@123'),'employee','Phòng Kỹ thuật',true,ts]);
  }
  migrateProjectMembers();
}

// Back-fills ProjectMembers entries for pre-existing projects that have none yet.
function migrateProjectMembers() {
  const membersSheet = getSheet(SHEETS.PROJECT_MEMBERS);
  const existing = sheetToObjects(membersSheet);
  const projects = sheetToObjects(getSheet(SHEETS.PROJECTS));
  if (!projects.length) return;

  const needsMigration = projects.filter(function(p) {
    return !existing.some(function(m) { return m.projectId === p.id; });
  });
  if (!needsMigration.length) return;

  const tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const ts = nowIso();

  needsMigration.forEach(function(p) {
    const seen = new Set();
    if (p.managerId) {
      const mgr = users.find(function(u) { return u.id === p.managerId; });
      membersSheet.appendRow([generateId(), p.id, p.managerId, (mgr ? mgr.name : p.managerName) || '', 'owner', ts, 'migration']);
      seen.add(p.managerId);
    }
    tasks.filter(function(t) { return t.projectId === p.id; }).forEach(function(t) {
      [t.assigneeId, t.reviewerId].forEach(function(uid) {
        if (uid && !seen.has(uid)) {
          seen.add(uid);
          const u = users.find(function(u) { return u.id === uid; });
          if (u) membersSheet.appendRow([generateId(), p.id, uid, u.name, 'member', ts, 'migration']);
        }
      });
    });
  });
}

// ==================== AUTH ====================
function login(data) {
  if (!data.email || !data.password) return { success: false, message: 'Vui lòng nhập đầy đủ thông tin' };
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const hashed = hashPassword(data.password);
  const user = users.find(function(u) { return u.email === data.email && u.password === hashed && !!u.isActive; });
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
  const allMembers = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  const allTasks = sheetToObjects(getSheet(SHEETS.TASKS));
  let projects = sheetToObjects(getSheet(SHEETS.PROJECTS));

  if (user.role !== 'admin') {
    const myProjIds = new Set(
      allMembers.filter(function(m) { return m.userId === user.userId; }).map(function(m) { return m.projectId; })
    );
    projects = projects.filter(function(p) { return myProjIds.has(p.id); });
  }

  const today = new Date();
  projects = projects.map(function(p) {
    const pt = allTasks.filter(function(t) { return t.projectId === p.id; });
    const userProjectRole = user.role === 'admin' ? 'admin' : getProjectRoleFrom(allMembers, user.userId, p.id);
    return Object.assign({}, p, {
      taskCount: pt.length,
      completedTasks: pt.filter(function(t) { return t.status === 'completed'; }).length,
      overdueTasks: pt.filter(function(t) { return t.status !== 'completed' && t.dueDate && new Date(t.dueDate) < today; }).length,
      userProjectRole: userProjectRole,
      memberCount: allMembers.filter(function(m) { return m.projectId === p.id; }).length
    });
  });
  return { success: true, data: projects };
}

function getProjectById(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canAccessProject(user, data.id, members)) return { success: false, message: 'Bạn không có quyền truy cập dự án này' };
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
  const managerId = data.managerId || user.userId;
  const managerName = data.managerName || user.name;
  sheet.appendRow([id, data.code||'', data.name, data.description||'', managerId, managerName, data.startDate||'', data.endDate||'', data.status||'active', Number(data.budget)||0, ts, ts, data.projectType||'waterfall']);

  // Auto-add creator as owner
  addMemberEntry(id, user.userId, user.name, 'owner', user.userId);
  // Also add assigned manager as owner if different from creator
  if (managerId !== user.userId) {
    const users = sheetToObjects(getSheet(SHEETS.USERS));
    const mgr = users.find(function(u) { return u.id === managerId; });
    if (mgr) addMemberEntry(id, managerId, managerName, 'owner', user.userId);
    addNotification(managerId, 'Dự án mới được giao', 'Bạn quản lý dự án: ' + data.name, 'project', id, 'project');
  }
  return { success: true, id: id, message: 'Tạo dự án thành công' };
}

function updateProject(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.id, members)) return { success: false, message: 'Bạn không có quyền chỉnh sửa dự án này' };
  const sheet = getSheet(SHEETS.PROJECTS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Dự án không tồn tại' };
  updateRow(sheet, rowIndex, { code: data.code, name: data.name, description: data.description, managerId: data.managerId, managerName: data.managerName, startDate: data.startDate, endDate: data.endDate, status: data.status, budget: Number(data.budget)||0, updatedAt: nowIso(), projectType: data.projectType||'waterfall' });
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
  // Cascade: delete all tasks
  const tasksSheet = getSheet(SHEETS.TASKS);
  sheetToObjects(tasksSheet).filter(function(t) { return t.projectId === data.id; }).reverse().forEach(function(t) {
    const r = findRowById(tasksSheet, t.id); if (r !== -1) tasksSheet.deleteRow(r);
  });
  // Cascade: delete all member entries
  const membersSheet = getSheet(SHEETS.PROJECT_MEMBERS);
  sheetToObjects(membersSheet).filter(function(m) { return m.projectId === data.id; }).reverse().forEach(function(m) {
    const r = findRowById(membersSheet, m.id); if (r !== -1) membersSheet.deleteRow(r);
  });
  return { success: true, message: 'Xóa dự án thành công' };
}

// ==================== PROJECT MEMBERS ====================
function getProjectMembers(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canAccessProject(user, data.projectId, members)) return { success: false, message: 'Không có quyền truy cập dự án' };
  const projMembers = members.filter(function(m) { return m.projectId === data.projectId; });
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const enriched = projMembers.map(function(m) {
    const u = users.find(function(u) { return u.id === m.userId; }) || {};
    return { id: m.id, projectId: m.projectId, userId: m.userId, userName: m.userName || u.name || '', projectRole: m.projectRole, addedAt: m.addedAt, systemRole: u.role || '', department: u.department || '', email: u.email || '' };
  });
  return { success: true, data: enriched };
}

function addProjectMember(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền quản lý thành viên dự án này' };
  if (members.find(function(m) { return m.projectId === data.projectId && m.userId === data.userId; })) {
    return { success: false, message: 'Người dùng đã là thành viên của dự án' };
  }
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const targetUser = users.find(function(u) { return u.id === data.userId; });
  if (!targetUser) return { success: false, message: 'Người dùng không tồn tại' };
  addMemberEntry(data.projectId, data.userId, targetUser.name, data.projectRole || 'member', user.userId);
  const project = sheetToObjects(getSheet(SHEETS.PROJECTS)).find(function(p) { return p.id === data.projectId; });
  const roleLabel = { owner: 'Quản lý dự án', member: 'Thành viên', viewer: 'Người xem' }[data.projectRole] || 'Thành viên';
  addNotification(data.userId, 'Được thêm vào dự án', 'Bạn được thêm vào dự án "' + (project ? project.name : '') + '" với vai trò: ' + roleLabel, 'project', data.projectId, 'project');
  return { success: true, message: 'Thêm thành viên thành công' };
}

function updateProjectMember(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền quản lý thành viên dự án này' };
  const entry = members.find(function(m) { return m.projectId === data.projectId && m.userId === data.userId; });
  if (!entry) return { success: false, message: 'Thành viên không tồn tại trong dự án' };
  if (entry.projectRole === 'owner' && data.projectRole !== 'owner') {
    const owners = members.filter(function(m) { return m.projectId === data.projectId && m.projectRole === 'owner'; });
    if (owners.length <= 1) return { success: false, message: 'Dự án phải có ít nhất một Quản lý dự án' };
  }
  const sheet = getSheet(SHEETS.PROJECT_MEMBERS);
  const rowIndex = findRowById(sheet, entry.id);
  if (rowIndex === -1) return { success: false, message: 'Lỗi tìm bản ghi thành viên' };
  updateRow(sheet, rowIndex, { projectRole: data.projectRole });
  return { success: true, message: 'Cập nhật quyền thành công' };
}

function removeProjectMember(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền quản lý thành viên dự án này' };
  const entry = members.find(function(m) { return m.projectId === data.projectId && m.userId === data.userId; });
  if (!entry) return { success: false, message: 'Thành viên không tồn tại trong dự án' };
  if (entry.projectRole === 'owner') {
    const owners = members.filter(function(m) { return m.projectId === data.projectId && m.projectRole === 'owner'; });
    if (owners.length <= 1) return { success: false, message: 'Không thể xóa Quản lý dự án duy nhất' };
  }
  const sheet = getSheet(SHEETS.PROJECT_MEMBERS);
  const rowIndex = findRowById(sheet, entry.id);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
  return { success: true, message: 'Đã xóa thành viên khỏi dự án' };
}

// ==================== TASKS ====================
function getTasks(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  let tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  if (user.role !== 'admin') {
    const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
    const myProjIds = new Set(members.filter(function(m) { return m.userId === user.userId; }).map(function(m) { return m.projectId; }));
    tasks = tasks.filter(function(t) { return myProjIds.has(t.projectId); });
  }
  if (data.projectId) tasks = tasks.filter(function(t) { return t.projectId === data.projectId; });
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

function validateParentChild(issueType, parentTask) {
  if (issueType === 'epic') {
    if (parentTask) return 'Epic không được có công việc cha';
  } else if (issueType === 'task' || issueType === 'user_story') {
    if (!parentTask) return 'Task/User Story phải thuộc một Epic';
    if ((parentTask.issueType || 'task') !== 'epic') return 'Task/User Story chỉ có thể thuộc Epic';
  } else if (issueType === 'bug') {
    if (parentTask && (parentTask.issueType || 'task') !== 'epic') return 'Bug chỉ có thể thuộc Epic (nếu có)';
  } else if (issueType === 'sub_task') {
    if (!parentTask) return 'Sub-task phải có công việc cha';
    if (['task','user_story','bug'].indexOf(parentTask.issueType || 'task') === -1) return 'Sub-task chỉ có thể thuộc Task, User Story hoặc Bug';
  }
  return null;
}

function createTask(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canContributeToProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền tạo công việc trong dự án này (cần vai trò Thành viên trở lên)' };
  const issueType = data.issueType || 'task';
  const allTasks = sheetToObjects(getSheet(SHEETS.TASKS));
  const parentTask = data.parentId ? allTasks.find(function(t) { return t.id === data.parentId; }) : null;
  const validErr = validateParentChild(issueType, parentTask);
  if (validErr) return { success: false, message: validErr };
  const sheet = getSheet(SHEETS.TASKS);
  const id = generateId(); const ts = nowIso();
  sheet.appendRow([id, data.projectId, data.projectName, data.title, data.description||'', data.assigneeId||'', data.assigneeName||'', data.reviewerId||'', data.reviewerName||'', data.startDate||'', data.dueDate||'', data.status||'todo', data.priority||'medium', Number(data.estimatedCost)||0, Number(data.actualCost)||0, Number(data.progress)||0, ts, ts, user.userId, issueType, data.parentId||'', data.sprintId||'', data.storyPoints!==undefined&&data.storyPoints!==''?Number(data.storyPoints):'', Number(data.backlogOrder)||0]);
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
  if (user.role !== 'admin') {
    const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
    const projRole = getProjectRoleFrom(members, user.userId, oldTask.projectId);
    if (!projRole) return { success: false, message: 'Bạn không có quyền truy cập dự án này' };
    if (projRole === 'viewer') return { success: false, message: 'Bạn chỉ có quyền xem dự án này' };
    if (projRole === 'member' && oldTask.assigneeId !== user.userId && oldTask.reviewerId !== user.userId) {
      return { success: false, message: 'Thành viên chỉ được chỉnh sửa công việc được phân công cho mình' };
    }
  }
  const newIssueType = data.issueType || oldTask.issueType || 'task';
  const newParentTask = data.parentId ? sheetToObjects(sheet).find(function(t) { return t.id === data.parentId; }) : null;
  const validErr2 = validateParentChild(newIssueType, newParentTask);
  if (validErr2) return { success: false, message: validErr2 };
  var spVal = data.storyPoints !== undefined ? (data.storyPoints === '' ? '' : Number(data.storyPoints)) : (oldTask.storyPoints !== undefined ? oldTask.storyPoints : '');
  updateRow(sheet, rowIndex, { projectId: data.projectId, projectName: data.projectName, title: data.title, description: data.description, assigneeId: data.assigneeId, assigneeName: data.assigneeName, reviewerId: data.reviewerId, reviewerName: data.reviewerName, startDate: data.startDate, dueDate: data.dueDate, status: data.status, priority: data.priority, estimatedCost: Number(data.estimatedCost)||0, actualCost: Number(data.actualCost)||0, progress: Number(data.progress)||0, updatedAt: nowIso(), issueType: newIssueType, parentId: data.parentId||'', sprintId: data.sprintId !== undefined ? data.sprintId : (oldTask.sprintId||''), storyPoints: spVal, backlogOrder: data.backlogOrder !== undefined ? Number(data.backlogOrder)||0 : Number(oldTask.backlogOrder)||0 });
  if (data.status === 'review' && oldTask && oldTask.status !== 'review' && oldTask.reviewerId) addNotification(oldTask.reviewerId, 'Công việc cần kiểm duyệt', '"' + (data.title||oldTask.title) + '" đã hoàn thành và cần kiểm duyệt', 'task', data.id, 'task');
  return { success: true, message: 'Cập nhật công việc thành công' };
}

function deleteTask(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sheet = getSheet(SHEETS.TASKS);
  const allTasks = sheetToObjects(sheet);
  const task = allTasks.find(function(t) { return t.id === data.id; });
  if (!task) return { success: false, message: 'Công việc không tồn tại' };
  if (user.role !== 'admin') {
    const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
    if (!canManageProject(user, task.projectId, members)) return { success: false, message: 'Chỉ Quản lý dự án hoặc Admin mới có thể xóa công việc' };
  }
  // Cascade: task + direct children + their sub-tasks
  var toDeleteIds = [data.id];
  allTasks.filter(function(t) { return t.parentId === data.id; }).forEach(function(c) {
    toDeleteIds.push(c.id);
    allTasks.filter(function(t) { return t.parentId === c.id; }).forEach(function(gc) { toDeleteIds.push(gc.id); });
  });
  var toDeleteSet = {};
  toDeleteIds.forEach(function(id) { toDeleteSet[id] = true; });
  var allData = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = 1; i < allData.length; i++) {
    if (toDeleteSet[String(allData[i][0])]) rowsToDelete.push(i + 1);
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });
  var msg = toDeleteIds.length > 1 ? 'Đã xóa công việc và ' + (toDeleteIds.length - 1) + ' công việc con' : 'Xóa công việc thành công';
  return { success: true, message: msg };
}

// ==================== SPRINTS ====================
function getSprints(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const sprints = sheetToObjects(getSheet(SHEETS.SPRINTS)).filter(function(s) { return s.projectId === data.projectId; });
  const tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  const SP_ISSUE_TYPES = ['task','user_story','bug'];
  const result = sprints.map(function(s) {
    const st = tasks.filter(function(t) { return t.sprintId === s.id; });
    const spEligible = st.filter(function(t) { return SP_ISSUE_TYPES.indexOf(t.issueType||'task') !== -1 && t.storyPoints !== '' && t.storyPoints !== null && t.storyPoints !== undefined; });
    const totalSp = spEligible.reduce(function(a, t) { return a + Number(t.storyPoints); }, 0);
    const completedSp = spEligible.filter(function(t) { return t.status === 'completed'; }).reduce(function(a, t) { return a + Number(t.storyPoints); }, 0);
    return Object.assign({}, s, { totalSp: totalSp, completedSp: completedSp, issueCount: st.length });
  });
  return { success: true, data: result.sort(function(a,b){ const ord={planning:0,active:1,completed:2}; return (ord[a.status]||0)-(ord[b.status]||0); }) };
}

function createSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền tạo Sprint' };
  const id = generateId(); const ts = nowIso();
  getSheet(SHEETS.SPRINTS).appendRow([id, data.projectId, data.name, data.goal||'', data.startDate||'', data.endDate||'', 'planning', Number(data.capacity)||0, 0, ts, user.userId]);
  return { success: true, id: id, message: 'Tạo Sprint thành công' };
}

function updateSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền cập nhật Sprint' };
  const sheet = getSheet(SHEETS.SPRINTS);
  const rowIndex = findRowById(sheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Sprint không tồn tại' };
  updateRow(sheet, rowIndex, { name: data.name, goal: data.goal||'', startDate: data.startDate||'', endDate: data.endDate||'', capacity: Number(data.capacity)||0 });
  return { success: true, message: 'Cập nhật Sprint thành công' };
}

function startSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền bắt đầu Sprint' };
  const sprintsSheet = getSheet(SHEETS.SPRINTS);
  const allSprints = sheetToObjects(sprintsSheet);
  if (allSprints.some(function(s) { return s.projectId === data.projectId && s.status === 'active'; })) return { success: false, message: 'Đã có Sprint đang chạy trong dự án này. Kết thúc Sprint hiện tại trước.' };
  const rowIndex = findRowById(sprintsSheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Sprint không tồn tại' };
  const tasks = sheetToObjects(getSheet(SHEETS.TASKS));
  const totalSp = tasks.filter(function(t) { return t.sprintId === data.id && ['task','user_story','bug'].indexOf(t.issueType||'task') !== -1 && t.storyPoints !== '' && t.storyPoints !== null && t.storyPoints !== undefined; }).reduce(function(a, t) { return a + Number(t.storyPoints); }, 0);
  updateRow(sprintsSheet, rowIndex, { status: 'active', totalSpSnapshot: totalSp, startDate: data.startDate || nowIso().split('T')[0] });
  return { success: true, message: 'Sprint đã bắt đầu!' };
}

function completeSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền kết thúc Sprint' };
  const sprintsSheet = getSheet(SHEETS.SPRINTS);
  const rowIndex = findRowById(sprintsSheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Sprint không tồn tại' };
  const taskSheet = getSheet(SHEETS.TASKS);
  const tasks = sheetToObjects(taskSheet);
  var moved = 0;
  tasks.forEach(function(t) {
    if (t.sprintId === data.id && t.status !== 'completed') {
      const tr = findRowById(taskSheet, t.id);
      if (tr !== -1) { updateRow(taskSheet, tr, { sprintId: '', backlogOrder: 0 }); moved++; }
    }
  });
  updateRow(sprintsSheet, rowIndex, { status: 'completed' });
  return { success: true, message: 'Sprint kết thúc. ' + moved + ' issue chưa done đã chuyển về Backlog.' };
}

function deleteSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canManageProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền xóa Sprint' };
  const sprintsSheet = getSheet(SHEETS.SPRINTS);
  const sprint = sheetToObjects(sprintsSheet).find(function(s) { return s.id === data.id; });
  if (!sprint) return { success: false, message: 'Sprint không tồn tại' };
  if (sprint.status === 'active') return { success: false, message: 'Không thể xóa Sprint đang chạy' };
  const taskSheet = getSheet(SHEETS.TASKS);
  sheetToObjects(taskSheet).forEach(function(t) {
    if (t.sprintId === data.id) { const tr = findRowById(taskSheet, t.id); if (tr !== -1) updateRow(taskSheet, tr, { sprintId: '' }); }
  });
  const rowIndex = findRowById(sprintsSheet, data.id);
  if (rowIndex !== -1) sprintsSheet.deleteRow(rowIndex);
  return { success: true, message: 'Đã xóa Sprint' };
}

function moveIssuesToSprint(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canContributeToProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền' };
  const taskSheet = getSheet(SHEETS.TASKS);
  var count = 0;
  (data.issueIds || []).forEach(function(issueId) {
    const rowIndex = findRowById(taskSheet, issueId);
    if (rowIndex !== -1) { updateRow(taskSheet, rowIndex, { sprintId: data.sprintId || '' }); count++; }
  });
  return { success: true, message: count + ' issue đã chuyển' };
}

function updateStoryPoints(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
  if (!canContributeToProject(user, data.projectId, members)) return { success: false, message: 'Bạn không có quyền' };
  const taskSheet = getSheet(SHEETS.TASKS);
  const rowIndex = findRowById(taskSheet, data.id);
  if (rowIndex === -1) return { success: false, message: 'Issue không tồn tại' };
  updateRow(taskSheet, rowIndex, { storyPoints: data.storyPoints !== '' ? Number(data.storyPoints) : '' });
  return { success: true, message: 'Cập nhật Story Points thành công' };
}

function updateBacklogOrder(data) {
  const user = getSessionUser(data.token);
  if (!user) return { success: false, message: 'Phiên đăng nhập hết hạn' };
  const taskSheet = getSheet(SHEETS.TASKS);
  (data.orders || []).forEach(function(item) {
    const rowIndex = findRowById(taskSheet, item.id);
    if (rowIndex !== -1) updateRow(taskSheet, rowIndex, { backlogOrder: Number(item.backlogOrder) || 0 });
  });
  return { success: true };
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
  if (user.role !== 'admin') {
    const members = sheetToObjects(getSheet(SHEETS.PROJECT_MEMBERS));
    const myProjIds = new Set(members.filter(function(m) { return m.userId === user.userId; }).map(function(m) { return m.projectId; }));
    projects = projects.filter(function(p) { return myProjIds.has(p.id); });
    tasks = tasks.filter(function(t) { return myProjIds.has(t.projectId); });
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
