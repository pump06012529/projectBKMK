/**
 * School Budget Management System - Backend
 * Developed for 2026 Modern Standards
 * Author: Antigravity AI
 */

const CONFIG = {
  SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  APP_TITLE: "ระบบบริหารงบประมาณโครงการโรงเรียน",
  VERSION: "1.0.0",
  FOLDERS: {
    UPLOADS: "SchoolBudget_Uploads"
  }
};

/**
 * Serves the web application
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Unified API Router
 * Handles all client-side requests
 */
function apiRouter(request) {
  const { action, data } = request;
  const logger = createLogger();
  
  try {
    switch (action) {
      case 'INIT_APP':
        return initDatabase();
      case 'LOGIN':
        return authService.login(data.username, data.password);
      case 'GET_DASHBOARD':
        return dashboardService.getStats();
      case 'PROJECT_GET_ALL':
        return projectService.getAll();
      case 'PROJECT_SAVE':
        return projectService.save(data);
      case 'PROJECT_BULK_SAVE':
        return projectService.bulkSave(data);
      case 'PROJECT_DELETE':
        return projectService.remove(data.id);
      case 'TRANSACTION_GET_ALL':
        return transactionService.getAll();
      case 'TRANSACTION_SAVE':
        return transactionService.save(data);
      case 'TRANSACTION_APPROVE':
        return transactionService.approve(data.id, data.status);
      case 'USER_GET_ALL':
        return userService.getAll();
      case 'USER_SAVE':
        return userService.save(data);
      case 'SETTING_GET':
        return settingService.get();
      case 'UPLOAD_FILE':
        return uploadService.handleFile(data);
      default:
        throw new Error('Unknown action: ' + action);
    }
  } catch (error) {
    logger.log('ERROR', `API Error: ${error.message}`, { action });
    return { success: false, message: error.message };
  }
}

/**
 * DATABASE INITIALIZATION
 */
function initDatabase() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheets = {
    'Users': ['userId', 'fullname', 'username', 'password', 'role', 'department', 'status'],
    'Projects': ['projectId', 'projectName', 'totalBudget', 'budgetType', 'subsidy', 'studentDev', 'manager', 'department', 'disbursementTerm1', 'disbursementTerm2', 'remainingSubsidy', 'remainingStudentDev', 'remainingBudget', 'status', 'createdAt'],
    'Transactions': ['transactionId', 'projectId', 'term', 'budgetCategory', 'amount', 'description', 'receiptUrl', 'status', 'submittedBy', 'approvedBy', 'timestamp'],
    'Departments': ['deptId', 'deptName'],
    'BudgetTypes': ['typeId', 'typeName'],
    'Logs': ['logId', 'userId', 'action', 'details', 'timestamp'],
    'Settings': ['key', 'value']
  };

  Object.keys(sheets).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length).setFontWeight('bold').setBackground('#f3f4f6');
      sheet.setFrozenRows(1);
    }
  });

  // Add Default Admin if no users
  const userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() === 1) {
    const adminData = ['U-001', 'System Administrator', 'admin', hashPassword('admin123'), 'Admin', 'IT', 'Active'];
    userSheet.appendRow(adminData);
  }

  // Add Default Departments
  const deptSheet = ss.getSheetByName('Departments');
  if (deptSheet.getLastRow() === 1) {
    [['D-001', 'ฝ่ายวิชาการ'], ['D-002', 'ฝ่ายงบประมาณ'], ['D-003', 'ฝ่ายบริหารทั่วไป'], ['D-004', 'ฝ่ายบุคคล']].forEach(row => deptSheet.appendRow(row));
  }

  // Add Default Budget Types
  const typeSheet = ss.getSheetByName('BudgetTypes');
  if (typeSheet.getLastRow() === 1) {
    [['B-001', 'งบอุดหนุนรายหัว'], ['B-002', 'งบพัฒนาผู้เรียน'], ['B-003', 'งบรายได้สถานศึกษา']].forEach(row => typeSheet.appendRow(row));
  }

  return { success: true, message: 'Database initialized successfully' };
}

/**
 * SERVICE MODULES
 */

const authService = {
  login(username, password) {
    const users = getSheetData('Users');
    const user = users.find(u => u.username === username && u.password === hashPassword(password));
    if (user) {
      if (user.status !== 'Active') throw new Error('Account is disabled');
      return { success: true, user: { id: user.userId, name: user.fullname, role: user.role, dept: user.department } };
    }
    throw new Error('Invalid username or password');
  }
};

const projectService = {
  getAll() {
    return { success: true, data: getSheetData('Projects') };
  },
  save(data) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Projects');
    const rows = getSheetData('Projects');
    
    if (data.projectId) {
      const index = rows.findIndex(r => r.projectId === data.projectId);
      if (index > -1) {
        const range = sheet.getRange(index + 2, 1, 1, 13);
        const values = [
          data.projectId, data.projectName, data.totalBudget, data.budgetType,
          data.subsidy, data.studentDev, data.manager, data.department,
          data.disbursementTerm1 || 0, data.disbursementTerm2 || 0,
          data.remainingSubsidy || data.subsidy,
          data.remainingStudentDev || data.studentDev,
          data.totalBudget - (data.disbursementTerm1 || 0) - (data.disbursementTerm2 || 0),
          data.status, data.createdAt || new Date()
        ];
        range.setValues([values]);
        return { success: true, message: 'Project updated' };
      }
    }
    
    // New Project
    const newId = 'PJ-' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMM") + '-' + (rows.length + 1).toString().padStart(3, '0');
    const newRow = [
      newId, data.projectName, data.totalBudget, data.budgetType,
      data.subsidy, data.studentDev, data.manager, data.department,
      0, 0, data.subsidy, data.studentDev, data.totalBudget, 'Active', new Date()
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'Project created', id: newId };
  },
  bulkSave(dataArray) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Projects');
    const existing = getSheetData('Projects');
    let count = 0;
    
    dataArray.forEach((data, index) => {
      const newId = 'PJ-' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMM") + '-' + (existing.length + count + 1).toString().padStart(3, '0');
      
      // Map Thai headers to keys
      const pName = data['งาน/กิจกรรม'] || data.projectName || '';
      const bType = data['ประเภทงบประมาณ'] || data.budgetType || 'งบประมาณผสม';
      const mng = data['ผู้รับผิดชอบ'] || data.manager || '';
      const dept = data['ฝ่ายงาน'] || data.department || '';
      const sub = Number(data['อุดหนุนรายหัว'] || data.subsidy || 0);
      const stu = Number(data['พัฒนาผู้เรียน'] || data.studentDev || 0);
      const d1 = Number(data['การเบิกจ่าย เทอม 1'] || data.disbursementTerm1 || 0);
      const d2 = Number(data['การเบิกจ่าย เทอม 2'] || data.disbursementTerm2 || 0);
      
      const total = sub + stu;
      
      const newRow = [
        newId, pName, total, bType,
        sub, stu, mng, dept,
        d1, d2, sub - d1, stu - d2, total - d1 - d2, 'Active', new Date()
      ];
      sheet.appendRow(newRow);
      count++;
    });
    return { success: true, message: `Imported ${count} projects successfully` };
  },
  remove(id) {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('Projects');
    const data = getSheetData('Projects');
    const index = data.findIndex(r => r.projectId === id);
    if (index > -1) {
      sheet.deleteRow(index + 2);
      return { success: true, message: 'Project deleted' };
    }
    throw new Error('Project not found');
  }
};

const transactionService = {
  getAll() {
    return { success: true, data: getSheetData('Transactions') };
  },
  save(data) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Transactions');
    const newId = 'TRX-' + Date.now();
    
    // Validation
    const project = getSheetData('Projects').find(p => p.projectId === data.projectId);
    if (!project) throw new Error('Project not found');
    if (data.amount > project.remainingBudget) throw new Error('Insufficient budget');

    const newRow = [
      newId, data.projectId, data.term, data.budgetCategory, data.amount, data.description,
      data.receiptUrl || '', 'Pending', data.submittedBy, '', new Date()
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'Transaction submitted for approval' };
  },
  approve(id, status) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Transactions');
    const rows = getSheetData('Transactions');
    const index = rows.findIndex(r => r.transactionId === id);
    
    if (index > -1) {
      const trx = rows[index];
      sheet.getRange(index + 2, 7).setValue(status); // Status
      sheet.getRange(index + 2, 9).setValue(Session.getActiveUser().getEmail()); // ApprovedBy
      
      if (status === 'Approved') {
        updateProjectBudget(trx.projectId, trx.term, trx.amount);
      }
      return { success: true, message: `Transaction ${status}` };
    }
    throw new Error('Transaction not found');
  }
};

const userService = {
  getAll() {
    const data = getSheetData('Users').map(u => {
      delete u.password;
      return u;
    });
    return { success: true, data };
  },
  save(data) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Users');
    const rows = getSheetData('Users');
    
    if (data.userId) {
      const index = rows.findIndex(r => r.userId === data.userId);
      if (index > -1) {
        sheet.getRange(index + 2, 2).setValue(data.fullname);
        sheet.getRange(index + 2, 5).setValue(data.role);
        sheet.getRange(index + 2, 6).setValue(data.department);
        sheet.getRange(index + 2, 7).setValue(data.status);
        if (data.password) sheet.getRange(index + 2, 4).setValue(hashPassword(data.password));
        return { success: true, message: 'User updated' };
      }
    }
    
    const newId = 'U-' + (rows.length + 1).toString().padStart(3, '0');
    const newRow = [newId, data.fullname, data.username, hashPassword(data.password), data.role, data.department, 'Active'];
    sheet.appendRow(newRow);
    return { success: true, message: 'User created' };
  }
};

const dashboardService = {
  getStats() {
    const projects = getSheetData('Projects');
    const trxs = getSheetData('Transactions');
    
    const totalBudget = projects.reduce((sum, p) => sum + Number(p.totalBudget), 0);
    const totalSpent = projects.reduce((sum, p) => sum + (Number(p.disbursementTerm1) + Number(p.disbursementTerm2)), 0);
    const remaining = totalBudget - totalSpent;
    
    return {
      success: true,
      stats: {
        projectsCount: projects.length,
        totalBudget,
        totalSpent,
        remaining,
        pendingApprovals: trxs.filter(t => t.status === 'Pending').length
      },
      charts: {
        byDept: aggregateBy(projects, 'department', 'totalBudget'),
        byType: aggregateBy(projects, 'budgetType', 'totalBudget'),
        monthly: trxs.filter(t => t.status === 'Approved').reduce((acc, t) => {
          const month = Utilities.formatDate(new Date(t.timestamp), "GMT+7", "MMM");
          acc[month] = (acc[month] || 0) + Number(t.amount);
          return acc;
        }, {})
      }
    };
  }
};

const uploadService = {
  handleFile(data) {
    const folder = getOrCreateFolder(CONFIG.FOLDERS.UPLOADS);
    const contentType = data.content.substring(5, data.content.indexOf(';'));
    const bytes = Utilities.base64Decode(data.content.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, data.name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl(), id: file.getId() };
  }
};

/**
 * HELPERS
 */
function getSheetData(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function hashPassword(pass) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass);
  return digest.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

function updateProjectBudget(projectId, term, amount) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('Projects');
  const data = getSheetData('Projects');
  const index = data.findIndex(p => p.projectId === projectId);
  
  if (index > -1) {
    const col = term == 1 ? 9 : 10;
    const current = Number(sheet.getRange(index + 2, col).getValue());
    sheet.getRange(index + 2, col).setValue(current + Number(amount));
    
    // Update category-specific remaining
    if (trx.budgetCategory === 'งบอุดหนุนรายหัว') {
      const rem = Number(sheet.getRange(index + 2, 11).getValue());
      sheet.getRange(index + 2, 11).setValue(rem - Number(amount));
    } else if (trx.budgetCategory === 'งบกิจกรรมพัฒนาผู้เรียน') {
      const rem = Number(sheet.getRange(index + 2, 12).getValue());
      sheet.getRange(index + 2, 12).setValue(rem - Number(amount));
    }

    // Recalculate total remaining
    const total = Number(sheet.getRange(index + 2, 3).getValue());
    const t1 = Number(sheet.getRange(index + 2, 9).getValue());
    const t2 = Number(sheet.getRange(index + 2, 10).getValue());
    sheet.getRange(index + 2, 13).setValue(total - t1 - t2);
  }
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function aggregateBy(data, key, valKey) {
  return data.reduce((acc, item) => {
    acc[item[key]] = (acc[item[key]] || 0) + Number(item[valKey]);
    return acc;
  }, {});
}

function createLogger() {
  return {
    log(level, message, details) {
      const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      const sheet = ss.getSheetByName('Logs');
      const user = Session.getActiveUser().getEmail();
      sheet.appendRow([Date.now(), user, level, message, JSON.stringify(details), new Date()]);
    }
  };
}
