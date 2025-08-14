// ===============================================================
// CONFIGURATION
// ===============================================================
const CONFIG = {
  COURSES_SHEET: 'Courses', // <-- THÊM DÒNG NÀY
  STUDENTS_SHEET: 'Students',
  CLASSES_SHEET: 'Classes',
  ENROLLMENTS_SHEET: 'Enrollments',
  CLASS_SESSIONS_SHEET: 'ClassSessions',
  ATTENDANCE_SHEET: 'Attendance',
  CONFIG_SHEET: 'Config',
  FEE_TYPES_SHEET: 'FeeTypes',
  TRANSACTIONS_SHEET: 'Transactions',
  USERS_SHEET: 'PhanQuyen',
  TEACHERS_SHEET: 'Teachers', // Them dong nay
  EXPENSES_SHEET: 'Expenses', // Them dong nay
  DEFAULT_TIMEZONE: 'Asia/Ho_Chi_Minh',
  STUDENT_ID_PREFIX: 'HV', // <-- Them dong nay
  CLASS_ID_PREFIX: 'C',
  ENROLLMENT_ID_PREFIX: 'E',
  SESSION_ID_PREFIX: 'SES',
  ATTENDANCE_ID_PREFIX: 'ATT',
  TRANSACTION_ID_PREFIX: 'GD'
};

// ===============================================================
// WEB APP SERVING
// ===============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('Quản Lý Trung Tâm')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===============================================================
// AUTHENTICATION & INITIAL DATA
// ===============================================================
function login(credentials) {
  const lock = getLock();
  try {
    const { username, password } = credentials; // Changed from email to username
    if (!username || !password) {
      throw new Error("Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu.");
    }

    const usersSheet = getSheet(CONFIG.USERS_SHEET);
    const usersData = usersSheet.getDataRange().getValues();
    
    // Compare username (column A) and password (column B)
    const userRow = usersData.find(row => row[0].toLowerCase() === username.toLowerCase() && row[1].toString() === password.toString());

    if (!userRow) {
      return { success: false, error: "Sai tên đăng nhập hoặc mật khẩu. Vui lòng thử lại." };
    }

    const userRole = userRow[2]; // Column C for Role
    const userEmail = userRow[0]; // Keep this for display purposes, it's the username now
    const teacherId = userRow[3] || null; // Lấy Teacher ID từ cột D (index 3)

    const config = getConfigData();

    return {
      success: true,
      data: {
        config,
        userRole,
        userEmail, // This variable now holds the username
        teacherId // Thêm teacherId vào dữ liệu trả về
      }
    };

  } catch (error) {
    Logger.log("!!! LỖI trong hàm login: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error);
    return { success: false, error: analysis };
  } finally {
    lock.releaseLock();
  }
}

// ===============================================================
// VALIDATION FUNCTIONS
// ===============================================================
function validateFormData(data) {
    const errors = [];
    if (data.name && data.name.trim() === '') errors.push('Tên không được để trống');
    if (data.date && !/^\d{4}-\d{2}-\d{2}$/.test(data.date)) errors.push('Định dạng ngày không hợp lệ');
    if (data.amount && (isNaN(data.amount) || Number(data.amount) <= 0)) errors.push('Số tiền không hợp lệ');
    if (errors.length > 0) throw new Error(errors.join('. '));
    return true;
}

// ===============================================================
// CONCURRENCY CONTROL
// ===============================================================
function getLock() {
    const lock = LockService.getScriptLock();
    const success = lock.tryLock(30000);
    if (!success) throw new Error('Hệ thống đang bận, vui lòng thử lại sau.');
    return lock;
}


// ===============================================================
// DATA RETRIEVAL
// ===============================================================
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Lỗi: Không tìm thấy trang tính có tên "${sheetName}". Vui lòng kiểm tra lại.`);
  return sheet;
}

function getConfigData() {
  const sheet = getSheet(CONFIG.CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return {};
  const headers = data.shift();
  const config = {};
  headers.forEach((header, index) => {
    if (header) {
      config[header] = data.map(row => row[index]).filter(String);
    }
  });
  return config;
}


// Thay the ham getInitialData cu bang phien ban nay
function getInitialData() {
  try {
    const config = getConfigData();
    // Giả lập vai trò Admin để có toàn quyền truy cập giao diện
    const userRole = 'Admin';
    const userEmail = 'public.user'; // Một tên định danh chung

    return { success: true, data: { config, userRole, userEmail } };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getInitialData: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getTransactionsData() {
  try {
    const transactionsSheet = getSheet(CONFIG.TRANSACTIONS_SHEET);
    const transactionsData = transactionsSheet.getLastRow() > 1 ? transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 8).getValues() : [];
    const transactions = transactionsData.map(row => ({
      id: row[0], studentId: row[1], enrollmentId: row[2], amount: row[3], 
      date: row[4] ? Utilities.formatDate(new Date(row[4]), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd') : '',
      content: row[5], collector: row[6], method: row[7]
    })).filter(t => t.id);
    return { success: true, data: { transactions } };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getTransactionsData: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getDashboardData() {
  try {
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const transactionsSheet = getSheet(CONFIG.TRANSACTIONS_SHEET);
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const expensesSheet = getSheet(CONFIG.EXPENSES_SHEET); // Them dong nay

    // Tinh Doanh thu & Chi phi trong thang nay
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);

    const transactionsData = transactionsSheet.getLastRow() > 1 ? transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 8).getValues() : [];
    const revenueThisMonth = transactionsData.reduce((sum, row) => {
      const transactionDate = new Date(row[4]);
      if (transactionDate >= startOfMonth && transactionDate <= endOfMonth) {
        return sum + Number(row[3] || 0);
      }
      return sum;
    }, 0);

    const expensesData = expensesSheet.getLastRow() > 1 ? expensesSheet.getRange(2, 1, expensesSheet.getLastRow() - 1, 6).getValues() : [];
    const expenseThisMonth = expensesData.reduce((sum, row) => {
      const expenseDate = new Date(row[1]);
      if (expenseDate >= startOfMonth && expenseDate <= endOfMonth) {
        return sum + Number(row[2] || 0);
      }
      return sum;
    }, 0);

    const profitThisMonth = revenueThisMonth - expenseThisMonth;

    // Cac chi so khac
    const totalStudents = studentsSheet.getLastRow() > 1 ? studentsSheet.getLastRow() - 1 : 0;
    const totalClasses = classesSheet.getLastRow() > 1 ? classesSheet.getLastRow() - 1 : 0;

    // Lich hoc hom nay (giu nguyen)
    const sessionsData = sessionsSheet.getLastRow() > 1 ? sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 6).getValues() : [];
    const todayStr = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
    const classMap = new Map(classesSheet.getLastRow() > 1 ? classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, 2).getValues().map(row => [row[0], row[1]]) : []);
    const todaysSchedule = sessionsData.filter(row => Utilities.formatDate(new Date(row[2]), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd') === todayStr)
      .map(row => ({
        className: classMap.get(row[1]) || 'N/A',
        startTime: Utilities.formatDate(new Date(row[3]), CONFIG.DEFAULT_TIMEZONE, 'HH:mm'),
        endTime: Utilities.formatDate(new Date(row[4]), CONFIG.DEFAULT_TIMEZONE, 'HH:mm'),
        teacher: row[5] || ''
      })).sort((a,b) => a.startTime.localeCompare(b.startTime));

    return {
      success: true,
      data: {
        totalStudents,
        totalClasses,
        revenueThisMonth,
        expenseThisMonth,
        profitThisMonth,
        todaysSchedule
      }
    };

  } catch (error) {
    Logger.log("!!! LỖI trong hàm getDashboardData: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getRevenueReport(options) {
  try {    
    const { startDate, endDate } = options;
    if (!startDate || !endDate) {
      throw new Error("Vui lòng cung cấp ngày bắt đầu và ngày kết thúc.");
    }

    const transactionsSheet = getSheet(CONFIG.TRANSACTIONS_SHEET);
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);

    // Tao map de tra cuu ten hoc vien va ten lop cho nhanh
    const studentData = studentsSheet.getLastRow() > 1 ? studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 2).getValues() : [];
    const studentMap = new Map(studentData.map(row => [row[0], row[1]]));

    const classData = classesSheet.getLastRow() > 1 ? classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, 2).getValues() : [];
    const classMap = new Map(classData.map(row => [row[0], row[1]]));

    const enrollmentData = enrollmentsSheet.getLastRow() > 1 ? enrollmentsSheet.getRange(2, 1, enrollmentsSheet.getLastRow() - 1, 3).getValues() : [];
    const enrollmentMap = new Map(enrollmentData.map(row => [row[0], row[2]])); // Map tu enrollmentId -> classId

    const allTransactions = transactionsSheet.getLastRow() > 1 ? transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 8).getValues() : [];
    
    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23, 59, 59, 999); // Tinh den het ngay ket thuc

    const filteredTransactions = allTransactions.filter(row => {
      const transactionDate = new Date(row[4]);
      return transactionDate >= startDateObj && transactionDate <= endDateObj;
    });

    let totalRevenue = 0;
    const transactionDetails = filteredTransactions.map(row => {
      totalRevenue += Number(row[3] || 0);
      const classId = enrollmentMap.get(row[2]);
      return {
        date: Utilities.formatDate(new Date(row[4]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy'),
        studentId: row[1],
        studentName: studentMap.get(row[1]) || 'Không rõ',
        amount: Number(row[3] || 0),
        content: row[5],
        className: classMap.get(classId) || 'N/A',
        collector: row[6],
        method: row[7]
      };
    }).sort((a, b) => new Date(b.date.split('/').reverse().join('-')) - new Date(a.date.split('/').reverse().join('-')));

    return {
      success: true,
      data: {
        totalRevenue: totalRevenue,
        transactions: transactionDetails
      }
    };

  } catch (error) {
    Logger.log("!!! LỖI trong hàm getRevenueReport: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getDebtReport() {
  try {
    // *** THAY DOI O DAY: Loc hoc vien "Dang hoc" ***
    const allStudents = getSheet(CONFIG.STUDENTS_SHEET).getDataRange().getValues().slice(1);
    const students = allStudents.filter(row => row[5] === 'Đang học');

    const enrollments = getSheet(CONFIG.ENROLLMENTS_SHEET).getDataRange().getValues().slice(1);
    const feeTypes = getSheet(CONFIG.FEE_TYPES_SHEET).getDataRange().getValues().slice(1);
    const transactions = getSheet(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues().slice(1);

    const feeTypeMap = new Map(feeTypes.map(ft => [ft[0], Number(ft[2] || 0)]));
    
    const totalPaidMap = new Map();
    transactions.forEach(t => {
      const studentId = t[1];
      const amount = Number(t[3] || 0);
      totalPaidMap.set(studentId, (totalPaidMap.get(studentId) || 0) + amount);
    });

    const debtList = [];
    students.forEach(s => {
      const studentId = s[0];
      const studentName = s[1];
      const studentPhone = s[4];

      const studentEnrollments = enrollments.filter(e => e[1] === studentId);
      const totalDue = studentEnrollments.reduce((sum, e) => {
        const feeAmount = feeTypeMap.get(e[4]) || 0;
        return sum + feeAmount;
      }, 0);
      
      const totalPaid = totalPaidMap.get(studentId) || 0;
      const remaining = totalDue - totalPaid;

      if (remaining > 0) {
        debtList.push({
          studentId: studentId,
          studentName: studentName,
          phone: studentPhone,
          totalDue: totalDue,
          totalPaid: totalPaid,
          remaining: remaining
        });
      }
    });

    return {
      success: true,
      data: debtList.sort((a,b) => b.remaining - a.remaining)
    };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getDebtReport: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getAttendanceReport(options) {
  try {    
    const { classId, startDate, endDate } = options;
    if (!classId || !startDate || !endDate) {
      throw new Error("Vui lòng cung cấp Lớp, ngày bắt đầu và ngày kết thúc.");
    }

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23, 59, 59, 999);

    const allEnrollments = getSheet(CONFIG.ENROLLMENTS_SHEET).getDataRange().getValues().slice(1);
    const studentIdsInClass = allEnrollments.filter(e => e[2] === classId).map(e => e[1]);
    if (studentIdsInClass.length === 0) { return { success: true, data: [] }; }

    const allStudents = getSheet(CONFIG.STUDENTS_SHEET).getDataRange().getValues().slice(1);
    const studentMap = new Map(allStudents.map(s => [s[0], s[1]]));

    const allSessions = getSheet(CONFIG.CLASS_SESSIONS_SHEET).getDataRange().getValues().slice(1);
    const sessionIdsInRange = allSessions
      .filter(s => {
        const sessionDate = new Date(s[2]);
        return s[1] === classId && sessionDate >= startDateObj && sessionDate <= endDateObj;
      })
      .map(s => s[0]);
    if (sessionIdsInRange.length === 0) { return { success: true, data: [] }; }

    const allAttendance = getSheet(CONFIG.ATTENDANCE_SHEET).getDataRange().getValues().slice(1);
    const relevantAttendance = allAttendance.filter(a => sessionIdsInRange.includes(a[1]));

    const reportData = [];
    for (const studentId of studentIdsInClass) {
      const studentAttendance = relevantAttendance.filter(a => a[2] === studentId);
      
      // *** LOGIC THAY DOI NAM O DAY ***
      const attended = studentAttendance.filter(a => a[3] === 'present').length;
      const absent = studentAttendance.filter(a => a[3] === 'absent' || a[3] === 'excused').length;
      // *** KET THUC THAY DOI ***

      const totalRecorded = attended + absent;

      if (totalRecorded > 0) {
        reportData.push({
          studentId: studentId,
          studentName: studentMap.get(studentId) || 'Không rõ',
          attended: attended,
          absent: absent,
          total: totalRecorded,
          rate: Math.round((attended / totalRecorded) * 100)
        });
      }
    }
    return { success: true, data: reportData.sort((a,b) => a.studentName.localeCompare(b.studentName)) };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getAttendanceReport: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function getStudentProfile(studentId) {
  try {
    // 1. Lay thong tin chi tiet hoc vien
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const allStudents = studentsSheet.getDataRange().getValues();
    const studentRow = allStudents.find(row => row[0] === studentId);
    if (!studentRow) throw new Error("Không tìm thấy học viên.");
    const studentDetails = {
      id: studentRow[0],
      name: studentRow[1],
      dob: studentRow[2] ? Utilities.formatDate(new Date(studentRow[2]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy') : '',
      parentName: studentRow[3],
      phone: studentRow[4]
    };

    // Lay du lieu tu cac sheet lien quan
    const allEnrollments = getSheet(CONFIG.ENROLLMENTS_SHEET).getDataRange().getValues().slice(1);
    const allTransactions = getSheet(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues().slice(1);
    const allAttendance = getSheet(CONFIG.ATTENDANCE_SHEET).getDataRange().getValues().slice(1);
    const classMap = new Map(getSheet(CONFIG.CLASSES_SHEET).getDataRange().getValues().slice(1).map(c => [c[0], c[1]]));
    const feeTypeMap = new Map(getSheet(CONFIG.FEE_TYPES_SHEET).getDataRange().getValues().slice(1).map(f => [f[0], { name: f[1], amount: f[2] }]));
    const sessionMap = new Map(getSheet(CONFIG.CLASS_SESSIONS_SHEET).getDataRange().getValues().slice(1).map(s => [s[0], { date: s[2], classId: s[1] }]));

    // 2. Xu ly cac lop da ghi danh
    const enrollments = allEnrollments.filter(e => e[1] === studentId).map(e => {
      const feeType = feeTypeMap.get(e[4]);
      return {
        className: classMap.get(e[2]) || 'N/A',
        enrollmentDate: e[3] ? Utilities.formatDate(new Date(e[3]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy') : '',
        feeTypeName: feeType ? feeType.name : 'Chưa có',
        feeAmount: feeType ? feeType.amount : 0
      };
    });

    // 3. Xu ly lich su giao dich
    const transactions = allTransactions.filter(t => t[1] === studentId).map(t => ({
      date: t[4] ? Utilities.formatDate(new Date(t[4]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy') : '',
      amount: Number(t[3] || 0),
      content: t[5],
      method: t[7]
    })).sort((a,b) => new Date(b.date.split('/').reverse().join('-')) - new Date(a.date.split('/').reverse().join('-')));

    // 4. Xu ly lich su diem danh
    const attendance = allAttendance.filter(a => a[2] === studentId).map(a => {
      const session = sessionMap.get(a[1]);
      return {
        date: session ? Utilities.formatDate(new Date(session.date), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy') : 'N/A',
        className: session ? classMap.get(session.classId) : 'N/A',
        status: a[3],
        note: a[4]
      };
    }).sort((a,b) => new Date(b.date.split('/').reverse().join('-')) - new Date(a.date.split('/').reverse().join('-')));

    return {
      success: true,
      data: {
        details: studentDetails,
        enrollments: enrollments,
        transactions: transactions,
        attendance: attendance
      }
    };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getStudentProfile: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

// ===============================================================
// PRINTING DATA FUNCTIONS
// ===============================================================

function getFeeNoticeData(studentId) {
  try {
    if (!studentId) throw new Error("Cần có mã học viên.");

    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const allStudents = studentsSheet.getDataRange().getValues();
    const studentRow = allStudents.find(row => row[0] === studentId);
    if (!studentRow) throw new Error("Không tìm thấy học viên.");

    const enrollments = getSheet(CONFIG.ENROLLMENTS_SHEET).getDataRange().getValues().slice(1).filter(e => e[1] === studentId);
    const transactions = getSheet(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues().slice(1).filter(t => t[1] === studentId);
    
    const feeTypeMap = new Map(getSheet(CONFIG.FEE_TYPES_SHEET).getDataRange().getValues().slice(1).map(f => [f[0], { name: f[1], amount: Number(f[2] || 0) }]));
    const classMap = new Map(getSheet(CONFIG.CLASSES_SHEET).getDataRange().getValues().slice(1).map(c => [c[0], c[1]]));

    let totalDue = 0;
    const feeDetails = enrollments.map(e => {
      const classId = e[2];
      const feeTypeId = e[4];
      const feeType = feeTypeMap.get(feeTypeId);
      if (feeType) {
        totalDue += feeType.amount;
        const className = classMap.get(classId) || 'N/A';
        return `<tr><td>Học phí lớp ${className} - ${feeType.name}</td><td>${feeType.amount.toLocaleString('vi-VN')}</td></tr>`;
      }
      return '';
    }).join('');

    const totalPaid = transactions.reduce((sum, t) => sum + Number(t[3] || 0), 0);
    const remaining = totalDue - totalPaid;

    return {
      success: true,
      data: {
        studentId: studentRow[0],
        studentName: studentRow[1],
        parentName: studentRow[3],
        phone: studentRow[4],
        totalDue: totalDue.toLocaleString('vi-VN'),
        totalPaid: totalPaid.toLocaleString('vi-VN'),
        remaining: remaining.toLocaleString('vi-VN'),
        feeDetails: feeDetails
      }
    };
  } catch (e) {
    Logger.log("Lỗi khi lấy dữ liệu phiếu báo học phí: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getReceiptData(studentId) {
  try {
    if (!studentId) throw new Error("Cần có mã học viên.");
    
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const allStudents = studentsSheet.getDataRange().getValues();
    const studentRow = allStudents.find(row => row[0] === studentId);
    if (!studentRow) throw new Error("Không tìm thấy học viên.");

    const transactions = getSheet(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues().slice(1).filter(t => t[1] === studentId);
    if (transactions.length === 0) {
      throw new Error("Học viên này chưa có thanh toán nào để in biên lai.");
    }
    
    let totalPaid = 0;
    const transactionDetails = transactions.map(t => {
      totalPaid += Number(t[3] || 0);
      const date = Utilities.formatDate(new Date(t[4]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy');
      return `<tr>
                <td>${date}</td>
                <td>${t[5] || 'Thanh toán học phí'}</td>
                <td>${t[7] || 'N/A'}</td>
                <td>${Number(t[3] || 0).toLocaleString('vi-VN')}</td>
              </tr>`;
    }).join('');

    return {
      success: true,
      data: {
        studentId: studentRow[0],
        studentName: studentRow[1],
        parentName: studentRow[3],
        phone: studentRow[4],
        totalPaid: totalPaid.toLocaleString('vi-VN'),
        transactionDetails: transactionDetails
      }
    };
  } catch (e) {
    Logger.log("Lỗi khi lấy dữ liệu biên lai: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ===============================================================
// UTILITY FUNCTIONS
// ===============================================================
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function generateNextId(sheet, column, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return `${prefix}001`;
  const columnValues = sheet.getRange(`${column}2:${column}${lastRow}`).getValues();
  const numericIds = columnValues
    .flat()
    .filter(id => id && typeof id === 'string' && id.startsWith(prefix))
    .map(id => {
      const numericPart = id.replace(/\D/g, '');
      return numericPart ? parseInt(numericPart, 10) : 0;
    });
  const nextIdNumber = (numericIds.length > 0 ? Math.max(...numericIds) : 0) + 1;
  return `${prefix}${String(nextIdNumber).padStart(3, '0')}`;
}

// ===============================================================
// === HÀM MỚI: TẠO MÃ LỚP HỌC TỰ ĐỘNG ===
// ===============================================================

/**
 * Hàm tạo mã lớp học mới tự động theo cấu trúc [CourseID]-K[NămTháng]-[SốThứTự]
 * @param {string} courseId - Mã của khóa học gốc, ví dụ "IE-6.5".
 * @param {string} startDateString - Ngày khai giảng dưới dạng chuỗi "yyyy-MM-dd".
 * @returns {string} Mã lớp học mới, ví dụ "IE-6.5-K2508-01".
 */
function generateNewClassId(courseId, startDateString) {
  const sheet = getSheet(CONFIG.CLASSES_SHEET);
  // Lấy toàn bộ dữ liệu ở cột A (cột mã lớp học) để kiểm tra
  const classIdColumn = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat() : [];

  // 1. Tạo phần tiền tố từ CourseID và ngày khai giảng
  const startDate = new Date(startDateString);
  const year = startDate.getFullYear().toString().slice(-2); // Lấy 2 số cuối của năm, ví dụ 2025 -> 25
  const month = (startDate.getMonth() + 1).toString().padStart(2, '0'); // Lấy tháng, đảm bảo có 2 chữ số, ví dụ 8 -> 08
  const prefix = `${courseId}-K${year}${month}`; // Ví dụ: IE-6.5-K2508

  // 2. Tìm số thứ tự (sequence) lớn nhất của các lớp có cùng tiền tố
  let maxSequence = 0;
  classIdColumn.forEach(existingId => {
    // Chỉ xét những mã lớp có cùng tiền tố
    if (existingId && existingId.startsWith(prefix)) {
      const parts = existingId.split('-');
      // Lấy phần tử cuối cùng (số thứ tự) và chuyển thành số
      const sequence = parseInt(parts[parts.length - 1], 10);
      if (sequence > maxSequence) {
        maxSequence = sequence;
      }
    }
  });

  // 3. Tạo mã hoàn chỉnh với số thứ tự tiếp theo, đảm bảo có 2 chữ số
  const nextSequence = (maxSequence + 1).toString().padStart(2, '0'); // Ví dụ 1 -> 01, 2 -> 02
  return `${prefix}-${nextSequence}`;
}

// ===============================================================
// CLASS MANAGEMENT
// ===============================================================
// Thay thế hàm addClass
function addClass(data) {
  const lock = getLock();
  try {
    // Dữ liệu nhận từ frontend sẽ có cấu trúc mới
    const { courseId, teacherId, startDate, maxSize, lichHoc, gioHoc } = data;

    // Kiểm tra các thông tin đầu vào quan trọng
    if (!courseId || !startDate) {
      throw new Error("Vui lòng chọn Khóa học và nhập Ngày khai giảng.");
    }

    // 1. Gọi hàm mới để tạo mã lớp học tự động
    const newClassId = generateNewClassId(courseId, startDate);

    const sheet = getSheet(CONFIG.CLASSES_SHEET);

    // 2. Lưu vào sheet với cấu trúc mới đã thống nhất
    // Cột A: Mã lớp mới (VD: IE-6.5-K2508-01)
    // Cột B: Mã khóa học gốc (VD: IE-6.5)
    // Các cột sau giữ nguyên
    sheet.appendRow([
        newClassId, 
        courseId, 
        teacherId || '', 
        maxSize || '', 
        lichHoc || '', 
        gioHoc || ''
    ]);
    
    return { success: true, message: `Đã tạo thành công lớp học với mã: ${newClassId}` };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm addClass: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error);
    return { success: false, error: analysis };
  } finally {
    lock.releaseLock();
  }
}

// Thay thế hàm editClass
function editClass(data) {
  const lock = getLock();
  try {
    const { id, teacherId, maxSize, lichHoc, gioHoc } = data; // Bỏ courseId vì không cho sửa
    const sheet = getSheet(CONFIG.CLASSES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === id);

    if (rowIndex === -1) throw new Error("Không tìm thấy lớp học để cập nhật.");

    // Dữ liệu cũ của hàng đó
    let rowData = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Chỉ cập nhật các cột được phép thay đổi
    // Cột C (index 2): teacherId, D(3): maxSize, E(4): lichHoc, F(5): gioHoc
    sheet.getRange(rowIndex + 1, 3, 1, 4).setValues([[
        teacherId || '',
        maxSize || '',
        lichHoc || '',
        gioHoc || ''
    ]]);

    return { success: true, message: "Cập nhật thông tin lớp học thành công." };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm editClass: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error);
    return { success: false, error: analysis };
  } finally {
    lock.releaseLock();
  }
}

function deleteClass(classId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const hasStudents = enrollmentsSheet.getDataRange().getValues().some(row => row[2] === classId);
    if (hasStudents) throw new Error('Không thể xóa lớp học đang có học viên.');
    const hasSessions = sessionsSheet.getDataRange().getValues().some(row => row[1] === classId);
    if (hasSessions) throw new Error('Không thể xóa lớp học đang có buổi học đã lên lịch.');
    const classData = classesSheet.getDataRange().getValues();
    const rowIndex = classData.findIndex(row => row[0] === classId);
    if (rowIndex === -1) throw new Error('Không tìm thấy lớp học để xóa.');
    classesSheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'Đã xóa lớp học thành công' };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm deleteClass: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// SESSION MANAGEMENT
// ===============================================================
function addSession(data) {
  const lock = getLock();
  try {
    const { classId, date, startTime, endTime, teacherId: substituteTeacherId } = data;

    // Lấy giáo viên chính của lớp
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const allClasses = classesSheet.getDataRange().getValues();
    const classRow = allClasses.find(row => row[0] === classId);
    if (!classRow) throw new Error("Không tìm thấy lớp học.");
    const mainTeacherId = classRow[2];

    // Giáo viên của buổi học này là giáo viên dạy thay, nếu không có thì là giáo viên chính
    const effectiveTeacherId = substituteTeacherId || mainTeacherId;

    if (effectiveTeacherId) {
        const newSessionStart = new Date(`${date}T${startTime}:00`);
        const newSessionEnd = new Date(`${date}T${endTime}:00`);

        const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
        const allExistingSessions = sessionsSheet.getDataRange().getValues();
        const teacherExistingSessions = allExistingSessions.filter(sessionRow => {
            const sessionClassId = sessionRow[1];
            const sessionTeacherId = sessionRow[5];
            const sessionClassInfo = allClasses.find(c => c[0] === sessionClassId);
            const classMainTeacherId = sessionClassInfo ? sessionClassInfo[2] : null;
            return sessionTeacherId === effectiveTeacherId || (!sessionTeacherId && classMainTeacherId === effectiveTeacherId);
        });

        for (const existingSession of teacherExistingSessions) {
            const existingStart = new Date(existingSession[3]);
            const existingEnd = new Date(existingSession[4]);
            if (newSessionStart < existingEnd && newSessionEnd > existingStart) {
                const conflictingClassInfo = allClasses.find(c => c[0] === existingSession[1]);
                throw new Error(`TRÙNG LỊCH: Giáo viên đã có lịch dạy lớp "${conflictingClassInfo[0]}" lúc ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'HH:mm')} ngày ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy')}.`);
            }
        }
    }

    const sheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const newSessionId = generateNextId(sheet, "A", CONFIG.SESSION_ID_PREFIX);
    sheet.appendRow([newSessionId, classId, new Date(date), new Date(`${date}T${startTime}:00`), new Date(`${date}T${endTime}:00`), substituteTeacherId || '']);
    return { success: true, message: `Đã thêm buổi học mới` };
  } catch (e) {
    Logger.log("!!! LỖI trong hàm addSession: " + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Thay thế hàm editSession
function editSession(data) {
  const lock = getLock();
  try {
    const { id, classId, date, startTime, endTime, teacherId: substituteTeacherId } = data;
    
    // Logic kiểm tra trùng lịch tương tự addSession
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const allClasses = classesSheet.getDataRange().getValues();
    const classRow = allClasses.find(row => row[0] === classId);
    if (!classRow) throw new Error("Không tìm thấy lớp học.");
    const mainTeacherId = classRow[2];
    const effectiveTeacherId = substituteTeacherId || mainTeacherId;

    if (effectiveTeacherId) {
        const newSessionStart = new Date(`${date}T${startTime}:00`);
        const newSessionEnd = new Date(`${date}T${endTime}:00`);

        const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
        const allExistingSessions = sessionsSheet.getDataRange().getValues();
        const teacherExistingSessions = allExistingSessions.filter(sessionRow => {
            // Loại trừ chính buổi học đang sửa ra khỏi danh sách kiểm tra
            if (sessionRow[0] === id) return false;
            
            const sessionClassId = sessionRow[1];
            const sessionTeacherId = sessionRow[5];
            const sessionClassInfo = allClasses.find(c => c[0] === sessionClassId);
            const classMainTeacherId = sessionClassInfo ? sessionClassInfo[2] : null;
            return sessionTeacherId === effectiveTeacherId || (!sessionTeacherId && classMainTeacherId === effectiveTeacherId);
        });

        for (const existingSession of teacherExistingSessions) {
             const existingStart = new Date(existingSession[3]);
             const existingEnd = new Date(existingSession[4]);
             if (newSessionStart < existingEnd && newSessionEnd > existingStart) {
                 const conflictingClassInfo = allClasses.find(c => c[0] === existingSession[1]);
                 throw new Error(`TRÙNG LỊCH: Giáo viên đã có lịch dạy lớp "${conflictingClassInfo[0]}" lúc ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'HH:mm')} ngày ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy')}.`);
             }
        }
    }

    const sheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === id);
    if(rowIndex === -1) throw new Error("Không tìm thấy buổi học");
    sheet.getRange(rowIndex + 1, 2, 1, 5).setValues([[ classId, new Date(date), new Date(`${date}T${startTime}:00`), new Date(`${date}T${endTime}:00`), substituteTeacherId || '' ]]);
    return { success: true, message: `Đã cập nhật buổi học` };
  } catch (e) {
    Logger.log("!!! LỖI trong hàm editSession: " + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Thay thế hàm deleteSession
function deleteSession(sessionId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const attendanceSheet = getSheet(CONFIG.ATTENDANCE_SHEET);
    const hasAttendance = attendanceSheet.getDataRange().getValues().some(row => row[1] === sessionId);
    if(hasAttendance) throw new Error("Không thể xóa buổi học đã có dữ liệu điểm danh.");
    const allData = sessionsSheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === sessionId);
    if(rowIndex === -1) throw new Error("Không tìm thấy buổi học");
    sessionsSheet.deleteRow(rowIndex + 1);
    return { success: true, message: "Đã xóa buổi học" };
  } catch(e) {
    Logger.log("!!! LỖI trong hàm deleteSession: " + e.toString());
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// Thay thế hàm createRecurringSessions
function createRecurringSessions(data) {
  const lock = getLock();
  try {
    const { classId, weekdays, startTime, endTime, startDate: startDateStr, endDate: endDateStr } = data;
    if (!weekdays || weekdays.length === 0) throw new Error("Vui lòng chọn ít nhất một ngày trong tuần.");
    
    // === PHẦN LOGIC KIỂM TRA TRÙNG LỊCH - BẮT ĐẦU ===

    // 1. Lấy thông tin giáo viên của lớp đang được xếp lịch
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const allClasses = classesSheet.getDataRange().getValues();
    const classRow = allClasses.find(row => row[0] === classId);
    if (!classRow) throw new Error("Không tìm thấy lớp học.");
    const teacherId = classRow[2]; // Cột C là TeacherID

    let newSessionsToCreate = []; // Mảng chứa các buổi học mới sẽ được tạo

    // Chỉ kiểm tra lịch nếu lớp này có giáo viên được phân công
    if (teacherId) {
        const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
        const allExistingSessions = sessionsSheet.getDataRange().getValues();

        // 2. Lấy TẤT CẢ các buổi học đã có của giáo viên này
        const teacherExistingSessions = allExistingSessions.filter(sessionRow => {
            const sessionClassId = sessionRow[1];
            const sessionTeacherId = sessionRow[5]; // GV dạy thay trong buổi học
            const sessionClassInfo = allClasses.find(c => c[0] === sessionClassId);
            const mainTeacherId = sessionClassInfo ? sessionClassInfo[2] : null; // GV chính của lớp
            return sessionTeacherId === teacherId || (!sessionTeacherId && mainTeacherId === teacherId);
        });

        // 3. Vòng lặp qua các ngày dự định tạo để kiểm tra
        let currentDate = new Date(startDateStr);
        let endDate = new Date(endDateStr);
        while (currentDate <= endDate) {
            if (weekdays.map(Number).includes(currentDate.getDay())) {
                const dateKey = Utilities.formatDate(currentDate, CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
                const newSessionStart = new Date(`${dateKey}T${startTime}:00`);
                const newSessionEnd = new Date(`${dateKey}T${endTime}:00`);

                // 4. Với mỗi buổi học mới, so sánh với toàn bộ lịch đã có của giáo viên
                for (const existingSession of teacherExistingSessions) {
                    const existingStart = new Date(existingSession[3]);
                    const existingEnd = new Date(existingSession[4]);

                    // Công thức kiểm tra 2 khoảng thời gian giao nhau: (StartA < EndB) AND (EndA > StartB)
                    if (newSessionStart < existingEnd && newSessionEnd > existingStart) {
                        const conflictingClassInfo = allClasses.find(c => c[0] === existingSession[1]);
                        const conflictingClassName = conflictingClassInfo ? conflictingClassInfo[0] : 'Không rõ'; // Lấy mã lớp bị trùng
                        // Nếu trùng, báo lỗi ngay lập tức và dừng lại
                        throw new Error(
                            `TRÙNG LỊCH: Giáo viên đã có lịch dạy lớp "${conflictingClassName}" ` +
                            `lúc ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'HH:mm')} ` +
                            `ngày ${Utilities.formatDate(existingStart, CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy')}.`
                        );
                    }
                }
                // Nếu không trùng, thêm vào danh sách chờ tạo
                newSessionsToCreate.push({ date: new Date(dateKey), startTime: newSessionStart, endTime: newSessionEnd });
            }
            currentDate.setDate(currentDate.getDate() + 1);
        }
    } else {
        // Nếu lớp không có giáo viên, chỉ cần tạo danh sách các buổi học
        let currentDate = new Date(startDateStr);
        let endDate = new Date(endDateStr);
        while (currentDate <= endDate) {
            if (weekdays.map(Number).includes(currentDate.getDay())) {
                 const dateKey = Utilities.formatDate(currentDate, CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
                 newSessionsToCreate.push({ 
                     date: new Date(dateKey), 
                     startTime: new Date(`${dateKey}T${startTime}:00`), 
                     endTime: new Date(`${dateKey}T${endTime}:00`) 
                 });
            }
            currentDate.setDate(currentDate.getDate() + 1);
        }
    }
    // === KẾT THÚC PHẦN LOGIC KIỂM TRA ===
    
    // 5. Nếu không có lỗi nào, tiến hành lưu các buổi học vào Sheet
    if (newSessionsToCreate.length > 0) {
      const sheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
      const rowsToAdd = newSessionsToCreate.map(session => {
        const newSessionId = generateNextId(sheet, "A", CONFIG.SESSION_ID_PREFIX) + `-${Math.random()}`; // Thêm số ngẫu nhiên để ID không trùng trong 1 lần
        return [newSessionId, classId, session.date, session.startTime, session.endTime, '']; // Để trống GV dạy thay
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
      return { success: true, message: `Đã tạo thành công ${rowsToAdd.length} buổi học.` };
    } else {
      return { success: true, message: 'Không có buổi học mới nào được tạo.' };
    }
  } catch (error) {
    Logger.log("!!! LỖI trong hàm createRecurringSessions: " + error.toString());
    return { success: false, error: error.toString() }; // Trả về lỗi trùng lịch cho người dùng
  } finally {
    lock.releaseLock();
  }
}

// Thay thế hàm deleteMultipleSessions
function deleteMultipleSessions(sessionIds) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!sessionIds || sessionIds.length === 0) return { success: true, message: "Không có buổi học nào được chọn để xóa." };
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const attendanceSheet = getSheet(CONFIG.ATTENDANCE_SHEET);
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const sessionsWithAttendance = new Set(attendanceData.map(row => row[1]));
    const idsToDelete = [];
    const idsSkipped = [];
    sessionIds.forEach(id => {
      if (sessionsWithAttendance.has(id)) {
        idsSkipped.push(id);
      } else {
        idsToDelete.push(id);
      }
    });
    if (idsToDelete.length > 0) {
      const allSessionsData = sessionsSheet.getDataRange().getValues();
      const rowsToDelete = [];
      allSessionsData.forEach((row, index) => {
        if (idsToDelete.includes(row[0])) {
          rowsToDelete.push(index + 1);
        }
      });
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sessionsSheet.deleteRow(rowsToDelete[i]);
      }
    }
    let message = '';
    if (idsToDelete.length > 0) message += `Đã xóa thành công ${idsToDelete.length} buổi học. `;
    if (idsSkipped.length > 0) message += `${idsSkipped.length} buổi học không thể xóa do đã có dữ liệu điểm danh.`;
    if (message === '') message = 'Không có buổi học nào được xóa.';
    return { success: true, message: message.trim() };
  } catch (e) {
    Logger.log("!!! LỖI trong hàm deleteMultipleSessions: " + e.toString());
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// STUDENT MANAGEMENT
// ===============================================================
// Thay the ham searchStudents bang ham getStudentsPage nay
function getStudentsPage(options) {
  try {    
    const searchTerm = options.searchTerm || '';
    const page = options.page || 1;
    const pageSize = options.pageSize || 50; 

    const sheet = getSheet(CONFIG.STUDENTS_SHEET);
    const allStudentsData = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues() : [];
    
    // *** THAY DOI O DAY: Loc hoc vien "Dang hoc" truoc khi tim kiem ***
    const activeStudentsData = allStudentsData.filter(row => row[5] === 'Đang học');

    let filteredData = activeStudentsData;
    if (searchTerm && searchTerm.trim() !== '') {
      const lowerCaseSearchTerm = searchTerm.toLowerCase();
      filteredData = activeStudentsData.filter(row => {
        const studentId = row[0] ? row[0].toString().toLowerCase() : '';
        const studentName = row[1] ? row[1].toString().toLowerCase() : '';
        return studentId.includes(lowerCaseSearchTerm) || studentName.includes(lowerCaseSearchTerm);
      });
    }
    
    const totalItems = filteredData.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = startIndex + pageSize;

    const studentsOnPageData = filteredData.slice(startIndex, endIndex);
    
    const students = studentsOnPageData.map(row => ({
      id: row[0], name: row[1],
      dob: row[2] ? Utilities.formatDate(new Date(row[2]), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd') : '',
      parentName: row[3], phone: row[4]
    }));

    return {
      success: true,
      data: {
        studentsOnPage: students,
        pagination: {
          currentPage: page,
          totalPages: totalPages,
          pageSize: pageSize,
          totalItems: totalItems
        }
      }
    };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm getStudentsPage: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }
}

function addStudent(data) {
  const lock = getLock();
  try {
    validateFormData(data);
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);
    if (!data.name || !data.dob) throw new Error('Họ tên và Ngày sinh là bắt buộc.');
    
    const newStudentId = generateNextId(studentsSheet, "A", CONFIG.STUDENT_ID_PREFIX);
    
    // *** THAY DOI O DAY: Them trang thai "Dang hoc" vao cuoi ***
    studentsSheet.appendRow([newStudentId, data.name, data.dob, data.parentName || '', data.phone || '', 'Đang học']);

    if (data.enrollments && Array.isArray(data.enrollments) && data.enrollments.length > 0) {
      const enrollmentDate = new Date();
      data.enrollments.forEach(enrollment => {
        if (enrollment.classId) {
          const newEnrollmentId = generateNextId(enrollmentsSheet, "A", CONFIG.ENROLLMENT_ID_PREFIX);
          enrollmentsSheet.appendRow([newEnrollmentId, newStudentId, enrollment.classId, enrollmentDate, enrollment.feeTypeId || '']);
        }
      });
    }

    return { success: true, message: `Đã thêm học viên ${data.name} thành công.` };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm addStudent: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
        lock.releaseLock();
  }
}

// Thay thế hàm editStudent
function editStudent(data) {
  const lock = getLock(); // <-- Thêm dòng này để khóa tiến trình
  try {
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);

    const allStudents = studentsSheet.getDataRange().getValues();
    const rowIndex = allStudents.findIndex(row => row[0] === data.id);
    if (rowIndex === -1) throw new Error('Không tìm thấy học viên để cập nhật.');

    studentsSheet.getRange(rowIndex + 1, 2, 1, 4).setValues([[data.name, data.dob, data.parentName || '', data.phone || '']]);

    const allEnrollments = enrollmentsSheet.getDataRange().getValues();
    const rowsToDelete = [];
    allEnrollments.forEach((row, index) => {
      if (row[1] === data.id) {
        rowsToDelete.push(index + 1);
      }
    });

    // Xóa các ghi danh cũ trước khi thêm mới
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      enrollmentsSheet.deleteRow(rowsToDelete[i]);
    }
    
    // Thêm các ghi danh mới
    if (data.enrollments && Array.isArray(data.enrollments) && data.enrollments.length > 0) {
      const enrollmentDate = new Date();
      data.enrollments.forEach(enrollment => {
        if(enrollment.classId) {
          const newEnrollmentId = generateNextId(enrollmentsSheet, "A", CONFIG.ENROLLMENT_ID_PREFIX);
          enrollmentsSheet.appendRow([newEnrollmentId, data.id, enrollment.classId, enrollmentDate, enrollment.feeTypeId || '']);
        }
      });
    }

    return { success: true, message: 'Cập nhật thông tin học viên thành công.' };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm editStudent: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error);
    return { success: false, error: analysis };
  } finally {
    lock.releaseLock(); // <-- Thêm dòng này để mở khóa tiến trình
  }
}

// Thay thế hàm deleteStudent
function deactivateStudent(studentId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sheet = getSheet(CONFIG.STUDENTS_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === studentId);
    
    if (rowIndex === -1) {
      throw new Error("Không tìm thấy học viên.");
    }
    
    // Cap nhat cot F (cot thu 6) thanh "Đã nghỉ"
    sheet.getRange(rowIndex + 1, 6).setValue('Đã nghỉ');
    
    return { success: true, message: "Đã cập nhật trạng thái học viên thành 'Đã nghỉ'." };
  } catch (error) {
    Logger.log("!!! LỖI trong hàm deactivateStudent: " + error.toString());
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// ATTENDANCE MANAGEMENT
// ===============================================================
// Thay thế hàm getAttendanceForSession
function getAttendanceForSession(sessionId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!sessionId) return { success: true, data: [] };
    const sheet = getSheet(CONFIG.ATTENDANCE_SHEET);
    const data = sheet.getDataRange().getValues();
    const attendanceRecords = data.map(row => ({
        sessionId: row[1], studentId: row[2], status: row[3], note: row[4]
      })).filter(record => record.sessionId === sessionId);
    return { success: true, data: attendanceRecords };
  } catch (e) {
    Logger.log("!!! LỖI trong hàm getAttendanceForSession: " + e.toString());
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function saveAttendanceForSession(sessionId, attendanceData) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sheet = getSheet(CONFIG.ATTENDANCE_SHEET);
    const allData = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
    attendanceData.forEach(studentAtt => {
      const rowIndex = allData.findIndex(row => row[1] === sessionId && row[2] === studentAtt.studentId);
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 2, 4, 1, 2).setValues([[ studentAtt.status, studentAtt.note ]]);
      } else {
        const newAttendanceId = generateNextId(sheet, "A", CONFIG.ATTENDANCE_ID_PREFIX);
        sheet.appendRow([ newAttendanceId, sessionId, studentAtt.studentId, studentAtt.status, studentAtt.note ]);
      }
    });
    return { success: true, message: "Đã lưu điểm danh thành công!" };
  } catch (e) {
    Logger.log("!!! LỖI trong hàm saveAttendanceForSession: " + e.toString());
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// FINANCIAL MANAGEMENT
// ===============================================================
// Thay thế hàm addTransaction
function addTransaction(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { studentId, enrollmentId, amount, date, content, collector, method } = data;
    if(!studentId || !enrollmentId || !amount || !date) {
      throw new Error("Thông tin giao dịch không đầy đủ.");
    }
    const sheet = getSheet(CONFIG.TRANSACTIONS_SHEET);
    const newId = generateNextId(sheet, "A", CONFIG.TRANSACTION_ID_PREFIX);
    const paymentDate = new Date(date);

    sheet.appendRow([newId, studentId, enrollmentId, amount, paymentDate, content || '', collector || '', method || '']);
    return { success: true, message: "Đã ghi nhận giao dịch thành công."};
  } catch (e) {
    Logger.log("!!! LỖI trong hàm addTransaction: " + e.toString());
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// TEACHER MANAGEMENT
// ===============================================================
function addTeacher(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sheet = getSheet(CONFIG.TEACHERS_SHEET);
    const newId = generateNextId(sheet, "A", 'GV');
    sheet.appendRow([
      newId, data.name, data.phone || '', data.email || '', data.dob || '', 
      data.startDate || '', data.specialization || '', data.status || 'Active',
      data.payRate || 0 // Them payRate vao
    ]);
    return { success: true, message: `Đã thêm giáo viên ${data.name}` };
  } catch (error) {
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function editTeacher(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sheet = getSheet(CONFIG.TEACHERS_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === data.id);
    if (rowIndex === -1) throw new Error("Không tìm thấy giáo viên.");

    sheet.getRange(rowIndex + 1, 2, 1, 8).setValues([[ // Tang so cot len 8
      data.name, data.phone || '', data.email || '', data.dob || '',
      data.startDate || '', data.specialization || '', data.status || '',
      data.payRate || 0 // Them payRate vao
    ]]);
    return { success: true, message: "Cập nhật thông tin giáo viên thành công." };
  } catch (error) {
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function deleteTeacher(teacherId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    // Kiem tra xem giao vien co dang phu trach lop nao khong
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const classData = classesSheet.getDataRange().getValues();
    const isTeachingClass = classData.some(row => row[2] === teacherId); // Kiem tra cot C - teacherId
    if (isTeachingClass) {
      throw new Error("Không thể xóa. Giáo viên này đang là giáo viên chính của một hoặc nhiều lớp học.");
    }
    
    // Kiem tra xem giao vien co dang duoc phan cong buoi hoc le nao khong
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const sessionData = sessionsSheet.getDataRange().getValues();
    const isTeachingSession = sessionData.some(row => row[5] === teacherId); // Kiem tra cot F - teacherId
    if (isTeachingSession) {
      throw new Error("Không thể xóa. Giáo viên này đang được phân công dạy một hoặc nhiều buổi học lẻ.");
    }
    
    // Neu khong co rang buoc, tien hanh xoa
    const sheet = getSheet(CONFIG.TEACHERS_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === teacherId);
    if (rowIndex === -1) throw new Error("Không tìm thấy giáo viên.");
    
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: "Đã xóa giáo viên." };
  } catch (error) {
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function getTeacherPayroll(options) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { teacherId, startDate, endDate } = options;
    if (!teacherId || !startDate || !endDate) {
      throw new Error("Vui lòng cung cấp đủ thông tin.");
    }

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23, 59, 59, 999);

    const teachersSheet = getSheet(CONFIG.TEACHERS_SHEET);
    const teacherData = teachersSheet.getDataRange().getValues();
    const teacherRow = teacherData.find(row => row[0] === teacherId);
    if (!teacherRow) throw new Error("Không tìm thấy giáo viên.");
    
    const teacherName = teacherRow[1];
    const payRate = Number(teacherRow[8] || 0);

    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const allSessions = sessionsSheet.getDataRange().getValues().slice(1);
    const classMap = new Map(getSheet(CONFIG.CLASSES_SHEET).getDataRange().getValues().slice(1).map(c => [c[0], {name: c[1], mainTeacherId: c[2]}])); // Doc teacherId tu cot C

    const taughtSessions = allSessions.filter(s => {
      const sessionDate = new Date(s[2]);
      const sessionSpecificTeacherId = s[5]; // teacherId tu buoi hoc
      const classInfo = classMap.get(s[1]);
      const classMainTeacherId = classInfo ? classInfo.mainTeacherId : '';
      
      // So sanh theo teacherId
      const isTaughtByTeacher = (sessionSpecificTeacherId && sessionSpecificTeacherId === teacherId) || (!sessionSpecificTeacherId && classMainTeacherId === teacherId);
      
      return isTaughtByTeacher && sessionDate >= startDateObj && sessionDate <= endDateObj;
    });

    let totalHoursDecimal = 0;
    const sessionDetails = taughtSessions.map(s => {
      const startTime = new Date(s[3]);
      const endTime = new Date(s[4]);
      if (isNaN(startTime) || isNaN(endTime) || endTime <= startTime) return null;
      const durationInMinutes = (endTime.getTime() - startTime.getTime()) / (1000 * 60);
      totalHoursDecimal += durationInMinutes / 60;
      const hours = Math.floor(durationInMinutes / 60);
      const minutes = durationInMinutes % 60;
      const classInfo = classMap.get(s[1]);
      return {
        date: Utilities.formatDate(new Date(s[2]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy'),
        className: classInfo ? classInfo.name : 'N/A',
        startTime: Utilities.formatDate(startTime, CONFIG.DEFAULT_TIMEZONE, 'HH:mm'),
        endTime: Utilities.formatDate(endTime, CONFIG.DEFAULT_TIMEZONE, 'HH:mm'),
        durationFormatted: `${hours} giờ ${minutes} phút`
      };
    }).filter(Boolean);
    
    const totalSalary = totalHoursDecimal * payRate;
    const totalHoursInt = Math.floor(totalHoursDecimal);
    const totalMinutesInt = Math.round((totalHoursDecimal - totalHoursInt) * 60);
    const totalHoursFormatted = `${totalHoursInt} giờ ${totalMinutesInt} phút`;

    return {
      success: true,
      data: {
        teacherName: teacherName,
        payRate: payRate,
        totalHoursFormatted: totalHoursFormatted,
        totalSalary: totalSalary,
        details: sessionDetails.sort((a,b) => new Date(a.date.split('/').reverse().join('-')) - new Date(b.date.split('/').reverse().join('-')))
      }
    };
  } catch (error) {
    Logger.log(`Lỗi tính lương: ${error.toString()}`);
    const analysis = getGeminiErrorAnalysis(error); // Goi Gemini de phan tich
    return { success: false, error: analysis };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// EXPENSE MANAGEMENT
// ===============================================================
function getExpenses() {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const sheet = getSheet(CONFIG.EXPENSES_SHEET);
    const data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues() : [];
    
    const expenses = data.map(row => ({
      id: row[0],
      date: Utilities.formatDate(new Date(row[1]), CONFIG.DEFAULT_TIMEZONE, 'dd/MM/yyyy'),
      amount: Number(row[2] || 0),
      category: row[3],
      description: row[4],
      payer: row[5]
    })).sort((a,b) => new Date(b.date.split('/').reverse().join('-')) - new Date(a.date.split('/').reverse().join('-')));
    
    return { success: true, data: expenses };
  } catch(e) {
    Logger.log(`Lỗi khi lấy dữ liệu chi phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function addExpense(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { date, amount, category, description, payer } = data;
    if (!date || !amount || !category) {
      throw new Error("Ngày, Số tiền và Loại chi phí là bắt buộc.");
    }

    const sheet = getSheet(CONFIG.EXPENSES_SHEET);
    const newId = generateNextId(sheet, "A", 'CP'); // CP = Chi Phi
    
    sheet.appendRow([newId, new Date(date), amount, category, description || '', payer || '']);
    return { success: true, message: "Đã thêm khoản chi thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi thêm chi phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function editExpense(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { id, date, amount, category, description, payer } = data;
    if (!id || !date || !amount || !category) {
      throw new Error("Thông tin chi phí không đầy đủ.");
    }
    const sheet = getSheet(CONFIG.EXPENSES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === id);

    if (rowIndex === -1) {
      throw new Error("Không tìm thấy khoản chi để cập nhật.");
    }

    sheet.getRange(rowIndex + 1, 2, 1, 5).setValues([[
      new Date(date), amount, category, description || '', payer || ''
    ]]);
    
    return { success: true, message: "Đã cập nhật khoản chi thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi sửa chi phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function deleteExpense(expenseId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!expenseId) {
      throw new Error("Không có mã chi phí.");
    }
    const sheet = getSheet(CONFIG.EXPENSES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === expenseId);
    
    if (rowIndex === -1) {
      throw new Error("Không tìm thấy khoản chi để xóa.");
    }
    
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: "Đã xóa khoản chi thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi xóa chi phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// SETTINGS MANAGEMENT
// ===============================================================
function addConfigItem(listName, newItem) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!listName || !newItem || newItem.trim() === '') {
      throw new Error("Dữ liệu không hợp lệ.");
    }
    const sheet = getSheet(CONFIG.CONFIG_SHEET);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(listName);
    
    if (colIndex === -1) {
      throw new Error(`Không tìm thấy cột cấu hình: ${listName}`);
    }
    
    const colNumber = colIndex + 1; // Chuyển index sang số thứ tự cột
    
    // Lấy toàn bộ giá trị trong cột đó
    const colValues = sheet.getRange(1, colNumber, sheet.getMaxRows(), 1).getValues();
    
    // Tìm hàng trống đầu tiên trong cột
    let nextEmptyRow = 0;
    for (let i = 0; i < colValues.length; i++) {
      if (colValues[i][0] === "") {
        nextEmptyRow = i + 1; // +1 vì hàng được đánh số từ 1
        break;
      }
    }

    // Nếu không tìm thấy hàng trống (cột đã đầy), thêm vào cuối
    if (nextEmptyRow === 0) {
      nextEmptyRow = sheet.getLastRow() + 1;
    }

    sheet.getRange(nextEmptyRow, colNumber).setValue(newItem.trim());

    return { success: true, message: `Đã thêm "${newItem}" vào danh sách.` };
  } catch(e) {
    Logger.log(`Lỗi khi thêm mục cài đặt: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// File: Code.js

// THAY THE TOAN BO HAM CU bang phien ban co kiem tra rang buoc
function deleteConfigItem(listName, itemToDelete) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!listName || !itemToDelete) {
      throw new Error("Dữ liệu không hợp lệ.");
    }

    // *** LOGIC KIEM TRA RANG BUOC MOI DUOC THEM VAO ***
    const usageCheckMap = {
      // Ten trong Config -> [Ten Sheet, So Thu Tu Cot De Kiem Tra]
      'TenKhoaHoc': [CONFIG.CLASSES_SHEET, 2],       // Cot B - name
      'LoaiChiPhi': [CONFIG.EXPENSES_SHEET, 4],     // Cot D - category
      'NguoiChi': [CONFIG.EXPENSES_SHEET, 6],         // Cot F - payer
      'NguoiThu': [CONFIG.TRANSACTIONS_SHEET, 7], // Cot G - collector
      'HinhThucTT': [CONFIG.TRANSACTIONS_SHEET, 8] // Cot H - method
    };

    if (usageCheckMap[listName]) {
      const [sheetName, colIndex] = usageCheckMap[listName];
      const sheet = getSheet(sheetName);
      if (sheet.getLastRow() > 1) {
          const data = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getValues().flat();
          const isBeingUsed = data.some(cell => cell.toString().trim() === itemToDelete.toString().trim());
          if (isBeingUsed) {
            throw new Error(`Không thể xóa. Mục "${itemToDelete}" đang được sử dụng ở nơi khác.`);
          }
      }
    }
    // *** KET THUC LOGIC KIEM TRA ***

    const sheet = getSheet(CONFIG.CONFIG_SHEET);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.indexOf(listName);
    
    if (columnIndex === -1) {
      throw new Error(`Không tìm thấy cột cấu hình: ${listName}`);
    }

    const columnValues = sheet.getRange(2, columnIndex + 1, sheet.getLastRow()).getValues().flat();
    const itemIndex = columnValues.findIndex(item => item.toString() === itemToDelete.toString());

    if (itemIndex === -1) {
      throw new Error(`Không tìm thấy mục "${itemToDelete}" để xóa.`);
    }

    sheet.getRange(itemIndex + 2, columnIndex + 1).clearContent();
    const range = sheet.getRange(1, columnIndex + 1, sheet.getLastRow());
    range.sort({column: columnIndex + 1, ascending: true});

    return { success: true, message: `Đã xóa "${itemToDelete}" khỏi danh sách.` };
  } catch(e) {
    Logger.log(`Lỗi khi xóa mục cài đặt: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// FEE TYPE MANAGEMENT
// ===============================================================
function addFeeType(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { name, amount, note } = data;
    if (!name || !amount) {
      throw new Error("Tên và Số tiền là bắt buộc.");
    }
    const sheet = getSheet(CONFIG.FEE_TYPES_SHEET);
    const newId = generateNextId(sheet, "A", 'HP'); // HP = Hoc Phi
    
    sheet.appendRow([newId, name, amount, note || '']);
    return { success: true, message: "Đã thêm loại học phí thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi thêm loại học phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function editFeeType(data) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    const { id, name, amount, note } = data;
    if (!id || !name || !amount) {
      throw new Error("Thông tin không đầy đủ.");
    }
    const sheet = getSheet(CONFIG.FEE_TYPES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === id);

    if (rowIndex === -1) {
      throw new Error("Không tìm thấy loại học phí để cập nhật.");
    }

    sheet.getRange(rowIndex + 1, 2, 1, 3).setValues([[
      name, amount, note || ''
    ]]);
    
    return { success: true, message: "Đã cập nhật loại học phí thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi sửa loại học phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

function deleteFeeType(feeTypeId) {
  const lock = getLock(); // <-- Thêm dòng này
  try {
    if (!feeTypeId) throw new Error("Không có mã loại học phí.");

    // Kiem tra xem loai phi nay co dang duoc su dung khong
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);
    const enrollmentData = enrollmentsSheet.getDataRange().getValues();
    const isBeingUsed = enrollmentData.some(row => row[4] === feeTypeId);

    if (isBeingUsed) {
      throw new Error("Không thể xóa. Loại học phí này đang được sử dụng cho học viên.");
    }

    const sheet = getSheet(CONFIG.FEE_TYPES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === feeTypeId);
    
    if (rowIndex === -1) throw new Error("Không tìm thấy loại học phí để xóa.");
    
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: "Đã xóa loại học phí thành công." };
  } catch(e) {
    Logger.log(`Lỗi khi xóa loại học phí: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }finally {
    lock.releaseLock(); // <-- Thêm dòng này
  }
}

// ===============================================================
// COURSE MANAGEMENT
// ===============================================================

/**
 * Thêm một khóa học mới vào sheet Courses.
 */
function addCourse(data) {
  const lock = getLock();
  try {
    const { courseId, programName, level, description, fee } = data;
    if (!courseId || !programName) {
      throw new Error("Mã Khóa Học và Tên Chương Trình là bắt buộc.");
    }

    const sheet = getSheet(CONFIG.COURSES_SHEET);
    const allIds = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    if (allIds.includes(courseId)) {
      throw new Error(`Mã Khóa Học "${courseId}" đã tồn tại. Vui lòng chọn một mã khác.`);
    }

    sheet.appendRow([courseId, programName, level || '', description || '', fee || 0]);
    return { success: true, message: `Đã thêm khóa học "${courseId}" thành công.` };
  } catch (e) {
    Logger.log(`Lỗi khi thêm Khóa học: ${e.toString()}`);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cập nhật thông tin một khóa học đã có.
 */
function editCourse(data) {
  const lock = getLock();
  try {
    const { courseId, programName, level, description, fee } = data;
    if (!courseId) throw new Error("Không tìm thấy Mã Khóa Học.");

    const sheet = getSheet(CONFIG.COURSES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === courseId);

    if (rowIndex === -1) {
      throw new Error(`Không tìm thấy khóa học với mã "${courseId}" để cập nhật.`);
    }

    // Cập nhật từ cột B đến cột E (4 cột)
    sheet.getRange(rowIndex + 1, 2, 1, 4).setValues([[
      programName, level || '', description || '', fee || 0
    ]]);
    
    return { success: true, message: `Đã cập nhật khóa học "${courseId}".` };
  } catch (e) {
    Logger.log(`Lỗi khi sửa Khóa học: ${e.toString()}`);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Xóa một khóa học, nhưng phải kiểm tra xem nó có đang được lớp nào sử dụng không.
 */
function deleteCourse(courseId) {
  const lock = getLock();
  try {
    if (!courseId) throw new Error("Không có Mã Khóa Học để xóa.");

    // Kiểm tra ràng buộc: Khóa học có đang được lớp nào dùng không?
    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    // Cột B trong sheet Classes bây giờ là CourseID
    const allClassCourseIds = classesSheet.getRange(2, 2, classesSheet.getLastRow(), 1).getValues().flat(); 
    const isBeingUsed = allClassCourseIds.includes(courseId);

    if (isBeingUsed) {
      throw new Error(`Không thể xóa. Khóa học "${courseId}" đang được sử dụng bởi một hoặc nhiều lớp học.`);
    }

    const sheet = getSheet(CONFIG.COURSES_SHEET);
    const allData = sheet.getDataRange().getValues();
    const rowIndex = allData.findIndex(row => row[0] === courseId);
    
    if (rowIndex === -1) {
      throw new Error(`Không tìm thấy khóa học "${courseId}".`);
    }
    
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: `Đã xóa khóa học "${courseId}".` };
  } catch (e) {
    Logger.log(`Lỗi khi xóa Khóa học: ${e.toString()}`);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
// ===============================================================
// LAZY LOADING DATA FUNCTIONS
// ===============================================================
function getCoreData() {
  try {   
    // === PHẦN MỚI: Tải dữ liệu từ sheet Courses ===
    const coursesSheet = getSheet(CONFIG.COURSES_SHEET);
    const coursesData = coursesSheet.getLastRow() > 1 ? coursesSheet.getRange(2, 1, coursesSheet.getLastRow() - 1, 5).getValues() : [];
    const courses = coursesData.map(c => ({
      id: c[0],           // CourseID (VD: IE-6.5)
      programName: c[1],
      level: c[2],
      description: c[3],
      fee: c[4]
    })).filter(c => c && c.id); // Lọc những hàng có CourseID
    // === KẾT THÚC PHẦN MỚI ===

    const classesSheet = getSheet(CONFIG.CLASSES_SHEET);
    const teachersSheet = getSheet(CONFIG.TEACHERS_SHEET);
    const feeTypesSheet = getSheet(CONFIG.FEE_TYPES_SHEET);
    const studentsSheet = getSheet(CONFIG.STUDENTS_SHEET);
    const enrollmentsSheet = getSheet(CONFIG.ENROLLMENTS_SHEET);

    const classesData = classesSheet.getLastRow() > 1 ? classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, 6).getValues() : [];
    const teachersData = teachersSheet.getLastRow() > 1 ? teachersSheet.getRange(2, 1, teachersSheet.getLastRow() - 1, 9).getValues() : [];
    const feeTypesData = feeTypesSheet.getLastRow() > 1 ? feeTypesSheet.getRange(2, 1, feeTypesSheet.getLastRow() - 1, 4).getValues() : [];
    
    const allStudentsData = studentsSheet.getLastRow() > 1 ? studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 6).getValues() : [];
    const activeStudentsData = allStudentsData.filter(row => row[5] === 'Đang học');
    
    const enrollmentsData = enrollmentsSheet.getLastRow() > 1 ? enrollmentsSheet.getRange(2, 1, enrollmentsSheet.getLastRow() - 1, 5).getValues() : [];
    
    const teacherMap = new Map(teachersData.map(t => [t[0], t[1]]));

    const classes = classesData.map(r => {
      const teacherName = teacherMap.get(r[2]) || '';
      return { id:r[0], courseId:r[1], teacherId:r[2] || '', teacherName: teacherName, maxSize:r[3]||0, lichHoc:r[4]||'', gioHoc:r[5]||''};
    }).filter(c => c.id);

    const teachers = teachersData.map(r => ({id:r[0], name:r[1], phone:r[2], email:r[3], dob:r[4]?'...':'', startDate:r[5]?'...':'', specialization:r[6], status:r[7], payRate:r[8]||0})).filter(t => t.id);
    const feeTypes = feeTypesData.map(r => ({id:r[0], name:r[1], amount:r[2], note:r[3]})).filter(f => f.id);
    
    // *** DONG CODE DUOC SUA O DAY ***
    // Lay day du thong tin hoc vien, khong chi ten va ID
    const students = activeStudentsData.map(r => ({
        id: r[0], 
        name: r[1],
        dob: r[2] ? Utilities.formatDate(new Date(r[2]), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd') : '',
        parentName: r[3],
        phone: r[4]
    })).filter(s => s.id);

    const enrollments = enrollmentsData.map(r => ({id:r[0], studentId:r[1], classId:r[2], enrollmentDate:r[3]?'...':'', feeTypeId:r[4]||''})).filter(e => e.id);

    return { success: true, data: { courses, classes, teachers, feeTypes, students, enrollments } };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSessionsData() {
  try {    
    const sessionsSheet = getSheet(CONFIG.CLASS_SESSIONS_SHEET);
    const sessionsData = sessionsSheet.getLastRow() > 1 ? sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 6).getValues() : [];
    
    // Can lay ca du lieu giao vien de tra cuu ten
    const teachersSheet = getSheet(CONFIG.TEACHERS_SHEET);
    const teachersData = teachersSheet.getLastRow() > 1 ? teachersSheet.getRange(2, 1, teachersSheet.getLastRow() - 1, 2).getValues() : [];
    const teacherMap = new Map(teachersData.map(t => [t[0], t[1]])); // Map tu ID -> Ten

    const sessions = sessionsData.map(r => {
      // Tu teacherId (r[5]), tim ra ten giao vien
      const teacherName = teacherMap.get(r[5]) || '';
      return {
        id:r[0], classId:r[1], 
        date:r[2]?Utilities.formatDate(new Date(r[2]),CONFIG.DEFAULT_TIMEZONE,'yyyy-MM-dd'):'', 
        startTime:r[3]?Utilities.formatDate(new Date(r[3]),CONFIG.DEFAULT_TIMEZONE,'HH:mm'):'', 
        endTime:r[4]?Utilities.formatDate(new Date(r[4]),CONFIG.DEFAULT_TIMEZONE,'HH:mm'):'', 
        teacherId:r[5], // Luu lai ca ID
        teacher:teacherName // Va ten da tra cuu duoc
      }
    }).filter(s => s.id);

    return { success: true, data: { sessions } };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getGeminiAnalysis() {
  try {
    // Lay du lieu tu cac sheet
    const classes = getSheet(CONFIG.CLASSES_SHEET).getDataRange().getValues().slice(1);
    const students = getSheet(CONFIG.STUDENTS_SHEET).getDataRange().getValues().slice(1);
    const transactions = getSheet(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues().slice(1);
    const attendance = getSheet(CONFIG.ATTENDANCE_SHEET).getDataRange().getValues().slice(1);

    // Chuan bi du lieu de gui cho Gemini
    const dataSummary = {
      totalClasses: classes.length,
      totalActiveStudents: students.filter(s => s[5] === 'Đang học').length,
      totalTransactions: transactions.length,
      totalAttendanceRecords: attendance.length,
      // Ban co the them du lieu chi tiet hon o day
    };

    // Tao cau lenh (prompt) cho Gemini
    const prompt = `Bạn là một chuyên gia phân tích kinh doanh cho một trung tâm gia sư. Dựa vào các số liệu tóm tắt sau: ${JSON.stringify(dataSummary)}, hãy đưa ra 3 nhận xét về tình hình vận hành của trung tâm và 2 đề xuất để cải thiện. Trình bày dưới dạng gạch đầu dòng, ngôn ngữ thân thiện.`;

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      throw new Error("Chưa cấu hình Gemini API Key.");
    }
    
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
    
    const payload = {
      "contents": [{
        "parts": [{ "text": prompt }]
      }]
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    const analysisText = result.candidates[0].content.parts[0].text;
    
    return { success: true, data: analysisText };
  } catch(e) {
    Logger.log("Loi khi goi Gemini API: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getGeminiErrorAnalysis(errorObject) {
  try {
    // Ham nay se khong tu throw error de tranh vong lap loi
    const errorMessage = errorObject.toString() + (errorObject.stack ? ('\nStack Trace: ' + errorObject.stack) : '');
    const prompt = `Bạn là một lập trình viên Google Apps Script chuyên nghiệp. Một lỗi vừa xảy ra trong hệ thống web của tôi. Dựa vào thông tin lỗi sau đây: "${errorMessage}", hãy giải thích nguyên nhân có thể gây ra lỗi này bằng ngôn ngữ đơn giản, dễ hiểu cho người không chuyên về kỹ thuật và đề xuất hướng khắc phục nếu có.`;
    
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "Lỗi: Chưa cấu hình Gemini API Key.";
    
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
    const payload = {"contents": [{"parts": [{ "text": prompt }]}]};
    const options = {'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload)};
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
  } catch(e) {
    // Neu chinh ham nay bi loi, tra ve thong bao loi goc
    return "Không thể kết nối đến Gemini để phân tích lỗi. Lỗi gốc là: " + errorObject.toString();
  }
}

function generateNewClassId(courseId, startDate) {
  const sheet = getSheet(CONFIG.CLASSES_SHEET);
  const classData = sheet.getDataRange().getValues();

  // 1. Tạo phần tiền tố từ CourseID và ngày khai giảng
  const year = startDate.getFullYear().toString().slice(-2); // 25
  const month = (startDate.getMonth() + 1).toString().padStart(2, '0'); // 10
  const prefix = `${courseId}-K${year}${month}`; // VD: IE-6.5-K2510

  // 2. Tìm số thứ tự lớn nhất của các lớp có cùng tiền tố
  let maxSequence = 0;
  classData.forEach(row => {
    const existingId = row[0]; // Giả sử mã lớp ở cột A
    if (existingId && existingId.startsWith(prefix)) {
      const parts = existingId.split('-');
      const sequence = parseInt(parts[parts.length - 1], 10);
      if (sequence > maxSequence) {
        maxSequence = sequence;
      }
    }
  });

  // 3. Tạo mã hoàn chỉnh với số thứ tự tiếp theo
  const nextSequence = (maxSequence + 1).toString().padStart(2, '0');
  return `${prefix}-${nextSequence}`;
}

function kiemTraDuLieuPhanQuyen() {
  try {
    // Cố gắng truy cập trực tiếp vào spreadsheet và sheet 'PhanQuyen'
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PhanQuyen');
    
    if (!sheet) {
      Logger.log('LỖI NGHIÊM TRỌNG: Không tìm thấy sheet "PhanQuyen" trong file Google Sheet này. Tên file Sheet hiện tại là: ' + ss.getName());
      return;
    }
    
    // Nếu tìm thấy sheet, đọc dữ liệu và ghi ra Log
    const data = sheet.getDataRange().getValues();
    Logger.log('Đã tìm thấy sheet "PhanQuyen". Dữ liệu đọc được như sau:');
    Logger.log(JSON.stringify(data));

  } catch (e) {
    Logger.log('Đã xảy ra lỗi nghiêm trọng khi cố gắng đọc sheet: ' + e.toString());
  }
}
function forceReAuthorization() {
  try {
    // Hành động này yêu cầu quyền được biết email người dùng
    const email = Session.getActiveUser().getEmail();
    Logger.log('Email của người thực thi: ' + email);

    // Hành động này yêu cầu quyền được đọc file Google Sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PhanQuyen');
    Logger.log('Đã truy cập được sheet PhanQuyen.');
    
    Logger.log('SUCCESS: Script đã được cấp đủ quyền cần thiết.');
  } catch (e) {
    Logger.log('ERROR: Có lỗi xảy ra trong quá trình cấp lại quyền. ' + e.toString());
  }
}