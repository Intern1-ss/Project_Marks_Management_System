const SHEET_NAME = "Project_mapping";
const EDIT_REQUESTS_SHEET = "EditRequests";
const FACULTY_DEADLINES_SHEET = "FacultyDeadlines";
const ADMIN_EMAIL = "intern1@sssihl.edu.in";
let WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxbPaYuY_7yittUQXd0XZTSLuEYFpFrDLen4F3XTRtCWqisAaXQs_F3GFZ2dSqTR0eBng/exec"; 
const COL_REGD_NO = "Regd. No.";
const COL_STUDENT_NAME = "Student Name";
const COL_PAPER_CODE = "Paper Code";
const COL_PAPER_TITLE = "Paper Title";
const COL_EXAMINER_EMAIL = "Examiner Email";
const COL_MARKS = "Marks (100)";
const COL_VERIFIED = "Verified";
const ANALYSIS_SHEET_NAME = "PaperAnalysis";
 
// ===== INITIALIZATION FUNCTIONS =====
function onOpen() {
  const context = detectExecutionContext();
  console.log(`Initializing in ${context} context`);
  
  try {
    initialize();
  } catch (error) {
    console.error("Error in sheet initialization:", error);
  }
  
  if (context === 'spreadsheet') {
    createMenus();
  } else {
    console.log("Skipping menu creation - running in web app context");
  }
}

function initialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create main sheet if it doesn't exist
  if (!ss.getSheetByName(SHEET_NAME)) {
    const mainSheet = ss.insertSheet(SHEET_NAME);
    mainSheet.appendRow([
      "Sl. No.", "Programme Name", "Campus", "Semester",
      COL_REGD_NO, COL_STUDENT_NAME, COL_PAPER_CODE, COL_PAPER_TITLE, "Examiner",
      COL_EXAMINER_EMAIL, "Exam", "Credits", "Max Marks", COL_MARKS, COL_VERIFIED
    ]);
    mainSheet.setFrozenRows(1);
  }
  
  // Create EditRequests sheet with correct headers
  let editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
  if (!editSheet) {
    editSheet = ss.insertSheet(EDIT_REQUESTS_SHEET);
  }
  
  // Check if headers are correct - if not, fix them
  const currentRange = editSheet.getDataRange();
  let needsHeaders = true;
  
  if (currentRange.getNumRows() > 0) {
    const currentHeaders = currentRange.getValues()[0];
    const requiredHeaders = ["Request ID", "Faculty Email", "Student Regd No", "Unlock Until"];
    const hasAllHeaders = requiredHeaders.every(header => 
      currentHeaders.some(h => h.toString().trim() === header)
    );
    needsHeaders = !hasAllHeaders;
  }
  
  if (needsHeaders) {
    editSheet.clear();
    editSheet.appendRow([
      "Request ID", "Faculty Email", "Student Regd No", "Student Name", 
      "Paper Code", "Current Marks", "Request Time", "Status", "Action Notes", "Unlock Until"
    ]);
    const headerRange = editSheet.getRange(1, 1, 1, 10);
    headerRange.setBackground("#4CAF50").setFontColor("white").setFontWeight("bold");
    editSheet.setFrozenRows(1);
    editSheet.setColumnWidth(1, 180).setColumnWidth(2, 200).setColumnWidth(3, 120);
    editSheet.setColumnWidth(4, 150).setColumnWidth(5, 200).setColumnWidth(6, 100);
    editSheet.setColumnWidth(7, 150).setColumnWidth(8, 100).setColumnWidth(9, 250).setColumnWidth(10, 150);
    editSheet.insertRowAfter(1);
    const instructionRange = editSheet.getRange(2, 1, 1, 10);
    instructionRange.merge().setValue("Instructions: Select a row containing a 'Pending' request and use 'CoE Functions > Approve/Disapprove Selected Request'. Pending requests are highlighted yellow.")
                    .setBackground("#E3F2FD").setFontStyle("italic").setHorizontalAlignment("center");
  }
  
  formatEditRequestsSheet();
}

function initializeFacultyDeadlines() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let deadlinesSheet = ss.getSheetByName(FACULTY_DEADLINES_SHEET);
    
    if (!deadlinesSheet) {
      deadlinesSheet = ss.insertSheet(FACULTY_DEADLINES_SHEET);
      
      // Create headers
      const headers = [
        "Faculty Email", "Due Date", "Total Students", "Students With Marks", 
        "Students Verified", "Completion Status", "Last Reminder Sent", 
        "Reminder Count", "Completion Confirmed Date", "Notes"
      ];
      
      deadlinesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      deadlinesSheet.getRange(1, 1, 1, headers.length)
        .setBackground("#4CAF50").setFontColor("white").setFontWeight("bold");
      deadlinesSheet.setFrozenRows(1);
      
      console.log("✅ FacultyDeadlines sheet created successfully");
    }
    
    return deadlinesSheet;
  } catch (error) {
    console.error("Error initializing faculty deadlines:", error);
    throw error;
  }
}

function setFacultyDeadline(email, dueDate, notes = "") {
  try {
    const deadlinesSheet = initializeFacultyDeadlines();
    const data = deadlinesSheet.getDataRange().getValues();
    
    // Find if faculty already exists
    let facultyRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        facultyRow = i + 1;
        break;
      }
    }
    
    // Get faculty's student count
    const studentStats = getFacultyStudentStats(email);
    
    // ✅ ENSURE PROPER COLUMN ORDER matching headers
    const rowData = [
      email,                          // Faculty Email
      new Date(dueDate),             // Due Date  
      studentStats.totalStudents,    // Total Students
      studentStats.studentsWithMarks, // Students With Marks
      studentStats.studentsVerified,  // Students Verified
      "Pending",                     // Completion Status
      "",                           // Last Reminder Sent
      0,                            // Reminder Count
      "",                           // Completion Confirmed Date
      notes                         // Notes
    ];
    
    if (facultyRow > 0) {
      deadlinesSheet.getRange(facultyRow, 1, 1, rowData.length).setValues([rowData]);
      console.log(`✅ Updated deadline for ${email}`);
    } else {
      deadlinesSheet.appendRow(rowData);
      console.log(`✅ Added new deadline for ${email}`);
    }
    
    return { success: true, message: `Deadline set for ${email}` };
  } catch (error) {
    console.error("❌ Error setting faculty deadline:", error);
    return { success: false, message: error.message };
  }
}


// Function to get faculty's student statistics
function getFacultyStudentStats(email) {
  try {
    // Validate email parameter
    if (!email || typeof email !== 'string' || email.trim() === '') {
      console.error('getFacultyStudentStats: Invalid email parameter:', email);
      return { totalStudents: 0, studentsWithMarks: 0, studentsVerified: 0, completionPercentage: 0, verificationPercentage: 0 };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.error(`getFacultyStudentStats: Sheet "${SHEET_NAME}" not found`);
      return { totalStudents: 0, studentsWithMarks: 0, studentsVerified: 0, completionPercentage: 0, verificationPercentage: 0 };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log('getFacultyStudentStats: No data in main sheet');
      return { totalStudents: 0, studentsWithMarks: 0, studentsVerified: 0, completionPercentage: 0, verificationPercentage: 0 };
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    let totalStudents = 0;
    let studentsWithMarks = 0;
    let studentsVerified = 0;
    
    const emailLower = email.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      
      if (rowEmail === emailLower) {
        totalStudents++;
        
        const marks = row[columnIndices.marks];
        const verified = row[columnIndices.verified];
        
        if (marks !== null && marks !== undefined && marks !== '') {
          studentsWithMarks++;
        }
        
        if (verified === '✅') {
          studentsVerified++;
        }
      }
    }
    
    const result = {
      totalStudents,
      studentsWithMarks,
      studentsVerified,
      completionPercentage: totalStudents > 0 ? Math.round((studentsWithMarks / totalStudents) * 100) : 0,
      verificationPercentage: totalStudents > 0 ? Math.round((studentsVerified / totalStudents) * 100) : 0
    };
    
    console.log(`📊 Stats for ${email}:`, result);
    return result;
    
  } catch (error) {
    console.error(`getFacultyStudentStats: Error getting stats for ${email}:`, error);
    return { 
      totalStudents: 0, 
      studentsWithMarks: 0, 
      studentsVerified: 0, 
      completionPercentage: 0, 
      verificationPercentage: 0,
      error: error.message 
    };
  }
}


function testDeadlineReminderSafe(email, dueDate, stats, reminderCount) {
  try {
    console.log("🧪 Testing reminder function (no email sent)...");
    
    // Validate inputs (same as real function)
    if (!email || typeof email !== 'string') {
      console.error('Invalid email:', email);
      return false;
    }
    
    if (!dueDate || !(dueDate instanceof Date) || isNaN(dueDate.getTime())) {
      console.error('Invalid due date:', dueDate);
      return false;
    }
    
    if (!stats || typeof stats !== 'object') {
      console.error('Invalid stats:', stats);
      return false;
    }
    
    const today = new Date();
    const daysPastDue = Math.ceil((today - dueDate) / (1000 * 60 * 60 * 24));
    
    const subject = `Urgent: Assessment Deadline Passed - Action Required [Reminder ${reminderCount}]`;
    
    const message = `Dear Faculty Member,

Sai Ram,

This is an urgent reminder regarding your assessment deadline.

📊 DEADLINE STATUS:
- Due Date: ${dueDate.toLocaleDateString()}
- Days Overdue: ${daysPastDue} days
- Reminder Count: ${reminderCount}

📈 YOUR PROGRESS:
- Total Students Assigned: ${stats.totalStudents || 0}
- Marks Uploaded: ${stats.studentsWithMarks || 0}/${stats.totalStudents || 0} (${stats.completionPercentage || 0}%)
- Students Verified: ${stats.studentsVerified || 0}/${stats.totalStudents || 0} (${stats.verificationPercentage || 0}%)

[TEST MODE - EMAIL NOT SENT]`;

    console.log("📧 Test email content:");
    console.log("To:", email);
    console.log("Subject:", subject);
    console.log("Message length:", message.length, "characters");
    
    return true;
    
  } catch (error) {
    console.error("❌ Test reminder error:", error);
    return false;
  }
}

// ===== ENHANCED GET STUDENTS FUNCTION WITH DYNAMIC MAX MARKS =====
function getStudents(email) {
  console.log("🔍 Enhanced getStudents called with email:", email);
  
  if (!email || typeof email !== 'string' || email.trim() === '') {
    console.error("getStudents: Invalid email parameter:", email);
    throw new Error("Invalid email parameter");
  }
  
  const trimmedEmailLower = email.trim().toLowerCase();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log("getStudents: No data found in sheet, returning empty array");
      return [];
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    const facultyStudents = [];
    
    // ✅ ENHANCED: Get unlocked students
    let unlockedRegdNos;
    try {
      unlockedRegdNos = getUnlockedRegdNos(trimmedEmailLower);
      console.log("✅ getUnlockedRegdNos returned successfully, unlocked count:", unlockedRegdNos.size);
    } catch (unlockError) {
      console.error("getStudents: Error in getUnlockedRegdNos:", unlockError);
      unlockedRegdNos = new Set();
    }

    // ✅ ENHANCED: Get pending edit requests with better detection
    let pendingRequests;
    try {
      pendingRequests = getDetailedPendingRequests(trimmedEmailLower);
      console.log("✅ getDetailedPendingRequests returned successfully, pending count:", pendingRequests.size);
    } catch (pendingError) {
      console.error("getStudents: Error getting pending requests:", pendingError);
      pendingRequests = new Set();
    }

    console.log("getStudents: Processing students for email:", trimmedEmailLower);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      
      if (studentEmail === trimmedEmailLower) {
        const regdNo = row[columnIndices.regdNo] ? row[columnIndices.regdNo].toString().trim() : '';
        const verifiedStatus = row[columnIndices.verified] ? row[columnIndices.verified].toString().trim() : '';
        const isVerified = verifiedStatus === '✅';
        const isUnlocked = unlockedRegdNos.has(regdNo);
        const hasPendingRequest = pendingRequests.has(regdNo);
        const canEdit = isUnlocked || !isVerified;

        facultyStudents.push({
          regdNo: regdNo,
          name: row[columnIndices.name] ? row[columnIndices.name].toString().trim() : '',
          paperCode: row[columnIndices.paperCode] ? row[columnIndices.paperCode].toString().trim() : '',
          paperTitle: row[columnIndices.paperTitle] ? row[columnIndices.paperTitle].toString().trim() : '',
          programmeName: row[columnIndices.programme] ? row[columnIndices.programme].toString().trim() : '',
          semester: row[columnIndices.semester] ? row[columnIndices.semester].toString().trim() : '',
          exam: row[columnIndices.exam] ? row[columnIndices.exam].toString().trim() : '',
          credits: row[columnIndices.credits] ? row[columnIndices.credits].toString().trim() : '',
          maxMarks: row[columnIndices.maxMarks] ? Number(row[columnIndices.maxMarks]) || 100 : 100,
          examinerName: row[columnIndices.examiner] ? row[columnIndices.examiner].toString().trim() : '',
          marks: row[columnIndices.marks],
          verified: verifiedStatus,
          isUnlocked: isUnlocked,
          canEdit: canEdit,
          isInEditMode: isUnlocked,
          hasPendingRequest: hasPendingRequest
        });
      }
    }
    
    console.log(`✅ Enhanced getStudents completed: ${facultyStudents.length} students found`);
    console.log(`📊 Stats - Unlocked: ${unlockedRegdNos.size}, Pending: ${pendingRequests.size}`);
    
    return facultyStudents;
    
  } catch (error) {
    console.error(`🚨 CRITICAL ERROR in enhanced getStudents for email ${email}:`, error);
    throw new Error(`Failed to retrieve students: ${error.message}`);
  }
}

// ===== ENHANCED HELPER FUNCTION FOR COLUMN INDICES =====
function getColumnIndices(headers) {
  const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
  const regdNoColIndex = headers.indexOf(COL_REGD_NO);
  const nameColIndex = headers.indexOf(COL_STUDENT_NAME);
  const paperCodeColIndex = headers.indexOf(COL_PAPER_CODE);
  const paperTitleColIndex = headers.indexOf(COL_PAPER_TITLE);
  const marksColIndex = headers.indexOf(COL_MARKS);
  const verifiedColIndex = headers.indexOf(COL_VERIFIED);
  const examinerColIndex = headers.indexOf("Examiner");
  const programmeColIndex = headers.indexOf("Programme Name");
  const semesterColIndex = headers.indexOf("Semester");
  const examColIndex = headers.indexOf("Exam");
  const creditsColIndex = headers.indexOf("Credits");
  const maxMarksColIndex = headers.indexOf("Max Marks");

  const missingCols = [];
  if (emailColIndex === -1) missingCols.push(COL_EXAMINER_EMAIL);
  if (regdNoColIndex === -1) missingCols.push(COL_REGD_NO);
  if (nameColIndex === -1) missingCols.push(COL_STUDENT_NAME);
  if (paperCodeColIndex === -1) missingCols.push(COL_PAPER_CODE);
  if (marksColIndex === -1) missingCols.push(COL_MARKS);
  if (verifiedColIndex === -1) missingCols.push(COL_VERIFIED);
  if (examinerColIndex === -1) missingCols.push("Examiner");

  if (missingCols.length > 0) {
    const errorMsg = `Required columns not found in sheet '${SHEET_NAME}'. Missing: ${missingCols.join(', ')}. Available headers: ${headers.join(', ')}`;
    console.error(errorMsg);
    throw new Error(errorMsg);
  }

  return {
    email: emailColIndex,
    regdNo: regdNoColIndex,
    name: nameColIndex,
    paperCode: paperCodeColIndex,
    paperTitle: paperTitleColIndex,
    marks: marksColIndex,
    verified: verifiedColIndex,
    examiner: examinerColIndex,
    programme: programmeColIndex,
    semester: semesterColIndex,
    exam: examColIndex,
    credits: creditsColIndex,
    maxMarks: maxMarksColIndex
  };
}

// ===== REQUEST EDIT ACCESS WRAPPER FUNCTION =====
function requestEditForSelectedStudents(email, regdNos) {
  console.log("🔄 Enhanced requestEditForSelectedStudents received:", {
    email, 
    regdNosCount: regdNos ? regdNos.length : 0,
    regdNos: regdNos
  });
  
  // ✅ ENHANCED PARAMETER VALIDATION
  if (!email || typeof email !== 'string' || !isValidEmail(email.trim())) {
    console.error("❌ Invalid user email:", email);
    return { 
      success: false, 
      message: "Invalid user email.", 
      processedCount: 0, 
      errors: ["Invalid email format"],
      code: "INVALID_EMAIL"
    };
  }
  
  if (!Array.isArray(regdNos) || regdNos.length === 0) {
    console.error("❌ No registration numbers provided:", regdNos);
    return { 
      success: false, 
      message: "No registration numbers provided.", 
      processedCount: 0, 
      errors: ["No students selected"],
      code: "NO_STUDENTS"
    };
  }
  
  const requestData = {
    email: email.trim(),
    regdNos: regdNos
  };
  
  return enhancedRequestEditAccess(requestData);
}

function enhancedRequestEditAccess(requestData) {
  console.log("🔄 Enhanced requestEditAccess received data:", JSON.stringify(requestData));
  
  if (!requestData || typeof requestData !== 'object') {
    return { 
      success: false, 
      message: "Invalid request data.", 
      processedCount: 0, 
      errors: ["Invalid request format"],
      code: "INVALID_REQUEST"
    };
  }
  
  const { email, regdNos } = requestData;
  
  if (!email || !isValidEmail(email.trim())) {
    return { 
      success: false, 
      message: "Invalid user email.", 
      processedCount: 0, 
      errors: ["Invalid email"],
      code: "INVALID_EMAIL"
    };
  }
  
  if (!Array.isArray(regdNos) || regdNos.length === 0) {
    return { 
      success: false, 
      message: "No registration numbers provided.", 
      processedCount: 0, 
      errors: ["No students provided"],
      code: "NO_STUDENTS"
    };
  }
  
  const trimmedEmail = email.trim().toLowerCase(); // ✅ Normalize email
  let processedCount = 0;
  const errors = [];
  const warnings = [];
  const requestsToAdd = [];
  const duplicateRequests = [];
  const invalidStudents = [];

  try {
    console.log("📊 Starting enhanced request processing...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!mainSheet || !editSheet) {
      throw new Error("Required sheets not found. Please check sheet configuration.");
    }

    const mainData = mainSheet.getDataRange().getValues();
    const mainHeaders = mainData[0].map(h => h.toString().trim());
    const mainColumnIndices = getColumnIndices(mainHeaders);

    const editData = editSheet.getDataRange().getValues();
    const editHeaders = editData[0].map(h => h.toString().trim());
    const editRegdNoCol = editHeaders.indexOf("Student Regd No");
    const editStatusCol = editHeaders.indexOf("Status");
    const editFacultyCol = editHeaders.indexOf("Faculty Email");
    
    // ✅ ENHANCED: Check for existing pending/approved requests
    const existingRequests = new Map(); // regdNo -> {status, requestTime, requestId}
    
    if (editRegdNoCol !== -1 && editStatusCol !== -1 && editFacultyCol !== -1) {
      for(let i = 1; i < editData.length; i++) {
        const rowFaculty = editData[i][editFacultyCol] ? editData[i][editFacultyCol].toString().trim().toLowerCase() : '';
        const rowRegdNo = editData[i][editRegdNoCol] ? editData[i][editRegdNoCol].toString().trim() : '';
        const rowStatus = editData[i][editStatusCol] ? editData[i][editStatusCol].toString().trim() : '';
        const rowRequestTime = editData[i][6] || ''; // Request Time column
        const rowRequestId = editData[i][0] || ''; // Request ID column
        
        if (rowFaculty === trimmedEmail && rowRegdNo) {
          existingRequests.set(rowRegdNo, {
            status: rowStatus,
            requestTime: rowRequestTime,
            requestId: rowRequestId,
            rowIndex: i
          });
          
          console.log(`📋 Found existing request: ${rowRegdNo} -> ${rowStatus}`);
        }
      }
    }

    // ✅ BUILD STUDENT DETAILS MAP
    const studentDetailsMap = {};
    for (let i = 1; i < mainData.length; i++) {
      const row = mainData[i];
      const rowEmail = row[mainColumnIndices.email] ? row[mainColumnIndices.email].toString().trim().toLowerCase() : '';
      
      if (rowEmail === trimmedEmail) {
        const regdNo = row[mainColumnIndices.regdNo] ? row[mainColumnIndices.regdNo].toString().trim() : '';
        if (regdNo) {
          studentDetailsMap[regdNo] = {
            name: row[mainColumnIndices.name] ? row[mainColumnIndices.name].toString().trim() : '',
            paperCode: row[mainColumnIndices.paperCode] ? row[mainColumnIndices.paperCode].toString().trim() : '',
            marks: row[mainColumnIndices.marks],
            verified: row[mainColumnIndices.verified] ? row[mainColumnIndices.verified].toString().trim() : '',
            rowIndex: i + 1 
          };
        }
      }
    }

    console.log(`📊 Found ${Object.keys(studentDetailsMap).length} students for faculty ${trimmedEmail}`);

    // ✅ PROCESS EACH STUDENT WITH ENHANCED VALIDATION
    const timestamp = new Date();
    
    for (const regdNo of regdNos) {
      const trimmedRegdNo = typeof regdNo === 'string' ? regdNo.trim() : '';
      if (!trimmedRegdNo) { 
        errors.push(`Invalid registration number format: ${regdNo}`); 
        continue; 
      }
      
      console.log(`🔍 Processing student: ${trimmedRegdNo}`);
      
      // Check if student exists and is assigned to this faculty
      const details = studentDetailsMap[trimmedRegdNo];
      if (!details) { 
        invalidStudents.push(trimmedRegdNo);
        errors.push(`Student ${trimmedRegdNo} not found or not assigned to your email.`); 
        continue; 
      }
      
      // Check if student is verified
      if (details.verified !== '✅') { 
        errors.push(`Student ${trimmedRegdNo} is not verified. Only verified students can have edit access requested.`); 
        continue; 
      }
      
      // ✅ ENHANCED: Check for existing requests with detailed status
      const existingRequest = existingRequests.get(trimmedRegdNo);
      if (existingRequest) {
        console.log(`⚠️ Existing request found for ${trimmedRegdNo}:`, existingRequest);
        
        if (existingRequest.status === 'Pending') {
          duplicateRequests.push({
            regdNo: trimmedRegdNo,
            name: details.name,
            status: existingRequest.status,
            requestTime: existingRequest.requestTime,
            requestId: existingRequest.requestId
          });
          warnings.push(`Student ${trimmedRegdNo} already has a PENDING edit request. Please wait for approval.`);
          continue;
        } else if (existingRequest.status === 'Approved') {
          // Check if approval is still valid (within 48 hours)
          const unlockUntilCol = editHeaders.indexOf("Unlock Until");
          if (unlockUntilCol !== -1) {
            const unlockUntil = editData[existingRequest.rowIndex][unlockUntilCol];
            if (unlockUntil && new Date(unlockUntil) > new Date()) {
              warnings.push(`Student ${trimmedRegdNo} already has ACTIVE edit access until ${new Date(unlockUntil).toLocaleString()}.`);
              continue;
            }
          }
        }
        // If request was disapproved or expired, allow new request
      }

      // ✅ CREATE NEW REQUEST
      const requestId = `REQ-${timestamp.getTime()}-${trimmedRegdNo}-${Math.random().toString(36).substring(2, 8)}`;
      
      requestsToAdd.push([
        requestId,                    // Request ID
        trimmedEmail,                 // Faculty Email  
        trimmedRegdNo,                // Student Regd No
        details.name,                 // Student Name
        details.paperCode,            // Paper Code
        details.marks,                // Current Marks
        timestamp,                    // Request Time
        'Pending',                    // Status
        'Requested via faculty portal', // Action Notes
        ''                            // Unlock Until (empty for pending)
      ]);
      
      processedCount++;
      console.log(`✅ Created request for student: ${trimmedRegdNo}`);
    }

    // ✅ ENHANCED RESPONSE HANDLING
    let responseData = {
      success: false,
      message: '',
      processedCount: processedCount,
      totalRequested: regdNos.length,
      errors: errors,
      warnings: warnings,
      duplicateRequests: duplicateRequests,
      invalidStudents: invalidStudents,
      timestamp: timestamp.toISOString()
    };

    // ✅ HANDLE DUPLICATE REQUESTS SPECIALLY
    if (duplicateRequests.length > 0 && processedCount === 0) {
      console.log("⚠️ All requests were duplicates");
      responseData.success = false;
      responseData.message = `All selected students already have pending edit requests. No new requests created.`;
      responseData.code = "ALL_DUPLICATE";
      
      // Don't send CoE notification for duplicates
      return responseData;
    }

    // ✅ ADD NEW REQUESTS TO SHEET
    if (requestsToAdd.length > 0) {
      console.log(`📝 Adding ${requestsToAdd.length} new requests to EditRequests sheet`);
      
      editSheet.getRange(editSheet.getLastRow() + 1, 1, requestsToAdd.length, requestsToAdd[0].length)
               .setValues(requestsToAdd);
      
      formatEditRequestsSheet();
      
      // ✅ SEND COE NOTIFICATION
      try {
        const notificationSent = sendCoeNotification(trimmedEmail, requestsToAdd);
        console.log(`📧 CoE notification sent: ${notificationSent}`);
      } catch (notificationError) {
        console.error("⚠️ Failed to send CoE notification:", notificationError);
        warnings.push("Requests created but notification email failed. Please contact examination section.");
      }
    }

    // ✅ BUILD COMPREHENSIVE RESPONSE MESSAGE
    if (processedCount > 0) {
      responseData.success = true;
      responseData.message = `Successfully submitted ${processedCount} edit request(s).`;
      
      if (warnings.length > 0) {
        responseData.message += ` ${warnings.length} warning(s) noted.`;
      }
      if (errors.length > 0) {
        responseData.message += ` ${errors.length} student(s) could not be processed.`;
      }
    } else {
      responseData.success = false;
      if (duplicateRequests.length > 0) {
        responseData.message = `No new requests created. ${duplicateRequests.length} student(s) already have pending requests.`;
        responseData.code = "ALL_DUPLICATE";
      } else if (invalidStudents.length > 0) {
        responseData.message = `No valid students found for edit request.`;
        responseData.code = "NO_VALID_STUDENTS";
      } else {
        responseData.message = `Unable to process edit requests. Please check student verification status.`;
        responseData.code = "PROCESSING_FAILED";
      }
    }

    console.log("✅ Enhanced request processing completed:", responseData);
    return responseData;
    
  } catch (error) {
    console.error(`🚨 CRITICAL ERROR in enhancedRequestEditAccess for ${email}:`, error, error.stack);
    return { 
      success: false, 
      processedCount: 0, 
      errors: errors.concat([`Server error: ${error.message}`]), 
      message: `Server error occurred: ${error.message}`,
      code: "SERVER_ERROR",
      timestamp: new Date().toISOString()
    };
  }
}

// ===== GET UNLOCKED STUDENTS FUNCTION =====
function getUnlockedRegdNos(email) {
  console.log("=== getUnlockedRegdNos ENHANCED DEBUG START ===");
  console.log("Parameter received:", email);
  console.log("Parameter type:", typeof email);
  console.log("Parameter stringified:", JSON.stringify(email));
  
  // ===== COMPREHENSIVE PARAMETER VALIDATION =====
  if (email === undefined) {
    console.error("getUnlockedRegdNos: ❌ Email parameter is UNDEFINED");
    console.error("This indicates a bug in the calling code!");
    console.error("Call stack:", new Error().stack);
    return new Set(); // Return empty Set instead of crashing
  }
  
  if (email === null) {
    console.error("getUnlockedRegdNos: ❌ Email parameter is NULL");
    console.error("Call stack:", new Error().stack);
    return new Set();
  }
  
  if (typeof email !== 'string') {
    console.error("getUnlockedRegdNos: ❌ Email parameter is not a string:", typeof email, email);
    console.error("Call stack:", new Error().stack);
    return new Set();
  }
  
  const trimmedEmail = email.trim();
  if (trimmedEmail === '') {
    console.error("getUnlockedRegdNos: ❌ Email parameter is empty after trimming");
    console.error("Original email:", JSON.stringify(email));
    console.error("Call stack:", new Error().stack);
    return new Set();
  }
  
  // Validate email format
  if (!isValidEmail(trimmedEmail)) {
    console.error("getUnlockedRegdNos: ❌ Invalid email format:", trimmedEmail);
    console.error("Call stack:", new Error().stack);
    return new Set();
  }
  
  console.log("getUnlockedRegdNos: ✅ Parameter validation passed for:", trimmedEmail);
  
  const unlockedRegdNos = new Set();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      console.log("getUnlockedRegdNos: EditRequests sheet not found - returning empty set");
      return unlockedRegdNos;
    }

    const editData = editSheet.getDataRange().getValues();
    if (editData.length < 2) {
      console.log("getUnlockedRegdNos: No data in EditRequests sheet - returning empty set");
      return unlockedRegdNos;
    }

    const headers = editData[0].map(h => h.toString().trim());
    console.log("getUnlockedRegdNos: EditRequests headers:", headers);
    
    const facultyCol = headers.indexOf("Faculty Email");
    const statusCol = headers.indexOf("Status");
    const regdNoCol = headers.indexOf("Student Regd No");
    const unlockUntilCol = headers.indexOf("Unlock Until");

    console.log("getUnlockedRegdNos: Column indices:", { facultyCol, statusCol, regdNoCol, unlockUntilCol });

    if ([facultyCol, statusCol, regdNoCol, unlockUntilCol].includes(-1)) {
      console.error("getUnlockedRegdNos: ❌ Required columns not found in EditRequests sheet");
      console.error("Available headers:", headers);
      return unlockedRegdNos;
    }

    const now = new Date();
    console.log("getUnlockedRegdNos: Current time:", now.toLocaleString());
    
    // Convert email to lowercase for comparison
    const emailLower = trimmedEmail.toLowerCase();
    console.log("getUnlockedRegdNos: Searching for email:", emailLower);

    let matchedRows = 0;
    let approvedRows = 0;
    let activeUnlocks = 0;

    for (let i = 1; i < editData.length; i++) {
      const row = editData[i];
      const rowFaculty = row[facultyCol] ? row[facultyCol].toString().trim().toLowerCase() : '';
      const rowStatus = row[statusCol] ? row[statusCol].toString().trim() : '';
      const rowRegdNo = row[regdNoCol] ? row[regdNoCol].toString().trim() : '';
      const rowUnlockUntil = row[unlockUntilCol];

      if (rowFaculty === emailLower) {
        matchedRows++;
        console.log(`getUnlockedRegdNos: Row ${i} - Faculty match: Status=${rowStatus}, RegdNo=${rowRegdNo}`);
        
        if (rowStatus === "Approved") {
          approvedRows++;
          console.log(`getUnlockedRegdNos: Row ${i} - Status is Approved`);
          
          // Enhanced date handling
          const unlockDate = parseUnlockDate(rowUnlockUntil);
          
          if (unlockDate && unlockDate > now) {
            activeUnlocks++;
            console.log(`getUnlockedRegdNos: ✅ Student ${rowRegdNo} is UNLOCKED until ${unlockDate.toLocaleString()}`);
            unlockedRegdNos.add(rowRegdNo);
          } else if (unlockDate) {
            console.log(`getUnlockedRegdNos: ⏰ Student ${rowRegdNo} unlock EXPIRED at ${unlockDate.toLocaleString()}`);
          } else {
            console.log(`getUnlockedRegdNos: ❌ Student ${rowRegdNo} has INVALID unlock date: ${rowUnlockUntil}`);
          }
        }
      }
    }

    console.log(`getUnlockedRegdNos: Summary for ${emailLower}:`);
    console.log(`  - Total rows checked: ${editData.length - 1}`);
    console.log(`  - Faculty email matches: ${matchedRows}`);
    console.log(`  - Approved requests: ${approvedRows}`);
    console.log(`  - Active unlocks: ${activeUnlocks}`);
    console.log(`  - Unlocked RegdNos:`, Array.from(unlockedRegdNos));

    console.log("=== getUnlockedRegdNos ENHANCED DEBUG END ===");
    return unlockedRegdNos;
    
  } catch (error) {
    console.error("getUnlockedRegdNos: ❌ CRITICAL ERROR:", error);
    console.error("Stack trace:", error.stack);
    console.error("Email parameter:", trimmedEmail);
    return unlockedRegdNos; // Return empty Set instead of crashing
  }
}

function parseUnlockDate(dateValue) {
  if (!dateValue) {
    console.log("parseUnlockDate: No date value provided");
    return null;
  }
  
  try {
    let unlockDate = null;
    
    if (dateValue instanceof Date) {
      unlockDate = dateValue;
      console.log("parseUnlockDate: Date object provided:", unlockDate);
    } else if (typeof dateValue === 'string' && dateValue.trim() !== '') {
      unlockDate = new Date(dateValue);
      console.log("parseUnlockDate: String converted to date:", dateValue, "->", unlockDate);
    } else if (typeof dateValue === 'number') {
      unlockDate = new Date(dateValue);
      console.log("parseUnlockDate: Number converted to date:", dateValue, "->", unlockDate);
    } else {
      console.log("parseUnlockDate: Unsupported date format:", typeof dateValue, dateValue);
      return null;
    }
    
    // Validate the parsed date
    if (unlockDate && !isNaN(unlockDate.getTime())) {
      console.log("parseUnlockDate: Valid date parsed:", unlockDate.toLocaleString());
      return unlockDate;
    } else {
      console.log("parseUnlockDate: Invalid date after parsing:", unlockDate);
      return null;
    }
  } catch (error) {
    console.error(`parseUnlockDate: Failed to parse date: ${dateValue}`, error);
    return null;
  }
}

// ===== ENHANCED AUTO-SAVE FUNCTION WITH DYNAMIC MAX MARKS =====
function autoSaveMarks(email, regdNo, marks) {
  console.log("=== autoSaveMarks DEBUG START ===");
  console.log("Parameters received:", {email, regdNo, marks});
  
  // Enhanced parameter validation
  if (!email || typeof email !== 'string' || !isValidEmail(email.trim())) {
    console.error("autoSaveMarks: Invalid email parameter:", email);
    return { success: false, message: "Invalid user email." };
  }
  
  if (!regdNo || typeof regdNo !== 'string' || regdNo.trim() === '') {
    console.error("autoSaveMarks: Invalid regdNo parameter:", regdNo);
    return { success: false, message: "Invalid registration number." };
  }
  
  const trimmedEmail = email.trim().toLowerCase();
  const trimmedRegdNo = regdNo.trim();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowRegdNo = row[columnIndices.regdNo] ? row[columnIndices.regdNo].toString().trim() : '';
      const rowEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      
      if (rowRegdNo === trimmedRegdNo && rowEmail === trimmedEmail) {
        // Get the student's max marks for validation
        const studentMaxMarks = row[columnIndices.maxMarks] ? Number(row[columnIndices.maxMarks]) : 100;
        
        // Validate marks with dynamic max marks
        if (marks !== null && marks !== undefined && marks !== '') {
          const numMarks = Number(marks);
          if (isNaN(numMarks) || numMarks < 0 || numMarks > studentMaxMarks) {
            console.error("autoSaveMarks: Invalid marks value:", marks, "Max allowed:", studentMaxMarks);
            return { success: false, message: `Invalid marks value (${marks}). Must be between 0-${studentMaxMarks}.` };
          }
        }
        
        const processedMarks = marks !== null && marks !== undefined && marks !== '' ? Number(marks) : '';
        
        const isVerified = row[columnIndices.verified] === '✅';
        const isUnlocked = isStudentUnlocked(trimmedEmail, trimmedRegdNo);

        if (isVerified && !isUnlocked) {
          return { success: false, message: "Marks are verified and locked." };
        }

        // Save marks
        sheet.getRange(i + 1, columnIndices.marks + 1).setValue(processedMarks);
        
        // Clear verification if it was set
        if (row[columnIndices.verified] !== '') {
           sheet.getRange(i + 1, columnIndices.verified + 1).setValue('');
        }
        
        SpreadsheetApp.flush();
        console.log(`autoSaveMarks: SUCCESS for ${trimmedRegdNo}: ${processedMarks} (max: ${studentMaxMarks})`);
        return { success: true, message: `Auto-saved for ${trimmedRegdNo}` };
      }
    }
    
    return { success: false, message: "Student not found or not assigned to you." };
    
  } catch (error) {
    console.error(`Error in autoSaveMarks:`, error, error.stack);
    return { success: false, message: `Server error: ${error.message}` };
  }
}

// ===== ENHANCED BULK OPERATIONS WITH DYNAMIC MAX MARKS =====
function bulkSaveMarks(email, marksData) {
  console.log("=== bulkSaveMarks DEBUG START ===");
  console.log("Parameters received:", {email, marksDataCount: marksData ? marksData.length : 0});
  
  if (!email || typeof email !== 'string' || !isValidEmail(email.trim())) {
    console.error("bulkSaveMarks: Invalid email parameter:", email);
    return { success: false, message: "Invalid user email.", results: [] };
  }
  
  if (!Array.isArray(marksData) || marksData.length === 0) {
    console.error("bulkSaveMarks: Invalid marksData parameter:", marksData);
    return { success: false, message: "No marks data provided.", results: [] };
  }
  
  const trimmedEmail = email.trim().toLowerCase();
  const results = [];
  let successCount = 0;
  let errorCount = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);

    for (const markEntry of marksData) {
      const { regdNo, marks } = markEntry;
      
      if (!regdNo || typeof regdNo !== 'string' || regdNo.trim() === '') {
        results.push({ regdNo: regdNo || 'Unknown', success: false, message: "Invalid registration number" });
        errorCount++;
        continue;
      }
      
      const trimmedRegdNo = regdNo.trim();
      
      // Find student in sheet and get their max marks
      let studentFound = false;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowRegdNo = row[columnIndices.regdNo] ? row[columnIndices.regdNo].toString().trim() : '';
        const rowEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
        
        if (rowRegdNo === trimmedRegdNo && rowEmail === trimmedEmail) {
          studentFound = true;
          
          // Get student's max marks for validation
          const studentMaxMarks = row[columnIndices.maxMarks] ? Number(row[columnIndices.maxMarks]) : 100;
          
          // Validate marks with dynamic max marks
          if (marks === null || marks === undefined || isNaN(Number(marks)) || Number(marks) < 0 || Number(marks) > studentMaxMarks) {
            results.push({ regdNo: trimmedRegdNo, success: false, message: `Invalid marks value (${marks}). Must be between 0-${studentMaxMarks}.` });
            errorCount++;
            break;
          }
          
          const numericMarks = Number(marks);
          
          const isVerified = row[columnIndices.verified] === '✅';
          const isUnlocked = isStudentUnlocked(trimmedEmail, trimmedRegdNo);

          if (isVerified && !isUnlocked) {
            results.push({ regdNo: trimmedRegdNo, success: false, message: "Marks are verified and locked" });
            errorCount++;
            break;
          }

          sheet.getRange(i + 1, columnIndices.marks + 1).setValue(numericMarks);
          
          if (row[columnIndices.verified] !== '') {
             sheet.getRange(i + 1, columnIndices.verified + 1).setValue('');
          }
          
          results.push({ regdNo: trimmedRegdNo, success: true, message: "Marks saved successfully" });
          successCount++;
          break;
        }
      }
      
      if (!studentFound) {
        results.push({ regdNo: trimmedRegdNo, success: false, message: "Student not found or not assigned to you" });
        errorCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    return { 
      success: successCount > 0, 
      message: `Bulk save completed: ${successCount} saved, ${errorCount} errors`,
      successCount: successCount,
      errorCount: errorCount,
      results: results
    };
    
  } catch (error) {
    console.error(`Error in bulkSaveMarks:`, error, error.stack);
    return { 
      success: false, 
      message: `Server error: ${error.message}`,
      results: results
    };
  }
}

function bulkVerifyMarks(email, regdNos) {
  console.log("bulkVerifyMarks called with:", {email, regdNosCount: regdNos ? regdNos.length : 0});
  
  if (!email || typeof email !== 'string' || !isValidEmail(email.trim())) {
    return { success: false, message: "Invalid user email.", results: [] };
  }
  
  if (!Array.isArray(regdNos) || regdNos.length === 0) {
    return { success: false, message: "No students provided for verification.", results: [] };
  }
  
  const trimmedEmail = email.trim().toLowerCase();
  const results = [];
  let successCount = 0;
  let errorCount = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);

    for (const regdNo of regdNos) {
      if (!regdNo || typeof regdNo !== 'string' || regdNo.trim() === '') {
        results.push({ regdNo: regdNo || 'Unknown', success: false, message: "Invalid registration number" });
        errorCount++;
        continue;
      }
      
      const trimmedRegdNo = regdNo.trim();
      
      let studentFound = false;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowRegdNo = row[columnIndices.regdNo] ? row[columnIndices.regdNo].toString().trim() : '';
        const rowEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';

        if (rowRegdNo === trimmedRegdNo && rowEmail === trimmedEmail) {
          studentFound = true;
          
          const currentMarks = row[columnIndices.marks];
          
          if (currentMarks === null || currentMarks === '' || currentMarks === undefined) {
            results.push({ regdNo: trimmedRegdNo, success: false, message: "Cannot verify empty marks" });
            errorCount++;
            break;
          }
          
          // Get student's max marks for validation
          const studentMaxMarks = row[columnIndices.maxMarks] ? Number(row[columnIndices.maxMarks]) : 100;
          const numericMarks = Number(currentMarks);
          
          if (isNaN(numericMarks) || numericMarks < 0 || numericMarks > studentMaxMarks) {
            results.push({ regdNo: trimmedRegdNo, success: false, message: `Cannot verify invalid marks (must be 0-${studentMaxMarks})` });
            errorCount++;
            break;
          }
          
          const wasUnlocked = isStudentUnlocked(trimmedEmail, trimmedRegdNo);
          
          if (row[columnIndices.verified] === '✅' && !wasUnlocked) {
            results.push({ regdNo: trimmedRegdNo, success: true, message: "Already verified and locked" });
            successCount++;
            break;
          }

          if (wasUnlocked) {
            lockStudentCompletely(trimmedEmail, trimmedRegdNo);
          }
          
          sheet.getRange(i + 1, columnIndices.verified + 1).setValue('✅');
          
          results.push({ regdNo: trimmedRegdNo, success: true, message: "Marks verified successfully" });
          successCount++;
          break;
        }
      }
      
      if (!studentFound) {
        results.push({ regdNo: trimmedRegdNo, success: false, message: "Student not found or not assigned to you" });
        errorCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    return { 
      success: successCount > 0, 
      message: `Bulk verification completed: ${successCount} verified, ${errorCount} errors`,
      successCount: successCount,
      errorCount: errorCount,
      results: results
    };
    
  } catch (error) {
    console.error(`Error in bulkVerifyMarks:`, error, error.stack);
    return { 
      success: false, 
      message: `Server error: ${error.message}`,
      results: results
    };
  }
}

// ===== UTILITY FUNCTIONS =====
function isStudentUnlocked(email, regdNo) {
  console.log(`=== isStudentUnlocked ENHANCED DEBUG ===`);
  console.log(`Called with email: ${email}, regdNo: ${regdNo}`);
  console.log(`Email type: ${typeof email}, RegdNo type: ${typeof regdNo}`);
  
  // ===== PARAMETER VALIDATION =====
  if (!email || typeof email !== 'string' || email.trim() === '') {
    console.error("isStudentUnlocked: Invalid email parameter:", email);
    return false;
  }
  
  if (!regdNo || typeof regdNo !== 'string' || regdNo.trim() === '') {
    console.error("isStudentUnlocked: Invalid regdNo parameter:", regdNo);
    return false;
  }
  
  const trimmedEmail = email.trim().toLowerCase();
  const trimmedRegdNo = regdNo.trim();
  
  console.log(`isStudentUnlocked: Processed params - email: ${trimmedEmail}, regdNo: ${trimmedRegdNo}`);
  
  try {
    // Get unlocked students for this email
    const unlockedRegdNos = getUnlockedRegdNos(trimmedEmail);
    const isUnlocked = unlockedRegdNos.has(trimmedRegdNo);
    
    console.log(`isStudentUnlocked: Student ${trimmedRegdNo} unlocked status: ${isUnlocked}`);
    console.log(`isStudentUnlocked: All unlocked students for ${trimmedEmail}:`, Array.from(unlockedRegdNos));
    
    return isUnlocked;
    
  } catch (error) {
    console.error("isStudentUnlocked: Error checking unlock status:", error);
    return false; // Safe fallback
  }
}

function lockStudentCompletely(email, regdNo) {
  try {
    console.log(`lockStudentCompletely called for ${email}, ${regdNo}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    if (!editSheet) {
      console.log("EditRequests sheet not found");
      return { success: false, message: "EditRequests sheet not found" };
    }
    
    const data = editSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log("No data in EditRequests sheet");
      return { success: true, message: "No unlock requests to expire" };
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const facultyCol = headers.indexOf("Faculty Email");
    const regdNoCol = headers.indexOf("Student Regd No");
    const statusCol = headers.indexOf("Status");
    const unlockUntilCol = headers.indexOf("Unlock Until");
    const actionNotesCol = headers.indexOf("Action Notes");
    
    if ([facultyCol, regdNoCol, statusCol, unlockUntilCol].includes(-1)) {
      console.error("Required columns not found:", { facultyCol, regdNoCol, statusCol, unlockUntilCol });
      return { success: false, message: "Required columns missing" };
    }
    
    const trimmedEmail = email.trim().toLowerCase();
    const trimmedRegdNo = regdNo.trim();
    let expiredCount = 0;
    const expiredDate = new Date();
    expiredDate.setDate(expiredDate.getDate() - 1);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowFaculty = row[facultyCol] ? row[facultyCol].toString().trim().toLowerCase() : '';
      const rowRegdNo = row[regdNoCol] ? row[regdNoCol].toString().trim() : '';
      const rowStatus = row[statusCol] ? row[statusCol].toString().trim() : '';
      
      if (rowFaculty === trimmedEmail && 
          rowRegdNo === trimmedRegdNo && 
          rowStatus === "Approved") {
        
        editSheet.getRange(i + 1, unlockUntilCol + 1).setValue(expiredDate);
        editSheet.getRange(i + 1, statusCol + 1).setValue("Completed");
        
        const currentNotes = row[actionNotesCol] || '';
        const newNotes = currentNotes + ` | LOCKED after verification on ${new Date().toLocaleString()}`;
        if (actionNotesCol !== -1) {
          editSheet.getRange(i + 1, actionNotesCol + 1).setValue(newNotes);
        }
        
        expiredCount++;
      }
    }
    
    if (expiredCount > 0) {
      SpreadsheetApp.flush();
    }
    
    return { 
      success: true, 
      message: `Expired ${expiredCount} unlock request(s)`,
      expiredCount: expiredCount
    };
    
  } catch (error) {
    console.error("Error in lockStudentCompletely:", error);
    return { success: false, message: error.message };
  }
}

// ===== EMAIL FUNCTIONS =====
function isValidEmail(email) {
  if (!email || typeof email !== "string") return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

function generateOTP() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

// ===== BULK EMAIL SYSTEM FUNCTIONS =====
function sendEmails() {
  try {
    console.log("=== BULK EMAIL SENDING STARTED ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    if (emailColIndex === -1) throw new Error(`Column '${COL_EXAMINER_EMAIL}' not found`);
    
    const facultyEmails = {};
    const errors = [];
    let successCount = 0;
    
    // ✅ FIXED: Collect unique faculty emails and normalize to lowercase
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailColIndex];
      if (email && typeof email === 'string') {
          const trimmedEmail = email.trim().toLowerCase(); // ✅ FIXED: Always lowercase
          if (isValidEmail(trimmedEmail) && !facultyEmails[trimmedEmail]) {
              facultyEmails[trimmedEmail] = generateOTP();
          } else if (!isValidEmail(trimmedEmail)) {
              errors.push(`Invalid email format: ${email}`);
          }
      }
    }
    
    const uniqueEmails = Object.keys(facultyEmails);
    if (uniqueEmails.length === 0) {
      const message = "No valid faculty emails found to send OTPs to.";
      console.error(message);
      safeUIAlert("No Recipients", message);
      return message;
    }
    
    // ✅ FIXED: Store OTPs with debug info
    try {
      PropertiesService.getScriptProperties().setProperty("OTPs", JSON.stringify(facultyEmails));
      console.log(`✅ STORED OTPs for ${uniqueEmails.length} faculty members`);
      console.log("✅ STORED EMAILS:", uniqueEmails.join(', '));
      
      // Verify storage worked
      const verifyStored = PropertiesService.getScriptProperties().getProperty('OTPs');
      if (verifyStored) {
        console.log("✅ STORAGE VERIFICATION: OTPs successfully stored and retrievable");
      } else {
        throw new Error("Storage verification failed");
      }
    } catch (storageError) {
      console.error("❌ STORAGE ERROR:", storageError);
      throw new Error(`Failed to store OTPs: ${storageError.message}`);
    }
    
    // Check email quota
    const emailQuota = MailApp.getRemainingDailyQuota();
    if (emailQuota < uniqueEmails.length) {
        throw new Error(`Insufficient email quota (${emailQuota}) to send ${uniqueEmails.length} emails.`);
    }
    
    const portalLink = WEB_APP_URL;
    
    // Send emails to all faculty
    for (const email of uniqueEmails) {
      try {
        const otp = facultyEmails[email];
        const subject = "Faculty Portal Access - Your One-Time Password [OTP:" + otp + "]";
        
        const message = `Dear Faculty Member,

Sai Ram,

Your One-Time Password (OTP) for accessing the Faculty Portal is ready:

Portal Access: ${portalLink}

Login Credentials:
- 📧 Email: ${email}
- 🔐 Your OTP: ${otp}

⚠️ Important:
• Use the EXACT email address: ${email}
• Use the 6-digit OTP: ${otp}
• Do not share your OTP with anyone

For technical support, please contact the Examinations Section.

Best Regards,
Examinations Section`;

        MailApp.sendEmail(email, subject, message);
        successCount++;
        logEmailActivity("BULK_OTP_EMAIL", email, subject, "SUCCESS");
        
        Utilities.sleep(200); // Rate limiting
        
      } catch (emailError) {
        console.error(`Failed to send email to ${email}:`, emailError);
        errors.push(`Failed to send to ${email}: ${emailError.message}`);
        logEmailActivity("BULK_OTP_EMAIL", email, subject || "OTP Email", `FAILED: ${emailError.message}`);
      }
    }
    
    // Show results
    let resultMessage = `✅ OTP Generation & Email Complete!\n\n`;
    resultMessage += `📊 Results:\n`;
    resultMessage += `• OTPs generated: ${uniqueEmails.length}\n`;
    resultMessage += `• Emails sent: ${successCount}\n`;
    resultMessage += `• Failed: ${errors.length}\n`;
    resultMessage += `• Stored in system: ✅\n`;
    
    if (errors.length > 0) {
      resultMessage += `\n⚠️ Errors: ${errors.length}`;
    }
    
    console.log("✅ BULK EMAIL COMPLETED:", { 
      generated: uniqueEmails.length, 
      sent: successCount, 
      errors: errors.length 
    });
    
    safeUIAlert("Email Generation Complete", resultMessage);
    return `OTP emails sent to ${successCount} faculty members. ${errors.length} errors.`;
    
  } catch (error) {
    console.error("❌ CRITICAL ERROR in sendEmails:", error, error.stack);
    const errorMsg = `❌ Email generation failed: ${error.message}`;
    safeUIAlert("Email Generation Failed", errorMsg);
    return errorMsg;
  }
}

function logEmailActivity(type, recipient, subject, status) {
    const timestamp = new Date().toLocaleString();
    console.log(`📧 EMAIL LOG [${timestamp}] -> Type: ${type}, Recipient: ${recipient}, Status: ${status}` + (subject ? `, Subject: ${subject}`: ''));
    
    // Optional: Store in a logging sheet
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName("EmailLog");
      
      if (!logSheet) {
        logSheet = ss.insertSheet("EmailLog");
        logSheet.appendRow(["Timestamp", "Type", "Recipient", "Subject", "Status"]);
        logSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f0f0f0");
      }
      
      logSheet.appendRow([timestamp, type, recipient, subject || "N/A", status]);
      
    } catch (logError) {
      console.warn("Failed to log email activity to sheet:", logError);
    }
}

// ===== EMAIL NOTIFICATION FUNCTIONS =====
function sendCoeNotification(facultyEmail, requests) {
  const emailSubject = 'Faculty Edit Request Notification';
  
  try {
    if (!facultyEmail || !isValidEmail(facultyEmail.trim())) {
      console.error("Invalid faculty email:", facultyEmail);
      return false;
    }
    
    if (!requests || !Array.isArray(requests) || requests.length === 0) {
      console.error("Invalid or empty requests array:", requests);
      return false;
    }
    
    const coeEmail = 'intern1@sssihl.edu.in';
    
    if (!isValidEmail(coeEmail)) {
      console.error("Invalid CoE email:", coeEmail);
      return false;
    }
    
    let emailBody = `📝 EDIT REQUEST NOTIFICATION

Faculty member (${facultyEmail}) has requested edit access for ${requests.length} student(s).

Student Details:
`;
    
    requests.forEach(request => {
      if (request && Array.isArray(request) && request.length >= 5) {
        emailBody += `📋 ${request[2]} - ${request[3]} (Paper Code: ${request[4]})\n`;
      } else {
        emailBody += `📋 Invalid request data\n`;
        console.warn("Invalid request format:", request);
      }
    });
    
    emailBody += `
📊 Next Steps:
1. Open your examination spreadsheet
2. Check the "EditRequests" sheet to review requests
3. Use "CoE Functions" menu to approve/disapprove
4. Faculty will be automatically notified of your decision

Timestamp: ${new Date().toLocaleString()}

Best regards,
Examination System`;

    MailApp.sendEmail(coeEmail, emailSubject, emailBody);
    console.log("CoE notification sent successfully");
    return true;
  } catch (error) {
    console.error("Error sending CoE notification:", error);
    return false;
  }
}

function sendEditApprovalNotification(facultyEmail, regdNo, studentName) {
  try {
    console.log(`Sending edit approval notification to: ${facultyEmail}`);
    
    if (!isValidEmail(facultyEmail)) {
      console.error("Invalid faculty email:", facultyEmail);
      return false;
    }
    
    const subject = 'Edit Request Approved - Student ' + regdNo;
    const message = `Dear Faculty,

✅ GOOD NEWS! Your edit request has been APPROVED by the Controller of Examination Section.

Student Details:
📋 Registration Number: ${regdNo}
👤 Student Name: ${studentName}

🔓 EDIT ACCESS GRANTED:
You can now log in to the Faculty Portal using your same OTP sent earlier and edit the marks for this student.
The edit access is valid for 48 hours from now.

🌐 Faculty Portal: ${WEB_APP_URL}

📝 How to Edit:
1. Click the portal link above
2. Log in with your institute email and OTP
3. Find the student (${regdNo}) - it will show as "Unlocked for Edit"
4. Enter/modify the marks
5. Use "Save" to save your changes
6. Use "Verify Selected" when you're ready to lock the marks

⏰ IMPORTANT: Edit access expires in 48 hours. After that, you'll need to request edit access again.

If you have any questions, please contact the examination committee.

Best regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning`;

    MailApp.sendEmail(facultyEmail, subject, message);
    console.log("Edit approval notification sent successfully to:", facultyEmail);
    logEmailActivity('EDIT_APPROVAL', facultyEmail, subject, 'SUCCESS');
    return true;
  } catch (error) {
    console.error("Error sending edit approval notification:", error);
    logEmailActivity('EDIT_APPROVAL', facultyEmail, subject || "Edit Approval", 'FAILED: ' + error.message);
    return false;
  }
}

function sendDisapprovalNotification(facultyEmail, regdNo, reason) {
  try {
    console.log(`Sending disapproval notification to: ${facultyEmail}`);
    
    if (!isValidEmail(facultyEmail)) {
      console.error("Invalid faculty email:", facultyEmail);
      return false;
    }
    
    const subject = 'Edit Request Disapproved - Student ' + regdNo;
    const message = `Dear Faculty,

❌ Your edit request has been DISAPPROVED by the Controller of Examinations Section.

Student Details:
📋 Registration Number: <strong>${regdNo}</strong>

📝 Reason for Disapproval: <strong>${reason}</strong>

Next Steps:
• If you believe this is an error, please contact the examination committee
• You may submit a new edit request with proper justification
• Ensure all required documentation is provided

For questions or appeals, please contact the Controller of Examinations Section directly.

Best regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning`;

    MailApp.sendEmail(facultyEmail, subject, message);
    console.log("Disapproval notification sent successfully to:", facultyEmail);
    logEmailActivity('EDIT_DISAPPROVAL', facultyEmail, subject, 'SUCCESS');
    return true;
  } catch (error) {
    console.error("Error sending disapproval notification:", error);
    logEmailActivity('EDIT_DISAPPROVAL', facultyEmail, subject || "Edit Disapproval", 'FAILED: ' + error.message);
    return false;
  }
}

// ===== WEB APP ENTRY POINT =====
function doGet(e) {
  console.log("doGet triggered with parameters:", JSON.stringify(e));
  try {
    const params = e && e.parameter ? e.parameter : {};
    if (params.action === "approveSecondEdit") {
      const requestId = params.requestId;
      if (!requestId) {
        console.error("approveSecondEdit action called without requestId.");
        return HtmlService.createHtmlOutput("<h1>Error</h1><p>No request ID provided.</p>").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
      const approvalResult = approveSecondEdit(requestId);
      return HtmlService.createHtmlOutput("Approval action response not fully implemented.");
    } else {
      console.log("Serving faculty portal HTML.");
      return HtmlService.createTemplateFromFile("faculty_portal")
        .evaluate()
        .setTitle("Faculty Portal")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (error) {
    console.error("Error in doGet:", error, error.stack);
    return HtmlService.createHtmlOutput(
      "<h1>Error</h1><p>Sorry, an unexpected error occurred while loading the page. Please try again later or contact support.</p>" +
      "<p>Error details logged for administrator.</p>"
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ===== ENHANCED OTP VERIFICATION FUNCTION =====
function verifyOTP(email, otp) {
  console.log(`=== verifyOTP DEBUG START ===`);
  console.log(`verifyOTP called with email: ${email}, otp: ${otp}`);
  console.log(`Email type: ${typeof email}, OTP type: ${typeof otp}`);
  
  // Enhanced parameter validation
  if (email === undefined || email === null || email === '') {
    console.error("verifyOTP: Invalid email parameter:", email);
    return false;
  }
  
  if (otp === undefined || otp === null || otp === '') {
    console.error("verifyOTP: Invalid OTP parameter:", otp);
    return false;
  }
  
  // Convert to string and validate
  const emailStr = String(email).trim().toLowerCase(); // ✅ FIXED: Always convert to lowercase
  const otpStr = String(otp).trim();
  
  if (!isValidEmail(emailStr)) {
    console.error("verifyOTP: Invalid email format:", emailStr);
    return false;
  }
  
  if (!/^\d{6}$/.test(otpStr)) {
    console.error("verifyOTP: Invalid OTP format:", otpStr);
    return false;
  }
  
  try {
    const storedOTPs = PropertiesService.getScriptProperties().getProperty('OTPs');
    if (!storedOTPs) {
      console.error("verifyOTP: No stored OTPs found");
      console.error("SOLUTION: Run sendEmails() function first to generate OTPs");
      return false;
    }
    
    let otpMap;
    try {
      otpMap = JSON.parse(storedOTPs);
    } catch (parseError) {
      console.error("verifyOTP: Failed to parse stored OTPs:", parseError);
      console.error("SOLUTION: Clear Properties and regenerate OTPs");
      PropertiesService.getScriptProperties().deleteProperty('OTPs');
      return false;
    }
    
    console.log("verifyOTP: Available emails in OTP map:", Object.keys(otpMap));
    
    // ✅ FIXED: Try both original email and lowercase email
    let expectedOTP = otpMap[emailStr]; // Try lowercase first
    if (!expectedOTP) {
      // Try to find email in different case
      const foundEmail = Object.keys(otpMap).find(storedEmail => 
        storedEmail.toLowerCase() === emailStr
      );
      if (foundEmail) {
        expectedOTP = otpMap[foundEmail];
        console.log(`verifyOTP: Found email with different case: ${foundEmail}`);
      }
    }
    
    if (expectedOTP && expectedOTP === otpStr) {
      console.log(`verifyOTP: SUCCESS for ${emailStr}`);
      return true;
    } else {
      console.log(`verifyOTP: FAILED for ${emailStr}`);
      console.log(`Expected OTP: ${expectedOTP}, Received OTP: ${otpStr}`);
      console.log(`Available emails: ${Object.keys(otpMap).join(', ')}`);
      return false;
    }
  } catch (error) {
    console.error("Error during verifyOTP:", error);
    return false;
  }
}

// ===== APPROVAL SYSTEM FUNCTIONS =====
function refreshEditRequestsView() {
  try {
    console.log("=== REFRESH EDIT REQUESTS VIEW ===");
    
    formatEditRequestsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (editSheet) {
      editSheet.activate();
      safeUIToast("Edit Requests view refreshed successfully.", "Success", 3);
    } else {
      safeUIAlert("Sheet Not Found", `The "${EDIT_REQUESTS_SHEET}" sheet was not found. Please run the initialization function first.`);
    }
    
    console.log("Edit requests view refreshed successfully");
    return "Edit requests view refreshed successfully";
    
  } catch (error) {
    console.error("Error refreshing edit requests view:", error);
    safeUIAlert("Refresh Error", `Failed to refresh view: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

function approveSelectedRequest() {
  try {
    console.log("=== ENHANCED APPROVE SELECTED REQUEST ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot approve request via menu");
      return "UI not available in this context.";
    }
    
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      throw new Error(`Sheet "${EDIT_REQUESTS_SHEET}" not found.`);
    }
    
    // Get the currently selected row
    const activeRange = editSheet.getActiveRange();
    if (!activeRange || activeRange.getNumRows() !== 1) {
      safeUIAlert("Selection Error", "Please select a single row containing the request to approve.");
      return "No valid row selected";
    }
    
    const row = activeRange.getRow();
    const dataStartRow = 3; // Account for header and instruction rows
    
    if (row < dataStartRow) {
      safeUIAlert("Selection Error", `Please select a data row (row ${dataStartRow} or below).`);
      return "Invalid row selected";
    }
    
    // ===== ENHANCED ROW DATA READING =====
    console.log(`Reading data from row ${row}`);
    const rowValues = editSheet.getRange(row, 1, 1, 10).getValues()[0];
    
    // Log all row values for debugging
    console.log("Raw row values:", rowValues);
    
    const requestId = rowValues[0];
    const facultyEmail = rowValues[1];
    const regdNo = rowValues[2];
    const studentName = rowValues[3];
    const paperCodeName = rowValues[4];
    const currentMarks = rowValues[5];
    const requestTime = rowValues[6];
    const status = rowValues[7];
    const actionNotes = rowValues[8];
    const unlockUntil = rowValues[9];
    
    console.log("Parsed request data:", {
      requestId, facultyEmail, regdNo, studentName, paperCodeName,
      currentMarks, requestTime, status, actionNotes, unlockUntil
    });
    
    // Validate the request
    if (!requestId) {
      safeUIAlert("Invalid Row", "Invalid row selected. Cannot find Request ID.");
      return "Invalid request ID";
    }
    
    if (status !== "Pending") {
      safeUIAlert("Invalid Status", `Request cannot be approved. Current status: ${status}`);
      return `Cannot approve request with status: ${status}`;
    }
    
    // Show confirmation dialog
    const confirmMessage = `Approve edit request?\n\n` +
                          `Request ID: ${requestId}\n` +
                          `Faculty: ${facultyEmail}\n` +
                          `Student: ${regdNo} - ${studentName}\n` +
                          `Paper Code: ${paperCodeName}\n` +
                          `Current Marks: ${currentMarks}\n\n` +
                          `This will grant 48-hour edit access to the faculty member.`;
    
    const response = safeUIAlert("Confirm Approval", confirmMessage, ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      // ===== ENHANCED APPROVAL PROCESS =====
      console.log("User confirmed approval, processing...");
      
      // Calculate unlock expiry (48 hours from now)
      const unlockUntilDate = new Date();
      unlockUntilDate.setHours(unlockUntilDate.getHours() + 48);
      
      console.log("Setting unlock until:", unlockUntilDate.toLocaleString());
      
      // Get current sheet headers to ensure we're updating the right columns
      const headers = editSheet.getRange(1, 1, 1, 10).getValues()[0];
      console.log("Sheet headers:", headers);
      
      // Find the correct column indices
      const statusColIndex = headers.indexOf("Status");
      const actionNotesColIndex = headers.indexOf("Action Notes");
      const unlockUntilColIndex = headers.indexOf("Unlock Until");
      
      console.log("Column indices:", { statusColIndex, actionNotesColIndex, unlockUntilColIndex });
      
      if (statusColIndex === -1 || actionNotesColIndex === -1 || unlockUntilColIndex === -1) {
        console.error("Could not find required columns in EditRequests sheet");
        safeUIAlert("Error", "Could not find required columns in EditRequests sheet. Please check sheet structure.");
        return "Column mapping error";
      }
      
      // Update the request status with proper column indices
      try {
        console.log(`Updating row ${row}, status column ${statusColIndex + 1} to "Approved"`);
        editSheet.getRange(row, statusColIndex + 1).setValue("Approved");
        
        const approvalNote = `Approved via Menu on ${new Date().toLocaleString()}`;
        console.log(`Updating row ${row}, action notes column ${actionNotesColIndex + 1} to: ${approvalNote}`);
        editSheet.getRange(row, actionNotesColIndex + 1).setValue(approvalNote);
        
        console.log(`Updating row ${row}, unlock until column ${unlockUntilColIndex + 1} to: ${unlockUntilDate}`);
        editSheet.getRange(row, unlockUntilColIndex + 1).setValue(unlockUntilDate);
        
        // Force spreadsheet to save
        SpreadsheetApp.flush();
        console.log("✅ Request data updated successfully");
        
      } catch (updateError) {
        console.error("Error updating request data:", updateError);
        safeUIAlert("Update Error", `Failed to update request: ${updateError.message}`);
        return "Update failed";
      }
      
      // Apply formatting
      try {
        formatEditRequestsSheet();
        console.log("✅ Sheet formatting applied");
      } catch (formatError) {
        console.warn("Warning: Sheet formatting failed:", formatError);
      }
      
      // Send notification email to faculty
      let emailSent = false;
      try {
        emailSent = sendEditApprovalNotification(facultyEmail, regdNo, studentName);
        console.log("Email notification result:", emailSent);
      } catch (emailError) {
        console.error("Error sending email notification:", emailError);
      }
      
      let successMessage = "✅ Request approved successfully!\n\n";
      successMessage += `Faculty: ${facultyEmail}\n`;
      successMessage += `Student: ${regdNo} - ${studentName}\n`;
      successMessage += `Edit access granted until: ${unlockUntilDate.toLocaleString()}\n`;
      
      if (emailSent) {
        successMessage += "\n📧 Notification email sent to faculty.";
      } else {
        successMessage += "\n⚠️ Request approved but email notification failed.";
      }
      
      safeUIAlert("Approval Successful", successMessage);
      console.log("✅ Request approval completed successfully");
      
      return "Request approved successfully";
    } else {
      console.log("Approval cancelled by user");
      return "Approval cancelled";
    }
    
  } catch (error) {
    console.error("❌ Error in approveSelectedRequest:", error, error.stack);
    safeUIAlert("Approval Error", `Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

function disapproveSelectedRequest() {
  try {
    console.log("=== DISAPPROVE SELECTED REQUEST ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot disapprove request via menu");
      return "UI not available in this context.";
    }
    
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      throw new Error(`Sheet "${EDIT_REQUESTS_SHEET}" not found.`);
    }
    
    const activeRange = editSheet.getActiveRange();
    if (!activeRange || activeRange.getNumRows() !== 1) {
      safeUIAlert("Selection Error", "Please select a single row containing the request to disapprove.");
      return "No valid row selected";
    }
    
    const row = activeRange.getRow();
    const dataStartRow = 3;
    
    if (row < dataStartRow) {
      safeUIAlert("Selection Error", `Please select a data row (row ${dataStartRow} or below).`);
      return "Invalid row selected";
    }
    
    const rowValues = editSheet.getRange(row, 1, 1, 10).getValues()[0];
    const requestId = rowValues[0];
    const facultyEmail = rowValues[1];
    const regdNo = rowValues[2];
    const studentName = rowValues[3];
    const status = rowValues[7];
    
    if (!requestId) {
      safeUIAlert("Invalid Row", "Invalid row selected. Cannot find Request ID.");
      return "Invalid request ID";
    }
    
    if (status !== "Pending") {
      safeUIAlert("Invalid Status", `Request cannot be disapproved. Current status: ${status}`);
      return `Cannot disapprove request with status: ${status}`;
    }
    
    // Get reason for disapproval
    const reasonResponse = safeUIPrompt("Reason for Disapproval", "Please provide a brief reason for disapproving this request:");
    
    if (!reasonResponse || reasonResponse.getSelectedButton() !== ui.Button.OK) {
      console.log("Disapproval cancelled - no reason provided");
      return "Disapproval cancelled";
    }
    
    const reason = reasonResponse.getResponseText().trim();
    if (!reason) {
      safeUIAlert("Missing Reason", "Reason for disapproval cannot be empty.");
      return "Reason required";
    }
    
    // Show final confirmation
    const confirmMessage = `Disapprove edit request?\n\n` +
                          `Request ID: ${requestId}\n` +
                          `Faculty: ${facultyEmail}\n` +
                          `Student: ${regdNo} - ${studentName}\n` +
                          `Reason: ${reason}`;
    
    const confirmResponse = safeUIAlert("Confirm Disapproval", confirmMessage, ui.ButtonSet.YES_NO);
    
    if (confirmResponse === ui.Button.YES) {
      // Update the request status
      editSheet.getRange(row, 8).setValue("Disapproved"); // Status column
      editSheet.getRange(row, 9).setValue(`Disapproved on ${new Date().toLocaleString()}. Reason: ${reason}`); // Action Notes
      
      // Apply formatting
      formatEditRequestsSheet();
      
      // Send notification email to faculty
      const emailSent = sendDisapprovalNotification(facultyEmail, regdNo, reason);
      
      let successMessage = "❌ Request disapproved successfully!\n\n";
      successMessage += `Faculty: ${facultyEmail}\n`;
      successMessage += `Student: ${regdNo} - ${studentName}\n`;
      successMessage += `Reason: ${reason}\n`;
      
      if (emailSent) {
        successMessage += "\n📧 Notification email sent to faculty.";
      } else {
        successMessage += "\n⚠️ Request disapproved but email notification failed.";
      }
      
      safeUIAlert("Disapproval Successful", successMessage);
      console.log("Request disapproved successfully:", requestId);
      
      return "Request disapproved successfully";
    } else {
      console.log("Disapproval cancelled by user");
      return "Disapproval cancelled";
    }
    
  } catch (error) {
    console.error("Error in disapproveSelectedRequest:", error, error.stack);
    safeUIAlert("Disapproval Error", `Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

function approveAllPending() {
  try {
    console.log("=== APPROVE ALL PENDING REQUESTS ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot approve all pending via menu");
      return "UI not available in this context.";
    }
    
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      throw new Error(`Sheet "${EDIT_REQUESTS_SHEET}" not found.`);
    }
    
    const data = editSheet.getDataRange().getValues();
    if (data.length < 3) { // Header + instruction + at least one data row
      safeUIAlert("No Data", "No edit requests found to process.");
      return "No requests found";
    }
    
    // Find all pending requests
    const pendingRequests = [];
    const headers = data[0];
    const statusCol = headers.indexOf("Status");
    const facultyCol = headers.indexOf("Faculty Email");
    const regdNoCol = headers.indexOf("Student Regd No");
    const studentNameCol = headers.indexOf("Student Name");
    
    if ([statusCol, facultyCol, regdNoCol].includes(-1)) {
      throw new Error("Required columns not found in EditRequests sheet");
    }
    
    for (let i = 2; i < data.length; i++) { // Start from row 3 (index 2)
      const row = data[i];
      if (row[statusCol] === "Pending") {
        pendingRequests.push({
          rowIndex: i + 1, // Convert to 1-based row index
          requestId: row[0],
          facultyEmail: row[facultyCol],
          regdNo: row[regdNoCol],
          studentName: row[studentNameCol] || "Unknown"
        });
      }
    }
    
    if (pendingRequests.length === 0) {
      safeUIAlert("No Pending Requests", "No pending edit requests found to approve.");
      return "No pending requests";
    }
    
    // Show confirmation
    let confirmMessage = `Approve ALL ${pendingRequests.length} pending edit requests?\n\n`;
    confirmMessage += "This will:\n";
    confirmMessage += `• Grant 48-hour edit access to ${pendingRequests.length} faculty members\n`;
    confirmMessage += "• Send notification emails to all faculty\n";
    confirmMessage += "• Cannot be undone\n\n";
    confirmMessage += "Pending requests:\n";
    
    // Show first few requests as preview
    const previewCount = Math.min(5, pendingRequests.length);
    for (let i = 0; i < previewCount; i++) {
      const req = pendingRequests[i];
      confirmMessage += `• ${req.facultyEmail} → ${req.regdNo}\n`;
    }
    
    if (pendingRequests.length > previewCount) {
      confirmMessage += `... and ${pendingRequests.length - previewCount} more requests`;
    }
    
    const response = safeUIAlert("Approve All Pending", confirmMessage, ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      console.log("Bulk approval cancelled by user");
      return "Bulk approval cancelled";
    }
    
    // Process all approvals
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    const unlockUntil = new Date();
    unlockUntil.setHours(unlockUntil.getHours() + 48);
    
    for (const request of pendingRequests) {
      try {
        const row = request.rowIndex;
        
        // Update status
        editSheet.getRange(row, statusCol + 1).setValue("Approved");
        editSheet.getRange(row, statusCol + 2).setValue(`Bulk approved on ${new Date().toLocaleString()}`); // Action Notes
        editSheet.getRange(row, statusCol + 3).setValue(unlockUntil); // Unlock Until
        
        // Send notification email
        const emailSent = sendEditApprovalNotification(request.facultyEmail, request.regdNo, request.studentName);
        
        if (!emailSent) {
          console.warn(`Email notification failed for ${request.facultyEmail}`);
        }
        
        successCount++;
        
        // Add small delay to avoid rate limiting
        Utilities.sleep(100);
        
      } catch (requestError) {
        console.error(`Error processing request ${request.requestId}:`, requestError);
        errors.push(`${request.requestId}: ${requestError.message}`);
        errorCount++;
      }
    }
    
    // Apply formatting
    formatEditRequestsSheet();
    
    // Show results
    let resultMessage = `✅ Bulk approval completed!\n\n`;
    resultMessage += `📊 Results:\n`;
    resultMessage += `• Successfully approved: ${successCount}\n`;
    resultMessage += `• Errors: ${errorCount}\n`;
    resultMessage += `• Edit access expires: ${unlockUntil.toLocaleString()}\n`;
    
    if (errorCount > 0) {
      resultMessage += `\n⚠️ Errors encountered:\n`;
      errors.slice(0, 3).forEach(error => resultMessage += `• ${error}\n`);
      if (errors.length > 3) {
        resultMessage += `• ... and ${errors.length - 3} more errors`;
      }
    }
    
    safeUIAlert("Bulk Approval Results", resultMessage);
    console.log("Bulk approval completed:", { successCount, errorCount, total: pendingRequests.length });
    
    return `Bulk approval completed: ${successCount} approved, ${errorCount} errors`;
    
  } catch (error) {
    console.error("Error in approveAllPending:", error, error.stack);
    safeUIAlert("Bulk Approval Error", `Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

// ===== OTHER FUNCTIONS (unchanged from original code) =====
function getDetailedPendingRequests(email) {
  console.log("🔍 Getting detailed pending requests for:", email);
  
  const pendingRequests = new Set();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      console.log("EditRequests sheet not found");
      return pendingRequests;
    }

    const editData = editSheet.getDataRange().getValues();
    if (editData.length < 2) {
      console.log("No data in EditRequests sheet");
      return pendingRequests;
    }

    const headers = editData[0].map(h => h.toString().trim());
    const facultyCol = headers.indexOf("Faculty Email");
    const statusCol = headers.indexOf("Status");
    const regdNoCol = headers.indexOf("Student Regd No");

    if ([facultyCol, statusCol, regdNoCol].includes(-1)) {
      console.error("Required columns not found in EditRequests sheet");
      return pendingRequests;
    }

    const emailLower = email.toLowerCase();
    
    for (let i = 1; i < editData.length; i++) {
      const row = editData[i];
      const rowFaculty = row[facultyCol] ? row[facultyCol].toString().trim().toLowerCase() : '';
      const rowStatus = row[statusCol] ? row[statusCol].toString().trim() : '';
      const rowRegdNo = row[regdNoCol] ? row[regdNoCol].toString().trim() : '';

      if (rowFaculty === emailLower && rowStatus === "Pending" && rowRegdNo) {
        pendingRequests.add(rowRegdNo);
        console.log(`📋 Found pending request for ${rowRegdNo}`);
      }
    }

    console.log(`✅ Detailed pending requests completed: ${pendingRequests.size} found`);
    return pendingRequests;
    
  } catch (error) {
    console.error("Error in getDetailedPendingRequests:", error);
    return pendingRequests;
  }
}

// ===== SAFE UI HELPER FUNCTIONS =====
function safeUIAlert(title, message, buttonSet = null) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui && typeof ui.alert === 'function') {
      const buttons = buttonSet || ui.ButtonSet.OK;
      return ui.alert(title, message, buttons);
    }
    console.log(`UI Alert (${title}): ${message}`);
    return null;
  } catch (error) {
    console.log(`UI Alert (${title}): ${message}`);
    return null;
  }
}

function safeUIPrompt(title, message, buttonSet = null) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui && typeof ui.prompt === 'function') {
      const buttons = buttonSet || ui.ButtonSet.OK_CANCEL;
      return ui.prompt(title, message, buttons);
    }
    console.log(`UI Prompt (${title}): ${message}`);
    return null;
  } catch (error) {
    console.log(`UI Prompt (${title}): ${message}`);
    return null;
  }
}

function safeUIToast(message, title = "Notification", timeoutSeconds = 5) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui && typeof SpreadsheetApp.getActiveSpreadsheet === 'function') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss && typeof ss.toast === 'function') {
        ss.toast(message, title, timeoutSeconds);
        return true;
      }
    }
    console.log(`Toast (${title}): ${message}`);
    return false;
  } catch (error) {
    console.log(`Toast (${title}): ${message}`);
    return false;
  }
}

function detectExecutionContext() {
  try {
    SpreadsheetApp.getUi();
    return 'spreadsheet';
  } catch (e) {
    return 'webapp';
  }
}

function createMenus() {
  try {
    if (!isUIAvailable()) {
      console.log("UI not available - skipping menu creation");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Exam Tools")
      .addItem("Send Email", "sendEmails")
      .addItem("Analysis","analyse")
      .addItem("🔍 Debug Sheet Headers", "debugSheetHeaders")
      .addSeparator()
      .addSubMenu(ui.createMenu("CoE Functions")
        .addItem("🔄 Refresh Edit Requests", "refreshEditRequestsView")
        .addSeparator()
        .addItem("✅ Approve Selected Request", "approveSelectedRequest")
        .addItem("❌ Disapprove Selected Request", "disapproveSelectedRequest")
        .addItem("✅ Approve All Pending", "approveAllPending")
        .addSeparator()
        .addItem("🧪 Test Approval System", "testApprovalSystem")
        .addItem("🔓 Test Unlock System (UI)", "testUnlockSystemWithUI")
        .addItem("🔍 Test Unlock System (Console)", "testUnlockSystem")
        .addItem("🌐 Open Faculty Portal", "openFacultyPortal"))
      .addSeparator()
      .addSubMenu(ui.createMenu("Email System")
        .addItem("📧 Test Email System", "testEmailSystem")
        .addItem("📊 Check Email Status", "checkEmailQuota")
        .addSeparator()
        .addItem("🔄 Resend OTP to Individual Faculty", "sendIndividualOTP")
        .addItem("📋 Send OTP to All Faculty", "sendEmails")
        .addItem("👥 View Faculty List", "showFacultyList")
        .addSeparator()
        .addItem("🔍 Diagnose Email Issues", "troubleshootEmailIssues")
        .addItem("🔧 Repair Email System", "repairEmailSystem")
        .addSeparator()
        .addItem("🔐 Authorize Email Permissions", "authorizeEmailPermissions")
        .addItem("📋 View Email Log", "viewEmailLog"))
      .addSeparator()
      .addSubMenu(ui.createMenu("Deadline Management")
        .addItem("📅 Initialize Deadlines Sheet", "initializeFacultyDeadlines")
        .addItem("⏰ Set Faculty Deadline", "setFacultyDeadlineUI")
        .addItem("📊 View Deadline Status", "viewDeadlineStatus")
        .addItem("🔍 Check Overdue Faculties", "checkOverdueFaculties")
        .addItem("📧 Send Test Reminder", "sendTestReminder"))
      .addToUi();
      
    console.log("Menus created successfully");
  } catch (error) {
    console.log("Menu creation failed (expected in web app context):", error.message);
  }
}

function setFacultyDeadlineUI() {
  try {
    if (!isUIAvailable()) return;
    
    const ui = SpreadsheetApp.getUi();
    const emailResponse = ui.prompt('Set Faculty Deadline', 'Enter faculty email:', ui.ButtonSet.OK_CANCEL);
    
    if (emailResponse.getSelectedButton() === ui.Button.OK) {
      const email = emailResponse.getResponseText().trim();
      const dateResponse = ui.prompt('Set Deadline Date', 'Enter due date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
      
      if (dateResponse.getSelectedButton() === ui.Button.OK) {
        const dateStr = dateResponse.getResponseText().trim();
        const result = setFacultyDeadline(email, dateStr);
        
        if (result.success) {
          ui.alert('Success', result.message, ui.ButtonSet.OK);
        } else {
          ui.alert('Error', result.message, ui.ButtonSet.OK);
        }
      }
    }
  } catch (error) {
    console.error('Error in setFacultyDeadlineUI:', error);
  }
}

function viewDeadlineStatus() {
  try {
    if (!isUIAvailable()) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deadlinesSheet = ss.getSheetByName(FACULTY_DEADLINES_SHEET);
    
    if (deadlinesSheet) {
      deadlinesSheet.activate();
    } else {
      SpreadsheetApp.getUi().alert('Sheet Not Found', 'Please initialize the deadlines sheet first.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    console.error('Error viewing deadline status:', error);
  }
}

function sendTestReminder() {
  try {
    if (!isUIAvailable()) return;
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Test Reminder', 'Enter email to send test reminder:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const email = response.getResponseText().trim();
      const stats = getFacultyStudentStats(email);
      const reminderSent = sendDeadlineReminder(email, new Date(), stats, 1);
      
      if (reminderSent) {
        ui.alert('Success', 'Test reminder sent successfully!', ui.ButtonSet.OK);
      } else {
        ui.alert('Error', 'Failed to send test reminder', ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    console.error('Error sending test reminder:', error);
  }
}

function isUIAvailable() {
  try {
    const ui = SpreadsheetApp.getUi();
    return ui && typeof ui.alert === 'function';
  } catch (error) {
    return false;
  }
}

function formatEditRequestsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
  if (!editSheet) return;
  
  editSheet.clearConditionalFormatRules();
  const startRow = 3;
  const lastRow = editSheet.getLastRow();
  if (lastRow < startRow) return;
  const numRows = lastRow - startRow + 1;
  const range = editSheet.getRange(startRow, 1, numRows, editSheet.getLastColumn());
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$H${startRow}="Pending"`).setBackground("#FFF9C4").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$H${startRow}="Approved"`).setBackground("#C8E6C9").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$H${startRow}="Disapproved"`).setBackground("#FFCDD2").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$H${startRow}="Completed"`).setBackground("#E8F5E8").setRanges([range]).build()
  ];
  
  editSheet.setConditionalFormatRules(rules);
  console.log("Conditional formatting applied to EditRequests sheet with Completed status.");
}

// ===== ANALYSIS FUNCTION =====
function analyse() {
  try {
    console.log("🔍 Starting comprehensive paper analysis...");
    
    if (!isUIAvailable()) {
      console.log("UI not available - analysis will run in console mode");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!mainSheet) {
      throw new Error(`Main data sheet "${SHEET_NAME}" not found.`);
    }
    
    const data = mainSheet.getDataRange().getValues();
    if (data.length < 2) {
      const message = "No student paper data found to analyze.";
      safeUIAlert("No Data", message);
      return message;
    }
    
    // Get column indices
    const headers = data[0].map(h => h.toString().trim());
    const emailCol = headers.indexOf(COL_EXAMINER_EMAIL);
    const marksCol = headers.indexOf(COL_MARKS);
    const verifiedCol = headers.indexOf(COL_VERIFIED);
    const regdNoCol = headers.indexOf(COL_REGD_NO);
    const nameCol = headers.indexOf(COL_STUDENT_NAME);
    const paperCodeCol = headers.indexOf(COL_PAPER_CODE);
    const campusCol = headers.indexOf("Campus");
    
    console.log("Column mapping:", { emailCol, marksCol, verifiedCol, regdNoCol, nameCol, paperCodeCol, campusCol });
    
    if ([emailCol, marksCol, verifiedCol, regdNoCol, paperCodeCol].includes(-1)) {
      throw new Error(`Required columns missing. Found headers: ${headers.join(', ')}`);
    }
    
    // Perform analysis
    const analysisResults = performPaperAnalysis(data, {
      emailCol, marksCol, verifiedCol, regdNoCol, nameCol, paperCodeCol, campusCol
    });
    
    // Create analysis sheet
    const analysisSheet = createOrUpdateAnalysisSheet(ss, analysisResults);
    
    // Show results
    showAnalysisResults(analysisResults);
    
    // Activate analysis sheet
    if (analysisSheet) {
      analysisSheet.activate();
    }
    
    return `Analysis complete. Results available in "${ANALYSIS_SHEET_NAME}" sheet.`;
    
  } catch (error) {
    console.error("❌ Analysis failed:", error, error.stack);
    const errorMsg = `Analysis failed: ${error.message}`;
    safeUIAlert("Analysis Error", errorMsg);
    return errorMsg;
  }
}

function showAvailableFacultyEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    
    if (emailColIndex === -1) {
      console.log("Email column not found");
      return;
    }
    
    const emails = new Set();
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailColIndex];
      if (email && typeof email === 'string' && email.includes('@')) {
        emails.add(email.trim());
      }
    }
    
    console.log("📧 Available faculty emails in your system:");
    Array.from(emails).forEach((email, index) => {
      console.log(`${index + 1}. ${email}`);
    });
    
    return Array.from(emails);
    
  } catch (error) {
    console.error("Error getting faculty emails:", error);
    return [];
  }
}

// Function to map column headers to indices
function getDeadlineColumnIndices(headers) {
  try {
    console.log("📋 Mapping deadline sheet columns...");
    console.log("Available headers:", headers);
    
    const columnMap = {
      facultyEmail: headers.indexOf("Faculty Email"),
      dueDate: headers.indexOf("Due Date"),
      totalStudents: headers.indexOf("Total Students"),
      studentsWithMarks: headers.indexOf("Students With Marks"),
      studentsVerified: headers.indexOf("Students Verified"),
      completionStatus: headers.indexOf("Completion Status"),
      lastReminderSent: headers.indexOf("Last Reminder Sent"),
      reminderCount: headers.indexOf("Reminder Count"),
      completionConfirmedDate: headers.indexOf("Completion Confirmed Date"),
      notes: headers.indexOf("Notes")
    };
    
    console.log("📊 Column mapping:", columnMap);
    
    // Check for missing critical columns
    const missingColumns = [];
    if (columnMap.facultyEmail === -1) missingColumns.push("Faculty Email");
    if (columnMap.dueDate === -1) missingColumns.push("Due Date");
    if (columnMap.completionStatus === -1) missingColumns.push("Completion Status");
    
    if (missingColumns.length > 0) {
      console.error("❌ Missing required columns:", missingColumns);
      throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
    }
    
    console.log("✅ Column mapping successful");
    return columnMap;
    
  } catch (error) {
    console.error("❌ Error mapping columns:", error);
    throw error;
  }
}

function checkOverdueFaculties() {
  try {
    console.log("🔍 Starting header-based overdue faculty check...");
    
    const deadlinesSheet = initializeFacultyDeadlines();
    const data = deadlinesSheet.getDataRange().getValues();
    
    console.log(`📊 Sheet data: ${data.length} rows total`);
    
    if (data.length <= 1) {
      console.log("ℹ️ No faculty deadlines data found");
      return { success: true, message: "No deadlines to check" };
    }
    
    // ✅ Map columns by header names
    const headers = data[0].map(h => h.toString().trim());
    console.log("📋 Sheet headers:", headers);
    
    const cols = getDeadlineColumnIndices(headers);
    console.log("🗂️ Using column mapping:", cols);
    
    // ✅ VALIDATION: Check if Faculty Email column exists
    if (cols.facultyEmail === -1) {
      console.error("❌ CRITICAL: Faculty Email column not found!");
      console.error("❌ Available headers:", headers);
      return {
        success: false,
        message: "Faculty Email column not found in deadlines sheet",
        availableHeaders: headers
      };
    }
    
    console.log(`✅ Faculty Email column found at index: ${cols.facultyEmail}`);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let remindersSet = 0;
    let completionChecks = 0;
    let errors = [];
    let processedCount = 0;
    let skippedCount = 0;
    
    console.log(`🔄 Processing ${data.length - 1} data rows with header-based access...`);

    for (let i = 1; i < data.length; i++) {
      console.log(`\n--- PROCESSING ROW ${i} ---`);
      
      const row = data[i];
      
      if (!row || row.length === 0) {
        console.log(`❌ Row ${i}: Empty row, skipping`);
        skippedCount++;
        continue;
      }
      
      // ✅ FIXED: Comprehensive email validation with early exit
      const emailRaw = row[cols.facultyEmail];
      console.log(`📧 Row ${i}: Raw email from column ${cols.facultyEmail}: "${emailRaw}" (${typeof emailRaw})`);
      
      // Check if email is null, undefined, or empty
      if (emailRaw === null || emailRaw === undefined || emailRaw === '') {
        console.log(`❌ Row ${i}: Faculty Email is null/undefined/empty, skipping`);
        skippedCount++;
        continue;
      }
      
      // Convert to string and validate
      let email;
      try {
        email = emailRaw.toString().trim();
        if (email === '' || email === 'undefined' || email === 'null') {
          console.log(`❌ Row ${i}: Email is invalid after string conversion: "${email}", skipping`);
          skippedCount++;
          continue;
        }
        
        // Validate email format
        if (!isValidEmail(email)) {
          console.log(`❌ Row ${i}: Invalid email format: "${email}", skipping`);
          skippedCount++;
          continue;
        }
        
        console.log(`✅ Row ${i}: Valid email: "${email}"`);
      } catch (stringError) {
        console.log(`❌ Row ${i}: Failed to process email:`, stringError);
        skippedCount++;
        continue;
      }
      
      // ✅ FIXED: Get other row data with validation
      const dueDateValue = row[cols.dueDate];
      const completionStatus = row[cols.completionStatus] || "Pending";
      const lastReminderSent = row[cols.lastReminderSent] ? new Date(row[cols.lastReminderSent]) : null;
      const reminderCount = parseInt(row[cols.reminderCount]) || 0;
      
      console.log(`📅 Row ${i}: Due date: "${dueDateValue}", Status: "${completionStatus}"`);
      
      // ✅ ENHANCED: Due date validation
      if (!dueDateValue) {
        console.log(`❌ Row ${i}: No due date for ${email}, skipping`);
        skippedCount++;
        continue;
      }
      
      let dueDate;
      try {
        dueDate = new Date(dueDateValue);
        if (isNaN(dueDate.getTime())) {
          console.log(`❌ Row ${i}: Invalid due date for ${email}: ${dueDateValue}`);
          errors.push(`Invalid due date for ${email}: ${dueDateValue}`);
          continue;
        }
        dueDate.setHours(0, 0, 0, 0);
        console.log(`✅ Row ${i}: Valid due date: ${dueDate.toLocaleDateString()}`);
      } catch (dateError) {
        console.log(`❌ Row ${i}: Date parsing error for ${email}:`, dateError);
        errors.push(`Date parsing error for ${email}: ${dateError.message}`);
        continue;
      }
      
      // Skip if already confirmed
      if (completionStatus === "Confirmed") {
        console.log(`✅ Row ${i}: ${email} already confirmed, skipping`);
        continue;
      }
      
      // Get faculty statistics
      console.log(`📊 Row ${i}: Getting stats for ${email}...`);
      const stats = getFacultyStudentStats(email);
      console.log(`📊 Row ${i}: Stats:`, stats);
      
      // ✅ UPDATE: Use column mapping for updates
      try {
        deadlinesSheet.getRange(i + 1, cols.totalStudents + 1, 1, 3)
                     .setValues([[stats.totalStudents, stats.studentsWithMarks, stats.studentsVerified]]);
        console.log(`✅ Row ${i}: Updated stats in sheet`);
      } catch (updateError) {
        console.log(`❌ Row ${i}: Failed to update stats:`, updateError);
      }
      
      // Check completion status
      if (stats.totalStudents > 0 && stats.studentsWithMarks === stats.totalStudents && completionStatus === "Pending") {
        deadlinesSheet.getRange(i + 1, cols.completionStatus + 1).setValue("Ready for Confirmation");
        console.log(`✅ Row ${i}: ${email} ready for confirmation`);
        completionChecks++;
      }
      
      // Check if reminder is needed
      const isOverdue = today > dueDate;
      const needsReminder = completionStatus !== "Ready for Confirmation" && completionStatus !== "Confirmed";
      
      if (isOverdue && needsReminder) {
        const daysPastDue = Math.ceil((today - dueDate) / (1000 * 60 * 60 * 24));
        const daysSinceLastReminder = lastReminderSent ? 
          Math.floor((today - lastReminderSent) / (1000 * 60 * 60 * 24)) : 999;
        
        console.log(`📅 Row ${i}: ${daysPastDue} days overdue, ${daysSinceLastReminder} days since last reminder`);
        
        if (daysSinceLastReminder >= 1) {
          console.log(`📧 Row ${i}: SENDING reminder to: "${email}"`);
          
          // ✅ FIXED: Final validation before calling sendDeadlineReminder
          if (!email || typeof email !== 'string' || email.trim() === '') {
            console.error(`❌ Row ${i}: CRITICAL - email validation failed before sending: "${email}"`);
            errors.push(`Email validation failed for row ${i}: ${email}`);
            skippedCount++;
            continue;
          }
          
          // ✅ FIXED: Don't create another email variable, use the validated one
          console.log(`📧 Row ${i}: Calling sendDeadlineReminder with email: "${email}"`);
          const reminderSent = sendDeadlineReminder(email, dueDate, stats, reminderCount + 1);
          
          if (reminderSent) {
            try {
              deadlinesSheet.getRange(i + 1, cols.lastReminderSent + 1, 1, 2)
                           .setValues([[today, reminderCount + 1]]);
              remindersSet++;
              console.log(`✅ Row ${i}: Reminder sent successfully to ${email}`);
            } catch (updateError) {
              console.error(`❌ Row ${i}: Failed to update reminder data in sheet:`, updateError);
              errors.push(`Update failed for ${email}: ${updateError.message}`);
            }
          } else {
            console.log(`❌ Row ${i}: Failed to send reminder to ${email}`);
            errors.push(`Failed to send reminder to ${email}`);
          }
        } else {
          console.log(`⏰ Row ${i}: Too soon for reminder to ${email} (last sent ${daysSinceLastReminder} days ago)`);
        }
      }
      
      processedCount++;
    }
    
    // Check if all faculties have completed
    try {
      checkAllFacultiesCompleted();
    } catch (completionError) {
      console.error("❌ Error checking completion:", completionError);
    }
    
    console.log(`\n📊 FINAL SUMMARY:`);
    console.log(`  - Processed: ${processedCount}, Skipped: ${skippedCount}`);
    console.log(`  - Reminders sent: ${remindersSet}, Ready for confirmation: ${completionChecks}`);
    console.log(`  - Errors: ${errors.length}`);
    
    return {
      success: true,
      message: `Check completed: ${processedCount} processed, ${remindersSet} reminders sent`,
      details: { processedCount, skippedCount, remindersSet, completionChecks, errors }
    };
    
  } catch (error) {
    console.error("🚨 CRITICAL ERROR in checkOverdueFaculties:", error);
    return { success: false, message: error.message };
  }
}


// Fixed sendDeadlineReminder function with OTP retrieval
function sendDeadlineReminder(email, dueDate = null, stats = null, reminderCount = null) {
  let subject = "Assessment Deadline Reminder";
  
  try {
    console.log(`📧 sendDeadlineReminder called with: email="${email}", dueDate=${dueDate}, stats=${stats}, reminderCount=${reminderCount}`);
    
    // ✅ VALIDATION: Check email parameter
    if (!email || typeof email !== 'string') {
      console.error('sendDeadlineReminder: Invalid email:', email);
      return false;
    }
    
    // ✅ AUTO-GET: Due date if not provided
    if (!dueDate) {
      console.log('sendDeadlineReminder: Auto-getting due date...');
      // Get from deadlines sheet
      const deadlinesSheet = initializeFacultyDeadlines();
      const data = deadlinesSheet.getDataRange().getValues();
      const headers = data[0];
      const cols = getDeadlineColumnIndices(headers);
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowEmail = row[cols.facultyEmail];
        if (rowEmail && rowEmail.toString().trim().toLowerCase() === email.toLowerCase()) {
          dueDate = new Date(row[cols.dueDate]);
          break;
        }
      }
      
      if (!dueDate) {
        console.error('sendDeadlineReminder: No due date found for:', email);
        return false;
      }
    }
    
    // ✅ AUTO-GET: Stats if not provided
    if (!stats) {
      console.log('sendDeadlineReminder: Auto-getting faculty stats...');
      stats = getFacultyStudentStats(email);
    }
    
    // ✅ AUTO-GET: Reminder count if not provided
    if (!reminderCount) {
      console.log('sendDeadlineReminder: Using default reminder count...');
      reminderCount = 1;
    }
    
    // ✅ FIX: GET THE FACULTY'S OTP WITH ENHANCED LOGIC
    let otp = null;
    try {
      console.log(`🔍 Retrieving OTP for: ${email}`);
      
      const storedOTPs = PropertiesService.getScriptProperties().getProperty('OTPs');
      if (!storedOTPs) {
        console.warn('🚨 No stored OTPs found - generating new OTP for this faculty');
        // Generate new OTP for this faculty member
        otp = generateOTP();
        
        // Store it
        const newOTPMap = {};
        newOTPMap[email.toLowerCase()] = otp;
        PropertiesService.getScriptProperties().setProperty('OTPs', JSON.stringify(newOTPMap));
        console.log(`✅ Generated and stored new OTP for ${email}: ${otp}`);
        
      } else {
        const otpMap = JSON.parse(storedOTPs);
        console.log(`📋 Available emails in OTP storage: ${Object.keys(otpMap).join(', ')}`);
        
        const emailLower = email.toLowerCase();
        console.log(`🔍 Looking for: ${emailLower}`);
        
        // Try multiple variations to find the OTP
        otp = otpMap[email] || otpMap[emailLower];
        
        if (!otp) {
          // Try to find email with different case
          const foundEmail = Object.keys(otpMap).find(storedEmail => 
            storedEmail.toLowerCase() === emailLower
          );
          if (foundEmail) {
            otp = otpMap[foundEmail];
            console.log(`✅ Found OTP using case-insensitive match: ${foundEmail} -> ${otp}`);
          }
        }
        
        if (!otp) {
          console.warn(`⚠️ No existing OTP found for ${email}, generating new one`);
          // Generate new OTP for this faculty member
          otp = generateOTP();
          
          // Add to existing OTP map
          otpMap[emailLower] = otp;
          PropertiesService.getScriptProperties().setProperty('OTPs', JSON.stringify(otpMap));
          console.log(`✅ Generated and stored new OTP for ${email}: ${otp}`);
        } else {
          console.log(`✅ Found existing OTP for ${email}: ${otp}`);
        }
      }
    } catch (otpError) {
      console.error('❌ Error retrieving/generating OTP:', otpError);
      // Fallback: generate a new OTP
      otp = generateOTP();
      console.log(`🔄 Fallback: Generated new OTP: ${otp}`);
    }
    
    // Ensure we have a valid OTP
    if (!otp || typeof otp !== 'string' || otp.length !== 6) {
      otp = generateOTP();
      console.log(`🔄 Final fallback: Generated OTP: ${otp}`);
    }
    
    // ✅ VALIDATION: Check all parameters are now valid
    if (!(dueDate instanceof Date) || isNaN(dueDate.getTime())) {
      console.error('sendDeadlineReminder: Invalid due date:', dueDate);
      return false;
    }
    
    if (!stats || typeof stats !== 'object') {
      console.error('sendDeadlineReminder: Invalid stats:', stats);
      return false;
    }
    
    // ✅ FIX: Normalize both dates to midnight for accurate calculation
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const normalizedDueDate = new Date(dueDate);
    normalizedDueDate.setHours(0, 0, 0, 0);
    
    const daysPastDue = Math.floor((today - normalizedDueDate) / (1000 * 60 * 60 * 24));
    
    // Only send if actually overdue
    if (daysPastDue <= 0) {
      console.log(`sendDeadlineReminder: Not overdue yet for ${email}`);
      return false;
    }
    
    subject = `Urgent: Assessment Deadline Passed - Action Required [Reminder ${reminderCount}]`;
    
    const message = `Dear Faculty Member,

Sai Ram,

This is an urgent reminder regarding your assessment deadline.

📊 DEADLINE STATUS:
- Due Date: ${normalizedDueDate.toLocaleDateString('en-GB')}
- Days Overdue: ${daysPastDue} day${daysPastDue === 1 ? '' : 's'}
- Reminder Count: ${reminderCount}

📈 YOUR PROGRESS:
- Total Students Assigned: ${stats.totalStudents || 0}
- Marks Uploaded: ${stats.studentsWithMarks || 0}/${stats.totalStudents || 0} (${stats.completionPercentage || 0}%)
- Students Verified: ${stats.studentsVerified || 0}/${stats.totalStudents || 0} (${stats.verificationPercentage || 0}%)

⚠️ IMMEDIATE ACTION REQUIRED:
${(stats.studentsWithMarks || 0) < (stats.totalStudents || 0) ? 
  `• Upload marks for ${(stats.totalStudents || 0) - (stats.studentsWithMarks || 0)} remaining students` : 
  '• All marks uploaded - Please verify and confirm completion'}

🌐 Faculty Portal: ${WEB_APP_URL}
🔒 One Time Password (OTP): ${otp}

📝 Steps to Complete:
1. Login to the faculty portal using your OTP
2. Upload marks for all assigned students
3. Verify all marks using the "Verify" button
4. Click "Confirm Completion" when all marks are uploaded

⏰ Please complete this immediately to avoid further escalation.

For technical support, contact the examination committee.

Best regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning

---
This is an automated reminder. Please complete your assessment to stop these notifications.`;

    MailApp.sendEmail(email, subject, message);
    logEmailActivity("DEADLINE_REMINDER", email, subject, "SUCCESS");
    
    console.log(`✅ Reminder sent successfully to ${email}`);
    return true;
    
  } catch (error) {
    console.error(`❌ Error sending reminder to ${email}:`, error);
    logEmailActivity("DEADLINE_REMINDER", email || "undefined", subject, `FAILED: ${error.message}`);
    return false;
  }
}

// Function to confirm faculty completion
function confirmFacultyCompletion(email) {
  try {
    console.log(`🔄 Confirming completion for faculty: ${email}`);
    
    const deadlinesSheet = initializeFacultyDeadlines();
    const data = deadlinesSheet.getDataRange().getValues();
    
    // Find faculty row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        // Update completion status
        deadlinesSheet.getRange(i + 1, 6, 1, 2).setValues([["Confirmed", new Date()]]);
        
        console.log(`✅ Faculty ${email} completion confirmed`);
        
        // Send confirmation email to faculty
        sendCompletionConfirmationEmail(email);
        
        return { success: true, message: "Completion confirmed successfully" };
      }
    }
    
    return { success: false, message: "Faculty not found in deadlines sheet" };
  } catch (error) {
    console.error("Error confirming completion:", error);
    return { success: false, message: error.message };
  }
}

// Function to send completion confirmation email
function sendCompletionConfirmationEmail(email) {
  try {
    // ✅ FIX: Add validation before sending email
    console.log(`📧 sendCompletionConfirmationEmail called with: "${email}" (type: ${typeof email})`);
    
    // Check if email is undefined, null, or empty
    if (!email || typeof email !== 'string' || email.trim() === '') {
      console.error(`❌ Invalid email parameter: "${email}" (type: ${typeof email})`);
      console.error(`❌ Cannot send completion confirmation - invalid recipient`);
      return false; // Return false to indicate failure
    }
    
    const trimmedEmail = email.trim();
    
    // Validate email format
    if (!isValidEmail(trimmedEmail)) {
      console.error(`❌ Invalid email format: "${trimmedEmail}"`);
      return false;
    }
    
    console.log(`✅ Email validation passed for: "${trimmedEmail}"`);
    
    const subject = "Assessment Completion Confirmed - Thank You!";
    const message = `Dear Faculty Member,

Sai Ram,

Thank you for confirming the completion of your assessment!

✅ COMPLETION CONFIRMED:
- All marks have been uploaded and verified
- Assessment deadline requirements fulfilled
- No further reminder notifications will be sent

📊 STATUS: Assessment Complete ✅

Your contribution to the examination process is greatly appreciated.

Best regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning`;

    // Send email only if validation passed
    MailApp.sendEmail(trimmedEmail, subject, message);
    console.log(`✅ Completion confirmation email sent successfully to: ${trimmedEmail}`);
    
    logEmailActivity("COMPLETION_CONFIRMATION", trimmedEmail, subject, "SUCCESS");
    return true; // Return true to indicate success
    
  } catch (error) {
    console.error(`❌ Error sending completion confirmation to "${email}":`, error);
    
    // Log with more context about the email parameter
    const emailInfo = email ? `"${email}" (${typeof email})` : 'undefined/null';
    logEmailActivity("COMPLETION_CONFIRMATION", emailInfo, "Completion Confirmation", `FAILED: ${error.message}`);
    return false; // Return false to indicate failure
  }
}

function debugEmailCalls() {
  console.log("🔍 Searching for functions that call sendCompletionConfirmationEmail...");
  
  // Check the most likely culprit - confirmFacultyCompletion
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deadlinesSheet = ss.getSheetByName(FACULTY_DEADLINES_SHEET);
    
    if (!deadlinesSheet) {
      console.log("❌ FacultyDeadlines sheet not found");
      return;
    }
    
    const data = deadlinesSheet.getDataRange().getValues();
    console.log(`📊 FacultyDeadlines sheet has ${data.length} rows`);
    
    if (data.length > 0) {
      console.log("📋 Headers:", data[0]);
      
      // Show sample email data
      for (let i = 1; i < Math.min(data.length, 3); i++) {
        console.log(`Row ${i} emails: Column 0="${data[i][0]}", Full row:`, data[i]);
      }
    }
    
    console.log("\n🧪 Test: What happens if we call confirmFacultyCompletion with undefined?");
    // This will show us the exact path the undefined email takes
    
  } catch (error) {
    console.error("Debug failed:", error);
  }
}

// Function to check if all faculties have completed
function checkAllFacultiesCompleted() {
  try {
    console.log("🔍 Checking if all faculties have completed assessments...");
    
    // ✅ FIX 1: Check completion based on MAIN SHEET data, not deadlines sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!mainSheet) {
      console.log("❌ Main sheet not found");
      return;
    }
    
    const data = mainSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log("❌ No data in main sheet");
      return;
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    // ✅ FIX 2: Analyze actual completion based on marks verification
    const facultyStats = new Map();
    let totalStudents = 0;
    let totalVerified = 0;
    
    // Collect faculty-wise statistics
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      const verified = row[columnIndices.verified] ? row[columnIndices.verified].toString().trim() : '';
      
      if (!email) continue;
      
      if (!facultyStats.has(email)) {
        facultyStats.set(email, {
          totalStudents: 0,
          verifiedStudents: 0
        });
      }
      
      const faculty = facultyStats.get(email);
      faculty.totalStudents++;
      totalStudents++;
      
      if (verified === '✅') {
        faculty.verifiedStudents++;
        totalVerified++;
      }
    }
    
    // ✅ FIX 3: Check if ALL faculty have completed ALL their students
    let allFacultiesCompleted = true;
    const incompleteFaculties = [];
    
    for (const [email, stats] of facultyStats) {
      if (stats.verifiedStudents < stats.totalStudents) {
        allFacultiesCompleted = false;
        incompleteFaculties.push({
          email: email,
          completed: stats.verifiedStudents,
          total: stats.totalStudents,
          percentage: Math.round((stats.verifiedStudents / stats.totalStudents) * 100)
        });
      }
    }
    
    const totalFaculties = facultyStats.size;
    const completedFaculties = totalFaculties - incompleteFaculties.length;
    
    console.log(`📊 Completion Analysis:`);
    console.log(`  - Total Faculties: ${totalFaculties}`);
    console.log(`  - Completed Faculties: ${completedFaculties}`);
    console.log(`  - Incomplete Faculties: ${incompleteFaculties.length}`);
    console.log(`  - Total Students: ${totalStudents}`);
    console.log(`  - Total Verified: ${totalVerified}`);
    console.log(`  - Overall Completion: ${totalStudents > 0 ? Math.round((totalVerified / totalStudents) * 100) : 0}%`);
    
    // ✅ FIX 4: Only send notification if ALL faculties are 100% complete
    if (allFacultiesCompleted && totalFaculties > 0 && totalStudents > 0) {
      console.log("🎉 All faculties have completed their assessments!");
      
      // ✅ FIX 5: Use a PERMANENT key that doesn't reset daily
      const completionKey = `admin_notified_all_complete_${totalFaculties}_${totalStudents}_${totalVerified}`;
      const alreadyNotified = PropertiesService.getScriptProperties().getProperty(completionKey);
      
      if (!alreadyNotified) {
        console.log("✅ Sending one-time completion notification to admin...");
        
        // Generate and send the completion report
        const success = sendAdminCompletionNotification(totalFaculties);
        
        if (success) {
          // ✅ FIX 6: Set permanent flag to prevent ANY future duplicate notifications
          PropertiesService.getScriptProperties().setProperty(completionKey, new Date().toISOString());
          
          // Also set a simple flag for easier checking
          PropertiesService.getScriptProperties().setProperty('ASSESSMENT_COMPLETION_NOTIFIED', 'true');
          
          console.log("✅ Completion notification sent and permanent flag set");
          console.log("🔒 Future duplicate notifications are now prevented");
        } else {
          console.log("❌ Failed to send completion notification");
        }
      } else {
        console.log("ℹ️ Admin completion notification already sent - skipping duplicate");
        console.log(`📅 Originally notified: ${alreadyNotified}`);
      }
    } else {
      console.log("⏳ Not all faculties have completed their assessments yet");
      if (incompleteFaculties.length > 0) {
        console.log("📋 Incomplete faculties:");
        incompleteFaculties.forEach(faculty => {
          console.log(`  - ${faculty.email}: ${faculty.completed}/${faculty.total} (${faculty.percentage}%)`);
        });
      }
    }
    
  } catch (error) {
    console.error("❌ Error checking faculties completion:", error);
  }
}

function hasCompletionNotificationBeenSent() {
  try {
    const flag = PropertiesService.getScriptProperties().getProperty('ASSESSMENT_COMPLETION_NOTIFIED');
    return flag === 'true';
  } catch (error) {
    console.error("Error checking notification flag:", error);
    return false;
  }
}

function resetCompletionNotification() {
  try {
    console.log("🔄 Resetting completion notification flags...");
    
    // Get all stored properties
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    // Remove all completion-related keys
    Object.keys(properties).forEach(key => {
      if (key.startsWith('admin_notified_all_complete_') || key === 'ASSESSMENT_COMPLETION_NOTIFIED') {
        PropertiesService.getScriptProperties().deleteProperty(key);
        console.log(`🗑️ Removed property: ${key}`);
      }
    });
    
    console.log("✅ Completion notification flags reset - notification can be sent again");
    return "Completion notification flags reset successfully";
    
  } catch (error) {
    console.error("❌ Error resetting completion notification:", error);
    return `Error: ${error.message}`;
  }
}

// Function to send admin notification when all faculties complete
function sendAdminCompletionNotification(totalFaculties) {
  try {
    // ✅ EXTRA SAFETY: Check if notification was already sent
    if (hasCompletionNotificationBeenSent()) {
      console.log("🚫 Completion notification already sent - preventing duplicate");
      return false;
    }
    
    console.log("🎉 Generating admin completion notification with paper-wise breakdown...");
    
    // Get paper-wise completion data
    const paperWiseReport = generatePaperWiseCompletionReport();
    
    if (!paperWiseReport || paperWiseReport.paperCodes.length === 0) {
      console.error("❌ Failed to generate paper-wise report");
      // Fallback to original notification
      return sendOriginalAdminNotification(totalFaculties);
    }
    
    const subject = "🎉 Assessment Completion Report - All Paper Codes Completed";
    const message = generateEnhancedCompletionEmail(paperWiseReport, totalFaculties);
    
    // Send email to admin
    MailApp.sendEmail(ADMIN_EMAIL, subject, message);
    logEmailActivity("ADMIN_COMPLETION_NOTIFICATION", ADMIN_EMAIL, subject, "SUCCESS");
    
    console.log("✅ Enhanced admin completion notification sent with paper-wise breakdown");
    return true;
    
  } catch (error) {
    console.error("❌ Error sending enhanced admin notification:", error);
    // Fallback to original notification
    return sendOriginalAdminNotification(totalFaculties);
  }
}

function generatePaperWiseCompletionReport() {
  try {
    console.log("📊 Generating paper-wise completion report...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!mainSheet) {
      console.error("❌ Main sheet not found");
      return null;
    }
    
    const data = mainSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error("❌ No data in main sheet");
      return null;
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    // Group data by paper code
    const paperCodeMap = new Map();
    let totalStudentsOverall = 0;
    let totalWithMarksOverall = 0;
    let totalVerifiedOverall = 0;
    const facultySet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const paperCode = row[columnIndices.paperCode] ? row[columnIndices.paperCode].toString().trim() : 'Not Specified';
      const paperTitle = row[columnIndices.paperTitle] ? row[columnIndices.paperTitle].toString().trim() : 'Title Not Available';
      const examinerEmail = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      const examinerName = row[columnIndices.examiner] ? row[columnIndices.examiner].toString().trim() : 'Not Assigned';
      const regdNo = row[columnIndices.regdNo] ? row[columnIndices.regdNo].toString().trim() : '';
      const studentName = row[columnIndices.name] ? row[columnIndices.name].toString().trim() : '';
      const marks = row[columnIndices.marks];
      const verified = row[columnIndices.verified] ? row[columnIndices.verified].toString().trim() : '';
      const programme = row[columnIndices.programme] ? row[columnIndices.programme].toString().trim() : 'Programme Not Specified';
      const semester = row[columnIndices.semester] ? row[columnIndices.semester].toString().trim() : 'Semester Not Specified';
      const exam = row[columnIndices.exam] ? row[columnIndices.exam].toString().trim() : 'Exam Not Specified';
      const credits = row[columnIndices.credits] ? row[columnIndices.credits].toString().trim() : 'Credits Not Specified';
      const maxMarks = row[columnIndices.maxMarks] ? Number(row[columnIndices.maxMarks]) || 100 : 100;
      
      // Skip rows without essential data
      if (!regdNo || !examinerEmail) continue;
      
      // Add to overall faculty count
      if (examinerEmail) {
        facultySet.add(examinerEmail);
      }
      
      // Initialize paper code entry if not exists
      if (!paperCodeMap.has(paperCode)) {
        paperCodeMap.set(paperCode, {
          paperCode: paperCode,
          paperTitle: paperTitle,
          programme: programme,
          semester: semester,
          exam: exam,
          credits: credits,
          maxMarks: maxMarks,
          totalStudents: 0,
          studentsWithMarks: 0,
          studentsVerified: 0,
          facultyEmails: new Set(),
          facultyNames: new Set(),
          students: []
        });
      }
      
      const paperData = paperCodeMap.get(paperCode);
      
      // Add faculty info
      if (examinerEmail) {
        paperData.facultyEmails.add(examinerEmail);
      }
      if (examinerName && examinerName !== 'Not Assigned') {
        paperData.facultyNames.add(examinerName);
      }
      
      // Count students
      paperData.totalStudents++;
      totalStudentsOverall++;
      
      // Count students with marks
      const hasMarks = marks !== null && marks !== undefined && marks !== '';
      if (hasMarks) {
        paperData.studentsWithMarks++;
        totalWithMarksOverall++;
      }
      
      // Count verified students
      if (verified === '✅') {
        paperData.studentsVerified++;
        totalVerifiedOverall++;
      }
      
      // Store student info
      paperData.students.push({
        regdNo: regdNo,
        name: studentName,
        marks: marks,
        verified: verified,
        hasMarks: hasMarks,
        examinerEmail: examinerEmail,
        examinerName: examinerName
      });
    }
    
    // Convert map to array and calculate completion rates
    const paperCodes = Array.from(paperCodeMap.values()).map(paper => {
      const completionRate = paper.totalStudents > 0 ? 
        Math.round((paper.studentsWithMarks / paper.totalStudents) * 100) : 0;
      const verificationRate = paper.totalStudents > 0 ? 
        Math.round((paper.studentsVerified / paper.totalStudents) * 100) : 0;
      
      return {
        ...paper,
        completionRate: completionRate,
        verificationRate: verificationRate,
        facultyCount: paper.facultyEmails.size,
        facultyEmailsList: Array.from(paper.facultyEmails),
        facultyNamesList: Array.from(paper.facultyNames),
        isComplete: completionRate === 100 && verificationRate === 100
      };
    });
    
    // Sort by paper code
    paperCodes.sort((a, b) => a.paperCode.localeCompare(b.paperCode));
    
    const overallCompletionRate = totalStudentsOverall > 0 ? 
      Math.round((totalWithMarksOverall / totalStudentsOverall) * 100) : 0;
    const overallVerificationRate = totalStudentsOverall > 0 ? 
      Math.round((totalVerifiedOverall / totalStudentsOverall) * 100) : 0;
    
    const report = {
      paperCodes: paperCodes,
      summary: {
        totalPaperCodes: paperCodes.length,
        totalFaculty: facultySet.size,
        totalStudents: totalStudentsOverall,
        totalWithMarks: totalWithMarksOverall,
        totalVerified: totalVerifiedOverall,
        overallCompletionRate: overallCompletionRate,
        overallVerificationRate: overallVerificationRate,
        fullyCompletedPapers: paperCodes.filter(p => p.isComplete).length,
        timestamp: new Date()
      }
    };
    
    console.log("✅ Paper-wise report generated:", {
      paperCount: report.summary.totalPaperCodes,
      facultyCount: report.summary.totalFaculty,
      studentCount: report.summary.totalStudents,
      fullyCompleted: report.summary.fullyCompletedPapers
    });
    
    return report;
    
  } catch (error) {
    console.error("❌ Error generating paper-wise report:", error);
    return null;
  }
}

function generateEnhancedCompletionEmail(paperWiseReport, totalFaculties) {
  const { paperCodes, summary } = paperWiseReport;
  const timestamp = summary.timestamp.toLocaleString();
  
  let emailBody = `📊 COMPREHENSIVE ASSESSMENT COMPLETION REPORT

Dear Administrator,

Sai Ram,

🎉 Excellent news! All faculty assessments have been completed across all paper codes.

═══════════════════════════════════════════════════════════════

📈 OVERALL SUMMARY:
═══════════════════════════════════════════════════════════════
• Report Generated: ${timestamp}
• Total Paper Codes: ${summary.totalPaperCodes}
• Total Faculty Members: ${summary.totalFaculty}
• Total Students: ${summary.totalStudents}
• Overall Completion Rate: ${summary.overallCompletionRate}% ✅
• Overall Verification Rate: ${summary.overallVerificationRate}% ✅
• Fully Completed Papers: ${summary.fullyCompletedPapers}/${summary.totalPaperCodes}

═══════════════════════════════════════════════════════════════

📋 PAPER-WISE DETAILED BREAKDOWN:
═══════════════════════════════════════════════════════════════

`;

  paperCodes.forEach((paper, index) => {
    const paperNumber = index + 1;
    const statusIcon = paper.isComplete ? '✅' : (paper.completionRate >= 80 ? '🟡' : '🔴');
    
    emailBody += `
${paperNumber}. ${statusIcon} PAPER CODE: ${paper.paperCode}
   ────────────────────────────────────────────────────────────
   📖 Title: ${paper.paperTitle}
   📚 Programme: ${paper.programme}
   📅 Semester: ${paper.semester} | Exam: ${paper.exam}
   💯 Max Marks: ${paper.maxMarks} | Credits: ${paper.credits}
   
   👥 Faculty Assignment:
   • Faculty Count: ${paper.facultyCount}`;

    if (paper.facultyNamesList.length > 0) {
      emailBody += `
   • Examiner(s): ${paper.facultyNamesList.join(', ')}`;
    }
    
    emailBody += `
   • Email(s): ${paper.facultyEmailsList.join(', ')}
   
   📊 Student Statistics:
   • Total Students: ${paper.totalStudents}
   • With Marks: ${paper.studentsWithMarks}/${paper.totalStudents} (${paper.completionRate}%)
   • Verified: ${paper.studentsVerified}/${paper.totalStudents} (${paper.verificationRate}%)
   • Status: ${paper.isComplete ? 'FULLY COMPLETED ✅' : 'IN PROGRESS 🔄'}
   
`;
  });

  emailBody += `
═══════════════════════════════════════════════════════════════

🎯 COMPLETION ANALYSIS:
═══════════════════════════════════════════════════════════════`;

  // Show completion status breakdown
  const fullyCompleted = paperCodes.filter(p => p.isComplete);
  const partiallyCompleted = paperCodes.filter(p => !p.isComplete && p.completionRate > 0);
  const notStarted = paperCodes.filter(p => p.completionRate === 0);

  emailBody += `
• Fully Completed Papers: ${fullyCompleted.length}`;
  if (fullyCompleted.length > 0) {
    emailBody += `
  └─ ${fullyCompleted.map(p => p.paperCode).join(', ')}`;
  }

  if (partiallyCompleted.length > 0) {
    emailBody += `
• Partially Completed: ${partiallyCompleted.length}
  └─ ${partiallyCompleted.map(p => `${p.paperCode} (${p.completionRate}%)`).join(', ')}`;
  }

  if (notStarted.length > 0) {
    emailBody += `
• Not Started: ${notStarted.length}
  └─ ${notStarted.map(p => p.paperCode).join(', ')}`;
  }

  emailBody += `

═══════════════════════════════════════════════════════════════

🚀 NEXT STEPS:
═══════════════════════════════════════════════════════════════
✅ All paper codes have achieved 100% faculty completion
✅ All marks have been uploaded and verified
✅ All faculty members have confirmed completion
✅ Assessment phase is officially complete across all papers
✅ System is ready for result compilation and processing
═══════════════════════════════════════════════════════════════

📈 SYSTEM STATUS: ALL CLEAR ACROSS ALL PAPER CODES ✅

The enhanced assessment management system has successfully tracked and 
ensured completion of all faculty assessments across ${summary.totalPaperCodes} 
different paper codes with ${summary.totalFaculty} faculty members and 
${summary.totalStudents} students.


Best regards,
Enhanced Assessment Management System
Sri Sathya Sai Institute of Higher Learning

───────────────────────────────────────────────────────────────
This is an automated system notification with detailed analytics.
Report ID: COMP-${Date.now()}
Generated: ${timestamp}
───────────────────────────────────────────────────────────────`;

  return emailBody;
}

function sendOriginalAdminNotification(totalFaculties) {
  try {
    const subject = "🎉 All Faculty Assessments Completed - System Notification";
    
    const message = `Dear Administrator,

Sai Ram,

Excellent news! All faculty members have completed their assessments.

📊 COMPLETION SUMMARY:
- Total Faculty Members: ${totalFaculties}
- Completed Assessments: ${totalFaculties}
- Completion Rate: 100% ✅
- Notification Date: ${new Date().toLocaleDateString()}

🎯 NEXT STEPS:
- All marks have been uploaded and verified
- Faculty members have confirmed completion
- Assessment phase is officially complete
- You may proceed with result compilation

📈 SYSTEM STATUS: All Clear ✅

The assessment management system has successfully tracked and ensured completion of all faculty assessments.

Best regards,
Assessment Management System
Sri Sathya Sai Institute of Higher Learning

---
This is an automated system notification.`;

    MailApp.sendEmail(ADMIN_EMAIL, subject, message);
    logEmailActivity("ADMIN_COMPLETION_NOTIFICATION", ADMIN_EMAIL, subject, "SUCCESS");
    
    console.log("✅ Original admin completion notification sent as fallback");
    
  } catch (error) {
    console.error("❌ Error sending original admin notification:", error);
  }
}


function performPaperAnalysis(data, columnIndices) {
  console.log("📊 Performing comprehensive paper analysis...");
  
  const stats = {
    totalPapers: data.length - 1,
    papersWithMarks: 0,
    papersVerified: 0,
    papersWithoutMarks: 0,
    averageMarks: 0,
    highestMarks: 0,
    lowestMarks: 100,
    totalMarksSum: 0,
    validMarksCount: 0
  };
  
  const facultyProgress = {};
  const campusStats = {};
  const marksDistribution = {
    'A+ (90-100)': 0,
    'A (80-89)': 0,
    'B+ (70-79)': 0,
    'B (60-69)': 0,
    'C+ (50-59)': 0,
    'C (40-49)': 0,
    'F (0-39)': 0
  };
  
  const hasCampusData = columnIndices.campusCol !== -1;
  
  // Process each paper record
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[columnIndices.emailCol] ? row[columnIndices.emailCol].toString().trim() : '';
    const marks = row[columnIndices.marksCol];
    const verified = row[columnIndices.verifiedCol] ? row[columnIndices.verifiedCol].toString().trim() : '';
    const regdNo = row[columnIndices.regdNoCol] ? row[columnIndices.regdNoCol].toString().trim() : '';
    const campus = hasCampusData && row[columnIndices.campusCol] ? 
                   row[columnIndices.campusCol].toString().trim() : 'Unknown Campus';
    
    // Skip rows without essential data
    if (!email || !regdNo) continue;
    
    // Initialize faculty progress tracking
    if (!facultyProgress[email]) {
      facultyProgress[email] = {
        email: email,
        totalPapers: 0,
        papersWithMarks: 0,
        papersVerified: 0,
        completionRate: 0,
        verificationRate: 0
      };
    }
    
    // Initialize campus stats
    if (!campusStats[campus]) {
      campusStats[campus] = {
        name: campus,
        totalPapers: 0,
        papersWithMarks: 0,
        papersVerified: 0,
        completionRate: 0,
        verificationRate: 0,
        averageMarks: 0,
        totalMarksSum: 0,
        validMarksCount: 0
      };
    }
    
    facultyProgress[email].totalPapers++;
    campusStats[campus].totalPapers++;
    
    // Analyze marks and verification status
    const hasMarks = marks !== null && marks !== undefined && marks !== '';
    const isVerified = verified === '✅';
    
    if (hasMarks) {
      stats.papersWithMarks++;
      facultyProgress[email].papersWithMarks++;
      campusStats[campus].papersWithMarks++;
      
      const numericMarks = Number(marks);
      if (!isNaN(numericMarks) && numericMarks >= 0 && numericMarks <= 100) {
        stats.totalMarksSum += numericMarks;
        stats.validMarksCount++;
        stats.highestMarks = Math.max(stats.highestMarks, numericMarks);
        stats.lowestMarks = Math.min(stats.lowestMarks, numericMarks);
        
        // Campus marks statistics
        campusStats[campus].totalMarksSum += numericMarks;
        campusStats[campus].validMarksCount++;
        
        // Grade distribution
        if (numericMarks >= 90) marksDistribution['A+ (90-100)']++;
        else if (numericMarks >= 80) marksDistribution['A (80-89)']++;
        else if (numericMarks >= 70) marksDistribution['B+ (70-79)']++;
        else if (numericMarks >= 60) marksDistribution['B (60-69)']++;
        else if (numericMarks >= 50) marksDistribution['C+ (50-59)']++;
        else if (numericMarks >= 40) marksDistribution['C (40-49)']++;
        else marksDistribution['F (0-39)']++;
      }
    } else {
      stats.papersWithoutMarks++;
    }
    
    if (isVerified) {
      stats.papersVerified++;
      facultyProgress[email].papersVerified++;
      campusStats[campus].papersVerified++;
    }
  }
  
  // Calculate derived statistics
  stats.averageMarks = stats.validMarksCount > 0 ? 
    Math.round((stats.totalMarksSum / stats.validMarksCount) * 100) / 100 : 0;
  
  if (stats.validMarksCount === 0) {
    stats.lowestMarks = 0;
    stats.highestMarks = 0;
  }
  
  // Calculate faculty completion rates
  Object.keys(facultyProgress).forEach(email => {
    const faculty = facultyProgress[email];
    faculty.completionRate = faculty.totalPapers > 0 ? 
      Math.round((faculty.papersWithMarks / faculty.totalPapers) * 100) : 0;
    faculty.verificationRate = faculty.totalPapers > 0 ? 
      Math.round((faculty.papersVerified / faculty.totalPapers) * 100) : 0;
  });
  
  // Calculate campus statistics
  Object.keys(campusStats).forEach(campusName => {
    const campus = campusStats[campusName];
    campus.completionRate = campus.totalPapers > 0 ? 
      Math.round((campus.papersWithMarks / campus.totalPapers) * 100) : 0;
    campus.verificationRate = campus.totalPapers > 0 ? 
      Math.round((campus.papersVerified / campus.totalPapers) * 100) : 0;
    campus.averageMarks = campus.validMarksCount > 0 ? 
      Math.round((campus.totalMarksSum / campus.validMarksCount) * 100) / 100 : 0;
  });
  
  return {
    stats,
    facultyProgress,
    campusStats,
    marksDistribution,
    hasCampusData,
    timestamp: new Date().toLocaleString()
  };
}

function createOrUpdateAnalysisSheet(ss, results) {
  try {
    // Delete existing analysis sheet if it exists
    const existingSheet = ss.getSheetByName(ANALYSIS_SHEET_NAME);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // Create new analysis sheet
    const analysisSheet = ss.insertSheet(ANALYSIS_SHEET_NAME);
    
    // Build analysis report
    let currentRow = 1;
    
    // Title
    analysisSheet.getRange(currentRow, 1).setValue("📊 PAPER ANALYSIS REPORT");
    analysisSheet.getRange(currentRow, 1).setFontSize(16).setFontWeight("bold");
    currentRow += 2;
    
    // Timestamp
    analysisSheet.getRange(currentRow, 1).setValue(`Generated: ${results.timestamp}`);
    analysisSheet.getRange(currentRow, 1).setFontStyle("italic");
    currentRow += 2;
    
    // Overall Statistics
    analysisSheet.getRange(currentRow, 1).setValue("📈 OVERALL STATISTICS");
    analysisSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#e3f2fd");
    currentRow++;
    
    const overallStats = [
      ["Total Papers", results.stats.totalPapers],
      ["Papers with Marks", results.stats.papersWithMarks],
      ["Papers Verified", results.stats.papersVerified],
      ["Papers Pending", results.stats.papersWithoutMarks],
      ["Average Marks", results.stats.averageMarks],
      ["Highest Marks", results.stats.highestMarks],
      ["Lowest Marks", results.stats.lowestMarks]
    ];
    
    analysisSheet.getRange(currentRow, 1, overallStats.length, 2).setValues(overallStats);
    currentRow += overallStats.length + 2;
    
    // Faculty Progress
    analysisSheet.getRange(currentRow, 1).setValue("👥 FACULTY PROGRESS");
    analysisSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#e8f5e8");
    currentRow++;
    
    const facultyHeaders = [["Faculty Email", "Total Papers", "With Marks", "Verified", "Completion %", "Verification %"]];
    analysisSheet.getRange(currentRow, 1, 1, 6).setValues(facultyHeaders);
    analysisSheet.getRange(currentRow, 1, 1, 6).setFontWeight("bold");
    currentRow++;
    
    const facultyData = Object.values(results.facultyProgress).map(faculty => [
      faculty.email,
      faculty.totalPapers,
      faculty.papersWithMarks,
      faculty.papersVerified,
      faculty.completionRate + "%",
      faculty.verificationRate + "%"
    ]);
    
    if (facultyData.length > 0) {
      analysisSheet.getRange(currentRow, 1, facultyData.length, 6).setValues(facultyData);
      currentRow += facultyData.length + 2;
    }
    
    // Campus Statistics (if available)
    if (results.hasCampusData && Object.keys(results.campusStats).length > 0) {
      analysisSheet.getRange(currentRow, 1).setValue("🏫 CAMPUS STATISTICS");
      analysisSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#fff3e0");
      currentRow++;
      
      const campusHeaders = [["Campus", "Total Papers", "With Marks", "Verified", "Completion %", "Avg Marks"]];
      analysisSheet.getRange(currentRow, 1, 1, 6).setValues(campusHeaders);
      analysisSheet.getRange(currentRow, 1, 1, 6).setFontWeight("bold");
      currentRow++;
      
      const campusData = Object.values(results.campusStats).map(campus => [
        campus.name,
        campus.totalPapers,
        campus.papersWithMarks,
        campus.papersVerified,
        campus.completionRate + "%",
        campus.averageMarks
      ]);
      
      analysisSheet.getRange(currentRow, 1, campusData.length, 6).setValues(campusData);
      currentRow += campusData.length + 2;
    }
    
    // Marks Distribution
    analysisSheet.getRange(currentRow, 1).setValue("📊 MARKS DISTRIBUTION");
    analysisSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#f3e5f5");
    currentRow++;
    
    const distributionHeaders = [["Grade", "Count", "Percentage"]];
    analysisSheet.getRange(currentRow, 1, 1, 3).setValues(distributionHeaders);
    analysisSheet.getRange(currentRow, 1, 1, 3).setFontWeight("bold");
    currentRow++;
    
    const distributionData = Object.entries(results.marksDistribution).map(([grade, count]) => {
      const percentage = results.stats.validMarksCount > 0 ? 
        Math.round((count / results.stats.validMarksCount) * 100) : 0;
      return [grade, count, percentage + "%"];
    });
    
    analysisSheet.getRange(currentRow, 1, distributionData.length, 3).setValues(distributionData);
    
    // Format the sheet
    analysisSheet.autoResizeColumns(1, 6);
    
    return analysisSheet;
    
  } catch (error) {
    console.error("Error creating analysis sheet:", error);
    return null;
  }
}

function showAnalysisResults(results) {
  const completionRate = Math.round((results.stats.papersVerified / results.stats.totalPapers) * 100);
  const uploadRate = Math.round((results.stats.papersWithMarks / results.stats.totalPapers) * 100);
  const facultyCount = Object.keys(results.facultyProgress).length;
  const campusCount = Object.keys(results.campusStats).length;
  
  let message = `📊 Paper Analysis Complete!\n\n`;
  message += `📈 Overall Results:\n`;
  message += `• Total Papers: ${results.stats.totalPapers}\n`;
  message += `• Upload Progress: ${uploadRate}% (${results.stats.papersWithMarks}/${results.stats.totalPapers})\n`;
  message += `• Verification Progress: ${completionRate}% (${results.stats.papersVerified}/${results.stats.totalPapers})\n`;
  message += `• Average Marks: ${results.stats.averageMarks}\n\n`;
  message += `👥 Faculty & Campus:\n`;
  message += `• Total Faculty: ${facultyCount}\n`;
  if (results.hasCampusData) {
    message += `• Total Campuses: ${campusCount}\n`;
  }
  message += `\n📊 Detailed analysis available in "${ANALYSIS_SHEET_NAME}" sheet.`;
  
  safeUIAlert("Analysis Complete", message);
  console.log("Analysis results shown to user");
}

// ===== DEBUG FUNCTIONS =====
function debugSheetHeaders() {
  try {
    console.log("=== DEBUG SHEET HEADERS ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      const message = `Sheet "${SHEET_NAME}" not found.`;
      console.error(message);
      safeUIAlert("Debug Headers", message);
      return message;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      const message = "Sheet has no data.";
      console.error(message);
      safeUIAlert("Debug Headers", message);
      return message;
    }
    
    const headers = data[0].map(h => h.toString().trim());
    console.log("Current headers:", headers);
    
    const requiredHeaders = [
      COL_REGD_NO, COL_STUDENT_NAME, COL_PAPER_CODE, COL_PAPER_TITLE,
      COL_EXAMINER_EMAIL, COL_MARKS, COL_VERIFIED, "Semester", "Exam", "Credits", "Max Marks"
    ];
    
    const missingHeaders = requiredHeaders.filter(h => !headers.includes(h));
    const extraHeaders = headers.filter(h => !requiredHeaders.includes(h) && h !== "" && h !== "Sl. No." && h !== "Programme Name" && h !== "Campus" && h !== "Examiner");
    
    let debugMessage = `📋 SHEET HEADERS DEBUG\n\n`;
    debugMessage += `📊 Sheet: ${SHEET_NAME}\n`;
    debugMessage += `📈 Total columns: ${headers.length}\n`;
    debugMessage += `📋 Total rows: ${data.length}\n\n`;
    
    debugMessage += `✅ Current Headers:\n`;
    headers.forEach((header, index) => {
      debugMessage += `${index + 1}. "${header}"\n`;
    });
    
    debugMessage += `\n🎯 Required Headers:\n`;
    requiredHeaders.forEach((header, index) => {
      const found = headers.includes(header);
      debugMessage += `${index + 1}. "${header}" ${found ? '✅' : '❌'}\n`;
    });
    
    if (missingHeaders.length > 0) {
      debugMessage += `\n❌ Missing Headers (${missingHeaders.length}):\n`;
      missingHeaders.forEach((header, index) => {
        debugMessage += `${index + 1}. "${header}"\n`;
      });
    }
    
    if (extraHeaders.length > 0) {
      debugMessage += `\nℹ️ Additional Headers (${extraHeaders.length}):\n`;
      extraHeaders.forEach((header, index) => {
        debugMessage += `${index + 1}. "${header}"\n`;
      });
    }
    
    console.log("Headers debug completed");
    safeUIAlert("Sheet Headers Debug", debugMessage);
    
    return debugMessage;
    
  } catch (error) {
    const errorMsg = `Debug failed: ${error.message}`;
    console.error("Debug error:", error);
    safeUIAlert("Debug Error", errorMsg);
    return errorMsg;
  }
}

// ===== ADDITIONAL HELPER FUNCTIONS =====
function openFacultyPortal() {
  try {
    console.log("=== OPENING FACULTY PORTAL ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot open portal via menu");
      return "UI not available in this context.";
    }
    
    const message = `🌐 Faculty Portal Access\n\nPortal URL: ${WEB_APP_URL}\n\n💡 Instructions:\n1. Copy the URL above\n2. Open it in a new browser tab\n3. Use faculty email and OTP to login\n\nFor OTP issues:\n• Use "Send OTP to All Faculty" for bulk emails\n• Use "Resend OTP to Individual Faculty" for specific faculty\n\nNote: Make sure faculty have received their OTP emails before accessing the portal.`;
    
    safeUIAlert("Faculty Portal Information", message);
    console.log("Faculty portal information displayed");
    
    return `Portal URL provided: ${WEB_APP_URL}`;
    
  } catch (error) {
    console.error("Error in openFacultyPortal:", error);
    safeUIAlert("Portal Error", `Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

function testApprovalSystem() {
  try {
    console.log("=== TESTING APPROVAL SYSTEM ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot test approval system via menu");
      return "UI not available in this context.";
    }
    
    const testResults = [];
    
    // Test 1: Check EditRequests sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (editSheet) {
      testResults.push("✅ EditRequests sheet found");
      
      const data = editSheet.getDataRange().getValues();
      testResults.push(`✅ EditRequests has ${data.length - 1} rows of data`);
      
      if (data.length > 1) {
        const headers = data[0];
        const requiredCols = ["Status", "Faculty Email", "Student Regd No", "Unlock Until"];
        const missingCols = requiredCols.filter(col => !headers.includes(col));
        
        if (missingCols.length === 0) {
          testResults.push("✅ All required columns present");
        } else {
          testResults.push(`❌ Missing columns: ${missingCols.join(', ')}`);
        }
      }
    } else {
      testResults.push("❌ EditRequests sheet not found");
    }
    
    // Test 2: Check main sheet
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    if (mainSheet) {
      testResults.push(`✅ Main sheet "${SHEET_NAME}" found`);
      
      const mainData = mainSheet.getDataRange().getValues();
      testResults.push(`✅ Main sheet has ${mainData.length - 1} rows of data`);
    } else {
      testResults.push(`❌ Main sheet "${SHEET_NAME}" not found`);
    }
    
    // Test 3: Test approval workflow functions
    try {
      // Test unlock system
      const testEmail = "test@example.com";
      const unlockedStudents = getUnlockedRegdNos(testEmail);
      testResults.push(`✅ Unlock system functional - found ${unlockedStudents.size} unlocked students for test email`);
    } catch (unlockError) {
      testResults.push(`❌ Unlock system error: ${unlockError.message}`);
    }
    
    const testReport = `🧪 APPROVAL SYSTEM TEST RESULTS\n\n${testResults.join('\n')}\n\n📅 Test completed: ${new Date().toLocaleString()}`;
    
    console.log("Approval system test completed");
    safeUIAlert("Approval System Test", testReport);
    
    return testReport;
    
  } catch (error) {
    const errorMsg = `Approval system test failed: ${error.message}`;
    console.error("Test error:", error);
    safeUIAlert("Test Error", errorMsg);
    return errorMsg;
  }
}

function testUnlockSystemWithUI() {
  try {
    console.log("=== TESTING UNLOCK SYSTEM WITH UI ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot test unlock system via menu");
      return "UI not available in this context.";
    }
    
    const testResults = testUnlockSystem();
    safeUIAlert("Unlock System Test Results", testResults);
    
    return testResults;
    
  } catch (error) {
    const errorMsg = `Unlock system test failed: ${error.message}`;
    console.error("Test error:", error);
    safeUIAlert("Test Error", errorMsg);
    return errorMsg;
  }
}

function testUnlockSystem() {
  console.log("=== COMPREHENSIVE UNLOCK SYSTEM TEST ===");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      console.error("❌ EditRequests sheet not found");
      return "EditRequests sheet not found";
    }
    
    console.log("✅ EditRequests sheet found");
    
    // Test 1: Check sheet structure
    const data = editSheet.getDataRange().getValues();
    console.log("Sheet data rows:", data.length);
    
    if (data.length > 0) {
      const headers = data[0];
      console.log("Headers:", headers);
      
      // Check required columns
      const requiredColumns = ["Faculty Email", "Student Regd No", "Status", "Unlock Until"];
      const columnIndices = {};
      const missingColumns = [];
      
      requiredColumns.forEach(col => {
        const index = headers.indexOf(col);
        columnIndices[col] = index;
        if (index === -1) {
          missingColumns.push(col);
        }
      });
      
      console.log("Column indices:", columnIndices);
      
      if (missingColumns.length > 0) {
        console.error("❌ Missing columns:", missingColumns);
        return `Missing columns: ${missingColumns.join(', ')}`;
      }
      
      console.log("✅ All required columns found");
      
      // Test 2: Check for approved requests
      let approvedCount = 0;
      let expiredCount = 0;
      let activeCount = 0;
      const now = new Date();
      
      console.log("Current time for comparison:", now.toLocaleString());
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = row[columnIndices["Status"]];
        const facultyEmail = row[columnIndices["Faculty Email"]];
        const regdNo = row[columnIndices["Student Regd No"]];
        const unlockUntil = row[columnIndices["Unlock Until"]];
        
        if (status === "Approved") {
          approvedCount++;
          console.log(`Row ${i + 1}: Approved request - Faculty: ${facultyEmail}, Student: ${regdNo}, Unlock until: ${unlockUntil}`);
          
          if (unlockUntil) {
            const unlockDate = new Date(unlockUntil);
            if (!isNaN(unlockDate.getTime())) {
              if (unlockDate > now) {
                activeCount++;
                console.log(`  ✅ ACTIVE unlock until ${unlockDate.toLocaleString()}`);
              } else {
                expiredCount++;
                console.log(`  ⏰ EXPIRED unlock (expired at ${unlockDate.toLocaleString()})`);
              }
            } else {
              console.log(`  ❌ INVALID unlock date: ${unlockUntil}`);
            }
          } else {
            console.log(`  ❌ NO unlock date set`);
          }
        }
      }
      
      console.log("\n📊 UNLOCK SUMMARY:");
      console.log(`  Total approved requests: ${approvedCount}`);
      console.log(`  Active unlocks: ${activeCount}`);
      console.log(`  Expired unlocks: ${expiredCount}`);
      
      // Test 3: Test specific unlock function
      if (activeCount > 0) {
        console.log("\nTesting getUnlockedRegdNos with real data...");
        
        // Find a real faculty email from approved requests
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const status = row[columnIndices["Status"]];
          const facultyEmail = row[columnIndices["Faculty Email"]];
          
          if (status === "Approved" && facultyEmail) {
            console.log(`Testing getUnlockedRegdNos with email: ${facultyEmail}`);
            const unlockedStudents = getUnlockedRegdNos(facultyEmail);
            console.log(`  Result: ${unlockedStudents.size} unlocked students`);
            console.log(`  Students:`, Array.from(unlockedStudents));
            break;
          }
        }
      }
      
      const summary = `Test completed. Approved: ${approvedCount}, Active: ${activeCount}, Expired: ${expiredCount}`;
      console.log("\n" + summary);
      return summary;
      
    } else {
      console.log("⚠️ EditRequests sheet is empty");
      return "EditRequests sheet is empty";
    }
    
  } catch (error) {
    console.error("❌ Test failed:", error);
    return `Test failed: ${error.message}`;
  }
}

// ===== ADDITIONAL APPROVAL FUNCTIONS =====
function approveSecondEdit(requestId) {
  console.log("=== APPROVE SECOND EDIT REQUEST ===");
  console.log("Request ID:", requestId);
  
  if (!requestId) {
    console.error("No request ID provided");
    return { success: false, message: "No request ID provided" };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = ss.getSheetByName(EDIT_REQUESTS_SHEET);
    
    if (!editSheet) {
      throw new Error(`Sheet "${EDIT_REQUESTS_SHEET}" not found.`);
    }
    
    const data = editSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the request by ID
    let requestRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) {
        requestRow = i;
        break;
      }
    }
    
    if (requestRow === -1) {
      console.error("Request not found:", requestId);
      return { success: false, message: "Request not found" };
    }
    
    // Update the request
    const statusCol = headers.indexOf("Status");
    const unlockUntilCol = headers.indexOf("Unlock Until");
    const actionNotesCol = headers.indexOf("Action Notes");
    
    if ([statusCol, unlockUntilCol, actionNotesCol].includes(-1)) {
      console.error("Required columns not found");
      return { success: false, message: "Sheet structure error" };
    }
    
    const unlockUntil = new Date();
    unlockUntil.setHours(unlockUntil.getHours() + 48);
    
    editSheet.getRange(requestRow + 1, statusCol + 1).setValue("Approved");
    editSheet.getRange(requestRow + 1, unlockUntilCol + 1).setValue(unlockUntil);
    editSheet.getRange(requestRow + 1, actionNotesCol + 1).setValue(`Approved via URL on ${new Date().toLocaleString()}`);
    
    console.log("✅ Second edit request approved successfully");
    return { success: true, message: "Second edit request approved" };
    
  } catch (error) {
    console.error("Error in approveSecondEdit:", error);
    return { success: false, message: error.message };
  }
}

// ===== INDIVIDUAL OTP FUNCTIONALITY =====
function sendIndividualOTP() {
  try {
    console.log("=== INDIVIDUAL OTP REQUEST STARTED ===");
    
    if (!isUIAvailable()) {
      console.log("UI not available - cannot send individual OTP via menu");
      return "UI not available in this context.";
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Step 1: Get faculty email from admin
    const emailResponse = safeUIPrompt(
      "Resend OTP to Faculty Member", 
      "Enter the faculty member's institutional email address:"
    );
    
    if (!emailResponse || emailResponse.getSelectedButton() !== ui.Button.OK) {
      safeUIToast("OTP resend cancelled by admin.");
      return "Cancelled by admin.";
    }
    
    const facultyEmail = emailResponse.getResponseText().trim();
    console.log("Admin requested OTP for:", facultyEmail);
    
    // Step 2: Validate email format
    if (!facultyEmail || !isValidEmail(facultyEmail)) {
      safeUIAlert("Invalid Email", "Please enter a valid email address.");
      return "Invalid email format.";
    }
    
    // Step 3: Check if faculty exists in system
    const facultyExists = validateFacultyExists(facultyEmail);
    if (!facultyExists.exists) {
      safeUIAlert(
        "Faculty Not Found", 
        `Email "${facultyEmail}" is not found in the faculty database.\n\n` +
        `${facultyExists.totalFaculty} faculty members are currently in the system.\n\n` +
        `Please verify the email address or add the faculty member to the main sheet first.`
      );
      return "Faculty not found in system.";
    }
    
    // Step 4: Generate fresh OTP
    const newOTP = generateOTP();
    console.log("Generated new OTP for:", facultyEmail);
    
    // Step 5: Update stored OTPs
    const updateResult = updateStoredOTP(facultyEmail, newOTP);
    if (!updateResult.success) {
      safeUIAlert("Storage Error", `Failed to store OTP: ${updateResult.message}`);
      return "OTP storage failed.";
    }
    
    // Step 6: Send email with new OTP
    const emailResult = sendOTPEmail(facultyEmail, newOTP);
    
    // Step 7: Show result to admin
    if (emailResult.success) {
      safeUIAlert(
        "✅ OTP Sent Successfully!", 
        `New OTP has been sent to: ${facultyEmail}\n\n` +
        `OTP: ${newOTP}\n` +
        `Faculty can now use this OTP to access the portal.\n\n` +
        `Email quota remaining: ${MailApp.getRemainingDailyQuota()}`
      );
      console.log("Individual OTP sent successfully to:", facultyEmail);
      return `OTP sent successfully to ${facultyEmail}`;
    } else {
      safeUIAlert(
        "❌ Email Sending Failed", 
        `Failed to send OTP to: ${facultyEmail}\n\n` +
        `Error: ${emailResult.message}\n\n` +
        `Please check email permissions and quota.`
      );
      console.error("Failed to send individual OTP:", emailResult.message);
      return `Email sending failed: ${emailResult.message}`;
    }
    
  } catch (error) {
    console.error("Error in sendIndividualOTP:", error, error.stack);
    safeUIAlert("System Error", `An error occurred: ${error.message}\n\nPlease try again or contact IT support.`);
    return `Error: ${error.message}`;
  }
}

// Helper function to validate if faculty exists in the main sheet
function validateFacultyExists(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { exists: false, message: `Sheet "${SHEET_NAME}" not found.`, totalFaculty: 0 };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { exists: false, message: "No faculty data found in sheet.", totalFaculty: 0 };
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    
    if (emailColIndex === -1) {
      return { exists: false, message: `Column '${COL_EXAMINER_EMAIL}' not found.`, totalFaculty: 0 };
    }
    
    const trimmedEmail = email.trim().toLowerCase();
    const uniqueFacultyEmails = new Set();
    let facultyFound = false;
    
    // Check all rows for faculty emails
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailColIndex];
      if (rowEmail && typeof rowEmail === 'string') {
        const rowEmailTrimmed = rowEmail.trim().toLowerCase();
        if (isValidEmail(rowEmailTrimmed)) {
          uniqueFacultyEmails.add(rowEmailTrimmed);
          if (rowEmailTrimmed === trimmedEmail) {
            facultyFound = true;
          }
        }
      }
    }
    
    return { 
      exists: facultyFound, 
      message: facultyFound ? "Faculty found in system." : "Faculty not found in system.",
      totalFaculty: uniqueFacultyEmails.size,
      allFacultyEmails: Array.from(uniqueFacultyEmails)
    };
    
  } catch (error) {
    console.error("Error in validateFacultyExists:", error);
    return { exists: false, message: `Validation error: ${error.message}`, totalFaculty: 0 };
  }
}

// Helper function to update stored OTP for specific faculty
function updateStoredOTP(email, newOTP) {
  try {
    // Get existing OTPs
    const storedOTPs = PropertiesService.getScriptProperties().getProperty('OTPs');
    let otpMap = {};
    
    if (storedOTPs) {
      try {
        otpMap = JSON.parse(storedOTPs);
      } catch (parseError) {
        console.warn("Failed to parse existing OTPs, creating new map:", parseError);
        otpMap = {};
      }
    }
    
    // Update with new OTP
    const trimmedEmail = email.trim();
    otpMap[trimmedEmail] = newOTP;
    
    // Store back
    PropertiesService.getScriptProperties().setProperty('OTPs', JSON.stringify(otpMap));
    
    console.log(`Updated OTP for ${trimmedEmail}. Total OTPs stored: ${Object.keys(otpMap).length}`);
    
    return { 
      success: true, 
      message: "OTP stored successfully.",
      totalOTPs: Object.keys(otpMap).length
    };
    
  } catch (error) {
    console.error("Error in updateStoredOTP:", error);
    return { success: false, message: error.message };
  }
}

// Helper function to send OTP email to individual faculty
function sendOTPEmail(email, otp) {
  try {
    // Check email quota
    const emailQuota = MailApp.getRemainingDailyQuota();
    if (emailQuota < 1) {
      return { success: false, message: `Insufficient email quota (${emailQuota} remaining)` };
    }
    
    const portalLink = WEB_APP_URL;
    const subject = "Faculty Portal Access - Your New One-Time Password [" + otp + "]";
    
    const message = `Dear Faculty Member,

Sai Ram

A new One-Time Password (OTP) has been generated for your Faculty Portal access as requested by the administrator.

Portal Access: ${portalLink}

Login Credentials:
- 📧 Email: ${email}
- 🔐 Your New One-Time Password (OTP): "${otp}"

Quick Login Steps:
1. Click the portal link above
2. Enter your SSSIHL email address
3. Input the new OTP: ${otp}
4. Access your student data

⚠️ Important Notes:
• This new OTP replaces any previous OTP you may have had
• Use this OTP for your next login session
• Keep this OTP secure and do not share it with others
• Contact the examination team if you need further assistance

If you did not request this new OTP, please contact the examination administrator immediately.

Best Regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning (Deemed to be University)
Prasanthi Nilayam - 515134
Sri Sathya Sai District, Andhra Pradesh
INDIA

---
This is an automated message. Please do not reply to this email.`;

    // Send the email
    MailApp.sendEmail(email, subject, message);
    
    // Log the activity
    logEmailActivity("INDIVIDUAL_OTP_EMAIL", email, subject, "SUCCESS");
    
    console.log("Individual OTP email sent successfully to:", email);
    return { success: true, message: "Email sent successfully" };
    
  } catch (error) {
    console.error("Error in sendOTPEmail:", error);
    logEmailActivity("INDIVIDUAL_OTP_EMAIL", email, subject || "OTP Email", `FAILED: ${error.message}`);
    return { success: false, message: error.message };
  }
}

// ===== EMAIL TESTING AND DIAGNOSTIC FUNCTIONS =====
function testEmailSystem() { 
  try {
    console.log("=== EMAIL SYSTEM TEST STARTED ===");
    
    const quota = MailApp.getRemainingDailyQuota();
    const testEmail = Session.getActiveUser().getEmail();
    
    if (quota < 1) {
      const message = `❌ Email test failed: No remaining quota (${quota})`;
      console.error(message);
      safeUIAlert("Email Test Failed", message);
      return message;
    }
    
    // Send test email
    const subject = 'Email System Test - ' + new Date().toLocaleString();
    const testMessage = `✅ Email System Test Successful

This is a test email to verify that the email system is working correctly.

Test Details:
- Time: ${new Date().toLocaleString()}
- Quota Before: ${quota}
- Test Email: ${testEmail}
- System Status: Operational

If you received this email, the system is functioning properly.

Best regards,
Faculty Portal Email System`;

    MailApp.sendEmail(testEmail, subject, testMessage);
    
    const newQuota = MailApp.getRemainingDailyQuota();
    const message = `✅ Email system test successful!\n\n📊 Details:\n• Test email sent to: ${testEmail}\n• Quota before: ${quota}\n• Quota after: ${newQuota}\n• System status: Operational`;
    
    console.log("Email test successful:", { testEmail, quotaBefore: quota, quotaAfter: newQuota });
    safeUIAlert("Email Test Results", message);
    
    logEmailActivity("EMAIL_TEST", testEmail, subject, "SUCCESS");
    
    return message;
  } catch (error) {
    const errorMsg = `❌ Email test failed: ${error.message}`;
    console.error("Email test error:", error);
    safeUIAlert("Email Test Failed", errorMsg);
    logEmailActivity("EMAIL_TEST", "system", "Email Test", `FAILED: ${error.message}`);
    return errorMsg;
  }
}

function checkEmailQuota() { 
  try {
    const quota = MailApp.getRemainingDailyQuota();
    const percentage = Math.round((quota / 100) * 100); // Assuming 100 daily limit
    
    let status = "✅ Good";
    let color = "green";
    
    if (quota < 10) {
      status = "🔴 Critical";
      color = "red";
    } else if (quota < 25) {
      status = "🟡 Low";
      color = "orange";
    }
    
    const message = `📊 Email Quota Status\n\n• Remaining emails: ${quota}\n• Status: ${status}\n• Last checked: ${new Date().toLocaleString()}\n\nNote: Quota resets daily at midnight GMT.`;
    
    console.log("Email quota check:", { quota, status });
    safeUIAlert("Email Quota Status", message);
    
    return { quota, status, timestamp: new Date().toLocaleString() };
  } catch (error) {
    const errorMsg = `❌ Failed to check email quota: ${error.message}`;
    console.error("Email quota check error:", error);
    safeUIAlert("Quota Check Failed", errorMsg);
    return { error: errorMsg };
  }
}

function troubleshootEmailIssues() { 
  try {
    console.log("=== EMAIL TROUBLESHOOTING STARTED ===");
    
    const issues = [];
    const solutions = [];
    
    // Check 1: Email quota
    const quota = MailApp.getRemainingDailyQuota();
    if (quota < 1) {
      issues.push("❌ No email quota remaining");
      solutions.push("Wait for quota reset (midnight GMT) or contact administrator");
    } else {
      issues.push(`✅ Email quota available: ${quota}`);
    }
    
    // Check 2: User permissions
    try {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        issues.push(`✅ User email accessible: ${userEmail}`);
      } else {
        issues.push("❌ Cannot access user email");
        solutions.push("Check script permissions and re-authorize");
      }
    } catch (permError) {
      issues.push("❌ Permission error accessing user email");
      solutions.push("Re-run authorization: Extensions > Apps Script > Run authorizationTest");
    }
    
    // Check 3: Sheet access
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        issues.push(`✅ Main sheet accessible: ${SHEET_NAME}`);
      } else {
        issues.push(`❌ Main sheet not found: ${SHEET_NAME}`);
        solutions.push("Create or rename sheet to match SHEET_NAME constant");
      }
    } catch (sheetError) {
      issues.push("❌ Cannot access spreadsheet");
      solutions.push("Check spreadsheet permissions and sheet names");
    }
    
    // Check 4: Properties Service
    try {
      PropertiesService.getScriptProperties().getProperty('test');
      issues.push("✅ Properties Service accessible");
    } catch (propError) {
      issues.push("❌ Properties Service error");
      solutions.push("Check script permissions for Properties Service");
    }
    
    // Compile report
    let report = "📋 EMAIL SYSTEM DIAGNOSTIC REPORT\n\n";
    report += "🔍 Issues Found:\n";
    issues.forEach(issue => report += `• ${issue}\n`);
    
    if (solutions.length > 0) {
      report += "\n💡 Recommended Solutions:\n";
      solutions.forEach(solution => report += `• ${solution}\n`);
    }
    
    report += `\n📅 Report generated: ${new Date().toLocaleString()}`;
    
    console.log("Email troubleshooting report:", report);
    safeUIAlert("Email System Diagnostics", report);
    
    return { issues, solutions, report };
    
  } catch (error) {
    const errorMsg = `❌ Troubleshooting failed: ${error.message}`;
    console.error("Troubleshooting error:", error);
    safeUIAlert("Diagnostic Error", errorMsg);
    return { error: errorMsg };
  }
}

function repairEmailSystem() { 
  try {
    console.log("=== EMAIL SYSTEM REPAIR STARTED ===");
    
    const repairs = [];
    let repairCount = 0;
    
    // Repair 1: Clear corrupted OTP data
    try {
      const storedOTPs = PropertiesService.getScriptProperties().getProperty('OTPs');
      if (storedOTPs) {
        JSON.parse(storedOTPs); // Test if valid JSON
        repairs.push("✅ OTP data is valid");
      } else {
        repairs.push("ℹ️ No OTP data found (normal for first run)");
      }
    } catch (otpError) {
      PropertiesService.getScriptProperties().deleteProperty('OTPs');
      repairs.push("🔧 Cleared corrupted OTP data");
      repairCount++;
    }
    
    // Repair 2: Create email log sheet if missing
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName("EmailLog");
      if (!logSheet) {
        logSheet = ss.insertSheet("EmailLog");
        logSheet.appendRow(["Timestamp", "Type", "Recipient", "Subject", "Status"]);
        logSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f0f0f0");
        repairs.push("🔧 Created EmailLog sheet");
        repairCount++;
      } else {
        repairs.push("✅ EmailLog sheet exists");
      }
    } catch (logError) {
      repairs.push("❌ Could not create EmailLog sheet: " + logError.message);
    }
    
    // Repair 3: Test email functionality
    try {
      const quota = MailApp.getRemainingDailyQuota();
      if (quota > 0) {
        repairs.push(`✅ Email quota available: ${quota}`);
      } else {
        repairs.push("⚠️ No email quota available");
      }
    } catch (emailError) {
      repairs.push("❌ Email system access error: " + emailError.message);
    }
    
    // Compile repair report
    let report = "🔧 EMAIL SYSTEM REPAIR REPORT\n\n";
    report += `📊 Repairs performed: ${repairCount}\n\n`;
    report += "🔍 Repair Details:\n";
    repairs.forEach(repair => report += `• ${repair}\n`);
    
    if (repairCount > 0) {
      report += "\n✅ System repairs completed. Try sending emails again.";
    } else {
      report += "\nℹ️ No repairs needed. System appears to be functioning normally.";
    }
    
    report += `\n📅 Repair completed: ${new Date().toLocaleString()}`;
    
    console.log("Email repair results:", { repairCount, repairs });
    safeUIAlert("Email System Repair", report);
    
    return { repairCount, repairs, report };
    
  } catch (error) {
    const errorMsg = `❌ Email repair failed: ${error.message}`;
    console.error("Email repair error:", error);
    safeUIAlert("Repair Failed", errorMsg);
    return { error: errorMsg };
  }
}

function authorizeEmailPermissions() { 
  try {
    console.log("=== EMAIL AUTHORIZATION TEST ===");
    
    // Test 1: Check basic MailApp access
    const quota = MailApp.getRemainingDailyQuota();
    console.log("✅ MailApp access successful, quota:", quota);
    
    // Test 2: Check user email access
    const userEmail = Session.getActiveUser().getEmail();
    console.log("✅ User email access successful:", userEmail);
    
    // Test 3: Check Properties Service
    PropertiesService.getScriptProperties().setProperty('authTest', 'success');
    const testValue = PropertiesService.getScriptProperties().getProperty('authTest');
    console.log("✅ Properties Service access successful:", testValue);
    
    // Clean up test property
    PropertiesService.getScriptProperties().deleteProperty('authTest');
    
    const message = `✅ Email authorization successful!\n\n📊 Authorization Details:\n• MailApp access: ✅ Granted\n• Email quota: ${quota}\n• User email: ${userEmail}\n• Properties Service: ✅ Granted\n\n🎉 All email permissions are properly configured.`;
    
    safeUIAlert("Email Authorization Status", message);
    return { success: true, quota, userEmail, message };
    
  } catch (error) {
    const errorMsg = `❌ Email authorization failed: ${error.message}\n\n💡 Solution:\n1. Go to Extensions > Apps Script\n2. Click the "Run" button on any function\n3. Grant all requested permissions\n4. Try again`;
    
    console.error("Email authorization error:", error);
    safeUIAlert("Authorization Required", errorMsg);
    return { success: false, error: error.message, message: errorMsg };
  }
}

function viewEmailLog() { 
  try {
    console.log("=== EMAIL LOG VIEWER ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("EmailLog");
    
    if (!logSheet) {
      const message = "📋 No email log found.\n\nThe EmailLog sheet will be created automatically when the first email is sent.\n\n💡 To view detailed logs:\n1. Check Apps Script execution logs (Extensions > Apps Script > Executions)\n2. Send a test email to create the log sheet\n3. Use this function again to view the log";
      safeUIAlert("Email Log Information", message);
      return message;
    }
    
    const data = logSheet.getDataRange().getValues();
    const logCount = data.length - 1; // Subtract header row
    
    if (logCount === 0) {
      const message = "📋 Email log is empty.\n\nNo emails have been sent yet or log data was cleared.\n\n💡 The log will populate automatically as emails are sent through the system.";
      safeUIAlert("Email Log Status", message);
      return message;
    }
    
    // Show recent log entries (last 10)
    const recentEntries = data.slice(-10).reverse(); // Last 10, most recent first
    let logMessage = `📧 EMAIL LOG SUMMARY\n\n📊 Total entries: ${logCount}\n📅 Showing recent entries:\n\n`;
    
    recentEntries.forEach((entry, index) => {
      if (index === 0) return; // Skip header
      const [timestamp, type, recipient, subject, status] = entry;
      logMessage += `${index}. ${timestamp}\n   ${type} → ${recipient}\n   Status: ${status}\n\n`;
    });
    
    logMessage += `💡 For complete logs:\n• Check the "EmailLog" sheet in this spreadsheet\n• View Apps Script execution logs (Extensions > Apps Script > Executions)`;
    
    console.log("Email log viewed:", { totalEntries: logCount, recentShown: Math.min(10, logCount) });
    safeUIAlert("Email Activity Log", logMessage);
    
    // Activate the log sheet for viewing
    logSheet.activate();
    
    return { totalEntries: logCount, logSheet: "EmailLog", message: logMessage };
    
  } catch (error) {
    const errorMsg = `❌ Failed to view email log: ${error.message}`;
    console.error("Email log viewer error:", error);
    safeUIAlert("Log Viewer Error", errorMsg);
    return { error: errorMsg };
  }
}

function showFacultyList() {
  try {
    console.log("=== SHOW FACULTY LIST ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      const message = `Sheet "${SHEET_NAME}" not found.`;
      safeUIAlert("Faculty List Error", message);
      return message;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      const message = "No faculty data found in the sheet.";
      safeUIAlert("Faculty List", message);
      return message;
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    const examinerColIndex = headers.indexOf("Examiner");
    
    if (emailColIndex === -1) {
      const message = `Column '${COL_EXAMINER_EMAIL}' not found.`;
      safeUIAlert("Faculty List Error", message);
      return message;
    }
    
    const facultyMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailColIndex];
      const examinerName = examinerColIndex !== -1 ? data[i][examinerColIndex] : 'Unknown';
      
      if (email && typeof email === 'string' && isValidEmail(email.trim())) {
        const trimmedEmail = email.trim();
        if (!facultyMap.has(trimmedEmail)) {
          facultyMap.set(trimmedEmail, {
            name: examinerName || 'Not Specified',
            studentCount: 0
          });
        }
        facultyMap.get(trimmedEmail).studentCount++;
      }
    }
    
    if (facultyMap.size === 0) {
      const message = "No valid faculty emails found in the system.";
      safeUIAlert("Faculty List", message);
      return message;
    }
    
    let facultyList = `👥 FACULTY LIST\n\n📊 Total Faculty: ${facultyMap.size}\n\n`;
    
    Array.from(facultyMap.entries()).sort().forEach(([email, info], index) => {
      facultyList += `${index + 1}. ${info.name}\n   📧 ${email}\n   👥 ${info.studentCount} student(s)\n\n`;
    });
    
    console.log("Faculty list displayed:", { totalFaculty: facultyMap.size });
    safeUIAlert("Faculty List", facultyList);
    
    return `Faculty list: ${facultyMap.size} faculty members`;
    
  } catch (error) {
    const errorMsg = `Failed to show faculty list: ${error.message}`;
    console.error("Faculty list error:", error);
    safeUIAlert("Faculty List Error", errorMsg);
    return errorMsg;
  }
}

// Simple test to verify the issue
function testSendReminderDirectly() {
  try {
    console.log("🧪 Testing sendDeadlineReminder directly...");
    
    // Test with valid parameters
    const testEmail = "sathyajain9@gmail.com";
    const testDueDate = new Date("2024-01-10");
    const testStats = {
      totalStudents: 5,
      studentsWithMarks: 3,
      studentsVerified: 2,
      completionPercentage: 60,
      verificationPercentage: 40
    };
    const testReminderCount = 1;
    
    console.log("Testing with valid parameters...");
    const result1 = sendDeadlineReminder(testEmail, testDueDate, testStats, testReminderCount);
    console.log("Result with valid params:", result1);
    
    // Test with undefined email
    console.log("Testing with undefined email...");
    const result2 = sendDeadlineReminder(undefined, testDueDate, testStats, testReminderCount);
    console.log("Result with undefined email:", result2);
    
    return { success: true, results: { valid: result1, undefined: result2 } };
    
  } catch (error) {
    console.error("Test failed:", error);
    return { success: false, error: error.message };
  }
}
function debugDeadlinesSheetContent() {
  try {
    console.log("🔍 DEBUGGING: Checking actual deadlines sheet content...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deadlinesSheet = ss.getSheetByName(FACULTY_DEADLINES_SHEET);
    
    if (!deadlinesSheet) {
      console.log("❌ FacultyDeadlines sheet not found!");
      return;
    }
    
    const data = deadlinesSheet.getDataRange().getValues();
    console.log(`📊 Sheet has ${data.length} rows`);
    
    // Show headers
    console.log("\n📋 HEADERS (Row 0):");
    const headers = data[0];
    headers.forEach((header, index) => {
      console.log(`  Column ${index}: "${header}"`);
    });
    
    // Test column mapping
    console.log("\n🗂️ TESTING COLUMN MAPPING:");
    try {
      const cols = getDeadlineColumnIndices(headers);
      console.log("Column mapping result:", cols);
      console.log(`Faculty Email column index: ${cols.facultyEmail}`);
    } catch (mappingError) {
      console.error("❌ Column mapping failed:", mappingError);
      return;
    }
    
    // Show actual data rows
    console.log("\n📋 DATA ROWS:");
    for (let i = 1; i < Math.min(data.length, 4); i++) { // Show max 3 data rows
      console.log(`\nRow ${i}:`);
      const row = data[i];
      row.forEach((cell, index) => {
        const header = headers[index] || `Col${index}`;
        console.log(`  ${header}: "${cell}" (${typeof cell})`);
      });
    }
    
    if (data.length > 4) {
      console.log(`\n... and ${data.length - 4} more rows`);
    }
    
    return data;
    
  } catch (error) {
    console.error("❌ Debug failed:", error);
    return null;
  }
}
function testColumnMapping() {
  try {
    console.log("🧪 TESTING: Column mapping function...");
    
    // Test with your actual headers
    const testHeaders = [
      "Faculty Email", "Due Date", "Total Students", "Students With Marks", 
      "Students Verified", "Completion Status", "Last Reminder Sent", 
      "Reminder Count", "Completion Confirmed Date", "Notes"
    ];
    
    console.log("Test headers:", testHeaders);
    
    const result = getDeadlineColumnIndices(testHeaders);
    console.log("Mapping result:", result);
    
    // Check Faculty Email specifically
    const facultyEmailIndex = testHeaders.indexOf("Faculty Email");
    console.log(`"Faculty Email" should be at index: ${facultyEmailIndex}`);
    console.log(`Function returned index: ${result.facultyEmail}`);
    
    if (result.facultyEmail === facultyEmailIndex) {
      console.log("✅ Column mapping is working correctly");
    } else {
      console.log("❌ Column mapping has issues");
    }
    
    return result;
    
  } catch (error) {
    console.error("❌ Column mapping test failed:", error);
    return null;
  }
}

function debugDeadlinesData() {
  try {
    console.log("🔍 DEBUGGING: Deadlines sheet data...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deadlinesSheet = ss.getSheetByName(FACULTY_DEADLINES_SHEET);
    
    if (!deadlinesSheet) {
      console.log("❌ FacultyDeadlines sheet not found!");
      return;
    }
    
    const data = deadlinesSheet.getDataRange().getValues();
    console.log(`📊 Sheet has ${data.length} rows`);
    
    if (data.length === 0) {
      console.log("❌ Sheet is completely empty!");
      return;
    }
    
    // Show headers
    console.log("\n📋 HEADERS (Row 0):");
    const headers = data[0];
    headers.forEach((header, index) => {
      console.log(`  Column ${index}: "${header}" (${typeof header})`);
    });
    
    // Test column mapping
    console.log("\n🗂️ TESTING COLUMN MAPPING:");
    let cols;
    try {
      cols = getDeadlineColumnIndices(headers);
      console.log("✅ Column mapping successful:", cols);
      
      if (cols.facultyEmail === -1) {
        console.error("❌ Faculty Email column not found!");
        return;
      }
      
    } catch (mappingError) {
      console.error("❌ Column mapping failed:", mappingError);
      return;
    }
    
    // Check first few data rows for email values
    console.log("\n📧 EMAIL VALIDATION CHECK:");
    for (let i = 1; i < Math.min(data.length, 6); i++) {
      const row = data[i];
      const emailRaw = row[cols.facultyEmail];
      
      console.log(`\nRow ${i}:`);
      console.log(`  Raw email value: "${emailRaw}" (${typeof emailRaw})`);
      console.log(`  Is null: ${emailRaw === null}`);
      console.log(`  Is undefined: ${emailRaw === undefined}`);
      console.log(`  Is empty string: ${emailRaw === ''}`);
      
      if (emailRaw !== null && emailRaw !== undefined && emailRaw !== '') {
        try {
          const emailStr = emailRaw.toString().trim();
          console.log(`  String conversion: "${emailStr}"`);
          console.log(`  Is valid email: ${isValidEmail(emailStr)}`);
        } catch (strError) {
          console.log(`  String conversion failed: ${strError.message}`);
        }
      } else {
        console.log(`  ❌ Email value is null/undefined/empty - this row will be skipped`);
      }
    }
    
    // Count valid vs invalid emails
    let validEmails = 0;
    let invalidEmails = 0;
    let emptyEmails = 0;
    
    for (let i = 1; i < data.length; i++) {
      const emailRaw = data[i][cols.facultyEmail];
      
      if (emailRaw === null || emailRaw === undefined || emailRaw === '') {
        emptyEmails++;
      } else {
        try {
          const emailStr = emailRaw.toString().trim();
          if (emailStr === '' || emailStr === 'undefined' || emailStr === 'null') {
            invalidEmails++;
          } else if (isValidEmail(emailStr)) {
            validEmails++;
          } else {
            invalidEmails++;
          }
        } catch (error) {
          invalidEmails++;
        }
      }
    }
    
    console.log(`\n📊 EMAIL SUMMARY:`);
    console.log(`  Valid emails: ${validEmails}`);
    console.log(`  Invalid emails: ${invalidEmails}`);
    console.log(`  Empty/null emails: ${emptyEmails}`);
    console.log(`  Total data rows: ${data.length - 1}`);
    
    if (validEmails === 0) {
      console.log("\n❌ WARNING: No valid emails found! This explains why sendDeadlineReminder gets undefined.");
      console.log("💡 SOLUTIONS:");
      console.log("  1. Check if Faculty Email column has the correct data");
      console.log("  2. Make sure email addresses are properly formatted");
      console.log("  3. Verify the deadlines sheet has been populated correctly");
    }
    
    return {
      totalRows: data.length - 1,
      validEmails,
      invalidEmails,
      emptyEmails,
      headers
    };
    
  } catch (error) {
    console.error("❌ Debug failed:", error);
    return { error: error.message };
  }
}

// Debug version of sendDeadlineReminder to see what's failing
function debugSendDeadlineReminder(email, dueDate = null, stats = null, reminderCount = null) {
  console.log("🔍 DEBUG: Starting sendDeadlineReminder analysis...");
  console.log(`📧 Called with: email="${email}", dueDate=${dueDate}, stats=${stats}, reminderCount=${reminderCount}`);
  
  try {
    // Step 1: Validate email
    console.log("\n📧 STEP 1: Email Validation");
    if (!email || typeof email !== 'string') {
      console.error('❌ Invalid email:', email);
      return { success: false, step: 1, reason: "Invalid email parameter" };
    }
    console.log(`✅ Email is valid: "${email}"`);
    
    // Step 2: Get/validate due date
    console.log("\n📅 STEP 2: Due Date Processing");
    if (!dueDate) {
      console.log('🔍 Auto-getting due date from deadlines sheet...');
      
      try {
        const deadlinesSheet = initializeFacultyDeadlines();
        const data = deadlinesSheet.getDataRange().getValues();
        
        console.log(`📊 Deadlines sheet has ${data.length} rows`);
        
        if (data.length <= 1) {
          console.error('❌ No data in deadlines sheet');
          return { success: false, step: 2, reason: "No deadlines data found" };
        }
        
        const headers = data[0];
        console.log('📋 Headers:', headers);
        
        const cols = getDeadlineColumnIndices(headers);
        console.log('🗂️ Column mapping:', cols);
        
        if (cols.facultyEmail === -1 || cols.dueDate === -1) {
          console.error('❌ Required columns not found');
          return { success: false, step: 2, reason: "Required columns missing" };
        }
        
        // Search for faculty email
        let foundEmail = false;
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const rowEmail = row[cols.facultyEmail];
          
          console.log(`🔍 Row ${i}: Checking "${rowEmail}" vs "${email}"`);
          
          if (rowEmail && rowEmail.toString().trim().toLowerCase() === email.toLowerCase()) {
            foundEmail = true;
            const rawDueDate = row[cols.dueDate];
            console.log(`✅ Found faculty! Raw due date: ${rawDueDate}`);
            
            dueDate = new Date(rawDueDate);
            if (isNaN(dueDate.getTime())) {
              console.error('❌ Invalid due date found:', rawDueDate);
              return { success: false, step: 2, reason: "Invalid due date in sheet" };
            }
            
            console.log(`✅ Due date parsed: ${dueDate.toLocaleDateString()}`);
            break;
          }
        }
        
        if (!foundEmail) {
          console.error(`❌ Faculty email "${email}" not found in deadlines sheet`);
          console.log('📋 Available emails in sheet:');
          for (let i = 1; i < Math.min(data.length, 6); i++) {
            const rowEmail = data[i][cols.facultyEmail];
            console.log(`  Row ${i}: "${rowEmail}"`);
          }
          return { success: false, step: 2, reason: "Faculty not found in deadlines sheet" };
        }
        
      } catch (sheetError) {
        console.error('❌ Error accessing deadlines sheet:', sheetError);
        return { success: false, step: 2, reason: "Sheet access error: " + sheetError.message };
      }
    } else {
      console.log(`✅ Due date provided: ${dueDate}`);
    }
    
    // Step 3: Check if overdue
    console.log("\n⏰ STEP 3: Overdue Check");
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dueDate.setHours(0, 0, 0, 0);
    
    const daysPastDue = Math.ceil((today - dueDate) / (1000 * 60 * 60 * 24));
    console.log(`📅 Today: ${today.toLocaleDateString()}`);
    console.log(`📅 Due Date: ${dueDate.toLocaleDateString()}`);
    console.log(`📊 Days past due: ${daysPastDue}`);
    
    if (daysPastDue <= 0) {
      console.log(`⚠️ Not overdue yet (${daysPastDue} days). Reminder not needed.`);
      return { success: false, step: 3, reason: `Not overdue yet (${daysPastDue} days)` };
    }
    console.log(`✅ Is overdue by ${daysPastDue} days`);
    
    // Step 4: Get/validate stats
    console.log("\n📊 STEP 4: Faculty Stats");
    if (!stats) {
      console.log('🔍 Auto-getting faculty stats...');
      try {
        stats = getFacultyStudentStats(email);
        console.log('📊 Retrieved stats:', stats);
        
        if (!stats || typeof stats !== 'object') {
          console.error('❌ Invalid stats retrieved');
          return { success: false, step: 4, reason: "Invalid faculty stats" };
        }
        
        if (stats.error) {
          console.error('❌ Stats retrieval error:', stats.error);
          return { success: false, step: 4, reason: "Stats error: " + stats.error };
        }
        
      } catch (statsError) {
        console.error('❌ Error getting faculty stats:', statsError);
        return { success: false, step: 4, reason: "Stats retrieval error: " + statsError.message };
      }
    } else {
      console.log(`✅ Stats provided:`, stats);
    }
    
    // Step 5: Set reminder count
    console.log("\n🔢 STEP 5: Reminder Count");
    if (!reminderCount) {
      reminderCount = 1;
      console.log('🔍 Using default reminder count: 1');
    } else {
      console.log(`✅ Reminder count provided: ${reminderCount}`);
    }
    
    // Step 6: Check email quota
    console.log("\n📧 STEP 6: Email Quota Check");
    const quota = MailApp.getRemainingDailyQuota();
    console.log(`📊 Remaining email quota: ${quota}`);
    
    if (quota < 1) {
      console.error('❌ No email quota remaining');
      return { success: false, step: 6, reason: "No email quota remaining" };
    }
    console.log(`✅ Email quota available: ${quota}`);
    
    // Step 7: Try sending email
    console.log("\n📧 STEP 7: Sending Email");
    
    const subject = `Urgent: Assessment Deadline Passed - Action Required [Reminder ${reminderCount}]`;
    
    const message = `Dear Faculty Member,

Sai Ram,

This is an urgent reminder regarding your assessment deadline.

📊 DEADLINE STATUS:
- Due Date: ${dueDate.toLocaleDateString()}
- Days Overdue: ${daysPastDue} days
- Reminder Count: ${reminderCount}

📈 YOUR PROGRESS:
- Total Students Assigned: ${stats.totalStudents || 0}
- Marks Uploaded: ${stats.studentsWithMarks || 0}/${stats.totalStudents || 0} (${stats.completionPercentage || 0}%)
- Students Verified: ${stats.studentsVerified || 0}/${stats.totalStudents || 0} (${stats.verificationPercentage || 0}%)

⚠️ IMMEDIATE ACTION REQUIRED:
${(stats.studentsWithMarks || 0) < (stats.totalStudents || 0) ? 
  `• Upload marks for ${(stats.totalStudents || 0) - (stats.studentsWithMarks || 0)} remaining students` : 
  '• All marks uploaded - Please verify and confirm completion'}

🌐 Faculty Portal: ${WEB_APP_URL}

📝 Steps to Complete:
1. Login to the faculty portal using your OTP
2. Upload marks for all assigned students
3. Verify all marks using the "Verify" button
4. Click "Confirm Completion" when all marks are uploaded

⏰ Please complete this immediately to avoid further escalation.

For technical support, contact the examination committee.

Best regards,
Controller of Examinations Section
Sri Sathya Sai Institute of Higher Learning

---
This is an automated reminder. Please complete your assessment to stop these notifications.`;

    try {
      MailApp.sendEmail(email, subject, message);
      console.log(`✅ Email sent successfully to ${email}`);
      
      logEmailActivity("DEBUG_DEADLINE_REMINDER", email, subject, "SUCCESS");
      
      return { 
        success: true, 
        step: 7, 
        message: "Email sent successfully",
        details: {
          email,
          daysPastDue,
          reminderCount,
          stats
        }
      };
      
    } catch (emailError) {
      console.error('❌ Failed to send email:', emailError);
      return { success: false, step: 7, reason: "Email sending failed: " + emailError.message };
    }
    
  } catch (error) {
    console.error('❌ Unexpected error in sendDeadlineReminder:', error);
    return { success: false, step: 0, reason: "Unexpected error: " + error.message };
  }
}

// Test function to run the debug version
function testSendReminderWithDebug() {
  console.log("🧪 Testing sendDeadlineReminder with debug info...");
  
  // Replace with a real faculty email from your system
  const testEmail = "sathyajain9@gmail.com"; // ⚠️ CHANGE THIS TO A REAL EMAIL
  
  console.log(`🔍 Testing with email: ${testEmail}`);
  
  const result = debugSendDeadlineReminder(testEmail);
  
  console.log("\n📊 FINAL RESULT:", result);
  
  if (!result.success) {
    console.log(`❌ Failed at step ${result.step}: ${result.reason}`);
    
    // Provide specific guidance based on failure step
    switch(result.step) {
      case 1:
        console.log("💡 SOLUTION: Provide a valid email address");
        break;
      case 2:
        console.log("💡 SOLUTION: Make sure the faculty exists in the FacultyDeadlines sheet");
        break;
      case 3:
        console.log("💡 SOLUTION: Set a past due date for testing, or use a faculty that's actually overdue");
        break;
      case 4:
        console.log("💡 SOLUTION: Make sure the faculty has students assigned in the main sheet");
        break;
      case 6:
        console.log("💡 SOLUTION: Wait for email quota to reset or contact admin");
        break;
      case 7:
        console.log("💡 SOLUTION: Check email permissions and try again");
        break;
    }
  } else {
    console.log("✅ SUCCESS: Email reminder sent successfully!");
  }
  
  return result;
}

// ===== ENHANCED DEADLINE VALIDATION FUNCTIONS =====


function checkFacultyDeadlineExists(email) {
  console.log(`🔍 Checking deadline for faculty: ${email}`);
  
  try {
    // Validate email parameter
    if (!email || typeof email !== 'string' || email.trim() === '') {
      return {
        hasDeadline: false,
        dueDate: null,
        message: 'Invalid email parameter'
      };
    }
    
    // Initialize deadlines sheet
    const deadlinesSheet = initializeFacultyDeadlines();
    const data = deadlinesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        hasDeadline: false,
        dueDate: null,
        message: 'No deadline data found - please set deadlines first'
      };
    }
    
    // Get column mapping
    const headers = data[0].map(h => h.toString().trim());
    const cols = getDeadlineColumnIndices(headers);
    
    if (cols.facultyEmail === -1 || cols.dueDate === -1) {
      return {
        hasDeadline: false,
        dueDate: null,
        message: 'Deadline sheet structure error - missing required columns'
      };
    }
    
    // Search for faculty deadline
    const emailLower = email.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = row[cols.facultyEmail];
      
      if (rowEmail && rowEmail.toString().trim().toLowerCase() === emailLower) {
        const dueDateValue = row[cols.dueDate];
        
        if (!dueDateValue) {
          return {
            hasDeadline: false,
            dueDate: null,
            message: 'Faculty found but no due date set'
          };
        }
        
        try {
          const dueDate = new Date(dueDateValue);
          if (isNaN(dueDate.getTime())) {
            return {
              hasDeadline: false,
              dueDate: null,
              message: 'Faculty found but due date is invalid'
            };
          }
          
          console.log(`✅ Deadline found for ${email}: ${dueDate.toLocaleDateString()}`);
          return {
            hasDeadline: true,
            dueDate: dueDate,
            message: `Deadline set for ${dueDate.toLocaleDateString()}`
          };
        } catch (dateError) {
          return {
            hasDeadline: false,
            dueDate: null,
            message: 'Faculty found but due date parsing failed'
          };
        }
      }
    }
    
    // Faculty not found in deadlines sheet
    return {
      hasDeadline: false,
      dueDate: null,
      message: 'Faculty not found in deadlines sheet - please add deadline'
    };
    
  } catch (error) {
    console.error(`❌ Error checking deadline for ${email}:`, error);
    return {
      hasDeadline: false,
      dueDate: null,
      message: `Error checking deadline: ${error.message}`
    };
  }
}


function validateFacultyDeadlines(facultyEmails) {
  console.log(`🔍 Validating deadlines for ${facultyEmails.length} faculty members...`);
  
  const validFaculty = [];
  const invalidFaculty = [];
  const validationResults = [];
  
  facultyEmails.forEach(email => {
    const result = checkFacultyDeadlineExists(email);
    
    if (result.hasDeadline) {
      validFaculty.push({
        email: email,
        dueDate: result.dueDate,
        message: result.message
      });
    } else {
      invalidFaculty.push({
        email: email,
        reason: result.message
      });
    }
    
    validationResults.push({
      email: email,
      valid: result.hasDeadline,
      dueDate: result.dueDate,
      message: result.message
    });
  });
  
  const summary = {
    totalFaculty: facultyEmails.length,
    validCount: validFaculty.length,
    invalidCount: invalidFaculty.length,
    validationPercentage: facultyEmails.length > 0 ? Math.round((validFaculty.length / facultyEmails.length) * 100) : 0
  };
  
  console.log(`📊 Deadline validation summary:`, summary);
  
  return {
    validFaculty: validFaculty,
    invalidFaculty: invalidFaculty,
    validationResults: validationResults,
    summary: summary
  };
}


function sendAdminDeadlineAlert(invalidFaculty, summary) {
  try {
    console.log(`📧 Sending admin alert for ${invalidFaculty.length} faculty skipped due to missing deadlines...`);
    
    if (invalidFaculty.length === 0) {
      console.log('✅ No faculty skipped - no admin alert needed');
      return true;
    }
    
    const subject = `⚠️ Faculty Skipped - Missing Deadlines Alert`;
    
    let emailBody = `Dear Administrator,

Sai Ram,

The Faculty Portal Email System has SELECTIVELY processed faculty members based on deadline availability.

📊 SELECTIVE PROCESSING SUMMARY:
═══════════════════════════════════════════════════════════════
• Total Faculty Found: ${summary.totalFaculty}
• Emails Sent (Valid Deadlines): ${summary.validCount} (${summary.validationPercentage}%)
• Emails Skipped (Missing Deadlines): ${summary.invalidCount}
• Processing Date: ${new Date().toLocaleString()}

✅ SYSTEM BEHAVIOR:
═══════════════════════════════════════════════════════════════
• OTP emails were successfully sent to faculty WITH deadlines
• Faculty WITHOUT deadlines were automatically skipped
• System continued processing rather than stopping completely
• This ensures maximum faculty participation while maintaining data integrity

❌ FACULTY SKIPPED (Missing Deadlines):
═══════════════════════════════════════════════════════════════
`;

    invalidFaculty.forEach((faculty, index) => {
      emailBody += `${index + 1}. ${faculty.email}
   Reason: ${faculty.reason}
   Status: SKIPPED - No OTP email sent

`;
    });

    emailBody += `
⚠️ IMPACT FOR SKIPPED FACULTY:
═══════════════════════════════════════════════════════════════
• NO OTP emails sent to faculty without deadlines
• These faculty cannot access the portal until deadlines are set
• Deadline reminders cannot be processed for these faculty
• Assessment tracking will be incomplete for these faculty
• Portal access is blocked until deadlines are configured

✅ REQUIRED ACTIONS TO INCLUDE SKIPPED FACULTY:
═══════════════════════════════════════════════════════════════
1. Set deadlines for all SKIPPED faculty members using:
   • Menu: "Deadline Management" > "Set Faculty Deadline"
   • Or use setFacultyDeadline(email, date) function

2. Recommended deadline format: YYYY-MM-DD

3. After setting deadlines, re-run the email system:
   • Menu: "Enhanced Exam Tools" > "Send Emails (With Deadline Validation)"
   • Only the newly configured faculty will receive OTP emails

4. Verify all faculty deadlines using:
   • Menu: "Deadline Management" > "Validate All Faculty Deadlines"



🔄 SYSTEM STATUS:
═══════════════════════════════════════════════════════════════
✅ Email system is working correctly for faculty with deadlines
⚠️ Skipped faculty need deadline configuration to receive future emails
🔄 System will automatically include newly configured faculty in next run

💡 BEST PRACTICE:
═══════════════════════════════════════════════════════════════
• Set deadlines for all faculty BEFORE running email system
• Use "Validate All Faculty Deadlines" to check before sending emails
• Regular deadline audits prevent faculty from being skipped

This is an automated alert - please do not reply to this email.

Best regards,
Enhanced Faculty Portal Email System
Sri Sathya Sai Institute of Higher Learning

---
Alert ID: SELECTIVE-${Date.now()}
Generated: ${new Date().toLocaleString()}
---`;

    // Send email to admin
    MailApp.sendEmail(ADMIN_EMAIL, subject, emailBody);
    
    // Log the activity
    logEmailActivity("ADMIN_SELECTIVE_ALERT", ADMIN_EMAIL, subject, "SUCCESS");
    
    console.log(`✅ Admin selective processing alert sent successfully to ${ADMIN_EMAIL}`);
    return true;
    
  } catch (error) {
    console.error(`❌ Error sending admin selective processing alert:`, error);
    logEmailActivity("ADMIN_SELECTIVE_ALERT", ADMIN_EMAIL || "admin", "Selective Processing Alert", `FAILED: ${error.message}`);
    return false;
  }
}

function sendEmailsWithDeadlineValidation() {
  try {
    console.log("=== ENHANCED BULK EMAIL SENDING WITH SELECTIVE DEADLINE VALIDATION ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    if (emailColIndex === -1) throw new Error(`Column '${COL_EXAMINER_EMAIL}' not found`);
    
    // ✅ STEP 1: Collect unique faculty emails
    const allFacultyEmails = {};
    const errors = [];
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailColIndex];
      if (email && typeof email === 'string') {
        const trimmedEmail = email.trim().toLowerCase();
        if (isValidEmail(trimmedEmail) && !allFacultyEmails[trimmedEmail]) {
          allFacultyEmails[trimmedEmail] = generateOTP();
        } else if (!isValidEmail(trimmedEmail)) {
          errors.push(`Invalid email format: ${email}`);
        }
      }
    }
    
    const facultyEmailList = Object.keys(allFacultyEmails);
    
    if (facultyEmailList.length === 0) {
      const message = "No valid faculty emails found to process.";
      console.error(message);
      safeUIAlert("No Recipients", message);
      return message;
    }
    
    console.log(`📊 Found ${facultyEmailList.length} unique faculty emails`);
    
    // ✅ STEP 2: VALIDATE DEADLINES FOR ALL FACULTY (INDIVIDUAL CHECK)
    console.log("🔍 Performing individual deadline validation for each faculty...");
    
    const validFacultyList = [];
    const invalidFacultyList = [];
    const validFacultyEmails = {};
    
    // Check each faculty individually
    facultyEmailList.forEach(email => {
      const deadlineCheck = checkFacultyDeadlineExists(email);
      
      if (deadlineCheck.hasDeadline) {
        // Faculty has valid deadline - add to valid list
        validFacultyList.push({
          email: email,
          dueDate: deadlineCheck.dueDate,
          message: deadlineCheck.message
        });
        validFacultyEmails[email] = allFacultyEmails[email]; // Keep OTP for this faculty
        console.log(`✅ ${email}: Deadline valid - Will receive OTP`);
      } else {
        // Faculty missing deadline - add to invalid list but continue processing others
        invalidFacultyList.push({
          email: email,
          reason: deadlineCheck.message
        });
        console.log(`❌ ${email}: ${deadlineCheck.message} - Will NOT receive OTP`);
      }
    });
    
    // ✅ STEP 3: REPORT VALIDATION RESULTS
    const validationSummary = {
      totalFaculty: facultyEmailList.length,
      validCount: validFacultyList.length,
      invalidCount: invalidFacultyList.length,
      validationPercentage: facultyEmailList.length > 0 ? Math.round((validFacultyList.length / facultyEmailList.length) * 100) : 0
    };
    
    console.log(`📊 Individual validation results:`, validationSummary);
    console.log(`✅ Will send emails to: ${validFacultyList.length} faculty`);
    console.log(`❌ Will skip emails for: ${invalidFacultyList.length} faculty`);
    
    // ✅ STEP 4: SEND ADMIN ALERT FOR SKIPPED FACULTY (IF ANY)
    if (invalidFacultyList.length > 0) {
      console.log(`⚠️ Sending admin alert for ${invalidFacultyList.length} faculty without deadlines...`);
      sendAdminDeadlineAlert(invalidFacultyList, validationSummary);
    }
    
    // ✅ STEP 5: PROCEED WITH EMAIL SENDING FOR VALID FACULTY ONLY
    if (validFacultyList.length === 0) {
      const message = `❌ No faculty members have valid deadlines set. Cannot send any OTP emails.\n\nAll ${invalidFacultyList.length} faculty members need deadlines configured.\nAdministrator has been notified.`;
      console.error("❌ No valid faculty found - no emails sent");
      safeUIAlert("No Valid Faculty", message);
      return `No emails sent: All ${invalidFacultyList.length} faculty members need deadlines set.`;
    }
    
    console.log(`✅ Proceeding with email sending for ${validFacultyList.length} faculty with valid deadlines`);
    
    // Check email quota
    const emailQuota = MailApp.getRemainingDailyQuota();
    if (emailQuota < validFacultyList.length) {
      throw new Error(`Insufficient email quota (${emailQuota}) to send ${validFacultyList.length} emails.`);
    }
    
    // ✅ STEP 6: STORE OTPS FOR VALID FACULTY ONLY
    try {
      PropertiesService.getScriptProperties().setProperty("OTPs", JSON.stringify(validFacultyEmails));
      console.log(`✅ STORED OTPs for ${validFacultyList.length} faculty members with valid deadlines`);
    } catch (storageError) {
      console.error("❌ STORAGE ERROR:", storageError);
      throw new Error(`Failed to store OTPs: ${storageError.message}`);
    }
    
    // ✅ STEP 7: SEND EMAILS ONLY TO FACULTY WITH VALID DEADLINES
    const portalLink = WEB_APP_URL;
    let successCount = 0;
    const emailErrors = [];
    
    for (const facultyData of validFacultyList) {
      try {
        const email = facultyData.email;
        const otp = validFacultyEmails[email];
        const dueDate = facultyData.dueDate;
        
        const subject = "Faculty Portal Access - Your One-Time Password [OTP:" + otp + "]";
        
        const message = `Dear Faculty Member,

Sai Ram,

Your One-Time Password (OTP) for accessing the Faculty Portal is ready:

Portal Access: ${portalLink}

Login Credentials:
- 📧 Email: ${email}
- 🔐 Your OTP: ${otp}

📅 Assessment Deadline: ${dueDate.toLocaleDateString()}
⏰ Days remaining: ${Math.ceil((dueDate - new Date()) / (1000 * 60 * 60 * 24))} days

⚠️ Important:
• Use the EXACT email address: ${email}
• Use the 6-digit OTP: ${otp}
• Complete assessment before deadline
• Do not share your OTP with anyone

For technical support, please contact the Examinations Section.

Best Regards,
Examinations Section`;

        MailApp.sendEmail(email, subject, message);
        successCount++;
        logEmailActivity("BULK_OTP_EMAIL_SELECTIVE", email, subject, "SUCCESS");
        
        Utilities.sleep(200); // Rate limiting
        
      } catch (emailError) {
        console.error(`Failed to send email to ${facultyData.email}:`, emailError);
        emailErrors.push(`Failed to send to ${facultyData.email}: ${emailError.message}`);
        logEmailActivity("BULK_OTP_EMAIL_SELECTIVE", facultyData.email, subject || "OTP Email", `FAILED: ${emailError.message}`);
      }
    }
    
    // ✅ STEP 8: SHOW COMPREHENSIVE RESULTS WITH SELECTIVE PROCESSING
    let resultMessage = `✅ SELECTIVE OTP EMAIL PROCESSING COMPLETE!\n\n`;
    resultMessage += `📊 Faculty Processing Results:\n`;
    resultMessage += `• Total Faculty Found: ${facultyEmailList.length}\n`;
    resultMessage += `• With Valid Deadlines: ${validFacultyList.length} ✅\n`;
    resultMessage += `• Missing Deadlines: ${invalidFacultyList.length} ❌\n`;
    resultMessage += `• Processing Rate: ${validationSummary.validationPercentage}%\n\n`;
    
    resultMessage += `📧 Email Sending Results:\n`;
    resultMessage += `• OTPs Generated: ${validFacultyList.length}\n`;
    resultMessage += `• Emails Sent: ${successCount} ✅\n`;
    resultMessage += `• Email Failures: ${emailErrors.length}\n`;
    resultMessage += `• Emails Skipped: ${invalidFacultyList.length} (No deadlines)\n`;
    resultMessage += `• Success Rate: ${validFacultyList.length > 0 ? Math.round((successCount / validFacultyList.length) * 100) : 0}%\n\n`;
    
    if (invalidFacultyList.length > 0) {
      resultMessage += `⚠️ Faculty SKIPPED (No Deadlines):\n`;
      invalidFacultyList.slice(0, 5).forEach((faculty, index) => {
        resultMessage += `${index + 1}. ${faculty.email}\n`;
      });
      if (invalidFacultyList.length > 5) {
        resultMessage += `... and ${invalidFacultyList.length - 5} more faculty\n`;
      }
      resultMessage += `\n📧 Administrator notified to set missing deadlines\n`;
    }
    
    if (emailErrors.length > 0) {
      resultMessage += `\n❌ Email Sending Errors: ${emailErrors.length}`;
    }
    
    resultMessage += `\n\n✅ SELECTIVE PROCESSING: Emails sent only to faculty with deadlines`;
    
    console.log("✅ SELECTIVE EMAIL PROCESSING COMPLETED:", { 
      totalFound: facultyEmailList.length,
      validDeadlines: validFacultyList.length,
      invalidDeadlines: invalidFacultyList.length,
      emailsSent: successCount, 
      emailErrors: emailErrors.length 
    });
    
    safeUIAlert("Selective Email Processing Complete", resultMessage);
    
    if (successCount > 0) {
      return `✅ Selective processing: ${successCount} emails sent to faculty with deadlines. ${invalidFacultyList.length} faculty skipped (no deadlines). ${emailErrors.length} email failures.`;
    } else {
      return `❌ No emails sent: ${invalidFacultyList.length} faculty need deadlines set, ${emailErrors.length} email failures.`;
    }
    
  } catch (error) {
    console.error("❌ CRITICAL ERROR in selective email sending:", error, error.stack);
    const errorMsg = `❌ Selective email processing failed: ${error.message}`;
    safeUIAlert("Selective Email Processing Failed", errorMsg);
    return errorMsg;
  }
}

// ===== ENHANCED DEADLINE REMINDER WITH VALIDATION =====
function sendDeadlineReminderWithValidation(email, dueDate = null, stats = null, reminderCount = null) {
  console.log(`📧 Enhanced sendDeadlineReminder with validation for: ${email}`);
  
  try {
    // ✅ STEP 1: VALIDATE EMAIL PARAMETER
    if (!email || typeof email !== 'string') {
      console.error('sendDeadlineReminderWithValidation: Invalid email:', email);
      return false;
    }
    
    // ✅ STEP 2: CHECK IF DEADLINE EXISTS FOR THIS FACULTY
    const deadlineCheck = checkFacultyDeadlineExists(email);
    
    if (!deadlineCheck.hasDeadline) {
      console.error(`❌ Cannot send reminder to ${email}: ${deadlineCheck.message}`);
      
      // Send admin notification about missing deadline
      sendAdminDeadlineAlert([{
        email: email,
        reason: deadlineCheck.message
      }], {
        totalFaculty: 1,
        validCount: 0,
        invalidCount: 1,
        validationPercentage: 0
      });
      
      return false;
    }
    
    console.log(`✅ Deadline validation passed for ${email}: ${deadlineCheck.message}`);
    
    // ✅ STEP 3: USE VALIDATED DEADLINE OR PROVIDED ONE
    const validatedDueDate = dueDate || deadlineCheck.dueDate;
    
    // ✅ STEP 4: PROCEED WITH ORIGINAL REMINDER LOGIC
    return sendDeadlineReminder(email, validatedDueDate, stats, reminderCount);
    
  } catch (error) {
    console.error(`❌ Error in sendDeadlineReminderWithValidation for ${email}:`, error);
    return false;
  }
}

// ===== ENHANCED OVERDUE FACULTY CHECK WITH DEADLINE VALIDATION =====
function checkOverdueFacultiesWithValidation() {
  try {
    console.log("🔍 Starting enhanced overdue faculty check with deadline validation...");
    
    const deadlinesSheet = initializeFacultyDeadlines();
    const data = deadlinesSheet.getDataRange().getValues();
    
    console.log(`📊 Sheet data: ${data.length} rows total`);
    
    if (data.length <= 1) {
      console.log("ℹ️ No faculty deadlines data found");
      return { success: true, message: "No deadlines to check" };
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const cols = getDeadlineColumnIndices(headers);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let remindersSet = 0;
    let completionChecks = 0;
    let errors = [];
    let processedCount = 0;
    let skippedCount = 0;
    let deadlineValidationFailures = 0;
    
    console.log(`🔄 Processing ${data.length - 1} data rows with enhanced deadline validation...`);

    for (let i = 1; i < data.length; i++) {
      console.log(`\n--- PROCESSING ROW ${i} WITH ENHANCED VALIDATION ---`);
      
      const row = data[i];
      
      if (!row || row.length === 0) {
        console.log(`❌ Row ${i}: Empty row, skipping`);
        skippedCount++;
        continue;
      }
      
      // ✅ ENHANCED: Email validation with deadline check
      const emailRaw = row[cols.facultyEmail];
      console.log(`📧 Row ${i}: Raw email from column ${cols.facultyEmail}: "${emailRaw}" (${typeof emailRaw})`);
      
      if (emailRaw === null || emailRaw === undefined || emailRaw === '') {
        console.log(`❌ Row ${i}: Faculty Email is null/undefined/empty, skipping`);
        skippedCount++;
        continue;
      }
      
      let email;
      try {
        email = emailRaw.toString().trim();
        if (email === '' || email === 'undefined' || email === 'null') {
          console.log(`❌ Row ${i}: Email is invalid after string conversion: "${email}", skipping`);
          skippedCount++;
          continue;
        }
        
        if (!isValidEmail(email)) {
          console.log(`❌ Row ${i}: Invalid email format: "${email}", skipping`);
          skippedCount++;
          continue;
        }
        
        console.log(`✅ Row ${i}: Valid email: "${email}"`);
      } catch (stringError) {
        console.log(`❌ Row ${i}: Failed to process email:`, stringError);
        skippedCount++;
        continue;
      }
      
      // ✅ ENHANCED: DEADLINE VALIDATION CHECK
      const deadlineCheck = checkFacultyDeadlineExists(email);
      
      if (!deadlineCheck.hasDeadline) {
        console.log(`❌ Row ${i}: Deadline validation failed for ${email}: ${deadlineCheck.message}`);
        deadlineValidationFailures++;
        errors.push(`Deadline validation failed for ${email}: ${deadlineCheck.message}`);
        skippedCount++;
        continue;
      }
      
      console.log(`✅ Row ${i}: Deadline validation passed for ${email}`);
      
      // Continue with existing logic...
      const dueDateValue = row[cols.dueDate];
      const completionStatus = row[cols.completionStatus] || "Pending";
      const lastReminderSent = row[cols.lastReminderSent] ? new Date(row[cols.lastReminderSent]) : null;
      const reminderCount = parseInt(row[cols.reminderCount]) || 0;
      
      console.log(`📅 Row ${i}: Due date: "${dueDateValue}", Status: "${completionStatus}"`);
      
      if (!dueDateValue) {
        console.log(`❌ Row ${i}: No due date for ${email}, skipping`);
        skippedCount++;
        continue;
      }
      
      let dueDate;
      try {
        dueDate = new Date(dueDateValue);
        if (isNaN(dueDate.getTime())) {
          console.log(`❌ Row ${i}: Invalid due date for ${email}: ${dueDateValue}`);
          errors.push(`Invalid due date for ${email}: ${dueDateValue}`);
          continue;
        }
        dueDate.setHours(0, 0, 0, 0);
        console.log(`✅ Row ${i}: Valid due date: ${dueDate.toLocaleDateString()}`);
      } catch (dateError) {
        console.log(`❌ Row ${i}: Date parsing error for ${email}:`, dateError);
        errors.push(`Date parsing error for ${email}: ${dateError.message}`);
        continue;
      }
      
      // Skip if already confirmed
      if (completionStatus === "Confirmed") {
        console.log(`✅ Row ${i}: ${email} already confirmed, skipping`);
        continue;
      }
      
      // Get faculty statistics
      console.log(`📊 Row ${i}: Getting stats for ${email}...`);
      const stats = getFacultyStudentStats(email);
      console.log(`📊 Row ${i}: Stats:`, stats);
      
      // Update stats in sheet
      try {
        deadlinesSheet.getRange(i + 1, cols.totalStudents + 1, 1, 3)
                     .setValues([[stats.totalStudents, stats.studentsWithMarks, stats.studentsVerified]]);
        console.log(`✅ Row ${i}: Updated stats in sheet`);
      } catch (updateError) {
        console.log(`❌ Row ${i}: Failed to update stats:`, updateError);
      }
      
      // Check completion status
      if (stats.totalStudents > 0 && stats.studentsWithMarks === stats.totalStudents && completionStatus === "Pending") {
        deadlinesSheet.getRange(i + 1, cols.completionStatus + 1).setValue("Ready for Confirmation");
        console.log(`✅ Row ${i}: ${email} ready for confirmation`);
        completionChecks++;
      }
      
      // Check if reminder is needed
      const isOverdue = today > dueDate;
      const needsReminder = completionStatus !== "Ready for Confirmation" && completionStatus !== "Confirmed";
      
      if (isOverdue && needsReminder) {
        const daysPastDue = Math.ceil((today - dueDate) / (1000 * 60 * 60 * 24));
        const daysSinceLastReminder = lastReminderSent ? 
          Math.floor((today - lastReminderSent) / (1000 * 60 * 60 * 24)) : 999;
        
        console.log(`📅 Row ${i}: ${daysPastDue} days overdue, ${daysSinceLastReminder} days since last reminder`);
        
        if (daysSinceLastReminder >= 1) {
          console.log(`📧 Row ${i}: SENDING enhanced reminder to: "${email}"`);
          
          // ✅ USE ENHANCED REMINDER FUNCTION WITH VALIDATION
          const reminderSent = sendDeadlineReminderWithValidation(email, dueDate, stats, reminderCount + 1);
          
          if (reminderSent) {
            try {
              deadlinesSheet.getRange(i + 1, cols.lastReminderSent + 1, 1, 2)
                           .setValues([[today, reminderCount + 1]]);
              remindersSet++;
              console.log(`✅ Row ${i}: Enhanced reminder sent successfully to ${email}`);
            } catch (updateError) {
              console.error(`❌ Row ${i}: Failed to update reminder data in sheet:`, updateError);
              errors.push(`Update failed for ${email}: ${updateError.message}`);
            }
          } else {
            console.log(`❌ Row ${i}: Failed to send enhanced reminder to ${email}`);
            errors.push(`Failed to send enhanced reminder to ${email}`);
          }
        } else {
          console.log(`⏰ Row ${i}: Too soon for reminder to ${email} (last sent ${daysSinceLastReminder} days ago)`);
        }
      }
      
      processedCount++;
    }
    
    // Check if all faculties have completed
    try {
      checkAllFacultiesCompleted();
    } catch (completionError) {
      console.error("❌ Error checking completion:", completionError);
    }
    
    console.log(`\n📊 ENHANCED FINAL SUMMARY:`);
    console.log(`  - Processed: ${processedCount}, Skipped: ${skippedCount}`);
    console.log(`  - Deadline validation failures: ${deadlineValidationFailures}`);
    console.log(`  - Reminders sent: ${remindersSet}, Ready for confirmation: ${completionChecks}`);
    console.log(`  - Errors: ${errors.length}`);
    
    // Send admin alert if there were deadline validation failures
    if (deadlineValidationFailures > 0) {
      console.log(`⚠️ Sending admin alert for ${deadlineValidationFailures} deadline validation failures`);
    }
    
    return {
      success: true,
      message: `Enhanced check completed: ${processedCount} processed, ${remindersSet} reminders sent, ${deadlineValidationFailures} deadline validation failures`,
      details: { 
        processedCount, 
        skippedCount, 
        remindersSet, 
        completionChecks, 
        deadlineValidationFailures,
        errors 
      }
    };
    
  } catch (error) {
    console.error("🚨 CRITICAL ERROR in enhanced checkOverdueFaculties:", error);
    return { success: false, message: error.message };
  }
}

// ===== MENU INTEGRATION FOR NEW FUNCTIONS =====
function createEnhancedMenus() {
  try {
    if (!isUIAvailable()) {
      console.log("UI not available - skipping enhanced menu creation");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Tools")
      .addItem("📧 Send Emails", "sendEmailsWithDeadlineValidation")
      .addItem("📊 Analysis", "analyse")
      .addItem("🔍 Debug Sheet Headers", "debugSheetHeaders")
      .addSeparator()
      .addSubMenu(ui.createMenu("CoE Functions")
        .addItem("🔄 Refresh Edit Requests", "refreshEditRequestsView")
        .addSeparator()
        .addItem("✅ Approve Selected Request", "approveSelectedRequest")
        .addItem("❌ Disapprove Selected Request", "disapproveSelectedRequest")
        .addItem("✅ Approve All Pending", "approveAllPending")
        .addSeparator()
        .addItem("🧪 Test Approval System", "testApprovalSystem")
        .addItem("🔓 Test Unlock System (UI)", "testUnlockSystemWithUI")
        .addItem("🔍 Test Unlock System (Console)", "testUnlockSystem")
        .addItem("🌐 Open Faculty Portal", "openFacultyPortal"))
      .addSeparator()
      .addSubMenu(ui.createMenu("Enhanced Email System")
        .addItem("📧 Test Email System", "testEmailSystem")
        .addItem("📊 Check Email Status", "checkEmailQuota")
        .addSeparator()
        .addItem("🔄 Send OTP", "sendEmailsWithDeadlineValidation")
        .addItem("📋 Send Individual OTP", "sendIndividualOTP")
        .addItem("👥 View Faculty List", "showFacultyList")
        .addSeparator()
        .addItem("🔍 Diagnose Email Issues", "troubleshootEmailIssues")
        .addItem("🔧 Repair Email System", "repairEmailSystem")
        .addSeparator()
        .addItem("🔐 Authorize Email Permissions", "authorizeEmailPermissions")
        .addItem("📋 View Email Log", "viewEmailLog"))
      .addSeparator()
      .addSubMenu(ui.createMenu("Enhanced Deadline Management")
        .addItem("📅 Initialize Deadlines Sheet", "initializeFacultyDeadlines")
        .addItem("⏰ Set Faculty Deadline", "setFacultyDeadlineUI")
        .addItem("📊 View Deadline Status", "viewDeadlineStatus")
        .addItem("🔍 Check Faculty Deadline", "checkFacultyDeadlineUI")
        .addItem("✅ Validate All Faculty Deadlines", "validateAllFacultyDeadlinesUI")
        .addSeparator()
        .addItem("🔍 Check Overdue (Selective)", "checkOverdueFacultiesWithValidation")
        .addItem("📧 Send Test Reminder", "sendTestReminder"))
      .addToUi();
      
    console.log("Enhanced menus created successfully");
  } catch (error) {
    console.log("Enhanced menu creation failed (expected in web app context):", error.message);
  }
}

// ===== UI HELPER FUNCTIONS FOR NEW FEATURES =====
function checkFacultyDeadlineUI() {
  try {
    if (!isUIAvailable()) return;
    
    const ui = SpreadsheetApp.getUi();
    const emailResponse = ui.prompt('Check Faculty Deadline', 'Enter faculty email to check:', ui.ButtonSet.OK_CANCEL);
    
    if (emailResponse.getSelectedButton() === ui.Button.OK) {
      const email = emailResponse.getResponseText().trim();
      const result = checkFacultyDeadlineExists(email);
      
      let message = `Faculty: ${email}\n\n`;
      if (result.hasDeadline) {
        message += `✅ Deadline Status: FOUND\n`;
        message += `📅 Due Date: ${result.dueDate.toLocaleDateString()}\n`;
        message += `📝 Message: ${result.message}`;
      } else {
        message += `❌ Deadline Status: NOT FOUND\n`;
        message += `📝 Reason: ${result.message}\n\n`;
        message += `Action Required:\n`;
        message += `1. Set deadline using "Set Faculty Deadline"\n`;
        message += `2. Verify deadline is properly configured\n`;
        message += `3. Re-check deadline status`;
      }
      
      ui.alert('Faculty Deadline Check Result', message, ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('Error in checkFacultyDeadlineUI:', error);
  }
}

function validateAllFacultyDeadlinesUI() {
  try {
    if (!isUIAvailable()) return;
    
    console.log('🔍 Starting faculty deadline validation...');
    
    // Get all faculty emails from main sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const emailColIndex = headers.indexOf(COL_EXAMINER_EMAIL);
    
    const facultyEmails = new Set();
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailColIndex];
      if (email && typeof email === 'string' && isValidEmail(email.trim())) {
        facultyEmails.add(email.trim().toLowerCase());
      }
    }
    
    const facultyEmailList = Array.from(facultyEmails);
    console.log(`📊 Validating deadlines for ${facultyEmailList.length} faculty members...`);
    
    const validation = validateFacultyDeadlines(facultyEmailList);
    
    console.log('✅ Validation completed');
    
    let message = `📊 FACULTY DEADLINE VALIDATION RESULTS\n\n`;
    message += `Total Faculty: ${validation.summary.totalFaculty}\n`;
    message += `With Deadlines: ${validation.summary.validCount} (${validation.summary.validationPercentage}%)\n`;
    message += `Missing Deadlines: ${validation.summary.invalidCount}\n\n`;
    
    if (validation.invalidFaculty.length > 0) {
      message += `❌ Faculty WITHOUT Deadlines:\n`;
      validation.invalidFaculty.slice(0, 10).forEach((faculty, index) => {
        message += `${index + 1}. ${faculty.email}\n`;
      });
      if (validation.invalidFaculty.length > 10) {
        message += `... and ${validation.invalidFaculty.length - 10} more\n`;
      }
      message += `\nAction Required: Set deadlines for missing faculty`;
    } else {
      message += `✅ All faculty have deadlines configured!`;
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Faculty Deadline Validation', message, ui.ButtonSet.OK);
    
    return validation;
    
  } catch (error) {
    console.error('❌ Error in validateAllFacultyDeadlinesUI:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Validation Error', `Error during validation: ${error.message}`, ui.ButtonSet.OK);
    return null;
  }
}


// ===== OVERRIDE ORIGINAL FUNCTIONS TO USE ENHANCED VERSIONS =====
function sendEmails() {
  return sendEmailsWithDeadlineValidation();
}

// Replace the original checkOverdueFaculties function  
function checkOverdueFaculties() {
  return checkOverdueFacultiesWithValidation();
}

// Replace the original createMenus function
function createMenus() {
  return createEnhancedMenus();
}
// ===== UTILITY FUNCTIONS FOR COMPLETION NOTIFICATION MANAGEMENT =====

function checkCompletionStatus() {
  try {
    console.log("🔍 Checking current completion status...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!mainSheet) {
      return { error: "Main sheet not found" };
    }
    
    const data = mainSheet.getDataRange().getValues();
    if (data.length < 2) {
      return { error: "No data in main sheet" };
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const columnIndices = getColumnIndices(headers);
    
    const facultyStats = new Map();
    let totalStudents = 0;
    let totalVerified = 0;
    
    // Collect statistics
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[columnIndices.email] ? row[columnIndices.email].toString().trim().toLowerCase() : '';
      const verified = row[columnIndices.verified] ? row[columnIndices.verified].toString().trim() : '';
      
      if (!email) continue;
      
      if (!facultyStats.has(email)) {
        facultyStats.set(email, {
          totalStudents: 0,
          verifiedStudents: 0
        });
      }
      
      const faculty = facultyStats.get(email);
      faculty.totalStudents++;
      totalStudents++;
      
      if (verified === '✅') {
        faculty.verifiedStudents++;
        totalVerified++;
      }
    }
    
    // Check completion
    let allCompleted = true;
    const incompleteFaculties = [];
    const completeFaculties = [];
    
    for (const [email, stats] of facultyStats) {
      const completionRate = stats.totalStudents > 0 ? Math.round((stats.verifiedStudents / stats.totalStudents) * 100) : 0;
      
      if (stats.verifiedStudents < stats.totalStudents) {
        allCompleted = false;
        incompleteFaculties.push({
          email: email,
          completed: stats.verifiedStudents,
          total: stats.totalStudents,
          percentage: completionRate
        });
      } else {
        completeFaculties.push({
          email: email,
          completed: stats.verifiedStudents,
          total: stats.totalStudents,
          percentage: completionRate
        });
      }
    }
    
    const overallCompletionRate = totalStudents > 0 ? Math.round((totalVerified / totalStudents) * 100) : 0;
    const notificationSent = hasCompletionNotificationBeenSent();
    
    const status = {
      allCompleted: allCompleted,
      totalFaculties: facultyStats.size,
      completedFaculties: completeFaculties.length,
      incompleteFaculties: incompleteFaculties.length,
      totalStudents: totalStudents,
      totalVerified: totalVerified,
      overallCompletionRate: overallCompletionRate,
      notificationAlreadySent: notificationSent,
      completeFacultyList: completeFaculties,
      incompleteFacultyList: incompleteFaculties,
      timestamp: new Date().toLocaleString()
    };
    
    console.log("📊 Current Status:", status);
    return status;
    
  } catch (error) {
    console.error("❌ Error checking completion status:", error);
    return { error: error.message };
  }
}

// Function to display completion status in UI (for menu access)
function showCompletionStatusUI() {
  try {
    if (!isUIAvailable()) {
      console.log("UI not available - showing status in console only");
      const status = checkCompletionStatus();
      console.log("Status:", JSON.stringify(status, null, 2));
      return status;
    }
    
    const status = checkCompletionStatus();
    
    if (status.error) {
      safeUIAlert("Completion Status Error", `Error: ${status.error}`);
      return status;
    }
    
    let message = `📊 ASSESSMENT COMPLETION STATUS\n\n`;
    message += `📈 Overall Progress:\n`;
    message += `• Total Faculties: ${status.totalFaculties}\n`;
    message += `• Completed Faculties: ${status.completedFaculties}\n`;
    message += `• Incomplete Faculties: ${status.incompleteFaculties}\n`;
    message += `• Overall Completion: ${status.overallCompletionRate}%\n`;
    message += `• Total Students: ${status.totalStudents}\n`;
    message += `• Verified Students: ${status.totalVerified}\n\n`;
    
    if (status.allCompleted) {
      message += `🎉 STATUS: ALL ASSESSMENTS COMPLETED! ✅\n\n`;
      message += `📧 Admin Notification:\n`;
      message += `• Already Sent: ${status.notificationAlreadySent ? 'YES ✅' : 'NO ❌'}\n`;
      if (status.notificationAlreadySent) {
        message += `• Action: No further emails will be sent\n`;
      } else {
        message += `• Action: Will be sent on next check\n`;
      }
    } else {
      message += `⏳ STATUS: ASSESSMENTS IN PROGRESS\n\n`;
      message += `❌ Incomplete Faculties (${status.incompleteFaculties}):\n`;
      status.incompleteFacultyList.slice(0, 5).forEach((faculty, index) => {
        message += `${index + 1}. ${faculty.email}: ${faculty.completed}/${faculty.total} (${faculty.percentage}%)\n`;
      });
      if (status.incompleteFaculties > 5) {
        message += `... and ${status.incompleteFaculties - 5} more\n`;
      }
    }
    
    message += `\n📅 Status checked: ${status.timestamp}`;
    
    safeUIAlert("Assessment Completion Status", message);
    return status;
    
  } catch (error) {
    const errorMsg = `Error checking status: ${error.message}`;
    console.error("❌ Error in showCompletionStatusUI:", error);
    safeUIAlert("Status Check Error", errorMsg);
    return { error: errorMsg };
  }
}

// Function to manually trigger completion check (for testing)
function manualCompletionCheck() {
  try {
    console.log("🔄 Manual completion check triggered...");
    
    const status = checkCompletionStatus();
    
    if (status.error) {
      console.error("❌ Status check failed:", status.error);
      return status;
    }
    
    console.log("📊 Current status:", status);
    
    if (status.allCompleted && !status.notificationAlreadySent) {
      console.log("🎉 All completed and notification not sent - triggering notification...");
      checkAllFacultiesCompleted();
      return { message: "Completion notification triggered", status: status };
    } else if (status.allCompleted && status.notificationAlreadySent) {
      console.log("ℹ️ All completed but notification already sent");
      return { message: "All completed - notification already sent", status: status };
    } else {
      console.log("⏳ Not all assessments completed yet");
      return { message: "Assessments still in progress", status: status };
    }
    
  } catch (error) {
    console.error("❌ Error in manual completion check:", error);
    return { error: error.message };
  }
}

// Function to view notification history
function viewNotificationHistory() {
  try {
    console.log("📜 Viewing notification history...");
    
    const properties = PropertiesService.getScriptProperties().getProperties();
    const completionKeys = [];
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('admin_notified_') || key === 'ASSESSMENT_COMPLETION_NOTIFIED') {
        completionKeys.push({
          key: key,
          value: properties[key],
          type: key.startsWith('admin_notified_all_complete_') ? 'Permanent Completion' :
                key === 'ASSESSMENT_COMPLETION_NOTIFIED' ? 'Simple Flag' : 'Daily (Old)'
        });
      }
    });
    
    if (completionKeys.length === 0) {
      console.log("📭 No notification history found");
      return { message: "No notification history found", keys: [] };
    }
    
    console.log("📧 Notification history:");
    completionKeys.forEach((item, index) => {
      console.log(`${index + 1}. ${item.type}: ${item.key} = ${item.value}`);
    });
    
    if (isUIAvailable()) {
      let message = `📧 NOTIFICATION HISTORY\n\n`;
      message += `Found ${completionKeys.length} notification record(s):\n\n`;
      
      completionKeys.forEach((item, index) => {
        message += `${index + 1}. ${item.type}\n`;
        message += `   Key: ${item.key}\n`;
        message += `   Value: ${item.value}\n\n`;
      });
      
      if (completionKeys.some(k => k.type === 'Permanent Completion')) {
        message += `✅ Permanent completion notification has been sent\n`;
        message += `🔒 Future duplicate emails are prevented\n`;
      } else {
        message += `⚠️ No permanent completion notification found\n`;
        message += `📧 Completion email can still be sent\n`;
      }
      
      safeUIAlert("Notification History", message);
    }
    
    return { message: "Notification history retrieved", keys: completionKeys };
    
  } catch (error) {
    console.error("❌ Error viewing notification history:", error);
    return { error: error.message };
  }
}

// Function to add to menus
function addCompletionManagementToMenu() {
  try {
    if (!isUIAvailable()) {
      console.log("UI not available - skipping menu creation");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Add to existing menu or create new one
    ui.createMenu("Assessment Status")
      .addItem("📊 Check Completion Status", "showCompletionStatusUI")
      .addItem("🔄 Manual Completion Check", "manualCompletionCheck")
      .addItem("📜 View Notification History", "viewNotificationHistory")
      .addSeparator()
      .addItem("🔄 Reset Completion Flags (Testing)", "resetCompletionNotification")
      .addToUi();
      
    console.log("✅ Assessment Status menu added");
    
  } catch (error) {
    console.error("❌ Error adding menu:", error);
  }
}
// ===== END OF GOOGLE APPS SCRIPT BACKEND CODE =====