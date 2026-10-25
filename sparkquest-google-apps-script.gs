/*
 * SparkQuest Google Apps Script
 *
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1St7t2R6eQteEkZpPBNhE2ADjZZIRPcAYXgIfgHzvTAk/edit
 *
 * Deploy:
 * 1. Open script.google.com and create a new project.
 * 2. Paste this complete file.
 * 3. Run setupSheet() once and approve the permissions.
 * 4. Deploy > New deployment > Web app.
 * 5. Execute as: Me.
 * 6. Who has access: Anyone.
 * 7. Copy the /exec URL into CONFIG.appsScriptUrl in the HTML game.
 *
 * Payment verification:
 * - The game records every payment submission as PENDING.
 * - Verify the ₹10 payment in your UPI app using the submitted UTR when available.
 * - In the SparkQuest Leads sheet, change that payment row's Payment Status
 *   from PENDING to PAID.
 * - The game will unlock all 100 stages when the child taps Check payment status.
 */

var SPARKQUEST_SHEET_ID = "1St7t2R6eQteEkZpPBNhE2ADjZZIRPcAYXgIfgHzvTAk";
var SPARKQUEST_SHEET_NAME = "SparkQuest Leads";
var SPARKQUEST_HEADERS = [
  "Timestamp",
  "Event Type",
  "User ID",
  "Name",
  "Gender",
  "Age",
  "Mobile Number",
  "Email",
  "Current Stage",
  "Unlocked Stage",
  "Stars JSON",
  "Payment ID",
  "UTR / Transaction ID",
  "Amount",
  "Payment Status",
  "Progress JSON",
  "Notes"
];

function getSparkQuestSheet_() {
  var spreadsheet = SpreadsheetApp.openById(SPARKQUEST_SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SPARKQUEST_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SPARKQUEST_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SPARKQUEST_HEADERS.length).setValues([SPARKQUEST_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SPARKQUEST_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#6c5ce7")
      .setFontColor("#ffffff");
  }

  return sheet;
}

function setupSheet() {
  getSparkQuestSheet_();
  return jsonOutput_({
    ok: true,
    message: "SparkQuest Leads sheet is ready."
  });
}

function doPost(event) {
  try {
    var data = readPostData_(event);
    var action = String(data.action || "signup").toLowerCase();
    var sheet = getSparkQuestSheet_();

    if (action === "signup") {
      appendEvent_(sheet, {
        eventType: "SIGNUP",
        userId: data.userId,
        name: data.name,
        gender: data.gender,
        age: data.age,
        mobile: data.mobile,
        email: data.email,
        unlockedStage: data.unlockedStage || 1,
        progressJson: data.progressData,
        paymentStatus: "NOT_REQUIRED",
        notes: "New child account created."
      });
    } else if (action === "payment_submitted") {
      appendEvent_(sheet, {
        eventType: "PAYMENT_SUBMITTED",
        userId: data.userId,
        name: data.name,
        gender: data.gender,
        age: data.age,
        mobile: data.mobile,
        email: data.email,
        paymentId: data.paymentId,
        utr: data.utr,
        amount: data.amount || 10,
        paymentStatus: "PENDING",
        progressJson: data.progressData,
        notes: "UTR is optional. Verify the payment using the available details and change Payment Status to PAID after the successful ₹10 payment."
      });
    } else if (action === "progress") {
      appendEvent_(sheet, {
        eventType: "PROGRESS",
        userId: data.userId,
        name: data.name,
        gender: data.gender,
        age: data.age,
        mobile: data.mobile,
        email: data.email,
        currentStage: data.currentStage,
        unlockedStage: data.unlockedStage,
        starsJson: data.starsJson,
        paymentStatus: data.paymentStatus || "NOT_SUBMITTED",
        progressJson: data.progressData,
        notes: "Progress snapshot."
      });
    } else {
      throw new Error("Unsupported action: " + action);
    }

    return jsonOutput_({ ok: true, action: action });
  } catch (error) {
    return jsonOutput_({
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function doGet(event) {
  try {
    var params = event && event.parameter ? event.parameter : {};
    var action = String(params.action || "health").toLowerCase();
    var result;

    if (action === "login") {
      result = findUserForLogin_(params.name, params.mobile);
    } else if (action === "payment_status") {
      result = getPaymentStatus_(params.userId);
    } else {
      result = {
        ok: true,
        message: "SparkQuest service is running."
      };
    }

    return jsonOutput_(result, params.callback);
  } catch (error) {
    return jsonOutput_({
      ok: false,
      error: String(error && error.message ? error.message : error)
    }, event && event.parameter ? event.parameter.callback : "");
  }
}

function findUserForLogin_(name, mobile) {
  var sheet = getSparkQuestSheet_();
  var rows = getDataRows_(sheet);
  var requestedName = normaliseName_(name);
  var requestedMobile = normaliseMobile_(mobile);
  var signup = null;

  for (var i = rows.length - 1; i >= 0; i--) {
    var row = rows[i];
    if (
      String(row[1]).toUpperCase() === "SIGNUP" &&
      normaliseName_(row[3]) === requestedName &&
      normaliseMobile_(row[6]) === requestedMobile
    ) {
      signup = row;
      break;
    }
  }

  if (!signup) {
    return {
      ok: true,
      found: false,
      message: "No account found for these details."
    };
  }

  var userId = String(signup[2] || "");
  var user = {
    ok: true,
    found: true,
    userId: userId,
    name: String(signup[3] || ""),
    gender: String(signup[4] || "Other"),
    age: String(signup[5] || ""),
    mobile: String(signup[6] || ""),
    email: String(signup[7] || ""),
    progress: defaultProgress_(),
    paid: false,
    paymentStatus: "NOT_SUBMITTED"
  };

  rows.forEach(function (row) {
    if (String(row[2] || "") !== userId) return;

    var eventType = String(row[1] || "").toUpperCase();
    var rowPaymentStatus = String(row[14] || "").toUpperCase();

    if (eventType === "PROGRESS" && row[15]) {
      var progress = parseJson_(row[15]);
      if (progress) {
        user.progress = progress;
      }
    }

    if (eventType === "PAYMENT_SUBMITTED") {
      user.paymentStatus = rowPaymentStatus || "PENDING";
    }

    if (rowPaymentStatus === "PAID") {
      user.paid = true;
      user.paymentStatus = "PAID";
    }
  });

  if (user.paid) {
    user.progress.paid = true;
    user.progress.unlocked = 100;
    user.progress.paymentStatus = "PAID";
  }

  return user;
}

function getPaymentStatus_(userId) {
  var sheet = getSparkQuestSheet_();
  var rows = getDataRows_(sheet);
  var paid = false;
  var latestStatus = "NOT_SUBMITTED";
  var latestUtr = "";

  rows.forEach(function (row) {
    if (String(row[2] || "") !== String(userId || "")) return;
    var eventType = String(row[1] || "").toUpperCase();
    var status = String(row[14] || "").toUpperCase();

    if (eventType === "PAYMENT_SUBMITTED") {
      latestStatus = status || "PENDING";
      latestUtr = String(row[12] || "");
    }

    if (status === "PAID") {
      paid = true;
      latestStatus = "PAID";
      latestUtr = String(row[12] || latestUtr);
    }
  });

  return {
    ok: true,
    found: Boolean(userId),
    paid: paid,
    paymentStatus: latestStatus,
    utr: latestUtr,
    message: paid
      ? "Payment verified."
      : "Payment is pending verification."
  };
}

function appendEvent_(sheet, data) {
  var row = [
    new Date(),
    safeCell_(data.eventType),
    safeCell_(data.userId),
    safeCell_(data.name),
    safeCell_(data.gender),
    safeCell_(data.age),
    safeCell_(data.mobile),
    safeCell_(data.email),
    safeCell_(data.currentStage),
    safeCell_(data.unlockedStage),
    safeCell_(data.starsJson),
    safeCell_(data.paymentId),
    safeCell_(data.utr),
    safeCell_(data.amount),
    safeCell_(data.paymentStatus),
    safeCell_(data.progressJson),
    safeCell_(data.notes)
  ];

  sheet.appendRow(row);
}

function getDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, SPARKQUEST_HEADERS.length).getValues();
}

function readPostData_(event) {
  var data = {};
  if (event && event.parameter) {
    Object.keys(event.parameter).forEach(function (key) {
      data[key] = event.parameter[key];
    });
  }

  if (event && event.postData && event.postData.contents) {
    try {
      var json = JSON.parse(event.postData.contents);
      Object.keys(json).forEach(function (key) {
        data[key] = json[key];
      });
    } catch (ignored) {
      // Form-encoded requests are already available through event.parameter.
    }
  }
  return data;
}

function defaultProgress_() {
  return {
    unlocked: 1,
    stars: {},
    paid: false,
    paymentStatus: "NOT_SUBMITTED",
    paymentId: "",
    paymentUtr: ""
  };
}

function parseJson_(value) {
  try {
    return JSON.parse(String(value || ""));
  } catch (error) {
    return null;
  }
}

function normaliseName_(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function normaliseMobile_(value) {
  return String(value || "").replace(/\D/g, "").slice(-10);
}

function safeCell_(value) {
  var text = value === undefined || value === null ? "" : String(value);
  if (/^[=+\-@]/.test(text)) return "'" + text;
  return text;
}

function jsonOutput_(value, callback) {
  var json = JSON.stringify(value);
  var safeCallback = String(callback || "").match(/^[A-Za-z_$][0-9A-Za-z_$]*$/);
  if (safeCallback) {
    return ContentService
      .createTextOutput(safeCallback[0] + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
