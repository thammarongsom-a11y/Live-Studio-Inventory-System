// 1. ดึงข้อมูลผู้ใช้งานและสิทธิ์
function getCurrentUserInfo() {
  var email = Session.getActiveUser().getEmail();
  if (!email.endsWith("@dappmaker.co.th")) {
    throw new Error("403_UNAUTHORIZED: สิทธิ์การเข้าใช้งานจำกัดเฉพาะบุคลากรภายใน @dappmaker.co.th เท่านั้น");
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users"); 
  var userData = userSheet.getDataRange().getValues();
  var nickname = email.split("@")[0]; 
  var role = "Agent"; 
  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0] === email) { nickname = userData[i][1]; role = userData[i][2]; break; }
  }
  return { email: email, nickname: nickname, role: role };
}

// 2. ดึงรายการสินค้าทั้งหมด
function getStockList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stock");
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; 
    items.push({
      id: data[i][0], name: data[i][1], qty: data[i][2],
      img: data[i][3] || "https://placehold.co/300x200?text=No+Image",
      category: data[i][4] || "Uncategorized", addedBy: data[i][5] || "-"
    });
  }
  return items;
}

// 3. เติมสต็อกสินค้าเดิมที่มีอยู่แล้วในคลัง
function restockExistingProduct(productId, addQty) {
  var userInfo = getCurrentUserInfo();
  if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์เติมสต็อกสินค้า" };
  if (addQty <= 0) return { success: false, message: "❌ จำนวนที่เติมต้องมากกว่า 0" };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stock");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === productId.toString()) {
      var currentQty = parseInt(data[i][2]) || 0;
      var newQty = currentQty + parseInt(addQty);
      sheet.getRange(i + 1, 3).setValue(newQty);
      SpreadsheetApp.flush(); 
      return { success: true, message: "➕ เติมสต็อกพัสดุรหัส " + productId + " เพิ่มจำนวน " + addQty + " ชิ้นสำเร็จ! (ยอดรวมใหม่: " + newQty + " ชิ้น)" };
    }
  }
  return { success: false, message: "❌ ไม่พบรหัสสินค้าพัสดุในคลัง" };
}

function getNextProductId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stock");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "PROD-001";
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var idStr = data[i][0].toString();
    var match = idStr.match(/\d+/);
    if (match) {
      var num = parseInt(match[0]);
      if (num > maxNum) maxNum = num;
    }
  }
  return "PROD-" + (maxNum + 1).toString().padStart(3, '0');
}

// 4. เพิ่มสินค้าใหม่แกะกล่อง
function addNewProduct(name, qty, img, category) {
  var userInfo = getCurrentUserInfo();
  if (userInfo.role === "Agent") return { success: false, message: "คุณไม่มีสิทธิ์เพิ่มสินค้า" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stock");
  var generatedId = getNextProductId();
  sheet.appendRow([generatedId, name, qty, img, category, userInfo.nickname]);
  SpreadsheetApp.flush();
  return { success: true, message: "📦 เพิ่มสินค้าใหม่สำเร็จ! รหัสสินค้าคือ: " + generatedId };
}

// 5. ลบสินค้าออกจากคลังสินค้าถาวร
function deleteProduct(productId) {
  var userInfo = getCurrentUserInfo();
  if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์ลบสินค้าออกจากคลัง" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stock");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === productId.toString()) {
      sheet.deleteRow(i + 1); 
      SpreadsheetApp.flush();
      return { success: true, message: "🗑️ ลบสินค้ารหัส " + productId + " ออกจากคลังเรียบร้อยแล้ว" };
    }
  }
  return { success: false, message: "❌ ไม่พบสินค้าที่ต้องการลบ" };
}

// 6. ส่งคำขอเบิก/ยืมสินค้า
function processRequisition(form) {
  var userInfo = getCurrentUserInfo(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log");
  var stockSheet = ss.getSheetByName("Stock");
  var stockData = stockSheet.getDataRange().getValues();
  var itemName = "ไม่ระบุ";
  for(var i = 1; i < stockData.length; i++) {
    if(stockData[i][0].toString() === form.itemID.toString()) { itemName = stockData[i][1]; break; }
  }
  var rowID = "REQ-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
  var timestamp = new Date();
  
  sheet.appendRow([timestamp, rowID, userInfo.nickname, form.itemID, itemName, form.qty, form.reason, "Pending", userInfo.email, "", "", 0, "", form.requestType, "", "", ""]);
  SpreadsheetApp.flush(); 
  return { success: true };
}

// 7. ดึงข้อมูลรายการเบิกส่วนกลางทั้งหมด
function getActiveRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log");
  var data = sheet.getDataRange().getValues();
  var requests = [];
  
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][1]) continue; 
    
    var rawDate = data[i][0];
    var isoString = "";
    var displayDate = "-";
    
    if (rawDate instanceof Date) {
      isoString = rawDate.toISOString();
      displayDate = Utilities.formatDate(rawDate, "GMT+7", "dd/MM/yyyy HH:mm");
    } else if (rawDate) {
      var d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        isoString = d.toISOString();
        displayDate = Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm");
      }
    }

    var retDateRaw = data[i][14];
    var retDateStr = "-";
    if (retDateRaw instanceof Date) {
      retDateStr = Utilities.formatDate(retDateRaw, "GMT+7", "dd/MM/yyyy HH:mm");
    } else if (retDateRaw) {
      retDateStr = retDateRaw.toString();
    }

    requests.push({
      isoDate: isoString, 
      timestamp: displayDate, 
      rowID: data[i][1], name: data[i][2], itemID: data[i][3], itemName: data[i][4], 
      qty: data[i][5], reason: data[i][6], status: data[i][7], email: data[i][8],
      approver: data[i][9] || "-", receiver: data[i][10] || "-",
      returnQty: data[i][11] !== "" ? parseInt(data[i][11]) : 0, 
      returnReason: data[i][12] || "-",
      requestType: data[i][13] || "เบิกพัสดุ", 
      returnDate: retDateStr,
      pendingReturnQty: data[i][15] !== "" && data[i][15] !== undefined ? parseInt(data[i][15]) : 0,
      pendingReturnReason: data[i][16] || "-"
    });
  }
  return requests;
}

// 8. อัปเดตสถานะ + ระบบส่งอีเมลแจ้งเตือน
function updateStatus(rowID, action, extraData) {
  var userInfo = getCurrentUserInfo();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log");
  var stockSheet = ss.getSheetByName("Stock");
  var logData = logSheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < logData.length; i++) {
    if (logData[i][1].toString() === rowID.toString()) { rowIndex = i + 1; break; }
  }
  if (rowIndex == -1) return { success: false, message: "❌ ไม่พบรายการนี้ในระบบ" };
  
  var itemID = logSheet.getRange(rowIndex, 4).getValue();         
  var qty = parseInt(logSheet.getRange(rowIndex, 6).getValue()) || 0;  
  var itemName = logSheet.getRange(rowIndex, 5).getValue();       
  var requesterEmail = logSheet.getRange(rowIndex, 9).getValue();
  var requesterName = logSheet.getRange(rowIndex, 3).getValue();
  var requestType = logSheet.getRange(rowIndex, 14).getValue() || "เบิกพัสดุ";

  if (action === "Approve") {
    if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์อนุมัติ" };
    var stockData = stockSheet.getDataRange().getValues();
    var itemRow = -1;
    for (var j = 1; j < stockData.length; j++) {
      if (stockData[j][0].toString() === itemID.toString()) { itemRow = j + 1; break; }
    }
    if (itemRow == -1) return { success: false, message: "❌ ไม่พบสินค้าในคลัง" };
    var currentQty = parseInt(stockSheet.getRange(itemRow, 3).getValue()) || 0;
    if (currentQty < qty) return { success: false, message: "❌ สินค้าคงเหลือไม่เพียงพอ" };
    
    stockSheet.getRange(itemRow, 3).setValue(currentQty - qty); 
    logSheet.getRange(rowIndex, 8).setValue("Approved");       
    logSheet.getRange(rowIndex, 10).setValue(userInfo.nickname); 
    SpreadsheetApp.flush();
    
    sendNotificationEmail(requesterEmail, "🟢 คำขอ " + requestType + " ของคุณได้รับการอนุมัติแล้ว", 
      "เรียนคุณ " + requesterName + ",\n\nรายการคำขอ " + requestType + " : " + itemName + " จำนวน " + qty + " ชิ้น ได้รับการอนุมัติโดยคุณ " + userInfo.nickname + " แล้ว\nกรุณามารับของที่คลังพัสดุและกดปุ่มยืนยันการรับของบนหน้าเว็บระบบครับ");
    
    return { success: true, message: "🟢 อนุมัติคำขอเรียบร้อยแล้ว ผู้อนุมัติ: " + userInfo.nickname };

  } else if (action === "Reject") {
    if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์ปฏิเสธ" };
    logSheet.getRange(rowIndex, 8).setValue("Rejected");       
    logSheet.getRange(rowIndex, 10).setValue(userInfo.nickname); 
    logSheet.getRange(rowIndex, 13).setValue(extraData); 
    SpreadsheetApp.flush();
    
    sendNotificationEmail(requesterEmail, "🔴 คำขอ " + requestType + " ของคุณถูกปฏิเสธ", 
      "เรียนคุณ " + requesterName + ",\n\nคำขอ " + requestType + " : " + itemName + " จำนวน " + qty + " ชิ้น ถูกปฏิเสธ\nโดยมีเหตุผลชี้แจง: " + extraData);
    
    return { success: true, message: "🔴 ปฏิเสธคำขอเรียบร้อยแล้ว" };

  } else if (action === "Complete") {
    logSheet.getRange(rowIndex, 8).setValue("Received");       
    logSheet.getRange(rowIndex, 11).setValue(userInfo.nickname); 
    SpreadsheetApp.flush();
    return { success: true, message: "🏁 บันทึกชื่อผู้รับของเรียบร้อยแล้ว" };

  } else if (action === "Return") {
    var inputRetQty = parseInt(extraData.returnQty) || 0;
    var retReason = extraData.returnReason;
    var currentReturned = parseInt(logSheet.getRange(rowIndex, 12).getValue()) || 0;
    
    if ((currentReturned + inputRetQty) > qty) {
      return { success: false, message: "❌ จำนวนส่งคืนรวมสะสมห้ามเกินยอดที่ยืมไปจริง (" + qty + " ชิ้น)" };
    }
    if (inputRetQty <= 0) return { success: false, message: "❌ กรุณาระบุจำนวนพัสดุส่งคืนให้ถูกต้อง" };

    logSheet.getRange(rowIndex, 8).setValue("Pending Return"); 
    logSheet.getRange(rowIndex, 16).setValue(inputRetQty); 
    logSheet.getRange(rowIndex, 17).setValue(retReason);  
    SpreadsheetApp.flush();

    sendNotificationEmail(requesterEmail, "⏳ แจ้งคืนพัสดุสำเร็จ (รอ Supervisor ตรวจสอบของจริง)", 
      "เรียนคุณ " + requesterName + ",\n\nระบบได้รับแจ้งการส่งคืนพัสดุรายการ: " + itemName + " จำนวน " + inputRetQty + " ชิ้น แล้ว\nสถานะปัจจุบันเปลี่ยนเป็น 'Pending Return (รอตรวจคืน)' กรุณานำของจริงส่งมอบให้ Supervisor ครับ");

    return { success: true, message: "⏳ ยื่นแจ้งคืนพัสดุจำนวน " + inputRetQty + " ชิ้นสำเร็จ! รอ Supervisor ตรวจสอบพัสดุจริง" };

  } else if (action === "ApproveReturn") {
    if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์อนุมัติการตรวจรับพัสดุคืน" };
    
    var currentReturned = parseInt(logSheet.getRange(rowIndex, 12).getValue()) || 0;
    var pendingRetQty = parseInt(logSheet.getRange(rowIndex, 16).getValue()) || 0;
    var pendingRetReason = logSheet.getRange(rowIndex, 17).getValue() || "-";
    var newReturnedTotal = currentReturned + pendingRetQty;
    var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");

    var stockData = stockSheet.getDataRange().getValues();
    var itemRow = -1;
    for (var j = 1; j < stockData.length; j++) {
      if (stockData[j][0].toString() === itemID.toString()) { itemRow = j + 1; break; }
    }
    if (itemRow != -1) {
      var currentQty = parseInt(stockSheet.getRange(itemRow, 3).getValue()) || 0;
      stockSheet.getRange(itemRow, 3).setValue(currentQty + pendingRetQty); 
    }

    var finalStatus = (newReturnedTotal >= qty) ? "Returned" : "Partially Returned";
    var statusText = (newReturnedTotal >= qty) ? "Returned (คืนของครบแล้ว)" : "ส่งคืนยังไม่หมด (ค้างส่ง " + (qty - newReturnedTotal) + " ชิ้น)";

    logSheet.getRange(rowIndex, 8).setValue(finalStatus); 
    logSheet.getRange(rowIndex, 12).setValue(newReturnedTotal); 
    logSheet.getRange(rowIndex, 13).setValue(pendingRetReason); 
    logSheet.getRange(rowIndex, 15).setValue(nowStr); 
    logSheet.getRange(rowIndex, 16).setValue(""); 
    logSheet.getRange(rowIndex, 17).setValue("");
    SpreadsheetApp.flush(); 

    sendNotificationEmail(requesterEmail, "↩️ อนุมัติการรับคืนพัสดุเรียบร้อยแล้ว (" + finalStatus + ")", 
      "เรียนคุณ " + requesterName + ",\n\nSupervisor (" + userInfo.nickname + ") ได้ตรวจสอบและกดยืนยันการรับคืนพัสดุจริงเรียบร้อยแล้ว\nจำนวนที่รับคืนรอบนี้: " + pendingRetQty + " ชิ้น");

    return { success: true, message: "🟢 อนุมัติการคืนพัสดุและนำของเข้าคลังเรียบร้อย" };

  } else if (action === "RejectReturn") {
    if (userInfo.role === "Agent") return { success: false, message: "❌ คุณไม่มีสิทธิ์ปฏิเสธการรับพัสดุคืน" };
    
    var currentReturned = parseInt(logSheet.getRange(rowIndex, 12).getValue()) || 0;
    var revertedStatus = (currentReturned > 0) ? "Partially Returned" : "Received";

    logSheet.getRange(rowIndex, 8).setValue(revertedStatus); 
    logSheet.getRange(rowIndex, 16).setValue(""); 
    logSheet.getRange(rowIndex, 17).setValue("");
    SpreadsheetApp.flush();

    sendNotificationEmail(requesterEmail, "🔴 คำแจ้งคืนพัสดุของคุณถูกปฏิเสธ (ตรวจสอบไม่ผ่าน)", 
      "เรียนคุณ " + requesterName + ",\n\nคำแจ้งขอส่งคืนพัสดุรายการ: " + itemName + " ถูกปฏิเสธโดย Supervisor เนื่องจากตรวจพัสดุไม่ผ่าน\nเหตุผลระบุ: " + extraData);

    return { success: true, message: "🔴 ปฏิเสธคำค้างคืนเรียบร้อยแล้ว รายการดีดกลับสู่สถานะเดิม" };
  }
}

function sendNotificationEmail(toEmail, subject, bodyText) {
  try {
    MailApp.sendEmail({
      to: toEmail,
      subject: "[dappmaker Inventory System] " + subject,
      body: bodyText + "\n\n------------------------------\nระบบบริหารจัดการพัสดุภายในบริษัท (dappmaker.co.th)"
    });
  } catch(e) {
    Logger.log("Email error: " + e.toString());
  }
}

function getDashboardSummary() {
  var userInfo = getCurrentUserInfo();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stockSheet = ss.getSheetByName("Stock");
  var logSheet = ss.getSheetByName("Log");
  var stockData = stockSheet.getDataRange().getValues();
  var logData = logSheet.getDataRange().getValues();
  var totalItemsInStock = 0, outOfStockCount = 0, totalApprovedRequests = 0, totalPendingRequests = 0, itemUsage = {};

  for (var i = 1; i < stockData.length; i++) {
    if(!stockData[i][0]) continue;
    var qty = parseInt(stockData[i][2]) || 0;
    totalItemsInStock += qty; if (qty === 0) outOfStockCount++;
  }
  for (var j = 1; j < logData.length; j++) {
    var status = logData[j][7]; var requesterName = logData[j][2]; var itemName = logData[j][4]; var itemQty = parseInt(logData[j][5]) || 0; 
    if (userInfo.role === "Agent" && requesterName !== userInfo.nickname) { continue; }
    if (status === "Pending") totalPendingRequests++;
    if (status === "Approved") totalApprovedRequests++;
    if (status === "Received") itemUsage[itemName] = (itemUsage[itemName] || 0) + itemQty;
  }
  var sortedItems = [];
  for (var name in itemUsage) { sortedItems.push({ name: name, qty: itemUsage[name] }); }
  sortedItems.sort(function(a, b) { return b.qty - a.qty; });
  return { success: true, totalItemsInStock: totalItemsInStock, outOfStockCount: outOfStockCount, totalApprovedRequests: totalApprovedRequests, totalPendingRequests: totalPendingRequests, top5Items: sortedItems.slice(0, 5), viewMode: userInfo.role === "Agent" ? "ทั้งหมด" : "ทั้งหมด" };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Live Studio Inventory System').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
