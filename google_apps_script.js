function doPost(e) {
  try {
    // Google Apps Script receives the POST data in e.postData.contents
    const data = JSON.parse(e.postData.contents);
    
    // Get the active spreadsheet and the first sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Add headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", 
        "Category", 
        "Team/Player Name", 
        "Sport", 
        "Email", 
        "Mobile",
        "Payment Method", 
        "Payment Proof URL", 
        "Members Info"
      ]);
      
      // Format header row (bold)
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold");
    }
    
    let fileUrl = "";
    
    // If a payment proof file was uploaded, decode and save to Google Drive
    if (data.file_base64 && data.file_name) {
      const decodedFile = Utilities.base64Decode(data.file_base64);
      const blob = Utilities.newBlob(decodedFile, data.file_type || 'application/octet-stream', data.file_name);
      
      // Save file to the root of your Google Drive (you can also specify a folder ID)
      const file = DriveApp.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    // Determine names and sports based on category
    let teamOrPlayerName = data.category === 'pb-single' ? data.player_name : data.team_name;
    let sport = data.category === 'pb-single' ? data.sport_single : data.sport;
    
    // Format member info into a readable string
    let membersInfo = "";
    if (data.category === 'pb-single') {
      membersInfo = `Age: ${data.player_age}, Size: ${data.player_size}`;
    } else if (data.members_json) {
      const members = JSON.parse(data.members_json);
      membersInfo = members.map((m, i) => `${i+1}. ${m.name} (Age: ${m.age}, Size: ${m.size})`).join("\n");
    }
    
    // Append the new registration to the sheet
    sheet.appendRow([
      data.submitted_at || new Date().toISOString(),
      data.category,
      teamOrPlayerName,
      sport,
      data.email,
      data.mobile,
      data.payment_method,
      fileUrl,
      membersInfo
    ]);

    // Return success response to the web app
    return ContentService.createTextOutput(JSON.stringify({"success": true}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Return error response to the web app
    return ContentService.createTextOutput(JSON.stringify({"success": false, "error": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
