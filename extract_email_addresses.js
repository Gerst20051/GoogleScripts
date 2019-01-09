/** This script will extract email address from your Gmail mailbox **/
/** Written by Amit Agarwal on 06/13/2013 **/

function extractEmailAddresses() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
        var ssNew = SpreadsheetApp.create('Extract Email Addresses');
        SpreadsheetApp.setActiveSpreadsheet(ssNew);
        ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    var sheet = ss.getSheets()[0];
    var monitor = sheet.getRange('A2').getValue();
    var processed = sheet.getRange('B2').getValue();
    var label = GmailApp.getUserLabelByName(processed);
    var search = 'in:' + monitor + ' -in:' + processed;
    var threads = GmailApp.search(search, 0, 50); // Process 50 Gmail threads in a batch to prevent script execution errors
    var row, messages, from, email;
    try {
        for (var x = 0; x < threads.length; x++) {
            from = threads[x].getMessages()[0].getFrom();
            from = from.match(/\S+@\S+\.\S+/g); // Use Regular Expression to extract valid email address
            if (from.length) {
                email = from[0].replace('>', '').replace('<', '');
                row = sheet.getLastRow() + 1;
                sheet.getRange(row, 1).setValue(email); // If an email address if found, add it to the sheet
            }
            threads[x].addLabel(label);
        }
    } catch (e) {
        Logger.log(e.toString());
        Utilities.sleep(5000);
    }
    if (threads.length === 0) { // All messages in the label have been processed?
        GmailApp.sendEmail(Session.getActiveUser().getEmail(), 'Extraction Done', 'Download the sheet from ' + ss.getUrl());
    }
}

function cleanList() { // Remove Duplicate Email addresses
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getRange(4, 1, sheet.getLastRow()).getValues();
    var newData = [];
    for (i in data) {
        var row = data[i];
        var duplicate = false;
        for (j in newData) {
            if (row[0] === newData[j][0]) {
                duplicate = true;
            }
        }
        if (!duplicate) {
            newData.push(row);
        }
    }
    sheet.getRange(4, 2, newData.length, newData[0].length).setValues(newData); // Put the unique email addresses in the Google sheet
}
