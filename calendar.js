// Only for testing
var ss = SpreadsheetApp.openById('') // Insert Calendar ID for testing - this is the only way this works for now but I'll be adding a user prompt further down the line
SpreadsheetApp.setActiveSpreadsheet(ss);
var sheet = ss.getSheets()[1];
var calId = sheet.getRange(3,11).getValue(); // get cal ID from Sheet (user inputted)
var cal = CalendarApp.getCalendarById(calId);

function findEmptyRow() { // Checks ID column to see if row exists already
    var range = sheet.getRange("F7:F"); 
    var values = range.getValues(); 
    var ct = 5;
    while (values[ct] && values[ct][0] != "") {
        ct++;
    }
    return (ct+1);
}

function eventExists(eventId) {
    var sId = eventId;
    var range = sheet.getRange("F5:F5");
    var tf = sheet.createTextFinder(sId).matchCase(false).findNext();   
    
    if (tf == null) {
        var row = findEmptyRow();
    } else {
        var row = tf.getRow();
    }
    return row;
}

function convertToDate(datetime) {
    var date = new Date(datetime);
    Logger.log(date);

}

function durationCalc(d2, d1) {
    var diff = (d2.getTime() - d1.getTime()) / 1000;
    diff /= 60;
    return Math.abs(Math.round(diff))
}

function getCalendarEvents() {
    var startDate = sheet.getRange(3,3).getValue();
    var endDate = sheet.getRange(3,4).getValue();
    var events = cal.getEvents(startDate, endDate);
    eventArray = [];

function ifNull(field) {
    if (field == null) {
        return 0;
    } else {
        return field;
    }
}

function isCancelledAppt(color, price) {
    if (color == 8 || color == 2) {
        takeHomePay = -Math.abs(price);
    } else {
        takeHomePay = price;
    }
    return takeHomePay;
}

function shopCut(earnings) {
    var percentage = 0.2
    if (earnings > 0) {
        return earnings * percentage;
    } else {
        return 0;
    }
}

function removeNegative(number) {
    if (number < 0) {
        return 0;
    } else {
        return number;
    }
}

// function convertToHours(time) {
//     return Math.abs(Math.round(time / 3.6e+6));
// }

    for (i = 0; i < events.length; i++) {
        var startDate = events[i].getStartTime();
        // var startTime = convertToHours(startDate);
        var endT = events[i].getEndTime();
        // var endTime = convertToHours(endT);
        var duration = durationCalc(endT, startDate);
        var eventTitle = events[i].getTitle().split("£", 2);
        var custName = eventTitle[0];
        var tatPrice = ifNull(eventTitle[1]);
        var eventNotes = events[i].getDescription().split("£", 2)
        var tatDesc = eventNotes[0];
        var tatDeposit = ifNull(eventNotes[1]);
        var eventColor = events[i].getColor();
        var totalEarned = isCancelledAppt(eventColor, tatPrice);
        var takeHomePay = removeNegative(totalEarned) - shopCut(totalEarned);
        var shopPerc = shopCut(totalEarned);
        var tatToPay = tatPrice - tatDeposit;
        var eventId = events[i].getId();

        var eventData = [startDate, startDate, endT, duration, custName, tatPrice, tatDesc, tatDeposit, tatToPay, totalEarned, takeHomePay, shopPerc, eventColor, eventId];

        eventArray.push(eventData);
    
    }

    Logger.log(eventArray);

    return eventArray;

        // var newRow = eventExists(eventId);
        // var range = ("A" +  newRow + ":L" + newRow);
        // Logger.log("Range is " +  range);
        // var cells = range.setValues(eventData);
}

function writeRows() {
    var values = getCalendarEvents()
    var length = values.length + 4; // Add  4 to include header rows which aren't indexed
    var range = sheet.getRange("A5:N" + length);
    range.setValues(values); 

}

function main() {
    writeRows();
}