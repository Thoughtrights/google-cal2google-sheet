/* Clay Webster, Thoughtrights 2024 */

/* CHANGE ME: enter the calendar year you want to compute
 *         vvvv  */
var year = 2024;


var email = Session.getActiveUser().getEmail();
var emailShort = email.split('@')[0];  /* presumes it's actually an email */
var emailDomain = email.split('@')[1]; /* presumes it's actually an email */

/* Sadly doesn't do what you really want. Not a big deal since I skip
 * full day meetings. Wire it for EST.
 * var timeZone = Session.getScriptTimeZone(); */
var timeZone = 'EST';


/*
 * The Google API for Calendar and Events are more expansive than what
 * is available via the Apps Sync API for Calendar. This is
 * unfortunate because it lacks eventTypes which would make
 * 'focusTime' and 'outOfOffice' available. So instead we can only
 * infer and not discriminate by using the invitee list where only one
 * person is invited.
 * 
 */

function etl_cal(){

    /* Access the google calendar */
    var cal = CalendarApp.getDefaultCalendar();
    Logger.log('Found %s matching calendars.', cal.getName());
    /*
     * alternative routes FWIW: 
     * var cal = CalendarApp.getCalendarsByName(email)[0];
     * var cal = CalendarApp.getOwnedCalendarById(email)
     */

    /* Grab a year's worth of events. Really. They're not chunked. */
    var yearStr = year.toString();
    var events = cal.getEvents(new Date("January 1, "+yearStr+" "+ timeZone), new Date("December 31, "+yearStr+" "+timeZone));
    console.log ("got " + events.length.toString() + " events.");

    /* Create/overwrite a tab in the active sheet as name-year */
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(emailShort+"-"+yearStr);
    if (sheet == null) {
	sheet = ss.insertSheet(emailShort+"-"+yearStr);
    } else {
	sheet.clearContents(); 
    }

    /* Create the header row */
    var header = [["Week","Name","Title","Status", "Date", "Attendee Count", "Total Hours", "Weekly Total", "", "SUM TOTAL:"]]
    var range = sheet.getRange(1,1,1,10);
    range.setValues(header);
    range.setFontWeight("bold");
    
    /* ... and the full hours sums */
    var range = sheet.getRange(1,11,1,1);
    range.setFormula('=SUM(G:G)');

    var sumWeek = 0;
    var newWeek = "";
    var off = 1;
    var row = 0;
    var j = 0;
    
    /* dump all calendar events into sheet */
    for (var i=0;i<events.length;i++) {
	
	var startTime = events[i].getStartTime();
	var endTime = events[i].getEndTime();
	var duration = ((endTime - startTime) / (3600*1000));
	var weekNumber = getWeek(startTime);
	var eventText = events[i].getTitle();
	    
	/* SKIP ALL DAY MEETINGS. exclude all-day events, we only want
	 * clearly specified time-periods */
	if (events[i].isAllDayEvent()) {
	    continue; 
	}
	/* SKIP NON-MEETING MEETINGS. exclude some types of meetings
	 * based on attendees */
	var guestList = events[i].getGuestList();
	if (guestList.length == 0) {  /* OOO or focus time? */
	    continue;
	}

	/* SKIP Declined or not added meetings. */
	var myStatus = events[i].getMyStatus().toString();
	if ((myStatus == "NO") || (myStatus == "INVITED")) { /* ENUMS are NO, YES, OWNER, MAYBE, INVITED */
	    continue;
	}
	
	/* SKIP NON-MEETINGS with yourself. They only have you invited
	 * AND you own the meeting. If you don't own the meeting and
	 * it's a 1:1 with someone else, then it's legit. If you own
	 * the meeting and it's a 1:1 you would not be listed as a
	 * guest. */
	if ((guestList.length == 1) &&
	    (guestList[0].getEmail() == email) &&
	    (myStatus == "OWNER")) {
	    continue;
	}

	/* Week number and week hours logic */
	row = j + off;
	var writeWeek = "";
	if (weekNumber != newWeek) {
	    /* write old week sum */
	    if (newWeek != "") {
		var sumCell=sheet.getRange(row-1,8,1,1);
		sumCell.setValue(sumWeek.toFixed(2));
	    }
	    sumWeek = 0;
	    newWeek = weekNumber;
	    writeWeek = weekNumber;
	    off++;
	}
	row = j + off;
	
	/* tally this shit if we got this far */
	sumWeek += duration;
	
	/* write out event details */
	var details=[[writeWeek, emailShort, eventText, myStatus, new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate()), guestList.length, duration.toFixed(2)]];
	var range=sheet.getRange(row,1,1,7);
	range.setValues(details);
	var fontStyles = [[ "bold","normal","normal","normal","normal","normal","bold" ]];
	var formats = [[ "w0","","","","DDD, MMM d, yyyy","0","0.00" ]];
	range.setNumberFormats(formats);
	range.setFontWeights(fontStyles);

	j++;
	
    }
    
    /* write week number */
    if (newWeek != "") {
	var sumCell=sheet.getRange(row,8,1,1);
	sumCell.setValue(sumWeek.toFixed(2));
	sumCell.setFontWeight("bold");
    }
}


/* define the week number */
function getWeek( d ){ 
    var firstDay = new Date(d.getFullYear(),0,1); 
    return Math.ceil((((d - firstDay) / 86400000) + firstDay.getDay()+1)/7); 
} 


/*
 * convert number to col name
 * with a offset of 10 cols
 * 0 > J, 17 > AA
 */
function numToCol( p ){
  var offset = 10;
  if (p<17) {
    /* return String.fromCharCode(p+offset + 65-1); */
    return String.fromCharCode(p+offset+64);
  } else {
    var firstChar = String.fromCharCode( Math.floor((p+offset)/26) +64 );
    var secondChar = String.fromCharCode( (p+offset)%26 + 64 );
    return firstChar+''+secondChar;
  }
}


/* A little onOpen notice */
function onOpen() {
  Browser.msgBox('To use', '1) Go to Extensions > Apps Script\\n2) Edit the calendar year\\n3) Click Run > etl_cal', Browser.Buttons.OK);
}
