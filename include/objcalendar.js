
/****************************************************************************
_____________________________________________________________________________  
        
        Class:     objCalendar
        Version:  1.1
        Created: 18-Jul-2001
        Author:  Darren Neimke
        Email:    darren@showusyourcode.com
        URL:      http://www.showusyourcode.com/
_____________________________________________________________________________        
        
        Constructor:
        -----------------
        var objCalendar = new objCalendar([month(Int)], [year(Int)])
        ________________________________________________________________________
        
        Method Selectors:
        -------------------------
        BuildCalendar           - Renders a new calendar with a default of current month.
        goToCurrent             - Renders the calendar to the current month.
        toggleCalendar ( int )  - int = number of months to toggle by.
        show                     - sets the visibility of the Canvass to visible
        hide                      - sets the visibility of the Canvass to not visible
        moveTo                 - accepts x,y as pixels moves the canvass to that position
        bindToElement           - Not yet implemented.
        
        Properties Exposed:
        ---------------------------
        hasEvents   (type: Boolean) - if true then interface raises events.
        posX          allows the user to set the left location (the number of pixels from the left edge of the browser window). 
        posY          allows the user to set the top location (the number of pixels from the top edge of the browser window). 
        
        Notes:
        ---------
        If hasEvents is set to True then you need to implement the following event handlers:
            - clickhandler (d,m,y)   - receives the event thrown by the user clicking on a day/date
            - toggleCalendar (i)      - receives the event thrown by the user clicking on << or >>
            - toggleCurrent ()        - receives the event thrown by the user clicking on [ Today ]
            
        Calendar also exposes the following CSS classes:
            - clickable               : refers to the <<, [ TODAY ], and >> navigation controls at the base of the calendar
            - calendar_normal    : the normal state of a calendar date 
            - calendar_clickable  : the onMouseOver state of a calendar date
            - TR.monthsheader  : the row that contains the abbreviated Day names near the top of the calendar
            - TD.monthsheader : the individual cells contained by the aforementioned row
            
        Enjoy!!

*****************************************************************************/

/****************************************************************************
        BEGINNING OF CLASS
*****************************************************************************/
function objCalendar(m,y)
{
    
    if (typeof(_calendar_prototype_called) == 'undefined')
    {
        _calendar_prototype_called = true ;
        
        // Object methods
        this.BuildCalendar = _create ;
        this.createCanvass = _createCanvass ;
        this.showCalendar = _showCalendar ;
        this.goToCurrent = _goToCurrent ;
        this.toggleCalendar = _toggleCalendar ;
        this.hide = _hide ;
        this.show = _show ;
        this.init = _init ;
    }
    
        // Object properties
        this.name = 'default' ;
        this.rowBGColor = 'palegoldenrod' ;
        this.currentDay = 0 ;
        this.currentMonth = 0 ;
        this.currentYear = 0 ;
        this.visible = false ;
        this.isIE4 = '';
        this.isNav4 = '' ;
        
        // If you set hasEvents fo FALSE then the calendar face is dumbed out.
        this.hasEvents = true ;
        
        this.canvass = '' ; // The DIV || LAYER that we display the calendar on.
        this.bindToElement = '' ;  // Bind to an ELEMENT on the page.
    
    // Array of day names
    this.days = new Array("Sunday", "Monday", "Tuesday", "Wednesday","Thursday", 
                                   "Friday", "Saturday");
    // Array of month names
    this.months = new Array("January", "February", "March", "April", "May","June", "July", 
                                        "August", "September", "October", "November", "December");
    // Array of total days in each month
    this.totalDays = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
    
    // Call the Initialize() event.
    this.init(m,y) ;
}

function _show()
// This function displays the objects canvass.
{
    if(this.isNav4)
    {
        this.canvass.visibility = 'show' ;
    } else {
        this.canvass.style.visibility = 'visible' ;

    }
    
    
    this.visible = true ; 
}

function _hide()
// This function hides the objects canvass.
{
    if(this.isNav4)
    {
        this.canvass.visibility = 'hide' ;   
    } else {
        this.canvass.style.visibility = 'hidden' ;
    }
    this.visible = false ;
}

function _init(m,y)
{
    if (parseInt(navigator.appVersion.charAt(0)) >= 4)
    // Browser check.
    {
        this.isNav4 = (navigator.appName == "Netscape") ? true : false ;
        this.isIE4 = (navigator.appName.indexOf("Microsoft") != -1) ? true : false ;
    }
    
    // Populate the current Day|Month|Year properties
    var obj = new Date();
    this.currentDay = obj.getDate();
    this.currentMonth = obj.getMonth() + 1;
    this.currentYear = (obj.getYear() < 1000) ? obj.getYear() + 1900 : obj.getYear();
    
    /* 
        The constructor optionally accepts m && y parameters
        if none are supplied, the calendar defaults to the current
        month 
    */
    this.month = m || this.currentMonth ;
    this.year = y || this.currentYear ;
    
    // Create the canvass that we will be displaying the calendar on
    this.createCanvass() ;
    obj = null ;
  
}

function _createCanvass()
{
    
    // Create canvass for NN4+
    if (this.isNav4) 
    { 
        this.canvass = new Layer(200) ;
    }
    
    // Create canvass for IE4+
    if (this.isIE4)
    { 
        var objDiv = document.createElement("<DIV>") ;
        document.body.appendChild (objDiv) ;
        this.canvass = objDiv ;
    }
}

function _goToCurrent()
{
    
    this.year = this.currentYear ;
    this.month = this.currentMonth ;
    this.BuildCalendar() ;
}

function _toggleCalendar(n)
{
    var currentMonth = this.month ;
    var currentYear = this.year ;
    
    if((currentMonth + n) == 0)
    { 
        this.year = currentYear-1 ; 
        this.month = 12 ; 
    }
    else if((currentMonth + n) == 13)
    { 
        this.year = currentYear+1 ; 
        this.month = 1 ; 
    }
    else
    {
        this.month = currentMonth + n ;
    }
    this.BuildCalendar() ;
}

// Create and Display Calendar.
function _create()
{
    // Counters to count rows and days, String to store calendar output.
    var rowCount = 0 ;
    var numRows = 0 ;
    var sOut = new String() ;
    var greyoutdate = ""; 
    
    // Leap year correction
    if (this.year % 4 == 0 && (this.year % 100 != 0 || this.year % 400 == 0)) {
    	this.totalDays[1] = 29 ;
    }
    
    var obj = new Date(this.year, this.month-1, 1);
    var firstDayOfMonth = obj.getDay();
    obj.setDate(31);
    var lastDayOfMonth = obj.getDay();
    obj = null ;
    
//the old calender header code used to go here
   sOut = "<table border=0 cellpadding=1 cellspacing=0 class=calCalendar>" ;

    document.title = this.months[this.month-1] + " " + this.year

    /*  Write the abbreviated day names */
    sOut += "<tr class='calMonthsheader'>" ;
    for (x=0; x<7; x++) {
        sOut += "<td class='calMonthsheader'><span style='font-size: smaller'>" + this.days[x].substring(0,3) + "</span></td>" ;
    }
    sOut += "</tr>" ;

    /* Start of BODY */
    sOut += "<tr  class='bodyMain'>" ;
    numRows++ ;
    
    for (x=1; x<=firstDayOfMonth; x++) {
        /* pad the blank days at the beginning of the month. */
        rowCount++;
        sOut += "<td><span style='font-size: smaller'>&nbsp;</span></td>" ;
    }

    /* Increment the current date */
    this.dayCount=1;

    while (this.dayCount <= this.totalDays[this.month-1]){
    	/* Display new row after each 7 day block.  */
    	if (rowCount % 7 == 0) {
    		sOut += "</tr>\n<tr  class='bodyMain'>" ;
    		numRows++ ;
    	}
		/* Add handling for reserved dates */

		ubdate = true;
		
		if (!ubdate){
			if (this.currentDay == this.dayCount && this.currentMonth == this.month && this.currentYear == this.year){
				sOut += "<td align=center ><p"
				 /* Insert the Date. */	
				sOut += " CLASS='calBookedNow'>"
				sOut = sOut + this.dayCount + "</p>"
			}
			else{
				sOut += "<td align=center ><p"
				 /* Insert the Date. */	
				sOut += " CLASS='calBooked'>"
				sOut = sOut + this.dayCount + "</p>"	
			}
		}
		else {
			if (this.currentDay == this.dayCount && this.currentMonth == this.month && this.currentYear == this.year){
				sOut += "<td align=center CLASS='calNow' onMouseOver=\"this.className = 'calClickable_hover';\" onMouseOut=\"this.className = 'calNow';\"><A HREF=\"javascript:  ;\""
				 /* Insert the Date. */
				if (this.hasEvents) sOut += " onClick='clickhandler(" + this.dayCount + "," + this.month + "," + this.year + ") ;'" 
				sOut += " CLASS='calClickable'>"
				sOut = sOut + this.dayCount + "</A>"
			}	
			else {
				sOut += "<td align=center CLASS='calClickable' onMouseOver=\"this.className = 'calClickable_hover';\" onMouseOut=\"this.className = 'calClickable';\"><A HREF=\"javascript:  ;\""
				 /* Insert the Date. */
				if (this.hasEvents) sOut += " onClick='clickhandler(" + this.dayCount + "," + this.month + "," + this.year + ") ;'" 
				sOut += " CLASS='calClickable'>"
				sOut = sOut + this.dayCount + "</A>"
			}
		}
			sOut += "</td>" ;
        
        this.dayCount++ ;
        rowCount++ ;
    }

    while (rowCount % 7 != 0) {
        /* pad the blank days at the end of the month. */
        rowCount++ ;
        sOut += "<td><span style='font-size: smaller'>&nbsp;</span></td>" ;
    }
    sOut += "</tr>" ;
    // End of BODY
    
    // Write the Calendar Navigator.
    if(this.hasEvents) {
        sOut += "<tr>" ;
        sOut += "<td colspan=2 align=left>" ;
        sOut += "<A HREF=\"javascript:  ;\" onClick='toggleCalendar(-1) ; return false ;' CLASS='calClickable'  onMouseOver=\"this.className = 'calClickable_hover';\" onMouseOut=\"this.className = 'calClickable';\" TITLE='Previous Month' />&lt;&lt;</A>" ;
        sOut += "</td>" ;
        sOut += "<td colspan=3 align=center>" ;
        sOut += "<A HREF=\"javascript:  ;\" onClick='toggleCurrent() ; return false ;' CLASS='calClickable'  onMouseOver=\"this.className = 'calClickable_hover';\" onMouseOut=\"this.className = 'calClickable';\" TITLE='Current Month' />[ Today ]</A>" ;
        sOut += "</td>" ;
        sOut += "<td colspan=2 align=right>" ;
        sOut += "<A HREF=\"javascript:  ;\" onClick='toggleCalendar(1) ; return false ;' CLASS='calClickable'  onMouseOver=\"this.className = 'calClickable_hover';\" onMouseOut=\"this.className = 'calClickable';\" TITLE='Next Month' />&gt;&gt;</A>" ;
        sOut += "</td>" ;
        sOut += "</tr>" ;
    }
    
    sOut += "</table>" ;
    
    // Render the calendar
    this.showCalendar (sOut) ;
//document.write(sOut)
}

function _showCalendar (s)
{
    if(this.isNav4)
    {
        this.canvass.document.open() ;
        this.canvass.document.writeln(s) ;
        this.canvass.document.close() ;
    } else {
        this.canvass.innerHTML = s ;
    }
    this.show() ;
}

/****************************************************************************
        END OF CLASS
*****************************************************************************/
