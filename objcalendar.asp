<%@ Language=VBScript %>

<% Option Explicit %>
<%  Response.CacheControl = "no-cache" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: objCalender.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Used to produce a pop up calendar
' ----------------------------------------------------------------------------------
%>

<HTML>
<HEAD>
<TITLE>Calendar</TITLE>

<META HTTP-EQUIV="Content-Type" CONTENT="text/html">
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="Include/objCalendar.js"></SCRIPT>
<SCRIPT>
/****************************************************************************
        THIS IS THE CODE NEEDED TO PRODUCE A CALENDAR!
*****************************************************************************/
var obj1 ;

function produceCalendar (bindToFld) {
    if(typeof(obj1) != 'object') obj1 = new objCalendar() ;
    
    // If the calendar is visible and the same button was pressed as was pressed
    // to create the Calendar then Hide the calendar and exit the Sub
    if( obj1.visible && bindToFld == obj1.bindToElement )
    {
        obj1.hide() ;
        return ;
    }else{
        obj1.bindToElement = bindToFld ;
    }

    obj1.BuildCalendar() ;
}

/****************************************************************************
        IMPLEMENT THESE 4 FUNCTIONS IF YOU WANT TO CATCH CLIENT
        SIDE EVENTS (To use this you must have hasEvents set to True) .
*****************************************************************************/
function clickhandler (d,m,y) {

  var strDay = '0' + d
  var strMonth = '0' + m
  var strYear = '20' + y
  
  if (strDay.length>2)
  {
    strDay = strDay.substr(1,2)
  }
  if (strMonth.length>2)
  {
    strMonth = strMonth.substr(1,2)
  }
  if (strYear.length>4)
  {
    strYear = strYear.substr(2,4)
  }

  switch(document.cal.tbxDateFormat.value)
  {
    case("dd/mm/yyyy"):
    	objstring = strDay + '/' + strMonth + '/' + strYear ;
      break

    case("mm/dd/yyyy"):
    	objstring = strMonth + '/' + strDay + '/' + strYear ;
      break

    case("dd-mmm-yyyy"):
  	  objstring = strDay + '-' + obj1.months[obj1.month-1].substring(0,3) + '-' + strYear ;
      break

    case("dd.mm.yyyy"):
  	  objstring = strDay + '.' + strMonth + '.' + strYear ;
      break

    case("mm.dd.yyyy"):
  	  objstring = strMonth + '.' + strDay + '.' + strYear ;
      break

    default:
    	objstring = strMonth + '/' + strDay + '/' + strYear ;
      break
  }

  // objstring = d + '-' + obj1.months[obj1.month-1].substring(0,3) + '-' + y ;
	
<%
	dim comstring
	comstring = "window.opener.changeItem(objstring)"
	response.write comstring
%>   
    window.close()
}

function reservedDate(){
	alert("This date is booked out")
}

function toggleCalendar (i) {
	document.cal.newpage.value = ""
    obj1.toggleCalendar(i) ;
    window.focus()
}

function toggleCurrent () {
	document.cal.newpage.value = ""
    obj1.goToCurrent() ;
    window.focus()
}

function calClose () {
    obj1.hide() ;
}
function runBlurTest(){
	document.cal.newpage.value = "lost focus"
	setTimeout("secondblurtest()",500)	
}
function secondblurtest(){
	if (document.cal.newpage.value =="lost focus"){
		window.close()
	}
}
/*****************************************************************************/
</SCRIPT>
</HEAD>
<LINK REL="STYLESHEET" TYPE="text/css" HREF="objCalendar.css">
<BODY BGCOLOR="#d4d0c8" onLoad="return produceCalendar('Calendar') ;"  onBlur="Javascript:runBlurTest()">
<div name='Calendar'></div>
</BODY>
<form name="cal">
<input type="hidden" name="newpage" value="">
<input type="hidden" name="tbxDateFormat" value="<%=Application("DATE_FORMAT")%>">
</form>

</HTML>
