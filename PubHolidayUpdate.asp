<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
PubID = Request.Form("PubID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(2) * FROM PublicHolidays where PubID = " & PubID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("HolidayName") = Request.Form("HolidayName")
rstSecond("HolidayDate") = Request.Form("HolidayDate")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

Response.redirect "Updated.asp?AppCat=7&AppSubCatID=23"
%>