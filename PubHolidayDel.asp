<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
PubID = Request.Form("PubID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM PublicHolidays where PubID = " & PubID
rstSecond.Open
set rstSecond = nothing	

Response.redirect "Updated.asp?AppCat=7&AppSubCatID=23"
%>