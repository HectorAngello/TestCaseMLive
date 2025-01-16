<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRING
rstSecond.Source = "Delete FROM SeniorZonerAllocations where ID = " + Replace(request.querystring("ID"), "'", "''") + ""
rstSecond.Open
set rstSecond = nothing



Response.Redirect("Display.asp?AppCat=7&AppSubCatID=13&ItemID=55&RID=" & request.querystring("RID"))
%>
