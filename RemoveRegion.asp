<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM UserRegion where ID = " + Replace(request.querystring("ID"), "'", "''") + ""
rstSecond.Open
set rstSecond = nothing

WhichURLSID = Session("UNID")
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecURLAudit = Server.CreateObject ( "ADODB.Recordset" )
RecURLAudit.Open "SELECT Top(1) * FROM URLAuditTrail", MM_Site_STRINGWrite, 1, 2
RecURLAudit.AddNEw
RecURLAudit("SID") = WhichURLSID
RecURLAudit("DateTime") = Now()
RecURLAudit("Server") = Session("Server")
RecURLAudit("URL") = "Http://www.mtnlive.co.za/RegionRemove.asp?ID="  & request.querystring("ID")
RecURLAudit.Update
RecURLAudit.Close

Response.Redirect("Display.asp?AppCat=3&AppSubCatID=1&ItemID=9&UserID=" & request.querystring("UserID"))
%>
