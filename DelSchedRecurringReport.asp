<!-- #include file="Connections/Site.asp" -->
<!-- #include file="Includes/MySubRegions.inc" -->
<%
ReportID = Request.QueryString("ReportID")


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM ScheduledRecurringReports Where ReportID = " & ReportID
rstSecond.Open
set rstSecond = nothing	

Response.redirect("Display.asp?AppCat=10&AppSubCatID=40#Schedule")

%>