<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%


If Request.Form("Region") = "0" then
WR = "All Regions"
Else
set RecWR = Server.CreateObject("ADODB.Recordset")
RecWR.ActiveConnection = MM_Site_STRING
RecWR.Source = "SELECT * FROM [Regions] Where CompanyID = " & Session("CompanyID") & " and RID = " & Request.Form("Region")
RecWR.CursorType = 0
RecWR.CursorLocation = 2
RecWR.LockType = 3
RecWR.Open()
RecWR_numRows = 0
WR = RecWR.Fields.Item("RegionName").Value
End If

OtherVariables = ""
ReportRunID = Request.Form("OnceOffReportID")
If ReportRunID = 2 or ReportRunID = 3 or ReportRunID = 4 or ReportRunID = 6 Then
ReportUserDescription = "Region: " & WR
End If

If ReportRunID = 8 Then
TempX = Request.Form("ReportDays")
If TempX = "0" Then
TempX = "More Than 14"
End If
ReportUserDescription = "Agents Not Banked In: " & TempX & " Days | Region: " & WR
OtherVariables = Request.Form("ReportDays")
End If

If ReportRunID = 5 Then
ReportUserDescription = "Region: " & WR & " | " & Request.Form("RepDataType")
OtherVariables = Request.Form("RepDataType")
End If

If ReportRunID = 1 Then
ShowLabel = "All Agents"
If Request.Form("Type") = "1" Then
ShowLabel = "Terminated Agents"
End If
If Request.Form("Type") = "2" Then
ShowLabel = "Active Agents"
End If
ReportUserDescription = "Region: " & WR & " | " & ShowLabel
OtherVariables = Request.Form("Type")
End If

If ReportRunID = 7 Then
ZonerTypes = Request.Form("Display")
ReconType = Request.Form("ReconType")

If ReconType = "0" then
ReconTypeLabel = "All Recon Types"
Else
set RecWR1 = Server.CreateObject("ADODB.Recordset")
RecWR1.ActiveConnection = MM_Site_STRING
RecWR1.Source = "SELECT * FROM TediReconTypes Where RTypeID = " & ReconType
RecWR1.CursorType = 0
RecWR1.CursorLocation = 2
RecWR1.LockType = 3
RecWR1.Open()
RecWR1_numRows = 0
ReconTypeLabel = RecWR1.Fields.Item("ReconTypeLabel").Value
End If

TransCount = Int(Request.Form("TransCount"))

DisplayZonerType = ""
If ZonerTypes="1" Then
DisplayZonerType = "All Agents"
end If
If ZonerTypes="2" Then
DisplayZonerType = "All Agents With A Recon For Period"
End If
If ZonerTypes="3" Then
DisplayZonerType = "All Agents Without A Recon For Period"
End If
ReportUserDescription = "Region: " & WR & " | " & DisplayZonerType & " | " & ReconTypeLabel & " | " & TransCount & " Transactions"
OtherVariables = ZonerTypes & "," & ReconType & "," & TransCount
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ScheduledRecurringReports", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("UserID") = Request.Form("UNID")
rstSecond("AddedDate") = Now()
rstSecond("ShedRecuringTypeID") = Request.Form("ShedRecuringTypeID")
rstSecond("RunReportTime") = Request.Form("ReportTime")
rstSecond("ReportRunID") = Request.Form("ReportRunID")
rstSecond("DayOfTheWeek") = Request.Form("DayOfTheWeek")
rstSecond("DayOfTheMonth") = Request.Form("DayOfTheMonth")
rstSecond("ReportVarRegion") = Request.Form("Region")
rstSecond("ReportVariables") = OtherVariables
rstSecond("OnceOffReportID") = Request.Form("OnceOffReportID")
rstSecond("ReportEmailAddress") = Request.Form("ReportEmailAddress")
rstSecond("ReportUserDescription") = ReportUserDescription
rstSecond("DayOfTheWeek") = Request.Form("DayOfTheWeek")
rstSecond("DayOfTheMonth") = Request.Form("DayOfTheMonth")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.redirect("Display.asp?AppCat=10&AppSubCatID=40#Schedule")

%>