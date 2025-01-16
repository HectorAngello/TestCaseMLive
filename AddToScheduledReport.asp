<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
' Checking if Schedule Date and Time have already expired
Date2 = Request.Form("ReportDate")
Date1 = Now()
CheckExpiredDate = DateDiff("d", Date1, Date2)
If CheckExpiredDate < 0 Then
%>
	<script language="JavaScript" type="text/JavaScript">
	<!--
	  alert("Error - The date selected to run this report has already passed, please select either today or a future date / time." );
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If
If DateDiff("d", Date1, Date2) = 0 and (TimeValue(Request.Form("ReportTime")) < Timevalue(now))  Then
%>
	<script language="JavaScript" type="text/JavaScript">
	<!--
	  alert("Error - The time selected to run this report has already passed, please select a time which has not already passed.");
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If
' End Check

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
ReportRunID = Request.Form("ReportRunID")
If ReportRunID = 2 or ReportRunID = 3 or ReportRunID = 4 or ReportRunID = 6  or ReportRunID = 20 Then
ReportUserDescription = Request.Form("StartDate") & " - " & Request.Form("EndDate") & " | Region: " & WR
End If

If ReportRunID = 8 or ReportRunID = 19 Then
TempX = Request.Form("ReportDays")
If TempX = "0" Then
TempX = "More Than 14"
End If
ReportUserDescription = "Agents Not Banked In: " & TempX & " Days | Region: " & WR
OtherVariables = Request.Form("ReportDays")
End If

If ReportRunID = 5 or ReportRunID = 18 Then
AgentsTypeLabel = ""
If Request.Form("StatusType") = "1" Then
AgentsTypeLabel = "Only Active Agents"
End If
If Request.Form("StatusType") = "2" Then
AgentsTypeLabel = "Only In-Active Agents"
End If
If Request.Form("StatusType") = "3" Then
AgentsTypeLabel = "Both Active and In-Active Agents"
End If

ReportUserDescription = Request.Form("StartDate") & " - " & Request.Form("EndDate") & " | Region: " & WR & " | " & Request.Form("RepDataType") & " | " & AgentsTypeLabel
OtherVariables = Request.Form("RepDataType") & "," & Request.Form("StatusType")
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
ReportUserDescription = Request.Form("StartDate") & " - " & Request.Form("EndDate") & " | Region: " & WR & " | " & DisplayZonerType & " | " & ReconTypeLabel & " | " & TransCount & " Transactions"
OtherVariables = ZonerTypes & "," & ReconType & "," & TransCount
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ScheduledReports", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("UserID") = Replace(Request.Form("UNID"), " ", "")
rstSecond("AddedDate") = Now()
rstSecond("RunReportDate") = Request.Form("ReportDate") & " " & Request.Form("ReportTime")
rstSecond("RunReportTime") = Request.Form("ReportTime")
rstSecond("ReportRunID") = Request.Form("ReportRunID")
rstSecond("ReportVarStartDate") = Request.Form("StartDate")
rstSecond("ReportVarEndDate") = Request.Form("EndDate")
rstSecond("ReportVarRegion") = Request.Form("Region")
rstSecond("ReportVariables") = OtherVariables
rstSecond("ReportStatusID") = "1"
rstSecond("ReportEmailSent") = "False"
rstSecond("ReportEmailAddress") = Request.Form("ReportEmailAddress")
rstSecond("ReportUserDescription") = ReportUserDescription
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.redirect("Display.asp?AppCat=10&AppSubCatID=41")

%>