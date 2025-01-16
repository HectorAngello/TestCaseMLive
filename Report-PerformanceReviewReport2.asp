<%
Region = Request.QueryString("Region")
StartMonth = Request.QueryString("StartMonth")
StartYear = Right(StartMonth,4)
StartMonth = Replace(StartMonth, "-" & StartYear, "")
EndMonth = Request.QueryString("EndMonth")
EndYear = Right(EndMonth,4)
EndMonth = Replace(EndMonth, "-" & EndYear, "")


OutFormat = Request.QueryString("OutFormat")
%>
<!-- #include file="includes/header.asp" -->
<%


If Region = "0" then
WR = "All Regions"
Else
set RecWR = Server.CreateObject("ADODB.Recordset")
RecWR.ActiveConnection = MM_Site_STRING
RecWR.Source = "SELECT * FROM [Regions] Where CompanyID = " & Session("CompanyID") & " and RID = " & Region
RecWR.CursorType = 0
RecWR.CursorLocation = 2
RecWR.LockType = 3
RecWR.Open()
RecWR_numRows = 0
WR = RecWR.Fields.Item("RegionName").Value
End If

SubRegionQry = "Select * from ViewUserSubRegions where CompanyID = " & Session("CompanyID") & " and UserID = " & Session("UNID")

If Region = "0" then
Else
SubRegionQry = SubRegionQry & " and RID = " & Region
End If

'response.write(SubRegionQry)
set RecRegions = Server.CreateObject("ADODB.Recordset")
RecRegions.ActiveConnection = MM_Site_STRING
RecRegions.Source = SubRegionQry
RecRegions.CursorType = 0
RecRegions.CursorLocation = 2
RecRegions.LockType = 3
RecRegions.Open()
RecRegions_numRows = 0
While Not RecRegions.EOF
SRRegionList = SRRegionList & RecRegions.Fields.Item("SRID").Value & ","
RecRegions.MoveNext
Wend
TempLenSRRegionList = Len(SRRegionList)
SRRegionList = Left(SRRegionList,TempLenSRRegionList - 1)

If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "PerformanceReview_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Region,Agent Code,Status,Mentor,Agent Name,Start Date,Term Date,MSISDN,ID,Last Deposit Date,Last Deposit Amount"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Airtime Sales Report</h3>
<p>Month Range: <b><%=MonthName(StartMonth) & " " & StartYear%>&nbsp;to&nbsp;<%=MonthName(EndMonth) & " " & EndYear%></b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Region</th>
	<th>Agent Code</th>
	<th>Status</th>
	<th>Mentor</th>
	<th>Agent Name</th>
	<th>Start Date</th>
	<th>Term Date</th>
	<th>MSISDN</th>
	<th>ID</th>
	<th>Last Deposit Date</th>
	<th>Last Deposit Amount</th>
<%
HeadStartMonth = StartMonth
HeadStartYear = StartYear
Stopheading = "No"

Do While Stopheading = "No"
If Int(HeadStartMonth) = Int(EndMonth) Then
If Int(HeadStartYear) = Int(EndYear) Then
Stopheading = "Yes"
End If
End If

Response.write("<th>Airtime " & MonthName(HeadStartMonth,True) & " " & HeadStartYear & "</th>")
response.flush

HeadStartMonth = HeadStartMonth + 1
If HeadStartMonth = 13 then
HeadStartMonth = 1
HeadStartYear = HeadStartYear + 1
End If




Loop
%>
<%
HeadStartMonth = StartMonth
HeadStartYear = StartYear
Stopheading = "No"

Do While Stopheading = "No"
If Int(HeadStartMonth) = Int(EndMonth) Then
If Int(HeadStartYear) = Int(EndYear) Then
Stopheading = "Yes"
End If
End If
Response.write("<th>Data " & MonthName(HeadStartMonth,True) & " " & HeadStartYear & "</th>")
response.flush

HeadStartMonth = HeadStartMonth + 1
If HeadStartMonth = 13 then
HeadStartMonth = 1
HeadStartYear = HeadStartYear + 1
End If

Loop
%>
<%
HeadStartMonth = StartMonth
HeadStartYear = StartYear
Stopheading = "No"

Do While Stopheading = "No"
If Int(HeadStartMonth) = Int(EndMonth) Then
If Int(HeadStartYear) = Int(EndYear) Then
Stopheading = "Yes"
End If
End If
Response.write("<th>Connections " & MonthName(HeadStartMonth,True) & " " & HeadStartYear & "</th>")
response.flush

HeadStartMonth = HeadStartMonth + 1
If HeadStartMonth = 13 then
HeadStartMonth = 1
HeadStartYear = HeadStartYear + 1
End If

Loop
%>
<%
HeadStartMonth = StartMonth
HeadStartYear = StartYear
Stopheading = "No"

Do While Stopheading = "No"
If Int(HeadStartMonth) = Int(EndMonth) Then
If Int(HeadStartYear) = Int(EndYear) Then
Stopheading = "Yes"
End If
End If
Response.write("<th>Ports " & MonthName(HeadStartMonth,True) & " " & HeadStartYear & "</th>")
response.flush

HeadStartMonth = HeadStartMonth + 1
If HeadStartMonth = 13 then
HeadStartMonth = 1
HeadStartYear = HeadStartYear + 1
End If

Loop
%>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where SRID in (" & SRRegionList & ")"

AgentSQl = AgentSQl & " Order By TediEmpCode Asc"
'Response.write(AgentSQl)
set RecAgentEdit = Server.CreateObject("ADODB.Recordset")
RecAgentEdit.ActiveConnection = MM_Site_STRING
RecAgentEdit.Source = AgentSQl
RecAgentEdit.CursorType = 0
RecAgentEdit.CursorLocation = 2
RecAgentEdit.LockType = 3
RecAgentEdit.Open()
RecAgentEdit_numRows = 0
While Not RecAgentEdit.EOF
TID = RecAgentEdit.Fields.Item("TID").Value
AgentStatus = "Active"
If RecAgentEdit.Fields.Item("TediActive").Value = "False" Then
AgentStatus = "In-Active"
End If

If IsNull(RecAgentEdit.Fields.Item("TediStartDate").Value) = "True" Then
AgentStartDate = "N/A"
Else
ASDDay = Day(RecAgentEdit.Fields.Item("TediStartDate").Value)
If Len(ASDDay) = 1 Then
ASDDay = "0" & ASDDay
End If
ASDMonth = Month(RecAgentEdit.Fields.Item("TediStartDate").Value)
If Len(ASDMonth) = 1 Then
ASDMonth = "0" & ASDMonth
End If
AgentStartDate = ASDDay & "/" & ASDMonth & "/" & Year(RecAgentEdit.Fields.Item("TediStartDate").Value)
End If

If IsNull(RecAgentEdit.Fields.Item("TediTermDate").Value) = "True" Then
AgentEndDate = "N/A"
Else
ASTDay = Day(RecAgentEdit.Fields.Item("TediTermDate").Value)
If Len(ASTDay) = 1 Then
ASTDay = "0" & ASTDay
End If
ASTMonth = Month(RecAgentEdit.Fields.Item("TediTermDate").Value)
If Len(ASTMonth) = 1 Then
ASTMonth = "0" & ASTMonth
End If
AgentEndDate = ASTDay & "/" & ASTMonth & "/" & Year(RecAgentEdit.Fields.Item("TediTermDate").Value)
End If

LastTransDate = "N/A"
LastTransAmount = "0"
set RecLasDeposit = Server.CreateObject("ADODB.Recordset")
RecLasDeposit.ActiveConnection = MM_Site_STRING
RecLasDeposit.Source = "EXECUTE SPLastTransAction @TID = " & TID & ", @ctype = 2"
RecLasDeposit.CursorType = 0
RecLasDeposit.CursorLocation = 2
RecLasDeposit.LockType = 3
RecLasDeposit.Open()
RecLasDeposit_numRows = 0
If Not RecLasDeposit.EOF and Not RecLasDeposit.BOF Then
LastTransDateDay = Day(RecLasDeposit.Fields.Item("CDate").Value)
If Len(LastTransDateDay) = 1 Then
LastTransDateDay = "0" & LastTransDateDay
End If
LastTransDateMonth = Month(RecLasDeposit.Fields.Item("CDate").Value)
If Len(LastTransDateMonth) = 1 Then
LastTransDateMonth = "0" & LastTransDateMonth
End If

LastTransDate = LastTransDateDay & "/" &  LastTransDateMonth  & "/" & Year(RecLasDeposit.Fields.Item("CDate").Value)
LastTransAmount = Replace(RecLasDeposit.Fields.Item("CAmount").Value, ",", ".")
End If

If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("RegionCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=AgentStatus%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(AgentStartDate)%></td>
	<td><%=AgentEndDate%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediCell").Value)%></td>
	<td><%=RecAgentEdit.Fields.Item("IDNumber").Value%></td>
	<td><%=LastTransDate%></td>
	<td><%=LastTransAmount%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & AgentStatus & "," & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value & "," & AgentStartDate & "," & AgentEndDate & "," & RecAgentEdit.Fields.Item("TediCell").Value & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & LastTransDate & "," & LastTransAmount)
End If
RecAgentEdit.MoveNext
Wend


If OutFormat = "B" Then
%>
</tbody>
</table>
<%
Else
response.redirect("Reports/" & SaveFileName)
End If
%>