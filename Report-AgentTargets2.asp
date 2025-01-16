<%
Region = Request.QueryString("Region")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
RType = Request.QueryString("RType")
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

TargetMonth = Month(StartDate)
TargetYear = Year(StartDate)

set RecTargets = Server.CreateObject("ADODB.Recordset")
RecTargets.ActiveConnection = MM_Site_STRING
RecTargets.Source = "SELECT Top(1)* FROM MonthlyTargets Where PeriodYear = " & TargetYear & " and PeriodMonth = " & TargetMonth
RecTargets.CursorType = 0
RecTargets.CursorLocation = 2
RecTargets.LockType = 3
RecTargets.Open()
RecTargets_numRows = 0
DailyAirtimeTarget = RecTargets.Fields.Item("AirtimeTarget").Value
DailyDataTarget = RecTargets.Fields.Item("DataTarget").Value
DailyConnectionsTarget = RecTargets.Fields.Item("ConnectionsTarget").Value

ReportDayCount = DateDiff("d",StartDate,EndDate) + 1

ReportAirtimeTarget = FormatNumber(RecTargets.Fields.Item("AirtimeTarget").Value * ReportDayCount,,,,0)
ReportDataTarget = FormatNumber(RecTargets.Fields.Item("DataTarget").Value * ReportDayCount,,,,0)
ReportConnectionsTarget = FormatNumber(RecTargets.Fields.Item("ConnectionsTarget").Value * ReportDayCount,,,,0)

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
SaveFileName = "Targets_Report-" & RType & "-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code,First Name,Last Name,Mentor,Region,Sub Region," & RType & ", Met Target"

TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline("Agent Targets Report - " & RType)
TheFile.Writeline("Date Range: " & StartDate & " to " & EndDate & " (" & ReportDayCount & " Days)")
TheFile.Writeline("Targets for this period: " & MonthName(TargetMonth) & " " & Targetyear)
If RType = "Airtime" Then
TheFile.Writeline("Airtime: R " & ReportAirtimeTarget & " (R " & DailyAirtimeTarget & " Daily)")
End If
If RType = "Data" Then
TheFile.Writeline("Data: " & ReportDataTarget & " (" & DailyDataTarget & " Daily)")
End If
If RType = "Connections" Then 
TheFile.Writeline("Sims Connections: " & ReportConnectionsTarget & "(" & DailyConnectionsTarget & " Daily)")
End If
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Agent Targets Report</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%> (<%=ReportDayCount%> Days)</b>
<br>Region: <b><%=WR%></b>
<br><br>Targets for this period: <%=MonthName(TargetMonth) & " " & Targetyear%>
<%If RType = "Airtime" Then%><br>Airtime: R <%=ReportAirtimeTarget%> (R <%=DailyAirtimeTarget%> Daily)<%End If%>
<%If RType = "Data" Then%><br>Data: <%=ReportDataTarget%> (<%=DailyDataTarget%> Daily)<%End If%>
<%If RType = "Connections" Then%><br>Sims Connections: <%=ReportConnectionsTarget%> (<%=DailyConnectionsTarget%> Daily)<%End If%>
<br><strong>Targets are based on the month the report starts in, if the report goes over multiple months, all figures will be based on the targets from the starting month.</strong>
<table>
<thead>
<tr>
	<th>Agent Code</th>
	<th>First Name</th>
	<th>Last Name</th>
	<th>Mentor</th>
	<th>Region</th>
	<th>Sub Region</th>
	<%If RType = "Airtime" Then%><th>Airtime</th><th>Met Target</th><%End If%>
	<%If RType = "Data" Then%><th>Data</th><th>Met Target</th><%End If%>
	<%If RType = "Connections" Then%><th>Sim Connections</th><th>Met Target</th><%End If%>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where SRID in (" & SRRegionList & ")"
AgentSQl = AgentSQl & " and TediActive = 'True' "
AgentSQl = AgentSQl & " Order By RegionName, SubRegionName, TediEmpCode Asc"
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
TediID = RecAgentEdit.Fields.Item("TID").Value
AgentTotal = 0
If RType = "Airtime" Then
set RecAirtime = Server.CreateObject("ADODB.Recordset")
RecAirtime.ActiveConnection = MM_Site_STRING
RecAirtime.Source = "SELECT SUM(CAmount) AS ATTotal FROM TediTransactions WHERE TediID = " & TediID & "  and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') AND (CType = 2)"
'response.write(RecAirtime.Source)
RecAirtime.CursorType = 0
RecAirtime.CursorLocation = 2
RecAirtime.LockType = 3
RecAirtime.Open()
RecAirtime_numRows = 0
If IsNULL(RecAirtime.Fields.Item("ATTotal").Value) = "False" Then
AgentTotal = RecAirtime.Fields.Item("ATTotal").Value
End If
End If

If RType = "Data" Then
set RecThisMonthsVendsData = Server.CreateObject("ADODB.Recordset")
RecThisMonthsVendsData.ActiveConnection = MM_Site_STRING
RecThisMonthsVendsData.Source = "SELECT Sum(VendAmount) AS TotalVends FROM ViewVendingDetailsOnTIDShort Where TID = " & TediID & " and (VendDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') and AmountType = 'DATA'"
'response.write(RecThisMonthsVendsData.Source)
RecThisMonthsVendsData.CursorType = 0
RecThisMonthsVendsData.CursorLocation = 2
RecThisMonthsVendsData.LockType = 3
RecThisMonthsVendsData.Open()
RecThisMonthsVendsData_numRows = 0
If IsNull(RecThisMonthsVendsData.Fields.Item("TotalVends").Value) = false then
AgentTotal = RecThisMonthsVendsData.Fields.Item("TotalVends").Value
End If
End If

If RType = "Connections" Then
set RecThisMonthsConnections = Server.CreateObject("ADODB.Recordset")
RecThisMonthsConnections.ActiveConnection = MM_Site_STRING
RecThisMonthsConnections.Source = "SELECT Count(ActID) AS TotalConnect FROM ViewSimActivationDetails Where TID = " & TediID & "  and (ActivationDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
RecThisMonthsConnections.CursorType = 0
RecThisMonthsConnections.CursorLocation = 2
RecThisMonthsConnections.LockType = 3
RecThisMonthsConnections.Open()
RecThisMonthsConnections_numRows = 0
If IsNull(RecThisMonthsConnections.Fields.Item("TotalConnect").Value) = false then
AgentTotal = RecThisMonthsConnections.Fields.Item("TotalConnect").Value
End If
End If

AgentMetTarget = "False"

If RType = "Connections" Then
If Int(AgentTotal) > Int(ReportConnectionsTarget) Then
AgentMetTarget = "True"
End If
End If

If RType = "Data" Then
If Int(AgentTotal) > Int(ReportDataTarget) Then
AgentMetTarget = "True"
End If
End If

If RType = "Airtime" Then
If Int(AgentTotal) > Int(ReportAirtimeTarget) Then
AgentMetTarget = "True"
End If
End If

TediType = "Agent"
If RecAgentEdit.Fields.Item("TediParent").Value <> 0 Then
TediType = "Sub-Agent"
End If
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASLastName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=AgentTotal%></td>
	<td><%=AgentMetTarget%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & "," & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("ASLastName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & "," & AgentTotal & "," & AgentMetTarget)
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