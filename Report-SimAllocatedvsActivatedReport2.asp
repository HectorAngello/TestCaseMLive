<%
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
Region = Request.QueryString("Region")
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
SaveFileName = "SimAllocatedvsActivated_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Name,Agent Code,Agent ID Number,Region,Sub Region,Mentor Name,Start Date,Total Sims Allocated,Total Sims Activated"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Sim Allocated vs Activated Report</h3>
<br>Start Date: <b><%=FormatDateTime(StartDate,1)%></b>
<br>End Date: <b><%=FormatDateTime(EndDate,1)%></b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Agent Name</th>
	<th>Agent Code</th>
	<th>Agent ID Number</th>
	<th>Region</th>
	<th>Sub Region</th>
	<th>Mentor Name</th>
	<th>Start Date</th>
	<th>Total Sims Allocated</th>
	<th>Total Sims Activated</th>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True'"

AgentSQL = AgentSQL & " and SRID in (" & SRRegionList & ")"

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

If IsDate(RecAgentEdit.Fields.Item("TediStartDate").Value) = "True" Then
TediDate = Day(RecAgentEdit.Fields.Item("TediStartDate").Value) & " " & MonthName(Month(RecAgentEdit.Fields.Item("TediStartDate").Value)) & " " & Year(RecAgentEdit.Fields.Item("TediStartDate").Value)
Else
TediDate = "N/A"
End If

TotalAllocated = 0

set RecAllocated = Server.CreateObject("ADODB.Recordset")
RecAllocated.ActiveConnection = MM_Site_STRING
RecAllocated.Source = "Select Count(SimID) as AllocatedTotal from ViewSimsAllocationDetails where TID = " & RecAgentEdit.Fields.Item("TID").Value & " and Allocated = 'True' and (AllocatedDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
RecAllocated.CursorType = 0
RecAllocated.CursorLocation = 2
RecAllocated.LockType = 3
RecAllocated.Open()
RecAllocated_numRows = 0
If IsNull(RecAllocated.Fields.Item("AllocatedTotal").Value) = false then
TotalAllocated = RecAllocated.Fields.Item("AllocatedTotal").Value
End If

TotalActivated = 0

set RecActivations = Server.CreateObject("ADODB.Recordset")
RecActivations.ActiveConnection = MM_Site_STRING
RecActivations.Source = "Select Count(ActID) as ActivatedTotal from SimActivations where TID = " & RecAgentEdit.Fields.Item("TID").Value & " and (ActivationDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
'response.write(RecActivations.Source)
RecActivations.CursorType = 0
RecActivations.CursorLocation = 2
RecActivations.LockType = 3
RecActivations.Open()
RecActivations_numRows = 0
If IsNull(RecActivations.Fields.Item("ActivatedTotal").Value) = false then
TotalActivated = RecActivations.Fields.Item("ActivatedTotal").Value
End If



If OutFormat = "B" Then
%>
<tr>


	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("IDNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=TediDate%></td>
	<td><%=TotalAllocated%></td>
	<td><%=TotalActivated%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & "," & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & ", " & TediDate & "," & TotalAllocated & "," & TotalActivated)
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