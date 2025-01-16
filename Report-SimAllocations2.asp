<%
Region = Request.QueryString("Region")
OutFormat = Request.QueryString("OutFormat")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
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
SaveFileName = "SimAllocations_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code, First Name,Last Name,ID Number," & SupervisorLabel & ",Region,Sub Region,Sims Allocated,Sims Activated,Sims Not Activated"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Sim Allocation Report</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%></b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Agent Code</th>
	<th>First Name</th>
	<th>Last Name</th>
	<th>ID Number</th>
	<th><%=SupervisorLabel%></th>
	<th>Region</th>
	<th>Sub Region</th>
	<th>Sims Allocated</th>
	<th>Sims Activated</th>
	<th>Sims Not Activated</th>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True' "

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
TotalSims = 0
TotalActived = 0
TotalNotActived = 0

set RecTotals = Server.CreateObject("ADODB.Recordset")
RecTotals.ActiveConnection = MM_Site_STRING
RecTotals.Source = "SELECT Count(dbo.BulkSimChildren.ChildID)  As SimTotal FROM dbo.BulkSims INNER JOIN dbo.BulkSimChildren ON dbo.BulkSims.BulkID = dbo.BulkSimChildren.BulkID LEFT OUTER JOIN dbo.SimActivations ON dbo.BulkSimChildren.SerialNo = dbo.SimActivations.SimNo Where dbo.BulkSimChildren.TID = " & RecAgentEdit.Fields.Item("TID").Value & "  and (ActivationDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
'RecTotals.Source = "EXECUTE SPAgentSimAllocationCount @TID = " & RecAgentEdit.Fields.Item("TID").Value
RecTotals.CursorType = 0
RecTotals.CursorLocation = 2
RecTotals.LockType = 3
RecTotals.Open()
RecTotals_numRows = 0
If IsNull(RecTotals.Fields.Item("SimTotal").Value) = "False" Then
TotalActived = RecTotals.Fields.Item("SimTotal").Value
End If

set RecTotals2 = Server.CreateObject("ADODB.Recordset")
RecTotals2.ActiveConnection = MM_Site_STRING
RecTotals2.Source = "SELECT Count(dbo.BulkSimChildren.ChildID)  As SimTotal FROM dbo.BulkSims INNER JOIN dbo.BulkSimChildren ON dbo.BulkSims.BulkID = dbo.BulkSimChildren.BulkID LEFT OUTER JOIN dbo.SimActivations ON dbo.BulkSimChildren.SerialNo = dbo.SimActivations.SimNo Where dbo.BulkSimChildren.TID = " & RecAgentEdit.Fields.Item("TID").Value & " and (ActivationDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
'RecTotals2.Source = "EXECUTE SPAgentSimNonAllocationCount @TID = " & RecAgentEdit.Fields.Item("TID").Value
RecTotals2.CursorType = 0
RecTotals2.CursorLocation = 2
RecTotals2.LockType = 3
RecTotals2.Open()
RecTotals2_numRows = 0
If IsNull(RecTotals2.Fields.Item("SimTotal").Value) = "False" Then
TotalNotActived = RecTotals2.Fields.Item("SimTotal").Value
End If


TotalSims = TotalActived + TotalNotActived
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("IDNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=TotalSims%></td>
	<td><%=TotalActived%></td>
	<td><%=TotalNotActived%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & "," & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & "," & TotalSims & "," & TotalActived & "," & TotalNotActived)
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