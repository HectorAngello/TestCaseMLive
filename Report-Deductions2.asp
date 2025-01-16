<%
Region = Request.QueryString("Region")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
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
SaveFileName = "Deductions_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code,  First Name, Last Name, Agent Type, Region, Sub Region, Deduction Date, Deduction Category, Value"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Agent Deductions Report</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%></b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Agent Code</th>
	<th>First Name</th>
	<th>Last Name</th>
	<th>Agent Type</th>
	<th>Region</th>
	<th>Sub Region</th>
	<th>Deduction Date</th>
	<th>Deduction Category</th>
	<th>Value</th>
</tr>
</thead>

<tbody>
<%
End If
set RecReconRegions = Server.CreateObject("ADODB.Recordset")
RecReconRegions.ActiveConnection = MM_Site_STRING
If Region = "0" Then
RecReconRegions.Source = "SELECT Distinct RID, RegionName, SubRegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName, SubRegionName Asc"
Else
RecReconRegions.Source = "SELECT Distinct RID, RegionName, SubRegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and RID = " & Region & " order by SubRegionName"
End If
RecReconRegions.CursorType = 0
RecReconRegions.CursorLocation = 2
RecReconRegions.LockType = 3
RecReconRegions.Open()
RecReconRegions_numRows = 0
While Not RecReconRegions.EOF

AgentSQl = "SELECT * FROM ViewDeductionsDetails where RID = " & RecReconRegions.Fields.Item("RID").Value

AgentSQL = AgentSQL & " and SRID in (" & SRRegionList & ")"

AgentSQl = AgentSQl & " and TediActive = 'True' "

AgentSQl = AgentSQl & " and (DeductionDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"

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

DeductDate = Day(RecAgentEdit.Fields.Item("DeductionDate").Value) & " " & MonthName(Month(RecAgentEdit.Fields.Item("DeductionDate").Value)) & " " & Year(RecAgentEdit.Fields.Item("DeductionDate").Value)

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
	<td><%=(TediType)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=DeductDate%></td>
	<td><%=(RecAgentEdit.Fields.Item("DeductionLabel").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("DeductionValue").Value)%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & "," & RecAgentEdit.Fields.Item("TediLastName").Value & "," & TediType & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & "," & DeductDate & "," & RecAgentEdit.Fields.Item("DeductionLabel").Value & "," & RecAgentEdit.Fields.Item("DeductionValue").Value)
End If
RecAgentEdit.MoveNext
Wend

RecReconRegions.MoveNext
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