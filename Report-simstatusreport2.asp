<%
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
%>
<!-- #include file="includes/header.asp" -->
<%


If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "SimStatus_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Sim Number,Serial Number,Box Number,Brick Number,Import Date,Allocated Date,Mentor Name,Mentor Code,Agent Name,Agent Code,Agent ID NUmber,Region,Sub Region"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
ReportRange = ReportDays
If ReportDays = 0 Then
ReportRange = "More Than 14 "
End If
%>
        <h3>Sim Status Report</h3>
<br>Start Date: <b><%=FormatDateTime(StartDate,1)%></b>
<br>End Date: <b><%=FormatDateTime(EndDate,1)%></b>
<table>
<thead>
<tr>
	<th>Sim Number</th>
	<th>Serial Number</th>
	<th>Box Number</th>
	<th>Brick Number</th>
	<th>Import Date</th>
	<th>Allocated Date</th>
	<th>Mentor Name</th>
	<th>Mentor Code</th>
	<th>Agent Name</th>
	<th>Agent Code</th>
	<th>Agent ID Number</th>
	<th>Region</th>
	<th>Sub Region</th>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewSimsAllocationDetails where KitNo <> ''"

AgentSQL = AgentSQL & " and (AllocatedDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"

AgentSQl = AgentSQl & " Order By Kitno Asc"
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
Importdate = Day(RecAgentEdit.Fields.Item("ImportDate").Value) & " " & MonthName(Month(RecAgentEdit.Fields.Item("ImportDate").Value)) & " " & Year(RecAgentEdit.Fields.Item("ImportDate").Value)

If IsDate(RecAgentEdit.Fields.Item("AllocatedDate").Value) = "True" Then
AllocateDate = Day(RecAgentEdit.Fields.Item("AllocatedDate").Value) & " " & MonthName(Month(RecAgentEdit.Fields.Item("AllocatedDate").Value)) & " " & Year(RecAgentEdit.Fields.Item("AllocatedDate").Value)
Else
AllocateDate = "N/A"
End If
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("KitNo").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SerialNo").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("BoxNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("BrickNumber").Value)%></td>
	<td><%=Importdate%></td>
	<td><%=AllocateDate%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("IDNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("KitNo").Value & "," & RecAgentEdit.Fields.Item("SerialNo").Value & "," & RecAgentEdit.Fields.Item("BoxNumber").Value & "," & RecAgentEdit.Fields.Item("BrickNumber").Value & "," & Importdate & "," & AllocateDate & ", " & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & ", " & RecAgentEdit.Fields.Item("ASEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value)
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