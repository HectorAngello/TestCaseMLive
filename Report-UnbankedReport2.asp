<%
Region = Request.QueryString("Region")
ReportDays = Int(Request.QueryString("ReportDays"))
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
SaveFileName = "Not-Banked_Report-MCharge-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code,  First Name, Last Name, Agent Type, Mentor, Region, Sub Region, Last Banked Date, Day Count"
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
        <h3>Agent Not Banked Report MCharge</h3>
<p>Agents Not Banked For: <b><%=ReportRange%> Days</b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Agent Code</th>
	<th>First Name</th>
	<th>Last Name</th>
	<th>Agent Type</th>
	<th>Mentor</th>
	<th>Region</th>
	<th>Sub Region</th>
	<th>Last Banked Date</th>
	<th>Day Count</th>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True' and MChargeTedi = 'True' "

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
If IsNull(RecAgentEdit.Fields.Item("LastBankedDate").Value) = "False" Then
LastBankedDay = Day(RecAgentEdit.Fields.Item("LastBankedDate").Value)
If Len(LastBankedDay) = 1 Then
LastBankedDay = "0" & LastBankedDay
End If
LastBankedMonth = Month(RecAgentEdit.Fields.Item("LastBankedDate").Value)
If Len(LastBankedMonth) = 1 then
LastBankedMonth = "0" & LastBankedMonth
End If
LastBankedDate = LastBankedDay & "/" & LastBankedMonth & "/" & Year(RecAgentEdit.Fields.Item("LastBankedDate").Value)

TediType = "Agent"
If RecAgentEdit.Fields.Item("TediParent").Value <> 0 Then
TediType = "Sub-Agent"
End If

ShowTedi = "No"

LastBankedDay = 0
TodayDate = Now()
If ReportDays <> 0 Then
If DateDiff("d",LastBankedDate, TodayDate) >= ReportDays Then
ShowTedi = "Yes"
LastBankedDay = DateDiff("d",LastBankedDate, TodayDate)
End If
Else
If DateDiff("d",LastBankedDate, TodayDate) > 13 Then
ShowTedi = "Yes"
LastBankedDay = DateDiff("d",LastBankedDate, TodayDate)
End If
End If
'Response.write(DateDiff("d",LastBankedDate, TodayDate) & " - ")

If ShowTedi = "Yes" Then
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(TediType)%></td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASlastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=LastBankedDate%></td>
	<td><%=LastBankedDay%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & "," & RecAgentEdit.Fields.Item("TediLastName").Value & "," & TediType & ", " & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASlastName").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & ", " & LastBankedDate & ", " & LastBankedDay)
End If
End If
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