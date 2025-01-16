<%
Region = Request.QueryString("Region")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
RepDataType = Request.QueryString("RepDataType")
%>
<!-- #include file="includes/header.asp" -->
<%
'on error resume next

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

ListStartDay = StartDate
BrowserOut = ""
ExcelOut = ""
ListEndDate = DateAdd("d",1,EndDate)
Stoper = "No"
Do While Stoper = "No"
BrowserOut = BrowserOut & "<th>" & Day(ListStartDay) & " " & MonthName(Month(ListStartDay),True) &  " " & Year(ListStartDay) & "</th>"
ExcelOut = ExcelOut & ", " & Day(ListStartDay) & " " & MonthName(Month(ListStartDay)) &  " " & Year(ListStartDay)
ListStartDay = DateAdd("d",1,ListStartDay)
If Day(ListStartDay) = Day(ListEndDate) Then
If Month(ListStartDay) = Month(ListEndDate) Then
If Year(ListStartDay) = Year(ListEndDate) Then
Stoper = "Yes"
End If
End If
End If
Loop

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
SaveFileName = "MTD_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code, Region, Sub Region, Status, " & SupervisorLabel & ", Name, Phone Number, Last " & RepDataType & " Date, Last " & RepDataType & " Amount, Purse Limit, Transaction Type " & ExcelOut & ", Total"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>MTD Report</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%></b>
<br>Region: <b><%=WR%></b>
<br>Report Data: <b><%=RepDataType%></b>
<table style="table-layout: fixed;">
<thead>
<tr>
	<th>Agent Code</th>
	<th>Region</th>
	<th>Sub Region</th>
	<th>Status</th>
	<th><%=SupervisorLabel%></th>
	<th>Name</th>
	<th>Phone Number</th>

	<th>Last <%=RepDataType%> Date</th>
	<th>Last <%=RepDataType%> Amount</th>
	<th>Purse Limit</th>
	<th>Transaction Type</th>
<%=BrowserOut%>
	<th>Total</th>
</tr>
</thead>

<tbody>
<%
End If


AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True' "

AgentSQL = AgentSQL & " and SRID in (" & SRRegionList & ")"

AgentSQl = AgentSQl & " and TediActive = 'True' "

AgentSQl = AgentSQl & " Order By RegionName, TediEmpCode Asc"
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
LastTransDate = ""
LastTransAmount = "0"
BrowserOut1 = ""
ExcelOut1 = ""

If RepDataType = "Deductions" Then
LastTransQry = "Select Top(1) CDate, CAmount From viewTediTransactions where CType = 3 and TID = " & TID & " order by CDate Desc"
End If

If RepDataType = "Deposits" Then
LastTransQry = "Select Top(1) CDate, CAmount From viewTediTransactions where CType = 2 and TID = " & TID & " order by CDate Desc"
End If

If RepDataType = "Airtime" Then
LastTransQry = "Select Top(1) CDate, CAmount From viewTediTransactions where CType = 1 and TID = " & TID & " order by CDate Desc"
End If

set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRING
RecLasDeductions.Source = LastTransQry
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
If Not RecLasDeductions.EOF and Not RecLasDeductions.BOF Then
LastTransDateDay = Day(RecLasDeductions.Fields.Item("CDate").Value)
If Len(LastTransDateDay) = 1 Then
LastTransDateDay = "0" & LastTransDateDay
End If
LastTransDateMonth = Month(RecLasDeductions.Fields.Item("CDate").Value)
If Len(LastTransDateMonth) = 1 Then
LastTransDateMonth = "0" & LastTransDateMonth
End If

LastTransDate = LastTransDateDay & "/" &  LastTransDateMonth  & "/" & Year(RecLasDeductions.Fields.Item("CDate").Value)
LastTransAmount = RecLasDeductions.Fields.Item("CAmount").Value
End If

If RepDataType = "Airtime" Then
RecQuery = "Select * From ViewMTDByTedi where CType = 1 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If

If RepDataType = "Deposits" Then
RecQuery = "Select * From ViewMTDByTedi where CType = 2 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If

If RepDataType = "Deductions" Then
RecQuery = "Select * From ViewMTDByTedi where CType = 3 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If

set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRINGWrite
RecLasDeductions.Source = RecQuery
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0

LineOutTotal = 0
ListStartDay1 = StartDate

ListEndDate = DateAdd("d",1,EndDate)
Stoper = "No"
Do While Stoper = "No"
DayValue = 0


If Not RecLasDeductions.EOF and Not RecLasDeductions.BOF Then
DayValue = 0
If Day(RecLasDeductions.Fields.Item("CDate").Value) = Day(ListStartDay1) Then
If Month(RecLasDeductions.Fields.Item("CDate").Value) = Month(ListStartDay1) Then
If Year(RecLasDeductions.Fields.Item("CDate").Value) = Year(ListStartDay1) Then
DayValue = RecLasDeductions.Fields.Item("DayTotal").Value
RecLasDeductions.MoveNext
End If
End If
End If
End If


BrowserOut1 = BrowserOut1 & "<td>" & DayValue & "</td>"
ExcelOut1 = ExcelOut1 & ", " & DayValue
LineOutTotal = LineOutTotal + DayValue
ListStartDay1 = DateAdd("d",1,ListStartDay1)
If Day(ListStartDay1) = Day(ListEndDate) Then
If Month(ListStartDay1) = Month(ListEndDate) Then
If Year(ListStartDay1) = Year(ListEndDate) Then
Stoper = "Yes"
End If
End If
End If
Loop

BrowserOut1 = BrowserOut1 & "<td>" & LineOutTotal & "</td>"
ExcelOut1 = ExcelOut1 & ", " & LineOutTotal
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td>Active</td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediCell").Value)%></td>
	<td><%=LastTransDate%></td>
	<td><%=LastTransAmount%></td>
	<td><%=(RecAgentEdit.Fields.Item("Purselimit").Value)%></td>
	<td><%=RepDataType%></td>
	<%=BrowserOut1%>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & ", Active," & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("TediCell").Value & "," & LastTransDate & "," & LastTransAmount & ", " & RecAgentEdit.Fields.Item("Purselimit").Value & ", " & RepDataType & ExcelOut1)
End If
RecAgentEdit.MoveNext
Wend

DaylyTotalsBrowser = "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Totals:</td>"
DaylyTotalsExcel = ",,,,,,,,,,Totals"

ListStartDay2 = StartDate
BrowserOut = ""
ExcelOut = ""
ReportTotal = 0
ListEndDate = DateAdd("d",1,EndDate)
Stoper = "No"
Do While Stoper = "No"
DaylyTotal = 0

If RepDataType = "Deductions" Then
set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRING
RecLasDeductions.Source = "Select  * From ViewDeductionsDetails where SRID in (" & SRRegionList & ") and Day(DeductionDate) = " & Day(ListStartDay2) & " and Month(DeductionDate) = " & Month(ListStartDay2) & " and Year(DeductionDate) = " & Year(ListStartDay2)
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
While Not RecLasDeductions.EOF
DaylyTotal = DaylyTotal + RecLasDeductions.Fields.Item("DeductionValue").Value
RecLasDeductions.MoveNext
Wend
End If

If RepDataType = "Airtime" Then
set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRING
RecLasDeductions.Source = "Select  Sum(CAmount) As CAmountTotal From viewTediTransactions where CType = '1' and SRID in (" & SRRegionList & ") and Day(CDate) = " & Day(ListStartDay2) & " and Month(CDate) = " & Month(ListStartDay2) & " and Year(CDate) = " & Year(ListStartDay2)
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
If IsNull(RecLasDeductions.Fields.Item("CAmountTotal").Value) = false then
DaylyTotal = RecLasDeductions.Fields.Item("CAmountTotal").Value
End If
End If

If RepDataType = "Deposits" Then
set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRING
RecLasDeductions.Source = "Select Sum(CAmount) As CAmountTotal From viewTediTransactions where CType = '2' and SRID in (" & SRRegionList & ") and Day(CDate) = " & Day(ListStartDay2) & " and Month(CDate) = " & Month(ListStartDay2) & " and Year(CDate) = " & Year(ListStartDay2)
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
If IsNull(RecLasDeductions.Fields.Item("CAmountTotal").Value) = false then
DaylyTotal = RecLasDeductions.Fields.Item("CAmountTotal").Value
End If
End If

DaylyTotalsBrowser = DaylyTotalsBrowser & "<td>" & DaylyTotal & "</td>"
DaylyTotalsExcel = DaylyTotalsExcel & ", " & DaylyTotal
ReportTotal = ReportTotal + DaylyTotal
ListStartDay2 = DateAdd("d",1,ListStartDay2)
If Day(ListStartDay2) = Day(ListEndDate) Then
If Month(ListStartDay2) = Month(ListEndDate) Then
If Year(ListStartDay2) = Year(ListEndDate) Then
Stoper = "Yes"
End If
End If
End If
Loop


If OutFormat = "B" Then
DaylyTotalsBrowser = DaylyTotalsBrowser & "<td>" & ReportTotal & "</td></tr>"
Response.write(DaylyTotalsBrowser)
Else
TheFile.Writeline(DaylyTotalsExcel & ", " & ReportTotal)
End If
If OutFormat = "B" Then
%>
</tbody>
</table>
<%
Else
response.redirect("Reports/" & SaveFileName)
End If
%>