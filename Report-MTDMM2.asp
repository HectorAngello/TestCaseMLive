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
SaveFileName = "MTD_Report-MobileMoney-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code Airtime,Agent Code Mobile Money,ID Number,Region,Sub Region, Status," & SupervisorLabel & ", Name,Phone Number,Last " & RepDataType & " Date,Last " & RepDataType & "Amount,Purse Limit,Transaction Type " & ExcelOut & ",Total"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>MTD Report MCharge</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%></b>
<br>Region: <b><%=WR%></b>
<br>Report Data: <b><%=RepDataType%></b>
<table style="table-layout: fixed;">
<thead>
<tr>
	<th>Agent Code Airtime</th>
	<th>Agent Code Mobile Money</th>
	<th>ID Number</th>
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
AgentList = ""

AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True' and MobileMoneyTedi = 'True'"

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

AgentCodeMC = ""
AgentCodeMM = ""
If RecAgentEdit.Fields.Item("MobileMoneyTedi").Value = "True" Then
AgentCodeMM = RecAgentEdit.Fields.Item("TediEmpCode").Value
If Left(AgentCodeMM,1) = "P" Then
AgentCodeMM = "M" & AgentCodeMM
End If
End If

If RecAgentEdit.Fields.Item("MChargeTedi").Value = "True" Then
AgentCodeMC = RecAgentEdit.Fields.Item("TediEmpCode").Value
If Left(AgentCodeMC,1) = "M" Then
AgentCodeMCT = Len(AgentCodeMC)
AgentCodeMC = Right(AgentCodeMC, AgentCodeMCT - 1)
End If
End If

AgentList = AgentList & TID & ", "

If RepDataType = "Deductions" Then
LastTransQry = "EXECUTE SPLastTransActionMM @TID = " & TID & ", @ctype = 3"
End If

If RepDataType = "Deposits" Then
LastTransQry = "EXECUTE SPLastTransActionMM @TID = " & TID & ", @ctype = 2"
End If

If RepDataType = "Mobile Money" Then
LastTransQry = "EXECUTE SPLastTransActionMM @TID = " & TID & ", @ctype = 1"
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
LastTransAmount = Replace(RecLasDeductions.Fields.Item("CAmount").Value, ",", ".")
End If


If RepDataType = "Mobile Money" Then
RecQuery = "EXECUTE SPTediMTDReportAirtimeMM @TID = " & TID & ", @date1 = '" & StartDate & "', @date2 = '" & EndDate & "'"
'RecQuery = "Select * From ViewMTDByTedi where CType = 1 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If

If RepDataType = "Deposits" Then
RecQuery = "EXECUTE SPTediMTDReportBankingMM @TID = " & TID & ", @date1 = '" & StartDate & "', @date2 = '" & EndDate & "'"
'RecQuery = "Select * From ViewMTDByTedi where CType = 2 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If

If RepDataType = "Deductions" Then
RecQuery = "EXECUTE SPTediMTDReportDeductionsMM @TID = " & TID & ", @date1 = '" & StartDate & "', @date2 = '" & EndDate & "'"
'RecQuery = "Select * From ViewMTDByTedi where CType = 3 and TediID = " & TID & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') order by CDate asc"
End If
'response.write(RecQuery)

set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRINGWrite
RecLasDeductions.Source = RecQuery
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
LineOutTotal = 0
ExcelOut1 = ""
While Not RecLasDeductions.EOF

DayValue = 0
If RecLasDeductions.Fields.Item("DayTotal").Value <> "" Then
DayValue = RecLasDeductions.Fields.Item("DayTotal").Value
End If

BrowserOut1 = BrowserOut1 & "<td>" & DayValue & "</td>"
ExcelOut1 = ExcelOut1 & ", " & Replace(DayValue, ",", ".")
LineOutTotal = LineOutTotal + DayValue
			
RecLasDeductions.MoveNext
Wend


BrowserOut1 = BrowserOut1 & "<td>" & LineOutTotal & "</td>"
ExcelOut1 = ExcelOut1 & ", " & Replace(LineOutTotal, ",", ".")
If OutFormat = "B" Then
%>
<tr>
	<td><%=AgentCodeMC%></td>
	<td><%=AgentCodeMM%></td>
	<td><%=(RecAgentEdit.Fields.Item("IDNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RegionName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td>Active</td>
	<td><%=(RecAgentEdit.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("ASLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediCell").Value)%></td>
	<td><%=LastTransDate%></td>
	<td><%=LastTransAmount%></td>
	<td><%=(RecAgentEdit.Fields.Item("PurselimitMM").Value)%></td>
	<td><%=RepDataType%></td>
	<%=BrowserOut1%>
</tr>
<%
Response.flush
Else
TheFile.Writeline(AgentCodeMC & "," & AgentCodeMM & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & RecAgentEdit.Fields.Item("RegionName").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & ", Active," & RecAgentEdit.Fields.Item("ASFirstName").Value & " " & RecAgentEdit.Fields.Item("ASLastName").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & " " & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("TediCell").Value & "," & LastTransDate & "," & LastTransAmount & ", " & RecAgentEdit.Fields.Item("PurselimitMM").Value & ", " & RepDataType & ExcelOut1)
End If
RecAgentEdit.MoveNext
Wend

AgentListT = Len(AgentList)
AgentList = Left(AgentList, AgentListT - 2)

DaylyTotalsBrowser = "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Totals:</td>"
DaylyTotalsExcel = ",,,,,,,,,,,Totals"

ListStartDay2 = StartDate
BrowserOut = ""
ExcelOut = ""
ReportTotal = 0
ListEndDate = DateAdd("d",1,EndDate)
Stoper = "No"
Do While Stoper = "No"

DaylyTotal = 0


set RecLasDeductions = Server.CreateObject("ADODB.Recordset")
RecLasDeductions.ActiveConnection = MM_Site_STRING
If RepDataType = "Mobile Money" Then
RecLasDeductions.Source = "SELECT  Sum(dbo.ViewMTDByTediMM.DayTotal) As DayTotal FROM dbo.ViewMTDByTediMM RIGHT OUTER JOIN  dbo.CalenderDetail ON dbo.ViewMTDByTediMM.TransDay = dbo.CalenderDetail.day AND dbo.ViewMTDByTediMM.TransMonth = dbo.CalenderDetail.month AND dbo.ViewMTDByTediMM.TransYear = dbo.CalenderDetail.year AND dbo.ViewMTDByTediMM.TediID IN (" & AgentList & ") AND dbo.ViewMTDByTediMM.CType = 1 WHERE dbo.CalenderDetail.day = " & Day(ListStartDay2) & " and dbo.CalenderDetail.month = " & Month(ListStartDay2) & " and dbo.CalenderDetail.year = " & Year(ListStartDay2)
'RecLasDeductions.Source = "EXECUTE SPTediMTDReportAirtimeDay @TID = " & TID & ", @date1 = '" & ListStartDay2 & "'"
'RecLasDeductions.Source = "Select  Sum(CAmount) As CAmountTotal From viewTediTransactions where CType = '1' and SRID in (" & SRRegionList & ") and Day(CDate) = " & Day(ListStartDay2) & " and Month(CDate) = " & Month(ListStartDay2) & " and Year(CDate) = " & Year(ListStartDay2)
End If
If RepDataType = "Deposits" Then
RecLasDeductions.Source = "SELECT  Sum(dbo.ViewMTDByTediMM.DayTotal) As DayTotal FROM dbo.ViewMTDByTediMM RIGHT OUTER JOIN  dbo.CalenderDetail ON dbo.ViewMTDByTediMM.TransDay = dbo.CalenderDetail.day AND dbo.ViewMTDByTediMM.TransMonth = dbo.CalenderDetail.month AND dbo.ViewMTDByTediMM.TransYear = dbo.CalenderDetail.year AND dbo.ViewMTDByTediMM.TediID IN (" & AgentList & ") AND dbo.ViewMTDByTediMM.CType = 2 WHERE dbo.CalenderDetail.day = " & Day(ListStartDay2) & " and dbo.CalenderDetail.month = " & Month(ListStartDay2) & " and dbo.CalenderDetail.year = " & Year(ListStartDay2)
'RecLasDeductions.Source = "EXECUTE SPTediMTDReportBankingDay @TID = " & TID & ", @date1 = '" & ListStartDay2 & "'"
'RecLasDeductions.Source = "Select Sum(CAmount) As CAmountTotal From viewTediTransactions where CType = '2' and SRID in (" & SRRegionList & ") and Day(CDate) = " & Day(ListStartDay2) & " and Month(CDate) = " & Month(ListStartDay2) & " and Year(CDate) = " & Year(ListStartDay2)
End If
If RepDataType = "Deductions" Then
RecLasDeductions.Source = "SELECT  Sum(dbo.ViewMTDByTediMM.DayTotal) As DayTotal FROM dbo.ViewMTDByTediMM RIGHT OUTER JOIN  dbo.CalenderDetail ON dbo.ViewMTDByTediMM.TransDay = dbo.CalenderDetail.day AND dbo.ViewMTDByTediMM.TransMonth = dbo.CalenderDetail.month AND dbo.ViewMTDByTediMM.TransYear = dbo.CalenderDetail.year AND dbo.ViewMTDByTediMM.TediID IN (" & AgentList & ") AND dbo.ViewMTDByTediMM.CType = 3 WHERE dbo.CalenderDetail.day = " & Day(ListStartDay2) & " and dbo.CalenderDetail.month = " & Month(ListStartDay2) & " and dbo.CalenderDetail.year = " & Year(ListStartDay2)
'RecLasDeductions.Source = "EXECUTE SPTediMTDReportDeductionsDay @TID = " & TID & ", @date1 = '" & ListStartDay2 & "'"
'RecLasDeductions.Source = "Select  * From ViewDeductionsDetails where SRID in (" & SRRegionList & ") and Day(DeductionDate) = " & Day(ListStartDay2) & " and Month(DeductionDate) = " & Month(ListStartDay2) & " and Year(DeductionDate) = " & Year(ListStartDay2)
End If
'response.write(RecLasDeductions.Source)
RecLasDeductions.CursorType = 0
RecLasDeductions.CursorLocation = 2
RecLasDeductions.LockType = 3
RecLasDeductions.Open()
RecLasDeductions_numRows = 0
If IsNull(RecLasDeductions.Fields.Item("DayTotal").Value) = false then
DaylyTotal = DaylyTotal +  RecLasDeductions.Fields.Item("DayTotal").Value
'Response.write("<br>" & RecLasDeductions.Source)
End If




DaylyTotalsBrowser = DaylyTotalsBrowser & "<td>" & DaylyTotal & "</td>"
DaylyTotalsExcel = DaylyTotalsExcel & ", " & DaylyTotal
ReportTotal = ReportTotal + DaylyTotal

ListStartDay2 = DateAdd("d",1,ListStartDay2)
ListStartDay2 = Day(ListStartDay2) & " " & MonthName(Month(ListStartDay2)) & " " & Year(ListStartDay2)
If Day(ListStartDay2) = Day(ListEndDate) Then
If Month(ListStartDay2) = Month(ListEndDate) Then
If Year(ListStartDay2) = Year(ListEndDate) Then
Stoper = "Yes"
DaylyTotal = 0
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



