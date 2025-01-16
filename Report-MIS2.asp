<%
Region = Request.QueryString("Region")
ConStartDate = Request.QueryString("StartDate")
ConEndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
If OutFormat <> "P" Then
%>
<!-- #include file="includes/header.asp" -->
<%
Else
%><!--#include file="Connections/Site.asp" -->
<%
End If

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

DayCount = 0
SundayCount = 0
DateX = ConStartDate

DayCount = DateDiff("d",ConStartDate,ConEndDate) + 1
'Response.Write("DayCount:" & DayCount & "<br>")
'Response.Write(ConStartDate & "<br>")
'Response.Write(ConEndDate & "<br>")
TotDays = (DayCount * -1)
PubHolsCounter = 0
TryDate = ConEndDate
Do While TotDays < 0
'TryDate = DateDiff("d",TotDays,ConEndDate)
'Response.Write("TryDate: " & TryDate & "<br>")
' Check Public Holiday Starts
set RecIsPublicHoliday = Server.CreateObject("ADODB.Recordset")
RecIsPublicHoliday.ActiveConnection = MM_Site_STRING
RecIsPublicHoliday.Source = "SELECT * FROM PublicHolidays Where CompanyID = " & Session("CompanyID") & " and Day(HolidayDate) = " & Day(TryDate) & " and Month(HolidayDate) = " & Month(TryDate) & " and Year(HolidayDate) = " & Year(TryDate)
RecIsPublicHoliday.CursorType = 0
RecIsPublicHoliday.CursorLocation = 2
RecIsPublicHoliday.LockType = 3
RecIsPublicHoliday.Open()
RecIsPublicHoliday_numRows = 0
If Not RecIsPublicHoliday.EOF and Not RecIsPublicHoliday.BOF Then
PubHolsCounter = PubHolsCounter + 1
End If
' Check Public Holiday Ends
TotDays = TotDays + 1
'Response.Write(TryDate & " - " & WeekDayName(Weekday(TryDate)) & "<br>")
If WeekDayName(Weekday(TryDate)) = "Sunday" Then
SundayCount = SundayCount + 1
End If
TryDate = DateAdd("d",-1,TryDate)
Loop

If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "MIS_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Region, Active Agents, Total M-Charge Credit, Total M-Charge Deposits, Daily Ave Agent Sales"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>MIS Report</h3>
<p>Region: <b><%=WR%></b>
<br>Start Date: <b><%=FormatDateTime(ConStartDate,1)%></b>
<br>End Date: <b><%=FormatDateTime(ConEndDate,1)%></b>
<br>Days Used to Calculate = <b><%=DayCount%></b> Days - <b><%=SunDayCount%></b> Sundays = <b><%=DayCount - SunDayCount%></b> Days - <b><%=PubHolsCounter%></b> Public Holidays = <b><%=DayCount - SunDayCount - PubHolsCounter%></b> Days
</p>
<table>
<thead>
<tr>
	<th>Region</th>
	<th>Active Agents</th>
	<th>Total M-Charge Credit</th>
	<th>Total M-Charge Deposits</th>
	<th>Daily Ave Agent Sales</th>
</tr>
</thead>

<tbody>
<%
End If
set RecUsrRegCount = Server.CreateObject("ADODB.Recordset")
RecUsrRegCount.ActiveConnection = MM_Site_STRING
If Region = "0" Then
RecUsrRegCount.Source = "SELECT Distinct RegionName, RID FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
Else
RecUsrRegCount.Source = "SELECT Distinct RegionName, RID FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " and RID = " & Region
End If
RecUsrRegCount.CursorType = 0
RecUsrRegCount.CursorLocation = 2
RecUsrRegCount.LockType = 3
RecUsrRegCount.Open()
RecUsrRegCount_numRows = 0
While Not RecUsrRegCount.EOF
ActiveAgents = 0
RID = RecUsrRegCount.Fields.Item("RID").Value

set RecTotalAgents = Server.CreateObject("ADODB.Recordset")
RecTotalAgents.ActiveConnection = MM_Site_STRING
RecTotalAgents.Source = "SELECT Count(TID) AS AgentTotal FROM ViewTediDetail WHERE RID = '" & RID & "' and CompanyID = " & Session("CompanyID")
'Response.Write(RecFNB.Source)
RecTotalAgents.CursorType = 0
RecTotalAgents.CursorLocation = 2
RecTotalAgents.LockType = 3
RecTotalAgents.Open()
RecTotalAgents_numRows = 0
If IsNull(RecTotalAgents.Fields.Item("AgentTotal").Value) = false then
ActiveAgents = RecTotalAgents.Fields.Item("AgentTotal").Value
End If

TotalMchargeCredit = 0

set RecMChargeAllo = Server.CreateObject("ADODB.Recordset")
RecMChargeAllo.ActiveConnection = MM_Site_STRING
RecMChargeAllo.Source = "SELECT SUM(CAmount) AS TotalCredit FROM ViewTediTransactions WHERE CType = '1' and CompanyID = " & Session("CompanyID") & " and RID = '" & RID & "' and (CDate BETWEEN '" & ConStartDate & "' AND '" & ConEndDate & " 23:59:59')"
RecMChargeAllo.CursorType = 0
RecMChargeAllo.CursorLocation = 2
RecMChargeAllo.LockType = 3
RecMChargeAllo.Open()
RecMChargeAllo_numRows = 0
If IsNull(RecMChargeAllo.Fields.Item("TotalCredit").Value) = false then
TotalMchargeCredit = RecMChargeAllo.Fields.Item("TotalCredit").Value
End If

TotalMchargeDeposits = 0

set RecMChargeFNB = Server.CreateObject("ADODB.Recordset")
RecMChargeFNB.ActiveConnection = MM_Site_STRING
RecMChargeFNB.Source = "SELECT SUM(CAmount) AS TotalCredit FROM ViewTediTransactions WHERE CType = '2' and CompanyID = " & Session("CompanyID") & " and RID = '" & RID & "' and (CDate BETWEEN '" & ConStartDate & "' AND '" & ConEndDate & " 23:59:59')"
RecMChargeFNB.CursorType = 0
RecMChargeFNB.CursorLocation = 2
RecMChargeFNB.LockType = 3
RecMChargeFNB.Open()
RecMChargeFNB_numRows = 0
If IsNull(RecMChargeFNB.Fields.Item("TotalCredit").Value) = false then
TotalMchargeDeposits = RecMChargeFNB.Fields.Item("TotalCredit").Value
End If

DailyAgentAve = 0
If ActiveAgents > 0 Then
DX = DayCount - SunDayCount - PubHolsCounter
DailyAgentAve = (TotalMchargeDeposits / ActiveAgents) / DX
End If
If OutFormat = "B" Then
%><tr>
	<td><%=(RecUsrRegCount.Fields.Item("RegionName").Value)%></td>
	<td><%=ActiveAgents%></td>
	<td><%=FormatNumber(TotalMchargeCredit,,,,0)%></td>
	<td><%=FormatNumber(TotalMchargeDeposits,,,,0)%></td>
	<td><%=FormatNumber(DailyAgentAve,,,,0)%></td>
</tr><%
Response.flush
else
TheFile.Writeline(RecUsrRegCount.Fields.Item("RegionName").Value & "," & ActiveAgents & "," & Replace(TotalMchargeCredit, ",", ".") & "," & Replace(TotalMchargeDeposits, ",", ".") & "," & Replace(DailyAgentAve, ",", "."))
end if
RecUsrRegCount.Movenext
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