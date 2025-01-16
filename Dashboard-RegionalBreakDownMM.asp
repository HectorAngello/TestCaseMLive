<h3>Regional Breakdown <%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%></h3>
<%

set RecRegionalDash = Server.CreateObject("ADODB.Recordset")
RecRegionalDash.ActiveConnection = MM_Site_STRING
RecRegionalDash.Source = "SELECT Distinct RegionName, RegionCode, RID FROM viewUserRegion where UserID = " & Session("UNID") & " and Active = 'Yes' and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
'response.write(RecRegionalDash.Source)
RecRegionalDash.CursorType = 0
RecRegionalDash.CursorLocation = 2
RecRegionalDash.LockType = 3
RecRegionalDash.Open()
RecRegionalDash_numRows = 0

%>



                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th style="font-size: 12px">Region</th>
                                <th style="font-size: 12px">Cur HC</th>
                                <th style="font-size: 12px">Tar HC</th>
                                <th style="font-size: 12px">Mobile Money Target</th>
                                <th style="font-size: 12px">Mobile Money Banked</th>
                                <th style="font-size: 12px">%</th>
                                <th style="font-size: 12px">Deductions</th>

                            </tr>
                        </thead>
<tbody>
<%
RC = 0

TotalCurrentHC = 0
TotalTargetHC = 0
TotalAirTimeTarget = 0
TotalAirtimeBanked = 0
TotalDeductionsAmount = 0


IsCurrentMonth = "No"
If RegionMonth = Month(Now) Then
If RegionYear = Year(Now) Then
IsCurrentMonth = "Yes"
End If
End If

While Not RecRegionalDash.EOF
RC = RC + 1
RID = RecRegionalDash.Fields.Item("RID").Value
CurrentHC = 0
TargetHC = 0
AirTimeTarget = 0
AirtimeBanked = 0
Deductions = 0

set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT Sum(HeadCountTarget) As HCT FROM SubRegions Where RID = " & RID & " and SubRegionActive = 'True'"
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
If IsNull(RecSubRegion.Fields.Item("HCT").Value) = "False" Then
TargetHC = RecSubRegion.Fields.Item("HCT").Value
End If
' Use pregenerated Data

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "SELECT Sum(CurrentHC) as HCTotal, Sum(HCTarget) as HCTartGetTotal, Sum(Banked) As TotalATB, Sum(Deductions) As DeductionsToDate FROM PrerenderSubRegionsDashboardMM where RID = " & RID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If IsNull(RecCheck.Fields.Item("HCTotal").Value) = false then
CurrentHC = RecCheck.Fields.Item("HCTotal").Value
End If

If IsNull(RecCheck.Fields.Item("TotalATB").Value) = false then
AirtimeBanked = RecCheck.Fields.Item("TotalATB").Value
End If

If IsNull(RecCheck.Fields.Item("DeductionsToDate").Value) = false then
Deductions = RecCheck.Fields.Item("DeductionsToDate").Value
End If


AirtimeTediMonthlyTarget = 0
set RecGetTargets = Server.CreateObject("ADODB.Recordset")
RecGetTargets.ActiveConnection = MM_Site_STRING
RecGetTargets.Source = "SELECT Top(1)* FROM MonthlyTargetsMM where PeriodMonth = " & RegionMonth & " and PeriodYear = " & RegionYear
'response.write(RecGetTargets.Source)
RecGetTargets.CursorType = 0
RecGetTargets.CursorLocation = 2
RecGetTargets.LockType = 3
RecGetTargets.Open()
RecGetTargets_numRows = 0
AirtimeTediMonthlyTarget = RecGetTargets.Fields.Item("AirtimeTarget").Value

NextMonthDate = DateAdd("m", 1, Date())
NextMonthDate = "1 " & MonthName(Month(NextMonthDate)) & " " & Year(NextMonthDate)
LastDayThisMonthDate = DateAdd("d", -1, NextMonthDate)
ThisMonthDays = Day(LastDayThisMonthDate)

AirTimeTarget = AirtimeTediMonthlyTarget * CurrentHC * ThisMonthDays

ConnectPerc = 0

AirtimePerc = 0

If AirtimeBanked > 0 then
If AirTimeTarget > 0 then
AirtimePerc = FormatNumber((AirtimeBanked / AirTimeTarget) * 100,0)
End if
End if

TheFile.Writeline(MonthName(RegionMonth) & " " & RegionYear & "," & RecRegionalDash.Fields.Item("RegionCode").Value & "," & Formatnumber(CurrentHC,,,,0) & "," & Formatnumber(TargetHC,,,,0) & "," & Formatnumber(AirTimeTarget,,,,0) & "," & FormatNumber(AirtimeBanked,,,,0) & "," & AirtimePerc & "," & FormatNumber(Deductions,,,,0))
%>

<tr>
	<td><%=RC%>. <a href="ViewRegBreakdownMM.asp?RID=<%=RID%>&RegionMonth=<%=RegionMonth%>&RegionYear=<%=RegionYear%>"><%=RecRegionalDash.Fields.Item("RegionCode").Value%></a></td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td><%=Formatnumber(TargetHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td>R <%=FormatNumber(Deductions,0)%></td>
</tr>

<%
TotalCurrentHC = TotalCurrentHC + CurrentHC
TotalTargetHC = TotalTargetHC + TargetHC
TotalAirTimeTarget = TotalAirTimeTarget + AirTimeTarget
TotalAirtimeBanked = TotalAirtimeBanked + AirtimeBanked
TotalDeductionsAmount = TotalDeductionsAmount + Deductions
response.flush
RecRegionalDash.MoveNext
Wend

CurrentHC = 0
TargetHC = 0
AirTimeTarget = 0
AirtimeBanked = 0
Deductions = 0


set RecUnAlloAirtime = Server.CreateObject("ADODB.Recordset")
RecUnAlloAirtime.ActiveConnection = MM_Site_STRING
RecUnAlloAirtime.Source = "SELECT Sum(TransAmount) As UnAlloAT FROM MChargeFNBTransMM where Allocated = 'False' and  Month(FNBDate) = " & RegionMonth & " and Year(FNBDate) = " & RegionYear
RecUnAlloAirtime.CursorType = 0
RecUnAlloAirtime.CursorLocation = 2
RecUnAlloAirtime.LockType = 3
RecUnAlloAirtime.Open()
RecUnAlloAirtime_numRows = 0
If IsNull(RecUnAlloAirtime.Fields.Item("UnAlloAT").Value) = "False" Then
AirtimeBanked = RecUnAlloAirtime.Fields.Item("UnAlloAT").Value
End If

ConnectPerc = 0


AirtimePerc = 0

%>
<tr>
	<td>UnAllo</td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td><%=Formatnumber(TargetHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td>R <%=FormatNumber(Deductions,0)%></td>
</tr>
<%






TotalAirtimePerc = 0

If TotalAirtimeBanked > 0 then
If TotalAirTimeTarget > 0 then
TotalAirtimePerc = FormatNumber((TotalAirtimeBanked / TotalAirTimeTarget) * 100,0)
End if
End if

%></tbody>
<thead>
<tr>
	<th style="font-size: 12px">Totals</th>
	<th style="font-size: 12px"><%=Formatnumber(TotalCurrentHC,0)%></th>
	<th style="font-size: 12px"><%=Formatnumber(TotalTargetHC,0)%></th>
	<th style="font-size: 12px" nowrap>R <%=Formatnumber(TotalAirTimeTarget,0)%></th>
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalAirtimeBanked,0)%></th>
	<th style="font-size: 12px"><%=TotalAirtimePerc%>%</th>
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalDeductionsAmount,0)%></th>
</tr>
</thead>

</table>
<p><strong>HC</strong> = Head Count
<br><strong>Mobile Money Target</strong> = Current HC X R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> X <%=ThisMonthDays%> Days This Month (Mobile Money Deposits per Agent per Month)
<br><strong>NB:</strong> Data is pre-generated, Data is updated every hour.
</p>
<hr>
