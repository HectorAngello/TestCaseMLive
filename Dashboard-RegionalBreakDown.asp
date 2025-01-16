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
                                <th style="font-size: 12px">Airtime Target</th>
                                <th style="font-size: 12px">Airtime Banked</th>
                                <th style="font-size: 12px">%</th>
                                <th style="font-size: 12px">Vends Airtime</th>
                                <th style="font-size: 12px">Vends Data</th>
                                <th style="font-size: 12px">Vends SMS</th>
                                <th style="font-size: 12px">Deductions</th>
                                <th style="font-size: 12px">Gross Con Target</th>
                                <th style="font-size: 12px">Gross Cons To Date</th>
                                <th style="font-size: 12px">%</th>
                            </tr>
                        </thead>
<tbody>
<%
RC = 0

TotalCurrentHC = 0
TotalTargetHC = 0
TotalAirTimeTarget = 0
TotalAirtimeBanked = 0
TotalConnectionsTarget = 0
TotalConnectionsToDate = 0
TotalVendsAmount = 0
TotalVendsAmountData = 0
TotalDeductionsAmount = 0
TotalVendsAmountSMS = 0

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
ConnectionsTarget = 0
ConnectionsToDate = 0
VendsAmount = 0
VendsAmountData = 0
VendsAmountSMS = 0
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
RecCheck.Source = "SELECT Sum(VendsSMS) as VendSMSTotal, Sum(CurrentHC) as HCTotal, Sum(HCTarget) as HCTartGetTotal, Sum(AirtimeBanked) As TotalATB, Sum(VendsAirtime) As TotalVended, Sum(VendsData) As TotalVendedData, Sum(ConnectionsToDate) As ConsToDate, Sum(Deductions) As DeductionsToDate FROM PrerenderSubRegionsDashboard where RID = " & RID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
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
If IsNull(RecCheck.Fields.Item("TotalVended").Value) = false then
VendsAmount = RecCheck.Fields.Item("TotalVended").Value
End If
If IsNull(RecCheck.Fields.Item("ConsToDate").Value) = false then
ConnectionsToDate = RecCheck.Fields.Item("ConsToDate").Value
End If
If IsNull(RecCheck.Fields.Item("DeductionsToDate").Value) = false then
Deductions = RecCheck.Fields.Item("DeductionsToDate").Value
End If
If IsNull(RecCheck.Fields.Item("TotalVendedData").Value) = false then
VendsAmountData = RecCheck.Fields.Item("TotalVendedData").Value
End If
If IsNull(RecCheck.Fields.Item("VendSMSTotal").Value) = false then
VendsAmountSMS = RecCheck.Fields.Item("VendSMSTotal").Value
End If

AirtimeTediMonthlyTarget = 0
set RecGetTargets = Server.CreateObject("ADODB.Recordset")
RecGetTargets.ActiveConnection = MM_Site_STRING
RecGetTargets.Source = "SELECT Top(1)* FROM MonthlyTargets where PeriodMonth = " & RegionMonth & " and PeriodYear = " & RegionYear
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
ConnectionsTarget = CurrentHC * RecGetTargets.Fields.Item("ConnectionsTarget").Value * ThisMonthDays

ConnectPerc = 0

If ConnectionsToDate > 0 then
If ConnectionsTarget > 0 then
ConnectPerc = FormatNumber((ConnectionsToDate / ConnectionsTarget) * 100,0)
End if
End if

AirtimePerc = 0

If AirtimeBanked > 0 then
If AirTimeTarget > 0 then
AirtimePerc = FormatNumber((AirtimeBanked / AirTimeTarget) * 100,0)
End if
End if

TheFile.Writeline(MonthName(RegionMonth) & " " & RegionYear & "," & RecRegionalDash.Fields.Item("RegionCode").Value & "," & Formatnumber(CurrentHC,,,,0) & "," & Formatnumber(TargetHC,,,,0) & "," & Formatnumber(AirTimeTarget,,,,0) & "," & FormatNumber(AirtimeBanked,,,,0) & "," & AirtimePerc & "," & FormatNumber(VendsAmount,,,,0) & "," & FormatNumber(VendsAmountData,,,,0) & "," & FormatNumber(VendsAmountSMS,,,,0) & "," & FormatNumber(Deductions,,,,0) & "," & FormatNumber(ConnectionsTarget,,,,0) & "," & FormatNumber(ConnectionsToDate,,,,0) & "," & ConnectPerc)
%>

<tr>
	<td><%=RC%>. <a href="ViewRegBreakdown.asp?RID=<%=RID%>&RegionMonth=<%=RegionMonth%>&RegionYear=<%=RegionYear%>"><%=RecRegionalDash.Fields.Item("RegionCode").Value%></a></td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td><%=Formatnumber(TargetHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td>R <%=FormatNumber(VendsAmount,0)%></td>
	<td>R <%=FormatNumber(VendsAmountData,0)%></td>
	<td>R <%=FormatNumber(VendsAmountSMS,0)%></td>
	<td>R <%=FormatNumber(Deductions,0)%></td>
	<td><%=FormatNumber(ConnectionsTarget,0)%></td>
	<td><%=FormatNumber(ConnectionsToDate,0)%></td>
	<td><%=ConnectPerc%>%</td>
</tr>

<%
TotalCurrentHC = TotalCurrentHC + CurrentHC
TotalTargetHC = TotalTargetHC + TargetHC
TotalAirTimeTarget = TotalAirTimeTarget + AirTimeTarget
TotalAirtimeBanked = TotalAirtimeBanked + AirtimeBanked
TotalConnectionsTarget = TotalConnectionsTarget + ConnectionsTarget
TotalConnectionsToDate = TotalConnectionsToDate + ConnectionsToDate
TotalVendsAmount = TotalVendsAmount + VendsAmount
TotalDeductionsAmount = TotalDeductionsAmount + Deductions
TotalVendsAmountData = TotalVendsAmountData + VendsAmountData
TotalVendsAmountSMS = TotalVendsAmountSMS + VendsAmountSMS
response.flush
RecRegionalDash.MoveNext
Wend

CurrentHC = 0
TargetHC = 0
AirTimeTarget = 0
AirtimeBanked = 0
ConnectionsTarget = 0
ConnectionsToDate = 0
VendsAmount = 0
VendsAmountData = 0
Deductions = 0
VendsAmountSMS = 0

set RecUnAlloVendAir = Server.CreateObject("ADODB.Recordset")
RecUnAlloVendAir.ActiveConnection = MM_Site_STRING
RecUnAlloVendAir.Source = "SELECT Sum(VendAmount) As UnAlloAir FROM Vends where TID = '0' and AmountType = 'Airtime' and Month(Venddate) = " & RegionMonth & " and Year(Venddate) = " & RegionYear
RecUnAlloVendAir.CursorType = 0
RecUnAlloVendAir.CursorLocation = 2
RecUnAlloVendAir.LockType = 3
RecUnAlloVendAir.Open()
RecUnAlloVendAir_numRows = 0
If IsNull(RecUnAlloVendAir.Fields.Item("UnAlloAir").Value) = "False" Then
VendsAmount = FormatNumber(RecUnAlloVendAir.Fields.Item("UnAlloAir").Value,0)
End If

set RecUnAlloVendData = Server.CreateObject("ADODB.Recordset")
RecUnAlloVendData.ActiveConnection = MM_Site_STRING
RecUnAlloVendData.Source = "SELECT Sum(VendAmount) As UnAlloAir FROM Vends where TID = '0' and AmountType = 'Data' and Month(Venddate) = " & RegionMonth & " and Year(Venddate) = " & RegionYear
RecUnAlloVendData.CursorType = 0
RecUnAlloVendData.CursorLocation = 2
RecUnAlloVendData.LockType = 3
RecUnAlloVendData.Open()
RecUnAlloVendData_numRows = 0
If IsNull(RecUnAlloVendData.Fields.Item("UnAlloAir").Value) = "False" Then
VendsAmountData = FormatNumber(RecUnAlloVendData.Fields.Item("UnAlloAir").Value,0)
End If

set RecUnAlloVendSMS = Server.CreateObject("ADODB.Recordset")
RecUnAlloVendSMS.ActiveConnection = MM_Site_STRING
RecUnAlloVendSMS.Source = "SELECT Sum(VendAmount) As UnAlloAir FROM Vends where TID = '0' and AmountType = 'SMS' and Month(Venddate) = " & RegionMonth & " and Year(Venddate) = " & RegionYear
RecUnAlloVendSMS.CursorType = 0
RecUnAlloVendSMS.CursorLocation = 2
RecUnAlloVendSMS.LockType = 3
RecUnAlloVendSMS.Open()
RecUnAlloVendSMS_numRows = 0
If IsNull(RecUnAlloVendSMS.Fields.Item("UnAlloAir").Value) = "False" Then
VendsAmountSMS = FormatNumber(RecUnAlloVendSMS.Fields.Item("UnAlloAir").Value,0)
End If

set RecUnAlloConnect = Server.CreateObject("ADODB.Recordset")
RecUnAlloConnect.ActiveConnection = MM_Site_STRING
RecUnAlloConnect.Source = "SELECT Count(ActID) As UnAlloCons FROM SimActivations where TID = '0' and  Month(ActivationDate) = " & RegionMonth & " and Year(ActivationDate) = " & RegionYear
RecUnAlloConnect.CursorType = 0
RecUnAlloConnect.CursorLocation = 2
RecUnAlloConnect.LockType = 3
RecUnAlloConnect.Open()
RecUnAlloConnect_numRows = 0
If IsNull(RecUnAlloConnect.Fields.Item("UnAlloCons").Value) = "False" Then
ConnectionsToDate = RecUnAlloConnect.Fields.Item("UnAlloCons").Value
End If

set RecUnAlloAirtime = Server.CreateObject("ADODB.Recordset")
RecUnAlloAirtime.ActiveConnection = MM_Site_STRING
RecUnAlloAirtime.Source = "SELECT Sum(TransAmount) As UnAlloAT FROM MChargeFNBTrans where Allocated = 'False' and  Month(FNBDate) = " & RegionMonth & " and Year(FNBDate) = " & RegionYear
RecUnAlloAirtime.CursorType = 0
RecUnAlloAirtime.CursorLocation = 2
RecUnAlloAirtime.LockType = 3
RecUnAlloAirtime.Open()
RecUnAlloAirtime_numRows = 0
If IsNull(RecUnAlloAirtime.Fields.Item("UnAlloAT").Value) = "False" Then
AirtimeBanked = RecUnAlloAirtime.Fields.Item("UnAlloAT").Value
End If

ConnectPerc = 0

If ConnectionsToDate > 0 then
If ConnectionsTarget > 0 then
ConnectPerc = FormatNumber((ConnectionsToDate / ConnectionsTarget) * 100,0)
End if
End if

AirtimePerc = 0

If AirtimeBanked > 0 then
If AirTimeTarget > 0 then
AirtimePerc = FormatNumber((AirtimeBanked / AirTimeTarget) * 100,0)
End if
End if
%>
<tr>
	<td>UnAllo</td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td><%=Formatnumber(TargetHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td><a href="ExportUnknownVendMSISDNs.asp">R <%=FormatNumber(VendsAmount,0)%></a></td>
	<td>R <%=FormatNumber(VendsAmountData,0)%></td>
	<td>R <%=FormatNumber(VendsAmountSMS,0)%></td>
	<td>R <%=FormatNumber(Deductions,0)%></td>
	<td><%=FormatNumber(ConnectionsTarget,0)%></td>
	<td><%=FormatNumber(ConnectionsToDate,0)%></td>
	<td><%=ConnectPerc%>%</td>
</tr>
<%




TotalConnectPerc = 0

If TotalConnectionsToDate > 0 then
If TotalConnectionsTarget > 0 then
TotalConnectPerc = FormatNumber((TotalConnectionsToDate / TotalConnectionsTarget) * 100,0)
End if
End if

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
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalVendsAmount,0)%></th>
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalVendsAmountData,0)%></th>
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalVendsAmountSMS,0)%></th>
	<th style="font-size: 12px" nowrap>R <%=FormatNumber(TotalDeductionsAmount,0)%></th>
	<th style="font-size: 12px"><%=FormatNumber(TotalConnectionsTarget,0)%></th>
	<th style="font-size: 12px"><%=FormatNumber(TotalConnectionsToDate,0)%></th>
	<th style="font-size: 12px"><%=TotalConnectPerc%>%</th>
</tr>
</thead>

</table>
<p><strong>HC</strong> = Head Count
<br><strong>Airtime Target</strong> = Current HC X R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> X <%=ThisMonthDays%> Days This Month (Airtime Deposits per Agent per Month)
<br><strong>Connections Target</strong> = Current HC X Connections Target (<%=RecGetTargets.Fields.Item("ConnectionsTarget").Value%> X <%=ThisMonthDays%> Days This Month per Agent per Month)
<br><strong>NB:</strong> Data is pre-generated, Data is updated every hour.
</p>
<hr>
