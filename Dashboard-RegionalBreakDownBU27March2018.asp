<!--#include file="Connections/Site.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<%
RegionMonth = Month(Now)
RegionYear = Year(Now)

If Request.QueryString("RegMonth") <> "" Then
RegionMonth = Request.QueryString("RegMonth")
End If

If Request.QueryString("RegYear") <> "" Then
RegionYear = Request.QueryString("RegYear")
End If
%>

<h3>Regional Breakdown <%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%></h3>

<form name="form1" method="get" action="Telephone2.asp">
       
             
<select name="menu2" onChange="MM_jumpMenu2('parent',this,0)" Class="text3_frm">
<Option Value="Dashboard.asp">Select Period</Option>
                <%
YearStart = 2015
MonthStart = 1
StopRegLoop = "No"

Do While StopRegLoop = "No"

%>
                <option value="Dashboard.asp?RegMonth=<%=MonthStart%>&RegYear=<%=YearStart%>"><%=MonthName(MonthStart)%>&nbsp;<%=YearStart%></option>
                <%
MonthStart = MonthStart + 1

If MonthStart = 13 Then
MonthStart = 1
YearStart = YearStart + 1
End If


If MonthStart = Month(Now) Then
If YearStart = Year(Now) Then
StopRegLoop = "Yes"
End If
End If



Loop
%>
                <option value="Dashboard.asp?RegMonth=<%=MonthStart%>&RegYear=<%=YearStart%>"  <%=IsSelected%>><%=MonthName(MonthStart)%>&nbsp;<%=YearStart%></option>
              </select> <font size=1>This selection does not affect the head count values.
            
          
        
        </form>

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
                                <th>Region</th>
                                <th>Current HC</th>
                                <th>HC Targ</th>
                                <th>Airtime Target</th>
                                <th>Airtime Banked</th>
                                <th>%</th>
                                <th>Vends</th>
                                <th>Connect Target</th>
                                <th>Connect<br>To Date</th>
                                <th>%</th>
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

While Not RecRegionalDash.EOF
RC = RC + 1
RID = RecRegionalDash.Fields.Item("RID").Value
CurrentHC = 0
TargetHC = 0
AirTimeTarget = 0
AirtimeBanked = 0
ConnectionsTarget = 0
ConnectionsToDate = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "EXECUTE SPCurrentHCTarget @RID = " & RID
'RecHCTarget.Source = "SELECT Sum(HeadCountTarget) AS HCT FROM SubRegions Where RID = " & RID & " and SubRegionActive = 'True'"
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If IsNull(RecHCTarget.Fields.Item("HCT").Value) = false then
TargetHC = RecHCTarget.Fields.Item("HCT").Value
End If


set RecCurrentHC = Server.CreateObject("ADODB.Recordset")
RecCurrentHC.ActiveConnection = MM_Site_STRING
RecCurrentHC.Source = "EXECUTE SPCurrentHC @RID = " & RID
'RecCurrentHC.Source = "SELECT Count(TID) AS TediRegionTotal FROM ViewTediDetail Where RID = " & RID & " and TediActive = 'True'"
RecCurrentHC.CursorType = 0
RecCurrentHC.CursorLocation = 2
RecCurrentHC.LockType = 3
RecCurrentHC.Open()
RecCurrentHC_numRows = 0
If IsNull(RecCurrentHC.Fields.Item("TediRegionTotal").Value) = false then
CurrentHC = RecCurrentHC.Fields.Item("TediRegionTotal").Value
End If

AirTimeTarget = AirtimeTediMonthlyTarget * CurrentHC

set RecThisMonthsSales = Server.CreateObject("ADODB.Recordset")
RecThisMonthsSales.ActiveConnection = MM_Site_STRING
RecThisMonthsSales.Source = "EXECUTE SPThisMonthsSales @RID = " & RID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsSales.Source = "SELECT Sum(CAmount) AS TotalBanked FROM viewTediTransactions Where RID = " & RID & " and TediActive = 'True' and CType = '2' and Month(CDate) = " & RegionMonth & " and Year(CDate) = " & RegionYear
RecThisMonthsSales.CursorType = 0
RecThisMonthsSales.CursorLocation = 2
RecThisMonthsSales.LockType = 3
RecThisMonthsSales.Open()
RecThisMonthsSales_numRows = 0
If IsNull(RecThisMonthsSales.Fields.Item("TotalBanked").Value) = false then
AirtimeBanked = RecThisMonthsSales.Fields.Item("TotalBanked").Value
End If

ConnectionsTarget = CurrentHC * MonthlyTediGrossConnectionsTarget
VendsAmount = 0

set RecThisMonthsVends = Server.CreateObject("ADODB.Recordset")
RecThisMonthsVends.ActiveConnection = MM_Site_STRING
RecThisMonthsVends.Source =  "EXECUTE SPThisMonthsVends @RID = " & RID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsVends.Source = "SELECT Sum(VendAmount) AS TotalVends FROM ViewVendingDetailsOnTIDShort Where RID = " & RID & "  and Month(VendDate) = " & RegionMonth & " and Year(VendDate) = " & RegionYear
'response.write(RecThisMonthsVends.Source)
RecThisMonthsVends.CursorType = 0
RecThisMonthsVends.CursorLocation = 2
RecThisMonthsVends.LockType = 3
RecThisMonthsVends.Open()
RecThisMonthsVends_numRows = 0
If IsNull(RecThisMonthsVends.Fields.Item("TotalVends").Value) = false then
VendsAmount = RecThisMonthsVends.Fields.Item("TotalVends").Value
End If

set RecThisMonthsConnections = Server.CreateObject("ADODB.Recordset")
RecThisMonthsConnections.ActiveConnection = MM_Site_STRING
RecThisMonthsConnections.Source =  "EXECUTE SPThisMonthsConnections @RID = " & RID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsConnections.Source = "SELECT Count(ActID) AS TotalConnect FROM ViewSimActivationDetails Where RID = " & RID & "  and Month(ActivationDate) = " & RegionMonth & " and Year(ActivationDate) = " & RegionYear
RecThisMonthsConnections.CursorType = 0
RecThisMonthsConnections.CursorLocation = 2
RecThisMonthsConnections.LockType = 3
RecThisMonthsConnections.Open()
RecThisMonthsConnections_numRows = 0
If IsNull(RecThisMonthsConnections.Fields.Item("TotalConnect").Value) = false then
ConnectionsToDate = RecThisMonthsConnections.Fields.Item("TotalConnect").Value
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
	<td><%=RC%>. <a href="ViewRegBreakdown.asp?RID=<%=RID%>&RegionMonth=<%=RegionMonth%>&RegionYear=<%=RegionYear%>"><%=RecRegionalDash.Fields.Item("RegionCode").Value%></a></td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td><%=Formatnumber(TargetHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td>R <%=FormatNumber(VendsAmount,0)%></td>
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
response.flush
RecRegionalDash.MoveNext
Wend

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
	<th>Totals</th>
	<th><%=Formatnumber(TotalCurrentHC,0)%></th>
	<th><%=Formatnumber(TotalTargetHC,0)%></th>
	<th>R <%=Formatnumber(TotalAirTimeTarget,0)%></th>
	<th>R <%=FormatNumber(TotalAirtimeBanked,0)%></th>
	<th><%=TotalAirtimePerc%>%</th>
	<th>R <%=FormatNumber(TotalVendsAmount,0)%></th>
	<th><%=FormatNumber(TotalConnectionsTarget,0)%></th>
	<th><%=FormatNumber(TotalConnectionsToDate,0)%></th>
	<th><%=TotalConnectPerc%>%</th>
</tr>
</thead>

</table>
<p><strong>HC</strong> = Head Count
<br><strong>Airtime Target</strong> = Current HC X R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> (Airtime Deposits per Agent per Month)
<br><strong>Connections Target</strong> = Current HC X Gross Connections Target (<%=MonthlyTediGrossConnectionsTarget%> per Agent per Month)
</p>
<hr>
