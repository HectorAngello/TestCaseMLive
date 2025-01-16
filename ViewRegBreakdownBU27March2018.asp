<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

RID = Request.QueryString("RID")

RegionMonth = Month(Now)
RegionYear = Year(Now)

If Request.QueryString("RegMonth") <> "" Then
RegionMonth = Request.QueryString("RegMonth")
End If

If Request.QueryString("RegYear") <> "" Then
RegionYear = Request.QueryString("RegYear")
End If

set RecRegionalDash = Server.CreateObject("ADODB.Recordset")
RecRegionalDash.ActiveConnection = MM_Site_STRING
RecRegionalDash.Source = "SELECT Top(1)* FROM viewUserRegion where UserID = " & Session("UNID") & " and Active = 'Yes' and CompanyID = " & Session("CompanyID") & " and RID = " & RID
'response.write(RecRegionalDash.Source)
RecRegionalDash.CursorType = 0
RecRegionalDash.CursorLocation = 2
RecRegionalDash.LockType = 3
RecRegionalDash.Open()
RecRegionalDash_numRows = 0
%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
            </div>
            <div class="nine columns">
                <div class="content panel">

                        <div class="nine columns"><h2>View Region Breakdown: <%=RecRegionalDash.Fields.Item("RegionCode").Value%> (<%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%>)</h2></div>
                        <div class="three columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<p>&nbsp;</p>		</div>

</div>
<div class="row
		<div class="twelve columns">
  <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Sub Region</th>
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
TotalCurrentHC = 0
TotalTargetHC = 0
TotalAirTimeTarget = 0
TotalAirtimeBanked = 0
TotalConnectionsTarget = 0
TotalConnectionsToDate = 0
TotalVendsAmount = 0

set RecSubRegions = Server.CreateObject("ADODB.Recordset")
RecSubRegions.ActiveConnection = MM_Site_STRING
RecSubRegions.Source = "SELECT * FROM SubRegions Where RID = " & RID & " and SubRegionActive = 'True' order by SubRegionName asc"
'response.write(RecSubRegions.Source)
RecSubRegions.CursorType = 0
RecSubRegions.CursorLocation = 2
RecSubRegions.LockType = 3
RecSubRegions.Open()
RecSubRegions_numRows = 0

RC = 0
While Not RecSubRegions.EOF
RC = RC + 1

CurrentHC = 0
TargetHC = 0
AirTimeTarget = 0
AirtimeBanked = 0
ConnectionsTarget = 0
ConnectionsToDate = 0



SRID = RecSubRegions.Fields.Item("SRID").Value
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "EXECUTE SPTargetHCSubRegion @SRID = " & SRID
'RecHCTarget.Source = "SELECT Sum(HeadCountTarget) AS HCT FROM SubRegions Where SRID = " & SRID & " and SubRegionActive = 'True'"
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
RecCurrentHC.Source = "EXECUTE SPCurrentHCSubRegion @SRID = " & SRID
'RecCurrentHC.Source = "SELECT Count(TID) AS TediRegionTotal FROM ViewTediDetail Where SRID = " & SRID & " and TediActive = 'True'"
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
RecThisMonthsSales.Source = "EXECUTE SPThisMonthsSalesSubRegion @SRID = " & SRID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsSales.Source = "SELECT Sum(CAmount) AS TotalBanked FROM viewTediTransactions Where SRID = " & SRID & " and TediActive = 'True' and CType = '2' and Month(CDate) = " & RegionMonth & " and Year(CDate) = " & RegionYear
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
RecThisMonthsVends.Source =  "EXECUTE SPThisMonthsVendsSubRegion @SRID = " & SRID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsVends.Source = "SELECT Sum(VendAmount) AS TotalVends FROM ViewVendingDetailsOnTIDShort Where SRID = " & SRID & "  and Month(VendDate) = " & RegionMonth & " and Year(VendDate) = " & RegionYear
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
RecThisMonthsConnections.Source = "EXECUTE SPThisMonthsConnectionsSubRegion @SRID = " & SRID & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecThisMonthsConnections.Source = "SELECT Count(ActID) AS TotalConnect FROM ViewSimActivationDetails Where SRID = " & SRID & "  and Month(ActivationDate) = " & RegionMonth & " and Year(ActivationDate) = " & RegionYear
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
	<td><%=RC%>. <%=RecSubRegions.Fields.Item("SubRegionName").Value%></a></td>
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
Response.flush
RecSubRegions.MoveNext
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
%>
</tbody>
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
</div>
                    </div>

                             </div>        

<!-- #include file="includes/footer.asp" -->

