<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

RID = Request.QueryString("RID")

RegionMonth = Request.QueryString("Regionmonth")
RegionYear = Request.QueryString("RegionYear")


set RecRegionalDash = Server.CreateObject("ADODB.Recordset")
RecRegionalDash.ActiveConnection = MM_Site_STRING
RecRegionalDash.Source = "SELECT Top(1)* FROM viewUserRegion where UserID = " & Session("UNID") & " and Active = 'Yes' and CompanyID = " & Session("CompanyID") & " and RID = " & RID
'response.write(RecRegionalDash.Source)
RecRegionalDash.CursorType = 0
RecRegionalDash.CursorLocation = 2
RecRegionalDash.LockType = 3
RecRegionalDash.Open()
RecRegionalDash_numRows = 0

DashFileName = "DashboardMentors_" & Session("UNID") & "_" & Day(Now) & MonthName(Month(Now),true) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now)
TheFilePath=(AppPath & "Exports\" & DashFileName & ".csv")
'response.write(TheFilePath)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
'************ beginning of the file body ***********
TheFile.Writeline("Period,Mentor,Current Head Count,Airtime Target,Airtime Banked,Airtime %,Vends Airtime,Vends Data,Deductions,Gross Connections Target,Gross Connections To Date,Connections %")

%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" --><br><a href="ViewSubRegionBreakDown.asp?RID=<%=Request.QueryString("RID")%>&RegionYear=<%=Request.QueryString("RegionYear")%>&RegionMonth=<%=Request.QueryString("RegionMonth")%>" class="nice blue radius button">View <%=RecRegionalDash.Fields.Item("RegionCode").Value%> Sub Regions For The Same Period</a>
		<br><br><A href="Exports/<%=DashFileName%>.csv"  class="nice blue radius button">Export</a>
            </div>
            <div class="nine columns">
                <div class="content panel">

                        <div class="nine columns"><h2>Mentor Breakdown: <%=RecRegionalDash.Fields.Item("RegionCode").Value%> (<%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%>)</h2></div>
                        <div class="three columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<p>Viewing a breakdown of mentors which have agents in <%=RecRegionalDash.Fields.Item("RegionCode").Value%>.</p>
<script type="text/javascript">
FusionCharts.ready(function() {
  var energyChart = new FusionCharts({
    type: 'mscolumn3d',
    renderAt: 'chart-container',
    width: '100%',
    height: '350',
    dataFormat: 'json',
    dataSource: {
      "chart": {
        "caption": "Mentors",
        "subCaption": "",
        "canvasbgalpha": "0",
        "legendbgalpha": "0",
	"legendPosition": "RIGHT",
        "numbersuffix": "",
	"rotatevalues": "1",
	"rotateLabels": "1",
        "exportEnabled": "1"
      },

      "categories": [{
        "category": [
<%
set RecASs = Server.CreateObject("ADODB.Recordset")
RecASs.ActiveConnection = MM_Site_STRING
RecASs.Source = "SELECT * FROM ASs Where RID = " & RID & " and ASActive = 'True' order by ASFirstName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
%>
        {
          "label": "<%=RecASs.Fields.Item("ASFirstName").Value & " " & RecASs.Fields.Item("ASLastName").Value%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
%>
        ,{
          "label": "<%=RecASs.Fields.Item("ASFirstName").Value & " " & RecASs.Fields.Item("ASLastName").Value%>"
        }
<%
RecASs.MoveNext
Wend
%>
]
      }],
      "dataset": [
<%
set RecASs = Server.CreateObject("ADODB.Recordset")
RecASs.ActiveConnection = MM_Site_STRING
RecASs.Source = "SELECT * FROM ASs Where RID = " & RID & " and ASActive = 'True' order by ASFirstName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
ASID = RecASs.Fields.Item("ASID").Value
AirtimeBanked = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
AirtimeBanked = RecHCTarget.Fields.Item("AirtimeBanked").Value
End If
%>
	{
        "seriesname": "Airtime Banked",
        "data": [
	{
          "value": "<%=AirtimeBanked%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
ASID = RecASs.Fields.Item("ASID").Value
AirtimeBanked = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
AirtimeBanked = RecHCTarget.Fields.Item("AirtimeBanked").Value
End If
%>
	, {
          "value": "<%=AirtimeBanked%>"
        }
<%
RecASs.MoveNext
Wend
%>
]
      }

, {
<%
set RecASs = Server.CreateObject("ADODB.Recordset")
RecASs.ActiveConnection = MM_Site_STRING
RecASs.Source = "SELECT * FROM ASs Where RID = " & RID & " and ASActive = 'True' order by ASFirstName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
ASID = RecASs.Fields.Item("ASID").Value
VendsAmount = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
VendsAmount = RecHCTarget.Fields.Item("Vends").Value
End If
%>
        "seriesname": "Vends",
        "data": [
	{
          "value": "<%=VendsAmount%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
ASID = RecASs.Fields.Item("ASID").Value
VendsAmount = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
VendsAmount = RecHCTarget.Fields.Item("Vends").Value
End If
%>
	, {
          "value": "<%=VendsAmount%>"
        }
<%
RecASs.MoveNext
Wend
%>
]
      }
, {
<%
set RecASs = Server.CreateObject("ADODB.Recordset")
RecASs.ActiveConnection = MM_Site_STRING
RecASs.Source = "SELECT * FROM ASs Where RID = " & RID & " and ASActive = 'True' order by ASFirstName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
ASID = RecASs.Fields.Item("ASID").Value
ConnectionsToDate = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
ConnectionsToDate = RecHCTarget.Fields.Item("ConnectionsToDate").Value
End If
%>
        "seriesname": "Gross Connections",
        "data": [
	{
          "value": "<%=ConnectionsToDate%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
ASID = RecASs.Fields.Item("ASID").Value
ConnectionsToDate = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
ConnectionsToDate = RecHCTarget.Fields.Item("ConnectionsToDate").Value
End If
%>
	, {
          "value": "<%=ConnectionsToDate%>"
        }
<%
RecASs.MoveNext
Wend
%>
]
      }
]
    }
  });

  energyChart.render();
});
</script>
<div id="chart-container">Chart Rendering</div>		</div>

</div>
<div class="row
		<div class="twelve columns">
  <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Mentor</th>
                                <th>Current HC</th>
                                <th>Airtime Target</th>
                                <th>Airtime Banked</th>
                                <th>%</th>
                                <th>Vends Airtime</th>
                                <th>Vends Data</th>
                                <th>Deductions</th>
                                <th>Gross Con Target</th>
                                <th>Gross Cons To Date</th>
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
TotalDeductionsAmount = 0
TotalVendsAmountData = 0

set RecSubRegions = Server.CreateObject("ADODB.Recordset")
RecSubRegions.ActiveConnection = MM_Site_STRING
RecSubRegions.Source = "SELECT * FROM ASs Where RID = " & RID & " and ASActive = 'True' order by ASFirstName asc"
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
VendsAmount = 0
Deductions = 0
VendsAmountData = 0
ASID = RecSubRegions.Fields.Item("ASID").Value
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard Where ASID = " & ASID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
CurrentHC = RecHCTarget.Fields.Item("CurrentHC").Value
TargetHC = RecHCTarget.Fields.Item("HCTarget").Value
AirtimeBanked = RecHCTarget.Fields.Item("AirtimeBanked").Value
ConnectionsToDate = RecHCTarget.Fields.Item("ConnectionsToDate").Value
VendsAmount = RecHCTarget.Fields.Item("VendsAirtime").Value
VendsAmountData = RecHCTarget.Fields.Item("VendsData").Value
Deductions = RecHCTarget.Fields.Item("Deductions").Value
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

TheFile.Writeline(MonthName(RegionMonth) & " " & RegionYear & "," & RecSubRegions.Fields.Item("ASFirstName").Value & " " & RecSubRegions.Fields.Item("ASLastName").Value & "," & Formatnumber(CurrentHC,,,,0) & "," & Formatnumber(AirTimeTarget,,,,0) & "," & FormatNumber(AirtimeBanked,,,,0) & "," & AirtimePerc & "," & FormatNumber(VendsAmount,,,,0) & "," & FormatNumber(VendsAmountData,,,,0) & "," & FormatNumber(Deductions,,,,0) & "," & FormatNumber(ConnectionsTarget,,,,0) & "," & FormatNumber(ConnectionsToDate,,,,0) & "," & ConnectPerc)
%>


<tr>
	<td><a href="MentorBreakDown.asp?ASID=<%=RecSubRegions.Fields.Item("ASID").Value%>&Regionmonth=<%=Request.QueryString("Regionmonth")%>&RegionYear=<%=Request.QueryString("RegionYear")%>"><%=RC%>. <%=RecSubRegions.Fields.Item("ASFirstName").Value%>&nbsp;<%=RecSubRegions.Fields.Item("ASLastName").Value%></a></td>
	<td><%=Formatnumber(CurrentHC,0)%></td>
	<td>R <%=Formatnumber(AirTimeTarget,0)%></td>
	<td>R <%=FormatNumber(AirtimeBanked,0)%></td>
	<td><%=AirtimePerc%>%</td>
	<td>R <%=FormatNumber(VendsAmount,0)%></td>
	<td>R <%=FormatNumber(VendsAmountData,0)%></td>
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
	<th>R <%=Formatnumber(TotalAirTimeTarget,0)%></th>
	<th>R <%=FormatNumber(TotalAirtimeBanked,0)%></th>
	<th><%=TotalAirtimePerc%>%</th>
	<th nowrap>R <%=FormatNumber(TotalVendsAmount,0)%></th>
	<th nowrap>R <%=FormatNumber(TotalVendsAmountData,0)%></th>
	<th>R <%=FormatNumber(TotalDeductionsAmount,0)%></th>
	<th><%=FormatNumber(TotalConnectionsTarget,0)%></th>
	<th><%=FormatNumber(TotalConnectionsToDate,0)%></th>
	<th><%=TotalConnectPerc%>%</th>
</tr>
</thead>

</table>
<p><strong>HC</strong> = Head Count
<br><strong>Airtime Target</strong> = Current HC X R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> X <%=ThisMonthDays%> Days This Month (Airtime Deposits per Agent per Month)
<br><strong>Connections Target</strong> = Current HC X Connections Target (<%=RecGetTargets.Fields.Item("ConnectionsTarget").Value%> X <%=ThisMonthDays%> Days This Month per Agent per Month)
<br><strong>NB:</strong> Data is pre-generated, Data is updated every hour.
</p>
<%
TheFile.close
Set FSO = nothing
%>
</div>
                    </div>

                             </div>        

<!-- #include file="includes/footer.asp" -->

