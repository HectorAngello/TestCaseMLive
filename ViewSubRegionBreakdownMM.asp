<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

RID = Request.QueryString("RID")

RegionMonth = Request.QueryString("Regionmonth")
RegionYear = Request.QueryString("RegionYear")

DashFileName = "DashboardSubRegions_MM_" & Session("UNID") & "_" & Day(Now) & MonthName(Month(Now),true) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now)
TheFilePath=(AppPath & "Exports\" & DashFileName & ".csv")
'response.write(TheFilePath)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
'************ beginning of the file body ***********
TheFile.Writeline("Period,Sub Region,Current Head Count,Target Head Count,Airtime Target,Airtime Banked,Airtime %,Vends Airtime,Vends Data,Deductions,Gross Connections Target,Gross Connections To Date,Connections %")



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
                <!-- #include file="Includes/sidebar.asp" --><br><a href="ViewRegBreakDownMM.asp?RID=<%=Request.QueryString("RID")%>&RegionYear=<%=Request.QueryString("RegionYear")%>&RegionMonth=<%=Request.QueryString("RegionMonth")%>" class="nice blue radius button">View <%=RecRegionalDash.Fields.Item("RegionCode").Value%> Mentors For The Same Period</a>
		<br><br><A href="Exports/<%=DashFileName%>.csv"  class="nice blue radius button">Export</a>
            </div>
            <div class="nine columns">
                <div class="content panel">

                        <div class="nine columns"><h2>Mobile Money Mentor Breakdown: <%=RecRegionalDash.Fields.Item("RegionCode").Value%> (<%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%>)</h2></div>
                        <div class="three columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<p>Viewing a breakdown of sub regions which have agents in <%=RecRegionalDash.Fields.Item("RegionCode").Value%>.</p>
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
        "caption": "Sub Region",
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
RecASs.Source = "SELECT * FROM SubRegions Where RID = " & RID & " and SubRegionActive = 'True' order by SubRegionName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
%>
        {
          "label": "<%=RecASs.Fields.Item("SubRegionName").Value%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
%>
        ,{
          "label": "<%=RecASs.Fields.Item("SubRegionName").Value%>"
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
RecASs.Source =  "SELECT * FROM SubRegions Where RID = " & RID & " and SubRegionActive = 'True' order by SubRegionName asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
SRID = RecASs.Fields.Item("SRID").Value
AirtimeBanked = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard2MM Where SRID = " & SRID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
AirtimeBanked = RecHCTarget.Fields.Item("Banked").Value
End If
%>
	{
        "seriesname": "Mobile Money Banked",
        "data": [
	{
          "value": "<%=AirtimeBanked%>"
        }
<%
RecASs.MoveNext
While Not RecASs.EOF
SRID = RecASs.Fields.Item("SRID").Value
AirtimeBanked = 0
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard2MM Where SRID = " & SRID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
AirtimeBanked = RecHCTarget.Fields.Item("Banked").Value
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
                                <th>Sub Region</th>
                                <th>Current HC</th>
                                <th>HC Targ</th>
                                <th>Airtime Target</th>
                                <th>Airtime Banked</th>
                                <th>%</th>
                                <th>Deductions</th>

                            </tr>
                        </thead>
<tbody>
<%
TotalCurrentHC = 0
TotalTargetHC = 0
TotalAirTimeTarget = 0
TotalAirtimeBanked = 0
TotalDeductionsAmount = 0


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
VendsAmount = 0
VendsAmountData = 0
Deductions = 0

SRID = RecSubRegions.Fields.Item("SRID").Value
set RecHCTarget = Server.CreateObject("ADODB.Recordset")
RecHCTarget.ActiveConnection = MM_Site_STRING
RecHCTarget.Source = "SELECT Top(1)* FROM  PrerenderSubRegionsDashboard2MM Where SRID = " & SRID & " and RepMonth = " & RegionMonth & " and RepYear = " & RegionYear
'response.write(RecHCTarget.Source)
RecHCTarget.CursorType = 0
RecHCTarget.CursorLocation = 2
RecHCTarget.LockType = 3
RecHCTarget.Open()
RecHCTarget_numRows = 0
If Not RecHCTarget.EOF and Not RecHCTarget.BOF Then
CurrentHC = RecHCTarget.Fields.Item("CurrentHC").Value
TargetHC = RecHCTarget.Fields.Item("HCTarget").Value
AirtimeBanked = RecHCTarget.Fields.Item("Banked").Value
Deductions = RecHCTarget.Fields.Item("Deductions").Value
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


AirtimePerc = 0

If AirtimeBanked > 0 then
If AirTimeTarget > 0 then
AirtimePerc = FormatNumber((AirtimeBanked / AirTimeTarget) * 100,0)
End if
End if

TheFile.Writeline(MonthName(RegionMonth) & " " & RegionYear & "," & RecSubRegions.Fields.Item("SubRegionName").Value & "," & Formatnumber(CurrentHC,,,,0) & "," & Formatnumber(TargetHC,,,,0) & "," & Formatnumber(AirTimeTarget,,,,0) & "," & FormatNumber(AirtimeBanked,,,,0) & "," & AirtimePerc & "," & FormatNumber(Deductions,,,,0))
%>


<tr>
	<td><a href="SubRegionBreakDownMM.asp?SRID=<%=RecSubRegions.Fields.Item("SRID").Value%>&Regionmonth=<%=Request.QueryString("Regionmonth")%>&RegionYear=<%=Request.QueryString("RegionYear")%>"><%=RC%>. <%=RecSubRegions.Fields.Item("SubRegionName").Value%></a></td>
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

Response.flush
RecSubRegions.MoveNext
Wend


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
	<th>R <%=FormatNumber(TotalDeductionsAmount,0)%></th>

</tr>
</thead>

</table>
<%
TheFile.close
Set FSO = nothing
%>
<p><strong>HC</strong> = Head Count
<br><strong>Mobile Money Target</strong> = Current HC X R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> X <%=ThisMonthDays%> Days This Month (Mobile Money Deposits per Agent per Month)
<br><strong>NB:</strong> Data is pre-generated, Data is updated every hour.
</p>
</div>
                    </div>

                             </div>        

<!-- #include file="includes/footer.asp" -->

