<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

SRID = Request.QueryString("SRID")

RegionMonth = Request.QueryString("Regionmonth")
RegionYear = Request.QueryString("RegionYear")

DashFileName = "DashboardSubRegionAgents_MM_" & Session("UNID") & "_" & Day(Now) & MonthName(Month(Now),true) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now)
TheFilePath=(AppPath & "Exports\" & DashFileName & ".csv")
'response.write(TheFilePath)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
'************ beginning of the file body ***********
TheFile.Writeline("Sub Region,Period,Agent,Empcode,Airtime Target,Mobile Money Banked,Mobile Money %,Deductions")


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


set RecRegionalDash = Server.CreateObject("ADODB.Recordset")
RecRegionalDash.ActiveConnection = MM_Site_STRING
RecRegionalDash.Source = "SELECT Top(1)* FROM SubRegions where SRID = " & SRID
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

                        <div class="nine columns"><h2>Mobile Money Sub Region Breakdown: <%=RecRegionalDash.Fields.Item("SubRegionName").Value%> (<%=RecRegionalDash.Fields.Item("SubRegionCode").Value%>&nbsp;|&nbsp;<%=MonthName(RegionMonth)%>&nbsp;<%=RegionYear%>)</h2></div>
                        <div class="three columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<p>&nbsp;</p>		</div>

</div>
<div class="row
		<div class="twelve columns">
  <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Agent</th>
                                <th>EmpCode</th>
                                <th>Mobile Money Target</th>
                                <th>Mobile Money Banked</th>
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
TotalDeductions = 0

set RecSubRegions = Server.CreateObject("ADODB.Recordset")
RecSubRegions.ActiveConnection = MM_Site_STRING
RecSubRegions.Source = "SELECT * FROM Tedis Where SRID = " & SRID & " and TediActive = 'True' and MobileMoneyTedi = 'True' order by TediFirstName asc"
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
Deductions = 0

set RecDeductions = Server.CreateObject("ADODB.Recordset")
RecDeductions.ActiveConnection = MM_Site_STRING
RecDeductions.Source = "EXECUTE SPTediMonthlyDeductionsMM @TID = " & RecSubRegions.Fields.Item("TID").Value & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecDeductions.Source = "SELECT SUM(CAmount) AS ATTotal FROM TediTransactions WHERE (MONTH(CDate) = " & RegionMonth & ") AND (YEAR(CDate) = " & RegionYear & ") AND (TediID = " & RecSubRegions.Fields.Item("TID").Value & ") AND (CType = 2)"
'response.write(RecAirtime.Source)
RecDeductions.CursorType = 0
RecDeductions.CursorLocation = 2
RecDeductions.LockType = 3
RecDeductions.Open()
RecDeductions_numRows = 0
If IsNULL(RecDeductions.Fields.Item("DeductionsTotal").Value) = "False" Then
Deductions = RecDeductions.Fields.Item("DeductionsTotal").Value
End If

set RecAirtime = Server.CreateObject("ADODB.Recordset")
RecAirtime.ActiveConnection = MM_Site_STRING
RecAirtime.Source = "EXECUTE SPTediMonthlyBankedMM @TID = " & RecSubRegions.Fields.Item("TID").Value & ", @RegMonth = " & RegionMonth & ", @RegYear = " & RegionYear
'RecAirtime.Source = "SELECT SUM(CAmount) AS ATTotal FROM TediTransactions WHERE (MONTH(CDate) = " & RegionMonth & ") AND (YEAR(CDate) = " & RegionYear & ") AND (TediID = " & RecSubRegions.Fields.Item("TID").Value & ") AND (CType = 2)"
'response.write(RecAirtime.Source)
RecAirtime.CursorType = 0
RecAirtime.CursorLocation = 2
RecAirtime.LockType = 3
RecAirtime.Open()
RecAirtime_numRows = 0
If IsNULL(RecAirtime.Fields.Item("ATTotal").Value) = "False" Then
AirtimeBanked = RecAirtime.Fields.Item("ATTotal").Value
End If


AirTimeTarget = AirtimeTediMonthlyTarget *  ThisMonthDays


AirtimePerc = 0

If AirtimeBanked > 0 then
If AirTimeTarget > 0 then
AirtimePerc = FormatNumber((AirtimeBanked / AirTimeTarget) * 100,0)
End if
End if
TheFile.Writeline(RecRegionalDash.Fields.Item("SubRegionName").Value & "," & MonthName(RegionMonth) & " " & RegionYear & "," & RecSubRegions.Fields.Item("TediFirstName").Value & " " & RecSubRegions.Fields.Item("TediLastName").Value & "," & RecSubRegions.Fields.Item("TediEmpCode").Value & "," & Formatnumber(AirTimeTarget,,,,0) & "," & FormatNumber(AirtimeBanked,,,,0) & "," & AirtimePerc & "," & FormatNumber(Deductions,,,,0))

%>


<tr>
	<td><a href="TediView.asp?TID=<%=RecSubRegions.Fields.Item("TID").Value%>"><%=RC%>. <%=RecSubRegions.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecSubRegions.Fields.Item("TediLastName").Value%></a></td>
	<td><%=RecSubRegions.Fields.Item("TediEmpCode").Value%></td>
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
TotalDeductions = TotalDeductions + Deductions
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
	<th>&nbsp;</th>
	<th>R <%=Formatnumber(TotalAirTimeTarget,0)%></th>
	<th>R <%=FormatNumber(TotalAirtimeBanked,0)%></th>
	<th><%=TotalAirtimePerc%>%</th>
	<th>R <%=FormatNumber(TotalDeductions,0)%></th>

</tr>
</thead>

</table>
<br><br><A href="Exports/<%=DashFileName%>.csv"  class="nice blue radius button">Export</a>
<%
TheFile.close
Set FSO = nothing
%>
<p><strong>Mobile Money Target</strong> = R <%=FormatNumber(AirtimeTediMonthlyTarget,,,,0)%> X <%=ThisMonthDays%> Days This Month (Mobile Money Deposits per Agent per Month)
</p>
</div>
                    </div>

                             </div>        

<!-- #include file="includes/footer.asp" -->

