<!-- #include file="includes/header.asp" -->
<%
Region = Request.Form("Region")
RepType = Request.Form("Type")
OrderBy = Request.Form("OrderBy")
Outformat = Request.Form("OutFormat")

'Response.write("Region: " & Region )

If Region = "0" then
WR = "All My Regions"
Else
set RecWR = Server.CreateObject("ADODB.Recordset")
RecWR.ActiveConnection = MM_Site_STRING
RecWR.Source = "SELECT * FROM Regions Where RID = " & Region
'Response.write(RecWR.Source)
RecWR.CursorType = 0
RecWR.CursorLocation = 2
RecWR.LockType = 3
RecWR.Open()
RecWR_numRows = 0
WR = RecWR.Fields.Item("RegionName").Value
End If

SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID")

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

'response.write(SRRegionList)
ShowLabel = "All Agents"
If RepType = "1" Then
ShowLabel = "Terminated Agents"
End If
If RepType = "2" Then
ShowLabel = "Active Agents"
End If

SQLqry = "Select * From ViewTediDetail where CompanyID =  " & Session("CompanyID")
If RepType = "1" Then
SQLqry = SQLqry & " and TediActive = 'False' "
End If
If RepType = "2" Then
SQLqry = SQLqry & " and TediActive = 'True' "
End If


SQLqry = SQLqry & " and SRID in (" & SRRegionList & ")"


If OrderBy = "1" Then
SQLqry = SQLqry & " Order By TediFirstName Asc"
End If

If OrderBy = "2" Then
SQLqry = SQLqry & " Order By TediLastName Asc"
End If

If OrderBy = "3" Then
SQLqry = SQLqry & " Order By TediEmpCode Asc"
End If
'Response.write(SQLqry)
set RecReport = Server.CreateObject("ADODB.Recordset")
RecReport.ActiveConnection = MM_Site_STRING
RecReport.Source = SQLqry
RecReport.CursorType = 0
RecReport.CursorLocation = 2
RecReport.LockType = 3
RecReport.Open()
RecReport_numRows = 0

DDListBrowserHeading = "<th>Consent Form Submitted</th><th>Crim Check</th><th>Crim Record</th><th>AML Trained</th><th>AML Passed</th><th>Phone Allocated</th><th>MSISDN Allocated</th><th>TDR Onboarded</th><th>Validated</th><th>Status</th>"
DDListExcelHeading = "Consent Form Submitted,Crim Check,Crim Record,AML Trained,AML Passed,Phone Allocated,MSISDN Allocated,TDR Onboarded,Validated,Status,"


If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "Master_File_Export-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = ""
if Request.Form("AgentRegion") = "Yes" Then TableHead = TableHead & "Region," End If
if Request.Form("AgentEmpCode") = "Yes" Then TableHead = TableHead & "Airtime Agent Code,Mobile Money Agent Code," End If
if Request.Form("ShowAgentStatus") = "Yes" Then TableHead = TableHead & "Agent Status," End If
if Request.Form("BDOName") = "Yes" Then TableHead = TableHead & SupervisorLabel & "," End If
if Request.Form("AgentName") = "Yes" Then TableHead = TableHead & "Agent Name," End If
if Request.Form("StartDate") = "Yes" Then TableHead = TableHead & "Start Date," End If
if Request.Form("TermDate") = "Yes" Then TableHead = TableHead & "Termination Date," End If
if Request.Form("AgentCell") = "Yes" Then TableHead = TableHead & "Primary Mobile," End If
if Request.Form("IDNo") = "Yes" Then TableHead = TableHead & "ID No," End If
if Request.Form("WPExpiryDate") = "Yes" Then TableHead = TableHead & "Work Permit Expiry Date," End If
if Request.Form("LastBankedDate") = "Yes" Then TableHead = TableHead & "Last Banked Date," End If
if Request.Form("ShowLastRefund") = "Yes" Then TableHead = TableHead & "Last Banked Amount," End If
if Request.Form("LastAirtimeDate") = "Yes" Then TableHead = TableHead & "Last Airtime Date," End If
if Request.Form("ShowLastAirtimeAmount") = "Yes" Then TableHead = TableHead & "Last Airtime Amount," End If
if Request.Form("MChargeBal") = "Yes" Then TableHead = TableHead & "M-Charge Balance," End If
if Request.Form("PurseLimit") = "Yes" Then TableHead = TableHead & "Purse Limit," End If
if Request.Form("Bank") = "Yes" Then TableHead = TableHead & "Bank," End If
if Request.Form("AccNo") = "Yes" Then TableHead = TableHead & "Account No," End If
if Request.Form("Branch") = "Yes" Then TableHead = TableHead & "Branch," End If
if Request.Form("AccType") = "Yes" Then TableHead = TableHead & "Account Type," End If
if Request.Form("ShowMobiSitePass") = "Yes" Then TableHead = TableHead & "Mobi Site Password," End If
if Request.Form("AgentCell2") = "Yes" Then TableHead = TableHead & "Secondary Mobile," End If
if Request.Form("TertiaryMobileNumber") = "Yes" Then TableHead = TableHead & "Tertiary Mobile Number," End If
if Request.Form("ExcludeFromMchargeBulkFile") = "Yes" Then TableHead = TableHead & "Exclude From Bulk File," End If
if Request.Form("ShowRegionAlManager") = "Yes" Then TableHead = TableHead & "Regional Manager," End If
if Request.Form("AgentSubRegion") = "Yes" Then TableHead = TableHead & "Sub Region," End If
if Request.Form("Gender") = "Yes" Then TableHead = TableHead & "Gender," End If
if Request.Form("Race") = "Yes" Then TableHead = TableHead & "Race," End If
if Request.Form("TaxOffice") = "Yes" Then TableHead = TableHead & "BEE Signature Date," End If
if Request.Form("ResidentialAddress") = "Yes" Then TableHead = TableHead & "Residential Address," End If
if Request.Form("TermReason") = "Yes" Then TableHead = TableHead & "Termination Reason," End If
if Request.Form("BDOEmpCode") = "Yes" Then TableHead = TableHead & SupervisorLabel & " Code," End If
if Request.Form("SARS") = "Yes" Then TableHead = TableHead & "Cluster," End If
if Request.Form("SubRegionCode") = "Yes" Then TableHead = TableHead & "Sub Region Code," End If
if Request.Form("RealTimeCommOptIn") = "Yes" Then TableHead = TableHead & "Realtime Commission Opt In," End If
if Request.Form("ShowIfMchargeAgent") = "Yes" Then TableHead = TableHead & "Is A MCharge Agent," End If

if Request.Form("ShowIfMobileMoneyAgent") = "Yes" Then TableHead = TableHead & "Is A Mobile Money Agent," End If
if Request.Form("DDList") = "Yes" Then 
TableHead = TableHead & DDListExcelHeading
End If
if Request.Form("MMPurseLimit") = "Yes" Then TableHead = TableHead & "Mobile Money Purse Limit," End If
if Request.Form("MobileMoneyBalance") = "Yes" Then TableHead = TableHead & "Mobile Money Balance," End If
if Request.Form("MMLastBankedDate") = "Yes" Then TableHead = TableHead & "Mobile Money Last Banked Date," End If
if Request.Form("WPED") = "Yes" Then TableHead = TableHead & "Work Permit Expiry Date," End If
if Request.Form("MoMoAccNo") = "Yes" Then TableHead = TableHead & "Mobile Money Account Number," End If

if Request.Form("SkhokhoGSM") = "Yes" Then TableHead = TableHead & "Made for Skhokho GSM," End If
if Request.Form("SkhokhoDedicated") = "Yes" Then TableHead = TableHead & "Made for Skhokho Dedicated," End If
if Request.Form("TradingSpot") = "Yes" Then TableHead = TableHead & "Trading Spot," End If

TableHeadT = Len(TableHead)
TableHead = Left(TableHead, TableHeadT - 1)
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If



If OutFormat = "B" Then
%>
Master Agent File Export:
<br>Region: <%=WR%>
<br>Status: <%=ShowLabel%>

<table>
<thead>
<tr>
<%if Request.Form("AgentRegion") = "Yes" Then%><th>Region</th><%End If%>
<%if Request.Form("AgentEmpCode") = "Yes" Then%><th>Airtime Agent Code</th><th>Mobile Money Agent Code</th><%End If%>
<%if Request.Form("ShowAgentStatus") = "Yes" Then%><th>Agent Status</th><%End If%>
<%if Request.Form("BDOName") = "Yes" Then%><th><%=SupervisorLabel%></th><%End If%>
<%if Request.Form("AgentName") = "Yes" Then%><th>Agent Name</th><%End If%>
<%if Request.Form("StartDate") = "Yes" Then%><th>Start Date</th><%End If%>
<%if Request.Form("TermDate") = "Yes" Then%><th>Termination Date</th><%End If%>
<%if Request.Form("AgentCell") = "Yes" Then%><th>Primary Mobile</th><%End If%>
<%if Request.Form("IDNo") = "Yes" Then%><th>ID No</th><%End If%>
<%if Request.Form("WPExpiryDate") = "Yes" Then%><th>Work Permit Expiry Date</th><%End If%>
<%if Request.Form("LastBankedDate") = "Yes" Then%><th>Last Banked Date</th><%End If%>
<%if Request.Form("ShowLastRefund") = "Yes" Then%><th>Last Banked Amount</th><%End If%>
<%if Request.Form("LastAirtimeDate") = "Yes" Then%><th>Last Airtime Date</th><%End If%>
<%if Request.Form("ShowLastAirtimeAmount") = "Yes" Then%><th>Last Airtime Amount</th><%End If%>
<%if Request.Form("MChargeBal") = "Yes" Then%><th>M-Charge Balance</th><%End If%>
<%if Request.Form("PurseLimit") = "Yes" Then%><th>Purse Limit</th><%End If%>
<%if Request.Form("Bank") = "Yes" Then%><th>Bank</th><%End If%>
<%if Request.Form("AccNo") = "Yes" Then%><th>Account No</th><%End If%>
<%if Request.Form("Branch") = "Yes" Then%><th>Branch</th><%End If%>
<%if Request.Form("AccType") = "Yes" Then%><th>Account Type</th><%End If%>
<%if Request.Form("ShowMobiSitePass") = "Yes" Then%><th>Mobi Site Password</th><%End If%>
<%if Request.Form("AgentCell2") = "Yes" Then%><th>Secondary Mobile</th><%End If%>
<%if Request.Form("TertiaryMobileNumber") = "Yes" Then%><th>Tertiary Mobile Number</th><%End If%>
<%if Request.Form("ExcludeFromMchargeBulkFile") = "Yes" Then%><th>Exclude From Bulk File</th><%End If%>
<%if Request.Form("ShowRegionAlManager") = "Yes" Then%><th>Regional Manager</th><%End If%>
<%if Request.Form("AgentSubRegion") = "Yes" Then%><th>Sub Region</th><%End If%>
<%if Request.Form("Gender") = "Yes" Then%><th>Gender</th><%End If%>
<%if Request.Form("Race") = "Yes" Then%><th>Race</th><%End If%>
<%if Request.Form("TaxOffice") = "Yes" Then%><th>BEE Signature Date</th><%End If%>
<%if Request.Form("ResidentialAddress") = "Yes" Then%><th>Residential Address</th><%End If%>
<%if Request.Form("TermReason") = "Yes" Then%><th>Termination Reason</th><%End If%>
<%if Request.Form("BDOEmpCode") = "Yes" Then%><th><%=SupervisorLabel%> Code</th><%End If%>
<%if Request.Form("SARS") = "Yes" Then%><th>Cluster</th><%End If%>
<%if Request.Form("SubRegionCode") = "Yes" Then%><th>Sub Region Code</th><%End If%>
<%if Request.Form("RealTimeCommOptIn") = "Yes" Then%><th>Realtime Commission Opt In</th><%End If%>
<%if Request.Form("ShowIfMchargeAgent") = "Yes" Then%><th>Is A MCharge Agent</th><%End If%>
<%if Request.Form("ShowIfMobileMoneyAgent") = "Yes" Then%><th>Is A Mobile Money Agent</th><%End If%>
<%if Request.Form("DDList") = "Yes" Then%><%=DDListBrowserHeading%><%End If%>
<%if Request.Form("MMPurseLimit") = "Yes" Then%><th>Mobile Money Purse Limit</th><%End If%>
<%if Request.Form("MobileMoneyBalance") = "Yes" Then%><th>Mobile Money Balance</th><%End If%>
<%if Request.Form("MMLastBankedDate") = "Yes" Then%><th>Mobile Money Last Banked Date</th><%End If%>
<%if Request.Form("WPED") = "Yes" Then%><th>Work Permit Expiry Date</th><%End If%>
<%if Request.Form("MoMoAccNo") = "Yes" Then%><th>Mobile Money Account Number</th><%End If%>
<%if Request.Form("SkhokhoGSM") = "Yes" Then%><th>Made for Skhokho GSM</th><%End If%>
<%if Request.Form("SkhokhoDedicated") = "Yes" Then%><th>Made for Skhokho Dedicated</th><%End If%>
<%if Request.Form("TradingSpot") = "Yes" Then%><th>Trading Spot</th><%End If%>
</tr>
</thead>
<tbody>
<%
End If
While Not RecReport.EOF
Outline = ""

DDListBrowser = ""
DDListExcel = ""

DDConsentForm = "No"
If RecReport.Fields.Item("DDConsentForm").Value = "True" Then
DDConsentForm = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDConsentForm & "</td>"
DDListExcel = DDListExcel & DDConsentForm & ","

DDCrimCheck = "No"
If RecReport.Fields.Item("DDCrimCheck").Value = "True" Then
DDCrimCheck = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDCrimCheck & "</td>"
DDListExcel = DDListExcel & DDCrimCheck & ","

DDCrimRecord = "No"
If RecReport.Fields.Item("DDCrimRecord").Value = "True" Then
DDCrimRecord = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDCrimRecord & "</td>"
DDListExcel = DDListExcel & DDCrimRecord & ","

DDAMLTrained = "No"
If RecReport.Fields.Item("DDAMLTrained").Value = "True" Then
DDAMLTrained = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDAMLTrained & "</td>"
DDListExcel = DDListExcel & DDAMLTrained & ","

DDAMLPassed = "No"
If RecReport.Fields.Item("DDAMLPassed").Value = "True" Then
DDAMLPassed = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDAMLPassed & "</td>"
DDListExcel = DDListExcel & DDAMLPassed & ","


DDPhoneAllocated = "No"
If RecReport.Fields.Item("DDPhoneAllocated").Value = "True" Then
DDPhoneAllocated = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDPhoneAllocated & "</td>"
DDListExcel = DDListExcel & DDPhoneAllocated & ","

DDMSISDNAllocated = "No"
If RecReport.Fields.Item("DDMSISDNAllocated").Value = "True" Then
DDMSISDNAllocated = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDMSISDNAllocated & "</td>"
DDListExcel = DDListExcel & DDMSISDNAllocated & ","

DDTDROboarded = "No"
If RecReport.Fields.Item("DDTDROboarded").Value = "True" Then
DDTDROboarded = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDTDROboarded & "</td>"
DDListExcel = DDListExcel & DDTDROboarded & ","

DDValidated = "No"
If RecReport.Fields.Item("DDValidated").Value = "True" Then
DDValidated = "Yes"
End If
DDListBrowser = DDListBrowser & "<td>" & DDValidated & "</td>"
DDListExcel = DDListExcel & DDValidated & ","

DDStatusLabel = RecReport.Fields.Item("DDStatusLabel").Value
DDListBrowser = DDListBrowser & "<td>" & DDStatusLabel & "</td>"
DDListExcel = DDListExcel & DDStatusLabel & ","

%>
<tr>
<%if Request.Form("AgentRegion") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("RegionName").Value)%></td><%
Else
Outline = Outline & RecReport.Fields.Item("RegionName").Value & ","
End If
End If%>
<%if Request.Form("AgentEmpCode") = "Yes" Then
AgentCodeMC = ""
AgentCodeMM = ""
If RecReport.Fields.Item("MobileMoneyTedi").Value = "True" Then
AgentCodeMM = RecReport.Fields.Item("TediEmpCode").Value
If Left(AgentCodeMM,1) = "P" Then
AgentCodeMM = "M" & AgentCodeMM
End If
End If

If RecReport.Fields.Item("MChargeTedi").Value = "True" Then
AgentCodeMC = RecReport.Fields.Item("TediEmpCode").Value
If Left(AgentCodeMC,1) = "M" Then
AgentCodeMCT = Len(AgentCodeMC)
AgentCodeMC = Right(AgentCodeMC, AgentCodeMCT - 1)
End If
End If




If OutFormat = "B" Then
%><td><%=(AgentCodeMC)%></td><td><%=(AgentCodeMM)%></td><%
Else
Outline = Outline & AgentCodeMC & "," & AgentCodeMM & ","
End If
End If

if Request.Form("ShowAgentStatus") = "Yes" Then
AgentStatus = "In-Active"
If RecReport.Fields.Item("TediActive").Value = "True" Then
AgentStatus = "Active"
End If
If OutFormat = "B" Then
%><td><%=AgentStatus%></td><%
Else
Outline = Outline & AgentStatus & ","
End If
End If
%>
<%if Request.Form("BDOName") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecReport.Fields.Item("ASLastName").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("ASFirstName").Value & " " & RecReport.Fields.Item("ASLastName").Value & ","
End If
End If%>
<%if Request.Form("AgentName") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecReport.Fields.Item("TediLastName").Value)%></td><%
Else
Outline = Outline & RecReport.Fields.Item("TediFirstName").Value & " " & RecReport.Fields.Item("TediLastName").Value & ","
End If
End If%>



<%if Request.Form("StartDate") = "Yes" Then
SDate = RecReport.Fields.Item("TediStartDate").Value
If SDate <> "" Then

SDateDay = Day(RecReport.Fields.Item("TediStartDate").Value)
If Len(SDateDay) = 1 Then
SDateDay = "0" & SDateDay
End If
SDayMonth = Month(RecReport.Fields.Item("TediStartDate").Value)
If Len(SDayMonth) = 1 Then
SDayMonth = "0" & SDayMonth
End If
SDate = SDateDay & "/" & SDayMonth & "/" & Year(RecReport.Fields.Item("TediStartDate").Value)

Else
SDate = "n/a"
End If
If OutFormat = "B" Then
%><td><%=SDate%></td>
<%
Else
Outline = Outline & SDate & ","
End If
%>
<%End If%>
<%if Request.Form("TermDate") = "Yes" Then
TermDate = RecReport.Fields.Item("TediTermDate").Value
If TermDate <> "" Then
TermDateDay = Day(RecReport.Fields.Item("TediTermDate").Value)
If Len(TermDateDay) = 1 Then
TermDateDay = "0" & TermDateDay
End If
TermDateMonth = Month(RecReport.Fields.Item("TediTermDate").Value)
If Len(TermDateMonth) = 1 Then
TermDateMonth = "0" & TermDateMonth
End If
TermDate = TermDateDay & "/" & TermDateMonth & "/" & Year(RecReport.Fields.Item("TediTermDate").Value)
Else
TermDate = "n/a"
End If
TediParent = "n/a"
If RecReport.Fields.Item("TediParent").Value <> 0 Then
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where  TID = " & RecReport.Fields.Item("TediParent").Value
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
TediParent = RecEdit.Fields.Item("TediFirstName").Value & " " & RecEdit.Fields.Item("TediLastName").Value
End If
If OutFormat = "B" Then
%><td><%=TermDate%></td>
<%
Else
Outline = Outline & TermDate & ","
End If
End If%>
<%if Request.Form("AgentCell") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TediCell").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TediCell").Value & ","
End If
End If%>

<%if Request.Form("IDNo") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=(RecReport.Fields.Item("IDNumber").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("IDNumber").Value & ","
End If
%>
<%
End If

If Request.Form("WPExpiryDate") = "Yes" Then
WorkPermitExpiryDate = "NA"
If IsDate(RecReport.Fields.Item("WorkPermitExpiryDate").Value) = "True" Then
WorkPermitExpiryDate = Day(RecReport.Fields.Item("WorkPermitExpiryDate").Value) & " " & MonthName(Month(RecReport.Fields.Item("WorkPermitExpiryDate").Value)) & " " & Year(RecReport.Fields.Item("WorkPermitExpiryDate").Value)
End If
If OutFormat = "B" Then%>
<td><%=WorkPermitExpiryDate%></td>
<%
Else
Outline = Outline & WorkPermitExpiryDate & ","
End If
End If

if Request.Form("LastBankedDate") = "Yes" Then
LastBankedDate = RecReport.Fields.Item("LastBankedDate").Value
If LastBankedDate <> "" Then
LastBankedDay = Day(RecReport.Fields.Item("LastBankedDate").Value)
If Len(LastBankedDay) = 1 Then
LastBankedDay = "0" & LastBankedDay
End If
LastBankedMonth = Month(RecReport.Fields.Item("LastBankedDate").Value)
If Len(LastBankedMonth) = 1 Then
LastBankedMonth = "0" & LastBankedMonth
End If  
LastBankedDate = LastBankedDay & "/" & LastBankedMonth & "/" & Year(RecReport.Fields.Item("LastBankedDate").Value)
Else
LastBankedDate = "n/a"
End If
If OutFormat = "B" Then
%><td><%=LastBankedDate%></td>
<%
Else
Outline = Outline & LastBankedDate & ","
End If
End If
if Request.Form("ShowLastRefund") = "Yes" Then
LastAmount = 0

set RecLastTrans = Server.CreateObject("ADODB.Recordset")
RecLastTrans.ActiveConnection = MM_Site_STRING
RecLastTrans.Source = "EXECUTE SPLastTransAction @TID = " & RecReport.Fields.Item("TID").Value & ", @ctype = 2"
'RecLastTrans.Source = "SELECT Top(1)* FROM TediTransactions Where TediID = " & RecReport.Fields.Item("TID").Value & " and CType = 2 order by CDate Desc"
RecLastTrans.CursorType = 0
RecLastTrans.CursorLocation = 2
RecLastTrans.LockType = 3
RecLastTrans.Open()
RecLastTrans_numRows = 0
If Not RecLastTrans.EOF and Not RecLastTrans.BOF Then
LastAmount = RecLastTrans.Fields.Item("CAmount").Value
End If
If OutFormat = "B" Then
%><td><%=LastAmount%></td><%
Else
Outline = Outline & LastAmount & ","
End If
End If


ATAmount = "0"
ATDate = "N/A"
set RecLastAirtime = Server.CreateObject("ADODB.Recordset")
RecLastAirtime.ActiveConnection = MM_Site_STRING
RecLastAirtime.Source = "EXECUTE SPLastTransAction @TID = " & RecReport.Fields.Item("TID").Value & ", @ctype = 1"
RecLastAirtime.CursorType = 0
RecLastAirtime.CursorLocation = 2
RecLastAirtime.LockType = 3
RecLastAirtime.Open()
RecLastAirtime_numRows = 0
If Not RecLastAirtime.EOF and Not RecLastAirtime.BOF then
ATAmount = RecLastAirtime.Fields.Item("CAmount").Value
ATDate = RecLastAirtime.Fields.Item("CDate").Value

If ATDate <> "" Then
ATDateDay = Day(RecLastAirtime.Fields.Item("CDate").Value)
If Len(ATDateDay) = 1 Then
ATDateDay = "0" & ATDateDay
End If
ATDateMonth = Month(RecLastAirtime.Fields.Item("CDate").Value)
If Len(ATDateMonth) = 1 Then
ATDateMonth = "0" & ATDateMonth
End If  
ATDate = ATDateDay & "/" & ATDateMonth & "/" & Year(RecLastAirtime.Fields.Item("CDate").Value)
End If
End If


if Request.Form("LastAirtimeDate") = "Yes" Then
If OutFormat = "B" Then
%><td><%=ATDate%></td>
<%
Else
Outline = Outline & ATDate & ","
End If
End If

if Request.Form("ShowLastAirtimeAmount") = "Yes" Then
If OutFormat = "B" Then
%><td><%=ATAmount%></td>
<%
Else
Outline = Outline & ATAmount & ","
End If
End If

if Request.Form("MChargeBal") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("MChargeBalance").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("MChargeBalance").Value & ","
End If
End If%>
<%if Request.Form("PurseLimit") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("PurseLimit").Value)%></td><%
Else
Outline = Outline & RecReport.Fields.Item("PurseLimit").Value & ","
End If
End If
%>
<%if Request.Form("Bank") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("BankLabel").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("BankLabel").Value & ","
End If
End If%>
<%if Request.Form("AccNo") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("AccNo").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("AccNo").Value & ","
End If
End If%>
<%if Request.Form("Branch") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("BranchCode").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("BranchCode").Value & ","
End If
End If%>
<%if Request.Form("AccType") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("AccountLabel").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("AccountLabel").Value & ","
End If
End If

if Request.Form("ShowMobiSitePass") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TediPassword").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TediPassword").Value & ","
End If
End If
if Request.Form("AgentCell2") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TediCell2").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TediCell2").Value & ","
End If
End If

if Request.Form("TertiaryMobileNumber") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TertiaryMobileNumber").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TertiaryMobileNumber").Value & ","
End If
End If



if Request.Form("ExcludeFromMchargeBulkFile") = "Yes" Then
ExcludeTedi = "No"
If RecReport.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then
ExcludeTedi = "Yes"
End If
If OutFormat = "B" Then
%><td><%=ExcludeTedi%></td><%
Else
Outline = Outline & ExcludeTedi & ","
End If
End If

if Request.Form("ShowRegionAlManager") = "Yes" Then
If OutFormat = "B" Then
%><td><%=RecReport.Fields.Item("RegManagerFirstName").Value & " " & RecReport.Fields.Item("RegManagerLastName").Value%></td><%
Else
Outline = Outline & RecReport.Fields.Item("RegManagerFirstName").Value & " " & RecReport.Fields.Item("RegManagerLastName").Value & ","
End If
End If

if Request.Form("AgentSubRegion") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=(RecReport.Fields.Item("SubRegionName").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("SubRegionName").Value & ","
End If
End If

if Request.Form("Gender") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("GenderType").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("GenderType").Value & ","
End If
End If%>
<%if Request.Form("Race") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("RaceLabel").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("RaceLabel").Value & ","
End If
End If%>




<%if Request.Form("TaxOffice") = "Yes" Then

If BEESigDate <> "" Then
BEEDay = Day(RecReport.Fields.Item("TaxOffice").Value)
If Len(BEEDay) = 1 Then
BEEDay = "0" & BEEDay
End If
BEEMonth = Month(RecReport.Fields.Item("TaxOffice").Value)
If Len(BEEMonth) = 1 Then
BEEMonth = "0" & BEEMonth
End If
BEESigDate = BEEDay & "/" & BEEMonth & "/" & Year(RecReport.Fields.Item("TaxOffice").Value)
Else
BEESigDate = "n/a"
End If


If OutFormat = "B" Then
%><td><%=BEESigDate%></td>
<%
Else
Outline = Outline & BEESigDate & ","
End If
End If%>


<%if Request.Form("ResidentialAddress") = "Yes" Then

ResAddress = RecReport.Fields.Item("ResidentialAddress1").Value & " " & RecReport.Fields.Item("ResidentialAddress2").Value & " " & RecReport.Fields.Item("ResidentialAddress3").Value & " " & RecReport.Fields.Item("ResidentialCode").Value
If ResAddress <> "" Then
ResAddress = Replace(ResAddress, ",", " ")
End If

If OutFormat = "B" Then
%><td><%=(ResAddress)%></td>
<%
Else
Outline = Outline & ResAddress & ","
End If
End If%>



<%if Request.Form("TermReason") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("TermReason").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TermReason").Value & ","
End If
End If%>

<%if Request.Form("BDOEmpCode") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("ASEmpCode").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("ASEmpCode").Value & ","
End If
End If%>

<%if Request.Form("SARS") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("ClusterName").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("ClusterName").Value & ","
End If
End If%>

<%if Request.Form("SubRegionCode") = "Yes" Then
If OutFormat = "B" Then
%><td><%=(RecReport.Fields.Item("SubRegionCode").Value)%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("SubRegionCode").Value & ","
End If
End If%>

<%if Request.Form("RealTimeCommOptIn") = "Yes" Then
RealTimeCommOptInLabel = "Yes"
If RecReport.Fields.Item("RealTimeCommOptIn").Value = "False" Then
RealTimeCommOptInLabel = "No"
End If
If OutFormat = "B" Then
%><td><%=RealTimeCommOptInLabel%></td>
<%
Else
Outline = Outline & RealTimeCommOptInLabel & ","
End If
End If%>


<%
If Request.Form("ShowIfMchargeAgent") = "Yes" Then
IsMchargeAgent = "Yes"
If RecReport.Fields.Item("MChargeTedi").Value = "False" Then
IsMchargeAgent = "No"
End If
If OutFormat = "B" Then
%>
<td><%=IsMchargeAgent%></td>
<%
Else
Outline = Outline & IsMchargeAgent & ","
End If
End If
%>
<%
If Request.Form("ShowIfMobileMoneyAgent") = "Yes" Then
IsMobileMoneyAgent = "Yes"
If RecReport.Fields.Item("MobileMoneyTedi").Value = "False" Then
IsMobileMoneyAgent = "No"
End If

If OutFormat = "B" Then%>
<td><%=IsMobileMoneyAgent%></td>
<%
Else
Outline = Outline & IsMobileMoneyAgent & ","
End If
End If

if Request.Form("DDList") = "Yes" Then
If OutFormat = "B" Then%>
<%=DDListBrowser%>
<%
Else
Outline = Outline & DDListExcel & ","
End If
End If
%>
<%
If Request.Form("MMPurseLimit") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=RecReport.Fields.Item("PurseLimitMM").Value%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("PurseLimitMM").Value & ","
End If
End If

If Request.Form("MobileMoneyBalance") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=RecReport.Fields.Item("MobileMoneyBalance").Value%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("MobileMoneyBalance").Value & ","
End If
End If



If Request.Form("MMLastBankedDate") = "Yes" Then
LastBankedDateMM = RecReport.Fields.Item("LastBankedDateMM").Value
If LastBankedDateMM <> "" Then
LastBankedMMDay = Day(RecReport.Fields.Item("LastBankedDateMM").Value)
If Len(LastBankedMMDay) = 1 Then
LastBankedMMDay = "0" & LastBankedMMDay
End If
LastBankedMMMonth = Month(RecReport.Fields.Item("LastBankedDateMM").Value)
If Len(LastBankedMMMonth) = 1 Then
LastBankedMMMonth = "0" & LastBankedMMMonth
End If  
LastBankedDateMM = LastBankedMMDay & "/" & LastBankedMMMonth & "/" & Year(RecReport.Fields.Item("LastBankedDateMM").Value)
Else
LastBankedDateMM = "n/a"
End If
If OutFormat = "B" Then%>
<td><%=LastBankedDateMM%></td>
<%
Else
Outline = Outline & LastBankedDateMM & ","
End If
End If

If Request.Form("WPED") = "Yes"  Then
WPED = "N/A"
If IsDate(RecReport.Fields.Item("WorkPermitExpiryDate").Value) = "True" Then
WPEDDay = Day(RecReport.Fields.Item("WorkPermitExpiryDate").Value)
If Len(WPEDDay) = 1 Then
WPEDDay = "0" & WPEDDay
End If
WPEDMonth = Month(RecReport.Fields.Item("WorkPermitExpiryDate").Value)
If Len(WPEDMonth) = 1 Then
WPEDMonth = "0" & WPEDMonth
End If

WPED =  WPEDDay & "/" & WPEDMonth & "/" & Year(RecReport.Fields.Item("WorkPermitExpiryDate").Value)
End If
If OutFormat = "B" Then%>
<td><%=WPED%></td>
<%
Else
Outline = Outline & WPED & ","
End If
End If


If Request.Form("MoMoAccNo") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=RecReport.Fields.Item("MoMoAccNo").Value%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("MoMoAccNo").Value & ","
End If
End If

If Request.Form("SkhokhoGSM") = "Yes" Then
DDSkhokhoGSMLab = "No"
If RecReport.Fields.Item("DDSkhokhoGSM").Value = "True" Then
DDSkhokhoGSMLab = "Yes"
End If

If OutFormat = "B" Then%>
<td><%=DDSkhokhoGSMLab%></td>
<%
Else
Outline = Outline & DDSkhokhoGSMLab & ","
End If
End If

If Request.Form("SkhokhoDedicated") = "Yes" Then
DDSkhokhoDedicatedLab = "No"
If RecReport.Fields.Item("DDSkhokhoDedicated").Value = "True" Then
DDSkhokhoDedicatedLab = "Yes"
End If
If OutFormat = "B" Then%>
<td><%=DDSkhokhoDedicatedLab%></td>
<%
Else
Outline = Outline & DDSkhokhoDedicatedLab & ","
End If
End If

If Request.Form("TradingSpot") = "Yes" Then
If OutFormat = "B" Then%>
<td><%=RecReport.Fields.Item("TradingSpot").Value%></td>
<%
Else
Outline = Outline & RecReport.Fields.Item("TradingSpot").Value & ","
End If
End If
%>
<%
If OutFormat = "B" Then
%>
</tr>
<%
End If
If OutFormat <> "B" Then
OutlineT = Len(Outline)
Outline = Left(Outline, OutlineT - 1)
TheFile.Writeline(Outline)
End If
RecReport.MoveNext
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