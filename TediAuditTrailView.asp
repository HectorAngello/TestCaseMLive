<!-- #include file="includes/header.asp" -->
<%
set RecZonerHistory = Server.CreateObject("ADODB.Recordset")
RecZonerHistory.ActiveConnection = MM_Site_STRING
RecZonerHistory.Source = "SELECT * FROM ViewTediAuditTrail Where ID = '" & Request.QueryString("ID") & "'"
RecZonerHistory.CursorType = 0
RecZonerHistory.CursorLocation = 2
RecZonerHistory.LockType = 3
RecZonerHistory.Open()
RecZonerHistory_numRows = 0

set RecSiteInfo = Server.CreateObject("ADODB.Recordset")
RecSiteInfo.ActiveConnection = MM_Site_STRING
RecSiteInfo.Source = "SELECT * FROM SiteInfo Where ID = 1"
RecSiteInfo.CursorType = 0
RecSiteInfo.CursorLocation = 2
RecSiteInfo.LockType = 3
RecSiteInfo.Open()
RecSiteInfo_numRows = 0

HistoryAvail = "No"
set RecEarlierHistory = Server.CreateObject("ADODB.Recordset")
RecEarlierHistory.ActiveConnection = MM_Site_STRING
RecEarlierHistory.Source = "SELECT * FROM ViewTediAuditTrail Where ID < '" & Request.QueryString("ID") & "' and TID = " & RecZonerHistory.Fields.Item("TID").Value & " Order By ID Desc"
RecEarlierHistory.CursorType = 0
RecEarlierHistory.CursorLocation = 2
RecEarlierHistory.LockType = 3
RecEarlierHistory.Open()
RecEarlierHistory_numRows = 0

If Not RecEarlierHistory.EOF and Not RecEarlierHistory.BOF Then
HistoryAvail = "Yes"
End If

If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
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
                    <div class="row heading">
                        <div class="eight columns"><h1>View Agent Audit Trail:</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
                    </div>
<%If HistoryAvail = "Yes" Then%>
<table border="0" cellspacing="2" cellpadding="2">
<tr>
            <td><strong>Change By:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("UserFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("UserLastName").Value)%></td>
            <td><strong>Change Date:</strong></td>
            <td><%=FormatDateTime(RecZonerHistory.Fields.Item("Transdate").Value,1)%> : <%=FormatDateTime(RecZonerHistory.Fields.Item("Transdate").Value,3)%></td>
          </tr>
<tr>
            <td><strong>Type:</strong></td>
            <td class="
<%If RecZonerHistory.Fields.Item("TransType").Value <> RecEarlierHistory.Fields.Item("TransType").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TransType").Value)%></td>
            <td width="120" class="quote"><strong>Zoner Active:</strong></td>
<%
AgentActiveLabel = "Yes"
If RecZonerHistory.Fields.Item("TediActive").Value = "False" Then
AgentActiveLabel = "No"
End If
%>
            <td class="<%If RecZonerHistory.Fields.Item("TediActive").Value <> RecEarlierHistory.Fields.Item("TediActive").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=AgentActiveLabel%></td>
          </tr>

<tr>
            <td><strong>Name:</strong></td>
            <td class="<%If (RecZonerHistory.Fields.Item("TediFirstName").Value & " " & RecZonerHistory.Fields.Item("TediLastName").Value) <> (RecEarlierHistory.Fields.Item("TediFirstName").Value & " " & RecEarlierHistory.Fields.Item("TediLastName").Value) Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("TediLastName").Value)%></td>
            <td><strong>Agent Code:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TediEmpCode").Value <> RecEarlierHistory.Fields.Item("TediEmpCode").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediEmpCode").Value)%></td>
          </tr>
          <tr>
            <td><strong>Primary Mobile:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TediCell").Value <> RecEarlierHistory.Fields.Item("TediCell").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediCell").Value)%>&nbsp;</td>
            <td><strong>Email Address:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TediEmail").Value <> RecEarlierHistory.Fields.Item("TediEmail").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediEmail").Value)%>&nbsp;</td>
          </tr>
          <tr>
            <td><strong>Secondary Mobile:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TediCell2").Value <> RecEarlierHistory.Fields.Item("TediCell2").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediCell2").Value)%>&nbsp;</td>
            <td><strong>Tertiary Mobile Number:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TertiaryMobileNumber").Value <> RecEarlierHistory.Fields.Item("TertiaryMobileNumber").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TertiaryMobileNumber").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Mobi Site Password:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TediPassword").Value <> RecEarlierHistory.Fields.Item("TediPassword").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TediPassword").Value)%>&nbsp;</td>
            <td><Strong>ID Number:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("IDNumber").Value <> RecEarlierHistory.Fields.Item("IDNumber").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("IDNumber").Value)%>&nbsp;</td>
          </tr>

          <tr>
            <td><strong>BEE Signature Date:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TaxOffice").Value <> RecEarlierHistory.Fields.Item("TaxOffice").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TaxOffice").Value)%></td>
            <td><strong>Cluster:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TaxNumber").Value <> RecEarlierHistory.Fields.Item("TaxNumber").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ClusterName").Value)%></td>
          </tr>
	  <tr>
            <td><Strong>Race:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("RaceID").Value <> RecEarlierHistory.Fields.Item("RaceID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("RaceLabel").Value)%>&nbsp;</td>
            <td><strong>Gender:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("GenderID").Value <> RecEarlierHistory.Fields.Item("GenderID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("GenderType").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>Residential Address 1:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ResidentialAddress1").Value <> RecEarlierHistory.Fields.Item("ResidentialAddress1").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ResidentialAddress1").Value)%>&nbsp;</td>
            <td><Strong>Residential Address 2:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ResidentialAddress2").Value <> RecEarlierHistory.Fields.Item("ResidentialAddress2").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ResidentialAddress2").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Residential Address 3:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ResidentialAddress3").Value <> RecEarlierHistory.Fields.Item("ResidentialAddress3").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ResidentialAddress3").Value)%>&nbsp;</td>
            <td><strong>Residential Address Code:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ResidentialCode").Value <> RecEarlierHistory.Fields.Item("ResidentialCode").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ResidentialCode").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            
            <td width="120" class="quote"><strong>Termination Date:</strong></td>
<%
TermDate = RecZonerHistory.Fields.Item("TediTermDate").Value
If TermDate <> "" Then
TermDate = FormatDateTime(TermDate,1)
Else
TermDate = "n/a"
end If
%>
            <td class="auditsame"><%=(Termdate)%></td>
	    <td width="120" class="quote"><b>Termination Reason:</b></td>
            <td class="<%If RecZonerHistory.Fields.Item("TermReason").Value <> RecEarlierHistory.Fields.Item("TermReason").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TermReason").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td width="120" class="quote"><strong>Bank:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("BankID").Value <> RecEarlierHistory.Fields.Item("BankID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("BankLabel").Value)%>&nbsp;</td>
            <td width="120" class="quote"><Strong>Branch:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("BranchCode").Value <> RecEarlierHistory.Fields.Item("BranchCode").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("BranchCode").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td width="120" class="quote"><strong>Account Type:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("AccountType").Value <> RecEarlierHistory.Fields.Item("AccountType").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("AccountLabel").Value)%>&nbsp;</td>
            <td width="120" class="quote"><Strong>Account Number:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("AccNo").Value <> RecEarlierHistory.Fields.Item("AccNo").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("AccNo").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td width="120" class="quote"><strong>Region:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("RID").Value <> RecEarlierHistory.Fields.Item("RID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("RegionName").Value)%>&nbsp;</td>
            <td width="120" class="quote"><strong>Sub Region:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("SRID").Value <> RecEarlierHistory.Fields.Item("SRID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("SubRegionName").Value)%>&nbsp;</td>
          </tr>
<%
MChargeExclude = "No"
If RecZonerHistory.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then
MChargeExclude = "Yes"
End If

OnWatchList = "No"
If RecZonerHistory.Fields.Item("OnWatchList").Value = "True" Then
OnWatchList = "Yes"
End If
%>
	  <tr>
            <td width="120" class="quote"><strong>Purse Limit:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("PurseLimit").Value <> RecEarlierHistory.Fields.Item("PurseLimit").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("PurseLimit").Value)%>&nbsp;</td>
            <td width="120" class="quote"><Strong>Exclude From M-Charge File Creation:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ExcludeFromMchargeBulkFile").Value <> RecEarlierHistory.Fields.Item("ExcludeFromMchargeBulkFile").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=MChargeExclude%>&nbsp;</td>
          </tr>
<%
AgentRealTimeComm = "No"
If RecZonerHistory.Fields.Item("RealTimeCommOptIn").Value = "True" Then
AgentRealTimeComm = "Yes"
End If
%>
	  <tr>
            <td width="120" class="quote"><strong>On Watch List:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("OnWatchList").Value <> RecEarlierHistory.Fields.Item("OnWatchList").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(OnWatchList)%>&nbsp;</td>
            <td width="120" class="quote"><strong>Agent opted in for real time airtime commission</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("RealTimeCommOptIn").Value <> RecEarlierHistory.Fields.Item("RealTimeCommOptIn").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=AgentRealTimeComm%></td>
          </tr>
	  <tr>
            <td><strong>Airtime Allocation Type:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("AirtimeTypeID").Value <> RecEarlierHistory.Fields.Item("AirtimeTypeID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("AirtimeAlloLabel").Value)%>&nbsp;</td>
            <td><strong>Mobile Money Purse Limit:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("PurseLimitMM").Value <> RecEarlierHistory.Fields.Item("PurseLimitMM").Value Then%>auditdifferent<%Else%>auditsame<%End If%>">R <%=(RecZonerHistory.Fields.Item("PurseLimitMM").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>Airtime Agent:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("MChargeTedi").Value <> RecEarlierHistory.Fields.Item("MChargeTedi").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("MChargeTedi").Value)%>&nbsp;</td>
            <td><strong>Mobile Money Agent:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("MobileMoneyTedi").Value <> RecEarlierHistory.Fields.Item("MobileMoneyTedi").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("MobileMoneyTedi").Value)%>&nbsp;</td>
          </tr>
<%
WPED = "N/A"
If IsDate(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value) = "True" Then
WPED = Day(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value) & " " & MonthName(Month(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value)) & " " & Year(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value)
End If
%>
          <tr>
            <td><strong>Work Permit Expiry Date:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value <> RecEarlierHistory.Fields.Item("WorkPermitExpiryDate").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=WPED%></td>
            <td><strong>Mobile Money AccNo</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("MoMoAccNo").Value <> RecEarlierHistory.Fields.Item("MoMoAccNo").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("MoMoAccNo").Value)%>&nbsp;</td>
          </tr>

  <tr>
            <td><strong>Consent Form Submitted:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDConsentForm").Value <> RecEarlierHistory.Fields.Item("DDConsentForm").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDConsentForm").Value)%>&nbsp;</td>
            <td><strong>Crim Check:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDCrimCheck").Value <> RecEarlierHistory.Fields.Item("DDCrimCheck").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDCrimCheck").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>Crim Record:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDCrimRecord").Value <> RecEarlierHistory.Fields.Item("DDCrimRecord").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDCrimRecord").Value)%>&nbsp;</td>
            <td><strong>AML Trained:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDAMLTrained").Value <> RecEarlierHistory.Fields.Item("DDAMLTrained").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDAMLTrained").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>AML Passed:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDAMLPassed").Value <> RecEarlierHistory.Fields.Item("DDAMLPassed").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDAMLPassed").Value)%>&nbsp;</td>
            <td><strong>Phone Allocated:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDPhoneAllocated").Value <> RecEarlierHistory.Fields.Item("DDPhoneAllocated").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDPhoneAllocated").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>MSISDN Allocated:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDMSISDNAllocated").Value <> RecEarlierHistory.Fields.Item("DDMSISDNAllocated").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDMSISDNAllocated").Value)%>&nbsp;</td>
            <td><strong>TDR Onboarded :</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDTDROboarded").Value <> RecEarlierHistory.Fields.Item("DDTDROboarded").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDTDROboarded").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>Validated:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDValidated").Value <> RecEarlierHistory.Fields.Item("DDValidated").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDValidated").Value)%>&nbsp;</td>
            <td><strong>Agent Status:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDStatusID").Value <> RecEarlierHistory.Fields.Item("DDStatusID").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDStatusLabel").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Made for Skhokho GSM:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDSkhokhoGSM").Value <> RecEarlierHistory.Fields.Item("DDSkhokhoGSM").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDSkhokhoGSM").Value)%>&nbsp;</td>
            <td><strong>Made for Skhokho Dedicated:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("DDSkhokhoDedicated").Value <> RecEarlierHistory.Fields.Item("DDSkhokhoDedicated").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("DDSkhokhoDedicated").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Trading Spot:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TradingSpot").Value <> RecEarlierHistory.Fields.Item("TradingSpot").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TradingSpot").Value)%>&nbsp;</td>
            <td><strong></strong></td>
            <td></td>
          </tr>
</table>
<%Else%>
<table border="0" cellspacing="2" cellpadding="2">
<tr>
            <td><strong>Change By:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("UserFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("UserLastName").Value)%></td>
            <td><strong>Change Date:</strong></td>
            <td><%=FormatDateTime(RecZonerHistory.Fields.Item("Transdate").Value,1)%> : <%=FormatDateTime(RecZonerHistory.Fields.Item("Transdate").Value,3)%></td>
          </tr>
<tr>
            <td><strong>Type:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TransType").Value)%></td>
            <td><strong>Zoner Active:</strong></td>
<%
AgentActiveLabel = "Yes"
If RecZonerHistory.Fields.Item("TediActive").Value = "False" Then
AgentActiveLabel = "No"
End If
%>
            <td><%=AgentActiveLabel%></td>
          </tr>

<tr>
            <td><strong>Name:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("TediLastName").Value)%></td>
            <td><strong>Agent Code:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediEmpCode").Value)%></td>
          </tr>
          <tr>
            <td><strong>Primary Mobile:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediCell").Value)%>&nbsp;</td>
            <td><strong>Email Address:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediEmail").Value)%>&nbsp;</td>
          </tr>
          <tr>
            <td><strong>Secondary Mobile:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediCell2").Value)%>&nbsp;</td>
            <td><strong>Tertiary Mobile Number:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TertiaryMobileNumber").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Mobi Site Password:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TediPassword").Value)%>&nbsp;</td>
            <td><Strong>ID Number:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("IDNumber").Value)%>&nbsp;</td>
          </tr>

          <tr>
            <td><strong>BEE Signature Date:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TaxOffice").Value)%></td>
            <td><strong>Cluster:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ClusterName").Value)%></td>
          </tr>
		            <tr>
            <td><Strong>Race:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("RaceLabel").Value)%>&nbsp;</td>
            <td><strong>Gender:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("GenderType").Value)%>&nbsp;</td>
          </tr>

	  <tr>
            <td><strong>Residential Address 1:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ResidentialAddress1").Value)%>&nbsp;</td>
            <td><Strong>Residential Address 2:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ResidentialAddress2").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Residential Address 3:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ResidentialAddress3").Value)%>&nbsp;</td>
            <td><Strong>Residential Address Code:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ResidentialCode").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Termination Date:</strong></td>
<%
TermDate = RecZonerHistory.Fields.Item("TediTermDate").Value
If TermDate <> "" Then
TermDate = FormatDateTime(TermDate,1)
Else
TermDate = "n/a"
end If
%>
            <td><%=(Termdate)%></td>
	    <td><b>Termination Reason:</b></td>
            <td><%=(RecZonerHistory.Fields.Item("TermReason").Value)%>&nbsp;</td>
          </tr>


	  <tr>
            <td><strong>Bank:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("BankLabel").Value)%>&nbsp;</td>
            <td><Strong>Branch:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("BranchCode").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Account Type:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("AccountLabel").Value)%>&nbsp;</td>
            <td><Strong>Account Number:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("AccNo").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Region:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("RegionName").Value)%>&nbsp;</td>
            <td><strong>Sub Region:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("SubRegionName").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Purse Limit:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("PurseLimit").Value)%>&nbsp;</td>
            <td><strong>Exclude From M-Charge File Creation:</strong></td>
<%
MChargeExclude = "No"
If RecZonerHistory.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then
MChargeExclude = "Yes"
End If

OnWatchList = "No"
If RecZonerHistory.Fields.Item("OnWatchList").Value = "True" Then
OnWatchList = "Yes"
End If

AgentRealTimeComm = "No"
If RecZonerHistory.Fields.Item("RealTimeCommOptIn").Value = "True" Then
AgentRealTimeComm = "Yes"
End If
%>
            <td><%=(MChargeExclude)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>On Watch List:</strong></td>
            <td><%=(OnWatchList)%>&nbsp;</td>
            <td><strong>Agent opted in for real time airtime commission</strong></td>
            <td><%=AgentRealTimeComm%></td>
          </tr>
	  <tr>
            <td><strong>Airtime Allocation Type:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("AirtimeAlloLabel").Value)%>&nbsp;</td>
            <td><strong>Mobile Money Purse Limit:</strong></td>
            <td>R <%=(RecZonerHistory.Fields.Item("PurseLimitMM").Value)%></td>
          </tr>

	  <tr>
            <td><strong>Airtime Agent:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("MChargeTedi").Value)%>&nbsp;</td>
            <td><strong>Mobile Money Agent:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("MobileMoneyTedi").Value)%></td>
          </tr>

<%
WPED = "N/A"
If IsDate(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value) = "True" Then
WPED = Day(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value) & " " & MonthName(Month(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value)) & " " & Year(RecZonerHistory.Fields.Item("WorkPermitExpiryDate").Value)
End If
%>

          <tr>
            <td><strong>Work Permit Expiry Date:</strong></td>
            <td><%=WPED%></td>
            <td><strong>Mobile Money AccNo</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("MoMoAccNo").Value)%></td>
          </tr>

	  <tr>
            <td><strong>Consent Form Submitted:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDConsentForm").Value)%>&nbsp;</td>
            <td><strong>Crim Check:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDCrimCheck").Value)%></td>
          </tr>

	  <tr>
            <td><strong>Crim Record:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDCrimRecord").Value)%>&nbsp;</td>
            <td><strong>AML Trained:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDAMLTrained").Value)%></td>
          </tr>

	  <tr>
            <td><strong>AML Passed:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDAMLPassed").Value)%>&nbsp;</td>
            <td><strong>Phone Allocated:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDPhoneAllocated").Value)%></td>
          </tr>

	  <tr>
            <td><strong>MSISDN Allocated:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDMSISDNAllocated").Value)%>&nbsp;</td>
            <td><strong>TDR Onboarded :</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDTDROboarded").Value)%></td>
          </tr>

	  <tr>
            <td><strong>Validated:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDValidated").Value)%>&nbsp;</td>
            <td><strong>Agent Status:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDStatusLabel").Value)%></td>
          </tr>
	  <tr>
            <td><strong>Made for Skhokho GSM:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDSkhokhoGSM").Value)%>&nbsp;</td>
            <td><strong>Made for Skhokho Dedicated:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("DDSkhokhoDedicated").Value)%></td>
          </tr>
	  <tr>
            <td><strong>Trading Spot:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TradingSpot").Value)%>&nbsp;</td>
            <td><strong></strong></td>
            <td></td>
          </tr>
</table>
<%End If%>

<br><br><%If HistoryAvail = "No" Then%>No History Available Prior To This To Compare Data With<%Else%>Data Being Compared To Changes Made on <a href="TediAuditTrailView.asp?ID=<%=(RecEarlierHistory.Fields.Item("ID").Value)%>"><%=FormatDateTime(RecEarlierHistory.Fields.Item("TransDate").Value,1)%> at <%=FormatDateTime(RecEarlierHistory.Fields.Item("TransDate").Value,3)%></a><%End If%>

<!-- #include file="includes/footer.asp" -->

