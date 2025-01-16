<!-- #include file="includes/header.asp" -->
<%
set RecZonerHistory = Server.CreateObject("ADODB.Recordset")
RecZonerHistory.ActiveConnection = MM_Site_STRING
RecZonerHistory.Source = "SELECT * FROM ViewASAuditTrail Where ID = '" & Request.QueryString("ID") & "'"
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
RecEarlierHistory.Source = "SELECT * FROM ViewASAuditTrail Where ID < '" & Request.QueryString("ID") & "' and ASID = " & RecZonerHistory.Fields.Item("ASID").Value & " Order By ID Desc"
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
                        <div class="eight columns"><h1>View <%=SupervisorLabel%> Audit Trail:</h1></div>
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
            <td width="120" class="quote"><strong>Zone Manager Active:</strong></td>
<%
AgentActiveLabel = "Yes"
If RecZonerHistory.Fields.Item("ASActive").Value = "False" Then
AgentActiveLabel = "No"
End If
%>
            <td class="<%If RecZonerHistory.Fields.Item("ASActive").Value <> RecEarlierHistory.Fields.Item("ASActive").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=AgentActiveLabel%></td>
          </tr>

<tr>
            <td><strong>Name:</strong></td>
            <td class="<%If (RecZonerHistory.Fields.Item("ASFirstName").Value & " " & RecZonerHistory.Fields.Item("ASLastName").Value) <> (RecEarlierHistory.Fields.Item("ASFirstName").Value & " " & RecEarlierHistory.Fields.Item("ASLastName").Value) Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("ASLastName").Value)%></td>
            <td><strong><%=SupervisorLabel%> Code:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ASEmpCode").Value <> RecEarlierHistory.Fields.Item("ASEmpCode").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ASEmpCode").Value)%></td>
          </tr>
          <tr>
            <td><strong>Mobile:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ASCell").Value <> RecEarlierHistory.Fields.Item("ASCell").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ASCell").Value)%>&nbsp;</td>
            <td><strong>Email Address:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ASEmail").Value <> RecEarlierHistory.Fields.Item("ASEmail").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ASEmail").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Mobi Site Password:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("ASPassword").Value <> RecEarlierHistory.Fields.Item("ASPassword").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("ASPassword").Value)%>&nbsp;</td>
            <td><Strong>ID Number:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("IDNumber").Value <> RecEarlierHistory.Fields.Item("IDNumber").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("IDNumber").Value)%>&nbsp;</td>
          </tr>

          <tr>
            <td><strong>Tax Office:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TaxOffice").Value <> RecEarlierHistory.Fields.Item("TaxOffice").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TaxOffice").Value)%></td>
            <td><strong>SARS Tax Number:</strong></td>
            <td class="<%If RecZonerHistory.Fields.Item("TaxNumber").Value <> RecEarlierHistory.Fields.Item("TaxNumber").Value Then%>auditdifferent<%Else%>auditsame<%End If%>"><%=(RecZonerHistory.Fields.Item("TaxNumber").Value)%></td>
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
TermDate = RecZonerHistory.Fields.Item("ASTermDate").Value
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
            <td><strong>Zone Manager Active:</strong></td>
<%
AgentActiveLabel = "Yes"
If RecZonerHistory.Fields.Item("ASActive").Value = "False" Then
AgentActiveLabel = "No"
End If
%>
            <td><%=AgentActiveLabel%></td>
          </tr>

<tr>
            <td><strong>Name:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecZonerHistory.Fields.Item("ASLastName").Value)%></td>
            <td><strong><%=SupervisorLabel%> Code:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ASEmpCode").Value)%></td>
          </tr>
          <tr>
            <td><strong>Mobile:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ASCell").Value)%>&nbsp;</td>
            <td><strong>Email Address:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ASEmail").Value)%>&nbsp;</td>
          </tr>
	  <tr>
            <td><strong>Mobi Site Password:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("ASPassword").Value)%>&nbsp;</td>
            <td><Strong>ID Number:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("IDNumber").Value)%>&nbsp;</td>
          </tr>

          <tr>
            <td><strong>Tax Office:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TaxOffice").Value)%></td>
            <td><strong>SARS Tax Number:</strong></td>
            <td><%=(RecZonerHistory.Fields.Item("TaxNumber").Value)%></td>
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
TermDate = RecZonerHistory.Fields.Item("ASTermDate").Value
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
          </tr>

</table>
<%End If%>

<br><br><%If HistoryAvail = "No" Then%>No History Available Prior To This To Compare Data With<%Else%>Data Being Compared To Changes Made on <a href="ASAuditTrailView.asp?ID=<%=(RecEarlierHistory.Fields.Item("ID").Value)%>"><%=FormatDateTime(RecEarlierHistory.Fields.Item("TransDate").Value,1)%> at <%=FormatDateTime(RecEarlierHistory.Fields.Item("TransDate").Value,3)%></a><%End If%>

<!-- #include file="includes/footer.asp" -->

