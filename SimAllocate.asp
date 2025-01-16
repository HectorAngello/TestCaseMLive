<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetailWithTotals where CompanyID = " & Session("CompanyID") & " and TID = " & Request.QueryString("TID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

UType = 1
UserID = Request.QueryString("TID")
ASID = RecEdit.Fields.Item("ASID").Value
TediType = Session("AgentLabel")
If RecEdit.Fields.Item("TediParent").Value <> 0 Then


set RecParent = Server.CreateObject("ADODB.Recordset")
RecParent.ActiveConnection = MM_Site_STRING
RecParent.Source = "SELECT * FROM ViewTediDetailWithTotals where  TID = " & RecEdit.Fields.Item("TediParent").Value
RecParent.CursorType = 0
RecParent.CursorLocation = 2
RecParent.LockType = 3
RecParent.Open()
RecParent_numRows = 0
TediType = "Sub-" & Session("AgentLabel") & " - <a href='TediView.asp?TID=" & RecEdit.Fields.Item("TediParent").Value & "'>" & RecParent.Fields.Item("TediFirstName").Value & " " & RecParent.Fields.Item("TediLastName").Value & "</a>"
End If
%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
		<!-- #include file="Includes/Tedisidebar.asp" -->
            </div>
            <div class="nine columns">
<%If Request.QueryString("TediUpdated") = "True" Then%><div class="alert-box success">Agent Updated In The System.</div><%End If%>
                <div class="content panel">

                        <div class="eight columns"><h1>Agent: <%=RecEdit.Fields.Item("TediEmpCode").Value%></h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br>


                                

<div class="row">
<div class="six columns">
        <span class="spacer-120">First Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("TediFirstName").Value%></label><br>
        <span class="spacer-120">Last Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("TediLastName").Value%></label><br>
        <span class="spacer-120">Email:</span> <label for="agentCell"><%=RecEdit.Fields.Item("TediEmail").Value%></label><br>
        <span class="spacer-120">Mobile:</span> <label for="agentCell"><%=RecEdit.Fields.Item("TediCell").Value%></label><br>
        <span class="spacer-120">Region:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%> - <%=RecEdit.Fields.Item("SubRegionName").Value%></label><br>

</div>
<div class="six columns">
<%
AgentStatus = "Active"
If RecEdit.Fields.Item("TediActive").Value = "False" Then
AgentStatus = "In-Active"
End If
AgentOnWatchList = "No"
If RecEdit.Fields.Item("OnwatchList").Value = "True" Then
AgentOnWatchList = "Yes"
End If

AgentMChargeExclude = "No"
If RecEdit.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then
AgentMChargeExclude = "Yes"
End If
%>
Agent Status: <label for="agentEmail"><%=AgentStatus%></label>
<%If AgentStatus = "Active" then%>
<br>Agent on watchlist: <label for="agentEmail"><%=AgentOnWatchList%></label>
<%End If%>
<br>Agent excluded from Airtime file generation: <label for="agentEmail"><%=AgentMChargeExclude%></label>

</div>
</div>

<%


TotalFNBDeposits = 0
TotalMchargeAllocations = 0
MChargeBalance = 0

If IsNull(RecEdit.Fields.Item("TediTotalBanked").Value) = false then
TotalFNBDeposits = RecEdit.Fields.Item("TediTotalBanked").Value
End If
If IsNull(RecEdit.Fields.Item("TediTotalAllocated").Value) = false then
TotalMchargeAllocations = RecEdit.Fields.Item("TediTotalAllocated").Value
End If
MChargeBalance = TotalMchargeAllocations - TotalFNBDeposits



LastBankedDate = "N/A"
If RecEdit.Fields.Item("LastBankedDate").Value <> "" Then
LastBankedDate = Day(RecEdit.Fields.Item("LastBankedDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("LastBankedDate").Value)) & " " & Year(RecEdit.Fields.Item("LastBankedDate").Value)
End If
%>


<hr>
			<h2>Capture Sims Against <%=Session("AgentLabel")%> Profile:</h2>
<form action="SimAllocate2.asp" method="get">

<table border="0" cellspacing="2" cellpadding="2">
<tr>
<td Class="quote">Brick / Box Code 1</td><td><input type="text" Name="BrickCode1"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 2</td><td><input type="text" Name="BrickCode2"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 3</td><td><input type="text" Name="BrickCode3"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 4</td><td><input type="text" Name="BrickCode4"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 5</td><td><input type="text" Name="BrickCode5"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 6</td><td><input type="text" Name="BrickCode6"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 7</td><td><input type="text" Name="BrickCode7"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 8</td><td><input type="text" Name="BrickCode8"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 9</td><td><input type="text" Name="BrickCode9"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 10</td><td><input type="text" Name="BrickCode10"></td>
</tr>
<tr>
            <td colspan="2" align="center"><label>
              <input name="button2" type="submit" class="orange nice button radius" id="button2" value="Capture">
            </label></td>
          </tr>
  </table>
<input type="Hidden" Name="TID" Value="<%=Request.QueryString("TID")%>"></form>
<input type="Hidden" Name="ASID" Value="<%=ASID%>">
</form>
                    </div>
<!-- #include file="includes/footer.asp" -->

