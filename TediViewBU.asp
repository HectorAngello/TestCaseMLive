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

TediType = "Agent"
If RecEdit.Fields.Item("TediParent").Value <> 0 Then


set RecParent = Server.CreateObject("ADODB.Recordset")
RecParent.ActiveConnection = MM_Site_STRING
RecParent.Source = "SELECT * FROM ViewTediDetailWithTotals where  TID = " & RecEdit.Fields.Item("TediParent").Value
RecParent.CursorType = 0
RecParent.CursorLocation = 2
RecParent.LockType = 3
RecParent.Open()
RecParent_numRows = 0
TediType = "Sub-Agent - <a href='TediView.asp?TID=" & RecEdit.Fields.Item("TediParent").Value & "'>" & RecParent.Fields.Item("TediFirstName").Value & " " & RecParent.Fields.Item("TediLastName").Value & "</a>"
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
First Name: <label for="agentEmail"><%=RecEdit.Fields.Item("TediFirstName").Value%></label>
                                <br>Last Name: <label for="agentEmail"><%=RecEdit.Fields.Item("TediLastName").Value%></label>
                                <br>Email: <label for="agentCell"><%=RecEdit.Fields.Item("TediEmail").Value%></label>
                                <br>Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("TediCell").Value%></label>
  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%> - <%=RecEdit.Fields.Item("SubRegionName").Value%></label>
  				<br>Agent Type: <label for="agentEmail"><%=TediType%></label>
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
<%If AgentStatus = "Active" then%>
<%
SystemItem = "222"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%><br><br><h4>Watch List Management</h4>
Update Status: 
<form name="WatchListUpdate">
<select name="menu2" onChange="MM_jumpMenu2('parent',this,0)">


                <option value="TediWatchListUpdate.asp?TID=<%=Request.QueryString("TID")%>&Watchlist=False" <%If RecEdit.Fields.Item("OnwatchList").Value = "False" Then%>Selected<%End If%>>Remove From Watchlist</option>
                <option value="TediWatchListUpdate.asp?TID=<%=Request.QueryString("TID")%>&Watchlist=True" <%If RecEdit.Fields.Item("OnwatchList").Value = "True" Then%>Selected<%End If%>>Add To Watchlist</option>
              </select>
            
        </form>    
<%
End If
End If
%>
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

<%If Request.QueryString("Item") = "" Then%>
<hr>
			<h2>Financial Information:</h2>
Purse Limit: <label for="agentEmail">R <%=RecEdit.Fields.Item("PurseLimit").Value%></label>
<br>M-Charge Balance: <label for="agentEmail">R <%=FormatNumber(MChargeBalance,2)%></label>
<br>Last Banked Date: <label for="agentEmail"><%=LastBankedDate%></label>
<br>Total FNB Deposits: <label for="agentEmail">R <%=FormatNumber(TotalFNBDeposits,2)%></label>
<br>Total M-Charge Allocations: <label for="agentEmail">R <%=FormatNumber(TotalMchargeAllocations,2)%></label>
<%End If%>

<%If Request.QueryString("Item") = "4" Then%><!-- #include file="includes/UserFiles.inc" --><%End If%>
<%If Request.QueryString("Item") = "6" Then%><!-- #include file="includes/TediAuditTrial.inc" --><%End If%>
<%If Request.QueryString("Item") = "1" Then%><!-- #include file="includes/TediPersonalInfo.inc" --><%End If%>
<%If Request.QueryString("Item") = "9" Then%><!-- #include file="includes/TediMChargeHistory.inc" --><%End If%>
<%If Request.QueryString("Item") = "8" Then%><!-- #include file="includes/TediDeductions.inc" --><%End If%>
<%If Request.QueryString("Item") = "7" Then%><!-- #include file="includes/TediReCons.inc" --><%End If%>
<%If Request.QueryString("Item") = "2" Then
UserType = 1
AlloID = Request.QueryString("TID")
%><!-- #include file="includes/ComHistory.inc" --><%End If%>
<%
If RecEdit.Fields.Item("TediParent").Value <> 0 Then
If Request.QueryString("Item") = "7" Then%><!-- #include file="includes/TediUpgrade.inc" --><%
End If
End If%>
                    </div>
<!-- #include file="includes/footer.asp" -->

