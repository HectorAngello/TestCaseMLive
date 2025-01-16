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
<div class="five columns">
First Name: <label for="agentEmail"><%=RecEdit.Fields.Item("TediFirstName").Value%></label>
                                <br>Last Name: <label for="agentEmail"><%=RecEdit.Fields.Item("TediLastName").Value%></label>
                                <br>Email: <label for="agentCell"><%=RecEdit.Fields.Item("TediEmail").Value%></label>
                                <br>Primary Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("TediCell").Value%></label>
                                <br>Secondary Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("TediCell2").Value%></label>
                                <br>Tertiary Mobile Number:  <label for="agentCell"><%=RecEdit.Fields.Item("TertiaryMobileNumber").Value%></label>
  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%> - <%=RecEdit.Fields.Item("SubRegionName").Value%></label>
  				<br>Agent Type: <label for="agentEmail"><%=TediType%></label>
  				<br><%=SupervisorLabel%>: <label for="agentEmail"><a href="ASView.asp?ASID=<%=RecEdit.Fields.Item("ASID").Value%>"><%=RecEdit.Fields.Item("ASFirstName").Value%> - <%=RecEdit.Fields.Item("ASLastName").Value%></a></label>
				<br>Airtime Allocation Type: <label for="agentCell"><%=RecEdit.Fields.Item("AirtimeAlloLabel").Value%></label>
				<br>Cluster: <strong><%=RecEdit.Fields.Item("ClusterName").Value%></strong>
				<br>Trading Spot: <strong><%=RecEdit.Fields.Item("TradingSpot").Value%></strong>
<%
DDConsentForm = "No"
If RecEdit.Fields.Item("DDConsentForm").Value = "True" Then
DDConsentForm = "Yes"
End If

DDCrimCheck = "No"
If RecEdit.Fields.Item("DDCrimCheck").Value = "True" Then
DDCrimCheck = "Yes"
End If

DDCrimRecord = "No"
If RecEdit.Fields.Item("DDCrimRecord").Value = "True" Then
DDCrimRecord = "Yes"
End If

DDAMLTrained = "No"
If RecEdit.Fields.Item("DDAMLTrained").Value = "True" Then
DDAMLTrained = "Yes"
End If

DDAMLPassed = "No"
If RecEdit.Fields.Item("DDAMLPassed").Value = "True" Then
DDAMLPassed = "Yes"
End If


DDPhoneAllocated = "No"
If RecEdit.Fields.Item("DDPhoneAllocated").Value = "True" Then
DDPhoneAllocated = "Yes"
End If

DDMSISDNAllocated = "No"
If RecEdit.Fields.Item("DDMSISDNAllocated").Value = "True" Then
DDMSISDNAllocated = "Yes"
End If

DDTDROboarded = "No"
If RecEdit.Fields.Item("DDTDROboarded").Value = "True" Then
DDTDROboarded = "Yes"
End If

DDValidated = "No"
If RecEdit.Fields.Item("DDValidated").Value = "True" Then
DDValidated = "Yes"
End If

DDSkhokhoGSM = "No"
If RecEdit.Fields.Item("DDSkhokhoGSM").Value = "True" Then
DDSkhokhoGSM = "Yes"
End If
DDSkhokhoDedicated = "No"
If RecEdit.Fields.Item("DDSkhokhoDedicated").Value = "True" Then
DDSkhokhoDedicated = "Yes"
End If
%>
<br>
<br><br>Consent Form Submitted: <strong><%=DDConsentForm%></strong>
<br>Crim Check: <strong><%=DDCrimCheck%></strong>
<br>Crim Record: <strong><%=DDCrimRecord%></strong>
<br>AML Trained: <strong><%=DDAMLTrained%></strong>
<br>AML Passed: <strong><%=DDAMLPassed%></strong>
<br>Phone Allocated: <strong><%=DDPhoneAllocated%></strong>
<br>MSISDN Allocated: <strong><%=DDMSISDNAllocated%></strong>
<br>TDR Onboarded: <strong><%=DDTDROboarded%></strong>
<br>Validated: <strong><%=DDValidated%></strong>
<br>Status: <strong><%=RecEdit.Fields.Item("DDStatusLabel").Value%></strong>
<br><br>Made for Skhokho GSM: <strong><%=DDSkhokhoGSM%></strong>
<br>Made for Skhokho Dedicated: <strong><%=DDSkhokhoDedicated%></strong>
</div>
<div class="four columns">
<%
AgentStatus = "Active"
If RecEdit.Fields.Item("TediActive").Value = "False" Then
AgentStatus = "In-Active"
End If
AgentOnWatchList = "No"
If RecEdit.Fields.Item("OnwatchList").Value = "True" Then
AgentOnWatchList = "Yes"
End If

AgentRealTimeComm = "No"
AgentRealTimeComText = ""
If RecEdit.Fields.Item("RealTimeCommOptIn").Value = "True" Then
AgentRealTimeComm = "Yes"
AgentRealTimeComText = "(Agent opted in, but will only start receiving realtime airtime commission from the 1st of next month.)"
End If



If AgentRealTimeComm = "Yes" Then
set RecInCommOpt = Server.CreateObject("ADODB.Recordset")
RecInCommOpt.ActiveConnection = MM_Site_STRING
RecInCommOpt.Source = "SELECT Top(1)* FROM TediRealTimeCommAllocations Where TediID = " & Request.QueryString("TID")
RecInCommOpt.CursorType = 0
RecInCommOpt.CursorLocation = 2
RecInCommOpt.LockType = 3
RecInCommOpt.Open()
RecInCommOpt_numRows = 0
If not RecInCommOpt.EOF and Not RecInCommOpt.BOF Then
AgentRealTimeComText = "(Agent currently receiving real time airtime.)"
End If
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
<br>Agent opted in for real time airtime commission: <label for="agentEmail"><%=AgentRealTimeComm%></label> <font size=1><%=AgentRealTimeComText%></font>
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
<div class="three columns" align="center">
<%
Randomize
Seed = FormatNumber((9999999 * Rnd),0,,,0)
%>
<a href="TediImages/<%=Replace(RecEdit.Fields.Item("TediPic").Value, "-avatar", "")%>" Target="_New"><img src="TediImages/<%=RecEdit.Fields.Item("TediPic").Value%>?Seed=<%=Seed%>" id="avatar2" width="150" Border="0"></a><br>
<%
SystemItem = "244"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%>
                <a href data-reveal-id="avatarModal" class="button">Edit Image</a>
<%
End If
%>
             
</div>
</div>
<!-- Avatar Modal -->
                <div class="reveal-modal" id="avatarModal">


                            <div class="ip-modal-header">
                                
                                <h4 class="ip-modal-title">Change Tedi Image</h4>
<script language="javascript">      
       //Create an iframe and turn on the design mode for it 
       document.write ('<iframe src="TediImage.asp?TediPic=<%=RecEdit.Fields.Item("TediPic").Value%>&TID=<%=Request.QueryString("TID")%>" id="Abstract" width="100%" height="450" frameborder="0" scrolling="auto"></iframe>');             
 </script>
                            </div>
  			    <div>
                                <a class="close-reveal-modal">×</a>
                            </div>
   

                </div>
                <!-- end Modal -->
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

TotalFNBDepositsMM = 0
TotalMchargeAllocationsMM = 0
MChargeBalanceMM = 0

If IsNull(RecEdit.Fields.Item("TediTotalBankedMM").Value) = false then
TotalFNBDepositsMM = RecEdit.Fields.Item("TediTotalBankedMM").Value
End If
If IsNull(RecEdit.Fields.Item("TediTotalAllocatedMM").Value) = false then
TotalMchargeAllocationsMM = RecEdit.Fields.Item("TediTotalAllocatedMM").Value
End If


MChargeBalanceMM = TotalMchargeAllocationsMM - TotalFNBDepositsMM

LastBankedDate = "N/A"
If RecEdit.Fields.Item("LastBankedDate").Value <> "" Then
LastBankedDate = Day(RecEdit.Fields.Item("LastBankedDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("LastBankedDate").Value)) & " " & Year(RecEdit.Fields.Item("LastBankedDate").Value)
End If

LastBankedDateMM = "N/A"
If RecEdit.Fields.Item("LastBankedDateMM").Value <> "" Then
LastBankedDateMM = Day(RecEdit.Fields.Item("LastBankedDateMM").Value) & " " & MonthName(Month(RecEdit.Fields.Item("LastBankedDateMM").Value)) & " " & Year(RecEdit.Fields.Item("LastBankedDateMM").Value)
End If
%>

<%If Request.QueryString("Item") = "" Then%>
<hr>
<h2>Financial Information:</h2>
<div class="row">
<%
If RecEdit.Fields.Item("MChargeTedi").Value = "True" Then
%>
<div class="six columns">
<h3>MCharge:</h3>
Purse Limit: <label for="agentEmail">R <%=RecEdit.Fields.Item("PurseLimit").Value%></label>
<br>M-Charge Balance: <label for="agentEmail">R <%=FormatNumber(MChargeBalance,2)%></label>
<br>Last Banked Date: <label for="agentEmail"><%=LastBankedDate%></label>
<br>Total FNB Deposits: <label for="agentEmail">R <%=FormatNumber(TotalFNBDeposits,2)%></label>
<br>Total M-Charge Allocations: <label for="agentEmail">R <%=FormatNumber(TotalMchargeAllocations,2)%></label>
</div>
<%
End If
If RecEdit.Fields.Item("MobileMoneyTedi").Value = "True" Then
%><div class="six columns">
<h3>Mobile Money:</h3>
Purse Limit: <label for="agentEmail">R <%=RecEdit.Fields.Item("PurseLimitMM").Value%></label>
<br>Mobile Money Balance: <label for="agentEmail">R <%=FormatNumber(MChargeBalanceMM,2)%></label>
<br>Last Banked Date: <label for="agentEmail"><%=LastBankedDateMM%></label>
<br>Total FNB Deposits: <label for="agentEmail">R <%=FormatNumber(TotalFNBDepositsMM,2)%></label>
<br>Total Mobile Money Allocations: <label for="agentEmail">R <%=FormatNumber(TotalMchargeAllocationsMM,2)%></label>
</div>
<%
End If
End If
%>
</div>
<%If Request.QueryString("Item") = "4" Then%><!-- #include file="includes/UserFiles.inc" --><%End If%>
<%If Request.QueryString("Item") = "6" Then%><!-- #include file="includes/TediAuditTrial.inc" --><%End If%>
<%If Request.QueryString("Item") = "1" Then%><!-- #include file="includes/TediPersonalInfo.inc" --><%End If%>
<%If Request.QueryString("Item") = "9" Then%><!-- #include file="includes/TediMChargeHistory.inc" --><%End If%>
<%If Request.QueryString("Item") = "8" Then%><!-- #include file="includes/TediDeductions.inc" --><%End If%>
<%If Request.QueryString("Item") = "7" Then%><!-- #include file="includes/TediReCons.inc" --><%End If%>
<%If Request.QueryString("Item") = "12" Then%><!-- #include file="includes/TediConnections.inc" --><%End If%>
<%If Request.QueryString("Item") = "10" Then%><!-- #include file="includes/TediVends.inc" --><%End If%>
<%If Request.QueryString("Item") = "11" Then%><!-- #include file="includes/TediTransfers.inc" --><%End If%>
<%If Request.QueryString("Item") = "14" Then%><!-- #include file="includes/TediSims.inc" --><%End If%>
<%If Request.QueryString("Item") = "15" Then%><!-- #include file="includes/TediRealTimeCommHistory.inc" --><%End If%>
<%If Request.QueryString("Item") = "16" Then%><!-- #include file="includes/TediSimActivations.inc" --><%End If%>
<%If Request.QueryString("Item") = "17" Then%><!-- #include file="includes/TediMobileMoneyHistory.inc" --><%End If%>
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

