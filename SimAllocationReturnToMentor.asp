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
                                <br>Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("TediCell").Value%></label>
  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%> - <%=RecEdit.Fields.Item("SubRegionName").Value%></label>
  				<br>Agent Type: <label for="agentEmail"><%=TediType%></label>
  				<br><%=SupervisorLabel%>: <label for="agentEmail"><a href="ASView.asp?ASID=<%=RecEdit.Fields.Item("ASID").Value%>"><%=RecEdit.Fields.Item("ASFirstName").Value%> - <%=RecEdit.Fields.Item("ASLastName").Value%></a></label>
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
set RecSimBreakdown = Server.CreateObject("ADODB.Recordset")
RecSimBreakdown.ActiveConnection = MM_Site_STRING
RecSimBreakdown.Source = "EXECUTE SPAgentSimAllocationBreakdown @BulkID = " & Request.QueryString("BulkID")
'RecSimBreakdown.Source = "Select * FROM ViewAgentSimAllocationBreakdown where BulkID = " & Request.QueryString("BulkID") & " Order By SerialNo Asc"
'Response.write(RecSimBreakdown.Source)
RecSimBreakdown.CursorType = 0
RecSimBreakdown.CursorLocation = 2
RecSimBreakdown.LockType = 3
RecSimBreakdown.Open()
RecSimBreakdown_numRows = 0
%>


<hr>
			<h2>Return Sims To Mentor:</h2>
<p>
<strong>Bulk File Creation Date:</strong> <%=Day(RecSimBreakdown.Fields.Item("BulkDate").Value) & " " & MonthName(Month(RecSimBreakdown.Fields.Item("BulkDate").Value))  & " " & Year(RecSimBreakdown.Fields.Item("BulkDate").Value)%>
</p>
<p>ONLY Sims not activated can be moved back to the mentor.</p>
<form name="ReallocateToMentor" Action="SimAllocationReturnToMentor2.asp" Method="Post">
<table>
<thead>
<tr>
	<th>Serial No</th>
	<th>Activation Date</th>
	<th>Agent ID No</th>
	<th><%If Request.Querystring("All") = "" Then%><a href="SimAllocationReturnToMentor.asp?TID=<%=Request.QueryString("TID")%>&BulkID=<%=Request.QueryString("BulkID")%>&All=True">All</a><%End If%><%If Request.Querystring("All") = "True" Then%><a href="SimAllocationReturnToMentor.asp?TID=<%=Request.QueryString("TID")%>&BulkID=<%=Request.QueryString("BulkID")%>">None</a><%End If%></th>
</tr>
</thead>
<tbody>
<%
SC = 0
While Not RecSimBreakdown.EOF
SC = SC + 1

If RecSimBreakdown.Fields.Item("ActivationDate").Value <> "" Then
ActDate = Day(RecSimBreakdown.Fields.Item("ActivationDate").Value) & " " & MonthName(Month(RecSimBreakdown.Fields.Item("ActivationDate").Value))  & " " & Year(RecSimBreakdown.Fields.Item("ActivationDate").Value)
Else
ActDate = "N/A"
End If

CheckedMe = ""
If Request.Querystring("All") = "True" Then
CheckedMe = "Checked"
End If
%><tr>
	<td><%=SC%>. <%=RecSimBreakdown.Fields.Item("SerialNo").Value%></td>
	<td><%=ActDate%></td>
	<td><%=RecSimBreakdown.Fields.Item("AgentIDNo").Value%></td>
	<td><%If ActDate = "N/A" Then%><input type="checkbox" Name="ChildID<%=(RecSimBreakdown.Fields.Item("ChildID").Value)%>" Value="Yes" <%=CheckedMe%>><%End If%></td>
</tr><%
RecSimBreakdown.MoveNext
Wend
%>
</tbody>
</table>
<input type="Submit" Value="Return Sims To Mentor" class="orange nice button radius">
<input type="Hidden" Name="BulkID" Value="<%=Request.Querystring("BulkID")%>">
<input type="Hidden" Name="TID" Value="<%=Request.Querystring("TID")%>">
</form>
                    </div>



<!-- #include file="includes/footer.asp" -->

