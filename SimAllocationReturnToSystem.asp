<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewASDetail where CompanyID = " & Session("CompanyID") & " and  ASID = " & Request.QueryString("ASID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

UType = 2
UserID = Request.QueryString("ASID")


%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
		<!-- #include file="Includes/EDIsidebar.asp" -->
            </div>
            <div class="nine columns">
<%If Request.QueryString("TediUpdated") = "True" Then%><div class="alert-box success">Agent Updated In The System.</div><%End If%>
                <div class="content panel">

                        <div class="eight columns"><h1><%=SupervisorLabel%>: <%=RecEdit.Fields.Item("ASEmpCode").Value%></h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br>


                                

<div class="row">
<div class="five columns">
                                First Name: <label for="agentEmail"><%=RecEdit.Fields.Item("ASFirstName").Value%></label>
                                <br>Last Name: <label for="agentEmail"><%=RecEdit.Fields.Item("ASLastName").Value%></label>
    
                                <br>Email: <label for="agentCell"><%=RecEdit.Fields.Item("ASEmail").Value%></label>
    
                                <br>Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("ASCell").Value%></label>
    

  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%></label></div>
<div class="four columns">

</div>
<div class="three columns" align="center">
<%
Randomize
Seed = FormatNumber((9999999 * Rnd),0,,,0)
%>
<a href="ASImages/<%=Replace(RecEdit.Fields.Item("ASProfilePic").Value, "-avatar", "")%>" Target="_New"><img src="ASImages/<%=RecEdit.Fields.Item("ASProfilePic").Value%>?Seed=<%=Seed%>" id="avatar2" width="150" Border="0"></a><br>
<%
SystemItem = "245"
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
                                
                                <h4 class="ip-modal-title">Change <%=SupervisorLabel%> Image</h4>
<script language="javascript">      
       //Create an iframe and turn on the design mode for it 
       document.write ('<iframe src="ASImage.asp?ASPic=<%=RecEdit.Fields.Item("ASProfilePic").Value%>&ASID=<%=Request.QueryString("ASID")%>" id="Abstract" width="100%" height="450" frameborder="0" scrolling="auto"></iframe>')
frames.Abstract.document.designMode = "off";               
 </script>
                            </div>
  			    <div>
                                <a class="close-reveal-modal">×</a>
                            </div>
   

                </div>
                <!-- end Modal -->
<%
set RecBulkSims = Server.CreateObject("ADODB.Recordset")
RecBulkSims.ActiveConnection = MM_Site_STRING
RecBulkSims.Source = "SELECT * FROM BulkSimsAS Where BulkID = " & Request.QueryString("BulkID")
'response.write(RecBulkSims.Source)
RecBulkSims.CursorType = 0
RecBulkSims.CursorLocation = 2
RecBulkSims.LockType = 3
RecBulkSims.Open()
RecBulkSims_numRows = 0

set RecSimBreakdown = Server.CreateObject("ADODB.Recordset")
RecSimBreakdown.ActiveConnection = MM_Site_STRING
RecSimBreakdown.Source = "EXECUTE SPMentorSimAllocationBreakdown @BulkID = " & Request.QueryString("BulkID")
'RecSimBreakdown.Source = "SELECT * FROM BulkSimChildrenAS Where BulkID = " & Request.QueryString("BulkID")
'Response.write(RecSimBreakdown.Source)
RecSimBreakdown.CursorType = 0
RecSimBreakdown.CursorLocation = 2
RecSimBreakdown.LockType = 3
RecSimBreakdown.Open()
RecSimBreakdown_numRows = 0
%>


<hr>
			<h2>Return Sims To The System:</h2>
<p>
<strong>Bulk File Creation Date:</strong> <%=Day(RecBulkSims.Fields.Item("BulkDate").Value) & " " & MonthName(Month(RecBulkSims.Fields.Item("BulkDate").Value))  & " " & Year(RecBulkSims.Fields.Item("BulkDate").Value)%>
</p>
<p>ONLY Sims not activated or allocated to an agent can be moved back to the system.</p>
<form name="ReallocateToSystem" Action="SimAllocationReturnToSystem2.asp" Method="Post">
<table>
<thead>
<tr>
	<th>Serial No</th>
	<th>Activation Date</th>
	<th>Agent</th>
	<th><%If Request.Querystring("All") = "" Then%><a href="SimAllocationReturnToSystem.asp?ASID=<%=Request.QueryString("ASID")%>&BulkID=<%=Request.QueryString("BulkID")%>&All=True">All</a><%End If%><%If Request.Querystring("All") = "True" Then%><a href="SimAllocationReturnToSystem.asp?ASID=<%=Request.QueryString("ASID")%>&BulkID=<%=Request.QueryString("BulkID")%>">None</a><%End If%></th>
</tr>
</thead>
<tbody>
<%
SC = 0
While Not RecSimBreakdown.EOF
SC = SC + 1
CanBeReturned = "Yes"
If RecSimBreakdown.Fields.Item("ActivationDate").Value <> "" Then
ActDate = Day(RecSimBreakdown.Fields.Item("ActivationDate").Value) & " " & MonthName(Month(RecSimBreakdown.Fields.Item("ActivationDate").Value))  & " " & Year(RecSimBreakdown.Fields.Item("ActivationDate").Value)
CanBeReturned = "No"
Else
ActDate = "N/A"
End If

AgentLink = "UnAllocated"
If RecSimBreakdown.Fields.Item("AllocatedTo").Value <> 0 Then
AgentLink = "<a href=TediView.asp?TID=" & RecSimBreakdown.Fields.Item("TID").Value & "&Item=14>" & RecSimBreakdown.Fields.Item("TediEmpCode").Value & "</a>"
CanBeReturned = "No"
End If

CheckedMe = ""
If Request.Querystring("All") = "True" Then
CheckedMe = "Checked"
End If
%><tr>
	<td><%=SC%>. <%=RecSimBreakdown.Fields.Item("SerialNo").Value%></td>
	<td><%=ActDate%></td>
	<td><%=AgentLink%></td>
	<td><%If CanBeReturned = "Yes" Then%><input type="checkbox" Name="ChildID<%=(RecSimBreakdown.Fields.Item("ChildID").Value)%>" Value="Yes" <%=CheckedMe%>><%End If%></td>
</tr><%
RecSimBreakdown.MoveNext
Wend
%>
</tbody>
</table>
<input type="Submit" Value="Return Sims To System" class="orange nice button radius">
<input type="Hidden" Name="BulkID" Value="<%=Request.Querystring("BulkID")%>">
<input type="Hidden" Name="ASID" Value="<%=Request.Querystring("ASID")%>">
</form>
                    </div>



<!-- #include file="includes/footer.asp" -->

