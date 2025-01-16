<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

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
                <div class="content panel">
                    <div class="row heading"><h1><%=SupervisorLabel%>: <%=RecEdit.Fields.Item("ASEmpCode").Value%></h1>
			</div>
<%If Request.QueryString("ASUpdated") = "True" Then%><div class="alert-box success">Zone Manager <strong><%=Request.QueryString("ASName")%> (<%=Request.QueryString("ASEmpCode")%>)</strong> Updated In The System.</div><%End If%>
		
                        <fieldset>
                            <div class="nine columns">

                                First Name: <label for="agentEmail"><%=RecEdit.Fields.Item("ASFirstName").Value%></label>
                                <br>Last Name: <label for="agentEmail"><%=RecEdit.Fields.Item("ASLastName").Value%></label>
    
                                <br>Email: <label for="agentCell"><%=RecEdit.Fields.Item("ASEmail").Value%></label>
    
                                <br>Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("ASCell").Value%></label>
    

  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%></label>
 

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
                        </fieldset>
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
<%If Request.QueryString("Item") = "1" Then%><!-- #include file="includes/ASPersonalInfo.inc" --><%End If%>
<%If Request.QueryString("Item") = "4" Then%><!-- #include file="includes/UserFiles.inc" --><%End If%>
<%If Request.QueryString("Item") = "6" Then%><!-- #include file="includes/ASAuditTrial.inc" --><%End If%>
<%If Request.QueryString("Item") = "7" Then%><!-- #include file="includes/ASSims.inc" --><%End If%>
<%If Request.QueryString("Item") = "" or Request.QueryString("Item") = "8" Then%><!-- #include file="includes/ASTedis.inc" --><%End If%>
<%If Request.QueryString("Item") = "2" Then
UserType = 2
AlloID = Request.QueryString("ASID")
%><!-- #include file="includes/ComHistory.inc" --><%End If%>
		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

