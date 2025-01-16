<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT * FROM ViewRegionsDetail Where RID = " & Request.QueryString("RID")
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0

set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT * FROM SubRegions Where SRID = " & Request.QueryString("SRID")
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0

CanDelete = "Yes"
set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "SELECT * FROM ViewTediDetail where  SRID = " & Request.QueryString("SRID") & " Order By TediFirstName Asc"
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
CanDelete = "No"
End If

CanView = "No"
SystemItem = "61"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
CanView = "Yes"
End If

CanEdit = "No"
SystemItem = "59"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
CanEdit = "Yes"
End If

CanDel = "No"
SystemItem = "60"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
CanDel = "Yes"
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
                        <div class="eight columns"><h1><%=RecRegion.Fields.Item("RegionName").Value%> Sub Regions</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
                    </div>
                <h3>Confirm Sub Region Deletion</h3>

                        <fieldset>
                            <div class="twelve columns">
   
                                <label for="agencyCode">Region Name</label>
                                <%=RecSubRegion.Fields.Item("SubRegionName").Value%>
    
                               <br><label for="agencyName">Region Code</label>
                               <%=RecSubRegion.Fields.Item("SubRegionCode").Value%>
    

                                <br><label for="agentEmail">Head Count Target</label>
                                <%=RecSubRegion.Fields.Item("HeadCountTarget").Value%>
                             
<%
If CanDelete = "Yes" Then
%><form name="AddRegion" action="SubRegionDel2.asp" method="post"  class="nice">
                                <p>
                                    <input type="Submit" class="orange nice button radius" value="Confirm Sub Region Deletion">
                                </p>
<input type="Hidden" Name="RID" Value="<%=Request.Querystring("RID")%>">
<input type="Hidden" Name="SRID" Value="<%=Request.Querystring("SRID")%>">
</form>
<%Else%>

<br><br><h4>Unable to delete this Sub Region</h4>
<form name="bulkmove" Action="BulkSubRegionAllocate.asp" Method="Post">
<p>The following Agents are currently allocated to this region, please change their sub region and try again.</p>
                    <table>
                        <thead>
                            <tr>
                                <th>Name</th>
                                <th>Agent Code</th>
                                <th><%=SupervisorLabel%></th>
                                <th>Status</th>
                                <th><%If Request.Querystring("All") = "" Then%><a href="SubRegionDel.asp?RID=<%=Request.QueryString("RID")%>&SRID=<%=Request.QueryString("SRID")%>&All=True">All</a><%End If%><%If Request.Querystring("All") = "True" Then%><a href="SubRegionDel.asp?RID=<%=Request.QueryString("RID")%>&SRID=<%=Request.QueryString("SRID")%>">None</a><%End If%></th>
                                <th></th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                       <tbody>
<%
ZC = 0
While Not RecCheck.EOF
ZC = ZC + 1
TType = "Agent"
If RecCheck.Fields.Item("TediParent").Value <> 0 Then
TType = "Sub-Agent"
End If
AgentStatus = "Active"
If RecCheck.Fields.Item("TediActive").Value = "False" Then
AgentStatus = "In-Active"
End If

CheckedMe = ""
If Request.Querystring("All") = "True" Then
CheckedMe = "Checked"
End If
%>
                            <tr>
                                <td><%=ZC%>. <%=(RecCheck.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecCheck.Fields.Item("TediLastName").Value)%></td>
                                <td><%=(RecCheck.Fields.Item("TediEmpCode").Value)%></td>
                                <td><%=(RecCheck.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCheck.Fields.Item("ASLastName").Value)%>&nbsp;(<%=(RecCheck.Fields.Item("ASEmpCode").Value)%>)</td>
                                <td><%=AgentStatus%></td>
				<td><input type="checkbox" Name="Tedi<%=(RecCheck.Fields.Item("TID").Value)%>" Value="Yes" <%=CheckedMe%>></td>
<%
If CanView = "Yes" then
%>
                                <td class="action-td"><a href="TediView.asp?TID=<%=(RecCheck.Fields.Item("TID").Value)%>" class="view-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanEdit = "Yes" then
%>
                                <td class="action-td"><a href="TediEdit.asp?TID=<%=(RecCheck.Fields.Item("TID").Value)%>" class="edit-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanDel = "Yes" then
%>
                                <td class="action-td"><a href="TediDel.asp?TID=<%=(RecCheck.Fields.Item("TID").Value)%>" class="delete-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>


                            </tr>
<%
response.flush
RecCheck.MoveNext
Wend


%>
                        </tbody>

</table>
Select New Sub Region:	
<select Name="NewSRID">
<%

set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT Distinct SRID, SubregionName FROM ViewUserRegion where SubRegionActive = 'True' and RID = " & Request.Querystring("RID") & " Order By SubRegionName Asc"
'Response.write(RecRegion.Source)
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0
While Not RecRegion.EOF


Selected = ""
If RecRegion.Fields.Item("SRID").Value = Int(Request.Querystring("SRID")) Then
Selected = "Selected"
End If
%>
<option value="<%=RecRegion.Fields.Item("SRID").Value%>" <%=Selected%>><%=RecRegion.Fields.Item("SubRegionName").Value%></option>
<%
RecRegion.Movenext
Wend
%>
</select>
<input type="Hidden" Name="RID" Value="<%=Request.Querystring("RID")%>">
<input type="Hidden" Name="SRID" Value="<%=Request.Querystring("SRID")%>">
<input type="Hidden" Name="UNID" Value="<%=Session("UNID")%>">
<input type="Submit" Value="Reallocate Agents" class="orange nice button radius">
</form>
<%End If%>
                            </div>
                            
                        </fieldset>
                    </form>
<!-- #include file="includes/footer.asp" -->

