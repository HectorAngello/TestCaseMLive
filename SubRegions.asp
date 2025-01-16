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
RecSubRegion.Source = "SELECT * FROM SubRegions Where RID = " & Request.QueryString("RID") & " and SubRegionActive = 'True' order by SubRegionName Asc"
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
HC = 0
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
                    <table>
                        <thead>
                            <tr>
                                <th>Sub Region Name</th>
                                <th>Sub Region Code</th>
                                <th>HC Target</th>
				<th>Current Count</th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
<%
LC = 0

AllocatedCount = 0
While Not RecSubRegion.EOF
LC = LC + 1
CurrentCount = 0
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and SRID = " & RecSubRegion.Fields.Item("SRID").Value
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
While Not RecEdit.EOF
CurrentCount = CurrentCount + 1
RecEdit.MoveNext
Wend



AllocatedCount = AllocatedCount + 1
%>
                            <tr>
                                <td><%=LC%>. <%=RecSubRegion.Fields.Item("SubRegionName").Value%></td>
                                <td><%=RecSubRegion.Fields.Item("SubRegionCode").Value%></td>
                                <td><%=RecSubRegion.Fields.Item("HeadCountTarget").Value%></td>
				<td><%=CurrentCount%></td>
                                <td class="action-td"><a href="SubRegionEdit.asp?RID=<%=RecRegion.Fields.Item("RID").Value%>&SRID=<%=RecSubRegion.Fields.Item("SRID").Value%>" class="edit-button"></a></td>
                                <td class="action-td"><a href="SubRegionDel.asp?RID=<%=RecRegion.Fields.Item("RID").Value%>&SRID=<%=RecSubRegion.Fields.Item("SRID").Value%>" class="delete-button"></a></td>
                            </tr>
<%

Response.flush
HC = HC + RecSubRegion.Fields.Item("HeadCountTarget").Value
RecSubRegion.MoveNext
Wend
%>
                        </tbody>
                    </table>

<strong>Total Head Count Target: <%=HC%></strong>
<br><br>
<h3>Add A New Sub Region</h3>
<form name="AddRegion" action="SubRegionAdd.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">
   
                                <label for="agencyCode">Region Name *</label>
                                <input type="text" name="RegionName" class="input-text" Required />
    
                               <label for="agencyName">Region Code *</label>
                                <input type="text" name="RegionCode" class="input-text" Required />
                                    <label for="agentEmail">Head Count Target *</label>
                                <input type="text" name="HeadCountTarget" class="input-text" Required />

                                <p>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Add New Sub Region">
                                </p>
                            </div>
                            
                        </fieldset>
<input type="Hidden" Name="RID" Value="<%=Request.QueryString("RID")%>">
                    </form>
<!-- #include file="includes/footer.asp" -->

