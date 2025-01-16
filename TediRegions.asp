<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecMainRegions = Server.CreateObject("ADODB.Recordset")
RecMainRegions.ActiveConnection = MM_Site_STRING
If Request.querystring("RID") <> "0" Then
RecMainRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and RID = " & Request.querystring("RID")
Else
RecMainRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
End If
'Response.write(RecMainRegions.Source)
RecMainRegions.CursorType = 0
RecMainRegions.CursorLocation = 2
RecMainRegions.LockType = 3
RecMainRegions.Open()
RecMainRegions_numRows = 0

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

	
<%While Not RecMainRegions.EOF%>
<h1><%=(RecMainRegions.Fields.Item("RegionName").Value)%> Agents</h1>
                    <table>
                        <thead>
                            <tr>
                                <th><a href="TediRegions.asp?RID=<%=Request.QueryString("RID")%>&OD=TediFirstName">Name</a></th>
                                <th><a href="TediRegions.asp?RID=<%=Request.QueryString("RID")%>&OD=TediEmpCode">Agent Code</a></th>
                                <th><a href="TediRegions.asp?RID=<%=Request.QueryString("RID")%>&OD=SubRegionName">Sub Region</a></th>
                                <th>Agent Type</th>
                                <th><a href="TediRegions.asp?RID=<%=Request.QueryString("RID")%>&OD=ASFirstName"><%=SupervisorLabel%></a></th>
                                <th></th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
<%
ZC = 0
OB = "SubRegionName"
SubRegionQry = ""
SRRegionList = ""
'set RecSubRegion = Server.CreateObject("ADODB.Recordset")
'RecSubRegion.ActiveConnection = MM_Site_STRING
'RecSubRegion.Source = "SELECT * FROM ViewUserSubRegions where RID = " & RecMainRegions.Fields.Item("RID").Value & " and SubRegionActive = 'True' and UserID = " & Session("UNID") & " Order By " & OB & " Asc"
'RecSubRegion.CursorType = 0
'RecSubRegion.CursorLocation = 2
'RecSubRegion.LockType = 3
'RecSubRegion.Open()
'RecSubRegion_numRows = 0
'While Not RecSubRegion.EOF

SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID") & " and RID = " & RecMainRegions.Fields.Item("RID").Value

set RecWatchlistRegions = Server.CreateObject("ADODB.Recordset")
RecWatchlistRegions.ActiveConnection = MM_Site_STRING
RecWatchlistRegions.Source = SubRegionQry
RecWatchlistRegions.CursorType = 0
RecWatchlistRegions.CursorLocation = 2
RecWatchlistRegions.LockType = 3
RecWatchlistRegions.Open()
RecWatchlistRegions_numRows = 0
While Not RecWatchlistRegions.EOF
SRRegionList = SRRegionList & RecWatchlistRegions.Fields.Item("SRID").Value & ","
RecWatchlistRegions.MoveNext
Wend
TempLenSRRegionList = Len(SRRegionList)
SRRegionList = Left(SRRegionList,TempLenSRRegionList - 1)
%>
<%

OrderByval = "TediFirstName"
If Request.QueryString("OD") <> "" Then
OrderByval = Request.QueryString("OD")
End If
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and SRID IN  (" & SRRegionList & ") Order By " & OrderByval & " Asc"
'Response.write(RecCurrent.Source)
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
If Not RecCurrent.EOF and Not RecCurrent.Bof Then

While Not RecCurrent.EOF
ZC = ZC + 1
TType = "Agent"
If RecCurrent.Fields.Item("TediParent").Value <> 0 Then
TType = "Sub-Agent"
End If
%>
                            <tr>
                                <td><%=ZC%>. <%=(RecCurrent.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("TediLastName").Value)%></td>
                                <td><%=(RecCurrent.Fields.Item("TediEmpCode").Value)%></td>
                                <td><%=(RecCurrent.Fields.Item("SubRegionName").Value)%></td>
                                <td><%=(TType)%></td>
                                <td><%=(RecCurrent.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("ASLastName").Value)%></td>
<%
If CanView = "Yes" then
%>
                                <td class="action-td"><a href="TediView.asp?TID=<%=(RecCurrent.Fields.Item("TID").Value)%>" class="view-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanEdit = "Yes" then
%>
                                <td class="action-td"><a href="TediEdit.asp?TID=<%=(RecCurrent.Fields.Item("TID").Value)%>" class="edit-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanDel = "Yes" then
%>
                                <td class="action-td"><a href="TediDel.asp?TID=<%=(RecCurrent.Fields.Item("TID").Value)%>" class="delete-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>


                            </tr>
<%
response.flush
RecCurrent.MoveNext
Wend


%>


<%
End If
Response.flush
'RecSubRegion.MoveNext
'Wend
%>
                        </tbody>
                    </table>
<%
Response.flush
RecMainRegions.MoveNext
Wend
%>

 </div>

<!-- #include file="includes/footer.asp" -->

