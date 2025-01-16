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

SystemItem = "70"
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

SystemItem = "63"
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
SystemItem = "64"
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
<h1><%=(RecMainRegions.Fields.Item("RegionName").Value)%>&nbsp;<%=SupervisorLabel%>s</h1>


                   
                    <table>
                        <thead>
                            <tr>
                                <th><%=SupervisorLabel%></th>
                                <th><%=SupervisorLabel%> Code</th>
                                <th>Last Mobi Login</th>
                                <th>Agents</th>
                                <th></th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
<%
OrderBy = "ASFirstName"
If Request.QueryString("OD") <> "" Then
OrderBy = Request.QueryString("OD")
End If
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RecMainRegions.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
%>
                        <tbody>
<%
ZC = 0
TotalHeadCount = 0
TotalActualAgents = 0
While Not RecCurrent.EOF
ZC = ZC + 1
ZZZ = 0

set RecZonerCount = Server.CreateObject("ADODB.Recordset")
RecZonerCount.ActiveConnection = MM_Site_STRING
RecZonerCount.Source = "SELECT * FROM Tedis where TediActive = 'True' and ASID = " & RecCurrent.Fields.Item("ASID").Value
RecZonerCount.CursorType = 0
RecZonerCount.CursorLocation = 2
RecZonerCount.LockType = 3
RecZonerCount.Open()
RecZonerCount_numRows = 0
While Not RecZonerCount.EOF
ZZZ = ZZZ + 1
RecZonerCount.MoveNext
Wend

LastLogin = "N/A"

set RecLastLogIn = Server.CreateObject("ADODB.Recordset")
RecLastLogIn.ActiveConnection = MM_Site_STRING
RecLastLogIn.Source = "SELECT * FROM ChangeLogMobi Where Left(Changes,10) = 'Supervisor' and ChangeBy = " & RecCurrent.Fields.Item("ASID").Value & " Order By ID Desc"
'Response.Write(RecLastLogIn.Source)
RecLastLogIn.CursorType = 0
RecLastLogIn.CursorLocation = 2
RecLastLogIn.LockType = 3
RecLastLogIn.Open()
RecLastLogIn_numRows = 0
If Not RecLastLogIn.EOF and Not RecLastLogIn.BOF Then
LastLogin = Day(RecLastLogIn.Fields.Item("ChangeDate").Value) & " " & MonthName(Month(RecLastLogIn.Fields.Item("ChangeDate").Value),True) & " " & Year(RecLastLogIn.Fields.Item("ChangeDate").Value)
End If
%>
                            <tr>
                                <td><%=ZC%>. <%=(RecCurrent.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("ASLastName").Value)%></td>
                                <td><%=(RecCurrent.Fields.Item("ASEmpCode").Value)%></td>
                                <td><%=LastLogin%></td>
                                <td><%=ZZZ%></td>
<%
TotalHeadCount = TotalHeadCount + RecCurrent.Fields.Item("HeadCountTarget").Value
TotalActualAgents = TotalActualAgents + ZZZ
If CanView = "Yes" then
%>
                                <td class="action-td"><a href="ASView.asp?ASID=<%=(RecCurrent.Fields.Item("ASID").Value)%>" class="view-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanEdit = "Yes" then
%>
                                <td class="action-td"><a href="ASEdit.asp?ASID=<%=(RecCurrent.Fields.Item("ASID").Value)%>" class="edit-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>

<%
If CanDel = "Yes" then
%>
                                <td class="action-td"><a href="ASDel.asp?ASID=<%=(RecCurrent.Fields.Item("ASID").Value)%>" class="delete-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>


                            </tr>
<%
Response.flush
RecCurrent.MoveNext
Wend
%>
                        </tbody>
                    </table>
Total Head Count Target: <%=TotalHeadCount%>
<br>Total Agents: <%=TotalActualAgents%>
<hr>
<%
Response.flush
RecMainRegions.MoveNext
Wend
%>
 </div>

<!-- #include file="includes/footer.asp" -->

