<!--#include file="Connections/Site.asp" -->
<!-- #include file="Includes/MySubRegions.inc" -->

<%

SystemItem = "223"
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
<%

set RecWatchListAgents = Server.CreateObject("ADODB.Recordset")
RecWatchListAgents.ActiveConnection = MM_Site_STRING
RecWatchListAgents.Source = "SELECT * FROM ViewTediDetail Where SRID in (" & SRRegionList & ")  and TediActive = 'True' and MChargeTedi = 'True' and CompanyID = " & Session("CompanyID") & " and LastBankedDate IS NOT NULL and (DATEDIFF(day, LastBankedDate, GETDATE()) > 5) Order by RegionName, TediFirstName Asc"
'response.write(RecWatchListAgents.Source)
RecWatchListAgents.CursorType = 0
RecWatchListAgents.CursorLocation = 2
RecWatchListAgents.LockType = 3
RecWatchListAgents.Open()
RecWatchListAgents_numRows = 0

%>
<h3>Agents Not Banked In The Last 5 Days - MCharge</h3>
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>Agent Code</th>
				<th><%=SupervisorLabel%></th>
                                <th>Region</th>
                                <th>Sub Region</th>
                                <th>Last Banked</th>
                            	<th></th>
                            </tr>
                        </thead>
<tbody>
<%
SysClientCounter = 0
While Not RecWatchListAgents.EOF

LastTransAct = DateDiff("d",RecWatchListAgents.Fields.Item("LastBankedDate").Value,Now())
If LastTransAct > 5 Then
LastBankedDay = Day(RecWatchListAgents.Fields.Item("LastBankedDate").Value) & " " & MonthName(Month(RecWatchListAgents.Fields.Item("LastBankedDate").Value),True) & " " & Year(RecWatchListAgents.Fields.Item("LastBankedDate").Value)
SysClientCounter = SysClientCounter + 1
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecWatchListAgents.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecWatchListAgents.Fields.Item("TediLastName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("TediEmpCode").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("ASFirstName").Value & " " & RecWatchListAgents.Fields.Item("ASLastName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("RegionName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("SubRegionName").Value%></td>
                            <td><%=LastBankedDay%></td>
                            <td class="action-td"><a href="TediView.asp?TID=<%=RecWatchListAgents.Fields.Item("TID").Value%>" class="view-button"></a></td>
                        </tr>
<%
End If
Response.flush
RecWatchListAgents.MoveNext
Wend
%>
</tbody>

</table>
<%
End If
%>


<%

SystemItem = "2307"
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
<%

set RecWatchListAgents = Server.CreateObject("ADODB.Recordset")
RecWatchListAgents.ActiveConnection = MM_Site_STRING
RecWatchListAgents.Source = "SELECT * FROM ViewTediDetail Where SRID in (" & SRRegionList & ")  and TediActive = 'True' and MobileMoneyTedi = 'True' and CompanyID = " & Session("CompanyID") & " and LastBankedDate IS NOT NULL and (DATEDIFF(day, LastBankedDate, GETDATE()) > 5) Order by RegionName, TediFirstName Asc"
'response.write(RecWatchListAgents.Source)
RecWatchListAgents.CursorType = 0
RecWatchListAgents.CursorLocation = 2
RecWatchListAgents.LockType = 3
RecWatchListAgents.Open()
RecWatchListAgents_numRows = 0

%>
<h3>Agents Not Banked In The Last 5 Days - Mobile Money</h3>
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>Agent Code</th>
				<th><%=SupervisorLabel%></th>
                                <th>Region</th>
                                <th>Sub Region</th>
                                <th>Last Banked</th>
                            	<th></th>
                            </tr>
                        </thead>
<tbody>
<%
SysClientCounter = 0
While Not RecWatchListAgents.EOF

LastTransAct = DateDiff("d",RecWatchListAgents.Fields.Item("LastBankedDateMM").Value,Now())
If LastTransAct > 5 Then
LastBankedDay = Day(RecWatchListAgents.Fields.Item("LastBankedDateMM").Value) & " " & MonthName(Month(RecWatchListAgents.Fields.Item("LastBankedDateMM").Value),True) & " " & Year(RecWatchListAgents.Fields.Item("LastBankedDateMM").Value)
SysClientCounter = SysClientCounter + 1
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecWatchListAgents.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecWatchListAgents.Fields.Item("TediLastName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("TediEmpCode").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("ASFirstName").Value & " " & RecWatchListAgents.Fields.Item("ASLastName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("RegionName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("SubRegionName").Value%></td>
                            <td><%=LastBankedDay%></td>
                            <td class="action-td"><a href="TediView.asp?TID=<%=RecWatchListAgents.Fields.Item("TID").Value%>" class="view-button"></a></td>
                        </tr>
<%
End If
Response.flush
RecWatchListAgents.MoveNext
Wend
%>
</tbody>

</table>
<%
End If
%>