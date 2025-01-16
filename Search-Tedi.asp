<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<!-- #include file="Includes/MyMainRegions.inc" -->
<!-- #include file="Includes/MySubRegions.inc" -->
<%
KW = Request.QueryString("KW")
AgentStatus = Request.QueryString("AgentStatus")
AgentStatusLabel = "All " & AgentLabel & "'s"
SearchAgentsQry = "SELECT * FROM ViewTediDetail Where (TediFirstName LIKE '%" + Replace(KW, "'", "''") + "%' or TediLastName LIKE '%" + Replace(KW, "'", "''") + "%' or TediCell LIKE '%" + Replace(KW, "'", "''") + "%' or TediCell2 LIKE '%" + Replace(KW, "'", "''") + "%' or TediEmpCode LIKE '%" + Replace(KW, "'", "''") + "%' or TediEmail LIKE '%" + Replace(KW, "'", "''") + "%' or IDNumber LIKE '%" + Replace(KW, "'", "''") + "%')"
SearchAgentsQry = SearchAgentsQry & " and CompanyID = " & Session("CompanyID") & " and SRID in (" & SRRegionList & ")"
If AgentStatus = 1 Then
SearchAgentsQry = SearchAgentsQry & " And TediActive = 'False'"
AgentStatusLabel = "Only Terminated " & AgentLabel & "'s"
End If
If AgentStatus = 2 Then
SearchAgentsQry = SearchAgentsQry & " And TediActive = 'True'"
AgentStatusLabel = "Only Active " & AgentLabel & "'s"
End If
SearchAgentsQry = SearchAgentsQry & " Order by RegionName, TediFirstName Asc"
set RecZoners = Server.CreateObject("ADODB.Recordset")
RecZoners.ActiveConnection = MM_Site_STRING
RecZoners.Source = SearchAgentsQry
RecZoners.CursorType = 0
RecZoners.CursorLocation = 2
RecZoners.LockType = 3
RecZoners.Open()
RecZoners_numRows = 0


%>
<strong><%=AgentStatusLabel%></strong>			
                       
                  <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>EmpCode</th>
                                <th>Region</th>
                                <th>Mobile</th>
                                <th>Status</th>
                            	<th></th>
                            </tr>
                        </thead>
<%
SysClientCounter = 0
While Not RecZoners.EOF

SysClientCounter = SysClientCounter + 1
ZonerStatus = "Active"
If RecZoners.Fields.Item("TediActive").Value = "False" Then
ZonerStatus = "In-Active"
End If
TediType = "Agent"
If RecZoners.Fields.Item("TediParent").Value <> 0 Then
TediType = "Sub-Agent"
End If
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecZoners.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecZoners.Fields.Item("TediLastName").Value%></td>
                            <td><%=RecZoners.Fields.Item("TediEmpCode").Value%></td>
                            <td><%=RecZoners.Fields.Item("RegionName").Value%> - <%=RecZoners.Fields.Item("SubRegionName").Value%></td>
                            <td><%=RecZoners.Fields.Item("TediCell").Value%></td>
                            <td><%=ZonerStatus%></td>
                            <td class="action-td"><a href="TediView.asp?TID=<%=RecZoners.Fields.Item("TID").Value%>" class="view-button"></a></td>
                        </tr>
<%

RecZoners.MoveNext
Wend
%>
                    </table>
                    </fieldset>
                    </div>
<p><font size="1">Fields Searched: First Name, Surname, Cell Number, EmpCode, ID Number and Email address.</font></p>
