<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<!-- #include file="Includes/MyMainRegions.inc" -->
<%
KW = Request.QueryString("KW")
set RecZoners = Server.CreateObject("ADODB.Recordset")
RecZoners.ActiveConnection = MM_Site_STRING
RecZoners.Source = "SELECT * FROM ViewASDetail Where (ASFirstName LIKE '%" + Replace(KW, "'", "''") + "%' or ASLastName LIKE '%" + Replace(KW, "'", "''") + "%' or ASCell LIKE '%" + Replace(KW, "'", "''") + "%' or ASEmpCode LIKE '%" + Replace(KW, "'", "''") + "%' or ASEmail LIKE '%" + Replace(KW, "'", "''") + "%' or IDNumber LIKE '%" + Replace(KW, "'", "''") + "%') and CompanyID = " & Session("CompanyID") & " and RID In (" & SRRegionMainList & ") Order by RegionName, ASFirstName Asc"
RecZoners.CursorType = 0
RecZoners.CursorLocation = 2
RecZoners.LockType = 3
RecZoners.Open()
RecZoners_numRows = 0


%>
		
                       
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>EmpCode</th>
                                <th>Region</th>
                                <th>Cell</th>
                                <th>Status</th>
                            	<th></th>
                            </tr>
                        </thead>
<%
SysClientCounter = 0
While Not RecZoners.EOF

SysClientCounter = SysClientCounter + 1
ZonerStatus = "Active"
If RecZoners.Fields.Item("ASActive").Value = "False" Then
ZonerStatus = "In-Active"
End If
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecZoners.Fields.Item("ASFirstName").Value%>&nbsp;<%=RecZoners.Fields.Item("ASLastName").Value%></td>
                            <td><%=RecZoners.Fields.Item("ASEmpCode").Value%></td>
                            <td><%=RecZoners.Fields.Item("RegionName").Value%></td>
                            <td><%=RecZoners.Fields.Item("ASCell").Value%></td>
                            <td><%=ZonerStatus%></td>
                            <td class="action-td"><a href="ASView.asp?ASID=<%=RecZoners.Fields.Item("ASID").Value%>" class="view-button"></a></td>
                        </tr>
<%

RecZoners.MoveNext
Wend
%>
                    </table>
                    </fieldset>
                    </div>
<p><font size="1">Fields Searched: First Name, Surname, Cell Number, EmpCode, ID Number and Email address.</font></p>