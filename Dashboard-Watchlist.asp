<!--#include file="Connections/Site.asp" -->
<%
SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID")

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

set RecWatchListAgents = Server.CreateObject("ADODB.Recordset")
RecWatchListAgents.ActiveConnection = MM_Site_STRING
RecWatchListAgents.Source = "SELECT * FROM ViewTediDetail Where SRID in (" & SRRegionList & ") and OnWatchList = 'True' and TediActive = 'True' and CompanyID = " & Session("CompanyID") & " Order by RegionName, TediFirstName Asc"
RecWatchListAgents.CursorType = 0
RecWatchListAgents.CursorLocation = 2
RecWatchListAgents.LockType = 3
RecWatchListAgents.Open()
RecWatchListAgents_numRows = 0

%>
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>Agent Code</th>
                                <th>Region</th>
                                <th>Sub Region</th>
                                <th>Mobile</th>
                            	<th></th>
                            </tr>
                        </thead>
<tbody>
<%
SysClientCounter = 0
While Not RecWatchListAgents.EOF
SysClientCounter = SysClientCounter + 1
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecWatchListAgents.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecWatchListAgents.Fields.Item("TediLastName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("TediEmpCode").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("RegionName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("SubRegionName").Value%></td>
                            <td><%=RecWatchListAgents.Fields.Item("TediCell").Value%></td>
                            <td class="action-td"><a href="TediView.asp?TID=<%=RecWatchListAgents.Fields.Item("TID").Value%>" class="view-button"></a></td>
                        </tr>
<%
Response.flush
RecWatchListAgents.MoveNext
Wend
%>
</tbody>

</table>
