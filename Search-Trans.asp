<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<!-- #include file="Includes/MyMainRegions.inc" -->
<!-- #include file="Includes/MySubRegions.inc" -->
<%
KW = Request.QueryString("KW")
STransType = Request.QueryString("STransType")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
set RecZoners = Server.CreateObject("ADODB.Recordset")
RecZoners.ActiveConnection = MM_Site_STRING
If STransType = "MCharge" Then
TransQRY = "SELECT * FROM viewTediTransactions Where (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
End If
If STransType = "MMoney" Then
TransQRY = "SELECT * FROM viewTediTransactionsMM Where (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
End If
If Isnumeric(KW) = "True" Then
TransQRY = TransQRY & " and CAmount = " & KW 
Else
TransQRY = TransQRY & " and CComments Like '%" & KW & "%'"
End If
TransQRY = TransQRY & " and CompanyID = " & Session("CompanyID") & " and SRID in (" & SRRegionList & ") Order by RegionName, TediFirstName, CDate Asc"

RecZoners.Source = TransQRY
'response.write(RecZoners.Source)
RecZoners.CursorType = 0
RecZoners.CursorLocation = 2
RecZoners.LockType = 3
RecZoners.Open()
RecZoners_numRows = 0


%>
		
                       <p>Start Date: <%=StartDate%>
			<br>End Date: <%=EndDate%></p>
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Name</th>
                                <th>EmpCode</th>
                                <th>Region</th>
                                <th>Value</th>
                                <th>TransAction Date</th>
                            	<th></th>
                            </tr>
                        </thead>
<%
SysClientCounter = 0
While Not RecZoners.EOF

SysClientCounter = SysClientCounter + 1
TransDate = Day(RecZoners.Fields.Item("CDate").Value) & " " & MonthName(Month(RecZoners.Fields.Item("CDate").Value),True) & " " & Year(RecZoners.Fields.Item("CDate").Value)
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecZoners.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecZoners.Fields.Item("TediLastName").Value%></td>
                            <td><%=RecZoners.Fields.Item("TediEmpCode").Value%></td>
                            <td><%=RecZoners.Fields.Item("RegionName").Value%> - <%=RecZoners.Fields.Item("SubRegionName").Value%></td>
                            <td>R <%=RecZoners.Fields.Item("CAmount").Value%></td>
                            <td><%=TransDate%></td>
                            <td class="action-td"><a href="TediView.asp?TID=<%=RecZoners.Fields.Item("TID").Value%>" class="view-button"></a></td>
                        </tr>
<%

RecZoners.MoveNext
Wend
%>
                    </table>
                    </fieldset>
                    </div>
<p><font size="1">Fields Searched: Transaction Date, Transaction Value.</font></p>