<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%
KW = Request.QueryString("KW")
KW = Replace(KW, "'", "''")
set RecSims = Server.CreateObject("ADODB.Recordset")
RecSims.ActiveConnection = MM_Site_STRING
RecSims.Source = "SELECT * FROM ViewSimsAllocationDetails Where (SerialNo = '" & KW & "' or BoxNumber = '" & KW & "' or BrickNumber = '" & KW & "' or KitNo = '" & KW & "') Order by Kitno, SerialNo, BoxNumber, BrickNumber Asc"
'response.write(RecSims.Source)
RecSims.CursorType = 0
RecSims.CursorLocation = 2
RecSims.LockType = 3
RecSims.Open()
RecSims_numRows = 0


%>
		
                       
                    <table>
                        <thead>
                            <tr style="width: 100% !important">
                                <th>Kit No</th>
                                <th>Serial No</th>
                                <th>Brick No</th>
                                <th>Box No</th>
                                <th>Agent</th>
                                <th>Agent Bulk</th>
                                <th>Mentor</th>
                                <th>Mentor Bulk</th>
                            </tr>
                        </thead>
<%
SysClientCounter = 0
While Not RecSims.EOF

SysClientCounter = SysClientCounter + 1
AgentLink = "UnAllocated"
AgentBulkFile = "N/A"
If RecSims.Fields.Item("TID").Value <> "" Then
AgentLink = "<a href=TediView.asp?TID=" & RecSims.Fields.Item("TID").Value & "&Item=14>" & RecSims.Fields.Item("TediEmpCode").Value & "</a>"

set RecFindBulkFile = Server.CreateObject("ADODB.Recordset")
RecFindBulkFile.ActiveConnection = MM_Site_STRING
RecFindBulkFile.Source = "SELECT Top(1)* FROM BulkSimChildren Where (SerialNo = '" & RecSims.Fields.Item("SerialNo").Value & "')"
RecFindBulkFile.CursorType = 0
RecFindBulkFile.CursorLocation = 2
RecFindBulkFile.LockType = 3
RecFindBulkFile.Open()
RecFindBulkFile_numRows = 0
If Not RecFindBulkFile.EOF and Not RecFindBulkFile.BOF Then
AgentBulkFile = "BULK" & RecFindBulkFile.Fields.Item("BulkID").Value
End If
End If

MentorLink = "UnAllocated"
MentorBulkFile = "N/A"
If RecSims.Fields.Item("ASID").Value <> "" Then
MentorLink = "<a href=ASView.asp?ASID=" & RecSims.Fields.Item("ASID").Value & "&Item=7>" & RecSims.Fields.Item("ASEmpCode").Value & "</a>"

set RecFindBulkFile2 = Server.CreateObject("ADODB.Recordset")
RecFindBulkFile2.ActiveConnection = MM_Site_STRING
RecFindBulkFile2.Source = "SELECT Top(1)* FROM BulkSimChildrenAS Where (SerialNo = '" & RecSims.Fields.Item("SerialNo").Value & "')"
RecFindBulkFile2.CursorType = 0
RecFindBulkFile2.CursorLocation = 2
RecFindBulkFile2.LockType = 3
RecFindBulkFile2.Open()
RecFindBulkFile2_numRows = 0
If Not RecFindBulkFile2.EOF and Not RecFindBulkFile2.BOF Then
MentorBulkFile = "BULK" & RecFindBulkFile2.Fields.Item("BulkID").Value
End If

End If
%>
                        <tr>
                            <td><%=SysClientCounter%>. <%=RecSims.Fields.Item("KitNo").Value%></td>
                            <td><%=RecSims.Fields.Item("SerialNo").Value%></td>
                            <td><%=RecSims.Fields.Item("BrickNumber").Value%></td>
                            <td><%=RecSims.Fields.Item("BoxNumber").Value%></td>
                            <td><%=AgentLink%></td>
                            <td><%=AgentBulkFile%></td>
                            <td><%=MentorLink%></td>
                            <td><%=MentorBulkFile%></td>
                        </tr>
<%
Response.flush
RecSims.MoveNext
Wend
%>
                    </table>
                    </fieldset>
                    </div>
<p><font size="1">Fields Searched: Sim Kit, Sim Serial, Sim Brick and Sim Box.</font></p>