<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%
Response.Buffer = False
%>
<font size="1" Face="Arial">
<% 
Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecFetch = Server.CreateObject("ADODB.Recordset")
RecFetch.ActiveConnection = MM_Site_STRING
RecFetch.Source = "SELECT * FROM MChargeFNBTrans where Allocated = 'False' Order by FNBDate DeSC"
RecFetch.CursorType = 0
RecFetch.CursorLocation = 2
RecFetch.LockType = 3
RecFetch.Open()
RecFetch_numRows = 0
X = 0
While Not RecFetch.EOF
X = X + 1
WhatDS = RecFetch.Fields.Item("TransDescription").Value
WhatDS = Replace(WhatDS, "ADT CASH DEPOSIT", "")
WhatDS = Replace(WhatDS, "ADT CASH DEPO", "")
FindSpace = InStrRev(WhatDS,"PMG")

'Response.write("<br>" & FindSpace & "<br>")

HowLong = Len(WhatDS)
If FindSpace > 0 Then
WhatDS = Replace(WhatDS, Left(WhatDS,FindSpace - 1), "")
WhatDS = Replace(WhatDS, " " , "")
End If


WhatDS2 = WhatDS
%>
<!--#include file="includes/decode.inc" -->
<%=X%>. <%=RecFetch.Fields.Item("TransDescription").Value%> - <%=RecFetch.Fields.Item("FNBDate").Value%> - '<%=WhatDS2%>'<br>
<%
Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT Top(1) * FROM Tedis Where TediEmpCode = '" & WhatDS & "'"
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0
'Response.write("<br>" & RecZoner.Source)
F = "<br>Can't Find<br>"
While Not RecZoner.EOF
ZC = RecZoner.Fields.Item("TediEmpCode").Value
ZC = Replace(ZC, " ", "")
'Response.Write("<br>'" & ZC & "' = '" & whatDS & "' ?")
If ZC = WhatDS Then
AllocateTo = RecZoner.Fields.Item("TID").Value
FNBID = RecFetch.Fields.Item("FNBID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT Top(1) * FROM MChargeFNBTrans where FNBID = " & FNBID, MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.Update
RecUpdateFNBTable("Allocated") = "True"
RecUpdateFNBTable("TediID") = AllocateTo
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecFNB = Server.CreateObject("ADODB.Recordset")
RecFNB.ActiveConnection = MM_Site_STRING
RecFNB.Source = "SELECT Top(1) FNBID,FNBDate,TransDescription, TransAmount FROM MChargeFNBTrans where FNBID = " & FNBID
RecFNB.CursorType = 0
RecFNB.CursorLocation = 2
RecFNB.LockType = 3
RecFNB.Open()
RecFNB_numRows = 0

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
Set RecUpdateZTTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateZTTable.Open "SELECT Top(1) * FROM TediTransactions order by CID Desc", MM_Site_STRINGWrite, 1, 2
RecUpdateZTTable.AddNew
RecUpdateZTTable("CAmount") = RecFNB.Fields.Item("TransAmount").Value
RecUpdateZTTable("FNBID") = RecFNB.Fields.Item("FNBID").Value
RecUpdateZTTable("CDate") = RecFNB.Fields.Item("FNBDate").Value
RecUpdateZTTable("TediID") = AllocateTo
RecUpdateZTTable("CType") = "2"
RecUpdateZTTable("CComments") = RecFNB.Fields.Item("TransDescription").Value
RecUpdateZTTable("AddedBy") = Session("UNID")

RecUpdateZTTable.Update
RecUpdateZTTable.Close



%>
<%
F = "Found It - " & RecZoner.Fields.Item("TediFirstName").Value & " - " & ZC & " Allocated<br>" & Msg & " - Length=" & Len(Msg) & "<br>"
End If
RecZoner.MoveNext
Wend
response.Write(F)
RecFetch.MoveNext
Wend
' Check to see if week exists in the system

%>


