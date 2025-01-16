<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")
NewASID = Request.Form("NewASID")
UNID = Request.Form("UNID")
set RecTediCount = Server.CreateObject("ADODB.Recordset")
RecTediCount.ActiveConnection = MM_Site_STRING
RecTediCount.Source = "SELECT * FROM Tedis where  ASID = " & ASID & " Order by TediFirstName Asc"
RecTediCount.CursorType = 0
RecTediCount.CursorLocation = 2
RecTediCount.LockType = 3
RecTediCount.Open()
RecTediCount_numRows = 0
While Not RecTediCount.EOF
UpdateMe = "Tedi" & RecTediCount.Fields.Item("TID").Value
If Request.Form(UpdateMe) = "Yes" Then
Response.write("<br>" & RecTediCount.Fields.Item("TID").Value)
TID = RecTediCount.Fields.Item("TID").Value
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("ASID") = NewASID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

TediUpdateType = "Agent " & SupervisorLabel & " Updated"
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%

End If
RecTediCount.MoveNext
Wend

Response.redirect("ASDel.asp?ASID=" & ASID)
%>