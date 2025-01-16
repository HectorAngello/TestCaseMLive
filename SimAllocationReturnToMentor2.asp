<!-- #include file="Connections/Site.asp" -->
<%
TID = Request.Form("TID")
BulkID = Request.Form("BulkID")

set RecSimBreakdown = Server.CreateObject("ADODB.Recordset")
RecSimBreakdown.ActiveConnection = MM_Site_STRING
RecSimBreakdown.Source = "EXECUTE SPAgentSimAllocationBreakdown @BulkID = " & BulkID
RecSimBreakdown.CursorType = 0
RecSimBreakdown.CursorLocation = 2
RecSimBreakdown.LockType = 3
RecSimBreakdown.Open()
RecSimBreakdown_numRows = 0

While Not RecSimBreakdown.EOF
CheckMe = "ChildID" & RecSimBreakdown.Fields.Item("ChildID").Value
If Request.Form(CheckMe) = "Yes" Then
Response.write("<br>" & RecSimBreakdown.Fields.Item("ChildID").Value)
SerialNo = RecSimBreakdown.Fields.Item("SerialNo").Value

' Delete Record From BulkSimChildren using ChildID
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM BulkSimChildren where ChildID = " & RecSimBreakdown.Fields.Item("ChildID").Value
rstSecond.Open
set rstSecond = nothing

' Update Sims Table and make sim available for allocation

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Sims where SerialNo = '" & SerialNo & "'", MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("Allocated") = "False"
RecADDSMS("AllocatedTo") = "0"
RecADDSMS("AllocatedDate") = NULL
RecADDSMS.Update
RecADDSMS.Close

End If
RecSimBreakdown.MoveNext
Wend

Response.redirect("ASView.asp?ASID=" & ASID & "&Item=7")
%>