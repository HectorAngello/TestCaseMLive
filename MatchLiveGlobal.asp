<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
CID = Request.Form("CID")
M = Request.Form("M")
Y = Request.Form("Y")
TID = Request.Form("TID")


set RecVendList = Server.CreateObject("ADODB.Recordset")
RecVendList.ActiveConnection = MM_Site_STRING
RecVendList.Source = "SELECT * From TediTransactions WHERE CID = " & CID & " and TediID = " & TID & " and TransID = 0"
'Response.write(RecVendList.Source)
RecVendList.CursorType = 0
RecVendList.CursorLocation = 2
RecVendList.LockType = 3
RecVendList.Open()
RecVendList_numRows = 0
MTNValue = RecVendList.Fields.Item("CAmount").Value

set RecUnmatched = Server.CreateObject("ADODB.Recordset")
RecUnmatched.ActiveConnection = MM_Site_STRING
RecUnmatched.Source = "SELECT * From ViewTransfersDetails WHERE CalMonth = " & M & " and CalYear = " & Y & " and TediID = " & TID & " and linkedCID = 0 order by DateCreated"
RecUnmatched.CursorType = 0
RecUnmatched.CursorLocation = 2
RecUnmatched.LockType = 3
RecUnmatched.Open()
RecUnmatched_numRows = 0
GlobalScapeValue = 0
While Not RecUnmatched.EOF
CheckMe = "TransID" & RecUnmatched.Fields.Item("TransID").Value
If Request.Form(CheckMe) = "Yes" Then
GlobalScapeValue = GlobalScapeValue + RecUnmatched.Fields.Item("TransAmount").Value
End If
RecUnmatched.MoveNext
Wend

If GlobalScapeValue <> MTNValue Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! The transaction totals have to match, if need be select multiple transactions.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
Else

set RecUnmatched = Server.CreateObject("ADODB.Recordset")
RecUnmatched.ActiveConnection = MM_Site_STRING
RecUnmatched.Source = "SELECT * From ViewTransfersDetails WHERE CalMonth = " & M & " and CalYear = " & Y & " and TediID = " & TID & " and linkedCID = 0 order by DateCreated"
RecUnmatched.CursorType = 0
RecUnmatched.CursorLocation = 2
RecUnmatched.LockType = 3
RecUnmatched.Open()
RecUnmatched_numRows = 0
GlobalScapeValue = 0
While Not RecUnmatched.EOF
CheckMe = "TransID" & RecUnmatched.Fields.Item("TransID").Value
If Request.Form(CheckMe) = "Yes" Then

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecInsert = Server.CreateObject ( "ADODB.Recordset" )
		RecInsert.Open "SELECT Top(1)* FROM Transfers where TransID = " & RecUnmatched.Fields.Item("TransID").Value, MM_Site_STRINGWrite, 1, 2
		RecInsert.Update
		RecInsert("LinkedCID") = CID
		RecInsert.Update
		RecInsert.Close

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecInsert = Server.CreateObject ( "ADODB.Recordset" )
		RecInsert.Open "SELECT Top(1)* FROM TediTransactions where CID = " & CID, MM_Site_STRINGWrite, 1, 2
		RecInsert.Update
		RecInsert("TransID") = RecUnmatched.Fields.Item("TransID").Value
		RecInsert.Update
		RecInsert.Close

End If
RecUnmatched.MoveNext
Wend
End If
Response.redirect("TediView.asp?TID=" & TID & "&Item=11&M=" & M & "&Y=" & Y)
%>