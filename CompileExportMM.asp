<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
UserSID = Request.Form("SID")
CompanyID = Request.Form("CompanyID")
If UserSID = "" or UserSID = " " Then
UserSID = Session("UNID")
End If

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecInsert = Server.CreateObject ( "ADODB.Recordset" )
		RecInsert.Open "SELECT Top(1) * FROM BulkMChargeMM", MM_Site_STRINGWrite, 1, 2
		RecInsert.AddNew
		RecInsert("BulkDate") = Now()
		RecInsert("UserID") = UserSID
		RecInsert("CompanyID") = CompanyID
		RecInsert("BulkStatus") = "Compiling"
		RecInsert("RID") = Request.Form("RID")
		RecInsert.Update
		RecInsert.Close

set RecNewestBulk = Server.CreateObject("ADODB.Recordset")
RecNewestBulk.ActiveConnection = MM_Site_STRING
RecNewestBulk.Source = "SELECT * FROM BulkMChargeMM Order By BulkID Desc"
RecNewestBulk.CursorType = 0
RecNewestBulk.CursorLocation = 2
RecNewestBulk.LockType = 3
RecNewestBulk.Open()
RecNewestBulk_numRows = 0

NewestBulkID = RecNewestBulk.Fields.Item("BulkID").Value
SFile = "MobileMoney-" & Day(Now) & Month(Now)  & Year(Now) & ".BulkID" & NewestBulkID & ".AdminID" & UserSID

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeMM where BulkID = " & NewestBulkID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("FileName") = SFile
		RecUpdateBulk.Update
		RecUpdateBulk.Close

set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ViewTediDetailWithTotals where TediActive = 'True' and RID = " & Request.Form("RID") & " and MobileMoneyTedi = 'True' Order By TediFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0


While Not RecCurrent.EOF

WZ = "Tedi" & RecCurrent.Fields.Item("TID").Value

If Request.Form(WZ) = "Yes" Then
Response.Write("Agent: " & WZ & "<br>")
Response.flush
		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecInsertTmp = Server.CreateObject ( "ADODB.Recordset" )
		RecInsertTmp.Open "SELECT Top(1) * FROM BulkMChargeTediTempMM", MM_Site_STRINGWrite, 1, 2
		RecInsertTmp.AddNew
		RecInsertTmp("BulkID") = NewestBulkID
		RecInsertTmp("TID") = RecCurrent.Fields.Item("TID").Value
		RecInsertTmp("AirtimeTypeID") = RecCurrent.Fields.Item("AirtimeTypeID").Value
		RecInsertTmp.Update
		RecInsertTmp.Close

End If

RecCurrent.MoveNext
Wend


		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeMM where BulkID = " & NewestBulkID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("BulkStatus") = "Pending"
		RecUpdateBulk.Update
		RecUpdateBulk.Close

'Response.Redirect("CompileExportComplete.asp?BulkID=" & NewestBulkID)
%>
<script type="text/javascript">
<!--
function delayer(){
	
	window.location = "CompileExportCompleteMM.asp?BulkID=<%=NewestBulkID%>"

}
//-->
</script>
<body onLoad="setTimeout('delayer()', 100)">
Done</body>