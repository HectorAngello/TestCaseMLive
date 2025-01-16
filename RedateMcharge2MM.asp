<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
'response.buffer = false
WhichBulkID = Request.QueryString("BulkID")
NewDate = Replace(Request.QueryString("StartDate"), ",","")



Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecBulkID = Server.CreateObject("ADODB.Recordset")
RecBulkID.ActiveConnection = MM_Site_STRING
RecBulkID.Source = "SELECT * FROM BulkMChargeMM Where BulkID = " & WhichBulkID
'Response.Write(RecBulkID.Source)
RecBulkID.CursorType = 0
RecBulkID.CursorLocation = 2
RecBulkID.LockType = 3
RecBulkID.Open()
RecBulkID_numRows = 0

DeleteBulk = "Yes"
If RecBulkID.Fields.Item("BulkStatus").Value = "Compiling" Then
DeleteBulk = "No"
End If
If RecBulkID.Fields.Item("BulkStatus").Value = "Processing" Then
DeleteBulk = "No"
End If

If DeleteBulk = "No" Then
%>
	<script language="JavaScript" type="text/JavaScript">
	<!--
	  alert("Error - This Bulk File Can't Be Updated, File is Busy Being Processed");
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If

' Find Bulk Children
'Response.Write("<br>Looking For children<br>")
Set conMain = Server.CreateObject ( "ADODB.Connection" )

set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM  BulkMChargeChildrenMM Where BulkID = " & WhichBulkID
'Response.Write(RecCurrent.Source)
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
'Response.Write("<br>Looking For children Query ends")
While Not RecCurrent.EOF
' End Find Bulk Children
CID = RecCurrent.Fields.Item("ChildID").Value
ZID = RecCurrent.Fields.Item("TID").Value

MChargeValue = RecCurrent.Fields.Item("MchargeAmount").Value
TransDate = FormatDateTime(RecCurrent.Fields.Item("ChildCreationDate").Value,1)
TransDay = Day(RecCurrent.Fields.Item("ChildCreationDate").Value)
TransMonth = Month(RecCurrent.Fields.Item("ChildCreationDate").Value)
TransYear = Year(RecCurrent.Fields.Item("ChildCreationDate").Value)
TransDescrition = "Bulk Update: " & RecBulkID.Fields.Item("FileName").Value
' Find and delete the transaction in ZonerTransaction

Set conMain = Server.CreateObject ( "ADODB.Connection" )

set RecFindZonerTrans = Server.CreateObject("ADODB.Recordset")
RecFindZonerTrans.ActiveConnection = MM_Site_STRING
RecFindZonerTrans.Source = "SELECT * FROM TediTransactionsMM Where TediID = " & ZID & " and CComments = '" & TransDescrition & "'  and CAmount = '" & MChargeValue & "'"
'Response.Write("<br>" & RecFindZonerTrans.Source)
RecFindZonerTrans.CursorType = 0
RecFindZonerTrans.CursorLocation = 2
RecFindZonerTrans.LockType = 3
RecFindZonerTrans.Open()
RecFindZonerTrans_numRows = 0
While Not RecFindZonerTrans.EOF

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM TediTransactionsMM Where CID = " & RecFindZonerTrans.Fields.Item("CID").Value, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("CDate") = NewDate
		RecUpdateBulk.Update
		RecUpdateBulk.Close
		Response.write("<br>Agent Transaction Child ID " & RecFindZonerTrans.Fields.Item("CID").Value & " Date Updated to " & NewDate & " Trans Value = " & RecFindZonerTrans.Fields.Item("CAmount").Value)

RecFindZonerTrans.MoveNext
Wend


		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeChildrenMM Where ChildID = " & CID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("ChildCreationDate") = NewDate
		RecUpdateBulk.Update
		RecUpdateBulk.Close
		Response.write("<br>Bulk Child ID " & CID & " Date Updated to " & NewDate & " Trans Value: " & MChargeValue)


RecCurrent.MoveNext
Wend

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM  BulkMChargeMM Where BulkID = " & WhichBulkID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("BulkDate") = NewDate
		RecUpdateBulk.Update
		RecUpdateBulk.Close
		Response.write("<br>Bulk File " & WhichBulkID & " Updated to: " & NewDate)

Response.Redirect("Display.asp?AppCat=18&AppSubCatID=1045")
%>