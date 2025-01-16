<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%


WhichBulkID = Request.QueryString("BulkID")

set RecBulkID = Server.CreateObject("ADODB.Recordset")
RecBulkID.ActiveConnection = MM_Site_STRING
RecBulkID.Source = "SELECT * FROM BulkMCharge Where BulkID = " & WhichBulkID
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
	  alert("Error - This Bulk File Can't Be Deleted, File is Busy Being Processed");
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If

' Find Bulk Children
'Response.Write("<br>Looking For children<br>")
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM BulkMChargeChildren Where BulkID = " & WhichBulkID
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
TransDescritionEVD = "Bulk Update: " & RecBulkID.Fields.Item("FileName").Value & "_EVD"
' Find and delete the transaction in ZonerTransaction
set RecFindZonerTrans = Server.CreateObject("ADODB.Recordset")
RecFindZonerTrans.ActiveConnection = MM_Site_STRING
RecFindZonerTrans.Source = "SELECT Top(1)* FROM TediTransactions Where TediID = " & ZID & " and (CComments = '" & TransDescrition & "' or CComments = '" & TransDescritionEVD & "')  and CAmount = '" & MChargeValue & "'"
Response.Write("<br>" & RecFindZonerTrans.Source)
RecFindZonerTrans.CursorType = 0
RecFindZonerTrans.CursorLocation = 2
RecFindZonerTrans.LockType = 3
RecFindZonerTrans.Open()
RecFindZonerTrans_numRows = 0
While Not RecFindZonerTrans.EOF
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM TediTransactions Where CID = " + Replace(RecFindZonerTrans.Fields.Item("CID").Value, "'", "''") + ""
Response.Write("<br>" & rstSecond.Source)
rstSecond.Open
set rstSecond = nothing	

' Create Notification SMS
set RecZID = Server.CreateObject("ADODB.Recordset")
RecZID.ActiveConnection = MM_Site_STRINGWrite
RecZID.Source = "SELECT Top(1)* FROM ViewTediDetail Where TID = " & ZID
RecZID.CursorType = 0
RecZID.CursorLocation = 2
RecZID.LockType = 3
RecZID.Open()
RecZID_numRows = 0
MobileNo = RecZID.Fields.Item("TediCell").Value
ZonerCurrentBalance = RecZID.Fields.Item("MChargeBalance").Value
Msg = "M-Charge Reversal of R " & FormatNumber(MChargeValue,2) & " allocated to " & RecZID.Fields.Item("TediFirstName").Value & " " & RecZID.Fields.Item("TediLastName").Value  & " on " & TransDate & ", Regards PMG Live"

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("AlloID") = ZID
RecADDSMS("UserType") = "1"
RecADDSMS("SMSMSG") = Msg
RecADDSMS("MobileNo") = MobileNo
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close
' SMS Created and Sent

set RecNewCalcZonerMCharge = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerMCharge.ActiveConnection = MM_Site_STRING
RecNewCalcZonerMCharge.Source = "SELECT Top(1)* FROM ViewMchargeTediTotalAllocated Where TediID = " & ZID
RecNewCalcZonerMCharge.CursorType = 0
RecNewCalcZonerMCharge.CursorLocation = 2
RecNewCalcZonerMCharge.LockType = 3
RecNewCalcZonerMCharge.Open()
RecNewCalcZonerMCharge_numRows = 0
If Not RecNewCalcZonerMCharge.EOF Then
TediNewMCharge = RecNewCalcZonerMCharge.Fields.Item("TediTotalAllocated").Value
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecNewCalcZonerBanked = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerBanked.ActiveConnection = MM_Site_STRING
RecNewCalcZonerBanked.Source = "SELECT Top(1)* FROM ViewMchargeTediTotalBanked Where TediID = " & ZID
RecNewCalcZonerBanked.CursorType = 0
RecNewCalcZonerBanked.CursorLocation = 2
RecNewCalcZonerBanked.LockType = 3
RecNewCalcZonerBanked.Open()
RecNewCalcZonerBanked_numRows = 0
If Not RecNewCalcZonerBanked.EOF Then
TediNewBanked = RecNewCalcZonerBanked.Fields.Item("TediTotalBanked").Value
End If

ZTransTotal = (TediNewMCharge - TediNewBanked)

Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & ZID, MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MChargeBalance") = ZTransTotal
RecADDSMS.Update
RecADDSMS.Close

RecFindZonerTrans.MoveNext
Wend

' Item Removed From ZonerTransactions

' Delete Item From BulkMChargeChildren
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM BulkMChargeChildren Where ChildID = " + Replace(CID, "'", "''") + ""
rstSecond.Open
set rstSecond = nothing	
' Item Removed From BulkMChargeChildren



RecCurrent.MoveNext
Wend


Response.Redirect("Display.asp?AppCat=17&AppSubCatID=33")
%>