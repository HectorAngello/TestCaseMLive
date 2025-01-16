<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
TID = Request.QueryString("TID") 
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM TediTransactions Where CID = " & Request.QueryString("CID")
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0

CID = RecCurrent.Fields.Item("CID").Value

If int(RecCurrent.Fields.Item("CType").Value) = 1 Then
Response.Write("Unallocate MCharge Allocation")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM TediTransactions Where CID = " + Replace(Request.Querystring("CID"), "'", "''") + ""
rstSecond.Open
set rstSecond = nothing	

Else
Response.Write("Unallocate FNB Deposit")
FNBID = RecCurrent.Fields.Item("FNBID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM TediTransactions Where CID = " + Replace(Request.Querystring("CID"), "'", "''") + ""
rstSecond.Open
set rstSecond = nothing	

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1) * FROM MChargeFNBTrans Where FNBID = " & FNBID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Allocated") = "False"
rstSecond("TediID") = "0"
rstSecond("Allocatedby") = Session("UNID")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

set RecHasCommission = Server.CreateObject("ADODB.Recordset")
RecHasCommission.ActiveConnection = MM_Site_STRING
RecHasCommission.Source = "SELECT Top(1)* FROM AirtimeCommission Where FNBID = " & FNBID
RecHasCommission.CursorType = 0
RecHasCommission.CursorLocation = 2
RecHasCommission.LockType = 3
RecHasCommission.Open()
RecHasCommission_numRows = 0
If Not RecHasCommission.EOF and Not RecHasCommission.BOF Then
ComID = RecHasCommission.Fields.Item("ComID").Value
AlreadyAllocated = RecHasCommission.Fields.Item("AlreadyAllocated").Value

If AlreadyAllocated = "True" Then
AllocatedToCID = RecHasCommission.Fields.Item("AllocatedToCID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1) * FROM TediTransactions Where CID = " & AllocatedToCID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Ctype") = "1"
rstSecond("CComments") = "Airtime Commission Reversal"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM AirtimeCommission Where ComID = " & ComID
rstSecond.Open
set rstSecond = nothing	
End If

End If

TediNewBanked = 0
TediNewMCharge = 0

set RecNewCalcZonerMCharge = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerMCharge.ActiveConnection = MM_Site_STRING
RecNewCalcZonerMCharge.Source = "SELECT * FROM ViewMchargeTediTotalAllocated Where TediID = " & TID
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
RecNewCalcZonerBanked.Source = "SELECT * FROM ViewMchargeTediTotalBanked Where TediID = " & TID
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
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MChargeBalance") = ZTransTotal
RecADDSMS.Update
RecADDSMS.Close

Response.Redirect("TediView.asp?TID=" & Request.QueryString("TID") & "&Item=9")
%>