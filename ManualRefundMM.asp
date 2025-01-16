<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
RefundType = Request.Form("RefundType")
' Find Zoner
set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT * FROM ViewTediDetail Where TID = " & Request.Form("TID")
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0

set RecUpdateSysUser = Server.CreateObject("ADODB.Recordset")
RecUpdateSysUser.ActiveConnection = MM_Site_STRING
RecUpdateSysUser.Source = "Select * FROM Users Where UserID = " & Request.Form("UserID")
RecUpdateSysUser.CursorType = 0
RecUpdateSysUser.CursorLocation = 2
RecUpdateSysUser.LockType = 3
RecUpdateSysUser.Open()
RecUpdateSysUser_numRows = 0
UpdateBy = RecUpdateSysUser.Fields.Item("UserFirstName").Value & " " & RecUpdateSysUser.Fields.Item("UserLastName").Value
If RefundType = "4" Then
' Add Transaction to FNBTrans
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTrans = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTrans.Open "SELECT Top(1) * FROM MChargeFNBTransMM", MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTrans.AddNew
RecUpdateFNBTrans("ImportDate") = Now()
RecUpdateFNBTrans("FNBDate") = Now()
RecUpdateFNBTrans("ServiceFee") = "0"
RecUpdateFNBTrans("TransAmount") = Replace(Request.Form("CAmount"), " ","")
RecUpdateFNBTrans("TransDescription") = "Manual Refund By " & UpdateBy
RecUpdateFNBTrans("TransChequeNo") = "0"
RecUpdateFNBTrans("AccountBalance") = "0"
RecUpdateFNBTrans("Allocated") = "True"
RecUpdateFNBTrans("TediID") = Request.Form("TID")
RecUpdateFNBTrans("AllocatedBy") = Request.Form("UserID")
RecUpdateFNBTrans("AllocatedDate") = Now()
RecUpdateFNBTrans.Update
RecUpdateFNBTrans.Close

set RecNewsestFNBTrans = Server.CreateObject("ADODB.Recordset")
RecNewsestFNBTrans.ActiveConnection = MM_Site_STRING
RecNewsestFNBTrans.Source = "SELECT Top(2)* FROM MChargeFNBTransMM Where TediID = '" & Request.Form("TID") & "' Order By FNBID Desc"
RecNewsestFNBTrans.CursorType = 0
RecNewsestFNBTrans.CursorLocation = 2
RecNewsestFNBTrans.LockType = 3
RecNewsestFNBTrans.Open()
RecNewsestFNBTrans_numRows = 0
NewestFNBID = RecNewsestFNBTrans.Fields.Item("FNBID").Value

'Update Zoner TransActions
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT Top(1) * FROM TediTransactionsMM", MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.AddNew
RecUpdateFNBTable("CAmount") = RecNewsestFNBTrans.Fields.Item("TransAmount").Value
RecUpdateFNBTable("CDate") = RecNewsestFNBTrans.Fields.Item("FNBDate").Value
RecUpdateFNBTable("TediID") = RecNewsestFNBTrans.Fields.Item("TediID").Value
RecUpdateFNBTable("CType") = "4"
RecUpdateFNBTable("CComments") = RecNewsestFNBTrans.Fields.Item("TransDescription").Value
RecUpdateFNBTable("AddedBy") = Request.Form("UserID")
RecUpdateFNBTable("FNBID") = RecNewsestFNBTrans.Fields.Item("FNBID").Value
RecUpdateFNBTable("SRID") = RecZoner.Fields.Item("SRID").Value
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close
End If



set RecNewCalcZonerMCharge = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerMCharge.ActiveConnection = MM_Site_STRING
RecNewCalcZonerMCharge.Source = "SELECT * FROM ViewMchargeTediTotalAllocatedMM Where TediID = " & Request.Form("TID")
'Response.write(RecNewCalcZonerMCharge.Source)
RecNewCalcZonerMCharge.CursorType = 0
RecNewCalcZonerMCharge.CursorLocation = 2
RecNewCalcZonerMCharge.LockType = 3
RecNewCalcZonerMCharge.Open()
RecNewCalcZonerMCharge_numRows = 0
If Not RecNewCalcZonerMCharge.EOF Then
TediNewMCharge = RecNewCalcZonerMCharge.Fields.Item("TediTotalAllocatedMM").Value
End If

If RefundType = "5" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT Top(1) * FROM TediTransactionsMM", MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.AddNew
RecUpdateFNBTable("CAmount") =  Replace(Request.Form("CAmount"), " ","")
RecUpdateFNBTable("CDate") = Now()
RecUpdateFNBTable("TediID") = Request.Form("TID")
RecUpdateFNBTable("CType") = "5"
RecUpdateFNBTable("CComments") = "Manual Refund By " & UpdateBy
RecUpdateFNBTable("AddedBy") = Request.Form("UserID")
RecUpdateFNBTable("SRID") = RecZoner.Fields.Item("SRID").Value
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecNewCalcZonerBanked = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerBanked.ActiveConnection = MM_Site_STRING
RecNewCalcZonerBanked.Source = "SELECT * FROM ViewMchargeTediTotalBankedMM Where TediID = " & Request.Form("TID")
RecNewCalcZonerBanked.CursorType = 0
RecNewCalcZonerBanked.CursorLocation = 2
RecNewCalcZonerBanked.LockType = 3
RecNewCalcZonerBanked.Open()
RecNewCalcZonerBanked_numRows = 0
If Not RecNewCalcZonerBanked.EOF Then
TediNewBanked = RecNewCalcZonerBanked.Fields.Item("TediTotalBankedMM").Value
End If

ZTransTotal = (TediNewMCharge - TediNewBanked)

Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & Request.Form("TID"), MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MobileMoneyBalance") = ZTransTotal
RecADDSMS.Update
RecADDSMS.Close

Response.Redirect("TediView.asp?TID=" & RecZoner.Fields.Item("TID").Value)
%>