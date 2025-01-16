<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
AllocateTo = Request.QueryString("AllocateTo")
FNBID = Request.QueryString("FNBID")
UNID = Request.QueryString("UNID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT * FROM MChargeFNBTrans where FNBID = " & FNBID, MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.Update
RecUpdateFNBTable("Allocated") = "True"
RecUpdateFNBTable("TediID") = AllocateTo
RecUpdateFNBTable("AllocatedBy") = UNID
RecUpdateFNBTable("AllocatedDate") = Now()
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close

set RecFNB = Server.CreateObject("ADODB.Recordset")
RecFNB.ActiveConnection = MM_Site_STRING
RecFNB.Source = "SELECT * FROM MChargeFNBTrans where FNBID = " & FNBID
RecFNB.CursorType = 0
RecFNB.CursorLocation = 2
RecFNB.LockType = 3
RecFNB.Open()
RecFNB_numRows = 0

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateZTTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateZTTable.Open "SELECT Top(2) * FROM TediTransactions", MM_Site_STRINGWrite, 1, 2
RecUpdateZTTable.AddNew
RecUpdateZTTable("CAmount") = RecFNB.Fields.Item("TransAmount").Value
RecUpdateZTTable("FNBID") = RecFNB.Fields.Item("FNBID").Value
RecUpdateZTTable("CDate") = RecFNB.Fields.Item("FNBDate").Value
RecUpdateZTTable("TediID") = AllocateTo
RecUpdateZTTable("CType") = "2"
RecUpdateZTTable("CComments") = RecFNB.Fields.Item("TransDescription").Value
RecUpdateZTTable("AddedBy") = UNID
RecUpdateZTTable.Update
RecUpdateZTTable.Close

Set RecUpdateAddCommission = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateAddCommission.Open "SELECT Top(1) * FROM AirtimeCommission", MM_Site_STRINGWrite, 1, 2
RecUpdateAddCommission.AddNew
RecUpdateAddCommission("ComDate") = RecFNB.Fields.Item("FNBDate").Value
RecUpdateAddCommission("ComPercentage") = AirtimeCommissionPercentage
RecUpdateAddCommission("BankedAmount") = RecFNB.Fields.Item("TransAmount").Value
RecUpdateAddCommission("ComAmount") = AirtimeCommissionPercentage * RecFNB.Fields.Item("TransAmount").Value
RecUpdateAddCommission("ComPaidOut") = "False"
RecUpdateAddCommission("TediID") = AllocateTo
RecUpdateAddCommission("FNBID") = RecFNB.Fields.Item("FNBID").Value
RecUpdateAddCommission.Update
RecUpdateAddCommission.Close

If AllocateTo <> "0" Then

set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT * FROM ViewTediDetail Where TID = " & AllocateTo
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0

GZDZID = AllocateTo
%><!--#include file="Includes/GetTediDetail.inc" --><%

Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & AllocateTo, MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MChargeBalance") = GZDZonerCurrentMChargeBalance
If IsNull(GZDLastBankedDate) = "True" Then
RecADDSMS("LastBankedDate") = RecFNB.Fields.Item("FNBDate").Value
Else
If DateDiff("d",GZDLastBankedDate,RecFNB.Fields.Item("FNBDate").Value) > 0 Then
RecADDSMS("LastBankedDate") = RecFNB.Fields.Item("FNBDate").Value
End If
End If

RecADDSMS.Update
RecADDSMS.Close

MobileNo = RecZoner.Fields.Item("TediCell").Value
Msg = "Dear " & RecZoner.Fields.Item("TediFirstName").Value & ", Your deposit of R " & FormatNumber(RecFNB.Fields.Item("TransAmount").Value,2) & " - Ref: " & Replace(RecFNB.Fields.Item("TransDescription").Value, "  ","") & " has been allocated to your account, Regards PMG"

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("AlloID") = RecZoner.Fields.Item("TID").Value
RecADDSMS("UserType") = "1"
RecADDSMS("SMSMSG") = Msg
RecADDSMS("MobileNo") = MobileNo
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close



Response.Redirect("Display.asp?AppCat=17&AppSubCatID=34")
Else
Response.Redirect("Display.asp?AppCat=17&AppSubCatID=34")
End If
%>

