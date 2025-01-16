<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecFetchSMS = Server.CreateObject("ADODB.Recordset")
RecFetchSMS.ActiveConnection = MM_Site_STRING
RecFetchSMS.Source = "SELECT * FROM viewUnsentBulkSMS Where BulkID = 48366" 
RecFetchSMS.CursorType = 0
RecFetchSMS.CursorLocation = 2
RecFetchSMS.LockType = 3
RecFetchSMS.Open()
RecFetchSMS_numRows = 0
While Not RecFetchSMS.EOF

GZDZID = RecFetchSMS.Fields.Item("TID").Value
%><!--#include file="Includes/GetTediDetail.inc" --><%
MobileNo = RecFetchSMS.Fields.Item("ASCell").Value
ASID = RecFetchSMS.Fields.Item("ASID").Value

Msg = "R " & FormatNumber(RecFetchSMS.Fields.Item("MChargeAmount").Value,2) & " allocated to Tedi " & RecFetchSMS.Fields.Item("TediEmpCode").Value & " - " & RecFetchSMS.Fields.Item("TediFirstName").Value & " " & RecFetchSMS.Fields.Item("TediLastName").Value & " By Administrator: " & RecFetchSMS.Fields.Item("UserFirstname").Value & " (" & RecFetchSMS.Fields.Item("TediEmpCode").Value & " Airtime Bal: R" & GZDZonerCurrentMChargeBalance & "), Regards PMG"
SendMRSMS = "Yes"
If SendMRSMS = "Yes" Then

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "2"
RecADDSMS("AlloID") = ASID
RecADDSMS("SMSMSG") = Msg
RecADDSMS("MobileNo") = MobileNo
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close

RecFetchSMS.MoveNext
End If
Response.Write("<br>" & Msg & " - Length: " & Len(Msg))
Wend
%>