<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "2"
RecADDSMS("AlloID") = ASID
RecADDSMS("SMSMSG") = Request.Form("message")
RecADDSMS("MobileNo") = Request.Form("ComSentTo")
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS("SentBy") = Session("UNID")
RecADDSMS.Update
RecADDSMS.Close


Response.redirect("ASView.asp?ASID=" & ASID & "&Item=2")
%>