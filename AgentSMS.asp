<!-- #include file="Connections/Site.asp" -->
<%
TID = Request.Form("TID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "1"
RecADDSMS("AlloID") = TID
RecADDSMS("SMSMSG") = Request.Form("message")
RecADDSMS("MobileNo") = Request.Form("ComSentTo")
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS("SentBy") = Session("UNID")
RecADDSMS.Update
RecADDSMS.Close


Response.redirect("TediView.asp?TID=" & TID & "&Item=2")
%>