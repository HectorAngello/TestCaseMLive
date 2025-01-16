<!-- #include file="Connections/Site.asp" -->
<%
Dim RecFindPass__MMColParam
RecFindPass__MMColParam = "1"
If (Request.Form("EmailAddress") <> "") Then 
  RecFindPass__MMColParam = Request.Form("EmailAddress")
End If
%>
<%
Dim RecFindPass
Dim RecFindPass_cmd
Dim RecFindPass_numRows

Set RecFindPass_cmd = Server.CreateObject ("ADODB.Command")
RecFindPass_cmd.ActiveConnection = MM_Site_STRING
RecFindPass_cmd.CommandText = "SELECT * FROM Users WHERE UEmail = ?" 
RecFindPass_cmd.Prepared = true
RecFindPass_cmd.Parameters.Append RecFindPass_cmd.CreateParameter("param1", 200, 1, 255, RecFindPass__MMColParam) ' adVarChar

Set RecFindPass = RecFindPass_cmd.Execute
RecFindPass_numRows = 0

If Not RecFindPass.EOF Or Not RecFindPass.BOF Then 

		strBody = "Dear " & RecFindPass.Fields.Item("UserFirstName").Value & "<br>"
		strBody = strBody & "You requested your login details for the MTN Live system using the Forgot Password feature, please find you Login details below." & "<br>"
		strBody = strBody & "" & "<br>"
		strBody = strBody & "User Name : " & RecFindPass.Fields.Item("UserName").Value & "<br>"
		strBody = strBody & "Password : " & RecFindPass.Fields.Item("Password").Value & "<br>"
		strBody = strBody & "" & "<br>"
		strBody = strBody & "Regards, MTN Live Support Team" & "<br>"
		strBody = strBody & "http://pmg.mtnlive.co.za" & "<br>"
		strBody = strBody & "" & "<br>"


					Set objCDOSYSMail = Server.CreateObject("CDO.Message")
					Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
					
					'Outgoing SMTP server
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = OutMailServer
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = OutMailServerPort
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 240
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = OutmailServerUsername
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = OutmailserverPassword
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendTLS") = "True"
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "True"
					objCDOSYSCon.Fields.Update
					
					' Update the CDOSYS Configuration
					Set objCDOSYSMail.Configuration = objCDOSYSCon
					objCDOSYSMail.From = OutmailServerUsername
					objCDOSYSMail.To = RecFindPass.Fields.Item("UEmail").Value
					objCDOSYSMail.Subject = "MTN Live Login Details"
					objCDOSYSMail.HTMLBody = strBody
					objCDOSYSMail.Send
					
					'Close the server mail object
					Set objCDOSYSMail = Nothing
					Set objCDOSYSCon = Nothing

Response.Redirect("Default.asp?ForgotPass=True&Email=" & Request.Form("EmailAddress"))
            
Else
Response.Redirect("Default.asp?ForgotPass=False&Email=" & Request.Form("EmailAddress"))
End If
%>