<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
TID = Request.form("TID")
UNID = Request.form("UNID")
ASID = Request.form("ASID")

ZonerEmail = Request.form("ZonerEmail")
ZonerName = Request.Form("ZonerName")
ZonerCell = Request.Form("ZonerCell")

EmailFrom = "noreply@ir.co.za"
EmailSubject = "Bulk Sim Notification"

EmailBody = "Dear " & ZonerName
EmailBody = EmailBody & "<br>"
EmailBody = EmailBody & "<br>The following sim numbers have been allocated against your MTN Live profile:"
EmailBody = EmailBody & "<br>"

EmailTail = "<br>"
EmailTail = "<br>Please do not reply to this email address as it is system generated."
EmailTail = "<br>"
EmailTail = EmailTail & "<br>Regards,"
EmailTail = EmailTail & "<br>MTN Live"

ThisToken = UNID & "-" & DatePart("h",Now()) & ":" & DatePart("n",Now()) & ":" & DatePart("s",Now())

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1) * FROM BulkSims", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("BulkDate") = Now()
rstSecond("AllocatedBy") = UNID
rstSecond("TID") = TID
rstSecond("Token") = ThisToken
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

set RecNewest = Server.CreateObject("ADODB.Recordset")
RecNewest.ActiveConnection = MM_Site_STRING
RecNewest.Source = "SELECT Top(1)* FROM BulkSims Where Token = '" & ThisToken & "'"
RecNewest.CursorType = 0
RecNewest.CursorLocation = 2
RecNewest.LockType = 3
RecNewest.Open()
RecNewest_numRows = 0
BulkID = RecNewest.Fields.Item("BulkID").Value

set RecBrick = Server.CreateObject("ADODB.Recordset")
RecBrick.ActiveConnection = MM_Site_STRING
RecBrick.Source = "SELECT * FROM Sims Where ((BoxNumber = '" & Request.Form("BrickCode1") & "' or BoxNumber = '" & Request.Form("BrickCode2") & "' or BoxNumber = '" & Request.Form("BrickCode3") & "' or BoxNumber = '" & Request.Form("BrickCode4") & "' or BoxNumber = '" & Request.Form("BrickCode5") & "' or BoxNumber = '" & Request.Form("BrickCode6") & "' or BoxNumber = '" & Request.Form("BrickCode7") & "' or BoxNumber = '" & Request.Form("BrickCode8") & "' or BoxNumber = '" & Request.Form("BrickCode9") & "' or BoxNumber = '" & Request.Form("BrickCode10") & "' ) or (SerialNo = '" & Request.Form("BrickCode1") & "' or  SerialNo = '" & Request.Form("BrickCode2") & "' or  SerialNo = '" & Request.Form("BrickCode3") & "' or  SerialNo = '" & Request.Form("BrickCode4") & "' or  SerialNo = '" & Request.Form("BrickCode5") & "' or  SerialNo = '" & Request.Form("BrickCode6") & "' or  SerialNo = '" & Request.Form("BrickCode7") & "' or  SerialNo = '" & Request.Form("BrickCode8") & "' or  SerialNo = '" & Request.Form("BrickCode9") & "' or SerialNo = '" & Request.Form("BrickCode10") & "' ) or (BrickNumber = '" & Request.Form("BrickCode1") & "' or BrickNumber = '" & Request.Form("BrickCode2") & "' or BrickNumber = '" & Request.Form("BrickCode3") & "' or BrickNumber = '" & Request.Form("BrickCode4") & "' or BrickNumber = '" & Request.Form("BrickCode5") & "' or BrickNumber = '" & Request.Form("BrickCode6") & "' or BrickNumber = '" & Request.Form("BrickCode7") & "' or BrickNumber = '" & Request.Form("BrickCode8") & "' or BrickNumber = '" & Request.Form("BrickCode9") & "' or BrickNumber = '" & Request.Form("BrickCode10") & "' )) and ASID = " & ASID & " and AllocatedTo = 0 Order By BrickNumber, BoxNumber, SerialNo Asc"
RecBrick.CursorType = 0
RecBrick.CursorLocation = 2
RecBrick.LockType = 3
RecBrick.Open()
RecBrick_numRows = 0

While Not RecBrick.EOF
DoInsert = "Yes"

set RecChecking = Server.CreateObject("ADODB.Recordset")
RecChecking.ActiveConnection = MM_Site_STRING
RecChecking.Source = "SELECT * FROM BulkSimChildren Where SerialNo = '" & RecBrick.Fields.Item("SerialNo").Value & "'"
RecChecking.CursorType = 0
RecChecking.CursorLocation = 2
RecChecking.LockType = 3
RecChecking.Open()
RecChecking_numRows = 0
While Not RecChecking.EOF
DoInsert = "No"
RecChecking.MoveNext
Wend
'Response.write("<br>" & DoInsert)
'Response.write("<br>" & RecBrick.Fields.Item("SerialNo").Value)
If DoInsert = "Yes" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1) * FROM BulkSimChildren", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("BulkID") = BulkID
rstSecond("ChildCreationDate") = Now()
rstSecond("TID") = TID
rstSecond("SerialNo") = RecBrick.Fields.Item("SerialNo").Value
rstSecond("SIMID") = RecBrick.Fields.Item("SimID").Value
rstSecond("Token") = ThisToken
rstSecond("AllocatedBy") = UNID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	
Response.write("<br>" & RecBrick.Fields.Item("SerialNo").Value)
EmailBody = EmailBody & "<br>" & RecBrick.Fields.Item("SerialNo").Value


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstUpDate = Server.CreateObject ( "ADODB.Recordset" )
rstUpDate.Open "SELECT Top(1)* FROM Sims Where SimID = " & RecBrick.Fields.Item("SimID").Value, MM_Site_STRINGWrite, 1, 2
rstUpDate.Update
rstUpDate("AllocatedTo") = TID
rstUpDate("Allocated") = "True"
rstUpDate("AllocatedDate") = Now()
rstUpDate.Update
rstUpDate.Close
set rstUpDate = nothing	
End If
RecBrick.MoveNext
Wend

If ZonerEmail <> "" Then
EmailMSG = EmailBody & EmailTail

					Set objCDOSYSMail = Server.CreateObject("CDO.Message")
					Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
					
					'Outgoing SMTP server
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = OutMailServer
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = OutMailServerPort
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = OutmailServerUsername
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = OutmailserverPassword
					'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/startTLS") = "True"
					objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "True"
					objCDOSYSCon.Fields.Update
					
					' Update the CDOSYS Configuration
					Set objCDOSYSMail.Configuration = objCDOSYSCon
					objCDOSYSMail.From = EmailFrom
					'objCDOSYSMail.To = ZonerEmail
					objCDOSYSMail.To = "webmaster@bump.co.za"
					objCDOSYSMail.Subject = EmailSubject
					objCDOSYSMail.HTMLBody = EmailMSG
					'objCDOSYSMail.Send
					
					'Close the server mail object
					Set objCDOSYSMail = Nothing
					Set objCDOSYSCon = Nothing

End If

Response.Redirect("TediView.asp?TID=" & TID & "&Item=14")


%>
