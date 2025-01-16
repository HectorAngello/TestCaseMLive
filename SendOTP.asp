<!-- #include file="includes/header.asp" -->
<%
NBBSMSCon = "Provider=sqloledb;Data Source=196.38.88.105;Initial Catalog=IRSmS;uid=sa;pwd=Xm3nSt0ry;"

If Session("UNID") = "" Then
   Response.Redirect "Loginerror.asp" 
End If 
'Response.Write(Session("UNID"))
set RecUser = Server.CreateObject("ADODB.Recordset")
RecUser.ActiveConnection = MM_Site_STRING
RecUser.Source = "SELECT * FROM Users Where UserID = " & Session("UNID")
RecUser.CursorType = 0
RecUser.CursorLocation = 2
RecUser.LockType = 3
RecUser.Open()
RecUser_numRows = 0
'Response.Write(RecUser.Fields.Item("MobileNumber").Value)
If Request.QueryString("ErrorCode") = "1" or Request.QueryString("ErrorCode") = "0" Then
Else
Randomize
Seed = FormatNumber((9999999 * Rnd),0,,,0)
'Response.Write(Seed)
MSG = Seed & " - Please enter this OTP into MTN Live - " & Day(Now) & " " & MonthName(Month(Now),True) & " " & Year(Now) & " " & Time & " - MTN Live"
MSG2 = "******** - Please enter this OTP into MTN Live - " & Day(Now) & " " & MonthName(Month(Now),True) & " " & Year(Now) & " " & Time & " - MTN Live"
'Response.Write(MSG)
'Response.END
If Request.QueryString("ResendOTP") = "Yes" Then

set RecLastOTP = Server.CreateObject("ADODB.Recordset")
RecLastOTP.ActiveConnection = MM_Site_STRING
RecLastOTP.Source = "SELECT * FROM OTP Where SID = '" & Session("UNID")  & "' Order By ID Desc"
'Response.Write(RecLastOTP.Source)
RecLastOTP.CursorType = 0
RecLastOTP.CursorLocation = 2
RecLastOTP.LockType = 3
RecLastOTP.Open()
RecLastOTP_numRows = 0

MSG3 =  RecLastOTP.Fields.Item("OTPKey").Value & " - Please enter this OTP into MTN Live - " & Day(RecLastOTP.Fields.Item("LogInDate").Value) & " " & MonthName(Month(RecLastOTP.Fields.Item("LogInDate").Value),True) & " " & Year(RecLastOTP.Fields.Item("LogInDate").Value) & " " & Time & " - MTN Live"
MSG4 = "******** - Please enter this OTP into MTN Live - " & Day(RecLastOTP.Fields.Item("LogInDate").Value) & " " & MonthName(Month(RecLastOTP.Fields.Item("LogInDate").Value),True) & " " & Year(RecLastOTP.Fields.Item("LogInDate").Value) & " " & Time & " - MTN Live"

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMS_Queue", SMSG8Con, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("SQ_ClientRef") = SMSUserID
rstChangeUpdate("SQ_Processed") = "False"
rstChangeUpdate("SQ_ImportedDate") = Now()
rstChangeUpdate("SQ_SendDateTime") = Now()
rstChangeUpdate("SQ_CellNumber") = RecLastOTP.Fields.Item("CellNo").Value
rstChangeUpdate("SQ_Message") = MSG3
rstChangeUpdate("SQ_Priority") = "9"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing



Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstAddOTPToDB = Server.CreateObject ( "ADODB.Recordset" )
rstAddOTPToDB.Open "SELECT Top(2) * FROM OTPSMS", MM_Site_STRINGWrite, 1, 2
rstAddOTPToDB.AddNew
rstAddOTPToDB("SMSNo") = RecLastOTP.Fields.Item("CellNo").Value
rstAddOTPToDB("SMSMSG") = MSG4
rstAddOTPToDB("Processed") = "No"
rstAddOTPToDB.Update
rstAddOTPToDB.Close
set rstAddOTPToDB = nothing

' New Code Ends

Else
' ----------------- Send OTP Starts ----------------- 
If RecUser.Fields.Item("UserActive").Value = "True" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(2) * FROM OTP", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("SID") = Session("UNID")
rstChangeUpdate("LogInDate") = Now()
rstChangeUpdate("OTPKey") = Seed
rstChangeUpdate("WasUsed") = "No"
rstChangeUpdate("CellNo") = RecUser.Fields.Item("CellNo").Value
rstChangeUpdate("MSG") = MSG
rstChangeUpdate("LoginSeconds") = "0"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(2) * FROM ChangeLog", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("ChangeType") = "Sending OTP to " & Session("UNID")
rstChangeUpdate("ChangeBy") = Session("UNID")
rstChangeUpdate("Changes") = MSG
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	
' Old OTP Sending Code To Nashua And TheSMSPro

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMS_Queue", SMSG8Con, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("SQ_ClientRef") = SMSUserID
rstChangeUpdate("SQ_Processed") = "False"
rstChangeUpdate("SQ_ImportedDate") = Now()
rstChangeUpdate("SQ_SendDateTime") = Now()
rstChangeUpdate("SQ_CellNumber") = RecUser.Fields.Item("CellNo").Value
rstChangeUpdate("SQ_Message") = MSG
rstChangeUpdate("SQ_Priority") = "9"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstAddOTPToDB = Server.CreateObject ( "ADODB.Recordset" )
rstAddOTPToDB.Open "SELECT Top(2) * FROM OTPSMS", MM_Site_STRINGWrite, 1, 2
rstAddOTPToDB.AddNew
rstAddOTPToDB("SMSNo") = RecUser.Fields.Item("CellNo").Value
rstAddOTPToDB("SMSMSG") = MSG
rstAddOTPToDB("Processed") = "No"
rstAddOTPToDB.Update
rstAddOTPToDB.Close
set rstAddOTPToDB = nothing

' New Code Ends
End If
' ----------------- Send OTP Ends ----------------- 
End If
End If
set RecSiteInfo = Server.CreateObject("ADODB.Recordset")
RecSiteInfo.ActiveConnection = MM_Site_STRING
RecSiteInfo.Source = "SELECT * FROM SiteInfo Where ID = 1"
RecSiteInfo.CursorType = 0
RecSiteInfo.CursorLocation = 2
RecSiteInfo.LockType = 3
RecSiteInfo.Open()
RecSiteInfo_numRows = 0


%>
<!-- header -->
    <div class="header">
        <div class="row">
            <div class="nine columns">
                <!-- #include file="includes/Logo.inc" -->
            </div>
            <div class="three columns">
                <div class="logged-in">
                    <!-- #include file="Includes/User.inc" -->
                </div>
            </div>
        </div>
    </div>
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            
            <div class="twelve columns">
                <div class="content">
                    <div class="row heading">
                        <div id="login-box" class="eight columns centered">
                <div class="panel">
                    <div class="row">
                        <div class="twelve columns">
                            <%
If RecUser.Fields.Item("UserActive").Value = "True" Then
If Request.QueryString("ErrorCode") = "" Then
Else
If Request.QueryString("ErrorCode") = "0" Then
ErrorText = "Incorrect OTP"
End If
if Request.QueryString("ErrorCode") = "1" Then
ErrorText = "OTP Has Expired, Please Log In Again To Request A New OTP .....<A href=default.asp>Click Here To Log In Again</a>"
End If
%>
<div class="alert-box error"><%=ErrorText%></div>
<%
End If
%>
<%If Request.QueryString("ResendOTP") = "Yes" Then%><div class="alert-box success">OTP Resent</div><%End If%>
<% If Request.QueryString("ErrorCode") = "1" or Request.QueryString("ErrorCode") = "0" Then
Else%>
<span class="span-heading">Welcome Back: <%=(RecUser.Fields.Item("UserName").Value)%></span>

Your OTP Has been SMS'ed To: <b><%=(RecUser.Fields.Item("CellNo").Value)%></b><br><br>
<%End If%>
<%If Request.QueryString("ErrorCode") = "1" Then
Else
%>
<form action="SendOTP2.asp" method="get" onSubmit="MM_validateForm('OTP','','RisNum');return document.MM_returnValue">
<table border="0" align="center" cellpadding="2" cellspacing="2">
<tr><td Class="quote">User:</td><td><%=(RecUser.Fields.Item("UserName").Value)%></td></tr>
<tr><td Class="quote">Mobile Number:</td><td><%=(RecUser.Fields.Item("CellNo").Value)%></td></tr>
<tr><td Class="quote">OTP:</td><td><input name="OTP" type="text" class="offtab2" id="OTP" Size="10"></td></tr>
<tr><td Colspan="2"><input name="button2" type="submit"  class="nice red radius button" value="Complete Login"></td>
</tr>
</table>
</form>
<br>
<br>If you have not received your OTP within 2 minutes, <a href="SendOTP.asp?ResendOTP=Yes">Please click here to resend it</a>.<br>(Please be patient as this can take up to 2 minutes to deliver the OTP)
<%
End If
Else%>
<div class="alert-box error">Your Account Is No Longer Active In The System,<br>Please Contact GenesisLive to Reactivate Your Account</div>
<%End If%>
                        </div>
                    </div>
                </div>
            </div>
                    </div>
                    
<!-- #include file="includes/footer.asp" -->
<% OTPCounter = 0
set RecSendOTPSMS = Server.CreateObject("ADODB.Recordset")
RecSendOTPSMS.ActiveConnection = MM_Site_STRING
RecSendOTPSMS.Source = "SELECT Top(1)* FROM OTPSMS Where Processed = 'No'"
RecSendOTPSMS.CursorType = 0
RecSendOTPSMS.CursorLocation = 2
RecSendOTPSMS.LockType = 3
RecSendOTPSMS.Open()
RecSendOTPSMS_numRows = 0
While Not RecSendOTPSMS.EOF
OTPCounter = OTPCounter + 1

MsgOriginal = RecSendOTPSMS.Fields.Item("SMSMSG").Value
MSGLen = Len(MsgOriginal)
MsgUpdated = Right(RecSendOTPSMS.Fields.Item("SMSMSG").Value, MSGLen - 6)
MsgUpdated = "********" & MsgUpdated
'Insert Orginal SMS Into Que


' Insert ******** Message into Schedule
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond2 = Server.CreateObject ( "ADODB.Recordset" )
rstSecond2.Open "SELECT Top(1)* FROM Schedule", TheSMSProDBConn, 1, 2
rstSecond2.AddNew
rstSecond2("UserID") = SMSUserID
rstSecond2("SchedDate") = Month(Now()) & "/" & Day(Now()) & "/" & Year(Now())
rstSecond2("SchedTime") = FormatDateTime(Now(),4)
rstSecond2("SendType") = "Single"
rstSecond2("SchedNumber") = "OTP"
rstSecond2("SchedMsg") = MsgUpdated
rstSecond2("SchedProcessed") = "No"
rstSecond2("Result") = "Pending"
rstSecond2("AddedToSched") = Now()
rstSecond2.Update
rstSecond2.Close
set rstSecond2 = nothing

' Update OTP In Database as sent
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT Top(1)* FROM OTPSMS where ID = " & RecSendOTPSMS.Fields.Item("ID").Value, MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.Update
RecUpdateFNBTable("Processed") = "Yes"
RecUpdateFNBTable("ProcessedTime") = Now()
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close

RecSendOTPSMS.MoveNext
Wend
%>

