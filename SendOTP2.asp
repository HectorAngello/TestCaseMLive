<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
' create Token
Part1 = Replace(Date, "/", "")
	Function RandomNumber(intHighestNumber)
		Randomize
		RandomNumber = Int(Rnd * intHighestNumber) + 1
	End Function

Part2 = Int(Right(5000 + RandomNumber(3000000000000),5))

If Len(Part2) < 8 Then
Part2 = Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2
Part2 = Left(Part2,8)
End If
Part2 = "-" & Part2

Part3 = Int(Right(5000 + RandomNumber(9000000000000),5))
If Len(Part3) < 10 Then
Part3 = Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3
Part3 = Left(Part3,10)
End If

Part3 = "-" & Part3

Part4 = "-" & DatePart("h",Now()) & "-" & DatePart("n",Now())
Token = Part1 & Part2 & Part3 & Part4
' End Create Token

OTP = Request.QueryString("OTP")

If Session("UNID") = "" Then
   Response.Redirect "Loginerror.asp" 
End If

SID = Session("UNID")

set RecLastOTP = Server.CreateObject("ADODB.Recordset")
RecLastOTP.ActiveConnection = MM_Site_STRING
RecLastOTP.Source = "SELECT * FROM OTP Where SID = '" & Session("UNID")  & "' and OTPKey = " & OTP
'Response.Write(RecLastOTP.Source)
RecLastOTP.CursorType = 0
RecLastOTP.CursorLocation = 2
RecLastOTP.LockType = 3
RecLastOTP.Open()
RecLastOTP_numRows = 0
If Not RecLastOTP.EOF and Not RecLastOTP.BOF Then
Response.Write("Found The OTP For The User")

SecNow = Now()
SecOTP = RecLastOTP.Fields.Item("LogInDate").Value
SecTotal = DateDiff("s",SecOTP, SecNow)
Response.Write("<br>" & SecTotal)

If Sectotal > 1201 Then
Response.Redirect("SendOTP.asp?ErrorCode=1")
End If

' Valid OTP Found - Has Not expired - Log the User In
set RecUser = Server.CreateObject("ADODB.Recordset")
RecUser.ActiveConnection = MM_Site_STRING
RecUser.Source = "SELECT * FROM ViewUserDetail Where UserID = " & Session("UNID")
RecUser.CursorType = 0
RecUser.CursorLocation = 2
RecUser.LockType = 3
RecUser.Open()
RecUser_numRows = 0

LastLoginText = "Last Login (n/a)"
CheckLastDayInterval = 0
UserQry = RecUser.Fields.Item("Username").Value & " (UserID=" & RecUser.Fields.Item("UserID").Value & ")"
set RecLastLogIn = Server.CreateObject("ADODB.Recordset")
RecLastLogIn.ActiveConnection = MM_Site_STRING
RecLastLogIn.Source = "SELECT * FROM ChangeLog Where ChangeType = 'User Logged In - Completed OTP Process' and ChangeBy = '" & UserQry & "' Order By ID Desc"
'Response.Write(RecLastLogIn.Source)
RecLastLogIn.CursorType = 0
RecLastLogIn.CursorLocation = 2
RecLastLogIn.LockType = 3
RecLastLogIn.Open()
RecLastLogIn_numRows = 0
If Not RecLastLogIn.EOF and Not RecLastLogIn.BOF Then
LastLoginText = "Last Login: " & Day(RecLastLogIn.Fields.Item("ChangeDate").Value) & " " & MonthName(Month(RecLastLogIn.Fields.Item("ChangeDate").Value)) & " " & Year(RecLastLogIn.Fields.Item("ChangeDate").Value) & " " & FormatDateTime(RecLastLogIn.Fields.Item("ChangeDate").Value,3)
CLDI = DateDiff("d",RecLastLogIn.Fields.Item("ChangeDate").Value,Date)
End If

	struserUNID = RecUser.Fields("UserID")
      	struserGroup = RecUser.Fields("UserSecurityGroupID")
	struserFullName = RecUser.Fields("UserFirstName") & " " & RecUser.Fields("UserLastName")
	struserComp = RecUser.Fields("CompanyID")
	struserCN = RecUser.Fields("CompanyName")
	
Session("UNID") = struserUNID
Session("UNGroupID") = struserGroup
Session("userUN") = struserFullName
Session("CompanyID") = struserComp
Session("CompanyName") = struserCN

' ----------------- Logging Starts ----------------- 

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(2) * FROM ChangeLog", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("ChangeType") = "User Logged In - Completed OTP Process"
rstChangeUpdate("ChangeBy") = Session("UNID")
rstChangeUpdate("ChangeDate") = Now()
rstChangeUpdate("Changes") = "User Logged in Successfully in " & Sectotal & " Seconds"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	
' ----------------- Logging Ends ----------------- 
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT * FROM OTP Where ID = " & RecLastOTP.Fields.Item("ID").Value, MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.Update
rstChangeUpdate("WasUsed") = "Yes"
rstChangeUpdate("LoginSeconds") = Sectotal
rstChangeUpdate("DMGToken") = Token
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing

' Login Complete

Session.Timeout=60

Response.Redirect "Dashboard.asp"

Else
Response.Redirect "Default.asp?Login=Fail"
End If
%>