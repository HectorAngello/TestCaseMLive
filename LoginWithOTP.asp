<!-- #include file="Connections/Site.asp" -->
<%
' Session("DatabasePath") = "Path to your database"
  If Request.Form("login") = "Login" Then

    '-- Declare your variables
    Dim DataConnection, cmdDC, RecordSet
    Dim RecordToEdit, Updated, strUserName, strPassword
	
	strName = Replace(Request.Form("UserName"), " ", "")
    strPassword = Replace(Request.Form("Password"), " ", "")
    strName = Replace(strName, "'", "_")
    strPassword = Replace(strPassword, "'", "_")
	strName = Replace(strName, """", "_")
    strPassword = Replace(strPassword, """", "_")
	
    '-- Create object and open database
 	Set DataConnection = Server.CreateObject("ADODB.Connection")
		DataConnection.Open MM_Site_STRING
		
	Set cmdDC = Server.CreateObject("ADODB.Command")
 	cmdDC.ActiveConnection = DataConnection
    
	SQL = "SELECT * FROM Users WHERE UserActive= 'True' and username='" & strName & _
	     "' AND password ='" & strPassword & "'"
    'Response.write SQL
    cmdDC.CommandText = SQL
    Set RecordSet = Server.CreateObject("ADODB.Recordset")

    '-- Cursor Type, Lock Type
    '-- ForwardOnly 0 - ReadOnly 1
    '-- KeySet 1 - Pessimistic 2
    '-- Dynamic 2 - Optimistic 3
    '-- Static 3 - BatchOptimistic 4
    RecordSet.Open cmdDC, , 0, 2

If Not RecordSet.EOF Then
	  Dim struserLevel, strusercat, struserUN
      	struserUNID = RecordSet.Fields("UserID")
      	Session("UNID") = struserUNID

' ----------------- Logging Starts ----------------- 

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(2) * FROM ChangeLog", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("ChangeType") = "User Logging Into The System"
rstChangeUpdate("ChangeDate") = Now()
rstChangeUpdate("ChangeBy") = Session("UNID")
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	
' ----------------- Logging Ends ----------------- 
If RecordSet.Fields("UserActive") = "No" Then
Response.Redirect "LogOut.asp"
End If
Session.Timeout=60
	  Response.Redirect "SendOTP.asp"
	Else
      'The user was not validated...
      'Take them to a page which tells them they were not validated...
     Response.Redirect "Default.asp?Login=Fail"
    End If
  End If
%>