<!-- #include file="Connections/Site.asp" -->
<%
' Session("DatabasePath") = "Path to your database"
  

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
    
	SQL = "SELECT * FROM Users WHERE username='" & strName & _
	     "' AND Password ='" & strPassword & "' and UserActive = 'True'"
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
      	struserGroup = RecordSet.Fields("UserSecurityGroupID")
	struserFullName = RecordSet.Fields("UserFirstName") & " " & RecordSet.Fields("UserLastName")

	
Session("UNID") = struserUNID
Session("UNGroupID") = struserGroup
Session("UNFullName") = struserFullName

' ----------------- Logging Starts ----------------- 

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT * FROM ChangeLog", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("ChangeType") = "User Logging Into The System"
rstChangeUpdate("ChangeBy") = Session("userUN") & " (SID=" & Session("UNID") & ")"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	
' ----------------- Logging Ends ----------------- 
If RecordSet.Fields("UserActive") = "False" Then
Response.Redirect "LogOut.asp"
End If
Session.Timeout=60
	Response.Redirect "Dashboard.asp"
	Else
        Response.Redirect "Default.asp?Login=Fail"
    End If

%>