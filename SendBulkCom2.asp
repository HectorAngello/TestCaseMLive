<!-- #include file="Connections/Site.asp" -->
<%
Region = Request.Form("Region")
SendType = Request.Form("SendType")
ComType = Request.Form("ComType")
EmailFrom = Request.Form("EmailFrom")
EmailSubJect = Request.Form("EmailSubJect")
MSG = Request.Form("MSG")
Function isEmailValid(email) 
        Set regEx = New RegExp 
        regEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$" 
        isEmailValid = regEx.Test(trim(email)) 
    End Function 

	Function testEmail(email) 
        response.write "<p>" & email & " (" & _ 
            isEmailValid(email) & ")" 
    End Function 

If ComType = "Email" Then


RFILE = Request.Form("Region")
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), "--")
Region = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
RepID = Replace(RFILE, "-", "")
SendType = Request.Form("SendType")
DisplaySendText = ""
If SendType = "1" Then
DisplaySendText = "All Agents And " & SupervisorLabel & "s"
End If
If SendType = "2" Then
DisplaySendText = "ONLY Agents"
End If
If SendType = "3" Then
DisplaySendText = "ONLY " & SupervisorLabel & "s"
End If

set RecRegionSelect = Server.CreateObject("ADODB.Recordset")
RecRegionSelect.ActiveConnection = MM_Site_STRING
If Region = "0" then
RecRegionSelect.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
Else
RecRegionSelect.Source = "SELECT * FROM [Regions] Where RID = " & Region
End If
RecRegionSelect.CursorType = 0
RecRegionSelect.CursorLocation = 2
RecRegionSelect.LockType = 3
RecRegionSelect.Open()
RecRegionSelect_numRows = 0

While Not RecRegionSelect.EOF
set RecWhichRep = Server.CreateObject("ADODB.Recordset")
RecWhichRep.ActiveConnection = MM_Site_STRING
If RepID = "0" Then
RecWhichRep.Source = "SELECT * FROM ASs Where ASActive = 'True' and RID = '" & RecRegionSelect.Fields.Item("RID").Value & "' order by ASFirstName Asc"
Else
RecWhichRep.Source = "SELECT * FROM ASs Where TLID = " & RepID
End If
'Response.write(RecWhichRep.Source)
RecWhichRep.CursorType = 0
RecWhichRep.CursorLocation = 2
RecWhichRep.LockType = 3
RecWhichRep.Open()
RecWhichRep_numRows = 0
RepCount = 0
While Not RecWhichRep.EOF
If (SendType = "1") Then
TLEmail = RecWhichRep.Fields.Item("ASEmail").Value
TLID = RecWhichRep.Fields.Item("ASID").Value
If isEmailValid(TLEmail) = "True" Then
EmailAdd = TLEmail
%><!--#include file="Includes/SendMailer.inc" --><%
End If

set RecZoners = Server.CreateObject("ADODB.Recordset")
RecZoners.ActiveConnection = MM_Site_STRING
RecZoners.Source = "SELECT * FROM Tedis where TediActive = 'True' and ASID = " & TLID & " order by TediFirstName Asc"
RecZoners.CursorType = 0
RecZoners.CursorLocation = 2
RecZoners.LockType = 3
RecZoners.Open()
RecZoners_numRows = 0
While Not RecZoners.EOF
AgentEmail = RecZoners.Fields.Item("TediEmail").Value

If isEmailValid(AgentEmail) = "True" Then
EmailAdd = AgentEmail
%><!--#include file="Includes/SendMailer.inc" --><%
End If

RecZoners.MoveNext
Wend
End If
If (SendType = "2") Then
TLID = RecWhichRep.Fields.Item("ASID").Value
set RecZoners = Server.CreateObject("ADODB.Recordset")
RecZoners.ActiveConnection = MM_Site_STRING
RecZoners.Source = "SELECT * FROM Tedis where TediActive = 'True' and ASID = " & TLID & " order by TediFirstName Asc"
RecZoners.CursorType = 0
RecZoners.CursorLocation = 2
RecZoners.LockType = 3
RecZoners.Open()
RecZoners_numRows = 0
While Not RecZoners.EOF
AgentEmail = RecZoners.Fields.Item("TediEmail").Value

If isEmailValid(AgentEmail) = "True" Then
EmailAdd = AgentEmail
%><!--#include file="Includes/SendMailer.inc" --><%
End If

RecZoners.MoveNext
Wend

End If

If (SendType = "3") Then
TLEmail = RecWhichRep.Fields.Item("ASEmail").Value

If isEmailValid(TLEmail) = "True" Then
EmailAdd = TLEmail
%><!--#include file="Includes/SendMailer.inc" --><%
End If

End If
RecWhichRep.MoveNext
Wend

RecRegionSelect.MoveNext
Wend

End If

%>
<script type="text/javascript">
<!--
function delayer(){
	window.parent.top.location = "Display.asp?AppCat=16&AppSubCatID=32&Success=Email"

}
//-->
</script>
<body onLoad="setTimeout('delayer()', 10)">

</body>