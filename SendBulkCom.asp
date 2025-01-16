<!-- #include file="Connections/Site.asp" -->
<%
Region = Request.Form("Region")
SendType = Request.Form("SendType")
ComType= Request.Form("ComType")

Function isEmailValid(email) 
        Set regEx = New RegExp 
        regEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$" 
        isEmailValid = regEx.Test(trim(email)) 
    End Function 

	Function testEmail(email) 
        response.write "<p>" & email & " (" & _ 
            isEmailValid(email) & ")" 
    End Function 

If ComType = "SMS" Then
MSG = Request.Form("MSG")

RFILE = Request.Form("Region")
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), "--")
Region = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
RepID = Replace(RFILE, "-", "")
SendType = Request.Form("SendType")
DisplaySendText = ""
If SendType = "1" Then
DisplaySendText = "All Tedis And EDI Mentors"
End If
If SendType = "2" Then
DisplaySendText = "ONLY Tedis"
End If
If SendType = "3" Then
DisplaySendText = "ONLY EDI Mentors"
End If

set RecRegionSelect = Server.CreateObject("ADODB.Recordset")
RecRegionSelect.ActiveConnection = MM_Site_STRING
If Region = "0" then
RecRegionSelect.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
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
RecWhichRep.Source = "SELECT Top(1)* FROM ASs Where ASID = " & RepID
End If
Response.write(RecWhichRep.Source)
RecWhichRep.CursorType = 0
RecWhichRep.CursorLocation = 2
RecWhichRep.LockType = 3
RecWhichRep.Open()
RecWhichRep_numRows = 0
RepCount = 0
While Not RecWhichRep.EOF
If (SendType = "1") Then
TLID = RecWhichRep.Fields.Item("ASID").Value
RepNo = Replace(RecWhichRep.Fields.Item("ASCell").Value, " ", "")
If IsNumeric(RepNo) = "True" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("UserType") = "2"
rstChangeUpdate("AlloID") = TLID
rstChangeUpdate("SMSMsg") = MSG
rstChangeUpdate("MobileNo") = RepNo
rstChangeUpdate("SMSDate") = Now()
rstChangeUpdate("IsSent") = "False"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing
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
AgentID = RecZoners.Fields.Item("TID").Value
AgentCell = Replace(RecZoners.Fields.Item("TediCell").Value, " ", "")
If IsNumeric(AgentCell) = "True" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("UserType") = "1"
rstChangeUpdate("AlloID") = AgentID 
rstChangeUpdate("SMSMsg") = MSG
rstChangeUpdate("MobileNo") = AgentCell
rstChangeUpdate("SMSDate") = Now()
rstChangeUpdate("IsSent") = "False"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing
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
AgentID = RecZoners.Fields.Item("TID").Value
AgentCell = Replace(RecZoners.Fields.Item("TediCell").Value, " ", "")
If IsNumeric(AgentCell) = "True" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("UserType") = "1"
rstChangeUpdate("AlloID") = AgentID 
rstChangeUpdate("SMSMsg") = MSG
rstChangeUpdate("MobileNo") = AgentCell
rstChangeUpdate("SMSDate") = Now()
rstChangeUpdate("IsSent") = "False"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing
End If
RecZoners.MoveNext
Wend

End If

If (SendType = "3") Then
TLID = RecWhichRep.Fields.Item("ASID").Value
RepNo = Replace(RecWhichRep.Fields.Item("ASCell").Value, " ", "")
If Isnumeric(RepNo) = "True" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("UserType") = "2"
rstChangeUpdate("AlloID") = TLID
rstChangeUpdate("SMSMsg") = MSG
rstChangeUpdate("MobileNo") = RepNo
rstChangeUpdate("SMSDate") = Now()
rstChangeUpdate("IsSent") = "False"
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing
End If
End If
RecWhichRep.MoveNext
Wend

RecRegionSelect.MoveNext
Wend


response.redirect("Display.asp?AppCat=16&AppSubCatID=32&Success=SMS")
End If

%>