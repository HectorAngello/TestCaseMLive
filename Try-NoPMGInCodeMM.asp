<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%
Response.Buffer = False
%>
<font size="1" Face="Arial">
<% 
Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecFetch = Server.CreateObject("ADODB.Recordset")
RecFetch.ActiveConnection = MM_Site_STRING
RecFetch.Source = "SELECT * FROM MChargeFNBTransMM where Allocated = 'False' Order by FNBDate DESC"
RecFetch.CursorType = 0
RecFetch.CursorLocation = 2
RecFetch.LockType = 3
RecFetch.Open()
RecFetch_numRows = 0
X = 0
While Not RecFetch.EOF
X = X + 1
WhatDS = RecFetch.Fields.Item("TransDescription").Value
WhatDS = Replace(WhatDS, "ADT CASH DEPOSIT", "")
WhatDS = Replace(WhatDS, "ADT CASH DEPO", "")
%>
<%=X%>. <%=RecFetch.Fields.Item("TransDescription").Value%><br>
<%
AllRearSpacesRemoved = "No"

Do while AllRearSpacesRemoved = "No"
If Right(WhatDS,1) = " " Then
WhatDSTT = Len(WhatDS)
WhatDS = Left(WhatDS, WhatDSTT - 1)
Else
AllRearSpacesRemoved = "Yes"
End If
Loop

FindSpace = 0
Response.write("Rear Space Removed and left with: '" & WhatDS & "'<br>")
FindSpace = InStrRev(WhatDS," ")
Response.write("Space found at: '" & FindSpace & "'<br>")
If FindSpace > 0 Then
WhatDSAA = Len(WhatDS)
WhatDS = Right(WhatDS, WhatDSAA - FindSpace)
End If
Response.write("Space and text Removed prior to tail value: '" & WhatDS & "'<br>")

WhatDS2 = WhatDS
%>
<!--#include file="includes/decode.inc" -->

<%
Response.write("<br>Value being searched: " & WhatDS & "<br>")
Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT Top(1) * FROM Tedis Where Replace(TediEmpCode, 'PMG', '') = '" & WhatDS & "' and MobileMoneyTedi = 'True'"
response.write(RecZoner.Source)
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0
'Response.write("<br>" & RecZoner.Source)
F = "<br>Can't Find<br>"
While Not RecZoner.EOF
ZC = RecZoner.Fields.Item("TediEmpCode").Value
ZC = Replace(ZC, " ", "")
'Response.Write("<br>'" & ZC & "' = '" & whatDS & "' ?")
'If ZC = WhatDS Then
AllocateTo = RecZoner.Fields.Item("TID").Value
FNBID = RecFetch.Fields.Item("FNBID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateFNBTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateFNBTable.Open "SELECT Top(1) * FROM MChargeFNBTransMM where FNBID = " & FNBID, MM_Site_STRINGWrite, 1, 2
RecUpdateFNBTable.Update
RecUpdateFNBTable("Allocated") = "True"
RecUpdateFNBTable("TediID") = AllocateTo
RecUpdateFNBTable.Update
RecUpdateFNBTable.Close

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecFNB = Server.CreateObject("ADODB.Recordset")
RecFNB.ActiveConnection = MM_Site_STRING
RecFNB.Source = "SELECT Top(1) FNBID,FNBDate,TransDescription, TransAmount FROM MChargeFNBTransMM where FNBID = " & FNBID
RecFNB.CursorType = 0
RecFNB.CursorLocation = 2
RecFNB.LockType = 3
RecFNB.Open()
RecFNB_numRows = 0

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
Set RecUpdateZTTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateZTTable.Open "SELECT Top(1) * FROM TediTransactionsMM order by CID Desc", MM_Site_STRINGWrite, 1, 2
RecUpdateZTTable.AddNew
RecUpdateZTTable("CAmount") = RecFNB.Fields.Item("TransAmount").Value
RecUpdateZTTable("FNBID") = RecFNB.Fields.Item("FNBID").Value
RecUpdateZTTable("CDate") = RecFNB.Fields.Item("FNBDate").Value
RecUpdateZTTable("TediID") = AllocateTo
RecUpdateZTTable("CType") = "2"
RecUpdateZTTable("CComments") = RecFNB.Fields.Item("TransDescription").Value
RecUpdateZTTable("AddedBy") = Session("UNID")
RecUpdateZTTable.Update
RecUpdateZTTable.Close


If Err.Number <> 0 Then
    Response.Redirect("TryMM.asp")
End If

If AllocateTo <> "0" Then
GZDZID = AllocateTo
%><!--#include file="Includes/GetTediDetail.inc" --><%

MobileNo = ""
MobileNo = GZDZonerCell
Msg = "Dear " & GZDZonerName & ", Mobile Money Deposit R " & FormatNumber(RecFNB.Fields.Item("TransAmount").Value,2) & " - Ref: " & Replace(RecFNB.Fields.Item("TransDescription").Value, "  ","") & " has been allocated to your account (Mobile Money Bal: R" & GZDZonerCurrentMobileMoneyBalance & "), " & GZDZonerCompanyName

SendNotSMS = "Yes"
If SendNotSMS = "Yes" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "1"
RecADDSMS("AlloID") = GZDZID
RecADDSMS("SMSMSG") = Msg
RecADDSMS("MobileNo") = MobileNo
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close
End If

Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & AllocateTo, MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MobileMoneyBalance") = GZDZonerCurrentMChargeBalance
'If IsNull(GZDLastBankedDate) = "True" Then
If GZDLastBankedDate = "" Then
RecADDSMS("LastBankedDateMM") = RecFNB.Fields.Item("FNBDate").Value
Else
Response.write("<br>GZDLastBankedDate: '" & GZDLastBankedDate & "'<br>")
If DateDiff("d",GZDLastBankedDate,RecFNB.Fields.Item("FNBDate").Value) > 0 Then
RecADDSMS("LastBankedDateMM") = RecFNB.Fields.Item("FNBDate").Value
End If
End If

RecADDSMS.Update
RecADDSMS.Close

End If

%>
<%
F = "Found It - " & RecZoner.Fields.Item("TediFirstName").Value & " - " & ZC & " Allocated<br>" & Msg & " - Length=" & Len(Msg) & "<br>"
'End If
RecZoner.MoveNext
Wend
response.Write(F)
RecFetch.MoveNext
Wend
' Check to see if week exists in the system

Week2Day = (DatePart("ww",Date))
Year2Day = Year(Now())

Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
set RecCheckWeek = Server.CreateObject("ADODB.Recordset")
RecCheckWeek.ActiveConnection = MM_Site_STRING
RecCheckWeek.Source = "SELECT * FROM MISWeeks Where MISYear = " & Year2Day & " and MISWeek = " & Week2Day
RecCheckWeek.CursorType = 0
RecCheckWeek.CursorLocation = 2
RecCheckWeek.LockType = 3
RecCheckWeek.Open()
RecCheckWeek_numRows = 0
If Not RecCheckWeek.EOF and Not RecCheckWeek.BOF Then
Else
Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
Set RecADDMISWeek2Sys = Server.CreateObject ( "ADODB.Recordset" )
RecADDMISWeek2Sys.Open "SELECT Top(2) * FROM MISWeeks", MM_Site_STRINGWrite, 1, 2
RecADDMISWeek2Sys.AddNew
RecADDMISWeek2Sys("MISYear") = Year2Day
RecADDMISWeek2Sys("MISWeek") = Week2Day
RecADDMISWeek2Sys.Update
RecADDMISWeek2Sys.Close
End If
%>

<script type="text/javascript">
<!--
function delayer(){
	
	window.location = "scripts/recalctedibalancesMM.asp"

}
//-->
</script>
<body onLoad="setTimeout('delayer()', 100)">
Done</body>

