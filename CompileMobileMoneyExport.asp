<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<font size="2" face="Arial">
<%
Response.Expires = 1
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
FileBody = ""
TotalFileValue = 0

SendMRSMS = "No"

'response.buffer = false

UN = Request.Querystring("UN")

RunReport = "Yes"
RunCounter = 1
RepList = ""
Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecRunReport = Server.CreateObject("ADODB.Recordset")
RecRunReport.ActiveConnection = MM_Site_STRING
RecRunReport.Source = "SELECT * FROM BulkMChargeMM Where ((BulkStatus = 'Processing') or (BulkStatus = 'Correcting'))"
RecRunReport.CursorType = 0
RecRunReport.CursorLocation = 2
RecRunReport.LockType = 3
RecRunReport.Open()
RecRunReport_numRows = 0
While Not RecRunReport.EOF
RepList = RepList & " : " & RecRunReport.Fields.Item("BulkID").Value
If RunCounter > 1 Then
RunReport = "No"
End If
RunCounter = RunCounter + 1
RecRunReport.MoveNext
WEnd

If RunReport = "No" Then
Response.Write("Waiting For MCharge BulkID" & RepList)
Else
Set conMain = Server.CreateObject ( "ADODB.Connection" )

set RecNewestBulk = Server.CreateObject("ADODB.Recordset")
RecNewestBulk.ActiveConnection = MM_Site_STRING
RecNewestBulk.Source = "SELECT * FROM BulkMChargeMM Where BulkStatus = 'Pending' Order By BulkID Asc"
RecNewestBulk.CursorType = 0
RecNewestBulk.CursorLocation = 2
RecNewestBulk.LockType = 3
RecNewestBulk.Open()
RecNewestBulk_numRows = 0
If RecNewestBulk.EOF and RecNewestBulk.BOF Then
Response.Write("Nothing To Process ;)")
Else
NewestBulkID = RecNewestBulk.Fields.Item("BulkID").Value


FileGenBy = RecNewestBulk.Fields.Item("UserID").Value

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeMM where BulkID = " & NewestBulkID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("BulkStatus") = "Processing"
		RecUpdateBulk.Update
		RecUpdateBulk.Close

set RecAlloocationType = Server.CreateObject("ADODB.Recordset")
RecAlloocationType.ActiveConnection = MM_Site_STRING
RecAlloocationType.Source = "SELECT Top(1)* FROM AirtimeAllocationTypesMM Where AirtimeAlloActive = 'True' Order By AirtimeAlloLabel Asc"
RecAlloocationType.CursorType = 0
RecAlloocationType.CursorLocation = 2
RecAlloocationType.LockType = 3
RecAlloocationType.Open()
RecAlloocationType_numRows = 0
While Not RecAlloocationType.EOF
FileExt = ".csv"
SFile = RecNewestBulk.Fields.Item("FileName").Value
AirtimeTypeID = RecAlloocationType.Fields.Item("AirtimeTypeID").Value

RegionStopper = "No"
TheFilePath=(AppPath & "MChargeBulkFiles\" & SFile & FileExt)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
'************ beginning of the file body ***********
FileTopLine = "HDR,Payments,XXXTransCount,XXXTransValue"

If AirtimeTypeID = 1 Then
'TheFile.Writeline("H,1,N" & NewestBulkID & ".BulkTransfer," & SFile & ",Yes,Yes")
%>
H,1,<%=NewestBulkID%>,<%=SFile%>,Yes,Yes<br>
<%
End If 

RegCount = 0
ZCount = 0

RegCount = RegCount + 1
Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecReps = Server.CreateObject("ADODB.Recordset")
RecReps.ActiveConnection = MM_Site_STRING
RecReps.Source = "SELECT * FROM BulkMChargeTediTempMM Where BulkID = '" & NewestBulkID & "'"
RecReps.CursorType = 0
RecReps.CursorLocation = 2
RecReps.LockType = 3
RecReps.Open()
RecReps_numRows = 0
RepCount = 0
While Not RecReps.EOF
TmpID = RecReps.Fields.Item("TmpID").Value
RepCount = RepCount + 1
Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ViewTediDetail Where TediActive = 'True' and TID = " & RecReps.Fields.Item("TID").Value & " order by TediFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecRegions = Server.CreateObject("ADODB.Recordset")
RecRegions.ActiveConnection = MM_Site_STRING
RecRegions.Source = "SELECT Top(1)* FROM Regions where RID = '" & RecCurrent.Fields.Item("RID").Value & "'"
RecRegions.CursorType = 0
RecRegions.CursorLocation = 2
RecRegions.LockType = 3
RecRegions.Open()
RecRegions_numRows = 0
If Not RecCurrent.EOF And Not RecCurrent.BOF Then

ZonerPurseLimit = RecCurrent.Fields.Item("PurseLimitMM").Value


DC = 0
ZTransTotal = 0
CreditAmount = ZonerPurseLimit
' New Calc Starts
TediNewBanked = 0
TediNewMCharge = 0
TID = RecCurrent.Fields.Item("TID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout

set RecNewCalcZonerMCharge = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerMCharge.ActiveConnection = MM_Site_STRING
RecNewCalcZonerMCharge.Source = "SELECT Top(1)* FROM ViewMchargeTediTotalAllocatedMM Where TediID = " & TID
RecNewCalcZonerMCharge.CursorType = 0
RecNewCalcZonerMCharge.CursorLocation = 2
RecNewCalcZonerMCharge.LockType = 3
RecNewCalcZonerMCharge.Open()
RecNewCalcZonerMCharge_numRows = 0
If Not RecNewCalcZonerMCharge.EOF Then
TediNewMCharge = RecNewCalcZonerMCharge.Fields.Item("TediTotalAllocatedMM").Value
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
conMain.ConnectionTimeout = DBConTimeout
conMain.CommandTimeout = DBCommandTimeout
set RecNewCalcZonerBanked = Server.CreateObject("ADODB.Recordset")
RecNewCalcZonerBanked.ActiveConnection = MM_Site_STRING
RecNewCalcZonerBanked.Source = "SELECT Top(1)* FROM ViewMchargeTediTotalBankedMM Where TediID = " & TID
RecNewCalcZonerBanked.CursorType = 0
RecNewCalcZonerBanked.CursorLocation = 2
RecNewCalcZonerBanked.LockType = 3
RecNewCalcZonerBanked.Open()
RecNewCalcZonerBanked_numRows = 0
If Not RecNewCalcZonerBanked.EOF Then
TediNewBanked = RecNewCalcZonerBanked.Fields.Item("TediTotalBankedMM").Value
End If

ZTransTotal = (TediNewMCharge - TediNewBanked)

AirtimeAllocation = 0
If ZTransTotal = "0" then
DC = 0
Else 
DC = ZTransTotal
End If
CreditAmount = DC - ZonerPurseLimit
'End If
ZTransTotal = "R " & FormatNumber(ZTransTotal,2)
ZonerTempMobile = RecCurrent.Fields.Item("TediCell2").Value

If IsNull(RecCurrent.Fields.Item("TediCell2").Value) = "True" Then
ZonerTempMobile = RecCurrent.Fields.Item("TediCell").Value
End If

ZonerMobile = Replace(ZonerTempMobile, " ", "")
ZonerMobile = "27" & Right(ZonerMobile,9)
RegionStopper = "Yes"
AirtimeAllocation = Int(ZonerPurseLimit - DC)
Response.write("<br>TID: " & TID & " | PurseLimit:" & ZonerPurseLimit & " | Current Balance: " & ZTransTotal & " | Top Up Value: " & AirtimeAllocation)
If Int(AirtimeAllocation) >= Int(MinTopUpMM) Then
ZCount = ZCount + 1
TotalFileValue = TotalFileValue + AirtimeAllocation
		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecInsert2 = Server.CreateObject ( "ADODB.Recordset" )
		RecInsert2.Open "SELECT Top(1)* FROM BulkMChargeChildrenMM order by ChildID Desc", MM_Site_STRINGWrite, 1, 2
		RecInsert2.AddNew
		RecInsert2("BulkID") = NewestBulkID
		RecInsert2("ChildCreationDate") = Now()
		RecInsert2("TID") = RecCurrent.Fields.Item("TID").Value
		RecInsert2("TediMSISDN") = ZonerMobile
		RecInsert2("ValBefore") = DC
		RecInsert2("ValAfter") = DC - CreditAmount
		RecInsert2("MChargeAmount") = AirtimeAllocation
		RecInsert2.Update
		RecInsert2.Close

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecAllocateMCharge = Server.CreateObject ( "ADODB.Recordset" )
		RecAllocateMCharge.Open "SELECT Top(1)* FROM TediTransactionsMM", MM_Site_STRINGWrite, 1, 2
		RecAllocateMCharge.AddNew
		RecAllocateMCharge("CAmount") = AirtimeAllocation
		RecAllocateMCharge("CDate") = Now()
		RecAllocateMCharge("TediID") = RecCurrent.Fields.Item("TID").Value
		RecAllocateMCharge("CType") = "1"
		RecAllocateMCharge("CComments") = "Bulk Update: " & SFile
		RecAllocateMCharge("AddedBy") = UN
		RecAllocateMCharge.Update
		RecAllocateMCharge.Close

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		conMain.ConnectionTimeout = DBConTimeout
		conMain.CommandTimeout = DBCommandTimeout
		Set RecAllocateMCharge = Server.CreateObject ( "ADODB.Recordset" )
		RecAllocateMCharge.Open "SELECT Top(1)* FROM BulkMChargeTediTempMM where TmpID = " & TmpID, MM_Site_STRINGWrite, 1, 2
		RecAllocateMCharge.Update
		RecAllocateMCharge("AllocatedValue") = AirtimeAllocation
		RecAllocateMCharge.Update
		RecAllocateMCharge.Close

' Check Me Start
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM Tedis Where TID = " & RecCurrent.Fields.Item("TID").Value, MM_Site_STRINGWrite, 1, 2
RecADDSMS.Update
RecADDSMS("MobileMoneyBalance") = DC - CreditAmount
RecADDSMS.Update
RecADDSMS.Close

' Check Me End
CheckXXX = Int(ZonerPurseLimit - DC)
If CheckXXX <> 0 Then
If AirtimeTypeID = 1 Then
FileBody = FileBody & ZCount & ",MSISDN," & ZonerMobile & "," & Int(ZonerPurseLimit - DC) & "," & RecCurrent.Fields.Item("TediEmpCode").Value & vbCrLf
'TheFile.Writeline("D,," & ZonerMobile & ",," & Int(ZonerPurseLimit - DC) & ",1,1") 
End If
%>D,,<%=ZonerMobile%>,,<%=Int(ZonerPurseLimit - DC)%>,,<%=RecCurrent.Fields.Item("TID").Value%><br>
<%
End If
End If
End If


Response.flush
RecReps.MoveNext
Wend

If AirtimeTypeID = 1 Then
'TheFile.Writeline("T," & ZCount)
FileTopLine = Replace(FileTopLine, "XXXTransCount", ZCount)
FileTopLine = Replace(FileTopLine, "XXXTransValue", TotalFileValue)
TheFile.Writeline(FileTopLine)
TheFile.Writeline(FileBody)
%>
T,<%=ZCount%>
<%
End If

'************ end of the file body *****************
TheFile.close
Set FSO = nothing

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeMM where BulkID = " & NewestBulkID, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("BulkStatus") = "Complete"
		RecUpdateBulk.Update
		RecUpdateBulk.Close

If SendMRSMS = "Yes" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecFetchSMS = Server.CreateObject("ADODB.Recordset")
RecFetchSMS.ActiveConnection = MM_Site_STRING
RecFetchSMS.Source = "SELECT * FROM viewUnsentBulkSMSMM Where BulkID = " & NewestBulkID
RecFetchSMS.CursorType = 0
RecFetchSMS.CursorLocation = 2
RecFetchSMS.LockType = 3
RecFetchSMS.Open()
RecFetchSMS_numRows = 0
While Not RecFetchSMS.EOF

GZDZID = RecFetchSMS.Fields.Item("TID").Value
%><!--#include file="Includes/GetTediDetail.inc" --><%
MobileNo = RecFetchSMS.Fields.Item("ASCell").Value
ASID = RecFetchSMS.Fields.Item("ASID").Value

Msg = "R " & FormatNumber(RecFetchSMS.Fields.Item("MChargeAmount").Value,2) & " allocated to Mobile Money ISA " & RecFetchSMS.Fields.Item("TediEmpCode").Value & " - " & RecFetchSMS.Fields.Item("TediFirstName").Value & " " & RecFetchSMS.Fields.Item("TediLastName").Value & " By Administrator: " & RecFetchSMS.Fields.Item("UserFirstname").Value & " (" & RecFetchSMS.Fields.Item("TediEmpCode").Value & " Mobile Money Bal: R" & GZDZonerCurrentMobileMoneyBalance & "), Regards PMG"



Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "2"
RecADDSMS("AlloID") = ASID
RecADDSMS("SMSMSG") = Msg
RecADDSMS("MobileNo") = MobileNo
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close

RecFetchSMS.MoveNext

Response.Write("<br>" & Msg & " - Length: " & Len(Msg))
Wend

Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecBulk2 = Server.CreateObject("ADODB.Recordset")
RecBulk2.ActiveConnection = MM_Site_STRING
RecBulk2.Source = "SELECT Top(1)* FROM BulkMChargeChildrenMM Where BulkID = " & NewestBulkID
RecBulk2.CursorType = 0
RecBulk2.CursorLocation = 2
RecBulk2.LockType = 3
RecBulk2.Open()
RecBulk2_numRows = 0

If Not RecBulk2.EOF and not RecBulk2.BOF Then

		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecUpdateBulk = Server.CreateObject ( "ADODB.Recordset" )
		RecUpdateBulk.Open "SELECT Top(1)* FROM BulkMChargeChildrenMM Where ChildID = " & RecBulk2.Fields.Item("ChildID").Value, MM_Site_STRINGWrite, 1, 2
		RecUpdateBulk.Update
		RecUpdateBulk("SMSSent") = "True"
		RecUpdateBulk.Update
		RecUpdateBulk.Close

End if
End If
RecAlloocationType.MoveNext
Wend

End If
End If
response.end
%>