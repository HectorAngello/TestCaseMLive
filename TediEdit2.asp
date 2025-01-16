<!-- #include file="Connections/Site.asp" -->
<%
RFILE = Request.Form("SRID")
TID = Request.Form("TID")
EmployeeCode = Request.Form("EmployeeCode")
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), "-")
SRID = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
ASID = RFILE

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where  TID = " & TID
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
CurrentSRID = RecEdit.Fields.Item("SRID").Value
CurrentASID = RecEdit.Fields.Item("SRID").Value
If CurrentSRID <> SRID Then
set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT Top(1)* FROM SubRegions Where SRID = " & SRID
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
HCLimit = RecSubRegion.Fields.Item("HeadCountTarget").Value
CurrentCount = 0
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and SRID = " & SRID
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
While Not RecEdit.EOF
CurrentCount = CurrentCount + 1
RecEdit.MoveNext
Wend

If CurrentCount > HCLimit Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! This Sub Region Has Already Reached It's Head Count Target of <%=HCLimit%>");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If
End If
TediFirstName = Request.Form("TediFirstName")
TediLastName = Request.Form("TediLastName")
TediCell = Request.Form("TediCell")
TediEmail = Request.Form("TediEmail")
TediActive = "True"
PurseLimit = Request.Form("PurseLimit")
TediStartDate = Request.Form("TediStartDate")


GenderID = Request.Form("GenderID")
RaceID =  Request.Form("RaceID")
TaxNumber =  Request.Form("TaxNumber")

BankID =  Request.Form("BankID")
BranchCode =  Request.Form("BranchCode")
AccountType =  Request.Form("AccountType")
AccNo =  Replace(Request.Form("AccNo"), " ", "")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM Tedis where (IDNumber = '" & Request.Form("IDnumber") & "' or  TediCell = '" & TediCell & "' or TediEmpCode = '" & EmployeeCode & "') and TID <> " & TID
'Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! A Tedi Already exists in the system, with either the same ID Number, Account Number, Employee Number or Mobile Number");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

If IsNumeric(PurseLimit) = "False" Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Please ensure the purse limit is a numeric value.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

AccNo = Request.Form("AccNo")
AccNo = Trim(AccNo)
AccNo = Replace(AccNo, " ", "")
set RecCheck3 = Server.CreateObject("ADODB.Recordset")
RecCheck3.ActiveConnection = MM_Site_STRING
RecCheck3.Source = "Select * FROM Tedis where (AccNo = '" & AccNo & "') and TID <> " & TID
'Response.Write(RecCheck.Source)
RecCheck3.CursorType = 0
RecCheck3.CursorLocation = 2
RecCheck3.LockType = 3
RecCheck3.Open()
RecCheck3_numRows = 0
If Not RecCheck3.EOF and Not RecCheck3.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! An Agent Already Has The Bank Acc Number Allocated To Them in The System");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("SRID") = SRID
rstSecond("TediFirstName") = Request.Form("TediFirstName")
rstSecond("TediLastName") = Request.Form("TediLastName")
rstSecond("TediCell") = Request.Form("TediCell")
rstSecond("TediCell2") = Request.Form("TediCell2")
rstSecond("TertiaryMobileNumber") = Request.Form("TertiaryMobileNumber")
rstSecond("TediEmail") = Request.Form("TediEmail")
rstSecond("IDNumber") = Request.Form("IDNumber")
rstSecond("TediStartDate") = Request.Form("TediStartDate")
rstSecond("GenderID") = Request.Form("GenderID")
rstSecond("RaceID") = Request.Form("RaceID")
rstSecond("TaxOffice") = Request.Form("TaxOffice")
rstSecond("TaxNumber") = Request.Form("TaxNumber")
rstSecond("BankID") = Request.Form("BankID")
rstSecond("BranchCode") = Request.Form("BranchCode")
rstSecond("AccountType") = Request.Form("AccountType")
rstSecond("AccNo") = AccNo
rstSecond("ResidentialAddress1") = Request.Form("ResidentialAddress1")
rstSecond("ResidentialAddress2") = Request.Form("ResidentialAddress2")
rstSecond("ResidentialAddress3") = Request.Form("ResidentialAddress3")
rstSecond("ResidentialCode") = Request.Form("ResidentialCode")
rstSecond("TediPassword") = Request.Form("TediPassword")
rstSecond("ASID") = ASID
rstSecond("TediEmpCode") = EmployeeCode
rstSecond("TediActive") = Request.Form("AgentActive")
rstSecond("ExcludeFromMchargeBulkFile") = Request.Form("ExcludeFromMchargeBulkFile")
rstSecond("PurseLimit") = Request.Form("PurseLimit")
rstSecond("LastChangedDate") = Now()
rstSecond("RealTimeCommOptIn") = Request.Form("RealTimeCommOptIn")
rstSecond("AirtimeTypeID") = Request.Form("AirtimeTypeID")
rstSecond("MChargeTedi") = Request.Form("MChargeTedi")
rstSecond("MobileMoneyTedi") = Request.Form("MobileMoneyTedi")
rstSecond("PurseLimitMM") = Request.Form("PurseLimitMM")
rstSecond("WorkPermitExpiryDate") = Request.Form("WorkPermitExpiryDate")
rstSecond("DDConsentForm") = Request.Form("DDConsentForm")
rstSecond("DDCrimCheck") = Request.Form("DDCrimCheck")
rstSecond("DDCrimRecord") = Request.Form("DDCrimRecord")
rstSecond("DDAMLTrained") = Request.Form("DDAMLTrained")
rstSecond("DDAMLPassed") = Request.Form("DDAMLPassed")
rstSecond("DDPhoneAllocated") = Request.Form("DDPhoneAllocated")
rstSecond("DDMSISDNAllocated") = Request.Form("DDMSISDNAllocated")
rstSecond("DDTDROboarded") = Request.Form("DDTDROboarded")
rstSecond("DDValidated") = Request.Form("DDValidated")
rstSecond("DDStatusID") = Request.Form("DDStatusID")
rstSecond("MoMoAccNo") = Request.Form("MoMoAccNo")
rstSecond("DDSkhokhoGSM") = Request.Form("DDSkhokhoGSM")
rstSecond("DDSkhokhoDedicated") = Request.Form("DDSkhokhoDedicated")
rstSecond("TradingSpot") = Request.Form("TradingSpot")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

If Request.Form("RealTimeCommOptIn") = "False" Then
set RecInCommOpt = Server.CreateObject("ADODB.Recordset")
RecInCommOpt.ActiveConnection = MM_Site_STRING
RecInCommOpt.Source = "SELECT Top(1)* FROM TediRealTimeCommAllocations Where TediID = " & TID
RecInCommOpt.CursorType = 0
RecInCommOpt.CursorLocation = 2
RecInCommOpt.LockType = 3
RecInCommOpt.Open()
RecInCommOpt_numRows = 0
If not RecInCommOpt.EOF and Not RecInCommOpt.BOF Then

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM TediRealTimeCommAllocations Where ACID = " & RecInCommOpt.Fields.Item("ACID").Value
rstSecond.Open
set rstSecond = nothing	

End If
End If

set RecNewestAgent = Server.CreateObject("ADODB.Recordset")
RecNewestAgent.ActiveConnection = MM_Site_STRING
RecNewestAgent.Source = "SELECT * FROM Tedis Where TID = " & TID
RecNewestAgent.CursorType = 0
RecNewestAgent.CursorLocation = 2
RecNewestAgent.LockType = 3
RecNewestAgent.Open()
RecNewestAgent_numRows = 0
NewestName = RecNewestAgent.Fields.Item("TediFirstName").Value & " " & RecNewestAgent.Fields.Item("TediLastName").Value
NewestID = RecNewestAgent.Fields.Item("TID").Value
ASID = RecNewestAgent.Fields.Item("ASID").Value
DoCodeUpdate = "No"


TediEmpCode = RecNewestAgent.Fields.Item("TediEmpCode").Value

If (Request.Form("MobileMoneyTedi") = "True") and (Left(TediEmpCode,1) <> "M") Then
TediEmpCode = "M" & TediEmpCode
End If

If (Request.Form("MobileMoneyTedi") = "False") and (Left(TediEmpCode,1) = "M") Then
TediEmpCodeT = Len(TediEmpCode)
TediEmpCode = Right(TediEmpCode, TediEmpCodeT - 1)
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis where TID = " & NewestID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TediEmpCode") = TediEmpCode
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


TediUpdateType = Request.Form("UpdateReason")
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%
Response.Redirect("TediView.asp?TID=" & TID & "&TediUpdated=True")
%>