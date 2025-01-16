<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")
set RecNewestAgent2 = Server.CreateObject("ADODB.Recordset")
RecNewestAgent2.ActiveConnection = MM_Site_STRING
RecNewestAgent2.Source = "SELECT Top(1)* FROM Tedis Order By TID Desc"
RecNewestAgent2.CursorType = 0
RecNewestAgent2.CursorLocation = 2
RecNewestAgent2.LockType = 3
RecNewestAgent2.Open()
RecNewestAgent2_numRows = 0
NextEmpCode = Replace(RecNewestAgent2.Fields.Item("TediEmpCode").Value, "MPMG", "")
NextEmpCode = Replace(NextEmpCode, "PMG", "")
NextEmpCode = Int(NextEmpCode) + 1
If Request.Form("TediParent") = "0" then
SRID = Request.Form("SRID")
Else



set RecCheckTedi = Server.CreateObject("ADODB.Recordset")
RecCheckTedi.ActiveConnection = MM_Site_STRING
RecCheckTedi.Source = "SELECT * FROM ViewTediDetail where TID = " & Request.Form("TediParent")
'Response.write(RecRegion.Source)
RecCheckTedi.CursorType = 0
RecCheckTedi.CursorLocation = 2
RecCheckTedi.LockType = 3
RecCheckTedi.Open()
RecCheckTedi_numRows = 0
SRID = RecCheckTedi.Fields.Item("SRID").Value
End If
set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT * FROM SubRegions Where SRID = " & SRID
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
HCLimit = RecSubRegion.Fields.Item("HeadCountTarget").Value
CurrentCount = 0
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and TediParent = '0' and SRID = " & SRID
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
While Not RecEdit.EOF
CurrentCount = CurrentCount + 1
RecEdit.MoveNext
Wend

If CurrentCount >= HCLimit Then
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


EmployeeCode = Request.Form("EmployeeCode")
TediFirstName = Request.Form("TediFirstName")
TediLastName = Request.Form("TediLastName")
TediCell = Request.Form("TediCell")
TediCell2 = Request.Form("TediCell2")
TediEmail = Request.Form("TediEmail")
TediActive = "True"
PurseLimit = Request.Form("PurseLimit")
TediStartDate = Request.Form("TediStartDate")
TediParent = Request.Form("TediParent")

If TediParent <> "0" Then
set RecTediParent = Server.CreateObject("ADODB.Recordset")
RecTediParent.ActiveConnection = MM_Site_STRING
RecTediParent.Source = "SELECT * FROM Tedis where TID = " & TediParent
RecTediParent.CursorType = 0
RecTediParent.CursorLocation = 2
RecTediParent.LockType = 3
RecTediParent.Open()
RecTediParent_numRows = 0
ASID = RecTediParent.Fields.Item("ASID").Value
SRID = RecTediParent.Fields.Item("SRID").Value
End If
GenderID = Request.Form("GenderID")
RaceID =  Request.Form("RaceID")
TaxNumber =  Request.Form("TaxNumber")

BankID =  Request.Form("BankID")
BranchCode =  Request.Form("BranchCode")
AccountType =  Request.Form("AccountType")
AccNo =  Replace(Request.Form("AccNo"), " ", "")
IDNumberT = Trim(Request.Form("IDNumber"))

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM Tedis where (IDNumber = '" & IDNumberT & "' or  TediCell = '" & TediCell & "')"
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
      window.alert ("Error ! An Agent Already exists in the system, with either the same ID Number or Mobile Number");
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
RecCheck3.Source = "Select * FROM Tedis where (AccNo = '" & AccNo & "')"
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


PMGTempCode = "PMG"
If Request.Form("MobileMoneyTedi") = "True" Then
PMGTempCode = "MPMG"
End If

'If EmployeeCode = "Generate" Then
TediEmpCode = PMGTempCode & NextEmpCode
'Else
'TediEmpCode = EmployeeCode
'End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("SRID") = SRID
rstSecond("TediEmpCode") = TediEmpCode
rstSecond("TediFirstName") = Request.Form("TediFirstName")
rstSecond("TediLastName") = Request.Form("TediLastName")
rstSecond("TediCell") = Trim(Request.Form("TediCell"))
rstSecond("TediCell2") = Trim(Request.Form("TediCell2"))
rstSecond("TertiaryMobileNumber") = Request.Form("TertiaryMobileNumber")
rstSecond("TediEmail") = Request.Form("TediEmail")
rstSecond("IDNumber") = IDNumberT
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
rstSecond("TediParent") = TediParent
rstSecond("ExcludeFromMchargeBulkFile") = Request.Form("ExcludeFromMchargeBulkFile")
rstSecond("PurseLimit") = Request.Form("PurseLimit")
rstSecond("AirtimeTypeID") = Request.Form("AirtimeTypeID")
rstSecond("LastChangedDate") = Now()
rstSecond("RealTimeCommOptIn") = Request.Form("RealTimeCommOptIn")
rstSecond("MChargeTedi") = Request.Form("MChargeTedi")
rstSecond("MobileMoneyTedi") = Request.Form("MobileMoneyTedi")
rstSecond("PurseLimitMM") = Request.Form("PurseLimitMM")
rstSecond("WorkPermitExpiryDate") = Request.Form("WorkPermitExpiryDate")
rstSecond("DDSkhokhoDedicated") = Request.Form("DDSkhokhoDedicated")
rstSecond("DDSkhokhoGSM") = Request.Form("DDSkhokhoGSM")
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
rstSecond("TradingSpot") = Request.Form("TradingSpot")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing



set RecNewestAgent = Server.CreateObject("ADODB.Recordset")
RecNewestAgent.ActiveConnection = MM_Site_STRING
RecNewestAgent.Source = "SELECT * FROM Tedis where IDNumber = '" & IDNumberT & "' Order By TID Desc"
RecNewestAgent.CursorType = 0
RecNewestAgent.CursorLocation = 2
RecNewestAgent.LockType = 3
RecNewestAgent.Open()
RecNewestAgent_numRows = 0
NewestName = RecNewestAgent.Fields.Item("TediFirstName").Value & " " & RecNewestAgent.Fields.Item("TediLastName").Value
NewestID = RecNewestAgent.Fields.Item("TID").Value
TediUpdateID = RecNewestAgent.Fields.Item("TID").Value
ASID = RecNewestAgent.Fields.Item("ASID").Value


set RecAS = Server.CreateObject("ADODB.Recordset")
RecAS.ActiveConnection = MM_Site_STRING
RecAS.Source = "SELECT * FROM ASs Where ASID = " & ASID
RecAS.CursorType = 0
RecAS.CursorLocation = 2
RecAS.LockType = 3
RecAS.Open()
RecAS_numRows = 0

set RecReg = Server.CreateObject("ADODB.Recordset")
RecReg.ActiveConnection = MM_Site_STRING
RecReg.Source = "SELECT * FROM ViewRegionsDetail Where RID = " & RecAS.Fields.Item("RID").Value
RecReg.CursorType = 0
RecReg.CursorLocation = 2
RecReg.LockType = 3
RecReg.Open()
RecReg_numRows = 0

set RecSubReg = Server.CreateObject("ADODB.Recordset")
RecSubReg.ActiveConnection = MM_Site_STRING
RecSubReg.Source = "SELECT * FROM SubRegions Where SRID = " & SRID
RecSubReg.CursorType = 0
RecSubReg.CursorLocation = 2
RecSubReg.LockType = 3
RecSubReg.Open()
RecSubReg_numRows = 0




%><!-- #include file="Includes/TediAudit-Add.inc" -->
<%
Response.Redirect("TediAdd.asp?TediAdded=True&TediName=" & NewestName & "&TediEmpCode=" & TediEmpCode)
%>