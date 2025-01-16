<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")
RID = Request.Form("RID")
ASFirstName = Request.Form("ASFirstName")
ASLastName = Request.Form("ASLastName")
ASCell = Request.Form("ASCell")
ASEmail = Request.Form("ASEmail")
ASActive = Request.Form("ASActive")
EmployeeCode = Request.Form("EmployeeCode")
ASStartDate = Request.Form("ASStartDate")

set RecCheckRegion = Server.CreateObject("ADODB.Recordset")
RecCheckRegion.ActiveConnection = MM_Site_STRING
RecCheckRegion.Source = "SELECT * FROM ASs Where ASID = " & ASID
RecCheckRegion.CursorType = 0
RecCheckRegion.CursorLocation = 2
RecCheckRegion.LockType = 3
RecCheckRegion.Open()
RecCheckRegion_numRows = 0
RegCodecheck = RecCheckRegion.Fields.Item("RID").Value


GenderID = Request.Form("GenderID")
RaceID =  Request.Form("RaceID")
TaxNumber =  Request.Form("TaxNumber")

BankID =  Request.Form("BankID")
BranchCode =  Request.Form("BranchCode")
AccountType =  Request.Form("AccountType")
AccNo =  Request.Form("AccNo")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM ASs where (IDNumber = '" & Request.Form("IDnumber") & "' or ASEmail = '" & ASEmail & "' or ASEmpCode = '" & EmployeeCode & "') and ASID <> " & ASID
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
      window.alert ("Error ! A <%=SupervisorLabel%> Already exists in the system, with either the same ID Number, Employee Number or Mobile Number");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ASs Where ASID = " & ASID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("RID") = Request.Form("RID")
rstSecond("ASFirstName") = Request.Form("ASFirstName")
rstSecond("ASLastName") = Request.Form("ASLastName")
rstSecond("ASCell") = Request.Form("ASCell")
rstSecond("ASEmail") = Request.Form("ASEmail")
rstSecond("IDNumber") = Request.Form("IDNumber")
rstSecond("ASStartDate") = Request.Form("ASStartDate")
rstSecond("GenderID") = Request.Form("GenderID")
rstSecond("RaceID") = Request.Form("RaceID")
rstSecond("TaxOffice") = Request.Form("TaxOffice")
rstSecond("TaxNumber") = Request.Form("TaxNumber")
rstSecond("BankID") = Request.Form("BankID")
rstSecond("BranchCode") = Request.Form("BranchCode")
rstSecond("AccountType") = Request.Form("AccountType")
rstSecond("AccNo") = Request.Form("AccNo")
rstSecond("ResidentialAddress1") = Request.Form("ResidentialAddress1")
rstSecond("ResidentialAddress2") = Request.Form("ResidentialAddress2")
rstSecond("ResidentialAddress3") = Request.Form("ResidentialAddress3")
rstSecond("ResidentialCode") = Request.Form("ResidentialCode")
rstSecond("ASPassword") = Request.Form("ASPassword")
rstSecond("ASActive") = ASActive
rstSecond("ASEmpCode") = EmployeeCode
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing



set RecNewestAgent = Server.CreateObject("ADODB.Recordset")
RecNewestAgent.ActiveConnection = MM_Site_STRING
RecNewestAgent.Source = "SELECT * FROM ASs Where ASID = " & ASID
RecNewestAgent.CursorType = 0
RecNewestAgent.CursorLocation = 2
RecNewestAgent.LockType = 3
RecNewestAgent.Open()
RecNewestAgent_numRows = 0
NewestName = RecNewestAgent.Fields.Item("ASFirstName").Value & " " & RecNewestAgent.Fields.Item("ASLastName").Value
NewestID = RecNewestAgent.Fields.Item("ASID").Value
ASID = RecNewestAgent.Fields.Item("ASID").Value
RegCode = RecNewestAgent.Fields.Item("RID").Value
ASEmpCode = RecNewestAgent.Fields.Item("ASEmpCode").Value
DoCodeChange= "No"
If DoCodeChange= "Yes" Then
If Request.Form("RID") <> RegCodecheck Then
If Len(NewestID) = "1" Then
NewestID = "000" & NewestID
End If
If Len(NewestID) = "2" Then
NewestID = "00" & NewestID
End If
If Len(NewestID) = "3" Then
NewestID = "0" & NewestID
End If
TLEmpCode = RegCode & "TL" & NewestID


set RecTeamLeader = Server.CreateObject("ADODB.Recordset")
RecTeamLeader.ActiveConnection = MM_Site_STRING
RecTeamLeader.Source = "SELECT * FROM ViewRegionsDetail Where RID = " & RID
RecTeamLeader.CursorType = 0
RecTeamLeader.CursorLocation = 2
RecTeamLeader.LockType = 3
RecTeamLeader.Open()
RecTeamLeader_numRows = 0

ASEmpCode = RecTeamLeader.Fields.Item("CompanyAbb").Value & RecTeamLeader.Fields.Item("RegionCode").Value & NewestID

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ASs where ASID = " & ASID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("ASEmpCode") = ASEmpCode
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

End If
End If
ASUpDateID = ASID
ASUpdateType = SupervisorLabel & " Updated"
%><!-- #include file="Includes/ASAudit-Update.inc" -->
<%




Response.Redirect("ASEdit.asp?ASID=" & ASID & "&ASUpdated=True&ASName=" & NewestName & "&ASEmpCode=" & ASEmpCode)

%>