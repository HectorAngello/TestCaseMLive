<%
Region = Request.QueryString("Region")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
OutFormat = Request.QueryString("OutFormat")
If OutFormat <> "P" Then
%>
<!-- #include file="includes/header.asp" -->
<%
Else
%><!--#include file="Connections/Site.asp" -->
<%
End If
LC = 0
If Region = "0" then
WR = "All My Regions"
Else
set RecWR = Server.CreateObject("ADODB.Recordset")
RecWR.ActiveConnection = MM_Site_STRING
RecWR.Source = "SELECT * FROM [Regions] Where RID = " & Region
RecWR.CursorType = 0
RecWR.CursorLocation = 2
RecWR.LockType = 3
RecWR.Open()
RecWR_numRows = 0
WR = RecWR.Fields.Item("RegionName").Value
End If

SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID")

If Region = "0" then
Else
SubRegionQry = SubRegionQry & " and RID = " & Region
End If

'response.write(SubRegionQry)
set RecRegions = Server.CreateObject("ADODB.Recordset")
RecRegions.ActiveConnection = MM_Site_STRING
RecRegions.Source = SubRegionQry
RecRegions.CursorType = 0
RecRegions.CursorLocation = 2
RecRegions.LockType = 3
RecRegions.Open()
RecRegions_numRows = 0
While Not RecRegions.EOF
SRRegionList = SRRegionList & RecRegions.Fields.Item("SRID").Value & ","
RecRegions.MoveNext
Wend
TempLenSRRegionList = Len(SRRegionList)
SRRegionList = Left(SRRegionList,TempLenSRRegionList - 1)

If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "New_Hire_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Agent Code,  First Name,Last Name ,Gender ,Race ,ID Number ,Tax Office ,Sars Tax Ref, Start Date ,Residential Address, Mobile No., Agent Type, Bank, Branch, Acc Type, Acc No, Sub Region, Airtime Deposits, Gross Connections"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Agent New Hired Report</h3>
<p>Date Range: <b><%=StartDate%>&nbsp;to&nbsp;<%=EndDate%></b>
<br>Region: <b><%=WR%></b>
<table>
<thead>
<tr>
	<th>Agent Code</th>
	<th>First Name</th>
	<th>Last Name</th>
	<th>Gender</th>
	<th>Race</th>
	<th>ID Number</th>
	<th>Tax Office</th>
	<th>Sars Tax Ref</th>
	<th>Start Date</th>
	<th>Residential Address</th>
	<th>Mobile No.</th>
	<th>Agent Type</th>
	<th>Bank</th>
	<th>Branch</th>
	<th>Acc Type</th>
	<th>Acc No</th>
	<th>Sub Region</th>
	<th>Airtime Deposits</th>
	<th>Gross Connections</th>
</tr>
</thead>

<tbody>
<%
End If

AgentSQl = "SELECT * FROM ViewTediDetail where TediActive = 'True' "

AgentSQL = AgentSQL & " and SRID in (" & SRRegionList & ")"

AgentSQl = AgentSQl & " and (TediStartDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"

AgentSQl = AgentSQl & " Order By TediEmpCode Asc"

'Response.write(AgentSQl)

set RecAgentEdit = Server.CreateObject("ADODB.Recordset")
RecAgentEdit.ActiveConnection = MM_Site_STRING
RecAgentEdit.Source = AgentSQl
RecAgentEdit.CursorType = 0
RecAgentEdit.CursorLocation = 2
RecAgentEdit.LockType = 3
RecAgentEdit.Open()
RecAgentEdit_numRows = 0
While Not RecAgentEdit.EOF
TID = RecAgentEdit.Fields.Item("TID").Value
LC = LC + 1
AirtimeDeposits = 0
set RecTotalAgents = Server.CreateObject("ADODB.Recordset")
RecTotalAgents.ActiveConnection = MM_Site_STRING
RecTotalAgents.Source = "SELECT Sum(CAmount) AS AgentTotal FROM viewTediTransactions WHERE TID = " & RecAgentEdit.Fields.Item("TID").Value & " and (CDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59') and CType = '2'"
'Response.Write(RecFNB.Source)
RecTotalAgents.CursorType = 0
RecTotalAgents.CursorLocation = 2
RecTotalAgents.LockType = 3
RecTotalAgents.Open()
RecTotalAgents_numRows = 0
If IsNull(RecTotalAgents.Fields.Item("AgentTotal").Value) = false then
AirtimeDeposits = RecTotalAgents.Fields.Item("AgentTotal").Value
End If

If IsNull(RecAgentEdit.Fields.Item("TediStartDate").Value) = "True" Then
AgentStartDate = "N/A"
Else
ASDDay = Day(RecAgentEdit.Fields.Item("TediStartDate").Value)
If Len(ASDDay) = 1 Then
ASDDay = "0" & ASDDay
End If
ASDMonth = Month(RecAgentEdit.Fields.Item("TediStartDate").Value)
If Len(ASDMonth) = 1 Then
ASDMonth = "0" & ASDMonth
End If
AgentStartDate = ASDDay & "/" & ASDMonth & "/" & Year(RecAgentEdit.Fields.Item("TediStartDate").Value)
End If
TediType = "Agent"
If RecAgentEdit.Fields.Item("TediParent").Value <> 0 Then
TediType = "Sub-Agent"
End If

GrossConnections = 0
set RecTotalConnections = Server.CreateObject("ADODB.Recordset")
RecTotalConnections.ActiveConnection = MM_Site_STRING
RecTotalConnections.Source = "SELECT Count(ActID) AS ConTotal FROM SimActivations WHERE TID = " & RecAgentEdit.Fields.Item("TID").Value & " and (ActivationDate BETWEEN '" & StartDate & "' AND '" & EndDate & " 23:59:59')"
RecTotalConnections.CursorType = 0
RecTotalConnections.CursorLocation = 2
RecTotalConnections.LockType = 3
RecTotalConnections.Open()
RecTotalConnections_numRows = 0
If IsNull(RecTotalConnections.Fields.Item("ConTotal").Value) = false then
GrossConnections = RecTotalConnections.Fields.Item("ConTotal").Value
End If


AirtimeDeposits = Replace(AirtimeDeposits, ",", ".")
If OutFormat = "B" Then
%>
<tr>
	<td><%=(RecAgentEdit.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediFirstName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediLastName").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("GenderType").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("RaceLabel").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("IDNumber").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TaxOffice").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TaxNumber").Value)%></td>
	<td><%=AgentStartDate%></td>
	<td><%=(RecAgentEdit.Fields.Item("ResidentialAddress1").Value)%>, <%=(RecAgentEdit.Fields.Item("ResidentialAddress2").Value)%>, <%=(RecAgentEdit.Fields.Item("ResidentialAddress3").Value)%>,<%=(RecAgentEdit.Fields.Item("ResidentialCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("TediCell").Value)%></td>
	<td><%=(TediType)%></td>
	<td><%=(RecAgentEdit.Fields.Item("BankLabel").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("BranchCode").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("AccountLabel").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("AccNo").Value)%></td>
	<td><%=(RecAgentEdit.Fields.Item("SubRegionName").Value)%></td>
	<td><%=(AirtimeDeposits)%></td>
	<td><%=(GrossConnections)%></td>
</tr>
<%
Response.flush
Else
TheFile.Writeline(RecAgentEdit.Fields.Item("TediEmpCode").Value & "," & RecAgentEdit.Fields.Item("TediFirstName").Value & "," & RecAgentEdit.Fields.Item("TediLastName").Value & "," & RecAgentEdit.Fields.Item("GenderType").Value & "," & RecAgentEdit.Fields.Item("RaceLabel").Value & "," & RecAgentEdit.Fields.Item("IDNumber").Value & "," & RecAgentEdit.Fields.Item("TaxOffice").Value & "," & RecAgentEdit.Fields.Item("TaxNumber").Value & "," & AgentStartDate & "," & Replace(RecAgentEdit.Fields.Item("ResidentialAddress1").Value, ",", " ") & " " & Replace(RecAgentEdit.Fields.Item("ResidentialAddress2").Value, ",", " ") & " " & Replace(RecAgentEdit.Fields.Item("ResidentialAddress3").Value, ",", " ") & " " & Replace(RecAgentEdit.Fields.Item("ResidentialCode").Value, ",", " ") & "," & RecAgentEdit.Fields.Item("TediCell").Value & "," & TediType & "," & RecAgentEdit.Fields.Item("BankLabel").Value & "," & RecAgentEdit.Fields.Item("BranchCode").Value & "," & RecAgentEdit.Fields.Item("AccountLabel").Value & "," & RecAgentEdit.Fields.Item("AccNo").Value & "," & RecAgentEdit.Fields.Item("SubRegionName").Value & "," & AirtimeDeposits & "," & GrossConnections)
End If
RecAgentEdit.MoveNext
Wend
If OutFormat = "B" Then
%>

</tbody>
</table>
<strong>Total New <%=AgentLabel%>s For This Period: <%=LC%></strong>
<%
Else
response.redirect("Reports/" & SaveFileName)
End If
%>