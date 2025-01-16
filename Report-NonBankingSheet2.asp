<%
Region = Request.QueryString("Region")
OutFormat = Request.QueryString("OutFormat")
%>
<!-- #include file="includes/header.asp" -->
<%

If OutFormat <> "B" Then
SavePath = AppPath & "Reports/"
SaveFileName = "NonBanking_Sheet-" & RType & "-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".xls"

TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
End If

If Region = "0" then
WR = "All Regions"
Else
set RecWR = Server.CreateObject("ADODB.Recordset")
RecWR.ActiveConnection = MM_Site_STRING
RecWR.Source = "SELECT * FROM [Regions] Where CompanyID = " & Session("CompanyID") & " and RID = " & Region
RecWR.CursorType = 0
RecWR.CursorLocation = 2
RecWR.LockType = 3
RecWR.Open()
RecWR_numRows = 0
WR = RecWR.Fields.Item("RegionName").Value
End If

TheFile.Writeline("Region - " & WR)
TheFile.Writeline("<br>Sheet Date - " & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now))
TheFile.Writeline("<table border=1>")
%>
<h3>Non-Banking Comment / Reason Sheet</h3>
<p>Region: <b><%=WR%></b>
<br>Sheet Date: <b><%=Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)%></b>

<table>
<%

SubRegionQry = "Select * from ViewUserSubRegions where CompanyID = " & Session("CompanyID") & " and UserID = " & Session("UNID")

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


set RecWR2 = Server.CreateObject("ADODB.Recordset")
RecWR2.ActiveConnection = MM_Site_STRING
If Region <> "0" then
RecWR2.Source = "SELECT Distinct RegionName, RID FROM viewUserRegion Where CompanyID = " & Session("CompanyID") & " and RID = " & Region
else
RecWR2.Source = "SELECT Distinct RegionName, RID FROM viewUserRegion Where CompanyID = " & Session("CompanyID") & " and UserID = " & Session("UNID") & " order by RegionName ASC"
End If
RecWR2.CursorType = 0
RecWR2.CursorLocation = 2
RecWR2.LockType = 3
RecWR2.Open()
RecWR2_numRows = 0
While Not RecWR2.EOF

set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RecWR2.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0

If Not RecCurrent.EOF and Not RecCurrent.BOF Then
'TheFile.Writeline(RecWR2.Fields.Item("RegionName").Value)
TheFile.Writeline("<tr><th colspan=4 bgcolor=#FFC000><div align=center><strong>" & RecWR2.Fields.Item("RegionName").Value & "</strong></div></th></tr>")
%>


<tr>
	<th colspan="4"><div align="center"><strong><%=RecWR2.Fields.Item("RegionName").Value%></strong></div></th>
</tr>
<%
While Not RecCurrent.EOF

set RecZonerCount = Server.CreateObject("ADODB.Recordset")
RecZonerCount.ActiveConnection = MM_Site_STRING
RecZonerCount.Source = "SELECT * FROM Tedis where TediActive = 'True' and ASID = " & RecCurrent.Fields.Item("ASID").Value & " order by TediEmpCode"
RecZonerCount.CursorType = 0
RecZonerCount.CursorLocation = 2
RecZonerCount.LockType = 3
RecZonerCount.Open()
RecZonerCount_numRows = 0
'TheFile.Writeline(RecCurrent.Fields.Item("ASFirstName").Value & " " & RecCurrent.Fields.Item("ASLastName").Value)
'TheFile.Writeline("Agent Code,Agent Name,Days Since Last Banked,Comments / Reason for Non-Banking")

TheFile.Writeline("<tr><th colspan=4 bgcolor=#FFC000><div align=center><strong>" & RecCurrent.Fields.Item("ASFirstName").Value & " " & RecCurrent.Fields.Item("ASLastName").Value & "</strong></div></th></tr>")
TheFile.Writeline("<tr  bgcolor=#FFC000><td width=100><strong>Agent Code</strong></td><td width=200><strong>Agent Name</strong></td><td width=150><strong>Days Since Last Banked</strong></td><td width=400><strong>Comments / Reason for Non-Banking</strong></td></tr>")
%>
<tr>
	<th colspan="4"><div align="center"><strong><%=(RecCurrent.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("ASLastName").Value)%></strong></div></th>
</tr>
<tr>
	<td width="100"><strong>Agent Code</strong></td>
	<td width="200"><strong>Agent Name</strong></td>
	<td width="150"><strong>Days Since Last Banked</strong></td>
	<td><strong>Comments / Reason for Non-Banking</strong></td>
</tr>
<%
While Not RecZonerCount.EOF
LastBankedDays = 0


If IsDate(RecZonerCount.Fields.Item("LastBankedDate").Value) = "True" Then
LastBankedDays = Day(RecZonerCount.Fields.Item("LastBankedDate").Value) & " " & MonthName(Month(RecZonerCount.Fields.Item("LastBankedDate").Value)) & " " & Year(RecZonerCount.Fields.Item("LastBankedDate").Value)
LastBankedDays= DateDiff("d",LastBankedDays,Date())
Else
LastBankedDays = "N/A"
End If
If LastBankedDays <> "0" Then
'TheFile.Writeline(RecZonerCount.Fields.Item("TediEmpCode").Value & "," & RecZonerCount.Fields.Item("TediFirstName").Value & " " & RecZonerCount.Fields.Item("TediLastName").Value & "," & LastBankedDays & ",")

TheFile.Writeline("<tr><td>" & RecZonerCount.Fields.Item("TediEmpCode").Value & "</td><td>" & RecZonerCount.Fields.Item("TediFirstName").Value & " " & RecZonerCount.Fields.Item("TediLastName").Value & "</td><td><div align=right>" & LastBankedDays & "</div></td><td></td></tr>")
%>
<tr>
	<td><%=(RecZonerCount.Fields.Item("TediEmpCode").Value)%></td>
	<td><%=(RecZonerCount.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecZonerCount.Fields.Item("TediLastName").Value)%></td>
	<td><%=LastBankedDays%></td>
	<td></td>
</tr>
<%
End If
RecZonerCount.MoveNext
Wend
%>

<%

RecCurrent.MoveNext
Wend

End If
RecWR2.MoveNext
Wend
TheFile.Writeline("<table>")
%>
</table>
<%
If OutFormat <> "B" Then
response.redirect("Reports/" & SaveFileName)
End If
%>
