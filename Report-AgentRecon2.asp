<!-- #include file="includes/header.asp" -->
<%
Region = Request.QueryString("Region")
ConStartDate = Replace(Request.QueryString("StartDate"), ",","")
ConEndDate = Replace(Request.QueryString("EndDate"), ",","")
ZonerTypes = Request.QueryString("Display")
OutFormat = Request.QueryString("OutFormat")
ReconType = Request.QueryString("ReconType")
TransCount = Int(Request.QueryString("TransCount"))
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

If ReconType = "0" then
ReconTypeLabel = "All Recon Types"
Else
set RecWR1 = Server.CreateObject("ADODB.Recordset")
RecWR1.ActiveConnection = MM_Site_STRING
RecWR1.Source = "SELECT * FROM TediReconTypes Where RTypeID = " & ReconType
RecWR1.CursorType = 0
RecWR1.CursorLocation = 2
RecWR1.LockType = 3
RecWR1.Open()
RecWR1_numRows = 0
ReconTypeLabel = RecWR1.Fields.Item("ReconTypeLabel").Value
End If

DisplayZonerType = ""
If ZonerTypes="1" Then
DisplayZonerType = "All Agents"
end If
If ZonerTypes="2" Then
DisplayZonerType = "All Agents With A Recon For Period"
End If
If ZonerTypes="3" Then
DisplayZonerType = "All Agents Without A Recon For Period"
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
SaveFileName = "Recon_Report-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
If ZonerTypes="1" or ZonerTypes="2" Then
TableHead = "Name, EmpCode, " & SupervisorLabel & ", Name, Region, Sub-Region, Recon Type, Recon Date, System Value, Phone / Stock Value, Cash / Banked, Difference, Comments"
Else
TableHead = "Name, EmpCode, " & SupervisorLabel & " Name, Region, Sub-Region, Has Recon"
End If
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)
End If
If OutFormat = "B" Then
%>
        <h3>Agent Recon Report</h3>
<p>Date Range: <b><%=ConStartDate%>&nbsp;to&nbsp;<%=ConEndDate%></b>
<br>Region: <b><%=WR%></b>
<br><b><%=DisplayZonerType%> | Recon Type: <%=ReconTypeLabel%></b>
<table border="0" align="center" cellpadding="2" cellspacing="2">
<%If ZonerTypes="1" or ZonerTypes="2" Then%>
<tr>
<th>Agent Name</th>
<th>Emp Code</th>
<th><%=SupervisorLabel%> Name</th>
<th>Region</th>
<th>Sub-Region</th>
<th>Recon Type</th>
<th>Date</th>
<th>System Value</th>
<th>Phone / Stock Val</th>
<th>Cash / Banked</th>
<th>Difference</th>
<th>Comments</th>
</tr>
<%Else%>
<thead>
<tr>
<th><b>Agent Name</b></th>
<th><b>Emp Code</b></th>
<th><%=SupervisorLabel%> Name</th>
<th>Region</th>
<th>Sub-Region</th>
<th>Has Recon</th>
</thead>
</tr>
<%End If%>
<tbody>
<%
End If
If Region = "0" then
SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID")
Else
SubRegionQry = "Select * from ViewUserSubRegions where UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " and RID = " & Region
End If
set RecWatchlistRegions = Server.CreateObject("ADODB.Recordset")
RecWatchlistRegions.ActiveConnection = MM_Site_STRING
RecWatchlistRegions.Source = SubRegionQry
RecWatchlistRegions.CursorType = 0
RecWatchlistRegions.CursorLocation = 2
RecWatchlistRegions.LockType = 3
RecWatchlistRegions.Open()
RecWatchlistRegions_numRows = 0
While Not RecWatchlistRegions.EOF
SRRegionList = SRRegionList & RecWatchlistRegions.Fields.Item("SRID").Value & ","
RecWatchlistRegions.MoveNext
Wend
TempLenSRRegionList = Len(SRRegionList)
SRRegionList = Left(SRRegionList,TempLenSRRegionList - 1)

set RecActiveZoners = Server.CreateObject("ADODB.Recordset")
RecActiveZoners.ActiveConnection = MM_Site_STRING
RecActiveZoners.Source = "SELECT * FROM ViewTediDetail Where SRID in (" & SRRegionList & ") and TediActive = 'True' Order By RegionName, SubRegionName, TediEmpCode Asc"
'Response.Write(RecActiveZoners.Source)
RecActiveZoners.CursorType = 0
RecActiveZoners.CursorLocation = 2
RecActiveZoners.LockType = 3
RecActiveZoners.Open()
RecActiveZoners_numRows = 0
While Not RecActiveZoners.EOF

HasTrans = "No"
ZWR = "SELECT * FROM ViewTediReconDetails where ReconActive = 'True' and TID = " & RecActiveZoners.Fields.Item("TID").Value

ZWR = ZWR & " and (ReconDate BETWEEN '" & ConStartDate & "' AND '" & ConEndDate & " 23:59:59')"

If ReconType <> "0" then
ZWR = ZWR & " and TypeID = " & ReconType
end If

ZWR = ZWR & " Order By ReconDate desc"
set RecZonerWithRecon = Server.CreateObject("ADODB.Recordset")
RecZonerWithRecon.ActiveConnection = MM_Site_STRING
RecZonerWithRecon.Source = ZWR
RecZonerWithRecon.CursorType = 0
RecZonerWithRecon.CursorLocation = 2
RecZonerWithRecon.LockType = 3
RecZonerWithRecon.Open()
RecZonerWithRecon_numRows = 0
ZT = 0
While Not RecZonerWithRecon.EOF and (ZT < TransCount)
ZT = ZT + 1
HasTrans = "Yes"
ReconVVal = 0
DMGVal = FormatNumber(RecZonerWithRecon.Fields.Item("SystemValue").Value,2)
StockVal = FormatNumber(RecZonerWithRecon.Fields.Item("StockValue").Value,2)
CashVal = FormatNumber(RecZonerWithRecon.Fields.Item("CashValue").Value,2)
ReconRowCol = "offtabred"
ReconVVal = ReconVVal + StockVal
ReconVVal = ReconVVal + CashVal
ReconVVal = FormatNumber(ReconVVal,2)
If CStr(DMGVal) = CStr(ReconVVal) Then
ReconRowCol = "offtabGreen"
'End If
End If
If (ZonerTypes="1" or ZonerTypes="2") and HasTrans = "Yes" Then
If OutFormat = "B" Then
%>
<tr>
<td><%=(RecActiveZoners.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecActiveZoners.Fields.Item("TediLastName").Value)%></div></td>
<td><%=(RecActiveZoners.Fields.Item("TediEmpCode").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecActiveZoners.Fields.Item("ASLastName").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("RegionName").Value)%></div></td>
<td><%=(RecActiveZoners.Fields.Item("SubRegionName").Value)%></div></td>
<td><%=(RecZonerWithRecon.Fields.Item("ReconTypeLabel").Value)%></td>
<td><div align="left"><%=Day(RecZonerWithRecon.Fields.Item("ReconDate").Value)%>&nbsp;<%=MonthName(Month(RecZonerWithRecon.Fields.Item("ReconDate").Value))%>&nbsp;<%=Year(RecZonerWithRecon.Fields.Item("ReconDate").Value)%></div></td>
<td><%=DMGVal%></td>
<td><div Align="Left"><%=FormatNumber(StockVal,2)%></div></td>
<td><%=FormatNumber(CashVal,2)%></td>
<td><%=FormatNumber(DMGVal - StockVal - CashVal,2)%></td>
<td><%=(RecZonerWithRecon.Fields.Item("RecComments").Value)%></td>
</tr>
<%
Else
TheFile.Writeline(RecActiveZoners.Fields.Item("TediFirstName").Value & " " & RecActiveZoners.Fields.Item("TediLastName").Value & "," & RecActiveZoners.Fields.Item("TediEmpCode").Value & "," & RecActiveZoners.Fields.Item("ASFirstName").Value & " " & RecActiveZoners.Fields.Item("ASLastName").Value & "," & RecActiveZoners.Fields.Item("RegionName").Value & "," & RecActiveZoners.Fields.Item("SubRegionName").Value & ", " & RecZonerWithRecon.Fields.Item("ReconTypeLabel").Value & ", " & Day(RecZonerWithRecon.Fields.Item("ReconDate").Value) & " " & MonthName(Month(RecZonerWithRecon.Fields.Item("ReconDate").Value)) & " " & Year(RecZonerWithRecon.Fields.Item("ReconDate").Value) & ", " & DMGVal & ", " & StockVal & ", " & CashVal & ", " & (DMGVal - StockVal - CashVal) & ", " & RecZonerWithRecon.Fields.Item("RecComments").Value)
End If
End If
RecZonerWithRecon.MoveNext
Wend
If (ZonerTypes="3") and HasTrans = "No" Then
If OutFormat = "B" Then
%>
<tr>
<td><%=(RecActiveZoners.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecActiveZoners.Fields.Item("TediLastName").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("TediEmpCode").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecActiveZoners.Fields.Item("ASLastName").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("RegionName").Value)%></td>
<td><%=(RecActiveZoners.Fields.Item("SubRegionName").Value)%></td>
<td>No</td>
</tr>
<%
Else
TheFile.Writeline(RecActiveZoners.Fields.Item("TediFirstName").Value & " " & RecActiveZoners.Fields.Item("TediLastName").Value & "," & RecActiveZoners.Fields.Item("TediEmpCode").Value & "," & RecActiveZoners.Fields.Item("ASFirstName").Value & " " & RecActiveZoners.Fields.Item("ASLastName").Value & ", No")
End If
End If
RecActiveZoners.MoveNext
Wend
If OutFormat = "B" Then
%>
</tbody>
</table>
<%
Else
response.redirect("Reports/" & SaveFileName)
End If
%>