<!-- #include file="includes/header.asp" -->

<%
set RecReconRegions = Server.CreateObject("ADODB.Recordset")
RecReconRegions.ActiveConnection = MM_Site_STRING
RecReconRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
RecReconRegions.CursorType = 0
RecReconRegions.CursorLocation = 2
RecReconRegions.LockType = 3
RecReconRegions.Open()
RecReconRegions_numRows = 0

set RecSysUser = Server.CreateObject("ADODB.Recordset")
RecSysUser.ActiveConnection = MM_Site_STRING
RecSysUser.Source = "Select * FROM Users Where UserID = " & Session("UNID")
RecSysUser.CursorType = 0
RecSysUser.CursorLocation = 2
RecSysUser.LockType = 3
RecSysUser.Open()
RecSysUser_numRows = 0
%>

<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td align="left" valign="top"><form action="AddToScheduledReport.asp" name="ZRCS" method="Post" Target="_parent">
<table>
<tr>
<td class="quote">Region:</td>
<td><select name="Region">

<Option selected value="0">Select Region</Option>
<option value="0">All</option>
                <% 
While (NOT RecReconRegions.EOF)
%>
                <option value="<%=(RecReconRegions.Fields.Item("RID").Value)%>" ><%=(RecReconRegions.Fields.Item("RegionName").Value)%></option>
                <%
  RecReconRegions.MoveNext()
Wend
%>
              </select></td></tr>
<link rel="stylesheet" href="assets/css/pikaday.css">
    <style>

    a { color: #2996cc; }
    a:hover { text-decoration: none; }

    p { line-height: 1.0em; }
    .small { color: #666; font-size: 8px; }
    .large { font-size: 10px; }

    label {
        font-weight: bold;
    }

    </style> 
<%
StartDate = "1 " & MonthName(Month(Now)) & " " & Year(Now)
EndDate = Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)
%>

<tr><td class="quote">Start Date:</td><td><input type="text" id="datepicker1" Name="StartDate"  Value="<%=StartDate%>"></td></tr>
<tr><td class="quote">End Date:</td><td><input type="text" id="datepicker2" Name="EndDate"  Value="<%=EndDate%>"></td></tr>


<tr><td colspan="2"><h4>User Information:</h4></td></tr>
<tr><td class="quote">Run Report Date:</td><td><input type="text" id="datepicker3" Name="ReportDate"  Value="<%=EndDate%>"></td></tr>
<tr><td class="quote">Run Report Time:</td><td>
<select name="ReportTime">
<%
RHour = 0

Do While RHour < 25
RHourDisplay = RHour

		select Case Len(RHourDisplay)
		Case "1"
		RHourDisplay = "0" & RHourDisplay
		Case "2"
		RHourDisplay = RHourDisplay
		End Select

RMin = 0
Do While RMin < 60
RMinDisplay = RMin
		select Case Len(RMinDisplay)
		Case "1"
		RMinDisplay = "0" & RMinDisplay
		Case "2"
		RMinDisplay = RMinDisplay
		End Select
IsSelected = ""
OP = RHourDisplay & ":" & RMinDisplay
If OP = (Hour(Now()) & ":" & RMinDisplay) Then
IsSelected = "Selected"
End If
%><Option Value="<%=OP%>" <%=IsSelected%>><%=OP%></option>
<%
RMin = RMin + 5
Loop
RHour = RHour + 1
Loop%>
</select>
</td></tr>
<tr><td class="quote">Email Report To:</td><td><input type="email" name="ReportEmailAddress" Value="<%=RecSysUser.Fields.Item("UEmail").Value%>" style="width:80%; "></td></tr>


<tr><td Align="Center" Colspan="2"><input name="button2" type="submit" class="nice red radius button" id="button2" value="Schedule Report"></td></tr>
</table>
<input type="Hidden" Name="ReportRunID" Value="20">
<input type="Hidden" Name="ReportTypeID" Value="1">
<input type="Hidden" Name="UNID" Value="<%=Session("UNID")%>">
</form></td>
    
  </tr>

</table>

  <script src="moment.js"></script>
   <script src="pikaday.js"></script>
    <script>

    var picker = new Pikaday(
    {
        field: document.getElementById('datepicker1'),
        firstDay: 1,
	format: 'D MMMM YYYY',
        minDate: new Date('1950-01-01'),
        maxDate: new Date('2040-12-31'),
        yearRange: [1950,2040]
    });

    var picker = new Pikaday(
    {
        field: document.getElementById('datepicker2'),
        firstDay: 1,
	format: 'D MMMM YYYY',
        minDate: new Date('1950-01-01'),
        maxDate: new Date('2040-12-31'),
        yearRange: [1950,2040]
    });
    var picker = new Pikaday(
    {
        field: document.getElementById('datepicker3'),
        firstDay: 1,
	format: 'D MMMM YYYY',
        minDate: new Date('1950-01-01'),
        maxDate: new Date('2040-12-31'),
        yearRange: [1950,2040]
    });
    </script>