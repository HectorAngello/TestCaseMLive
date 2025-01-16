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


%>

<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td Class="ontab"><h3>MTD Report MCharge</h3></td>
  </tr>
  <tr>
    <td align="left" valign="top"><form action="Report-MTD2.asp" name="ZRCS" method="get" Target="_parent">
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

<tr><td class="quote">Report Data:</td><td><select Name="RepDataType">
<option Value="Airtime">Airtime</option>
<option Value="Deductions">Deductions</option>
<option Value="Deposits">Deposits</option>
</select></td></tr>

<tr><td class="quote">Output Format:</td><td><select Name="OutFormat">
<option Value="B">Browser</option>
<option Value="E">Excel</option>
</select></td></tr>

<tr><td Align="Center" Colspan="2"><input name="button2" type="submit" class="nice red radius button" id="button2" value="Run Report"></td></tr>
</table>
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
    </script>