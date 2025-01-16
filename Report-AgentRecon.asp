<!-- #include file="includes/header.asp" -->

<%
set RecReconRegions = Server.CreateObject("ADODB.Recordset")
RecReconRegions.ActiveConnection = MM_Site_STRING
RecReconRegions.Source = "Select Distinct RID, RegionName from viewUserRegion where UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
RecReconRegions.CursorType = 0
RecReconRegions.CursorLocation = 2
RecReconRegions.LockType = 3
RecReconRegions.Open()
RecReconRegions_numRows = 0

set RecReconTypes = Server.CreateObject("ADODB.Recordset")
RecReconTypes.ActiveConnection = MM_Site_STRING
RecReconTypes.Source = "SELECT * FROM TediReconTypes Where TypeActive = 'True' Order By ReconTypeLabel Asc"
RecReconTypes.CursorType = 0
RecReconTypes.CursorLocation = 2
RecReconTypes.LockType = 3
RecReconTypes.Open()
RecReconTypes_numRows = 0
%>

<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td Class="ontab"><h3>Agent Recon Report</h3></td>
  </tr>
  <tr>
    <td align="left" valign="top"><form action="Report-AgentRecon2.asp" name="ZRCS" method="get" Target="_parent">
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

    p { line-height: 1.5em; }
    .small { color: #666; font-size: 10px; }
    .large { font-size: 12px; }

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
<tr><td class="quote">Display In Report:</td><td><select Name="Display">
<option Value="2">All Agents With A Recon</option>
<option Value="3">All Agent Without A Recon</option>
</select></td></tr>

<tr><td class="quote">Output Format:</td><td><select Name="OutFormat">
<option Value="B">Browser</option>
<option Value="E">Excel</option>
</select></td></tr>
<tr><td class="quote">Amount of transactions to view:</td><td><select Name="TransCount">
<option Value="1">1</option>
<option Value="2">2</option>
<option Value="3">3</option>
<option Value="4">4</option>
<option Value="5">5</option>
</select></td></tr>
<tr><td class="quote">Recon Types:</td><td><select Name="ReconType">
<option Value="0">All</option>
<%While Not RecReconTypes.EOF%>
<option Value="<%=(RecReconTypes.Fields.Item("RTypeID").Value)%>"><%=(RecReconTypes.Fields.Item("ReconTypeLabel").Value)%></option>
<%
RecReconTypes.MoveNext
Wend
%>
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