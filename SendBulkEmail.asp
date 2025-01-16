<!-- #include file="Connections/Site.asp" -->
<%
set RecAirTimeSalesRegions = Server.CreateObject("ADODB.Recordset")
RecAirTimeSalesRegions.ActiveConnection = MM_Site_STRING
RecAirTimeSalesRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
RecAirTimeSalesRegions.CursorType = 0
RecAirTimeSalesRegions.CursorLocation = 2
RecAirTimeSalesRegions.LockType = 3
RecAirTimeSalesRegions.Open()
RecAirTimeSalesRegions_numRows = 0



set RecUpdateSysUser = Server.CreateObject("ADODB.Recordset")
RecUpdateSysUser.ActiveConnection = MM_Site_STRING
RecUpdateSysUser.Source = "Select * FROM Users Where UserID = " & Session("UNID")
RecUpdateSysUser.CursorType = 0
RecUpdateSysUser.CursorLocation = 2
RecUpdateSysUser.LockType = 3
RecUpdateSysUser.Open()
RecUpdateSysUser_numRows = 0
FromEmail = RecUpdateSysUser.Fields.Item("UEmail").Value
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge">
<meta charset="utf-8">

<title>zurb-foundation-wysihtml5</title>


<link rel="stylesheet" type="text/css" href="lib/css/foundation.min.css"></link>
<link rel="stylesheet" type="text/css" href="lib/css/prettify.css"></link>
<link rel="stylesheet" type="text/css" href="lib/css/foundation-glyphicons.css"></link>
<link rel="stylesheet" type="text/css" href="src/zurb-foundation-wysihtml5.css"></link>

<style type="text/css" media="screen">
	.btn.jumbo {
		font-size: 20px;
		font-weight: normal;
		padding: 14px 24px;
		margin-right: 10px;
		-webkit-border-radius: 6px;
		-moz-border-radius: 6px;
		border-radius: 6px;
	}
</style>
</head>
<body>
<div class="row">
	<div class="twelve columns" style="margin-top:40px">
<h4>Compose A New Bulk Email:</h4>
<form name="AddUser" action="SendBulkCom2.asp" method="post" class="nice">
                        <fieldset>
                           
                                <label for="agentEmail">Select Recipients</label>
<select name="Region">
<Option selected value="0--0">Select Region</Option>
<option value="0--0">All My Regions</option>
                <% 
While (NOT RecAirTimeSalesRegions.EOF)
%>
                <option value="<%=(RecAirTimeSalesRegions.Fields.Item("RID").Value)%>--0"><%=(RecAirTimeSalesRegions.Fields.Item("RegionName").Value)%> (All <%=SupervisorLabel%>s)</option>
                <%
set RecAirTimeSalesRegionsMR = Server.CreateObject("ADODB.Recordset")
RecAirTimeSalesRegionsMR.ActiveConnection = MM_Site_STRING
RecAirTimeSalesRegionsMR.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RecAirTimeSalesRegions.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecAirTimeSalesRegionsMR.CursorType = 0
RecAirTimeSalesRegionsMR.CursorLocation = 2
RecAirTimeSalesRegionsMR.LockType = 3
RecAirTimeSalesRegionsMR.Open()
RecAirTimeSalesRegionsMR_numRows = 0
While Not RecAirTimeSalesRegionsMR.EOF
%>
<option value="<%=(RecAirTimeSalesRegionsMR.Fields.Item("RID").Value)%>--<%=(RecAirTimeSalesRegionsMR.Fields.Item("ASID").Value)%>" ><%=(RecAirTimeSalesRegions.Fields.Item("RegionName").Value)%> - <%=(RecAirTimeSalesRegionsMR.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecAirTimeSalesRegionsMR.Fields.Item("ASLastName").Value)%></option>
<%
RecAirTimeSalesRegionsMR.Movenext
Wend
  RecAirTimeSalesRegions.MoveNext()
Wend
%>
              </select>
                                <label for="agentEmail">Send Email To</label>
<Select Name="SendType">
<option value="1">All Agents And All <%=SupervisorLabel%>s</option>
<option value="2">ONLY Agents</option>
<option value="3">ONLY <%=SupervisorLabel%>s</option>
</select>
                                <label for="agentEmail">Email From:</label>
				<input type="Text" Name="EmailFrom" value="<%=FromEmail%>" Required>
                                <label for="agentEmail">Email Subject:</label>
				<input type="Text" Name="EmailSubJect" Required>
                                <label for="agentEmail">Email Body:</label>
				<textarea class="textarea" style="width: 100%; height: 300px" name="MSG"></textarea>
                                <p>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Send Email">
                                </p>
			
			</fieldset>
<input type="Hidden" Name="ComType" Value="Email">
</form>


</div>

</div>



<script src="lib/js/wysihtml5-0.3.0.js"></script>
<script src="lib/js/jquery-1.7.2.min.js"></script>
<script src="lib/js/prettify.js"></script>
<script src="lib/js/foundation.min.js"></script>
<script src="lib/js/jquery.foundation.reveal.js"></script>
<script src="lib/js/jquery.foundation.buttons.js"></script>
<script src="src/zurb-foundation-wysihtml5.js"></script>

<script>
	$('.textarea').wysihtml5();
</script>

<script type="text/javascript" charset="utf-8">
	$(prettyPrint);
</script>

</body>
</html>
