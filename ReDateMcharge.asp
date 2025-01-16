<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
            </div>
            <div class="nine columns">
                <div class="content panel">

                        <h1>Re-Date an Mcharge Bulk File</h1>

<p>Use this to redate all the transactions in a bulk file.</p>
<%
WhichBulkID = Request.QueryString("BulkID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecBulkID = Server.CreateObject("ADODB.Recordset")
RecBulkID.ActiveConnection = MM_Site_STRING
RecBulkID.Source = "SELECT * FROM BulkMCharge Where BulkID = " & WhichBulkID
'Response.Write(RecBulkID.Source)
RecBulkID.CursorType = 0
RecBulkID.CursorLocation = 2
RecBulkID.LockType = 3
RecBulkID.Open()
RecBulkID_numRows = 0

LineCount = 0
Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM  BulkMChargeChildren Where BulkID = " & WhichBulkID
'Response.Write(RecCurrent.Source)
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
While Not RecCurrent.EOF
LineCount = LineCount + 1
RecCurrent.MoveNext
Wend

BulkDate = Day(RecBulkID.Fields.Item("BulkDate").Value) & " " & MonthName(Month(RecBulkID.Fields.Item("BulkDate").Value)) & " " & Year(RecBulkID.Fields.Item("BulkDate").Value)
%>
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

<form action="RedateMcharge2.asp" name="ZRCS" method="get">
<table>
<td class="quote">Selected Bulk File:</td><td><b><%=WhichBulkID%></b></td></tr>
<td class="quote">Current Bulk File Date:</td><td><b><%=Day(RecBulkID.Fields.Item("BulkDate").Value)%>&nbsp;<%=MonthName(Month(RecBulkID.Fields.Item("BulkDate").Value))%>&nbsp;<%=Year(RecBulkID.Fields.Item("BulkDate").Value)%></b></td></tr>
<td class="quote">Transactions to be updated:</td><td><b><%=LineCount%></b></td></tr>
<td class="quote">New Bulk File Date:</td><td><input type="text" id="datepicker" Name="StartDate" class="input-text" Value="<%=BulkDate%>"></td></tr>
</table>
<input type="Hidden" name="BulkID" value="<%=WhichBulkID%>">
<input name="button2" type="submit" class="quote" id="button2" value="Update Bulk File Date">
</form>

<!-- #include file="includes/footer.asp" -->

