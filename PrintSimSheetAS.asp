<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<style type="text/css">
<!--
.style5 {font-size: 12px; font-family: Arial, Helvetica, sans-serif; }
.style6 {font-size: 12px; font-family: Arial, Helvetica, sans-serif; }
-->
</style>
<%


BulkID = Request.Querystring("BulkID")
TodayDate = FormatDateTime(Now(),1)

set RecBulkV = Server.CreateObject("ADODB.Recordset")
RecBulkV.ActiveConnection = MM_Site_STRING
RecBulkV.Source = "SELECT Top(1) * FROM BulkSimsAS Where BulkID = " & Request.QueryString("BulkID") 
RecBulkV.CursorType = 0
RecBulkV.CursorLocation = 2
RecBulkV.LockType = 3
RecBulkV.Open()
RecBulkV_numRows = 0

set RecBulkChildren = Server.CreateObject("ADODB.Recordset")
RecBulkChildren.ActiveConnection = MM_Site_STRING
RecBulkChildren.Source = "SELECT * FROM BulkSimChildrenAS Where BulkID = " & Request.QueryString("BulkID") & " Order By ChildCreationDate Asc"
RecBulkChildren.CursorType = 0
RecBulkChildren.CursorLocation = 2
RecBulkChildren.LockType = 3
RecBulkChildren.Open()
RecBulkChildren_numRows = 0

set RecTedi = Server.CreateObject("ADODB.Recordset")
RecTedi.ActiveConnection = MM_Site_STRING
RecTedi.Source = "SELECT * FROM ViewASDetail where ASID = " & RecBulkV.Fields.Item("ASID").Value
RecTedi.CursorType = 0
RecTedi.CursorLocation = 2
RecTedi.LockType = 3
RecTedi.Open()
RecTedi_numRows = 0

	  %>
<P STYLE="page-break-after: always;">
<table width="700" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#000000">
  <tr>
    <td><table width="700" border="0" cellpadding="1" cellspacing="1">
      <tr>
        <td colspan="2" align="left" bgcolor="#FFFFFF" width="200"><span class="style5"><img src="Images/mtn_logo.jpg" width="200"></span></td>
        <td colspan="2" align="center" bgcolor="#FFFFFF" width="500" Valign="top"><span class="style5"><h2><%=SupervisorLabel%> Sim Allocation Sheet</h2>
	<h3>Agent: <%=RecTedi.Fields.Item("ASEmpCode").Value%><br><%=RecTedi.Fields.Item("ASFirstName").Value%>&nbsp;<%=RecTedi.Fields.Item("ASLastName").Value%><br><br>Printed Date:<br><%=FormatDateTime(Now(),1)%><br><br>SIM Capture Date:<br><%=FormatDateTime(RecBulkV.Fields.Item("BulkDate").Value,1)%></h3></span>
	</td>
      </tr>

      <tr>
        <td bgcolor="#FFFFFF"><span class="style5">&nbsp;&nbsp;<b>Date<b/></span></td>
        <td bgcolor="#FFFFFF"><span class="style5">&nbsp;&nbsp;<b>Sim Number:</b></span></td>
        <td bgcolor="#FFFFFF" colspan="2"><span class="style5">&nbsp;&nbsp;<b>Customer Signature:</b></span></td>

      </tr>
<%
TotVoucher = 0
While Not RecBulkChildren.EOF
%>
      <tr>
        <td bgcolor="#FFFFFF"  width="250">&nbsp;&nbsp;</span></td>
        <td bgcolor="#FFFFFF">&nbsp;&nbsp;<span class="style5"><%=(RecBulkChildren.Fields.Item("SerialNo").Value)%></span></td>
        <td bgcolor="#FFFFFF" width="250">&nbsp;&nbsp;</span></td>

      </tr>
<%

RecBulkChildren.MoveNext
Wend
%>


      <tr>
        <td colspan="4" bgcolor="#FFFFFF">
	<table border="0" Width="100%" height="180">
<tr valign="bottom" align="Center">
<td width="33%"><span class="style5">------------------------------------<br>Administrator Signature</span></td>
<td width="33%"><span class="style5">------------------------------------<br>Date</span></td>
<td width="33%"><span class="style5">------------------------------------<br><%=SupervisorLabel%> Signature</span></td>
</tr>
	</table>
	</td>
        </tr>
    </table></td>
  </tr>
</table>
</p>
