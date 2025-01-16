<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
If IsNumeric(Request.Form("DeductValue")) = "False" Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Please only submit a numeric value deduction value field.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateZTTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateZTTable.Open "SELECT * From Deductions", MM_Site_STRINGWrite, 1, 2
RecUpdateZTTable.AddNew
RecUpdateZTTable("TID") = Request.Form("TID")
RecUpdateZTTable("DeductionDate") = Now()
RecUpdateZTTable("DeductionBy") = Request.Form("SID")
RecUpdateZTTable("DeductionCatID") = Request.Form("DeductType")
RecUpdateZTTable("DeductionValue") = Request.Form("DeductValue")
RecUpdateZTTable.Update
RecUpdateZTTable.Close

Response.Redirect("TediView.asp?TID=" & Request.Form("TID") & "&Item=8")
%>