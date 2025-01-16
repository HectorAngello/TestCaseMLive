<!-- #include file="Connections/Site.asp" -->
<%
DoInsert = "Yes"
ErrorMsg = "ERROR !!!! "

CashVal = Replace(Request.form("CashValue"), "R", "")
CashVal = Replace(CashVal, "r", "")
CashVal = Replace(CashVal, " ", "")

If IsNumeric(CashVal) = False Then
DoInsert = "No"
ErrorMsg = ErrorMsg & "The Value For Cash Value " & Request.Form("CashValue") & " is invalid, please only submit the value e.g. 10.00, Please do not include a 'R'"
Else
CashVal = FormatNumber(CashVal,2)
End If

StockVal = Replace(Request.Form("StockValue"), "R", "")
StockVal = Replace(StockVal, "r", "")
StockVal = Replace(StockVal, " ", "")

If IsNumeric(StockVal) = False Then
DoInsert = "No"
ErrorMsg = ErrorMsg & "The Value For Stock Value " & Request.Form("StockValue") & " is invalid, please only submit the value e.g. 10.00, Please do not include a 'R'"
Else
StockVal = FormatNumber(StockVal,2)
End If



If DoInsert = "Yes" Then
		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecBackUp = Server.CreateObject ( "ADODB.Recordset" )
		RecBackUp.Open "SELECT Top(1) * FROM TediRecons", MM_Site_STRINGWrite, 1, 2
		RecBackUp.AddNew
		RecBackUp("TypeID") = Request.Form("TypeID")
		RecBackUp("TID") = Request.Form("TID")
		RecBackUp("AddedBy") = Request.Form("AddedBy")
		RecBackUp("ReconDate") = Now()
		RecBackUp("SystemValue") = Request.Form("DMGVal")
		RecBackUp("StockValue") = StockVal
		RecBackUp("CashValue") = CashVal
		RecBackUp("RecComments") = Request.Form("RecComments")
		RecBackUp.Update
		RecBackUp.Close
Response.Redirect("TediView.asp?TID=" & Request.Form("TID") & "&Item=7")
Else
%>
	<script language="JavaScript" type="text/JavaScript">
	<!--
	  alert("Error - <%=ErrorMsg%>");
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If

%>