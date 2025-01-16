<!-- #include file="Connections/Site.asp" -->
<%
PeriodMonth = Request.Form("PeriodMonth")
PeriodYear = Request.Form("PeriodYear")
AirtimeTarget = Replace(Request.Form("AirtimeTarget"), ".", ",")

ErrorMSG = ""

If Isnumeric(AirtimeTarget) = "False" then
ErrorMSG = "Mobile Money Target"
End If

If ErrorMSG <> "" Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! <%=ErrorMSG%> Needs to be a numeric value");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select Top(1)* FROM MonthlyTargetsMM where PeriodMonth = " & PeriodMonth & " and PeriodYear = " & PeriodYear
'Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! The selected period already has a target set, rather edit the period");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM  MonthlyTargetsMM", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("PeriodMonth") = PeriodMonth
rstSecond("PeriodYear") = PeriodYear
rstSecond("AirtimeTarget") = AirtimeTarget
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.redirect("Updated.asp?AppCat=7&AppSubCatID=1044")
%>