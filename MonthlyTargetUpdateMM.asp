<!-- #include file="Connections/Site.asp" -->
<%
ID = Request.Form("ID")
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


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM  MonthlyTargetsMM where ID = " & ID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("AirtimeTarget") = AirtimeTarget
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.redirect("Updated.asp?AppCat=7&AppSubCatID=1044")
%>