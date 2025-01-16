<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")
CompanyID = Request.Form("CompanyID")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM DeductionCategories where (DeductionLabel = '" & Request.Form("CatName") & "') and CompanyID = " & Session("CompanyID") & " and DeductionActive = 'True'"
Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! A Deduction Category Already exists in the system, with the same Name.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM DeductionCategories", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("DeductionLabel") = Request.Form("CatName")
rstSecond("DeductionActive") = "True"
rstSecond("CompanyID") = CompanyID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
