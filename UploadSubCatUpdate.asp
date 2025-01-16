<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
CatID = Request.Form("CatID")
SubCatID = Request.Form("SubCatID")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM UploadSubCategories where SubName = '" & Request.Form("SubName") & "' and SubActive = 'True' and SubCatID <> " & SubCatID
Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
'Response.end
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! An Upload Category Already exists in the system with the same Name in the selected Main Category.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM UploadSubCategories where SubCatID = " & SubCatID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("SubName") = Request.Form("SubName")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&ItemID=77&CatID=" & CatID)
%>
