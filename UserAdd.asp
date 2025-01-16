<!-- #include file="Connections/Site.asp" -->
<%
EmailAllGood = "Yes"
CheckEmail = Request.Form("UserEmail")

EmailCheckSpace = instr (1,CheckEmail, " ", 1) 
if EmailCheckSpace > 0 then 
EmailAllGood = "No"
End If

EmailCheckComma = instr (1,CheckEmail, ",", 1) 
if EmailCheckComma > 0 then 
EmailAllGood = "No"
End If

If EmailAllGood = "No" Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Email address submitted contains invalid characters.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM Users where (UEmail = '" & Request.Form("UserEmail") & "' or CellNo = '" & Request.Form("UserMobile") & "') and UserActive = 'True' and CompanyID = " & Request.Form("CompanyID")
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
      window.alert ("Error ! A user Already exists in the system, with either the same Username, Email Address or Mobile Number");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM Users", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("UserName") = Replace(Request.Form("UserName"), " ", "")
rstSecond("Password") = Replace(Request.Form("UserPass"), " ", "")
rstSecond("UEmail") = Request.Form("UserEmail")
rstSecond("CellNo") = Request.Form("UserMobile")
rstSecond("UserSecurityGroupID") = Request.Form("UserSecurityGroupID")
rstSecond("UserFirstName") = Request.Form("FirstName")
rstSecond("UserLastName") = Request.Form("LastName")
rstSecond("CompanyID") = Request.Form("CompanyID")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&ItemID=" & ItemID)
%>
