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
UserID = Request.Form("UserID")



Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM Users where UserID = " & UserID, MM_Site_STRINGWrite, 1, 2
rstSecond.update
rstSecond("UserName") = Replace(Request.Form("UserName"), " ", "")
rstSecond("Password") = Replace(Request.Form("UserPass"), " ", "")
rstSecond("UEmail") = Request.Form("UserEmail")
rstSecond("UserFirstName") = Request.Form("FirstName")
rstSecond("UserLastName") = Request.Form("LastName")
rstSecond("CellNo") = Request.Form("UserMobile")
rstSecond("UserSecurityGroupID") = Request.Form("UserSecurityGroupID")

rstSecond.Update
rstSecond.Close
set rstSecond = nothing

ChangeType = "System User Updated By: " & Session("UNID")
%><!--#include file="Includes/ErrorTrap.inc" --><%

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
