<!-- #include file="Connections/Site.asp" -->
<%
AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")
UserID = Request.Form("UserID")


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM Users where UserID = " & UserID, MM_Site_STRINGWrite, 1, 2
rstSecond.update
rstSecond("UserActive") = "False"

rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
