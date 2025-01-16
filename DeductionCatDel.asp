<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ID = Request.Form("ID")




Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM DeductionCategories where ID = " & ID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("DeductionActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
