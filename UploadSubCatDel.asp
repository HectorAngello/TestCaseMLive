<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
CatID = Request.Form("CatID")
SubCatID = Request.Form("SubCatID")
ItemID = "77"


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM UploadSubCategories where SubCatID = " & SubCatID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("SubActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&ItemID=77&CatID=" & CatID)
%>
