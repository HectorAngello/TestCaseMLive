<!-- #include file="Connections/Site.asp" -->
<%
AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")
GroupID = Request.Form("GroupID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SecurityGroups where GroupID = " & GroupID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("GroupActive") = 0
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
