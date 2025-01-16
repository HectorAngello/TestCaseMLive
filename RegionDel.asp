<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
RID = Request.Form("RID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM Regions where RID = " & RID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Active") = "No"
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

%>
<!-- #include file="Includes/UpdateuserregionsAutomatic.inc" -->
<%

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
