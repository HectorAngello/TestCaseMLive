<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%
FNBID = Request.QueryString("FNBID")
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(2) * FROM MChargeFNBTransMM where FNBID = " & FNBID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Allocated") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

Response.redirect "Display.asp?AppCat=18&AppSubCatID=1046&ItemID=2301"
%>