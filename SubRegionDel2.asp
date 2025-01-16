<!-- #include file="Connections/Site.asp" -->
<%


RID = Request.Form("RID")
SRID = Request.Form("SRID")



Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SubRegions where SRID = " & SRID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("SubRegionActive") = "False"
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

%>
<!-- #include file="Includes/UpdateuserregionsAutomatic.inc" -->
<%

Response.Redirect("SubRegions.asp?RID=" & RID)
%>
