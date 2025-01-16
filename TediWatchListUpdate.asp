<!-- #include file="Connections/Site.asp" -->
<%
Watchlist = Request.Querystring("Watchlist")
TID = Request.Querystring("TID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("OnWatchList") = Watchlist
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

TediUpdateType = "Agent Watch List Status Updated"
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%

Response.redirect("TediView.asp?TID=" & TID)
%>