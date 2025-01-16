<!-- #include file="Connections/Site.asp" -->
<%
TID = Request.QueryString("TID")
TediPic = Replace(Request.QueryString("NewPic"), "TediImages/", "")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TediPic") = TediPic
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.Redirect("TediView.asp?TID=" & TID)

%>