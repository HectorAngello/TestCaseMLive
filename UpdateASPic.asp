<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.QueryString("ASID")
TediPic = Replace(Request.QueryString("NewPic"), "ASImages/", "")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ASs Where ASID = " & ASID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("ASProfilePic") = TediPic
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.Redirect("ASView.asp?ASID=" & ASID)

%>