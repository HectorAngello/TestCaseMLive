<!-- #include file="Connections/Site.asp" -->
<%
ID = Request.QueryString("ID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM TrainingFiles Where ID = " & ID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("FileActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing




Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=1043")

%>