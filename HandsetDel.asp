<!-- #include file="Connections/Site.asp" -->
<%

HID = Request.Form("HID")



Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Handsets where HID = " & HID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("HandsetAcitive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=36")

%>