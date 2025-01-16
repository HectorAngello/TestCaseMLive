<!-- #include file="Connections/Site.asp" -->
<%

ID = Request.QueryString("ID")
UserID = Request.QueryString("UserID")
UserType = Request.QueryString("UserType")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM UserUploadedFiles where ID = " & ID, MM_Site_STRINGWrite, 1, 2
rstSecond.update
rstSecond("FileActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

If UserType = "1" Then
Response.Redirect("TediView.asp?Item=4&TID=" & UserID)
End If
If UserType = "2" Then
Response.Redirect("ASView.asp?Item=4&ASID=" & UserID)
End If

%>