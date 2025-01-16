<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Site.asp" -->
<%
AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")
RID = Request.Form("RID")

IsIn = "No"
Dim RecChecking
Dim RecChecking_numRows

Set RecChecking = Server.CreateObject("ADODB.Recordset")
RecChecking.ActiveConnection = MM_Site_STRING
RecChecking.Source = "SELECT * FROM SeniorZonerAllocations WHERE UserID = '" & Request.form("SID") & "'"
RecChecking.CursorType = 0
RecChecking.CursorLocation = 2
RecChecking.LockType = 1
RecChecking.Open()

RecChecking_numRows = 0
While Not RecChecking.EOF
IsIn = "Yes"
RecChecking.MoveNext
Wend

If IsIn = "No" Then
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SeniorZonerAllocations", MM_Site_STRING, 1, 2
rstSecond.AddNew
rstSecond("RID") = RID
rstSecond("UserID") = Request.Form("SID")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing
End If

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&ItemID=" & ItemID & "&RID=" & RID)
%>