<!-- #include file="Connections/Site.asp" -->
<%
TID = Request.Form("TID")
TermedBy = Request.Form("TermedBy")
TermReason = Request.Form("TermReason")
ASTermDate = Request.Form("ASTermDate")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TediTermDate") = ASTermDate
rstSecond("TediTermReason") = TermReason
rstSecond("TediActive") = "False"
rstSecond("DeletedBy") = TermedBy
rstSecond("DeletedDate") = Now()
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


TediUpdateType = "Agent Deleted"
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%



Response.Redirect("DashBoard.asp")

%>