<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")
TermedBy = Request.Form("TermedBy")
TermReason = Request.Form("TermReason")
ASTermDate = Request.Form("ASTermDate")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ASs Where ASID = " & ASID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("ASTermDate") = ASTermDate
rstSecond("ASTermReason") = TermReason
rstSecond("ASActive") = "False"
rstSecond("DeletedBy") = TermedBy
rstSecond("DeletedDate") = Now()
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

ASUpDateID = ASID
ASUpdateType = SupervisorLabel & " Deleted"
%><!-- #include file="Includes/ASAudit-Update.inc" -->
<%



Response.Redirect("DashBoard.asp")

%>