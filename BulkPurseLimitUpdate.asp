<!-- #include file="Connections/Site.asp" -->
<%
RID = Request.Form("RID")
SRID = Request.Form("SRID")

set RecAgents = Server.CreateObject("ADODB.Recordset")
RecAgents.ActiveConnection = MM_Site_STRING
If SRID = "0" Then
RecAgents.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True'  and RID = " & RID & " Order By TediFirstName Asc"
Else
RecAgents.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True'  and SRID = " & SRID & " Order By TediFirstName Asc"
End If
RecAgents.CursorType = 0
RecAgents.CursorLocation = 2
RecAgents.LockType = 3
RecAgents.Open()
RecAgents_numRows = 0
While Not RecAgents.EOF
CheckMe = "Tedi" & RecAgents.Fields.Item("TID").Value
CurrentPurseLimit = RecAgents.Fields.Item("PurseLimit").Value
If Cstr(CurrentPurseLimit) <> Request.Form(CheckMe) Then
Response.write("<br>" & RecAgents.Fields.Item("TID").Value & " - " & CurrentPurseLimit & " - " & Request.Form(CheckMe))

TID = RecAgents.Fields.Item("TID").Value
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis Where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("PurseLimit") = Request.Form(CheckMe)
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

TediUpdateType = "Agent Purse Limit Updated"
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%


End If
RecAgents.MoveNext
Wend

Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=37")
%>
