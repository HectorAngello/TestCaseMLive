<!-- #include file="Connections/Site.asp" -->
<%
GroupID = Request.Form("GroupID")
AppCat = Request.Form("AppCat")
AppSubCatID = Request.Form("AppSubCatID")

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM SecurityGroupLinks where GroupID = " + Replace(GroupID, "'", "''") + ""
rstSecond.Open
set rstSecond = nothing

set RecRolesItems = Server.CreateObject("ADODB.Recordset")
RecRolesItems.ActiveConnection = MM_Site_STRING
RecRolesItems.Source = "Select * FROM SecurityGroupItems"
RecRolesItems.CursorType = 0
RecRolesItems.CursorLocation = 2
RecRolesItems.LockType = 3
RecRolesItems.Open()
RecRolesItems_numRows = 0
While Not RecRolesItems.EOF
UpdateItem = "ItemID" & RecRolesItems.Fields.Item("ItemID").Value
If Request.Form(UpdateItem) = "Yes" Then

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SecurityGroupLinks", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("GroupID") = GroupID
rstSecond("ItemID") = RecRolesItems.Fields.Item("ItemID").Value
rstSecond("LinkAddedDate") = Now()
rstSecond("LinkAddedBy") = Session("UNID")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

End If
RecRolesItems.MoveNext
Wend

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&GroupID=" & GroupID)
%>
