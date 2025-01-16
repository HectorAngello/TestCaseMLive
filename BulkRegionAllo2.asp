<!-- #include file="Connections/Site.asp" -->
<%

UserID = Request.Form("UserID")
RID = Request.Form("RID")

set RecClearSubRegions = Server.CreateObject("ADODB.Recordset")
RecClearSubRegions.ActiveConnection = MM_Site_STRING
RecClearSubRegions.Source = "SELECT * FROM RegionSubRegion where  RID = " & RID
RecClearSubRegions.CursorType = 0
RecClearSubRegions.CursorLocation = 2
RecClearSubRegions.LockType = 3
RecClearSubRegions.Open()
RecClearSubRegions_numRows = 0
While Not RecClearSubRegions.EOF
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM UserRegion where UserID = " & UserID & " and SRID = " & RecClearSubRegions.Fields.Item("SRID").Value
'response.write(rstSecond.Source)
'response.end
rstSecond.Open
set rstSecond = nothing
RecClearSubRegions.MoveNext
Wend

set RecSubRegions = Server.CreateObject("ADODB.Recordset")
RecSubRegions.ActiveConnection = MM_Site_STRING
RecSubRegions.Source = "SELECT * FROM SubRegions where SubRegionActive = 'True' and RID = " & RID & " Order By SubRegionName Asc"
RecSubRegions.CursorType = 0
RecSubRegions.CursorLocation = 2
RecSubRegions.LockType = 3
RecSubRegions.Open()
RecSubRegions_numRows = 0
While Not RecSubRegions.EOF
UpdateItem = "SRID" & RecSubRegions.Fields.Item("SRID").Value
If Request.Form(UpdateItem) = "Yes" Then

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM UserRegion", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("SRID") = RecSubRegions.Fields.Item("SRID").Value
rstSecond("UserID") = UserID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

End If
RecSubRegions.MoveNext
Wend


Response.Redirect("updated.asp?AppCat=3&AppSubCatID=1")
%>
