<!-- #include file="Connections/Site.asp" -->
<!-- #include file="Includes/MySubRegions.inc" -->
<%
ID = Request.QueryString("ID")

set RecWhichOne = Server.CreateObject("ADODB.Recordset")
RecWhichOne.ActiveConnection = MM_Site_STRING
RecWhichOne.Source = "SELECT * FROM ViewTediReconDetails Where ID = " & ID & " and SRID in (" & SRRegionList & ") and ReconActive = 'True' and CompanyID = " & Session("CompanyID")
'Response.write(RecDeductions.Source)
RecWhichOne.CursorType = 0
RecWhichOne.CursorLocation = 2
RecWhichOne.LockType = 3
RecWhichOne.Open()
RecWhichOne_numRows = 0
If Not RecWhichOne.EOF and Not RecWhichOne.BOF Then
TID = RecWhichOne.Fields.Item("TID").Value
DelID = RecWhichOne.Fields.Item("ID").Value

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM TediRecons where ID = " & DelID & " and TID = " & TID , MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("ReconActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing
Response.redirect("TediView.asp?TID=" & TID & "&Item=7")
End If
Response.redirect("Dashboard.asp")
%>