<!-- #include file="Connections/Site.asp" -->
<%
ASID = Request.Form("ASID")
TID = Request.Form("TID")
SRID = Request.Form("SRID")

set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT * FROM SubRegions Where SRID = " & SRID
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
HCLimit = RecSubRegion.Fields.Item("HeadCountTarget").Value
CurrentCount = 0
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and TediParent = '0' and SRID = " & SRID
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0
While Not RecEdit.EOF
CurrentCount = CurrentCount + 1
RecEdit.MoveNext
Wend

If CurrentCount >= HCLimit Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! This Sub Region Has Already Reached It's Head Count Target of <%=HCLimit%>");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

set RecAS = Server.CreateObject("ADODB.Recordset")
RecAS.ActiveConnection = MM_Site_STRING
RecAS.Source = "SELECT * FROM ASs Where ASID = " & ASID
RecAS.CursorType = 0
RecAS.CursorLocation = 2
RecAS.LockType = 3
RecAS.Open()
RecAS_numRows = 0

set RecReg = Server.CreateObject("ADODB.Recordset")
RecReg.ActiveConnection = MM_Site_STRING
RecReg.Source = "SELECT * FROM Regions Where RID = " & RecAS.Fields.Item("RID").Value
RecReg.CursorType = 0
RecReg.CursorLocation = 2
RecReg.LockType = 3
RecReg.Open()
RecReg_numRows = 0

set RecSubReg = Server.CreateObject("ADODB.Recordset")
RecSubReg.ActiveConnection = MM_Site_STRING
RecSubReg.Source = "SELECT * FROM SubRegions Where SRID = " & SRID
RecSubReg.CursorType = 0
RecSubReg.CursorLocation = 2
RecSubReg.LockType = 3
RecSubReg.Open()
RecSubReg_numRows = 0

TediNewEmpCode = RecReg.Fields.Item("RegionCode").Value & RecSubReg.Fields.Item("SubRegionCode").Value & TID

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Tedis where TID = " & TID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TediEmpCode") = TediNewEmpCode
rstSecond("TediParent") = "0"
rstSecond("SRID") = SRID
rstSecond("ASID") = ASID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


TediUpdateType = "Sub Agent Up-graded"
%><!-- #include file="Includes/TediAudit-Update.inc" -->
<%

response.redirect("TediView.asp?TID=" & TID)
%>