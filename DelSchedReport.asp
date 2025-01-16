<!-- #include file="Connections/Site.asp" -->
<!-- #include file="Includes/MySubRegions.inc" -->
<%
ReportID = Request.QueryString("ReportID")

set RecWhichOne = Server.CreateObject("ADODB.Recordset")
RecWhichOne.ActiveConnection = MM_Site_STRING
RecWhichOne.Source = "SELECT * FROM ScheduledReports where ReportID = " & ReportID
'Response.write(RecDeductions.Source)
RecWhichOne.CursorType = 0
RecWhichOne.CursorLocation = 2
RecWhichOne.LockType = 3
RecWhichOne.Open()
RecWhichOne_numRows = 0
If RecWhichOne.Fields.Item("ReportStatusID").Value = "1" Then

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.ActiveConnection = MM_Site_STRINGWrite
rstSecond.Source = "Delete FROM ScheduledReports Where ReportID = " & ReportID
rstSecond.Open
set rstSecond = nothing	

Response.redirect("Display.asp?AppCat=10&AppSubCatID=41")
Else
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Report already processing - Unable to remove this report.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If
%>