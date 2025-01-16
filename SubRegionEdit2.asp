<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
RID = Request.Form("RID")
SRID = Request.Form("SRID")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM SubRegions where (SubRegionName = '" & Request.Form("RegionName") & "' or SubRegionCode = '" & Request.Form("RegionCode") & "') and SubRegionActive = 'True' and RID = " & RID & "  and SRID <> " & SRID
Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
'Response.end
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! A Sub Region Already exists in the system, with either the same Name or Region Code.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SubRegions where SRID = " & SRID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("SubRegionName") = Request.Form("RegionName")
rstSecond("SubRegionCode") = Replace(Request.Form("RegionCode"), " ", "")
rstSecond("LastChangedDate") = Now()
rstSecond("HeadCountTarget") = Request.Form("HeadCountTarget")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

%>
<!-- #include file="Includes/UpdateuserregionsAutomatic.inc" -->
<%

Response.Redirect("SubRegions.asp?RID=" & RID)
%>
