<!-- #include file="Connections/Site.asp" -->
<%

AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
RID = Request.Form("RID")
CompanyID = Request.Form("CompanyID")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM Regions where (RegionName = '" & Request.Form("RegionName") & "' or RegionCode = '" & Request.Form("RegionCode") & "') and Active = 'Yes' and CompanyID = " & CompanyID & " and RID <> " & RID
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
      window.alert ("Error ! A Region Already exists in the system, with either the same Name or Region Code.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM Regions where RID = " & RID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("RegionName") = Request.Form("RegionName")
rstSecond("RegionCode") = Replace(Request.Form("RegionCode"), " ", "")
rstSecond("RegionalManager") = Request.Form("RegionalManager")
rstSecond("LastChangedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

%>
<!-- #include file="Includes/UpdateuserregionsAutomatic.inc" -->
<%

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID)
%>
