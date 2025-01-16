<!-- #include file="Connections/Site.asp" -->
<%
AppSubCatID = Request.Form("AppSubCatID")
AppCat = Request.Form("AppCat")
ItemID = Request.Form("ItemID")
CompanyID = Request.Form("CompanyID")
GroupName = Request.Form("GroupName")

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM SecurityGroups where GroupName = '" & GroupName & "' and GroupActive = '1' and CompanyID = " & CompanyID
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Security Group - <%=GroupName%> Already Exists in the system");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT * FROM SecurityGroups", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("GroupName") = GroupName
rstSecond("GroupAddedBy") = Session("UNID")
rstSecond("GroupAddedDate") = Now()
rstSecond("CompanyID") = CompanyID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

set RecNewest = Server.CreateObject("ADODB.Recordset")
RecNewest.ActiveConnection = MM_Site_STRING
RecNewest.Source = "Select * FROM SecurityGroups where GroupActive = '1' Order By GroupID Desc"
RecNewest.CursorType = 0
RecNewest.CursorLocation = 2
RecNewest.LockType = 3
RecNewest.Open()
RecNewest_numRows = 0

GroupID = RecNewest.Fields.Item("GroupID").Value

Response.Redirect("Updated.asp?AppCat=" & AppCat & "&AppSubCatID=" & AppSubCatID & "&GroupID=" & GroupID)
%>
