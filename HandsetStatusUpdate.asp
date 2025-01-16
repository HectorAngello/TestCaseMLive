<!-- #include file="Connections/Site.asp" -->
<%
HID = Request.Form("HID")
StaID = Request.Form("StaID")
TediID = Request.Form("TediID")

If (TediID = 0) and (StaID = 2) Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Please select a Tedi to allocate this handset to.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

If (TediID <> 0) and (StaID = 1) Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! Please select 'Not Allocated To A Agent, to update the handset status to 'Available For Allocation'.");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM HandsetUpdates", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("HID") = HID
rstSecond("StaID") = StaID
rstSecond("UpdateBy") = Session("UNID")
rstSecond("UpDateDate") = Now()
rstSecond("TediID") = TediID
rstSecond("Notes") = Request.Form("Notes")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Handsets where HID = " & HID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("HandsetStatus") = StaID
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.redirect("Updated.asp?AppCat=7&AppSubCatID=36&ItemID=215&HID=" & HID)
%>