<!-- #include file="Connections/Site.asp" -->
<%
ID = Request.Form("ID")
AirtimeTarget = Request.Form("AirtimeTarget")
'AirtimeTarget = Replace(Request.Form("AirtimeTarget"), ".", ",")
ConnectionsTarget = Request.Form("ConnectionsTarget")
'ConnectionsTarget = Replace(Request.Form("ConnectionsTarget"), ".", ",")
ActivationsTarget = Request.Form("ActivationsTarget")
'ActivationsTarget = Replace(Request.Form("ActivationsTarget"), ".", ",")
DataTarget = Request.Form("DataTarget")
'DataTarget = Replace(Request.Form("DataTarget"), ".", ",")
PortsTarget = Request.Form("PortsTarget")
'PortsTarget = Replace(Request.Form("PortsTarget"), ".", ",")
ErrorMSG = ""

If Isnumeric(AirtimeTarget) = "False" then
ErrorMSG = "Airtime Target"
End If

If Isnumeric(ConnectionsTarget) = "False" then
ErrorMSG = "Connections Target"
End If

If Isnumeric(ActivationsTarget) = "False" then
ErrorMSG = "Activations Target"
End If

If Isnumeric(DataTarget) = "False" then
ErrorMSG = "Data Target"
End If

If Isnumeric(PortsTarget) = "False" then
ErrorMSG = "Ports Target"
End If

If ErrorMSG <> "" Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! <%=ErrorMSG%> Needs to be a numeric value");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM  MonthlyTargets where ID = " & ID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("AirtimeTarget") = AirtimeTarget
rstSecond("ConnectionsTarget") = ConnectionsTarget
rstSecond("ActivationsTarget") = ActivationsTarget
rstSecond("DataTarget") = DataTarget
rstSecond("PortsTarget") = PortsTarget
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

Response.redirect("Updated.asp?AppCat=7&AppSubCatID=42")
%>