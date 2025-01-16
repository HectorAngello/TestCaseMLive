<!-- #include file="Connections/Site.asp" -->
<%
PeriodMonth = Request.Form("PeriodMonth")
PeriodYear = Request.Form("PeriodYear")
AirtimeTarget = Replace(Request.Form("AirtimeTarget"), ".", ",")
ConnectionsTarget = Replace(Request.Form("ConnectionsTarget"), ".", ",")
ActivationsTarget = Replace(Request.Form("ActivationsTarget"), ".", ",")
DataTarget = Replace(Request.Form("DataTarget"), ".", ",")
PortsTarget = Replace(Request.Form("PortsTarget"), ".", ",")

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

set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select Top(1)* FROM MonthlyTargets where PeriodMonth = " & PeriodMonth & " and PeriodYear = " & PeriodYear
'Response.Write(RecCheck.Source)
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
%>
      <script language="javascript">
      <!--
      window.alert ("Error ! The selected period already has a target set, rather edit the period");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM  MonthlyTargets", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("PeriodMonth") = PeriodMonth
rstSecond("PeriodYear") = PeriodYear
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