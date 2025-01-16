<!-- #include file="Connections/Site.asp" -->
<%
SimNumber = Request.Form("SimNumber")
PUKNumber = Request.Form("PUKNumber")
IMEI = Request.Form("IMEI")
CompanyID = Request.Form("CompanyID")


set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "Select * FROM Handsets where (SimNumber = '" & SimNumber & "' or PUKNumber = '" & PUKNumber & "' or IMEI = '" & IMEI & "') and CompanyID = " & CompanyID
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
      window.alert ("Error ! An Handset Already exists in the system, with either the same Sim Number, IMEI Number or PUK Number");
      window.history.go(-1);
      //-->
      </script>
      <%
      Response.End
End If


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM Handsets", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("Handset") = Request.Form("Handset")
rstSecond("IMEI") = Request.Form("IMEI")
rstSecond("HandsetAcitive") = "True"
rstSecond("SimNumber") = Request.Form("SimNumber")
rstSecond("PUKNumber") = Request.Form("PUKNumber")
rstSecond("Battery") = Request.Form("Battery")
rstSecond("Charger") = Request.Form("Charger")
rstSecond("HandsFree") = Request.Form("HandsFree")
rstSecond("CompanyID") = Request.Form("CompanyID")
rstSecond("AddedBy") = Session("UNID")
rstSecond("AddedDate") = Now()
rstSecond("HandsetStatus") = "1"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

set RecNewestAgent = Server.CreateObject("ADODB.Recordset")
RecNewestAgent.ActiveConnection = MM_Site_STRING
RecNewestAgent.Source = "SELECT * FROM Handsets Order By HID Desc"
RecNewestAgent.CursorType = 0
RecNewestAgent.CursorLocation = 2
RecNewestAgent.LockType = 3
RecNewestAgent.Open()
RecNewestAgent_numRows = 0
NewestName = RecNewestAgent.Fields.Item("Handset").Value
NewestID = RecNewestAgent.Fields.Item("HID").Value

Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=36&Added=True&HID=" & NewestID & "&Name=" & NewestName)

%>