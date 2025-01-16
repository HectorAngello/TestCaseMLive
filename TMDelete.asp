<!-- #include file="Connections/Site.asp" -->
<%
TrainID = Request.Form("TrainID")


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM TrainingMaterial where TrainID = " & TrainID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TrainingActive") = "False"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=1043")
%>
