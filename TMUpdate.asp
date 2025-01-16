<!-- #include file="Connections/Site.asp" -->
<%
TrainID = Request.Form("TrainID")


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM TrainingMaterial where TrainID = " & TrainID, MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("TrainingName") = Request.Form("TrainingName")
rstSecond("TraingDescription") = Request.Form("TraingDescription")
rstSecond("TrainingDoc") = Request.Form("TrainingDoc")
rstSecond("TrainingDoc2") = Request.Form("TrainingDoc2")
rstSecond("TrainingDoc3") = Request.Form("TrainingDoc3")
rstSecond("TrainingDoc4") = Request.Form("TrainingDoc4")
rstSecond("TrainingDoc5") = Request.Form("TrainingDoc5")
rstSecond("TrainingAudienceID") = Request.Form("TrainingAudienceID")
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=1043")
%>
