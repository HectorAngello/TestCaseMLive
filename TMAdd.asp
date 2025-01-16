<!-- #include file="Connections/Site.asp" -->
<%
UNID = Request.Form("UNID")


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM TrainingMaterial", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("TrainingName") = Request.Form("TrainingName")
rstSecond("TraingDescription") = Request.Form("TraingDescription")
rstSecond("TrainingActive") = "True"
rstSecond("TrainingDoc") = Request.Form("TrainingDoc")
rstSecond("TrainingDoc2") = Request.Form("TrainingDoc2")
rstSecond("TrainingDoc3") = Request.Form("TrainingDoc3")
rstSecond("TrainingDoc4") = Request.Form("TrainingDoc4")
rstSecond("TrainingDoc5") = Request.Form("TrainingDoc5")

rstSecond("TrainingAudienceID") = Request.Form("TrainingAudienceID")
rstSecond("AddedBy") = UNID
rstSecond("AddedDate") = Now()

rstSecond.Update
rstSecond.Close
set rstSecond = nothing


Response.Redirect("Updated.asp?AppCat=7&AppSubCatID=1043")
%>
