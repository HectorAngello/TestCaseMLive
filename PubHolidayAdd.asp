<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->

<%

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(2) * FROM PublicHolidays", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("HolidayName") = Request.Form("HolidayName")
rstSecond("HolidayDate") = Request.Form("HolidayDate")
rstSecond("CompanyID") = Request.Form("CompanyID")
rstSecond("AddedBy") = Session("UNID")
rstSecond("AddedDate") = Now()
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

Response.redirect "Updated.asp?AppCat=7&AppSubCatID=23"
%>