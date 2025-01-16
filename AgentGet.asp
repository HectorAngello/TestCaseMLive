<%@ Language="VBScript" %>
<!--#include file="Connections/Site.asp" -->
<%

If Len(request.querystring("q")) > 3 Then
	sql="SELECT * FROM ViewMyRegionTedis where UserID = " & Session("UNID") & " and  ((TediEmpCode Like '%" & request.querystring("q") & "%')) and TediActive = 'True' Order By TediEmpCode Asc"
'Response.write(sql)	
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open(MM_Site_STRING)
	set rs = Server.CreateObject("ADODB.recordset")
	rs.Open sql, conn
	
	if not rs.EOF then
	response.write("<select name=AllocateTo>")
		do until rs.EOF
		response.write("<option value=" & rs("TID") & ">" & rs("TediEmpCode") & " - " & rs("TediFirstName") & " " & rs("TediFirstName") & "</option>")
		rs.MoveNext
		loop
	response.write("</select>")
	response.write("<INPUT TYPE=SUBMIT VALUE=Allocate  class='orange nice button radius'>")
	else
		response.write "Agent Not found.<br>"
	end if

end if%>