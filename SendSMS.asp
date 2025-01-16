<!-- #include file="Connections/Site.asp" -->
<!-- #include file="Includes/Base64Encode.inc" -->
<%
SMSMsg = "Test5"
MyNumber = "0824119202"
If Left(MyNumber,1) = "0" Then
MyNumberT = Len(MyNumber)
MyNumber = Right(MyNumber, MyNumberT - 1)
MyNumber = "27" & MyNumber
End If

SMSURL = "https://sms01.umsg.co.za/xml/send?number=" & MyNumber & "&message=" & SMSMsg

data = SMSURL

APIUser = "u97581"
APIPass = "RNGFwUPA"

set xmldom = server.CreateObject("Microsoft.XMLDOM") 
Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

httpRequest.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
httpRequest.Open "GET", data, False
httpRequest.setRequestHeader "Authorization", "Basic " & Base64Encode(APIUser & ":" & APIPass)
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"


httpRequest.Send

postResponse = httpRequest.ResponseText
'Response.write(postResponse)

set xmldom = httpRequest.responseXML

For Each oNode In xmldom.SelectNodes("/sms/submitresult")
    SentSuccess = oNode.GetAttribute("result")
    SentKey = oNode.GetAttribute("key")
Next

Response.write("SentKey:" & SentKey)
%>