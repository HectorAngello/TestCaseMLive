<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0
%>
<script type="text/javascript">
function Ajax(){

var xmlHttpWL;
	try{	
		xmlHttpWL=new XMLHttpRequest();// Firefox, Opera 8.0+, Safari
	}
	catch (e){
		try{
			xmlHttpWL=new ActiveXObject("Msxml2.XMLHTTP"); // Internet Explorer
		}
		catch (e){
		    try{
				xmlHttpWL=new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e){
				alert("No AJAX!?");
				return false;
			}
		}
	}

xmlHttpWL.onreadystatechange=function(){
	if(xmlHttpWL.readyState==4){
		document.getElementById('Watchlist').innerHTML=xmlHttpWL.responseText;
		//setTimeout('Ajax()',1000);
	}
}
xmlHttpWL.open("GET","Dashboard-Watchlist.asp",true);
xmlHttpWL.send(null);
<%
SystemItem = "224"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%>
var xmlHttp3;
	try{	
		xmlHttp3=new XMLHttpRequest();// Firefox, Opera 8.0+, Safari
	}
	catch (e){
		try{
			xmlHttp3=new ActiveXObject("Msxml2.XMLHTTP"); // Internet Explorer
		}
		catch (e){
		    try{
				xmlHttp3=new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e){
				alert("No AJAX!?");
				return false;
			}
		}
	}

xmlHttp3.onreadystatechange=function(){
	if(xmlHttp3.readyState==4){
		document.getElementById('NotBankedList').innerHTML=xmlHttp3.responseText;
		
	}
}

xmlHttp3.open("GET","Dashboard-NotBankedList.asp",true);
xmlHttp3.send(null);
<%
End If
%>
//setTimeout('Ajax()',120000);
}


window.onload=function(){
	Ajax();
}
</script>
<!-- header -->
   <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
            </div>

            <div class="nine columns">
                <div class="content panel">
		 <h1>Dashboard:</h1>
<!-- #include file="includes/DashboardChart.inc" -->
</div>
</div>
<div class="twelve columns">
                <div class="content panel">
<!-- #include file="includes/DashRegBreakdownMcharge.inc" -->

<%
smsoutstanding = 0
set RecSMSCount = Server.CreateObject("ADODB.Recordset")
RecSMSCount.ActiveConnection = MM_Site_STRING
RecSMSCount.Source = "Select COUNT(SMSID) AS SMSC FROM SMSCommunications WHERE (IsSent = 'False')"
RecSMSCount.CursorType = 0
RecSMSCount.CursorLocation = 2
RecSMSCount.LockType = 3
RecSMSCount.Open()
RecSMSCount_numRows = 0
smsoutstanding = RecSMSCount.Fields.Item("SMSC").Value
%>
<h3>SMSs currently waiting to be sent out: <%=smsoutstanding%></h3>
<%
SystemItem = "223"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%>
<h3>Agents in My Regions on the Watch List</h3>
<div id="Watchlist"><img src="Images/pageloading3.gif"></div>
<%
End If
SystemItem1 = "224"
SystemItem2 = "2307"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where (ItemID = " & SystemItem1 & " or ItemID = " & SystemItem2 & ") and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%>
<div id="NotBankedList"><img src="Images/pageloading3.gif"></div>
<%
End If
%>

		</div>

			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

