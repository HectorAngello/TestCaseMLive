﻿<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
%>
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
		<div class="row">
		<div class="twelve columns">
			<div class="alert-box success">System Updated</div>
		</div>
		</div>
		<!-- #include file="Includes/ContentSelect.inc" -->

                   
<!-- #include file="includes/footer.asp" -->

