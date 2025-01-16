<!-- #include file="includes/header.asp" -->
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

                        <div class="eight columns"><h1>M-Charge Management</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br><br>
<%If Request.QueryString("LineCount") <> "0" Then%>
<b>File Upload Complete: <%=Request.QueryString("LineCount")%> Items Imported</b>
<br>System busy with allocation process in the background, depending on the amount of items imported,<br>this can take a couple of minutes to complete.
<%If Request.QueryString("ImportType") = "M-Charge" Then%>
<script language="javascript">      
       //Create an iframe and turn on the design mode for it 
       document.write ('<iframe src="http://pmg2.mtnlive.co.za/try.asp" id="Abstract" frameborder="0" width="1" height="1"></iframe>')
frames.Abstract.document.designMode = "off";               
 </script>
<%End If%>
<%Else%>
No New Line Items Imported.
<%End If%>
                    </div>
<!-- #include file="includes/footer.asp" -->

