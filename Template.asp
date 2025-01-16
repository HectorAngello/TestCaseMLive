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

                        <div class="eight columns"><h1>Pages</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Page Name</th>
                                <th>Page Description</th>
                                <th></th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Home</td>
                                <td>This is the Home page.</td>
                                <td class="action-td"><a href="#" class="view-button"></a></td>
                                <td class="action-td"><a href="#" class="edit-button"></a></td>
                                <td class="action-td"><a href="#" class="delete-button"></a></td>
                            </tr>
                            <tr>
                                <td>About</td>
                                <td>This is the About page.</td>
                                <td class="action-td"><a href="#" class="view-button"></a></td>
                                <td class="action-td"><a href="#" class="edit-button"></a></td>
                                <td class="action-td"><a href="#" class="delete-button"></a></td>
                            </tr>
                            <tr>
                                <td>Products</td>
                                <td>This is the Products page.</td>
                                <td class="action-td"><a href="#" class="view-button"></a></td>
                                <td class="action-td"><a href="#" class="edit-button"></a></td>
                                <td class="action-td"><a href="#" class="delete-button"></a></td>
                            </tr>
                            <tr>
                                <td>Services</td>
                                <td>This is the Services page.</td>
                                <td class="action-td"><a href="#" class="view-button"></a></td>
                                <td class="action-td"><a href="#" class="edit-button"></a></td>
                                <td class="action-td"><a href="#" class="delete-button"></a></td>
                            </tr>
                            <tr>
                                <td>Contact</td>
                                <td>This is the contact page.</td>
                                <td class="action-td"><a href="#" class="view-button"></a></td>
                                <td class="action-td"><a href="#" class="edit-button"></a></td>
                                <td class="action-td"><a href="#" class="delete-button"></a></td>
                            </tr>
                        </tbody>
                    </table>
<!-- #include file="includes/footer.asp" -->

