<!doctype html>
<%
TediPic = Request.Querystring("ASPic")
%>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ImagePicker Demo</title>

    <!-- CSS -->
    <link rel="stylesheet" href="ImagePicker/assets/css/demo.css">
    <link rel="stylesheet" href="ImagePicker/assets/css/bootstrap.css">
    <link rel="stylesheet" href="ImagePicker/assets/css/imgpicker.css">

    <!-- JavaScript -->
    <script src="ImagePicker/assets/js/jquery-1.11.0.min.js"></script>
    <script src="ImagePicker/assets/js/jquery.Jcrop.min.js"></script>
    <script src="ImagePicker/assets/js/jquery.imgpicker.js"></script>

</head>
<body>
   
    <div class="main">
        <div class="container"><div class="box">
            <div class="content clearfix">
                <img src="ASimages/<%=Request.Querystring("ASPic")%>" id="avatar2" width="150"><br>

                <!-- Inline version -->
                <div id="avatarInline">
                    <div class="btn btn-primary ip-upload">Upload <input type="file" name="file" class="ip-file"></div>
                    <button type="button" class="btn btn-primary ip-webcam">Webcam</button>
                    <!-- <button type="button" class="btn btn-info ip-edit">Edit</button>
                    <button type="button" class="btn btn-danger ip-delete">Delete</button> -->

                    <div class="alert ip-alert"></div>
                    <div class="ip-info">To crop this image, drag a region below and then click "Save Image"</div>
                    <div class="ip-preview"></div>
                    <div class="ip-rotate">
                        <button type="button" class="btn btn-default ip-rotate-ccw" title="Rotate counter-clockwise"><i class="icon-ccw"></i></button>
                        <button type="button" class="btn btn-default ip-rotate-cw" title="Rotate clockwise"><i class="icon-cw"></i></button>
                    </div>
                    <div class="ip-progress">
                        <div class="text">Uploading</div>
                        <div class="progress progress-striped active"><div class="progress-bar"></div></div>
                    </div>
                    <div class="ip-actions">
                        <button type="button" class="btn btn-success ip-save">Save Image</button>
                        <button type="button" class="btn btn-primary ip-capture">Capture</button>
                        <button type="button" class="btn btn-default ip-cancel">Cancel</button>
                    </div>
                </div>
                <!-- end Inline -->

            </div>
        </div></div>
    </div>

    <script>
        $(function() {
            var time = function(){return'?'+new Date().getTime()};

            // Avatar setup
            $('#avatarInline').imgPicker({
                url: 'server/upload_avatarAS.php',
                aspectRatio: 1,
                deleteComplete: function() {
                    $('#avatar2').attr('src', '/ASimages/<%=Request.Querystring("ASPic")%>');
                    this.modal('hide');
                },
		data: function() {
		return {
			tid: <%=Request.Querystring("ASID")%>,
			};
		},
                cropSuccess: function(image) {
                    $('#avatar2').attr('src', image.versions.avatar.url +time() );
		    window.parent.top.location = "UpdateASPic.asp?ASID=<%=Request.Querystring("ASID")%>&NewPic=" + image.versions.avatar.url;
                    this.modal('hide');
                }
            });

            // Demo only
            $('.navbar-toggle').on('click',function(){$('.navbar-nav').toggleClass('navbar-collapse')});
            $(window).resize(function(e){if($(document).width()>=430)$('.navbar-nav').removeClass('navbar-collapse')});
        });
    </script>
</body>
</html>
