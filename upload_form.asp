<html>
	<head>
		<title>Upload a File</title>
	</head>
	<body>
		<h4>Upload a File</h4>
		<form method="post" enctype="multipart/form-data" action="upload.asp?id=123&folder=NOTA&retUrl=<%=(Request.ServerVariables("URL"))%>">
			<p><input type="file" name="blob"></p>
			<p><input type="submit" name="btnSubmit" value="Upload"></p>
		</form>
	</body>
</html>