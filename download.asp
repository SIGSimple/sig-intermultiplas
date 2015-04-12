<%
	Dim filename, path
		filename = Request.QueryString("filename")
		path = Request.QueryString("path")

	Response.Buffer = True
	Response.AddHeader "Content-Type","application/x-msdownload"
	Response.AddHeader "Content-Disposition","attachment; filename=" & filename
	Response.Flush

	Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1
		objStream.LoadFromFile Request.QueryString("path") & Request.QueryString("filename")

	Response.BinaryWrite objStream.Read

	objStream.Close

	Set objStream = Nothing

	Response.Flush
%>