<%
Response.Buffer = true
Server.ScriptTimeout = 18000
On Error Resume Next
'---------------------------------------------------------------------------------------------------------
'declare all variables at top of page is best practice
'---------------------------------------------------------------------------------------------------------
Dim strFullFileName,strSplitArray,intUpper,strFileName,strPath,strServerName,strServerRoot,strFinalPath,strFinalDownload,strAbsFile,objFSO,objFile,objStream



'---------------------------------------------------------------------------------------------------------
'set variables
'---------------------------------------------------------------------------------------------------------
strFileNameReal = CStr(Request.QueryString("filename"))
strFullFileName = CStr(Request.QueryString("path")) & "/" & CStr(Request.QueryString("filename"))'get full href and convert to string
strSplitArray = Split(strFullFileName,"/")'split above string at each forward slash
intUpper = Ubound(strSplitArray)'takes last part of split which will be the last number ie: the file name
strFileName = SplitArray(intUpper)'file name is last part of url
strPath = Replace(strFullFileName,strFileName,"")'gets the path to the image
strServerName = Request.ServerVariables("Server_Name")'gets domain url ie: www.yourdomain.com
strServerRoot = Server.MapPath("\")'gets the proper path to the root of the server
strFinalPath = Replace(Replace(Replace(Replace(strFullFileName,strServerName,""),strFileName,""),"http://",""),"/","\")'leaves me with the path to the image
strFinalDownload = strFinalPath&strFileName'add our new path and file name together and we have what we needed!



'---------------------------------------------------------------------------------------------------------
'do some basic error checking for the QueryString
'---------------------------------------------------------------------------------------------------------
If strPath = "" Then
	Response.Clear
	Response.Write("No file specified.")
	Response.End
ElseIf InStr(strPath, "..") > 0 Then
	Response.Clear
	Response.Write("Illegal folder location.")
	Response.End
ElseIf Len(strPath) > 1024 Then
	Response.Clear
	Response.Write("Folder path too long.")
	Response.End
Else
	Call DownloadFile(strFinalDownload)
End If



'---------------------------------------------------------------------------------------------------------
'now call the function that does all the work
'---------------------------------------------------------------------------------------------------------
Private Sub DownloadFile(file)

	'set absolute file location which our new path from the websites root (thats what all the work above was for)
	strAbsFile = Server.MapPath("\") & strFinalDownload
	
	'create FSO object to check if file exists and get properties
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'check to see if the file exists
	If (objFSO.FileExists(strAbsFile)) Then
		'Set the content type to the specific type that you are sending.
		Response.ContentType = "application/octet-stream"
		Response.AddHeader "Content-Disposition", "attachment; filename=" & strFileNameReal

		set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1
		objStream.LoadFromFile(strAbsFile)

		Response.BinaryWrite objStream.Read

		objStream.Close
		Set objStream = nothing
	Else
		Response.Clear
		Response.Write("No such file exists.")
	End If

	'release memory
	objFSO.Close
	Set objFSO = Nothing
End Sub
%>