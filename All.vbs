Option Explicit
' On Error Resume Next

Dim url: url = "https://www.foobar"
Dim ie
Set ie = Create Object("InternetExplorer.Application")
ie.navigate2(url)
ie.Visible = True
Call WaitIE(ie)

Dim objFSO
Dim objFile

Dim objPath
Set objPath = CreateObject("Scripting FilesystemObject").GetFolder(".")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

If Err.Number = 0 Then
    Dim textContent
	For Each objFile In objPath.Files
	Dim FileName
	FileName = objFile.Name
	If InStr(FileName, ".txt") > 0 Then
	    Set objFile = objFSO.OpenTextFile(FileName)
		If Err.Number = 0 Then
		    ' Dim textContent: textContent = ""
			Do While objFile.AtEndOfStream <> True
			    ' WScript.Echo objFile.ReadLine 
				textContent = textContent + objFile.ReadLine & vbCrLf
			Loop
			obj.FileClose
		Else
		    WScript.Echo "ファイルオープンエラー：" & Err.Description
		End If
	Next
	' ie.document.getElementById("q").value = textContent
	' Dim elm
	' Set elm = ie.document.getElementById("sb_form_go")
	' elm.click
	WScript.Echo textContent
Else
     WScript.Echo "エラー：" & Err.Description
End If
			
Set objFile = Nothing
Set objFSO = Nothing

' 処理を終了する
WScript.Quit(100)



Sub WaitIE(IE)
    Do While IE.Busy Or IE.ReadyState <> 4
	    WScript.Sleep(1000)
	Loop
End Sub


