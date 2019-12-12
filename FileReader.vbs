Option Explicit
On Error Resume Next

'https://teratail.com/questions/25245
' ブラウザ　自動　操作

Dim objFSO      ' FileSystemObject
Dim objFile     ' ファイル読み込み用
'カレントパス取得
Dim objPath
Set objPath = CreateObject("Scripting.FileSystemObject").GetFolder(".")


Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    For Each objFile In objPath.Files
	Dim FileName
	FileName    =    objFile.Name
	If InStr(FileName, ".txt") > 0 Then
	FileName = FileName + ".txt"
    'Set objFile = objFSO.OpenTextFile("test.txt")
	Set objFile = objFSO.OpenTextFile(FileName)
    If Err.Number = 0 Then
	    Dim fileContent
        Do While objFile.AtEndOfStream <> True
            'WScript.Echo objFile.ReadLine
			fileContent = objFile.ReadLine
        Loop
        objFile.Close
		
    Else
        WScript.Echo "ファイルオープンエラー: " & Err.Description
    End If
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objFile = Nothing
Set objFSO = Nothing
