Option Explicit

Dim url: url = "https://www.foobar"
Dim ie
Set ie = Create Object("InternetExplorer.Application")
ie.navigate2(url)
ie.Visible = True
ie.document.getElementById("q").value = "hoge"
Dim elm
Set elm = ie.document.getElementById("sb_form_go")
elm.click

' 処理を終了する
WScript.Quit(100)

Sub WaitIE(IE)
    Do While IE.Busy Or IE.ReadyState <> 4
	    WScript.Sleep(1000)
	Loop
End Sub
