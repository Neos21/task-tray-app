Option Explicit

' ================================================================================
' Task Tray App Executor
' ================================================================================

' `0` で非表示起動してしまうとウィンドウを表示できなくなる・そこで `7` 最小化状態での起動を利用する https://admhelp.microfocus.com/uft/en/all/VBScript/Content/html/6f28899c-d653-4555-8a59-49640b0e32ea.htm
Const windowStyle = 7
' 本ファイルと同じディレクトリにある `.ps1` ファイルを使用する
Dim psFilePath : psFilePath = Replace(WScript.ScriptFullName, ".vbs", ".ps1")
' 第3引数に `True` を与えると PowerShell が終了するまで WSH 側も待機する (デフォルトは待機せず WSH を終了する `False` と同じ)
CreateObject("Wscript.Shell").run "powershell -NoLogo -NoProfile -ExecutionPolicy Unrestricted -File " & Chr(34) & psFilePath & Chr(34), windowStyle
