Option Explicit

' ================================================================================
' Task Tray App Executor
' ================================================================================

' `0` �Ŕ�\���N�����Ă��܂��ƃE�B���h�E��\���ł��Ȃ��Ȃ�E������ `7` �ŏ�����Ԃł̋N���𗘗p���� https://admhelp.microfocus.com/uft/en/all/VBScript/Content/html/6f28899c-d653-4555-8a59-49640b0e32ea.htm
Const windowStyle = 7
' �{�t�@�C���Ɠ����f�B���N�g���ɂ��� `.ps1` �t�@�C�����g�p����
Dim psFilePath : psFilePath = Replace(WScript.ScriptFullName, ".vbs", ".ps1")
' ��3������ `True` ��^����� PowerShell ���I������܂� WSH �����ҋ@���� (�f�t�H���g�͑ҋ@���� WSH ���I������ `False` �Ɠ���)
CreateObject("WScript.Shell").Run "powershell -NoLogo -NoProfile -ExecutionPolicy Unrestricted -File " & Chr(34) & psFilePath & Chr(34), windowStyle
