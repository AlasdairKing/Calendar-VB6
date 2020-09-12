Attribute VB_Name = "modCreateShortcut"
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

'modCreateShortcut
'Requires a reference to wshom.ocx and Windows Scripting.
'Alasdair King, September 2009.
'http://www.vbforums.com/showthread.php?t=234891

Public Sub CreateShortCut(ByVal sTargetPath As String, ByVal sShortCutFolder As String, ByVal sShortCutName As String, Optional ByVal sArguments As String, Optional ByVal sWorkPath As String, Optional ByVal eWinStyle As WshWindowStyle = vbNormalFocus, Optional ByVal iIconNum As Integer)
    On Error Resume Next
    ' Requires reference to Windows Script Host Object Model
    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim oShortCut As IWshRuntimeLibrary.WshShortcut

    Set oShell = New IWshRuntimeLibrary.WshShell
    Set oShortCut = oShell.CreateShortCut(sShortCutFolder & "\" & sShortCutName & ".lnk")
    With oShortCut
        .TargetPath = sTargetPath
        .arguments = sArguments
        .WorkingDirectory = sWorkPath
        .WindowStyle = eWinStyle
        .IconLocation = sTargetPath & "," & iIconNum
        .save
    End With
    Set oShortCut = Nothing
    Set oShell = Nothing
End Sub
