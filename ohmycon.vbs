' OhMyCon (https://github.com/reallukee/ohmycon)
' Creato da Realluke (https://github.com/reallukee/)

Option Explicit
On Error Resume Next

Dim Path
Dim Color
Dim Fso
Dim Ini
Dim Item

Color = WScript.Arguments(0)
Path = ""

For Item = 1 To WScript.Arguments.Length - 1
    If Item < WScript.Arguments.Length - 1 Then
        Path = Path & WScript.Arguments(Item) & " "
    Else
        Path = Path & WScript.Arguments(Item)
    End If
Next

Set Fso = CreateObject("Scripting.FileSystemObject")

If Fso.FileExists(Path & "\Desktop.ini") Then
    Fso.DeleteFile Path & "\Desktop.ini"
End If

Set Ini = Fso.CreateTextFile(Path & "\Desktop.ini", True)

Ini.Write "[.ShellClassInfo]" & VbCrlf
Ini.Write "IconFile=C:\OhMyCon\" & Color & ".ico" & VbCrlf
Ini.Write "IconIndex=0" & VbCrlf
Ini.Write "IconResource=C:\OhMyCon\" & Color & ".ico,0" & VbCrlf
Ini.Write "[ViewState]" & VbCrlf
Ini.Write "Mode=" & VbCrlf
Ini.Write "Vid=" & VbCrlf
Ini.Write "FolderType=Generic" & VbCrlf

Ini.Close

Set Ini = Fso.GetFile(Path & "\Desktop.ini")

If Not Ini.Attributes And 6 Then
    Ini.Attributes = Ini.Attributes + 6
End If

Set Ini = Fso.GetFolder(Path)

If Not Ini.Attributes And 4 Then
    Ini.Attributes = Ini.Attributes + 4
End If
