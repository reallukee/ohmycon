' OhMyCon (https://github.com/reallukee/ohmycon)
' Creato da Luca Pollicino (https://github.com/reallukee)

Option Explicit
On Error Resume Next

Dim MbTitle
Dim MbText
Dim MbResult
Dim Regedit
Dim Key
Dim Verbs
Dim Item
Dim Fso
Dim Files
Dim File

Set Regedit = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

MbTitle = "OhMyCon Uninstaller"
MbText = "Benvenuto in OhMyCon Uninstaller! Questo programma disinstaller" & chr(224) & " OhMyCon sul dispositivo in uso."
MbText = MbText & "Questo programma funziona solo con la versione 0.0.1 di OhMyCon."
MbText = MbText & VbCrlf & VbCrlf & "Si vuole procedere con la disinstallazione?"
MbText = MbText & VbCrlf & VbCrlf & "NOTA BENE: OhMyCon viene fornito sotto licenza MIT e senza alcuna garanzia. Nel caso "
MbText = MbText & "di danni al dispositivo in uso lo sviluppatore non potr" & chr(224) & " essere ritenuto responsabile!"

MbResult = MsgBox(MbText, VbYesNoCancel + VbQuestion, MbTitle)

If MbResult = VbYes Then
    Verbs = "OhMyCon.Predefinito;" & _
        "OhMyCon.Giallo;" & _
        "OhMyCon.Arancione;" & _
        "OhMyCon.Rosso;" & _
        "OhMyCon.Rosa;" & _
        "OhMyCon.Viola;" & _
        "OhMyCon.Blu;" & _
        "OhMyCon.Azzurro;" & _
        "OhMyCon.Acqua;" & _
        "OhMyCon.Lime;" & _
        "OhMyCon.Verde"

    For Each Item In Split(Verbs, ";")
        Key = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\"
        Key = Key & Item & "\"

        Regedit.RegDelete key & "Command\" & ""
        Regedit.RegDelete Key & "Icon"
        Regedit.RegDelete Key & ""
    Next

    Key = "HKCR\Directory\Shell\OhMyCon\"
    Regedit.RegDelete Key & "MUIVerb"
    Regedit.RegDelete Key & "SubCommands"
    Regedit.RegDelete Key & "Icon"
    Regedit.RegDelete Key & "Position"
    Regedit.RegDelete Key & "SeparatorBefore"
    Regedit.RegDelete Key & "SeparatorAfter"
    Regedit.RegDelete Key & "CommandFlags"

    Files = Array("ohmycon.vbs", _
        "ohmycon.ico", _
        "Predefinito.ico", _
        "Giallo.ico", _
        "Arancione.ico", _
        "Rosso.ico", _
        "Rosa.ico", _
        "Viola.ico", _
        "Blu.ico", _
        "Azzurro.ico", _
        "Acqua.ico", _
        "Lime.ico", _
        "Verde.ico" _
    )

    For Each File In Files
        Fso.DeleteFile File, "C:\OhMyCon\" & File
    Next

    Fso.DeleteFolder "C:\OhMyCon"

    MbText = "Disinstallazione completata!"

    MbResult = MsgBox(MbText, VbOkOnly + VbInformation, MbTitle)
Else
    MbText = "Disinstallazione annullata!"

    MbResult = MsgBox(MbText, VbOkOnly + VbInformation, MbTitle)
End If
