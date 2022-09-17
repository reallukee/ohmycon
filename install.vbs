' OhMyCon (https://github.com/reallukee/ohmycon)
' Creato da Realluke (https://github.com/reallukee/)

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

MbTitle = "OhMyCon Installer"
MbText = "Benvenuto in OhMyCon Installer! Questo programma installer" & chr(224) & " OhMyCon sul dispositivo in uso."
MbText = MbText & "Questo programma funziona solo con la versione 0.0.1 di OhMyCon."
MbText = MbText & VbCrlf & VbCrlf & "Si vuole procedere con la installazione?"
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

        Regedit.RegWrite Key & "", Split(Item, ".")(1), "REG_SZ"
        Regedit.RegWrite Key & "Icon", "C:\OhMyCon\" & Split(Item, ".")(1) & ".ico", "REG_SZ"
        Regedit.RegWrite key & "Command\" & "", "wscript.exe C:\OhMyCon\ohmycon.vbs " & Split(Item, ".")(1) & " %1", "REG_SZ"
    Next

    Key = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\"
    Key = Key & "OhMyCon.Predefinito\"

    Regedit.RegWrite Key & "Icon", "C:\OhMyCon\Predefinito.ico", "REG_SZ"
    Regedit.RegWrite Key & "CommandFlags", "64", "REG_DWORD"

    Key = "HKCR\Directory\Shell\OhMyCon\"
    Regedit.RegWrite Key & "MUIVerb", "OhMyCon", "REG_SZ"
    Regedit.RegWrite Key & "SubCommands", Verbs, "REG_SZ"
    Regedit.RegWrite Key & "Icon", "C:\OhMyCon\ohmycon.ico", "REG_SZ"
    Regedit.RegWrite Key & "Position", "Top", "REG_SZ"
    Regedit.RegWrite Key & "SeparatorBefore", "", "REG_SZ"
    Regedit.RegWrite Key & "SeparatorAfter", "Top", "REG_SZ"
    Regedit.RegWrite Key & "CommandFlags", "64", "REG_DWORD"

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

    Fso.CreateFolder "C:\OhMyCon"

    For Each File In Files
        Fso.CopyFile File, "C:\OhMyCon\" & File
    Next

    MbText = "Installazione completata!"

    MbResult = MsgBox(MbText, VbOkOnly + VbInformation, MbTitle)
Else
    MbText = "Installazione annullata!"

    MbResult = MsgBox(MbText, VbOkOnly + VbInformation, MbTitle)
End If
