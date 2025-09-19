' Popup32.vbs
' Safer and cleaner version of Popup32

Option Explicit

Dim cooldownEnabled, popups, popupCount
Set popups = CreateObject("Scripting.Dictionary")
popupCount = 0

' === Ask if cooldown is enabled ===
If MsgBox("Enable 10 second cooldown before showing popups?", vbYesNo + vbQuestion, "Popup32 Setup") = vbYes Then
    cooldownEnabled = True
Else
    cooldownEnabled = False
End If

Do
    ' === Get popup message ===
    Dim msg, title, popupType, msgStyle
    msg = InputBox("Enter the popup description (message):", "Popup32 - Message")
    If Trim(msg) = "" Then Exit Do

    title = InputBox("Enter the popup title (default=Popup32):", "Popup32 - Title", "Popup32")
    If Trim(title) = "" Then title = "Popup32"

    popupType = InputBox("Choose popup type:" & vbCrLf & _
                         "1 = Critical" & vbCrLf & _
                         "2 = Exclamation" & vbCrLf & _
                         "3 = Information" & vbCrLf & _
                         "4 = Question" & vbCrLf & _
                         "5 = Plain (no icon)", _
                         "Popup32 - Type", "3")

    Select Case popupType
        Case "1": msgStyle = 16   ' Critical stop
        Case "2": msgStyle = 48   ' Exclamation
        Case "3": msgStyle = 64   ' Information
        Case "4": msgStyle = 32   ' Question
        Case "5": msgStyle = 0    ' No icon
        Case Else: msgStyle = 0
    End Select

    ' Store popup
    popupCount = popupCount + 1
    popups.Add CStr(popupCount), title & "|" & msg & "|" & CStr(msgStyle)

    ' Ask next action
    Dim action
    action = MsgBox("Choose an action:" & vbCrLf & _
                    "Yes = Add another popup" & vbCrLf & _
                    "No = Show all popups" & vbCrLf & _
                    "Cancel = Exit", _
                    vbYesNoCancel + vbQuestion, "Popup32")

    If action = vbYes Then
        ' Loop to add another popup
    ElseIf action = vbNo Then
        Exit Do
    Else
        WScript.Quit
    End If
Loop

' === Show all popups ===
Dim key, parts
For Each key In popups.Keys
    parts = Split(popups(key), "|")
    title = parts(0)
    msg = parts(1)
    msgStyle = CInt(parts(2))

    If cooldownEnabled Then WScript.Sleep 10000 ' 10 seconds

    MsgBox msg, msgStyle, title
Next
