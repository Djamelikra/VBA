Attribute VB_Name = "mMsgboxPopup"
Option Explicit
Sub msgBoxTempo(secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup("Click OK (this window closes automatically after)." & " " & secondWait & " " & "seconds.", _
    secondWait, "This is your Message Box", vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
End Sub
Sub MessageBoxTimer()
    Dim AckTime As Integer, InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
    AckTime = 1
    Select Case InfoBox.Popup("Click OK (this window closes automatically after 10 seconds).", _
    AckTime, "This is your Message Box", vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
End Sub
Sub msgBoxTempMsgWait(sMessage As String, secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup(sMessage, secondWait, gAppName, vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
    
    Beep
    
End Sub
Sub msgBoxTempFullCustom(sMessage As String, secondWait As Byte, msgTitle As String, iButton As Integer)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
    'iButton vbCritical = 16; vbQuestion = 32; vbExclamation = 48; vbInformation = 64
    
        InfoBox.Popup sMessage, secondWait, msgTitle, iButton
    Beep
End Sub

Sub TestmsgBoxTempFullCustom()
    msgBoxTempFullCustom "ttt", 3, gAppName, 64
    
    
    
    
End Sub


Sub msgBoxTempFullCustomRun()
    Dim sMessage As String
    Dim iButton As Long
    Dim secondWait As Byte
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
    'iButton vbCritical = 16; vbQuestion = 32; vbExclamation = 48; vbInformation = 64
    
    
    sMessage = "test" '
    iButton = 64
    secondWait = 3
    
    
        InfoBox.Popup sMessage, secondWait, gAppName, iButton
    Beep
End Sub









