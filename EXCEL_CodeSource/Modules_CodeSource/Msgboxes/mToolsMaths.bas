Attribute VB_Name = "mToolsMaths"
Option Explicit

Function msgInfoTemp()
     CreateObject("WScript.Shell").Popup "The data copy has been performed", 1, gAppname, vbInformation
End Function
Public Function MsgBoxTemp(Content As String, Duration As Integer, _
                            Optional Button = vbInformation, Optional Title As String = "TSP Vertices case n")

On Error GoTo MsgBoxTempError
   CreateObject("WScript.Shell").Popup Content, Duration, Title, Button
Exit Function

MsgBoxTempError:
    MsgBoxTemp = CVErr(xlErrValue)
End Function

Sub MessageBoxTimer()
    Dim AckTime As Integer, InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 3 seconds
    AckTime = 3
    Select Case InfoBox.Popup("Click OK (this window closes automatically after 3 seconds).", _
    AckTime, "This is your Message Box", 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub
Sub testMsgBoxTemp()


'    CreateObject("Wscript.shell").Popup "Le Message", 3, "Le Titre", vbExclamation

'    msgTemp

    MsgBoxTemp "The data copy has been performed", 1, vbInformation, gAppname
    

End Sub


'    CreateObject("WScript.Shell").Popup "The data copy has been performed", 1, "Copy", vbInformation

Function arithMean2Oper(scalar1#, scalar2#) As Double
Attribute arithMean2Oper.VB_Description = "Returns arithmetic mean of 2 scalars."
'    Returns arithmetic mean of 2 scalars
    arithMean2Oper = (scalar1 + scalar2) / 2
    
End Function






