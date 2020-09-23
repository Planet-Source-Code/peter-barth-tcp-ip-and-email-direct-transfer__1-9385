Attribute VB_Name = "basForeGround"
'Declare our API functions

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'The hwnd argument specifies the window handle(hWnd) of the target window.

Public Sub SetForeground(hwnd As Long)
    SetForegroundWindow hwnd
End Sub
