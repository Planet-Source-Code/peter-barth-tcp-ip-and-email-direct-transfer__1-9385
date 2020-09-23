Attribute VB_Name = "basglobal"

Public gButtonForecolor As Long
Public gLightForecolor As Long
Public gLocation As String
Public gParentCompany As String
Public gPassword As String
Public gCommunication As String
Public gEmailAddress As String

Public Sub ReadSettings()
Dim Holder As String
Dim i As Integer
    gParentCompany = "XXX"
    gPassword = "0123456789"
    gCommunication = "DirectTransfer"
    gEmailAddress = "Input@informationdynamics.com"
    Open gLocation & "IDI\Desktop\Settings.idi" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Holder
        Select Case Left(Holder, 2)
            Case "01"
                gParentCompany = Right(Holder, Len(Holder) - 2)
            Case "04"
                gPassword = Right(Holder, Len(Holder) - 2)
            Case "05"
                gCommunication = Right(Holder, Len(Holder) - 2)
            Case "06"
                gEmailAddress = Right(Holder, Len(Holder) - 2)
        End Select
    Loop
    Close #1
End Sub

