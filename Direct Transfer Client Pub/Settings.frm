VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   5700
   ClientLeft      =   1050
   ClientTop       =   1350
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Settings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   120
      Picture         =   "Settings.frx":0442
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   120
      Picture         =   "Settings.frx":087B
      ScaleHeight     =   720
      ScaleWidth      =   9180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009D6644&
      Caption         =   "Communications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3375
      Begin VB.TextBox txtParent 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboAddress 
         Height          =   330
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cboCommunications 
         Height          =   330
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblParentTag 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.DriveListBox drv2 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&OK"
      ForeColor       =   &H00FFC0C0&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      Top             =   225
      Width           =   1365
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
      ForeColor       =   &H00FFC0C0&
      Height          =   210
      Left            =   2040
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      Top             =   225
      Width           =   1365
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   1320
      Picture         =   "Settings.frx":7F7D
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   2760
      Picture         =   "Settings.frx":999F
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblAbraData 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgOk 
      Height          =   405
      Left            =   240
      Picture         =   "Settings.frx":9D88
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
   Begin VB.Image imgClose 
      Height          =   405
      Left            =   2040
      Picture         =   "Settings.frx":B7AA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_cN As cNeoCaption

Public Sub Skin(f As Form, cN As cNeoCaption)
      cN.ActiveCaptionColor = &H9D6644
      cN.InActiveCaptionColor = &HCCCCCC
      cN.ActiveMenuColor = &HCCCCCC
      cN.ActiveMenuColorOver = &HFFFFFF
      cN.InActiveMenuColor = &H808080
      cN.MenuBackgroundColor = &H0&
      cN.CaptionFont = Me.Font
      cN.MenuFont = Me.Font
      cN.Attach f, picResource(2).Picture, picResource(3).Picture, 13, 14, 90, 142, 240, 372
      f.BackColor = &H9D6644
End Sub

Private Sub imgClose_Click()
    Call lblClose_Click
End Sub

Private Sub imgOk_Click()
    Call lblOK_Click
End Sub


Private Sub lblClose_Click()
    Unload frmSettings
End Sub

Private Sub lblOK_Click()
Dim Holder As String
    Holder = drv2.Drive & "\"
    SaveSetting "DIRECTTRANSFER", "STARTUP", "Location", _
         Holder
    WriteSettings
    If MsgBox("You must restart the program for all of these changes to take effect. Would you like to restart now?", vbYesNo + vbQuestion) = vbYes Then
        Unload frmMain
        Call frmMain.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set m_cN = New cNeoCaption
    Skin Me, m_cN
    Me.Width = 3900
    Me.Height = 4800
    txtParent.Text = gParentCompany & gPassword
    cboCommunications.AddItem ("DirectTransfer")
    cboCommunications.AddItem ("Email")
    If gCommunication = "DirectTransfer" Then
        cboCommunications.ListIndex = 0
    Else
        If gCommunication = "Email" Then
            cboCommunications.ListIndex = 1
        Else
            cboCommunications.ListIndex = -1
        End If
    End If
    cboAddress.AddItem ("Main")
    cboAddress.AddItem ("Backup 1")
    cboAddress.AddItem ("Backup 2")
    If gEmailAddress = "ABC@xyz.com" Then
        cboAddress.ListIndex = 0
    Else
        If gEmailAddress = "ABC1@xyz.com" Then
            cboAddress.ListIndex = 1
        Else
            If gEmailAddress = "Backup@xyz.net" Then
                cboAddress.ListIndex = 2
            End If
        End If
    End If
    drv2.Drive = Left(gLocation, 1)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
            Case vbKeyO
                Call lblOK_Click
            Case vbKeyC
                Call lblClose_Click
        End Select
    End If
End Sub

Private Sub WriteSettings()
Dim Line As String
    Open gLocation & "IDI\Desktop\Settings.idi" For Output As #1
    Line = "01" & Left(txtParent.Text, 3)
    Print #1, Line
    Line = "04"
    If txtParent.Text <> "" Then Line = Line & Right(txtParent.Text, Len(txtParent.Text) - 3)
    Print #1, Line
    Line = "05" & cboCommunications.Text
    Print #1, Line
    Line = "06"
    Select Case cboAddress.List(cboAddress.ListIndex)
        Case "Main"
            Line = Line & "Input@informationdynamics.com"
        Case "Backup 1"
            Line = Line & "Backup1@informationdynamics.com"
        Case "Backup 2"
            Line = Line & "tdougherty@ameritech.net"
    End Select
    Print #1, Line
    Close #1
End Sub

Private Sub lblOk_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgOk.Picture = imgButton(1).Picture
        lblOk.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblOk_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgOk.Picture = imgButton(0).Picture
    lblOk.ForeColor = gButtonForecolor
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgClose.Picture = imgButton(1).Picture
        lblClose.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgClose.Picture = imgButton(0).Picture
    lblClose.ForeColor = gButtonForecolor
End Sub



