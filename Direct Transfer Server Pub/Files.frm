VERSION 5.00
Begin VB.Form frmFiles 
   BackColor       =   &H00FF8080&
   Caption         =   "Logs"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "Files.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCompanies 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   9735
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtNotificationEmail 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8040
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   120
      Picture         =   "Files.frx":0442
      ScaleHeight     =   720
      ScaleWidth      =   9180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   540
      Picture         =   "Files.frx":7B44
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5700
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Notification Email Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Email Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   2040
      Picture         =   "Files.frx":7F7D
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   600
      Picture         =   "Files.frx":8366
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Delete"
      ForeColor       =   &H00FFC0C0&
      Height          =   405
      Left            =   4200
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Add"
      ForeColor       =   &H00FFC0C0&
      Height          =   405
      Left            =   1800
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Save"
      ForeColor       =   &H00FFC0C0&
      Height          =   405
      Left            =   6600
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      Top             =   240
      Width           =   1365
   End
   Begin VB.Image imgNew 
      Height          =   405
      Left            =   1800
      Picture         =   "Files.frx":9D88
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
   Begin VB.Image imgDelete 
      Height          =   405
      Left            =   4200
      Picture         =   "Files.frx":B7AA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
   Begin VB.Image imgSave 
      Height          =   405
      Left            =   6600
      Picture         =   "Files.frx":D1CC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
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

Private Sub imgDelete_Click()
    Call lblDelete_Click
End Sub

Private Sub imgNew_Click()
    Call lblNew_Click
End Sub

Private Sub imgSave_Click()
    Call lblSave_Click
End Sub

Private Sub lblNew_Click()
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = gLightForecolor
    If Len(txtCode.Text) = 3 And (txtEmail.Text <> "" Or txtNotificationEmail.Text <> "" Or txtPassword.Text <> "") Then
        lstCompanies.AddItem (txtCode.Text & "," & txtEmail.Text & "," & txtNotificationEmail.Text & "," & txtPassword.Text)
        txtCode.Text = ""
        txtEmail.Text = ""
        txtNotificationEmail.Text = ""
        txtPassword = ""
    Else
        MsgBox "You must enter both Company Code and Email Address or a Password to add."
    End If
    DoEvents
    imgNew.Picture = imgButton(0).Picture
    lblNew.ForeColor = gButtonForecolor
End Sub

Private Sub lblDelete_Click()
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = gLightForecolor
    If lstCompanies.ListIndex <> -1 Then
        lstCompanies.RemoveItem (lstCompanies.ListIndex)
    Else
        MsgBox "Please select a line to delete."
    End If
    DoEvents
    imgDelete.Picture = imgButton(0).Picture
    lblDelete.ForeColor = gButtonForecolor
End Sub

Private Sub lblSave_Click()
    WriteCompanies
End Sub

Private Sub Form_Load()
    Set m_cN = New cNeoCaption
    Skin Me, m_cN
    Me.Width = 10310
    Me.Height = 4000
    ReadCompanies
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
            Case vbKeyA
                Call lblNew_Click
            Case vbKeyD
                Call lblDelete_Click
            Case vbKeyS
                Call lblSave_Click
        End Select
    End If
End Sub

Private Sub ReadCompanies()
Dim f, ts
    Set f = fso.GetFile("C:\DirectTransfer\Companies.idi")
    Set ts = f.OpenAsTextStream(1, -2)
    Do Until ts.AtEndOfStream = True
        Holder = ts.ReadLine
        lstCompanies.AddItem (Holder)
    Loop
    ts.Close
End Sub

Private Sub WriteCompanies()
Dim txtfile
    Set txtfile = fso.CreateTextFile("C:\DirectTransfer\Companies.idi")
    If lstCompanies.ListCount <> 0 Then
        For i = 0 To lstCompanies.ListCount - 1
            txtfile.WriteLine (lstCompanies.List(i))
        Next i
    End If
    txtfile.Close
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = gLightForecolor
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSave.Picture = imgButton(0).Picture
    lblSave.ForeColor = gButtonForecolor
End Sub
