VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct Transfer"
   ClientHeight    =   7170
   ClientLeft      =   1395
   ClientTop       =   1305
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox drvAvail 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   3195
      Width           =   3015
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   0
      Picture         =   "Main.frx":08CA
      ScaleHeight     =   720
      ScaleWidth      =   9180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6120
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   420
      Picture         =   "Main.frx":7FCC
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6900
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5640
      Width           =   9915
   End
   Begin VB.TextBox txtOutBox 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   675
      Width           =   9495
   End
   Begin VB.DirListBox dirAvail 
      Height          =   2010
      Left            =   1920
      TabIndex        =   3
      Top             =   3480
      Width           =   3015
   End
   Begin VB.FileListBox filAvail 
      Height          =   1140
      Left            =   6600
      TabIndex        =   1
      Top             =   3195
      Width           =   3015
   End
   Begin VB.FileListBox filSend 
      Height          =   1140
      Left            =   6600
      TabIndex        =   0
      Top             =   4395
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   10680
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   10560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   9840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   9840
      Picture         =   "Main.frx":8405
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblSettings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   7560
      TabIndex        =   18
      Tag             =   "ButtonLabel"
      Top             =   210
      Width           =   1455
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Tag             =   "ButtonLabel"
      Top             =   4875
      Width           =   1215
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Tag             =   "ButtonLabel"
      Top             =   3675
      Width           =   1215
   End
   Begin VB.Label lblSendReceive 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send && Receive"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   720
      TabIndex        =   13
      Tag             =   "ButtonLabel"
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Directories"
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
      TabIndex        =   12
      Top             =   3555
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Drives"
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
      TabIndex        =   11
      Top             =   3195
      Width           =   1575
   End
   Begin VB.Label lblMessageOut 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
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
      Left            =   -1080
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblSend 
      BackStyle       =   0  'Transparent
      Caption         =   "Files To Be Sent"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   4395
      Width           =   1335
   End
   Begin VB.Label lblAvailF 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Files"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Label lblUnencrypt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unencrypt"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   5280
      TabIndex        =   7
      Tag             =   "ButtonLabel"
      Top             =   195
      Width           =   1455
   End
   Begin VB.Label lblEncrypt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   3000
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   180
      Width           =   1455
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   9840
      Picture         =   "Main.frx":87EE
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgDelete 
      Height          =   405
      Left            =   5160
      Picture         =   "Main.frx":A210
      Stretch         =   -1  'True
      Top             =   4770
      Width           =   1215
   End
   Begin VB.Image imgAdd 
      Height          =   405
      Left            =   5160
      Picture         =   "Main.frx":BC32
      Stretch         =   -1  'True
      Top             =   3570
      Width           =   1215
   End
   Begin VB.Image imgEncrypt 
      Height          =   315
      Left            =   3000
      Picture         =   "Main.frx":D654
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1455
   End
   Begin VB.Image imgUnencrypt 
      Height          =   315
      Left            =   5280
      Picture         =   "Main.frx":F076
      Stretch         =   -1  'True
      Top             =   150
      Width           =   1455
   End
   Begin VB.Image imgSendReceive 
      Height          =   315
      Left            =   720
      Picture         =   "Main.frx":10A98
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1455
   End
   Begin VB.Image imgSettings 
      Height          =   315
      Left            =   7560
      Picture         =   "Main.frx":124BA
      Stretch         =   -1  'True
      Top             =   165
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BytesRecieved As Double
Public Size As Double
Public Sent As Double
Public FileNum As Integer
Public Passes As Integer
Private fEncrypt As Boolean
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

Private Sub Form_Unload(Cancel As Integer)
    KillFiles
End Sub

Private Sub imgAdd_Click()
    Call lblAdd_Click
End Sub

Private Sub imgDelete_Click()
    Call lblDelete_Click
End Sub

Private Sub imgEncrypt_Click()
    Call lblEncrypt_Click
End Sub

Private Sub imgSendReceive_Click()
    Call lblSendReceive_Click
End Sub

Private Sub imgSettings_Click()
    Call lblSettings_Click
End Sub

Private Sub imgUnencrypt_Click()
    Call lblUnencrypt_Click
End Sub

Private Sub lblAdd_Click()
    If filAvail.filename <> "" Then
        FileCopy filAvail.Path & "\" & filAvail.filename, filSend.Path & "\" & filAvail.filename
        filSend.Refresh
    End If
End Sub

Private Sub lblDelete_Click()
    If filSend.filename <> "" Then
        Kill filSend.Path & "\" & filSend.filename
        filSend.Refresh
    End If
End Sub

Private Sub dirAvail_Change()
    filAvail.Path = dirAvail.Path
End Sub

Private Sub drvAvail_Change()
    dirAvail.Path = drvAvail.Drive
    filAvail.Path = drvAvail.Drive
End Sub

Private Sub Form_Load()
    gButtonForecolor = &HFFC0C0
    gLightForecolor = &H80000012
    Set m_cN = New cNeoCaption
    Skin Me, m_cN
    Me.Width = 10000
    Me.Height = 6765
    gLocation = GetSetting("DIRECTTRANSFER", "STARTUP", _
         "Location", "")
    If gLocation = "" Then
        gLocation = InputBox("Please enter the  network drive of the IDI Folder. Example: Z:\", "Confirm Location")
        If Len(gLocation) = 1 Then gLocation = gLocation & ":\"
        If Len(gLocation) = 2 Then gLocation = gLocation & "\"
        If Len(gLocation) = 0 Then gLocation = gLocation & "C:\"
        If Len(gLocation) > 3 Then gLocation = Left(gLocation, 3)
        SaveSetting "DIRECTTRANSFER", "STARTUP", "Location", _
            gLocation
    End If
    ReadSettings
    If gCommunication <> "Email" Then
        Passes = 0
        Size = 0
        hostPort = 4444  'Can be any port you want as long it is the same on both sides
        hostIP = "127.0.0.1" 'Your server's ip address goes here
        sckSend.RemoteHost = hostIP
        sckSend.RemotePort = hostPort
    End If
    drvAvail.Drive = Left(gLocation, 1)
    dirAvail.Path = gLocation
    filAvail.Path = gLocation
    filSend.Path = gLocation & "IDI\Desktop\SEND\TEMP"
End Sub

Private Sub lblEncrypt_Click()
    If filSend.ListCount <> 0 Then
        ZipMail
        EncryptZip
        MsgBox "Encryption Complete"
    Else
        MsgBox "No Files To Encrypt"
    End If
End Sub

Private Sub lblSettings_Click()
    Call frmSettings.Show
End Sub

Private Sub lblUnencrypt_Click()
    If Dir(gLocation & "IDI\Desktop\MESSAGES\TEMP\" & gParentCompany & "_zip.sen") <> "" Then
        UnencryptZip
        UnZipMail
        MsgBox "Unencryption Complete"
    Else
        MsgBox "No Files To Unencrypt"
    End If
End Sub

Private Sub sckSend_Close()
    txtStatus.Text = " Transfer Complete"
End Sub

Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
Dim Holder As String
Dim sItemData As String
    BytesRecieved = BytesRecieved + bytesTotal
    sckSend.GetData sItemData
    Passes = Passes + 1
    If Passes = 1 Then
        n = Left(sItemData, 3)
        One = InStr(1, sItemData, "@@@") + 2
        Two = InStr(1, sItemData, "###") + 2
        Three = InStr(1, sItemData, "$$$") - 1
        Holder = Right(Left(sItemData, Three), Three - Two)
        Size = CLng(Holder) + Three + 3
        sServerMsg = Format(BytesRecieved) & " of " & Format(Size) & " received. " & Format(Int(BytesRecieved / Size * 100)) & "%"
        txtStatus.Text = " " & sServerMsg
        Holder = gLocation & "IDI\Desktop\Messages\Temp\" & n & "_zip.sen"
        FileNum = FreeFile
        Open Holder For Binary Access Write As FileNum
        sItemData = Right(sItemData, Len(sItemData) - Three - 3)
        Put FileNum, , sItemData
    Else
        Put FileNum, , sItemData
        sServerMsg = Format(BytesRecieved) & " of " & Format(Size) & " received. " & Format(Int(BytesRecieved / Size * 100)) & "%"
        txtStatus.Text = " " & sServerMsg
    End If
    If BytesRecieved >= Size Then
        Close FileNum
        DoEvents
        sckSend.Close
        Passes = 0
        BytesRecieved = 0
        txtStatus.Text = " Transfer of " & gParentCompany & ".zip completed successfully."
        UnencryptZip
        UnZipMail
        txtStatus.Text = " Transfer Complete"
        MsgBox "Files Received. Please Load."
    End If
End Sub

Private Sub sckSend_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Sent = Sent + bytesSent
    If Size <> 0 Then txtStatus.Text = " " & Format(Sent) & " of " & Format(Size) & " sent. " & Format(Int(Sent / Size * 100)) & "%"
End Sub

Private Sub lblSendReceive_Click()
    ZipMail
    If fEncrypt = True Then EncryptZip
    If gCommunication = "Email" Then
        SendMail
        ReceiveMail
    Else
        txtStatus.Text = " Waiting for host to accept transfer"
        sckSend.Close
        sckSend.Connect
        Do Until sckSend.State = sckConnected ' Wait until connected
            DoEvents
        Loop
        txtStatus.Text = " Connected"
        SendFile gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen"
    End If
End Sub

Private Sub ReceiveMail()
Dim intLenFileName As Integer
Dim intStrPos As Integer
Dim i As Integer
Dim Msg As Boolean
Dim t As Integer
    Msg = False
    mapSess.NewSession = True
    mapSess.DownLoadMail = True
    mapSess.UserName = "IDI"
    mapSess.LogonUI = False
    mapSess.SignOn
    mapMess.SessionID = mapSess.SessionID
    mapMess.FetchUnreadOnly = True
    mapMess.Fetch
    If mapMess.MsgCount > 0 Then
        For t = mapMess.MsgCount - 1 To 0 Step -1
            mapMess.MsgIndex = t
            If mapMess.AttachmentCount > 0 Then
                Msg = True
                For i = 0 To mapMess.AttachmentCount - 1
                    mapMess.AttachmentIndex = i
                    intLenFileName = Len(mapMess.AttachmentPathName)
                    For intStrPos = intLenFileName To 1 Step -1
                        If InStr(1, Right$(mapMess.AttachmentPathName, intLenFileName - (intStrPos - 1)), "\", 1) Then
                            strNewFileName = Right$(mapMess.AttachmentPathName, intLenFileName - intStrPos)
                            Exit For
                        End If
                    Next intStrPos
                    FileCopy mapMess.AttachmentPathName, gLocation & "IDI\Desktop\MESSAGES\TEMP\" & Left(strNewFileName, 3) & "_zip.sen"
                    DoEvents
                    UnencryptZip
                    UnZipMail
                Next i
                Holder = mapMess.MsgNoteText
                mapMess.Delete
                DoEvents
            End If
        Next t
    End If
    If Msg = True Then
        MsgBox "Files Received. Please Load."
    Else
        MsgBox "No new files at this time."
    End If
    mapSess.SignOff
End Sub

Private Sub SendMail()
Dim EmailAddres As String
    If txtOutBox.Text <> "" Or Dir(gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen") <> "" Then
        mapSess.NewSession = True
        mapSess.UserName = "IDI"
        mapSess.DownLoadMail = False
        mapSess.LogonUI = False
        mapSess.SignOn
        mapMess.SessionID = mapSess.SessionID
        mapMess.Compose
        mapMess.RecipAddress = gEmailAddress
        mapMess.AddressResolveUI = True
        mapMess.ResolveName
        mapMess.MsgSubject = gParentCompany
        If txtOutBox.Text <> "" Then
            mapMess.MsgNoteText = txtOutBox.Text
        Else
            mapMess.MsgNoteText = "Sending Files"
        End If
        If Dir(gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen") <> "" Then
            mapMess.AttachmentPathName = gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen"
        End If
        mapMess.Send
        mapSess.SignOff
        MsgBox "Send complete."
        If Dir(gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen") <> "" Then Kill gLocation & "IDI\Desktop\SEND\" & gParentCompany & "_zip.sen"
    End If
End Sub

Private Sub ZipMail()
Dim returncode As Integer
    fEncrypt = False
    If filSend.ListCount > 0 Then
        zDate = vbNullString
        zJunkDir = 1     ' 1 = Throw Away Path Names
        zRecurse = 0     ' 1 = Recurse -R 2 = Recurse -r 2 = Most Useful :)
        zUpdate = 0      ' 1 = Update Only If Newer
        zFreshen = 0     ' 1 = Freshen - Overwrite Only
        zLevel = Asc(9)  ' Compression Level (0 - 9)
        zEncrypt = 0     ' Encryption = 1 For Password Else 0
        zNoDirEntries = 1
        '-- Select Some Files - Wildcards Are Supported
        zArgc = filSend.ListCount + 1        ' Number Of Elements Of mynames Array
        zZipFileName = gLocation & "IDI\Desktop\SEND\" & gParentCompany & ".zip"
        zZipFileNames.zFiles(0) = ""
        For i = 0 To filSend.ListCount - 1
            zZipFileNames.zFiles(i + 1) = gLocation & "IDI\Desktop\SEND\TEMP\" & filSend.List(i)
        Next i
        zRootDir = ""    ' This Affects The Stored Path Name
        returncode = VBZip32
        DoEvents
        For i = 0 To filSend.ListCount - 1
            Kill filSend.Path & "\" & filSend.List(i)
        Next i
        filSend.Refresh
        fEncrypt = True
    End If
End Sub

Private Sub UnZipMail()
    uZipInfo = ""
    uZipNumber = 0   ' Holds The Number Of Zip Files
    uPromptOverWrite = 0  ' 1 = Prompt To Overwrite
    uOverWriteFiles = 1   ' 1 = Always Overwrite Files
    uExtractList = 0       ' 1 = List Contents Of Zip 0 = Extract
    uHonorDirectories = 0  ' 1 = Honour Zip Directories
    uZipNames.uzFiles(0) = vbNullString
    uNumberFiles = 0
    uExcludeNames.uzFiles(0) = vbNullString
    uNumberXFiles = 0
    uZipFileName = gLocation & "IDI\Desktop\MESSAGES\TEMP\" & gParentCompany & ".zip"
    uExtractDir = gLocation & "IDI\RCV1"
    uVbSkip = 0
    Call VBUnZip32
    If Dir(gLocation & "IDI\Desktop\MESSAGES\TEMP\" & gParentCompany & ".zip") <> "" Then
        Kill gLocation & "IDI\Desktop\MESSAGES\TEMP\" & gParentCompany & ".zip"
    End If
    If Dir(gLocation & "IDI\RCV1\IDIDesk.exe") <> "" Then
        FileCopy gLocation & "IDI\RCV1\IDIDesk.exe", gLocation & "IDI\Desktop\MESSAGES\IDIDesk.exe"
        Kill gLocation & "IDI\RCV1\IDIDesk.exe"
    End If
End Sub

Public Sub SendFile(filename As String)
Dim hIn, fileLength, ret
Dim temp As String
Dim blockSize As Long
    blockSize = 2048
    On Error GoTo ErrorHandler
    If txtOutBox.Text <> "" Then
        If Dir(filename) <> "" Then
            hIn = FreeFile
            Open filename For Binary Access Read As hIn
            fileLength = LOF(hIn)
            Size = fileLength
            i = 0
            n = 0
            Do Until EOF(hIn)
                DoEvents
                n = n + 1
                If fileLength - Loc(hIn) <= blockSize Then
                    blockSize = fileLength - Loc(hIn)
                End If
                temp = Space$(blockSize)
                Get hIn, , temp
                DoEvents
                If n = 1 Then
                    temp = gParentCompany & "_zip.sen" & "@@@" & txtOutBox.Text & "###" & Format(fileLength) & "$$$" & temp
                End If
                If sckSend.State = 7 Then sckSend.SendData temp
                sizeOfFileSent = sizeOfFileSent + blockSize
                On Error GoTo Endit
                If sizeOfFileSent = fileLength Then i = i + 1
                If i = 100 Or sckSend.State <> 7 Then
                    Exit Do
                End If
            Loop
            Close hIn
            Kill filename
            txtOutBox.Text = ""
        Else
            temp = gParentCompany & "_zip.sen" & "@@@" & txtOutBox.Text & "###" & "0" & "$$$"
            If sckSend.State = 7 Then sckSend.SendData temp
        End If
    Else
        If Dir(filename) <> "" Then
            hIn = FreeFile
            Open filename For Binary Access Read As hIn
            fileLength = LOF(hIn)
            Size = fileLength
            i = 0
            n = 0
            Do Until EOF(hIn)
                DoEvents
                n = n + 1
                If fileLength - Loc(hIn) <= blockSize Then
                    blockSize = fileLength - Loc(hIn)
                End If
                temp = Space$(blockSize)
                Get hIn, , temp
                DoEvents
                If n = 1 Then
                    temp = gParentCompany & "_zip.sen" & "@@@" & "Sending Files" & "###" & Format(fileLength) & "$$$" & temp
                End If
                If sckSend.State = 7 Then sckSend.SendData temp
                sizeOfFileSent = sizeOfFileSent + blockSize
                On Error GoTo Endit
                If sizeOfFileSent = fileLength Then i = i + 1
                If i = 100 Or sckSend.State <> 7 Then
                    Exit Do
                End If
            Loop
            Close hIn
            Kill filename
        Else
            temp = gParentCompany & "_zip.sen" & "@@@" & "Just Checking" & "###" & "0" & "$$$"
            If sckSend.State = 7 Then sckSend.SendData temp
        End If
    End If
    Exit Sub
ErrorHandler:
    Close hIn
Endit:
End Sub

Private Sub lblAdd_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgAdd.Picture = imgButton(1).Picture
        lblAdd.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblAdd_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgAdd.Picture = imgButton(0).Picture
    lblAdd.ForeColor = gButtonForecolor
End Sub

Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgDelete.Picture = imgButton(1).Picture
        lblDelete.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgDelete.Picture = imgButton(0).Picture
    lblDelete.ForeColor = gButtonForecolor
End Sub

Private Sub lblSendReceive_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgSendReceive.Picture = imgButton(0).Picture
    lblSendReceive.ForeColor = gButtonForecolor
End Sub

Private Sub lblSendReceive_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgSendReceive.Picture = imgButton(1).Picture
        lblSendReceive.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblEncrypt_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgEncrypt.Picture = imgButton(1).Picture
        lblEncrypt.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblEncrypt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgEncrypt.Picture = imgButton(0).Picture
    lblEncrypt.ForeColor = gButtonForecolor
End Sub

Private Sub lblUnencrypt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgUnencrypt.Picture = imgButton(0).Picture
    lblUnencrypt.ForeColor = gButtonForecolor
End Sub

Private Sub lblUnencrypt_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgUnencrypt.Picture = imgButton(1).Picture
        lblUnencrypt.ForeColor = gLightForecolor
    End If
End Sub

Private Sub lblSettings_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgSettings.Picture = imgButton(0).Picture
    lblSettings.ForeColor = gButtonForecolor
End Sub

Private Sub lblSettings_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        imgSettings.Picture = imgButton(1).Picture
        lblSettings.ForeColor = gLightForecolor
    End If
End Sub
