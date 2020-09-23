VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   Caption         =   "Direct Transfer"
   ClientHeight    =   8445
   ClientLeft      =   2430
   ClientTop       =   630
   ClientWidth     =   9510
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   1920
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   420
      Picture         =   "Main.frx":08CA
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   0
      Picture         =   "Main.frx":0D03
      ScaleHeight     =   720
      ScaleWidth      =   9180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   4440
      Top             =   7680
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   2760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSWinsockLib.Winsock sckReceive 
      Left            =   3360
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   3360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2760
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMessages 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3000
      Width           =   9255
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      Height          =   2790
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEmail 
      Caption         =   "&Email"
      Begin VB.Menu mnuEmailNotify 
         Caption         =   "Notification Address"
         Begin VB.Menu mnuEmailNotifyAddress 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEmailTransfer 
         Caption         =   "Transfer Address"
         Begin VB.Menu mnuEmailTransferAddress 
            Caption         =   "-"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuSetupCompanies 
         Caption         =   "Companies"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fstrFiles() As String
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

Private Sub mnuEmail_Click()
    Dim f, ts
    Dim i As Integer
    mnuEmailTransferAddress(0).Visible = True
    mnuEmailNotifyAddress(0).Visible = True
    For i = 1 To mnuEmailNotifyAddress.UBound
        Unload mnuEmailNotifyAddress(i)
    Next i
    For i = 1 To mnuEmailTransferAddress.UBound
        Unload mnuEmailTransferAddress(i)
    Next i
    Set f = fso.GetFile("C:\DirectTransfer\Companies.idi")
    Set ts = f.OpenAsTextStream(1, -2)
    Do Until ts.AtEndOfStream = True
        Holder = ts.ReadLine
        One = InStr(1, Holder, ",")
        Two = InStr(One + 1, Holder, ",")
        Three = InStr(Two + 1, Holder, ",")
        If Two - One > 1 Then
            Load mnuEmailTransferAddress(mnuEmailTransferAddress.UBound + 1)
            mnuEmailTransferAddress(mnuEmailTransferAddress.UBound).Caption = Left(Holder, 3) & "-" & Mid(Holder, One + 1, Two - One - 1)
            mnuEmailTransferAddress(mnuEmailTransferAddress.UBound).Enabled = True
            mnuEmailTransferAddress(mnuEmailTransferAddress.UBound).Visible = True
        End If
        If Three - Two > 1 Then
            Load mnuEmailNotifyAddress(mnuEmailNotifyAddress.UBound + 1)
            mnuEmailNotifyAddress(mnuEmailNotifyAddress.UBound).Caption = Left(Holder, 3) & "-" & Mid(Holder, Two + 1, Three - Two - 1)
            mnuEmailNotifyAddress(mnuEmailNotifyAddress.UBound).Enabled = True
            mnuEmailNotifyAddress(mnuEmailNotifyAddress.UBound).Visible = True
        End If
    Loop
    ts.Close
    If mnuEmailTransferAddress.Count > 0 Then mnuEmailTransferAddress(0).Visible = False
    If mnuEmailNotifyAddress.Count > 0 Then mnuEmailNotifyAddress(0).Visible = False
End Sub

Private Sub mnuEmailNotifyAddress_Click(Index As Integer)
On Error Resume Next
    mapSess.NewSession = True
    'mapSess.UserName = "ProfileName" 'Uncomment for automatic profile use
    mapSess.DownLoadMail = False
    'mapSess.LogonUI = False    'Uncomment for automatic profile use
    mapSess.SignOn
    mapMess.SessionID = mapSess.SessionID
    mapMess.Compose
    mapMess.RecipAddress = Right(mnuEmailNotifyAddress(Index).Caption, Len(mnuEmailNotifyAddress(Index).Caption) - 4)
    mapMess.AddressResolveUI = True
    mapMess.ResolveName
    mapMess.Send True
End Sub

Private Sub mnuEmailTransferAddress_Click(Index As Integer)
On Error Resume Next
    mapSess.NewSession = True
    'mapSess.UserName = "ProfileName" 'Uncomment for automatic profile use
    mapSess.DownLoadMail = False
    'mapSess.LogonUI = False    'Uncomment for automatic profile use
    mapSess.SignOn
    mapMess.SessionID = mapSess.SessionID
    mapMess.Compose
    mapMess.RecipAddress = Right(mnuEmailTransferAddress(Index).Caption, Len(mnuEmailTransferAddress(Index).Caption) - 4)
    mapMess.AddressResolveUI = True
    mapMess.ResolveName
    mapMess.Send True
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    cdg1.filename = "c:\directtransfer\*.txt"
    cdg1.Filter = "*.txt"
    cdg1.CancelError = True
    On Error Resume Next
    cdg1.ShowOpen
    If cdg1.filename <> "c:\directtransfer\*.txt" Then Execute_Program cdg1.filename, " ", " "
End Sub

Private Sub mnuSetupCompanies_Click()
    Call frmFiles.Show
End Sub

Private Sub Form_Load()
    gButtonForecolor = &HFFC0C0
    gLightForecolor = &H80000012
    Set m_cN = New cNeoCaption
    Skin Me, m_cN
    File1.Path = "c:\RCV1\"
    Socket(0).Close
    Socket(0).LocalPort = 4444
    Socket(0).Listen
End Sub

Private Sub Form_Resize()
    If Me.Width > 550 Then txtMessages.Width = Me.Width - 550
    If Me.Height > 3000 Then txtMessages.Height = Me.Height - 4000
    If Me.Width > 4875 Then List1.Width = Me.Width - 4875
End Sub

Private Sub Socket_Close(Index As Integer)
    Socket(Index).Close
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    sServerMsg = "Connection request id " & requestID & " from " & Socket(Index).RemoteHostIP
    Call List1.AddItem(sServerMsg, List1.ListCount)
    If Index = 0 Then
        For i = 1 To (Socket.Count - 1)
            If Socket(i).Tag = "0" Then GoTo FoundSocket
        Next i
        Dim NewSocket As Integer
        NewSocket = Socket.Count
        Load Socket(NewSocket)
        Socket(NewSocket).LocalPort = 4444
        Socket(NewSocket).Accept requestID
        File(NewSocket).Passes = 0
        File(NewSocket).Message = ""
        File(NewSocket).Name = ""
        File(NewSocket).BytesRecieved = 0
        File(NewSocket).FileNum = 0
        File(NewSocket).Size = 0
        Exit Sub
FoundSocket:
        Socket(i).Accept requestID
        File(i).Passes = 0
        File(i).Message = ""
        File(i).Name = ""
        File(i).BytesRecieved = 0
        File(i).FileNum = 0
        File(i).Size = 0
    End If
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Holder As String
Dim sItemData As String
Dim strName As String
    File(Index).BytesRecieved = File(Index).BytesRecieved + bytesTotal
    Socket(Index).GetData sItemData
    File(Index).Passes = File(Index).Passes + 1
    If File(Index).Passes = 1 Then
        One = InStr(1, sItemData, "@@@") + 2
        Two = InStr(1, sItemData, "###") + 2
        Three = InStr(1, sItemData, "$$$") - 1
        File(Index).Name = Left(sItemData, One - 3)
        sServerMsg = "Receiving data from " & Left(File(Index).Name, 3) & " on " & Socket(Index).RemoteHostIP & "(" & sRequestID & ")"
        Call List1.AddItem(sServerMsg, List1.ListCount)
        File(Index).Message = Right(Left(sItemData, Two - 3), Two - 3 - One)
        Holder = Right(Left(sItemData, Three), Three - Two)
        File(Index).Size = CLng(Holder) + Three + 3
        txtMessages.Text = vbCrLf & Left(File(Index).Name, 3) & " " & Now & " " & vbCrLf & File(Index).Message & vbCrLf & txtMessages.Text
        D = date
        Two = InStrRev(D, "/")
        One = InStr(1, D, "/")
        Holder = "c:\DirectTransfer\" & Right("0" & Left(D, One - 1), 2) & Right("0" & Right(Left(D, Two - 1), Two - 1 - One), 2) & Right("0" & Right(D, Len(D) - Two), 2) & ".txt"
        temp = vbCrLf & Left(File(Index).Name, 3) & " " & Now & " " & vbCrLf & File(Index).Message & vbCrLf
        File(Index).FileNum = FreeFile
        Open Holder For Append As File(Index).FileNum
        Print #File(Index).FileNum, , temp
        Close File(Index).FileNum
        If File(Index).Size > 0 Then
            If Len(File(Index).Name) > 3 Then
                Holder = "c:\DirectTransfer\InBox\" & File(Index).Name
            Else
                Holder = "c:\DirectTransfer\InBox\" & File(Index).Name & ".zip"
            End If
            sItemData = Right(sItemData, Len(sItemData) - Three - 3)
            File(Index).FileNum = FreeFile
            Open Holder For Binary Access Write As File(Index).FileNum
            Put File(Index).FileNum, , sItemData
        End If
    Else
        Put File(Index).FileNum, , sItemData
    End If
    If File(Index).BytesRecieved >= File(Index).Size Then
        Close File(Index).FileNum
        File(Index).Passes = 0
        Call List1.AddItem("Transfer from " & Left(File(Index).Name, 3) & " complete.", List1.ListCount)
        strName = Left(File(Index).Name, 3)
        If Len(File(Index).Name) > 3 Then
            ZipFiles strName
            EncryptZip strName
            SendFile Index
            UnencryptZip strName
            UnZipFiles strName
        Else
            ZipFiles strName
            SendFile Index
            UnZipFiles strName
        End If
    End If
End Sub

Private Sub SendFile(Index As Integer)
Dim hIn, fileLength, ret
Dim temp As String
Dim blockSize As Long
    blockSize = 2048                                '// Set your read buffer size here
    On Error GoTo ErrorHandler
    If Len(File(Index).Name) > 3 Then
        filename = "C:\DirectTransfer\OutBox\" & File(Index).Name
    Else
        filename = "C:\DirectTransfer\OutBox\" & File(Index).Name & ".zip"
    End If
    If fso.FileExists(filename) Then
        hIn = FreeFile
        Open filename For Binary Access Read As hIn
        fileLength = LOF(hIn)
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
                temp = File(Index).Name & "@@@###" & Format(fileLength) & "$$$" & temp
            End If
            If Socket(Index).State = 7 Then Socket(Index).SendData temp
            sizeOfFileSent = sizeOfFileSent + blockSize
            On Error GoTo endIt
            If sizeOfFileSent = fileLength Then i = i + 1
            If i = 100 Or Socket(Index).State <> 7 Then
                Exit Do
            End If
        Loop
        Close hIn
        Call List1.AddItem("Transfer to " & Left(File(Index).Name, 3) & " complete.", List1.ListCount)
        DoEvents
        fso.DeleteFile (filename)
        DoEvents
    Else
        Socket(Index).Close
        Call List1.AddItem("Nothing to transfer to " & Left(File(Index).Name, 3) & ".", List1.ListCount)
    End If
    DoEvents
    Exit Sub

ErrorHandler:                                           '// Always close the file handle
        Close hIn
endIt:
End Sub

Private Sub ZipFiles(Name As String)
Dim returncode As Integer
Dim FileCount As Integer
Dim Path As String
Dim strFiles() As String
    Path = "c:\XMT1\" & Name & "\"
    FileCount = GetFiles(Path)
    strFiles = fstrFiles
    If FileCount > 0 Then
        For i = 1 To FileCount
            FileCopy Path & strFiles(i), "C:\DirectTransfer\OutBox\" & Name & "\" & strFiles(i)
            FileCopy Path & strFiles(i), "c:\SAVE\" & Name & "\" & strFiles(i)
            Set txtfile = fso.GetFile(Path & strFiles(i))
            txtfile.Delete
        Next i
    End If
    Path = "C:\DirectTransfer\OutBox\" & Name & "\"
    FileCount = GetFiles(Path)
    strFiles = fstrFiles
    If FileCount > 0 Then
        zDate = vbNullString
        zJunkDir = 1     ' 1 = Throw Away Path Names
        zRecurse = 0     ' 1 = Recurse -R 2 = Recurse -r 2 = Most Useful :)
        zUpdate = 0      ' 1 = Update Only If Newer
        zFreshen = 0     ' 1 = Freshen - Overwrite Only
        zLevel = Asc(9)  ' Compression Level (0 - 9)
        zEncrypt = 0     ' Encryption = 1 For Password Else 0
        zNoDirEntries = 1
        zArgc = FileCount + 1        ' Number Of Elements Of mynames Array
        zZipFileName = "C:\DirectTransfer\OutBox\" & Name & ".zip"
        zZipFileNames.zFiles(0) = ""
        D = date
        Two = InStrRev(D, "/")
        One = InStr(1, D, "/")
        Holder = "c:\DirectTransfer\" & Right("0" & Left(D, One - 1), 2) & Right("0" & Right(Left(D, Two - 1), Two - 1 - One), 2) & Right("0" & Right(D, Len(D) - Two), 2) & "f.txt"
        FileNum = FreeFile
        Open Holder For Append As FileNum
        For i = 1 To FileCount
            zZipFileNames.zFiles(i) = Path & strFiles(i)
            sItemData = ""
            If i = 1 Then
                sItemData = vbCrLf & vbCrLf & Name & " " & Now & " " & vbCrLf & "Sent"
                Print #FileNum, , sItemData
            End If
            sItemData = strFiles(i) & "  " & FileLen(Path & strFiles(i)) & "  " & FileDateTime(Path & strFiles(i))
            Print #FileNum, , sItemData
        Next i
        Close FileNum
        zRootDir = ""    ' This Affects The Stored Path Name
        returncode = VBZip32
        DoEvents
        For i = 1 To FileCount
            Set txtfile = fso.GetFile(Path & strFiles(i))
            txtfile.Delete
        Next i
    End If
End Sub

Private Sub UnZipFiles(Name As String)
    '-- Init Global Message Variables
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
    uZipFileName = "c:\DirectTransfer\InBox\" & Name & ".zip"
    uExtractDir = "c:\DirectTransfer\InBox\" & Name
    uVbSkip = 0
    Call VBUnZip32
    DoEvents
    If fso.FileExists("c:\DirectTransfer\InBox\" & Name & ".zip") Then
        Set txtfile = fso.GetFile("c:\DirectTransfer\InBox\" & Name & ".zip")
        txtfile.Delete
    End If
    DistributeFiles Name
End Sub

Private Sub DistributeFiles(Name As String)
Dim strFiles() As String
Dim Path As String
Dim FileCount As Integer
Dim i As Integer
    Path = "C:\DirectTransfer\InBox\" & Name & "\"
    FileCount = GetFiles(Path)
    strFiles = fstrFiles
    If FileCount <> 0 Then
        D = date
        Two = InStrRev(D, "/")
        One = InStr(1, D, "/")
        Holder = "c:\DirectTransfer\" & Right("0" & Left(D, One - 1), 2) & Right("0" & Right(Left(D, Two - 1), Two - 1 - One), 2) & Right("0" & Right(D, Len(D) - Two), 2) & "f.txt"
        FileNum = FreeFile
        Open Holder For Append As FileNum
        For i = 1 To FileCount
            sItemData = ""
            If i = 1 Then
                sItemData = vbCrLf & vbCrLf & Name & " " & Now & " " & vbCrLf & "Recieved"
                Print #FileNum, , sItemData
            End If
            sItemData = strFiles(i) & "  " & FileLen(Path & strFiles(i)) & "  " & FileDateTime(Path & strFiles(i))
            Print #FileNum, , sItemData
            If UCase(Right(strFiles(i), 4)) = ".RJE" Then
                FileCopy Path & strFiles(i), "c:\RCV1\" & strFiles(i)
                Set txtfile = fso.GetFile(Path & strFiles(i))
                txtfile.Delete
            End If
        Next i
        Close FileNum
    End If
    Me.WindowState = 0
    SetForeground Me.hWnd
End Sub

Private Sub ReadCompanies()
Dim Holder As String
Dim First As Integer
Dim Second As Integer
Dim Third As Integer
Dim f, ts
    Set f = fso.GetFile("c:\DirectTransfer\Companies.idi")
    Set ts = f.OpenAsTextStream(1, -2)
    ReDim gCompanies(0)
    Do Until ts.AtEndOfStream = True
        Holder = ts.ReadLine
        If Holder <> "" Then
            First = InStr(1, Holder, ",")
            Second = InStr(First + 1, Holder, ",")
            Third = InStrRev(Holder, ",")
            ReDim Preserve gCompanies(0 To (UBound(gCompanies) + 1))
            gCompanies(UBound(gCompanies)).Code = Left(Holder, First - 1)
            gCompanies(UBound(gCompanies)).Email = Right(Left(Holder, Second - 1), Second - First - 1)
            If Third <> Second Then gCompanies(UBound(gCompanies)).NotificationEmail = Right(Left(Holder, Third - 1), Third - Second - 1)
        End If
    Loop
    Call File1.Refresh
    DoEvents
    ScanForFiles
End Sub

Private Sub ScanForFiles()
Dim i As Integer
Dim FileCount As Integer
Dim Path As String
Dim strName As String
Dim Password As String
    mapSess.NewSession = True
    mapSess.UserName = "Input"
    mapSess.DownLoadMail = False
    mapSess.LogonUI = False
    mapSess.SignOn
    mapMess.SessionID = mapSess.SessionID
    For i = 1 To UBound(gCompanies)
        Path = "c:\XMT1\" & gCompanies(i).Code & "\"
        FileCount = GetFiles(Path)
        If FileCount > 0 And gCompanies(i).NotificationEmail <> "" Then SendNotification i
        Path = "c:\DirectTransfer\OutBox\" & gCompanies(i).Code & "\"
        FileCount = FileCount + GetFiles(Path)
        If FileCount > 0 And gCompanies(i).Email <> "" Then
            strName = gCompanies(i).Code
            Password = FindPassword(strName)
            If Password <> "" Then
                ZipFiles strName
                EncryptZip strName
            Else
                ZipFiles strName
            End If
            SendMail i, Password
        End If
        DoEvents
    Next i
    mapSess.SignOff
    DoEvents
    ScanForMail
End Sub

Private Sub SendMail(i As Integer, Extension As String)
Dim EmailAddres As String
    EmailAddress = gCompanies(i).Email
    mapMess.Compose
    mapMess.RecipAddress = EmailAddress
    mapMess.AddressResolveUI = True
    mapMess.ResolveName
    mapMess.MsgSubject = Now
    mapMess.MsgNoteText = "Nothing"
    If Extension <> "" Then
        mapMess.AttachmentPathName = "c:\DirectTransfer\OutBox\" & gCompanies(i).Code & "_zip.sen"
    Else
        mapMess.AttachmentPathName = "c:\DirectTransfer\OutBox\" & gCompanies(i).Code & ".zip"
    End If
    mapMess.Send
    DoEvents
    Call List1.AddItem("Transfer to " & gCompanies(i).Code & " complete.", List1.ListCount)
    If Extension <> "" Then
        fso.DeleteFile ("c:\DirectTransfer\OutBox\" & gCompanies(i).Code & "_zip.sen")
    Else
        fso.DeleteFile ("c:\DirectTransfer\OutBox\" & gCompanies(i).Code & ".zip")
    End If
End Sub

Private Sub SendNotification(i As Integer)
Dim EmailAddres As String
Dim Path As String
    Path = "c:\XMT1\" & gCompanies(i).Code & "\"
    FileCount = GetFiles(Path)
    strFiles = fstrFiles
    If FileCount > 0 Then
        For Index = 1 To FileCount
            FileCopy Path & strFiles(Index), "C:\DirectTransfer\OutBox\" & gCompanies(i).Code & "\" & strFiles(Index)
            FileCopy Path & strFiles(Index), "c:\SAVE\" & gCompanies(i).Code & "\" & strFiles(Index)
            Set txtfile = fso.GetFile(Path & strFiles(Index))
            txtfile.Delete
        Next Index
    End If
    EmailAddress = gCompanies(i).NotificationEmail
    mapMess.Compose
    mapMess.RecipAddress = EmailAddress
    mapMess.AddressResolveUI = True
    mapMess.ResolveName
    mapMess.MsgSubject = "Your payroll files are ready."
    mapMess.MsgNoteText = "Your payroll files are ready."
    mapMess.Send
    DoEvents
End Sub

Private Sub ScanForMail()
Dim intLenFileName As Integer
Dim intStrPos As Integer
Dim i As Integer
Dim Name As String
Dim t As Integer
    mapSess.NewSession = True
    mapSess.DownLoadMail = True
    mapSess.UserName = "Input"
    mapSess.LogonUI = False
    mapSess.SignOn
    mapMess.SessionID = mapSess.SessionID
    mapMess.FetchUnreadOnly = True
    mapMess.Fetch
    If mapMess.MsgCount > 0 Then
        For t = mapMess.MsgCount - 1 To 0 Step -1
            mapMess.MsgIndex = t
            Name = mapMess.MsgSubject
            txtMessages.Text = vbCrLf & vbCrLf & Name & " " & Now & " " & vbCrLf & mapMess.MsgNoteText & txtMessages.Text
            D = date
            Two = InStrRev(D, "/")
            One = InStr(1, D, "/")
            Holder = "c:\DirectTransfer\" & Right("0" & Left(D, One - 1), 2) & Right("0" & Right(Left(D, Two - 1), Two - 1 - One), 2) & Right("0" & Right(D, Len(D) - Two), 2) & ".txt"
            sItemData = vbCrLf & vbCrLf & Name & " " & Now & " " & vbCrLf & mapMess.MsgNoteText
            FileNum = FreeFile
            Open Holder For Append As FileNum
            DoEvents
            Print #FileNum, , sItemData
            Close FileNum
            Call List1.AddItem("Transfer from " & Name & " complete.", List1.ListCount)
            If mapMess.AttachmentCount > 0 Then
                For i = 0 To mapMess.AttachmentCount - 1
                    mapMess.AttachmentIndex = i
                    intLenFileName = Len(mapMess.AttachmentPathName)
                    For intStrPos = intLenFileName To 1 Step -1
                        If InStr(1, Right$(mapMess.AttachmentPathName, intLenFileName - (intStrPos - 1)), "\", 1) Then
                            strNewFileName = Right$(mapMess.AttachmentPathName, intLenFileName - intStrPos)
                            Exit For
                        End If
                    Next intStrPos
                    Password = FindPassword(Name)
                    If Password <> "" Then
                        FileCopy mapMess.AttachmentPathName, "c:\DirectTransfer\InBox\" & Left(strNewFileName, 3) & "_zip.sen"
                        DoEvents
                        UnencryptZip Name
                        UnZipFiles Name
                    Else
                        FileCopy mapMess.AttachmentPathName, "c:\DirectTransfer\InBox\" & strNewFileName
                        DoEvents
                        UnZipFiles Name
                    End If
                Next i
            End If
            mapMess.Delete
        Next t
    End If
    mapSess.SignOff
    DoEvents
    Timer1.Enabled = True
End Sub

Private Function GetFiles(Path As String) As Integer
Dim i As Integer
Dim Holder As String
    Holder = Dir(Path, vbNormal + vbArchive)
    Do While Holder <> ""
        i = i + 1
        ReDim Preserve fstrFiles(0 To i) As String
        fstrFiles(i) = Holder
        Holder = Dir
    Loop
    If i = 0 Then
        ReDim fstrFiles(0) As String
        GetFiles = 0
    Else
        GetFiles = i
    End If
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    ReadCompanies
End Sub
