Attribute VB_Name = "basExecute"
'    i = Button.Index
'    Execute_Program gPath(i - 2), " ", _
'                Left(gPath(i - 2), InStrRev(gPath(i - 2), "\"))

Option Explicit
'****************************************************************
'compatibility:VB3(ignore 32 bit dec and #if syntax), VB4-16,VB4-32, VB5
'Execute_Program
'Runs a program and handles all possible errors'
'(such as running out of memory, file can't be opened, etc.)'
'Also, unlike the VB Shell command, it allows you to specify'
'a 'default working directory'!
'Also, allows you to run a file that is only an association--
'example:you can run a .txt file with this function!'
'Inputs:strFilePath--program to runstrParms--program command line parms (if any)
'strDir--default working directory'
'Returns:returns TRUE=successful FALSE=failed'
'Assumes: None '
'Side Effects: None '
'**************************************************************
#If Win32 Then
'32 bit declare
Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
'16 bit declare
Declare Function ShellExecute Lib "shell.dll" _
    (ByVal hWnd As Integer, ByVal lpszOp As String, ByVal lpszFile As String, ByVal spszParams As String, ByVal lpszDir As String, ByVal fsShowCmd As Integer) As Integer
#End If

Const SW_SHOWNORMAL = 5

Function Execute_Program(ByVal strFilePath As String, _
    ByVal strParms As String, ByVal strDir As String) _
    As Integer

    'run program
    Dim hwndProgram As Integer
    hwndProgram = ShellExecute(0, "Open", strFilePath, strParms, strDir, SW_SHOWNORMAL)
  
  'evaluate errors (if any)
  Select Case (hwndProgram)
    Case 0
        MsgBox "Insufficent system memory or corrupt program file.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 2
        MsgBox "File not found.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 3
        MsgBox "Invalid path.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 5
        MsgBox "Sharing or Protection Error.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 6
        MsgBox "Seperate data segments are required for each task.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 8
        MsgBox "Insufficient memory to run the program.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
   Case 10
        MsgBox "Incorrect Windows version.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 11
        MsgBox "Invalid program file.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 12
        MsgBox "Program file requires a different operating system.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 13
        MsgBox "Program requires MS-DOS 4.0.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 14
        MsgBox "Unknown program file type.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 15
        MsgBox "Windows program does not support protected memory mode.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 16
        MsgBox "Invalid use of data segments when loading a second instance of a program.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 19
        MsgBox "Attempt to run a compressed program file.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 20
        MsgBox "Invalid dynamic link library.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    Case 21
        MsgBox "Program requires Windows 32-bit extensions.", 0, "Error running " & strFilePath
        Execute_Program = False
        Exit Function
    End Select
    
    Execute_Program = True
    
End Function

