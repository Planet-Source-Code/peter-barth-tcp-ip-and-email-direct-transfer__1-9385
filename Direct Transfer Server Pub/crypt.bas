Attribute VB_Name = "basCrypt"
Option Explicit

Global Const Encrypt = 1
Global Const Decrypt = 2

Global sKey As Double
Global sKey1 As Double
Global sKey2 As Double

Global sl() As Byte
Global sn() As Byte

Global pkey&(2), passw$, filename As String
Global passw2$, ZipCrypt As Boolean

Dim rbox(0 To 255) As Byte
Dim Key(0 To 1023) As Integer
Dim pwKey(0 To 1023) As Integer
Dim sbox(0 To 1023) As Integer
Dim mbox(0 To 16383) As Integer
Private Crc32Table(255) As Long
Public StartTimer As Single, StartTime As Single
Public SysTime$, SysDate$
Public TempFile1$, TempFile2$

Public Sub KillFiles()
If TempFile1$ > "" Then
    Kill TempFile1$
    TempFile1$ = ""
End If
If TempFile2$ > "" Then
    Kill TempFile2$
    TempFile2$ = ""
End If
End Sub

Public Sub EncryptZip(Name As String)
Dim A$, fname1$, fname2$, it&, x&, l%
Dim PbarVal As Single, IsZipped As Boolean
Dim Password As String
StartTimer = Timer
KillFiles
'----------------------- Get FileName -----------------------------
filename = "C:\DirectTransfer\OutBox\" & Name & ".zip"
If Dir(filename) = "" Then Exit Sub

'----------------------- Get Password -----------------------------
Password = FindPassword(Name)
l% = Len(filename)
If l% > 0 Then
    A$ = Password
    makePwKey (A$)
End If

'---------------------- En / Decrypt File -------------------------
l% = Len(filename)
If l% > 0 Then
        fname1$ = filename
        fname2$ = Left(filename, InStr(1, filename, ".") - 1) & "_" & Right(filename, Len(filename) - InStr(1, filename, ".")) & ".sen"
        makeSKey (StartTimer)
        makeStorePassw
        makeSBox
        SetTime
'-------------------- Read & Write 16k at a time ------------------
        Open fname1$ For Binary Access Read Write As 1
        Open fname2$ For Binary Access Write As 2
            
            Put #2, , passw$            'save encr. sKeys
            it& = LOF(1) \ 16384        'how many 16k buffers?
            PbarVal = 100 / (it& + 1)   'ProgressBar update incr.
            If it& Then
                ReDim sl(16383)
                ReDim sn(16383)
                makeMBox 16384
                For x& = 1 To it&
                    Get #1, , sl()
                    encrypt_it
                    Put #2, , sn()
                Next
            End If
            x& = LOF(1) Mod 16384   'Left-overs?
            If x& Then
                ReDim sl(x& - 1)
                ReDim sn(x& - 1)
                makeMBox (x&)   'a MUST! it's < 16k
                Get #1, , sl()
                encrypt_it
                Put #2, , sn()
            End If
        Close #2

'------------------------------------------------------------------
            If Right$(fname1$, 2) = "sn" Or Right$(fname1$, 2) = "sz" Then
                TempFile1$ = fname1$
                WipeFile 1, it&
            Else
                Close #1
            End If
            ResetTime
'-------------------- Read & Write 16k at a time ------------------
 
       Kill filename
End If
End Sub


Public Sub UnencryptZip(Name As String)
Dim A$, fname1$, fname2$, it&, x&, l%
Dim hIn, fileLength
Dim PbarVal As Single, IsZipped As Boolean
Dim Password As String
StartTimer = Timer
KillFiles
'----------------------- Get FileName -----------------------------
filename = "C:\DirectTransfer\InBox\" & Name & "_zip.sen"
hIn = FreeFile
Open filename For Binary Access Read As hIn
fileLength = LOF(hIn)
Close hIn
If fileLength = 0 Then Exit Sub
'----------------------- Get Password -----------------------------
Password = FindPassword(Name)
l% = Len(filename)
If l% > 0 Then
    A$ = Password
    makePwKey (A$)
End If

'---------------------- En / Decrypt File -------------------------
l% = Len(filename)
If l% > 0 Then
        fname1$ = Left$(filename, Len(filename) - 4)
        Open filename For Binary Access Read Write As 1
            
            A$ = Space$(120)
            Get #1, , A$        'retrieve sKeys
            makeLoadPassw A$
            makeSBox
        
        Open fname1$ For Binary Access Write As 2
                        
            it& = (LOF(1) - 120) \ 16384     'see notes above
            PbarVal = 100 / (it& + 1)
            If it& Then
                ReDim sl(16383)
                ReDim sn(16383)
                makeMBox 16384
                For x& = 1 To it&
                    Get #1, , sn()
                    decrypt_it
                    Put #2, , sl()
                Next
            End If
            x& = (LOF(1) - 120) Mod 16384    'Left_overs?
            If x& Then
                ReDim sl(x& - 1)
                ReDim sn(x& - 1)
                makeMBox (x&)           'MUST do!
                Get #1, , sn()
                decrypt_it
                Put #2, , sl()
            End If
            TempFile1$ = filename
            WipeFile 1, it&
       Close #2
End If
End Sub

Public Sub SetTime()
'---- because our sKeys are date & time dependent let's hide ------
'------ the true time and date of encryption file creation. -------
    
    StartTime = Timer
    SysDate$ = date$
    date$ = "01-01-1980"
    Time$ = "12:00:00"
End Sub

Public Sub ResetTime()
'Let's reset the true system time & date
Dim t1&, temp1&, temp2&, temp3&

t1& = Timer - 43200 + StartTime
temp1& = t1& \ 3600
t1& = t1& - temp1& * 3600
temp2& = t1& \ 60
temp3& = t1& - temp2& * 60
Time = CStr(temp1&) + ":" + CStr(temp2&) + ":" + CStr(temp3&)
date$ = SysDate$
End Sub

Public Sub WipeFile(FileNum As Byte, num As Long)
Dim x&, y As Byte, offset&, Dummy1$, Dummy2$

Dummy1$ = "XCZVDSKL.BJK"
Dummy2$ = "KJBGFCDI.PFW"

For y = 1 To 3
    Seek #FileNum, 1
    For x& = 1 To num
        Put #FileNum, , String$(16384, Chr$(x& And &HFF))  'fill file w/Junk
    Next
    x& = LOF(FileNum) Mod 16384
    Put #FileNum, , String$(x&, Chr$(x& And &HFF))         'ditto
Next
Close #FileNum

If TempFile1$ > "" Then
    Name TempFile1$ As Dummy1$
    TempFile1$ = Dummy1$
End If
If TempFile2$ > "" Then
    Name TempFile2$ As Dummy2$
    TempFile2$ = Dummy2$
End If
    
'patience with the killing of the files! If it is done as the next
'command it'll kill the file before the i/o is done and the file
'will be deleted before it is overwritten. I got gray hair over this!
End Sub

'------------------------ Start Encryption -------------------
Sub encrypt_it()
Dim x&, l&, k%, i%, j%, c As Byte

l& = UBound(sl())

'make bytes interdependent forwards & backwards
For x& = 2 To l&
    k% = sl(x& - 2) Xor mbox(x&) And &HFF
    c = rotR(ByVal sl(x&), ByVal rbox(Key(k%)))
    sl(x&) = c Xor (k% + sl(x& - 1)) And &HFF
Next
For x& = l& - 2 To 0 Step -1
    k% = sl(x& + 2) Xor sl(x& + 1)
    c = rotL(ByVal sl(x&), ByVal rbox(Key(k%))) 'left rotate
    sl(x&) = c Xor (k% + sl(x& + 1)) And &HFF
Next

k% = Key(921)   'initial offset - different for each 16k buffer
j% = Key(113)

For x& = 0 To l&
    i% = x& Mod 1023
    
    'generate new offset values <= array size
    k% = (k% + mbox(x&)) Mod 1023
    j% = (j% + k%) Mod 1023
    
    'mix-up the pwKey & sbox
    c = pwKey(i%)
    pwKey(i%) = pwKey(j%)
    pwKey(j%) = pwKey(k%)
    pwKey(k%) = sbox(i%)
    sbox(i%) = sbox(j%)
    sbox(j%) = sbox(k%)
    sbox(k%) = c
    
    'Scramble & XOR
    c = sl(mbox(x&)) Xor (sbox(j%) + sbox(k%)) And &HFF Xor pwKey(i%)
    c = (c + (pwKey(i%) + pwKey(j%)) And &HFF) And &HFF Xor Key(j%)
    c = c Xor (Key(i%) + Key(j%)) And &HFF Xor sbox(i%)
    sn(x&) = (c - sbox(i%) + 256) And &HFF Xor pwKey(j%)
Next

'Generate a new Key for next 16k buffer
For x& = 0 To 1023
    Key(x&) = Key(x&) Xor pwKey(x&)
Next

'mix-up the mbox for next 16k buffer so that the next buffer
'will be scrambled differently.
k% = Key(1) + 37       'offset

For x& = k% + 1 To l&
    i% = mbox(x&)
    mbox(x&) = mbox(x& - k%)
    mbox(x& - k%) = i%
Next
End Sub

'---------------- Start Decrypting ---------------------------
'--- for comments see Encrypt

Sub decrypt_it()
Dim l&, x&, k%, i%, j%, c As Byte

l& = UBound(sn())

k% = Key(921)
j% = Key(113)

For x& = 0 To l&
    i% = x& Mod 1023
    k% = (k% + mbox(x&)) Mod 1023
    j% = (j% + k%) Mod 1023
    
    c = pwKey(i%)
    pwKey(i%) = pwKey(j%)
    pwKey(j%) = pwKey(k%)
    pwKey(k%) = sbox(i%)
    sbox(i%) = sbox(j%)
    sbox(j%) = sbox(k%)
    sbox(k%) = c
    
    c = sn(x&) Xor pwKey(j%)
    c = (c + sbox(i%)) And &HFF
    c = c Xor (Key(i%) + Key(j%)) And &HFF Xor sbox(i%)
    c = c Xor Key(j%)
    c = (c - (pwKey(i%) + pwKey(j%)) And &HFF + 256) And &HFF
    sl(mbox(x&)) = c Xor (sbox(j%) + sbox(k%)) And &HFF Xor pwKey(i%)
Next

For x& = 0 To l& - 2
    k% = sl(x& + 2) Xor sl(x& + 1)
    c = sl(x&) Xor (k% + sl(x& + 1)) And &HFF
    sl(x&) = rotR(ByVal c, ByVal rbox(Key(k%)))
Next
For x& = l& To 2 Step -1
    k% = sl(x& - 2) Xor mbox(x&) And &HFF
    c = sl(x&) Xor (k% + sl(x& - 1)) And &HFF
    sl(x&) = rotL(ByVal c, ByVal rbox(Key(k%)))
Next

For x& = 0 To 1023
    Key(x&) = Key(x&) Xor pwKey(x&)
Next

k% = Key(1) + 37

For x& = k% + 1 To l&
    i% = mbox(x&)
    mbox(x&) = mbox(x& - k%)
    mbox(x& - k%) = i%
Next
End Sub

Sub makeSKey(StartTimer As Single)
Dim sKy$, sKy1$, sKy2$, temp$, pos%, k%, x%, t&

'------------------- Generate 3 session keys ------------------------
'any values will do as long as they're small enough to change rapidly

t& = Julian("s")    'Seconds
temp$ = pkey(0) / t& / Timer
sKey = Timer / (Val(temp$) * 365 * 24 * 3600 + Timer) * (pkey&(0) / Timer)

Randomize sKey
RandomScrap

sKey = (sKey ^ 0.315 + Timer ^ 0.763) * 23.237 * Rnd
temp$ = Left$(CStr(sKey), 16)
pos% = InStr(temp$, ".")
sKy$ = Left$(temp$, pos% - 1) + Right$(temp$, Len(temp$) - pos%)
sKey = Val(sKy$)

k% = Val(Mid$(sKy$, 3, 4))

Do                          'mix the numbers around
For x% = 1 To Len(sKy$)
    k% = (k% + Asc(Mid$(sKy$, x%, 1))) Mod Len(sKy$) + 1
    temp$ = Mid$(sKy$, x%, 1)
    Mid$(sKy$, x%, 1) = Mid$(sKy$, k%, 1)
    Mid$(sKy$, k%, 1) = temp$
Next
Loop Until Left$(sKy$, 1) <> "0"    'no leading zeros

sKy$ = Left$(sKy$ + Mid$(sKy$, 3, 4), 15)   '15 digits
sKey = Val(sKy$)

'-------------------- 2nd sKey ----------------------------
t& = Julian("n")   'Minutes (don't know why min. is "n"!)
temp$ = pkey(1) / t& / Timer
sKey1 = StartTimer / (Val(temp$) * 30 * 24 * 3600) * (pkey&(1) / StartTimer)

Randomize sKey1
RandomScrap

sKey1 = (sKey1 ^ 0.247 + Timer ^ 0.873) * 17.526 * Rnd
temp$ = Left$(CStr(sKey1), 16)
pos% = InStr(temp$, ".")
sKy1$ = Left$(temp$, pos% - 1) + Right$(temp$, Len(temp$) - pos%)
sKey1 = Val(sKy1$)

Do
For x% = 1 To Len(sKy1$)
    k% = (k% + Asc(Mid$(sKy1$, x%, 1))) Mod Len(sKy1$) + 1
    temp$ = Mid$(sKy1$, x%, 1)
    Mid$(sKy1$, x%, 1) = Mid$(sKy1$, k%, 1)
    Mid$(sKy1$, k%, 1) = temp$
Next
Loop Until Left$(sKy1$, 1) <> "0"

sKy1$ = Left$(sKy1$ + Mid(sKy1$, 4, 4), 15)
sKey1 = Val(sKy1$)

'-------------------- 3rd sKey ----------------------------
t& = Julian("h")    'Hours
temp$ = pkey(2) / t& / Timer
sKey2 = StartTimer / (Val(temp$) * 24 * 3600) * (pkey&(2) / (Timer + StartTimer))

Randomize sKey2
RandomScrap

sKey2 = (sKey2 ^ 0.277 + Timer ^ 0.457) * 53.341 * Rnd
temp$ = Left$(CStr(sKey2), 16)
pos% = InStr(temp$, ".")
sKy2$ = Left$(temp$, pos% - 1) + Right$(temp$, Len(temp$) - pos%)
sKey2 = Val(sKy2$)

Do
For x% = 1 To Len(sKy2$)
    k% = (k% + Asc(Mid$(sKy2$, x%, 1))) Mod Len(sKy2$) + 1
    temp$ = Mid$(sKy2$, x%, 1)
    Mid$(sKy2$, x%, 1) = Mid$(sKy2$, k%, 1)
    Mid$(sKy2$, k%, 1) = temp$
Next
Loop Until Left$(sKy2$, 1) <> "0"

sKy2$ = Left$(sKy2$ + Mid(sKy2$, 4, 4), 15)
sKey2 = Val(sKy2$)
End Sub

Sub makePwKey(A$)
Dim b$, x%, y%, l%, k%, pky&

pky& = 0
l% = Len(A$)
passw2$ = ""
b$ = StrReverse(A$)

GetPkey A$, pkey&(0)
GetPkey b$, pkey&(1)
GetPkey A$ + b$, pkey&(2)

'make a Random Number seed
For x% = 1 To l%
    k% = (k% + Asc(Mid$(A$, x%, 1))) Mod 16384
    pky& = pky& + x% * Asc(Mid$(A$, x%, 1)) ^ 1.27 + k%
Next

Rnd -1
Randomize pkey&(0)
RandomScrap     'ditch the first few rnd nums

'initialize the pwKey array
For x% = 0 To 1023
    k% = (k% + Asc(Mid$(A$, x% Mod l% + 1))) Mod 857
    pwKey(x%) = Int(Rnd * pky& * (x% + 1) + k%) And &HFF
Next

k% = Asc(Mid$(A$, l% \ 2, 1))   'offset

'scramble the pwKey array
For x% = 0 To 1023
    k% = (k% + pwKey(x%) + Asc(Mid$(A$, x% Mod l% + 1, 1))) Mod 1023
    y% = pwKey(x%)
    pwKey(x%) = pwKey(k%)
    pwKey(k%) = y%
Next

'add the hex$(crc) to the password
A$ = Hex$(pkey&(0)) + A$ + Hex$(pkey&(1)) + b$ + Hex$(pkey&(2))

'make the entered password longer & encrypt it
For x% = 1 To 1200
    k% = (k% + pwKey(k%)) Mod 1023
    A$ = A$ + Chr$(((Asc(Mid$(A$, x%, 1)) + Asc(Mid$(A$, x% + 1, 1))) And &HFF) Xor pwKey(k%))
Next

A$ = Mid$(A$, 151, 1024)

'Scramble the password
For x% = 1 To 1024
    k% = (k% + x% + pwKey(x% - 1) + Asc(Mid$(A$, x%, 1))) Mod 1024 + 1
    b$ = Mid$(A$, x%, 1)
    Mid$(A$, x%, 1) = Mid$(A$, k%, 1)
    Mid$(A$, k%, 1) = b$
Next


'finalize the pwKey by XORing it against the password
For x% = 0 To 1023
    k% = (k% + pwKey(x%)) Mod 1023
    pwKey(x%) = Asc(Mid$(A$, k% + 1, 1)) Xor (pwKey(k%) + pwKey(x%)) And &HFF
Next

'make PkZip password for PkZip encryption option
For x% = 1 To 15
    k% = pwKey(x%) Mod 126
    If k% < 35 Then k% = k% + 35
    passw2$ = passw2$ + Chr$(k%)
Next
End Sub

Sub makeSBox()
Dim x%, k%

Rnd -1
Randomize sKey
RandomScrap     'ditch a few

k% = pwKey(Rnd * 1023) 'initial offset

'make sbox
For x% = 0 To 1023
    k% = (k% + pwKey(x%)) Mod 1023
    sbox(x%) = Int(Rnd * 256) Xor (pwKey(x%) + pwKey(k%)) And &HFF
Next

Randomize sKey1
RandomScrap     'ditch a few

'make Key
For x% = 0 To 1023
    k% = (k% + sbox(x%)) Mod 1023
    Key(x%) = Int(Rnd * 238) Xor (pwKey(k%) + sbox(x%)) And &HFF
Next
End Sub

Sub makeMBox(l&)
Dim x%, temp%, k%
Erase mbox

Rnd -1
Randomize sKey2
RandomScrap     'ditch a few

k% = pwKey(Rnd * 1023)  'initial offset
l& = l& - 1

'first fill consecutively
For x% = 0 To l&
    mbox(x%) = x%
Next

'and now mix them around
For x% = 0 To l&
    k% = (k% + pwKey(x% Mod 1023) + Int(Rnd * l&)) Mod l& + 1
    temp% = mbox(x%)
    mbox(x%) = mbox(k%)
    mbox(k%) = temp%
Next

Randomize Abs(sKey - sKey1)
RandomScrap

'set up the rbox for rotating
For x% = 0 To 255
    rbox(x%) = Int(Rnd * 8)
Next

k% = pwKey(15)

For x% = 0 To 255
    k% = (k% + pwKey(k%)) And &HFF
    temp% = rbox(x%)
    rbox(x%) = rbox(k%)
    rbox(k%) = temp%
Next
End Sub

Sub makeStorePassw()
'--------------------------------------------------------------
'Since the session keys are vital to re-establish our encryption
'arrays and these keys are stored along with the file it is im-
'perative that these keys be very secure. We will create another
'60 digits at random, solve a sum of our 3 session keys, for a
'total of 120 digits and then mix them up and encrypt them.
'--------------------------------------------------------------

Dim sKy$, sKy1$, sKy2$, s%, x%, y%, z%
Dim sum$, dummy$, mb(1 To 120) As Byte

Randomize
RandomScrap

'make 60 dummy digits and XOR them against the pwKey
For x% = 1 To 60
    dummy$ = dummy$ + CStr(Int(Rnd * 9))
    Mid$(dummy$, x%, 1) = Chr$(Asc(Mid$(dummy$, x%, 1)) Xor pwKey(60 + x%))
Next

'set up for the real keys
passw$ = Space$(60)
sKy$ = CStr(sKey)
sKy1$ = CStr(sKey1)
sKy2$ = CStr(sKey2)
sum$ = Left$(CStr(sKey + sKey1 + sKey2), 15)    'a password check

For x% = 1 To 15
    y% = x% + 15
    z% = x% + 30
    s% = x% + 45
    Mid$(passw$, x%, 1) = Chr$(pwKey(x%) Xor Asc(Mid$(sKy$, x%, 1)))
    Mid$(passw$, y%, 1) = Chr$(pwKey(y%) Xor Asc(Mid$(sKy1$, x%, 1)))
    Mid$(passw$, z%, 1) = Chr$(pwKey(z%) Xor Asc(Mid$(sKy2$, x%, 1)))
    Mid$(passw$, s%, 1) = Chr$(pwKey(s%) Xor Asc(Mid$(sum$, x%, 1)))
Next

'join dummy & real
passw$ = passw$ + dummy$

s% = pwKey(745)     'offset

'make a small mbox for mixing
For x% = 1 To 120
    mb(x%) = x%
Next

For x% = 1 To 120
    s% = (s% + pwKey(x%)) Mod 120 + 1
    z% = mb(x%)
    mb(x%) = mb(s%)
    mb(s%) = z%
Next

'scramble all 120 digits
sKy$ = passw$
For x% = 1 To 120
    Mid$(passw$, x%, 1) = Mid$(sKy$, mb(x%), 1)
Next

'reverse the string
passw$ = StrReverse(passw$)

'and encrypt the whole thing one more time
For x% = 1 To 120
    Mid$(passw$, x%, 1) = Chr$(Asc(Mid$(passw$, x%, 1)) Xor pwKey(x%))
Next
End Sub

Sub makeLoadPassw(A$)
'for comments see makeStorePassw()

Dim sKy$, sKy1$, sKy2$, s%, x%, y%, z%
Dim sum$, mb(1 To 120) As Byte

For x% = 1 To 120
    Mid$(A$, x%, 1) = Chr$(Asc(Mid$(A$, x%, 1)) Xor pwKey(x%))
Next

A$ = StrReverse(A$)

s% = pwKey(745)

'rebuild our mixing box
For x% = 1 To 120
    mb(x%) = x%
Next

For x% = 1 To 120
    s% = (s% + pwKey(x%)) Mod 120 + 1
    z% = mb(x%)
    mb(x%) = mb(s%)
    mb(s%) = z%
Next

'lets unscramble that mess
sKy$ = A$
For x% = 1 To 120
    Mid$(A$, mb(x%), 1) = Mid$(sKy$, x%, 1)
Next

A$ = Left$(A$, 60)      'will the real sKeys please stand up?

sKy$ = Space$(15)
sKy1$ = Space$(15)
sKy2$ = Space$(15)
sum$ = Space$(15)

For x% = 1 To 15
    y% = x% + 15
    z% = x% + 30
    s% = x% + 45
    Mid$(sKy$, x%, 1) = Chr$((pwKey(x%)) Xor Asc(Mid$(A$, x%, 1)))
    Mid$(sKy1$, x%, 1) = Chr$((pwKey(y%)) Xor Asc(Mid$(A$, y%, 1)))
    Mid$(sKy2$, x%, 1) = Chr$((pwKey(z%)) Xor Asc(Mid$(A$, z%, 1)))
    Mid$(sum$, x%, 1) = Chr$((pwKey(s%)) Xor Asc(Mid$(A$, s%, 1)))
Next

're-establish the sKeys
sKey = Val(sKy$)
sKey1 = Val(sKy1$)
sKey2 = Val(sKy2$)
sKy$ = Left$(CStr(sKey + sKey1 + sKey2), 15)

If sum$ <> sKy$ Then    'is the password correct?
    MsgBox "*** Invalid Password ***", vbOKOnly, "Error"
    Close #1
    End                 'too bad!
End If

End Sub

Function rotL(ByVal num As Byte, ByVal pos As Byte) As Byte
Dim x As Byte

For x = 1 To pos
    If num And 128 Then
        num = num * 2 - 255
    Else
        num = num * 2
    End If
Next
rotL = num
End Function

Function rotR(ByVal num As Byte, ByVal pos As Byte) As Byte
Dim x As Byte

For x = 1 To pos
    If num And 1 Then
        num = (num + 255) \ 2
    Else
        num = num \ 2
    End If
Next
rotR = num
End Function

Sub RandomScrap()       'discard the first bunch of rnd bytes
Dim x%, c%, r As Single
c% = Int(Rnd * 256)
For x% = 1 To c%
    r = Rnd(r)
Next
End Sub

Function Julian(t$) As Long
'calculate the time passed since Midnight Jan. 1, 1980
    Julian = DateDiff(t$, "01/01/1980", Now)
End Function

Public Function InitCrc32(Optional ByVal seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    Dim iBytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    On Error Resume Next

    For iBytes = 0 To 255
        lCrc32 = iBytes

        For iBits = 0 To 7
            'Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF

            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor seed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        
        Crc32Table(iBytes) = lCrc32
    Next
    
    InitCrc32 = Precondition

End Function
'End of initialization function


Public Function AddCrc32(ByVal Item As String, ByVal Crc32 As Long) As Long
    Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    
    On Error Resume Next

    For iCounter = 1 To Len(Item)
        bCharValue = Asc(Mid$(Item, iCounter, 1))
        
        lAccValue = Crc32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        
        lIndex = Crc32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = Crc32Table(lIndex)
        
        Crc32 = lAccValue Xor lTableValue
    Next
    
    AddCrc32 = Crc32
End Function

Public Function GetCrc32(ByVal Crc32 As Long) As Long
    On Error Resume Next
    
    GetCrc32 = Crc32 Xor &HFFFFFFFF
End Function

Sub GetPkey(A$, var&)
    Dim lCrc32Value As Long
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(A$, lCrc32Value)
    var& = Abs(GetCrc32(lCrc32Value))
End Sub

Public Function FindPassword(Company As String) As String
Dim fso As New FileSystemObject
Dim Holder As String
Dim Third As Integer
Dim f, ts
    FindPassword = ""
    Set f = fso.GetFile("c:\DirectTransfer\Companies.idi")
    Set ts = f.OpenAsTextStream(1, -2)
    Do Until ts.AtEndOfStream = True
        Holder = ts.ReadLine
        If Holder <> "" Then
            Third = InStrRev(Holder, ",")
            If Company = Left(Holder, 3) Then
                FindPassword = Right(Holder, Len(Holder) - Third)
            End If
        End If
    Loop
End Function
