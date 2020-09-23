Attribute VB_Name = "basglobal"
Public Type FileType
    Name As String
    Passes As Integer
    Message As String
    FileNum As Integer
    BytesRecieved As Long
    Size As Long
End Type
Public File(200) As FileType

Public Type CompanyType
    Code As String
    Email As String
    NotificationEmail As String
End Type
Public gCompanies() As CompanyType
Public gButtonForecolor As Long
Public gLightForecolor As Long
