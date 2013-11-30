Attribute VB_Name = "mdlShare"
Option Explicit

Private Type gUpload
    uFilename As String
End Type
Private Type gUploads
    uEnabled As Boolean
    uCount As Integer
End Type
Private Type gFiles
    fEnabled As Boolean
    fFilename As String
    fFilepath As String
    fDescription As String
End Type
Private Type gFileshare
    fEnabled As Boolean
    
    fNickname As String
    fPassword As String
    fEmail As String
    fDirectory As String
    fFiles(1000) As gFiles
    fCount As Integer
    fReturn As String
    fFirstLoad As Boolean
End Type

Public lFileshare As gFileshare

Public Sub FileshareLoadSettings()
Dim msg As String, msg2, msg3 As String
lFileshare.fEnabled = GetSetting(App.Title, "FileshareSettings", "Enabled", False)
lFileshare.fNickname = GetSetting(App.Title, "FileshareSettings", "Nickname", "")
lFileshare.fPassword = GetSetting(App.Title, "FileshareSettings", "Password", "")
lFileshare.fEmail = GetSetting(App.Title, "FileshareSettings", "Email", "")
lFileshare.fDirectory = GetSetting(App.Title, "FileshareSettings", "Directory", "")
lFileshare.fCount = GetSetting(App.Title, "FileshareSettings", "Count", 0)
lFileshare.fFirstLoad = GetSetting(App.Title, "FileshareSettings", "FirstLoad", True)

If lFileshare.fEnabled = False Then
    msg2 = MsgBox("Would you like to enable your media filesharing network?", vbYesNo + vbQuestion, "Question ...")
    If msg2 = vbYes Then
        GoTo Start
    Else
        End
    End If
Else
    GoTo Start
End If

Exit Sub

Start:
If Len(lFileshare.fNickname) = 0 Then lFileshare.fNickname = InputBox("Please enter the nickname you would like to use", "Fileshare setup", "")
If Len(lFileshare.fPassword) = 0 Then lFileshare.fPassword = InputBox("Please enter the password you would like to use", "Fileshare setup", "")
If Len(lFileshare.fEmail) = 0 Then lFileshare.fEmail = InputBox("Enter your email address:", "Fileshare Setup")
If Len(lFileshare.fDirectory) = 0 Then
    frmFolder.Label1.Caption = "Please select the directory in which you store your media files, or select the directory in wish you'd like to store your media files"
    frmFolder.Show 1
    lFileshare.fDirectory = lFileshare.fReturn
    lFileshare.fReturn = ""
End If
If Len(lFileshare.fNickname) = 0 Or Len(lFileshare.fPassword) = 0 Or Len(lFileshare.fEmail) = 0 Or Len(lFileshare.fDirectory) = 0 Then
    GoTo Info_Blank
Else
    SaveSetting App.Title, "FileshareSettings", "Nickname", lFileshare.fNickname
    SaveSetting App.Title, "FileshareSettings", "Password", lFileshare.fPassword
    SaveSetting App.Title, "FileshareSettings", "Email", lFileshare.fEmail
    SaveSetting App.Title, "FileshareSettings", "Directory", lFileshare.fDirectory
    SaveSetting App.Title, "FileshareSettings", "FirstLoad", "False"
    SaveSetting App.Title, "FileshareSettings", "Enabled", "True"
    Dim i As Integer, X As Integer
    
    msg3 = FileSystem.Dir(lFileshare.fDirectory & "\", vbArchive)
    frmDir.Visible = False
    frmDir.Dir1.Path = lFileshare.fDirectory
    
    For i = 0 To frmDir.File1.ListCount
        If Right(LCase(frmDir.File1.List(i)), 3) = "mp3" Then AddFile frmDir.File1.Path & "\" & frmDir.File1.List(i), True
    Next i
End If
With frmShare
    .txtLocationOfMedia.Text = lFileshare.fDirectory
    .txtEmail.Text = lFileshare.fEmail
    .txtNickname.Text = lFileshare.fNickname
    .txtPassword.Text = lFileshare.fPassword
End With
Exit Sub

Info_Blank:
    msg = MsgBox("Not all items were proporly defined. This will cause abnormal operation of this program. Would you like to exit?", vbExclamation + vbYesNo, "Warning")
    If msg = vbYes Then
        End
    Else
        GoTo Start
    End If
End Sub

Public Function GetFTitle(lFilename As String) As String
On Local Error Resume Next

Dim msg As String
msg = lFilename
If Len(msg) <> 0 Then
Again:
    If InStr(msg, "\") Then
        msg = Right(lFilename, Len(msg) - InStr(msg, "\"))
        If InStr(msg, "\") Then
            GoTo Again
        Else
            GetFTitle = msg
        End If
    Else
        GetFTitle = msg
    End If
Else
    Exit Function
End If
End Function

Public Sub AddFile(lFilename As String, Optional lWrite As Boolean)
Dim lftitle As String, lfpath As String, lfull As String, i As Integer

lfull = lFilename
lftitle = GetFTitle(lfull)
lfpath = Left(lFilename, Len(lFilename) - Len(lftitle) - 1)
lFileshare.fCount = lFileshare.fCount + 1
i = lFileshare.fCount

lFileshare.fFiles(i).fDescription = Left(lFilename, Len(lFilename) - 4)
lFileshare.fFiles(i).fEnabled = True
lFileshare.fFiles(i).fFilename = lftitle
lFileshare.fFiles(i).fFilepath = lfpath

If lWrite = True Then
    SaveSetting App.Title, Str(i), "Files", lFilename
    SaveSetting App.Title, Str(i), "Description", lFileshare.fFiles(i).fDescription
    SaveSetting App.Title, Str(i), "Enabled", lFileshare.fFiles(i).fEnabled
    SaveSetting App.Title, Str(i), "Filepath", lFileshare.fFiles(i).fFilepath
    SaveSetting App.Title, Str(i), "Count", i
    frmShare.lblStatus.Caption = "Status: Adding " & lFileshare.fFiles(i).fFilename
    frmShare.Visible = True
    Pause 0.5
End If
End Sub
