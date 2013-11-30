Attribute VB_Name = "mdlSubs"
Option Explicit

Private Type gWindowPos
    wTitleBarHeight As Integer
    wWindowBorder As Integer
End Type

Global lMainWndSettings As gWindowPos

Public Function GetRnd(Num As Integer) As Integer
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
End Function

Public Sub InitDisplay()
Dim i As Integer
i = GetSetting(App.Title, "Main", "Interface", 1)
Select Case i
Case 0
    SetInterface eSmWindow, False, True
Case 1
    SetInterface eSmWindow, False, True
Case 2
    SetInterface eSmWindow, False, True
Case 3
    SetInterface eUtilityWindow, False, True
Case 4
    SetInterface eNexENCODEWindow, False, True
End Select
End Sub

Public Function GetFileTitle(lFilename As String) As String
On Local Error Resume Next

If Len(lFilename) <> 0 Then
Again:
    If InStr(lFilename, "\") Then
        lFilename = Right(lFilename, Len(lFilename) - InStr(lFilename, "\"))
        If InStr(lFilename, "\") Then
            GoTo Again
        Else
            GetFileTitle = lFilename
        End If
    Else
        GetFileTitle = lFilename
    End If
Else
    Exit Function
End If
End Function

Public Sub CloseMp3Player()
lInterface.iStoped = True
frmMain.ctlMp3Player.Stop
frmMain.ctlMp3Player.Close
frmMain.ctlMp3Player.SetOutDevice lSettings.sOutputDevice
frmMain.lblFileInfo.Caption = ""
frmMain.tmrStreamTitle.Enabled = False
frmMain.tmrFake.Enabled = False
Pause 0.5
End Sub

Public Function AddtoPlaylist(lFilename As String) As Integer
Dim i As Integer, fPath As String, fTitle As String, msg As String, msg2 As String, msg4 As String
On Local Error Resume Next

If DoesFileExist(lFilename) = False Then Exit Function
If InStr(LCase(lFilename), "mp3") Then
    msg2 = lFilename
    msg = Left(lFilename, Len(lFilename) - 1)
    If Right(LCase(msg), 3) = "mp3" Then lFilename = msg
    If Len(lFilename) <> 0 And Right(LCase(lFilename), 3) = "mp3" Then
        fTitle = GetFileTitle(msg2)
        AddRecientMedia fTitle
        fPath = Left(lFilename, Len(lFilename) - Len(fTitle))
        If DoesEntryExist(fTitle) = True Then Exit Function
        lPlaylist.pCount = lPlaylist.pCount + 1
        i = lPlaylist.pCount
        lPlaylist.pFiles(i).fFilename = fTitle
        lPlaylist.pFiles(i).fFilepath = fPath
        lPlaylist.pFiles(i).fFiletype = Mp3_File
        AddtoPlaylist = i
    End If
ElseIf InStr(LCase(lFilename), "wav") Then
    msg2 = lFilename
    msg = Left(lFilename, Len(lFilename) - 1)
    If Right(LCase(msg), 3) = "wav" Then lFilename = msg
    If Len(lFilename) <> 0 And Right(LCase(lFilename), 3) = "wav" Then
        fTitle = GetFileTitle(msg2)
        AddRecientMedia fTitle
        fPath = Left(lFilename, Len(lFilename) - Len(fTitle))
        If DoesEntryExist(fTitle) = True Then Exit Function
        lPlaylist.pCount = lPlaylist.pCount + 1
        i = lPlaylist.pCount
        lPlaylist.pFiles(i).fFilename = fTitle
        lPlaylist.pFiles(i).fFilepath = fPath
        lPlaylist.pFiles(i).fFiletype = Wav_File
        AddtoPlaylist = i
    End If
End If
DoEvents
End Function

Public Function FindPlaylistIndex(lSearch As String) As Integer
On Local Error Resume Next

Dim i As Integer
For i = 1 To lPlaylist.pCount
    If InStr(LCase(lPlaylist.pFiles(i).fFilename), LCase(lSearch)) Then
        FindPlaylistIndex = i
        Exit For
    End If
Next i
End Function

Public Function DoesEntryExist(lSearch As String) As Boolean
On Local Error Resume Next

Dim i As Integer

For i = 1 To lPlaylist.pCount
    If InStr(LCase(lPlaylist.pFiles(i).fFilename), LCase(lSearch)) Then
        DoesEntryExist = True
        Exit For
    End If
Next i
End Function

Public Sub RegisterComponents()
'frmMain.ctlMp3Player.Authorize "Leon J Aiossa", "812144397"
End Sub

Public Sub LoadPlaylist()
On Local Error Resume Next
Dim i As Integer, x As Integer, msg As String

lPlaylist.pFilename = GetSetting(App.Title, "Playlist", "Filename", App.Path & "\" & "audica.m3u")
x = GetSetting(App.Title, "Playlist", "Count", 0)
If x <> 0 Then
    For i = 1 To x
        msg = GetSetting(App.Title, "Playlist", i, "")
        If Len(msg) <> 0 Then AddtoPlaylist msg
    Next i
End If
End Sub

Public Sub SavePlaylist(Optional lFilename As String)
On Local Error Resume Next

Start:
If Len(lFilename) = 0 Then
    lFilename = SaveDialog(frmMain, "M3u Files (*.m3u)|*.m3u|All Files (*.*)|*.*", "Audica - Save as?", CurDir)
    If Len(lFilename) = 0 Then Exit Sub
    lFilename = Left(lFilename, Len(lFilename) - 1)
End If

If Right(LCase(lFilename), 4) <> ".m3u" Then lFilename = lFilename & ".m3u"

If DoesFileExist(lFilename) = True Then
    Dim msg As String
    msg = MsgBox("The file '" & lFilename & "' already exists. Would you like to save it as something else?", vbYesNo + vbQuestion)
    If msg = vbYes Then
        lFilename = ""
        GoTo Start
    ElseIf msg = vbNo Then
        GoTo Save
    End If
Else
Save:
    Dim i As Integer, msg2 As String
    For i = 1 To lPlaylist.pCount
        With lPlaylist.pFiles(i)
            If Len(.fFilename) <> 0 And Len(.fFilepath) <> 0 And .fFiletype <> 0 Then
                If Len(msg2) = 0 Then
                    msg2 = .fFilepath & .fFilename
                Else
                    msg2 = msg2 & vbCrLf & .fFilepath & .fFilename
                End If
            End If
        End With
    Next i
    SaveFile lFilename, msg2
    lPlaylist.pFilename = lFilename
    SaveSetting App.Title, "Playlist", "Filename", lFilename
End If
End Sub

Public Function SaveFile(lFilename As String, lText As String) As Boolean
On Local Error Resume Next

If Len(lFilename) <> 0 And Len(lText) <> 0 Then
    Open lFilename For Output As #1
    Print #1, lText
    Close #1
End If
End Function

Public Sub SaveSettings()
Dim i As Integer
For i = 1 To lPlaylist.pCount
    SaveSetting App.Title, "Playlist", i, lPlaylist.pFiles(i).fFilepath & lPlaylist.pFiles(i).fFilename
Next i
SaveSetting App.Title, "Playlist", "Count", lPlaylist.pCount
SaveSetting App.Title, "Main", "Interface", lInterface.iCurrentLayout
SaveSetting App.Title, "Main", "Left", frmMain.Left
SaveSetting App.Title, "Main", "Top", frmMain.Top
End Sub

Public Sub SetNexENCODEShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, tmp As Long
Dim x As Long, Y As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateEllipticRgn(0, 0, frmMain.Width, frmMain.Height) ' whole image
rgn1 = CreateEllipticRgn(x + 147, Y + 90, x + 326, Y + 267) 'big crl in back (transparency)
rgn2 = CreateEllipticRgn(x + 104, Y + 46, x + 367, Y + 310) ' big crl in back
rgn3 = CreateEllipticRgn(x + 48, Y + 74, x + 257, Y + 287) 'left crl
rgn4 = CreateEllipticRgn(x + 65, Y + 92, x + 241, Y + 268) 'left crl (transperancy)
rgn5 = CreateEllipticRgn(x + 212, Y + 72, x + 422, Y + 286) 'right crl
rgn6 = CreateEllipticRgn(x + 230, Y + 91, x + 404, Y + 268) 'right crl (transparency)
rgn7 = CreateRoundRectRgn(x + 39, Y + 120, x + 429, Y + 237, 110, 110)  'pill

tmp = CombineRgn(rgn1, rgn2, rgn1, RGN_DIFF) ' back crl
tmp = CombineRgn(rgn3, rgn3, rgn4, RGN_DIFF) ' left crl
tmp = CombineRgn(rgn5, rgn5, rgn6, RGN_DIFF) 'right crl

tmp = CombineRgn(rgn3, rgn3, rgn5, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn7, RGN_OR)
tmp = CombineRgn(rgn, rgn1, rgn3, RGN_OR)
frmMain.Width = 7000
frmMain.Height = 6000
tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub

Public Sub AppendToPlaylist(Optional lFilename As String)
Dim msg As String, msg2 As String, lefty As String
On Local Error Resume Next

If Len(lFilename) = 0 Then
    lFilename = OpenDialog(frmMain, "M3u files (*.m3u)|*.m3u|All Files (*.*)|*.*", "Audica - Select playlist ...", CurDir)
    If Len(lFilename) = 0 Then Exit Sub
End If

If DoesFileExist(lFilename) = True Then
    msg = ReadFile(lFilename)
    msg = Trim(msg)
Again:
    If Len(msg) <> 0 Then
        lefty = Left(msg, 1)
        msg2 = lefty & ParseString(msg, Left(msg, 1), Chr(13))
        msg = Right(msg, Len(msg) - Len(msg2) - 2)
        AddtoPlaylist msg2
        DoEvents
        If Len(msg) <> 0 Then
            If InStr(msg, Chr(13)) Then
                GoTo Again
            Else
                AddtoPlaylist Trim(msg)
                DoEvents
            End If
        End If
    End If
Else
    MsgBox "File does not exist"
End If
End Sub

Public Sub ClearPlaylist()
Dim i As Integer
On Local Error Resume Next

For i = 1 To lPlaylist.pCount
    With lPlaylist.pFiles(i)
        .fFilename = ""
        .fFilepath = ""
        .fFiletype = 0
    End With
Next i
lPlaylist.pCurrent = ""
For i = 1 To frmMain.mnuRecient.Count
    Unload frmMain.mnuRecient(i)
Next i

SaveSettings
End Sub

Public Sub RemoveFromPlaylist(lIndex As Integer)

End Sub

Public Sub SetInterface(lInterfaceType As eCurrentLayout, Optional lFadeOut As Boolean, Optional lInitVis As Boolean)
On Local Error Resume Next
Select Case lInterfaceType
Case eSmWindow
    If lFadeOut = True Then FadeOut
    DoEvents
    lInterface.iCurrentLayout = eSmWindow
    GetWindowSettings frmMain.hwnd
    SetPlayerShape
    With frmMain
        .mnuAudica.Checked = True
        .mnuUtility.Checked = False
        .mnuNexENCODE.Checked = False
        .imgLayout.Top = 0
        .imgLayout.Left = 0
        .imgLayout.Visible = True
        .imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture
        .Caption = "::AUDICA.PLAYER::"
        
'        .Spectrum1.Left = 57
'        .Spectrum1.Top = 50
'        .Spectrum1.Visible = True
        .imgSmVol.Picture = frmGFX.imgVolume.Picture
        .imgSmVol.Left = 158
        .imgSmVol.Top = 119
        .imgSmVol.Visible = True
        .lblFileInfo.Visible = True
        .lblFileInfo.Left = 64
        .lblFileInfo.Top = 184
        .imgLayout.Picture = frmGFX.imgSmWindow.Picture
        .imgSmBack.Picture = frmGFX.imgSmBack1.Picture
        .imgSmBack.Left = 6
        .imgSmBack.Top = 121
        .imgSmNext.Picture = frmGFX.imgSmNext1.Picture
        .imgSmNext.Left = 15
        .imgSmNext.Top = 171
        .imgSmPause.Picture = frmGFX.imgSmPause1.Picture
        .imgSmPause.Left = 3
        .imgSmPause.Top = 144
        .imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture
        .imgSmPlay.Left = 32
        .imgSmPlay.Top = 145
        .imgSmEject.Picture = frmGFX.imgSmEject1.Picture
        .imgSmEject.Top = 95
        .imgSmEject.Left = 189
        .imgSmOptions.Picture = frmGFX.imgSmOptions1.Picture
        .imgSmOptions.Left = 166
        .imgSmOptions.Top = 87
        .imgSmOptions.Visible = True
        .imgSmEject.Visible = True
        .imgSmPlay.Visible = True
        .imgSmPause.Visible = True
        .imgSmNext.Visible = True
        .imgSmBack.Visible = True
    
    End With
    FadeIn lInitVis

Case eUtilityWindow
    If lFadeOut = True Then FadeOut
    DoEvents
    frmMain.Caption = "::AUDICA.PLAYLIST::"
    lInterface.iCurrentLayout = eUtilityWindow
    GetWindowSettings frmMain.hwnd
    SetUtilityShape
    With frmMain
        .imgLayout.Top = 0
        .imgLayout.Left = 0
        .imgLayout.Visible = True
        .imgLayout.Picture = frmGFX.imgUtilityWind.Picture
        .imgSmOptions.Visible = False
        .imgSmEject.Visible = False
        .imgSmPlay.Visible = False
        .imgSmPause.Visible = False
        .imgSmNext.Visible = False
        .imgSmBack.Visible = False
        .mnuAudica.Checked = False
        .mnuUtility.Checked = True
        .mnuNexENCODE.Checked = False
    End With
    FadeIn
Case eAboutWindow
    If lFadeOut = True Then FadeOut
    DoEvents
    frmMain.Caption = "nexgen . audica - about"
    lInterface.iCurrentLayout = eAboutWindow
    GetWindowSettings frmMain.hwnd
    SetAboutShape
    With frmMain
        .imgLayout.Picture = frmGFX.imgAbout.Picture
        .imgLayout.Top = 0
        .imgLayout.Left = 0
        .imgLayout.Visible = True
    End With
    FadeIn True
End Select
End Sub

Public Sub FadeOut()
Dim x As Integer, i As Integer
x = 100
For i = 1 To 5
    x = x - 20
    MakeTransparent frmMain.hwnd, x
    DoEvents
Next i
End Sub

Public Sub FadeIn(Optional InitVis As Boolean)
Dim i As Integer, x As Integer
x = 0
If InitVis = True Then frmMain.Visible = True
For i = 1 To 5
    x = x + 20
    MakeTransparent frmMain.hwnd, x
    DoEvents
Next i
MakeOpaque frmMain.hwnd
End Sub

Public Sub PlayMp3(Optional lMp3File As String)
Dim msg As String, i As Integer

If Len(lMp3File) = 0 Then
    msg = PromptFile(Mp3_File)
Else
    msg = lMp3File
End If
If Len(msg) <> 0 Then
    i = OpenFile(msg)
    DoEvents
    Pause 0.2
    Playfile i, Mp3_File
Else
    Exit Sub
End If
End Sub

Public Sub PlayWav(Optional lWavFile As String)
Dim msg As String, i As Integer

If Len(lWavFile) = 0 Then
    msg = PromptFile(Wav_File)
Else
    msg = lWavFile
End If
If Len(msg) <> 0 Then
    OpenFile msg
    DoEvents
    i = FindPlaylistIndex(msg)
    Playfile i, Wav_File
Else
    Exit Sub
End If
End Sub

Public Sub GoNext()
Dim i As Integer, msg As String, x As Integer
frmMain.lblFileInfo.Caption = "Loading ..."
CloseMp3Player
DoEvents
Pause 0.2
If lPlaylist.pCount = 0 Or lPlaylist.pCount = 1 Then Exit Sub
If frmMain.mnuRandomize.Checked = True Then
Rand:
    x = GetRnd(lPlaylist.pCount)
    If Len(lPlaylist.pFiles(x).fFilename) <> 0 Then
        If x <> lPlaylist.pCurrent Then
            lPlaylist.pCurrent = x
            msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
            OpenFile msg
            Playfile lPlaylist.pCurrent, Mp3_File
            Exit Sub
        Else
            GoTo Rand
        End If
    Else
        GoTo Rand
    End If
End If
If lPlaylist.pCurrent = lPlaylist.pCount Then
    lPlaylist.pCurrent = 1
    msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    OpenFile msg
    Playfile lPlaylist.pCurrent, Mp3_File
ElseIf lPlaylist.pCurrent = 0 Then
    lPlaylist.pCurrent = 1
    msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    OpenFile msg
    Playfile lPlaylist.pCurrent, Mp3_File
Else
    lPlaylist.pCurrent = lPlaylist.pCurrent + 1
    msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    OpenFile msg
    Playfile lPlaylist.pCurrent, Mp3_File
End If
End Sub

Public Sub ProcessEvent(lEventType As Integer)

End Sub

Public Sub GoBack()
CloseMp3Player
DoEvents
Pause 0.5
Dim i As Integer, msg As String
If lPlaylist.pCount = 0 Then Exit Sub
If lPlaylist.pCurrent = 1 Then
    lPlaylist.pCurrent = lPlaylist.pCount
    OpenFile lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    Playfile lPlaylist.pCurrent, Mp3_File
ElseIf lPlaylist.pCurrent <> 0 Then
    i = lPlaylist.pCurrent
    lPlaylist.pCurrent = i - 1
    msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    OpenFile msg
    Playfile lPlaylist.pCurrent, Mp3_File
ElseIf lPlaylist.pCurrent = 0 Then
    lPlaylist.pCurrent = 1
    OpenFile lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
    Playfile lPlaylist.pCurrent, Mp3_File
End If
End Sub

Public Function OpenFile(lFilename As String) As Integer
On Local Error Resume Next

Dim msg As String, msg2 As String, i As Integer, msg3 As String, lext As String, lFile As String, x As Integer
If Len(lFilename) <> 0 Then
    lFile = lFilename
    lext = Right(LCase(lFilename), 3)
    msg2 = GetFileTitle(lFile)
    i = FindPlaylistIndex(msg2)
    If i = 0 Then x = AddtoPlaylist(lFilename)
    
    Select Case lext
    Case "mp3"
        With frmMain
            CloseMp3Player
            .ctlMp3Player.Open Trim(lFilename), ""
            DoEvents
            msg = .ctlMp3Player.GetArtist & " - " & .ctlMp3Player.GetTitle
            If msg = " - " Then
                msg3 = GetFileTitle(lFilename)
                lInterface.iStatusText = " " & Left(msg3, Len(msg3) - 4) & " ...   "
            Else
                lInterface.iStatusText = " " & msg & " ...   "
            End If
        End With
    Case "mp2"
        With frmMain
            CloseMp3Player
            .ctlMp3Player.Open Trim(lFilename), ""
            DoEvents
            msg = .ctlMp3Player.GetArtist & " - " & .ctlMp3Player.GetTitle
            If msg = " - " Then
                msg3 = GetFileTitle(lFilename)
                lInterface.iStatusText = " " & Left(msg3, Len(msg3) - 4) & " ...   "
            Else
                lInterface.iStatusText = " " & msg & " ...   "
            End If
        End With
    Case "mp1"
        With frmMain
            CloseMp3Player
            .ctlMp3Player.Open Trim(lFilename), ""
            DoEvents
            msg = .ctlMp3Player.GetArtist & " - " & .ctlMp3Player.GetTitle
            If msg = " - " Then
                msg3 = GetFileTitle(lFilename)
                lInterface.iStatusText = " " & Left(msg3, Len(msg3) - 4) & " ...   "
            Else
                lInterface.iStatusText = " " & msg & " ...   "
            End If
        End With
    Case "wav"
        msg = Left(lFilename, Len(lFilename) - 4)
        If Len(msg) <> 0 Then lInterface.iStatusText = " " & msg & " ...   "
    End Select
    OpenFile = x
End If
End Function

Public Sub Playfile(lIndex As Integer, lFiletype As eFiletypes)
On Local Error Resume Next

Dim msg As String, lFilename As String
lFilename = lPlaylist.pFiles(lIndex).fFilepath & "\" & lPlaylist.pFiles(lIndex).fFilename
If DoesFileExist(lFilename) = False Then
    msg = MsgBox("Audica cannot locate '" & lFilename & "'. Would you like to search for this file yourself?", vbYesNo + vbExclamation)
    If msg = vbYes Then
        PlayMp3
        Exit Sub
    Else
        Exit Sub
    End If
End If

If lIndex = 0 Then
    Exit Sub
Else
    Select Case lFiletype
    Case Mp3_File
        If frmMain.ctlMp3Player.GetHasTag = True Then
            'frmDDE.ddemIRC "10:12( 10Title: 12" & frmMain.ctlMp3Player.GetTitle & "10, Artist: 12" & frmMain.ctlMp3Player.GetArtist & ", 9[14" & frmMain.ctlMp3Player.GetGenreString(frmMain.ctlMp3Player.GetGenre) & "9] 0.11aud15ic14a" & " 12)10:"
        Else
            'frmDDE.ddemIRC "10:12( 10Playing: 12" & lPlaylist.pFiles(lIndex).fFilename & "9] 0.11aud15ic14a" & " 12)10:"
        End If
        lInterface.iStoped = False
        frmMain.tmrFake.Enabled = True
        lPlaylist.pCurrent = FindPlaylistIndex(GetFileTitle(lPlaylist.pFiles(lIndex).fFilename))
        frmMain.imgSmPlay.Picture = frmGFX.imgSmStop1.Picture
        frmMain.ctlMp3Player.Play
        lInterface.iPlaying = True
        frmMain.tmrStreamTitle.Enabled = True
    Case Wav_File
        'frmMain.tmrFake.Enabled = True
        lPlaylist.pCurrent = FindPlaylistIndex(GetFileTitle(lPlaylist.pFiles(lIndex).fFilename))
        frmMain.imgSmPlay.Picture = frmGFX.imgSmStop1.Picture
        sndPlaySound lPlaylist.pFiles(lIndex).fFilepath & "\" & lPlaylist.pFiles(lIndex).fFilename, SND_ASYNC

        lInterface.iPlaying = True
        frmMain.tmrStreamTitle.Enabled = True
    End Select
End If
End Sub

Public Function PromptFolder() As String
On Local Error Resume Next

Dim msg As String
lFileshare.fReturn = ""
With frmFolder
    .Label1.Caption = "Please select a folder"
    .Dir1.Path = CurDir
    .Show 1
End With
msg = lFileshare.fReturn
lFileshare.fReturn = ""
If Len(msg) <> 0 Then
    PromptFolder = msg
End If
End Function

Public Function PlayDirectory(lFiletype As eFiletypes)
Dim msg As String, msg2 As String, i As Integer, lext As String

msg = PromptFolder
If Len(msg) <> 0 Then
    frmDir.Dir1.Path = msg
    For i = 0 To frmDir.File1.ListCount
        msg2 = frmDir.File1.List(i)
        lext = Right(msg2, 3)
        Select Case lFiletype
        Case Mp3_File
            If lext = "mp3" Then AddtoPlaylist msg & "\" & msg2
        Case Wav_File
            If lext = "wav" Then AddtoPlaylist msg & "\" & msg2
        End Select
    Next i
    GoNext
End If
End Function

Public Function AddRecientMedia(lFileTitle As String)
On Local Error Resume Next
Dim i As Integer

i = frmMain.mnuRecient.Count
If Len(lFileTitle) <> 0 Then
'    frmMain.mnuRecient(0).Visible = False
    Load frmMain.mnuRecient(i)
    frmMain.mnuRecient(i).Visible = True
    frmMain.mnuRecient(i).Caption = "::" & UCase(Left(lFileTitle, Len(lFileTitle) - 4)) & "::"
    frmMain.mnuRecient(i).Enabled = True
End If
End Function

Public Function PromptFile(lFiletype As eFiletypes) As String
On Local Error Resume Next
Dim msg As String, msg2 As String

Select Case lFiletype
Case Mp3_File
    msg = OpenDialog(frmMain, "Mp3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*", "Nexgen Audica - Select File ...", ReturnDirectoryPath(Mp3_File))
    If Len(msg) = 0 Then Exit Function
    msg = Left(msg, Len(msg) - 1)
    If Len(msg) <> 0 Then
        msg2 = msg
        DoEvents
        frmMain.ctlMp3Player.Stop
        frmMain.ctlMp3Player.Close
        PromptFile = msg
    End If
Case Wav_File
    msg = OpenDialog(frmMain, "Wave Audio Files (*.wav)|*.wav|All Files (*.*)|*.*", "Nexgen Audica - Select File ...", ReturnDirectoryPath(Wav_File))
    msg = Left(msg, Len(msg) - 1)
    If Len(msg) <> 0 Then
        msg2 = msg
        DoEvents
        frmMain.ctlMp3Player.Stop
        frmMain.ctlMp3Player.Close
        PromptFile = msg
    End If
End Select
End Function

Public Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub GetWindowSettings(lHandle As Long)
On Local Error Resume Next

Dim lWindowPos As RECT, lClientPos As RECT
Dim lBorderWidth As Long, lTopOffset As Long
Dim i As Long

i = GetWindowRect(lHandle, lWindowPos)
i = GetClientRect(lHandle, lClientPos)

lMainWndSettings.wTitleBarHeight = lWindowPos.Bottom - lWindowPos.Top - lClientPos.Bottom - lBorderWidth
lMainWndSettings.wWindowBorder = lWindowPos.Right - lWindowPos.Left - lClientPos.Right - 2
End Sub

Public Sub FormDrag(lFormname As Form)
ReleaseCapture
Call SendMessage(lFormname.hwnd, &HA1, 2, 0&)
End Sub

Public Sub SetPlayerShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, tmp As Long
Dim x As Long, Y As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateEllipticRgn(x + 14, Y - 3, x + 172, Y + 152)
rgn1 = CreateEllipticRgn(x - 1.2, Y + 68, x + 190, Y + 234)
rgn2 = CreateEllipticRgn(x + 72, Y + 71, x + 237, Y + 227)
rgn3 = CreateEllipticRgn(x + 26, Y + 145, x + 161 + 23, Y + 153 + 150)

tmp = CombineRgn(rgn, rgn, rgn1, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn2, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn3, RGN_OR)

frmMain.Width = 3700
frmMain.Height = 5300
tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub

Public Sub SetUtilityShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, tmp As Long
Dim x As Long, Y As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateRoundRectRgn(x + 38, Y + 249, x + 288, Y + 268, 10, 10)
rgn1 = CreateRoundRectRgn(x + 286, Y - 3, x + 324, Y + 231, 40, 40)
rgn2 = CreateRectRgn(x + 286, Y + 192, x + 30 + 286, Y + 57 + 192)
rgn3 = CreateEllipticRgn(x + 286, Y + 228, x + 288 + 50, Y + 41 + 232)
rgn4 = CreateRoundRectRgn(x + 1, Y - 3, x + 300, Y + 34, 30, 30)
rgn5 = CreateRoundRectRgn(x + 35, Y + -100, x + 287, Y + 16, 20, 20)
rgn6 = CreateRectRgn(x + 35, Y + 20, x + 249 + 40, Y + 209 + 40)
rgn7 = CreateRoundRectRgn(x - 1, Y + 20, x + 70, Y + 250, 20, 20)

tmp = CombineRgn(rgn, rgn, rgn1, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn2, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn3, RGN_DIFF)
tmp = CombineRgn(rgn, rgn, rgn4, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn5, RGN_DIFF)
tmp = CombineRgn(rgn, rgn, rgn6, RGN_OR)
tmp = CombineRgn(rgn, rgn, rgn7, RGN_OR)

frmMain.Width = 5000
frmMain.Height = 4800

tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub

Public Sub SetAboutShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, tmp As Long
Dim x As Long, Y As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight
rgn = CreateRectRgn(x - 1, Y - 2, x + 200, Y + 229)
frmMain.Width = 3200
frmMain.Height = 4200

tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub
