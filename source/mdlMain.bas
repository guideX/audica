Attribute VB_Name = "mdlMain"
'Attribute VB_Name = "OpenFile32"
Option Explicit

Enum eSliderTypes
    Slider_Volume = 1
    Slider_Position = 2
    Slider_Balance = 3
End Enum
Enum eFiletypes
    Other_File = 0
    Mp3_File = 1
    Wav_File = 2
    Midi_File = 3
    Wma_File = 4
    Wmv_File = 5
    Mpeg_File = 6
End Enum
Enum eGFXFlash
    Audica_Logo = 1
End Enum
Private Type gFiles
    fFilename As String
    fFilepath As String
    fFiletype As eFiletypes
End Type
Private Type gPlaylist
    pFilename As String
    pFiles(200) As gFiles
    pCount As Integer
    pCurrent As Integer
    pVolumeButton As Boolean
End Type
Enum eCurrentLayout
    eSmWindow = 1
    eUtilityWindow = 2
    eAboutWindow = 3
    eNexENCODEWindow = 4
End Enum

Private Type gInterface
    iCurrentLayout As eCurrentLayout
    iCurrentFlash As eGFXFlash
    iFlashLoop As Integer
    iSliderType As eSliderTypes
    iStatusText As String
    iStatusDisplay As String
    iPlaying As Boolean
    iStoped As Boolean
    iOsSelected As Integer
    iPauseLayout As Boolean
End Type

Private Type gDirectory
    dPath As String
    dFiletype As eFiletypes
End Type
Private Type gSettings
    sDirectorys(6) As gDirectory
    sOutputDevice As Integer
End Type

Public lSettings As gSettings
Public lInterface As gInterface
Public lPlaylist As gPlaylist


Private Type gWindowPos
    wTitleBarHeight As Integer
    wWindowBorder As Integer
End Type

Global lMainWndSettings As gWindowPos


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
    Public Const OFN_READONLY = &H1
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_SHOWHELP = &H10
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
    Public Const OFN_EXPLORER = &H80000 ' new look commdlg
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
    Public Const OFN_SHAREFALLTHROUGH = 2
    Public Const OFN_SHARENOWARN = 1
    Public Const OFN_SHAREWARN = 0
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const RGN_MIN = RGN_AND
Public Const RGN_MAX = RGN_COPY
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Enum sndConst
    SND_ASYNC = &H1 ' play asynchronously
    SND_LOOP = &H8 ' loop the sound until Next sndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points To a memory file
    SND_NODEFAULT = &H2 ' silence Not default, If sound not found
    SND_NOSTOP = &H10 ' don't stop any currently playing sound
    SND_SYNC = &H0 ' play synchronously (default), halts prog use till done playing
End Enum

Private m_objGenres              As clsGenres
Private Type ID3V1Tag
   Album       As String * 30
   Artist      As String * 30
   Comment     As String * 30
   Genre       As Byte
   Identifier  As String * 3
   Title       As String * 30
   Year        As String * 4
End Type
Public Type Id3Tag
   Album       As String * 30
   Artist      As String * 30
   Comment     As String * 30
   Genre       As String * 30
   Identifier  As String * 3
   Title       As String * 30
   Year        As String * 4
End Type

Public Function ReadFile(lFile As String) As String
On Local Error Resume Next

Dim o As Integer, msg As String
o = FreeFile
Open lFile For Input As #o
    'ReadFile = Input(LOF(o), o)
    msg = StrConv(InputB(LOF(o), o), vbUnicode)
    ReadFile = Left(msg, Len(msg) - 2)
Close #o
End Function

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error Resume Next

Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function ReturnDirectoryPath(lFiletype As eFiletypes)
On Local Error Resume Next
Dim i As Integer

If lFiletype <> 0 Then
    For i = 1 To 6
        If lSettings.sDirectorys(i).dFiletype = lFiletype Then
            ReturnDirectoryPath = i
        End If
    Next i
End If
End Function

Public Function ParseString(lWhole As String, lStart As String, lEnd As String)
'On Local Error GoTo ErrHandler

Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
ParseString = Left(Str1, Len(Str1) - Len(Str2) - 1)

ErrHandler:
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim msg As Long
'On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
Dim msg As Long
On Error Resume Next
msg = GetWindowLong(hwnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function


Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
Dim ofn As OPENFILENAME, A As Long

ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hwnd
ofn.hInstance = App.hInstance
If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
For A = 1 To Len(Filter)
    If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
Next
ofn.lpstrFilter = Filter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = InitDir
ofn.lpstrTitle = Title
ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
A = GetSaveFileName(ofn)
If (A) Then
    SaveDialog = Trim$(ofn.lpstrFile)
Else
    SaveDialog = ""
End If
End Function

Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
Dim ofn As OPENFILENAME, A As Long
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hwnd
ofn.hInstance = App.hInstance
If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
For A = 1 To Len(Filter)
    If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
Next
ofn.lpstrFilter = Filter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = InitDir
ofn.lpstrTitle = Title
ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
A = GetOpenFileName(ofn)
    
If (A) Then
    OpenDialog = Trim$(ofn.lpstrFile)
Else
    OpenDialog = ""
End If
End Function



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
frmMain.Mp3OCX1.Stop
'frmMain.ctlMp3Player.Stop
'frmMain.ctlMp3Player.Close
'frmMain.ctlMp3Player.SetOutDevice lSettings.sOutputDevice
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
Dim i As Integer, X As Integer, msg As String

lPlaylist.pFilename = GetSetting(App.Title, "Playlist", "Filename", App.Path & "\" & "audica.m3u")
X = GetSetting(App.Title, "Playlist", "Count", 0)
If X <> 0 Then
    For i = 1 To X
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
Dim X As Long, Y As Long

X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateEllipticRgn(0, 0, frmMain.Width, frmMain.Height) ' whole image
rgn1 = CreateEllipticRgn(X + 147, Y + 90, X + 326, Y + 267) 'big crl in back (transparency)
rgn2 = CreateEllipticRgn(X + 104, Y + 46, X + 367, Y + 310) ' big crl in back
rgn3 = CreateEllipticRgn(X + 48, Y + 74, X + 257, Y + 287) 'left crl
rgn4 = CreateEllipticRgn(X + 65, Y + 92, X + 241, Y + 268) 'left crl (transperancy)
rgn5 = CreateEllipticRgn(X + 212, Y + 72, X + 422, Y + 286) 'right crl
rgn6 = CreateEllipticRgn(X + 230, Y + 91, X + 404, Y + 268) 'right crl (transparency)
rgn7 = CreateRoundRectRgn(X + 39, Y + 120, X + 429, Y + 237, 110, 110)  'pill

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
        .lblFileInfo.Visible = True
        .imgSlider.Visible = True
        .imgSmVol.Visible = True
        .Mp3OCX1.Visible = True
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
        .lblFileInfo.Visible = False
        .imgSlider.Visible = False
        .imgSmVol.Visible = False
        .Mp3OCX1.Visible = False
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
        .Mp3OCX1.Visible = False
        .imgLayout.Picture = frmGFX.imgAbout.Picture
        .imgLayout.Top = 0
        .imgLayout.Left = 0
        .imgLayout.Visible = True
    End With
    FadeIn True
End Select
End Sub

Public Sub FadeOut()
Dim X As Integer, i As Integer
X = 100
For i = 1 To 5
    X = X - 20
    MakeTransparent frmMain.hwnd, X
    DoEvents
Next i
End Sub

Public Sub FadeIn(Optional InitVis As Boolean)
Dim i As Integer, X As Integer
X = 0
If InitVis = True Then frmMain.Visible = True
For i = 1 To 5
    X = X + 20
    MakeTransparent frmMain.hwnd, X
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
        'If frmMain.ctlMp3Player.GetHasTag = True Then
            'frmDDE.ddemIRC "10:12( 10Title: 12" & frmMain.ctlMp3Player.GetTitle & "10, Artist: 12" & frmMain.ctlMp3Player.GetArtist & ", 9[14" & frmMain.ctlMp3Player.GetGenreString(frmMain.ctlMp3Player.GetGenre) & "9] 0.11aud15ic14a" & " 12)10:"
        'Else
            'frmDDE.ddemIRC "10:12( 10Playing: 12" & lPlaylist.pFiles(lIndex).fFilename & "9] 0.11aud15ic14a" & " 12)10:"
        'End If
      
    'Case Wav_File
        'lPlaylist.pCurrent = FindPlaylistIndex(GetFileTitle(lPlaylist.pFiles(lIndex).fFilename))
        'frmMain.imgSmPlay.Picture = frmGFX.imgSmStop1.Picture
        'sndPlaySound lPlaylist.pFiles(lIndex).fFilepath & "\" & lPlaylist.pFiles(lIndex).fFilename, SND_ASYNC
        'lInterface.iPlaying = True
        'frmMain.tmrStreamTitle.Enabled = True
    End Select
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
Dim i As Integer, msg As String, X As Integer
frmMain.lblFileInfo.Caption = "Loading ..."
CloseMp3Player
DoEvents
Pause 0.2
If lPlaylist.pCount = 0 Or lPlaylist.pCount = 1 Then Exit Sub
If frmMain.mnuRandomize.Checked = True Then
Rand:
    X = GetRnd(lPlaylist.pCount)
    If Len(lPlaylist.pFiles(X).fFilename) <> 0 Then
        If X <> lPlaylist.pCurrent Then
            lPlaylist.pCurrent = X
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
Dim msg As String, msg2 As String, i As Integer, msg3 As String, lext As String, lFile As String, X As Integer
If Len(lFilename) <> 0 Then
    lFile = lFilename
    lext = Right(LCase(lFilename), 3)
    msg2 = GetFileTitle(lFile)
    i = FindPlaylistIndex(msg2)
    If i = 0 Then X = AddtoPlaylist(lFilename)
    Select Case lext
    Case "mp3"
        With frmMain
            ms_InitialiseGenres
            Dim t As Id3Tag
            t = ms_ShowID3V1Tag(lFilename)
            Dim msg20 As String
            msg20 = Right(t.Title, 1)
            lInterface.iStatusText = Replace(Trim(t.Title), Chr(0), "") & " by " & Replace(Trim(t.Artist), Chr(0), "") & " album " & Replace(Trim(t.Album), Chr(0), "") & " " & Replace(Trim(t.Genre), Chr(0), "")
            lInterface.iStoped = False
            frmMain.tmrFake.Enabled = True
            lPlaylist.pCurrent = FindPlaylistIndex(GetFileTitle(lFilename))
            frmMain.imgSmPlay.Picture = frmGFX.imgSmStop1.Picture
            lInterface.iPlaying = True
            frmMain.tmrStreamTitle.Enabled = True
            .Mp3OCX1.Play lFilename
            frmMain.tmrStreamTitle.Enabled = True
        End With
    End Select
    OpenFile = X
End If
End Function

Function AlphaNumericOnly(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult
End Function
Public Function PromptFolder() As String
On Local Error Resume Next
Dim msg As String
With frmFolder
    .Label1.Caption = "Please select a folder"
    .Dir1.Path = CurDir
    .Show 1
    msg = .Dir1.Path
End With
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
    frmMain.mnuRecient(0).Visible = False
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
        frmMain.Mp3OCX1.Stop
        'frmMain.ctlMp3Player.Stop
        'frmMain.ctlMp3Player.Close
        PromptFile = msg
    End If
Case Wav_File
    msg = OpenDialog(frmMain, "Wave Audio Files (*.wav)|*.wav|All Files (*.*)|*.*", "Nexgen Audica - Select File ...", ReturnDirectoryPath(Wav_File))
    msg = Left(msg, Len(msg) - 1)
    If Len(msg) <> 0 Then
        msg2 = msg
        DoEvents
        frmMain.Mp3OCX1.Stop
        'frmMain.ctlMp3Player.Stop
        'frmMain.ctlMp3Player.Close
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
Dim X As Long, Y As Long

X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateEllipticRgn(X + 14, Y - 3, X + 172, Y + 152)
rgn1 = CreateEllipticRgn(X - 1.2, Y + 68, X + 190, Y + 234)
rgn2 = CreateEllipticRgn(X + 72, Y + 71, X + 237, Y + 227)
rgn3 = CreateEllipticRgn(X + 26, Y + 145, X + 161 + 23, Y + 153 + 150)

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
Dim X As Long, Y As Long

X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateRoundRectRgn(X + 38, Y + 249, X + 288, Y + 268, 10, 10)
rgn1 = CreateRoundRectRgn(X + 286, Y - 3, X + 324, Y + 231, 40, 40)
rgn2 = CreateRectRgn(X + 286, Y + 192, X + 30 + 286, Y + 57 + 192)
rgn3 = CreateEllipticRgn(X + 286, Y + 228, X + 288 + 50, Y + 41 + 232)
rgn4 = CreateRoundRectRgn(X + 1, Y - 3, X + 300, Y + 34, 30, 30)
rgn5 = CreateRoundRectRgn(X + 35, Y + -100, X + 287, Y + 16, 20, 20)
rgn6 = CreateRectRgn(X + 35, Y + 20, X + 249 + 40, Y + 209 + 40)
rgn7 = CreateRoundRectRgn(X - 1, Y + 20, X + 70, Y + 250, 20, 20)

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
Dim X As Long, Y As Long

X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight
rgn = CreateRectRgn(X - 1, Y - 2, X + 200, Y + 229)
frmMain.Width = 3200
frmMain.Height = 4200

tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub


Public Function ms_ShowID3V1Tag(sFileName As String) As Id3Tag
   On Local Error GoTo ErrHandler
   Const ID3V1TagSize   As Integer = 127
   Dim result As Id3Tag
   Dim t                As ID3V1Tag
   Dim lFileHandle      As Long
   Dim lll              As Long
   Dim sGenre           As String
   lFileHandle = FreeFile()
   Open sFileName For Binary As #lFileHandle
   lll = LOF(lFileHandle) 'Get the length of mp3 file
   Get #lFileHandle, lll - ID3V1TagSize, t.Identifier
   With t
      If .Identifier = "TAG" Then
         Get #lFileHandle, , .Title   '30 chars
         Get #lFileHandle, , .Artist  '30 chars
         Get #lFileHandle, , .Album   '30 chars
         Get #lFileHandle, , .Year    '4 chars
         Get #lFileHandle, , .Comment '30 chars
         Get #lFileHandle, , .Genre   '1 byte (i think)
         result.Album = Trim(.Album)
         result.Artist = Trim(.Artist)
         result.Comment = Trim(.Comment)
         result.Identifier = Trim(.Identifier)
         result.Title = Trim(.Title)
         result.Year = Trim(.Year)
         'sGenre = CStr(.Genre)
         'If m_objGenres.Exists(CStr(sGenre)) Then
            'Dim g As clsGenre
            'g = m_objGenres.Item(CStr(sGenre))
            'result.Genre = g.Description
         'End If
         'ms_ShowID3V1Tag = result
      End If
   End With
   ms_ShowID3V1Tag = result
   Close
Exit Function
ErrHandler:
   MsgBox "Error: " & Err.Description
End Function

Public Sub ms_InitialiseGenres()
   Dim objXMLDocument   As Object 'MSXML2.DOMDocument
   Dim objNodeList      As Object 'MSXML2.IXMLDOMNodeList
   Dim objRoot          As Object 'MSXML2.IXMLDOMElement
   Dim objNode          As Object 'MSXML2.IXMLDOMNode
   Dim objChild         As Object 'MSXML2.IXMLDOMNode
   Dim sIdentifier      As String
   Dim sGenre           As String
   Dim XML_FILE       As String
    XML_FILE = App.Path & "\Genres.xml"
   Set objXMLDocument = CreateObject("Microsoft.XMLDOM") '= New MSXML2.DOMDocument
   With objXMLDocument
      .async = False
      If .Load(XML_FILE) Then
         Set objRoot = .documentElement()
         For Each objNode In objRoot.childNodes
            sGenre = vbNullString
            sIdentifier = vbNullString
            For Each objChild In objNode.childNodes
               If objChild.nodeName = "id" Then
                  sIdentifier = objChild.Text
               ElseIf objChild.nodeName = "Description" Then
                  sGenre = objChild.Text
               End If
            Next
            If sGenre <> vbNullString Then
               If sIdentifier <> vbNullString Then
                  m_objGenres.Add sGenre, sIdentifier
               End If
            End If
         Next
         'ms_LoadGenreComboBox
      Else
         MsgBox "Error loading xml file: " & XML_FILE & vbCrLf & _
            "Check if the path to the file is correct", _
            vbExclamation, "Cannot Find XML File"
      End If
   End With
End Sub
