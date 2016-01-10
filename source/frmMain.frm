VERSION 5.00
Object = "{EEB96F74-14D2-11D3-A1BB-B6FC7F000000}#1.0#0"; "Mp3OCX.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audica"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picActiveControls 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Timer tmrStreamTitle 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   100
         Top             =   100
      End
      Begin VB.Timer tmrFake 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   600
         Top             =   120
      End
   End
   Begin MP3OCXLib.Mp3OCX Mp3OCX1 
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   1931
      _StockProps     =   161
      BackColor       =   0
   End
   Begin VB.Image imgSmPlay 
      Height          =   375
      Left            =   2160
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgSlider 
      Height          =   495
      Left            =   3480
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgSmVol 
      Height          =   495
      Left            =   3600
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFileInfo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   900
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgSmEject 
      Height          =   465
      Left            =   3600
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgSmOptions 
      Height          =   465
      Left            =   3600
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgSmNext 
      Height          =   435
      Left            =   3600
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image imgSmPlayOld 
      Height          =   435
      Left            =   6000
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgSmPause 
      Height          =   435
      Left            =   3600
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image imgSmBack 
      Height          =   405
      Left            =   3600
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image imgLayoutOld 
      Height          =   375
      Left            =   6000
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgLayout 
      Height          =   1695
      Left            =   1080
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Hidden"
      Begin VB.Menu mnuAudica 
         Caption         =   "::NEXGEN><AUDICA::"
         Checked         =   -1  'True
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuNexENCODE 
         Caption         =   "::NEXENCODE::"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNexMEDIA 
         Caption         =   "::NEXMEDIA::"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNexSynth 
         Caption         =   "::NEXSYNTH::"
         Enabled         =   0   'False
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUtility 
         Caption         =   "::PLAYLIST::"
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep98372973892 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "::SETTINGS::"
         Visible         =   0   'False
         Begin VB.Menu mnuOS 
            Caption         =   "::OPERATING><SYSTEM::"
         End
         Begin VB.Menu mnuOutputDevice 
            Caption         =   "::OUTPUT><DEVICE::"
         End
      End
      Begin VB.Menu mnuSep93782973 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoBack 
         Caption         =   "::GO.BACK::"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "::PLAY::"
         Begin VB.Menu mnuMP3File 
            Caption         =   "::MP3.FILE(S)::"
            Begin VB.Menu mnuPlayMp3File 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuOpenMp3Directory 
               Caption         =   "::PLAY.MP3.DIRECTORY::"
            End
            Begin VB.Menu mnuPlaymp3playlist 
               Caption         =   "::PLAY.PLAYLIST::"
            End
            Begin VB.Menu mnuSep9378937 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDecodeMp3 
               Caption         =   "::DECODE.MP3::"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuDecodeDirectory 
               Caption         =   "::DECODE.DIRECTORY::"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuDecodePlaylist 
               Caption         =   "::DECODE.PLAYLIST::"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu mnuWavFiles 
            Caption         =   "::WAV.FILE(S)::"
            Visible         =   0   'False
            Begin VB.Menu mnuPlayFile 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuPlayWavDirectory 
               Caption         =   "::PLAY.WAV.DIRECTORY::"
            End
            Begin VB.Menu mnuPlayWavPlaylist 
               Caption         =   "::PLAY.PLAYLIST::"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu mnuMIDIFILES 
            Caption         =   "::MIDI.FILES::"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuPlayMidiFIle 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuPlayMidiDirectory 
               Caption         =   "::PLAY.MIDI.DIRECTORY::"
            End
            Begin VB.Menu mnuPlayPlaylist 
               Caption         =   "::PLAY.PLAYLIST::"
            End
         End
         Begin VB.Menu mnuWmaFiles 
            Caption         =   "::WMA.FILES::"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuPlayWmaFile 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuPlayWmaDirectory 
               Caption         =   "::PLAY.WMA.DIRECTORY::"
            End
            Begin VB.Menu mnuPlaywmaPlaylist 
               Caption         =   "::PLAY.PLAYLIST::"
            End
         End
         Begin VB.Menu mnuWMVFILES 
            Caption         =   "::WMV.FILES::"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuPlayFiles 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuPlayWMVDirectory 
               Caption         =   "::PLAY.WMV.DIRECTORY::"
            End
            Begin VB.Menu mnuPlayWMVPlaylist 
               Caption         =   "::PLAY.PLAYLIST::"
            End
         End
         Begin VB.Menu mnuMpegFiles 
            Caption         =   "::MPEG.FILES::"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuPlayMpegFile 
               Caption         =   "::PLAY.FILE::"
            End
            Begin VB.Menu mnuplayMpegDirectory 
               Caption         =   "::PLAY.MPEG.DIRECTORY::"
            End
            Begin VB.Menu mnuPlayPlaylistMpeg 
               Caption         =   "::PLAY.PLAYLIST::"
            End
         End
         Begin VB.Menu mnuSep3872937 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPlayerSettings 
            Caption         =   "::PLAYER.SETTINGS::"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuPause 
         Caption         =   "::PAUSE::"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "::STOP::"
      End
      Begin VB.Menu mnuGoNext 
         Caption         =   "::GO.NEXT::"
      End
      Begin VB.Menu mnuSep9389723 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "::PLAYLIST::"
         Begin VB.Menu mnuRecientMedia 
            Caption         =   "::FILES::"
            Begin VB.Menu mnuRecient 
               Caption         =   "::EMPTY::"
               Index           =   0
            End
         End
         Begin VB.Menu mnuSep0938297392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLoadPlaylist 
            Caption         =   "::LOAD.PLAYLIST::"
         End
         Begin VB.Menu mnuAppendToPlaylist 
            Caption         =   "::APEND.TO.PLAYLIST::"
         End
         Begin VB.Menu mnuClearPlaylist 
            Caption         =   "::CLEAR.PLAYLIST::"
         End
         Begin VB.Menu mnuPlaylistEditor 
            Caption         =   "::PLAYLIST.EDITOR::"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSep3987932 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSavePlaylist 
            Caption         =   "::SAVE.PLAYLIST::"
         End
         Begin VB.Menu mnuSavePlaylistAs 
            Caption         =   "::SAVE.PLAYLIST.AS::"
         End
         Begin VB.Menu mnuSep937923 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRandomize 
            Caption         =   "::RANDOMIZE::"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuFileTransfer 
         Caption         =   "::FILE.TRANSFER::"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuShowFileTransfers 
            Caption         =   "::WINDOW::"
         End
         Begin VB.Menu mnuSep938292379273 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchForMedia 
            Caption         =   "::NEWS::"
         End
         Begin VB.Menu mnuCurrentFileTransfers 
            Caption         =   "::SEARCH::"
         End
         Begin VB.Menu mnuTrade 
            Caption         =   "::TRADE::"
         End
         Begin VB.Menu mnuDownloads 
            Caption         =   "::DOWNLOADS::"
         End
         Begin VB.Menu mnuUploads 
            Caption         =   "::UPLOADS::"
         End
         Begin VB.Menu mnuBrowse 
            Caption         =   "::BROWSE::"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "::OPTIONS::"
         End
         Begin VB.Menu mnuSep3987297392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "::DISABLE::"
         End
      End
      Begin VB.Menu mnuWAVToMp3 
         Caption         =   "::WAV.TO.MP3::"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuENCODEOneFile 
            Caption         =   "::ENCODE.ONE.FILE::"
         End
         Begin VB.Menu mnuEncodeSeveralFiles 
            Caption         =   "::ENCODE.SEVERAL.FILES::"
         End
         Begin VB.Menu mnuSep397297392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEncoderSettings 
            Caption         =   "::ENCODER.SETTINGS::"
         End
      End
      Begin VB.Menu mnuCDAToWav 
         Caption         =   "::CDA.TO.WAV::"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuRipOneTrack 
            Caption         =   "::RIP.ONE.TRACK::"
         End
         Begin VB.Menu mnuRipSeveralTracks 
            Caption         =   "::RIP.SEVERAL.TRACKS::"
         End
      End
      Begin VB.Menu mnuWavEffects 
         Caption         =   "::WAV.EFFECTS::"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuShowEffects 
            Caption         =   "::SHOW.EFFECTS::"
         End
      End
      Begin VB.Menu mnuSep9379273 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "::HIDE::"
      End
      Begin VB.Menu mnuPower 
         Caption         =   "::POWER::"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlMp3Player_ThreadEnded()
tmrStreamTitle.Enabled = False
imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture
tmrFake.Enabled = False
lInterface.iPlaying = False
If lInterface.iStoped = False And lPlaylist.pCount <> 0 And lPlaylist.pCount <> 1 Then GoNext
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim i As Integer, msg As String, X As Integer
If (Command = "mirc") Then
    ' Load program into mIRC
    Dim mircHwnd As Long
    Dim s As String
    mircHwnd = FindWindow(s, "mIRC Control Panel")
    SetParent Me.hwnd, mircHwnd
End If
SetInterface eAboutWindow, False, True
Me.BackColor = vbBlack
Mp3OCX1.OscilloType = otSpectrum
Mp3OCX1.BackColor = 0
Mp3OCX1.RightChanColor = &H800000
Mp3OCX1.LeftChanColor = &HFF0000
Mp3OCX1.Bands = 14
LoadSettings
'Pause 0.5
imgSlider.Left = 186
imgSlider.Visible = True
Dim lfname As String
lfname = Command
If Len(lfname) <> 0 Then
    PlayMp3 lfname
End If
FadeOut
InitDisplay
AppendToPlaylist lSettings.sLastPlaylist
Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
CloseMp3Player
DoEvents
SaveSettings
FadeOut
End
End Sub

Private Sub imgLayout_DblClick()
'If lInterface.iCurrentLayout = eSmWindow Then
    'SetInterface eUtilityWindow, True
'Else
    'SetInterface eSmWindow, True
'End If
End Sub

Private Sub imgLayout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    FormDrag Me
Else
    PopupMenu mnuMain
End If
End Sub

Private Sub imgSmBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgSmBack.Picture = frmGFX.imgSmBack2.Picture
End Sub

Private Sub imgSmBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    GoBack
    imgSmBack.Picture = frmGFX.imgSmBack1.Picture
End If
End Sub

Private Sub imgSmEject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgSmEject.Picture = frmGFX.imgSmEject2.Picture
End Sub

Private Sub imgSmEject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    CloseMp3Player
    PlayMp3
    imgSmEject.Picture = frmGFX.imgSmEject1.Picture
End If
End Sub

Private Sub imgSmNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgSmNext.Picture = frmGFX.imgSmNext2.Picture
End Sub

Private Sub imgSmNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    GoNext
    imgSmNext.Picture = frmGFX.imgSmNext1.Picture
End If
End Sub

Private Sub imgSmOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgSmOptions.Picture = frmGFX.imgSmOptions2.Picture
End Sub

Private Sub imgSmOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'SetInterface eUtilityWindow, True
    imgSmOptions.Picture = frmGFX.imgSmOptions1.Picture
End If
End Sub

Private Sub imgSmPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgSmPause.Picture = frmGFX.imgSmPause2.Picture
End Sub

Private Sub imgSmPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    imgSmPause.Picture = frmGFX.imgSmPause1.Picture
    If lInterface.iPlaying = True Then
        If tmrFake.Enabled = False Then
            tmrFake.Enabled = True
        Else
            tmrFake.Enabled = False
        End If
        Mp3OCX1.Pause
    End If
End If
End Sub

Private Sub imgSmPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture Then
        imgSmPlay.Picture = frmGFX.imgSmPlay2.Picture

    End If
    If imgSmPlay.Picture = frmGFX.imgSmStop1.Picture Then
        imgSmPlay.Picture = frmGFX.imgSmStop2.Picture
    End If
End If
End Sub

Private Sub imgSmPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    imgSmPause.Enabled = True
    If imgSmPlay.Picture = frmGFX.imgSmPlay2.Picture Then
        imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture
        If lblFileInfo.Caption = "" Then
            If lPlaylist.pCount <> 0 Then
                PlayMp3 lPlaylist.pFiles(1).fFilepath & "/" & lPlaylist.pFiles(1).fFilename
            Else
                PlayMp3
            End If
        Else
            Playfile lPlaylist.pCurrent, Mp3_File
        End If
    End If
    If imgSmPlay.Picture = frmGFX.imgSmStop2.Picture Then
        lInterface.iStoped = True
        CloseMp3Player
        imgSmPlay.Picture = frmGFX.imgSmPlay1.Picture
    End If
End If
End Sub

Private Sub imgSmVol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSlider.Left = imgSmVol.Left + 20
imgSlider.Top = imgSmVol.Top + 20
imgSlider.Visible = True
End Sub

Private Sub imgSmVol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    imgSlider.Top = Y / 14 + 110
    If Y < 500 Then
        imgSlider.Left = imgSlider.Left + 1
    Else
        imgSlider.Left = imgSlider.Left - 1
    End If
End If
End Sub

Private Sub lblFileInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub mnuAppendToPlaylist_Click()
AppendToPlaylist
End Sub

Private Sub mnuAudica_Click()
SetInterface eSmWindow, True, False
End Sub

Private Sub mnuClearPlaylist_Click()
ClearPlaylist
End Sub

Private Sub mnuGoNext_Click()
GoNext
End Sub

Private Sub mnuLoadPlaylist_Click()
ClearPlaylist
DoEvents
AppendToPlaylist
End Sub

Private Sub mnuMinimize_Click()
WindowState = vbMinimized
End Sub

Private Sub mnuOpenMp3Directory_Click()
PlayDirectory Mp3_File
End Sub

Private Sub mnuPause_Click()
Mp3OCX1.Pause
End Sub

Private Sub mnuPlayFile_Click()
sndPlaySound OpenDialog(frmMain, "Wav Files (*.wav)|*.wav|All Files (*.*)|*.*", "Open", CurDir), SND_ASYNC
End Sub

Private Sub mnuPlayMp3File_Click()
CloseMp3Player
PlayMp3
End Sub

Private Sub mnuPlaymp3playlist_Click()
ClearPlaylist
AppendToPlaylist
End Sub

Private Sub mnuPlayWavDirectory_Click()
PlayDirectory Wav_File
End Sub

Private Sub mnuPower_Click()
Unload frmMain
End Sub

Private Sub mnuRandomize_Click()
If mnuRandomize.Checked = False Then
    mnuRandomize.Checked = True
Else
    mnuRandomize.Checked = False
End If
End Sub

Private Sub mnuRecient_Click(Index As Integer)
Dim i As Integer, msg As String
CloseMp3Player
'PlayMp3
msg = Left(mnuRecient(Index).Caption, Len(mnuRecient(Index).Caption) - 2)
msg = LCase(Right(msg, Len(msg) - 2))
i = FindPlaylistIndex(msg)
OpenFile lPlaylist.pFiles(i).fFilepath & lPlaylist.pFiles(i).fFilename
Playfile i, lPlaylist.pFiles(i).fFiletype
End Sub

Private Sub mnuSavePlaylist_Click()
SavePlaylist lPlaylist.pFilename
End Sub

Private Sub mnuSavePlaylistAs_Click()
SavePlaylist
End Sub

Private Sub mnuUtility_Click()
SetInterface eUtilityWindow, True, False
End Sub

Private Sub tmrStreamTitle_Timer()
On Local Error Resume Next
lInterface.iStatusText = Right(lInterface.iStatusText, Len(lInterface.iStatusText) - 1)
lInterface.iStatusText = lInterface.iStatusText & Left(lInterface.iStatusText, 1)
lblFileInfo.Caption = lInterface.iStatusText
End Sub
