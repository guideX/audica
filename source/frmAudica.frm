VERSION 5.00
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "Mp3play.ocx"
Object = "{4DD7D078-F270-44BA-A3D5-6E2AF5B83F89}#1.0#0"; "Spectre.ocx"
Begin VB.Form frmAudica 
   Caption         =   "Audica"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAudica.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAudica.frx":08CA
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetshape 
      Caption         =   "XP Setshape"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer tmrFake 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   4200
   End
   Begin VB.ListBox lstPlaylist 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFC0FF&
      Height          =   645
      IntegralHeight  =   0   'False
      Left            =   3720
      TabIndex        =   1
      Top             =   3405
      Width           =   2655
   End
   Begin VB.PictureBox picGFXContainer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   5280
      Width           =   6615
      Begin MPEGPLAYLib.Mp3Play Mp3Player 
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   0
      End
      Begin VB.Timer tmrGFXFlash 
         Interval        =   20
         Left            =   840
         Top             =   600
      End
      Begin VB.Image imgEject2 
         Height          =   225
         Left            =   1680
         Picture         =   "frmAudica.frx":F410
         Top             =   240
         Width           =   465
      End
      Begin VB.Image imgEject1 
         Height          =   225
         Left            =   1680
         Picture         =   "frmAudica.frx":F9F2
         Top             =   0
         Width           =   465
      End
      Begin VB.Image imgPause2 
         Height          =   225
         Left            =   1440
         Picture         =   "frmAudica.frx":FFD4
         Top             =   240
         Width           =   285
      End
      Begin VB.Image imgPause1 
         Height          =   225
         Left            =   1440
         Picture         =   "frmAudica.frx":1039A
         Top             =   0
         Width           =   285
      End
      Begin VB.Image imgMinimize2 
         Height          =   180
         Left            =   5880
         Picture         =   "frmAudica.frx":10760
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgMinimize1 
         Height          =   180
         Left            =   5880
         Picture         =   "frmAudica.frx":10C52
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgExit2 
         Height          =   180
         Left            =   5880
         Picture         =   "frmAudica.frx":11144
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgExit1 
         Height          =   180
         Left            =   5880
         Picture         =   "frmAudica.frx":11636
         Top             =   720
         Width           =   495
      End
      Begin VB.Image imgForeward2 
         Height          =   240
         Left            =   960
         Picture         =   "frmAudica.frx":11B28
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgForeward1 
         Height          =   240
         Left            =   960
         Picture         =   "frmAudica.frx":1216A
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgPlay2 
         Height          =   270
         Left            =   720
         Picture         =   "frmAudica.frx":127AC
         Top             =   240
         Width           =   270
      End
      Begin VB.Image imgPlay1 
         Height          =   270
         Left            =   720
         Picture         =   "frmAudica.frx":12BDE
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgDel2 
         Height          =   165
         Left            =   5880
         Picture         =   "frmAudica.frx":13010
         Top             =   360
         Width           =   525
      End
      Begin VB.Image imgDel1 
         Height          =   165
         Left            =   5880
         Picture         =   "frmAudica.frx":13172
         Top             =   240
         Width           =   525
      End
      Begin VB.Image imgAdd2 
         Height          =   165
         Left            =   5880
         Picture         =   "frmAudica.frx":132C9
         Top             =   120
         Width           =   525
      End
      Begin VB.Image imgAdd1 
         Height          =   165
         Left            =   5880
         Picture         =   "frmAudica.frx":13426
         Top             =   0
         Width           =   525
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   0
         Left            =   2160
         Picture         =   "frmAudica.frx":13589
         Top             =   0
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   1
         Left            =   2160
         Picture         =   "frmAudica.frx":14901
         Top             =   360
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   2
         Left            =   2160
         Picture         =   "frmAudica.frx":1585C
         Top             =   720
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   3
         Left            =   2160
         Picture         =   "frmAudica.frx":16775
         Top             =   1080
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   4
         Left            =   2160
         Picture         =   "frmAudica.frx":172D3
         Top             =   1440
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   5
         Left            =   2160
         Picture         =   "frmAudica.frx":17F43
         Top             =   1800
         Width           =   3675
      End
      Begin VB.Image imgLogo 
         Height          =   405
         Index           =   6
         Left            =   2160
         Picture         =   "frmAudica.frx":18B85
         Top             =   2160
         Width           =   3675
      End
      Begin VB.Image imgBackward1 
         Height          =   255
         Left            =   0
         Picture         =   "frmAudica.frx":1D967
         Top             =   0
         Width           =   510
      End
      Begin VB.Image imgBackward2 
         Height          =   255
         Left            =   0
         Picture         =   "frmAudica.frx":1E091
         Top             =   240
         Width           =   510
      End
      Begin VB.Image imgStop1 
         Height          =   255
         Left            =   480
         Picture         =   "frmAudica.frx":1E7BB
         Top             =   0
         Width           =   300
      End
      Begin VB.Image imgStop2 
         Height          =   255
         Left            =   480
         Picture         =   "frmAudica.frx":1EBF9
         Top             =   240
         Width           =   300
      End
   End
   Begin Spectre.Spectrum Spectrum1 
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   2160
      TabIndex        =   9
      Top             =   1950
      Width           =   615
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1260
      TabIndex        =   7
      Top             =   1500
      Width           =   1530
   End
   Begin VB.Label lblArtist 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1260
      TabIndex        =   6
      Top             =   1380
      Width           =   1530
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1260
      TabIndex        =   5
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label lblBitrate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1245
      TabIndex        =   4
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label lblSamplerate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1245
      TabIndex        =   3
      Top             =   1005
      Width           =   1530
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   1245
      TabIndex        =   2
      Top             =   870
      Width           =   1530
   End
   Begin VB.Image imgEject 
      Height          =   225
      Left            =   6210
      Picture         =   "frmAudica.frx":1F037
      Top             =   2175
      Width           =   465
   End
   Begin VB.Image imgMinimize 
      Height          =   180
      Left            =   5640
      Picture         =   "frmAudica.frx":1F619
      Top             =   1740
      Width           =   495
   End
   Begin VB.Image imgExit 
      Height          =   180
      Left            =   6180
      Picture         =   "frmAudica.frx":1FB0B
      Top             =   1740
      Width           =   495
   End
   Begin VB.Image imgPause 
      Height          =   225
      Left            =   5400
      Picture         =   "frmAudica.frx":1FFFD
      Top             =   2160
      Width           =   285
   End
   Begin VB.Image imgForeward 
      Height          =   240
      Left            =   4635
      Picture         =   "frmAudica.frx":203C3
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgPlay 
      Height          =   270
      Left            =   4110
      Picture         =   "frmAudica.frx":20A05
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgClear 
      Height          =   165
      Left            =   5280
      Picture         =   "frmAudica.frx":20E37
      Top             =   4095
      Width           =   525
   End
   Begin VB.Image imgDel 
      Height          =   165
      Left            =   4680
      Picture         =   "frmAudica.frx":20F91
      Top             =   4088
      Width           =   525
   End
   Begin VB.Image imgAdd 
      Height          =   165
      Left            =   4080
      Picture         =   "frmAudica.frx":210E8
      Top             =   4095
      Width           =   525
   End
   Begin VB.Image imgAudicaLogo 
      Height          =   405
      Left            =   3015
      Picture         =   "frmAudica.frx":2124B
      Top             =   2760
      Width           =   3675
   End
   Begin VB.Image imgStop 
      Height          =   255
      Left            =   3525
      Picture         =   "frmAudica.frx":2602D
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image imgBackward 
      Height          =   255
      Left            =   2805
      Picture         =   "frmAudica.frx":2646B
      Top             =   2160
      Width           =   510
   End
   Begin VB.Image imgSlider 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2250
      Picture         =   "frmAudica.frx":26B95
      Top             =   2400
      Width           =   465
   End
   Begin VB.Image imgSliderBack 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   2400
      Picture         =   "frmAudica.frx":270B7
      Top             =   1920
      Width           =   135
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuAudicaAbout 
         Caption         =   "::AUDICA::1.0"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "::PLAYLIST::"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEqualizer 
         Caption         =   "::EQUALIZER::"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuNexENCODE4 
         Caption         =   "::NEXENCODE4::"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep128773 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "::PLAY::"
         Begin VB.Menu mnuAddtoPlaylist 
            Caption         =   "::ADD TO PLAYLIST::"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep9379273972 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSingleFile 
            Caption         =   "::SINGLE FILE::"
         End
         Begin VB.Menu mnuFileGroup 
            Caption         =   "::FILE GROUP::"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep93798237896728 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWholeDirectory 
            Caption         =   "::WHOLE DIRECTORY::"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDirectoryGroup 
            Caption         =   "::DIRECTORY GROUP::"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep3972973 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlaylistWindow 
            Caption         =   "::PLAYLIST WINDOW::"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuSep9378927392642 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "::EXIT::"
      End
   End
End
Attribute VB_Name = "frmAudica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'GetWindowSettings hwnd
'SetShape
End Sub

Private Sub cmdSetshape_Click()
'InitStartup
'GetWindowSettings hwnd
'SetShape
End Sub

Private Sub Form_Load()

DoEvents
If Len(Command$) <> 0 Then
    Dim msg2 As String, msg3 As String
    msg2 = Right(Command$, Len(Command$) - 1)
    msg3 = Left(msg2, Len(msg2) - 1)
    'Playfile msg3
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    FormDrag Me
ElseIf Button = 2 Then
    PopupMenu mnuMain
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SavePlaylist
SaveSetting App.Title, "Settings", "SliderType", lInterface.iSliderType

Mp3Player.Stop
Mp3Player.Close
End Sub

Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgAdd.Picture = imgAdd2.Picture
End Sub

Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim msg As String
    msg = OpenDialog(Me, "Mp3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*", "Nexgen Audica - Select File ...", CurDir)
    AddtoPlaylist msg
    imgAdd.Picture = imgAdd1.Picture
End If
End Sub

Private Sub imgAudicaLogo_Click()
lInterface.iCurrentFlash = Audica_Logo
lInterface.iFlashLoop = 0
tmrGFXFlash.interval = 20

End Sub

Private Sub imgAudicaLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub imgBackward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgBackward.Picture = imgBackward1.Picture And Button = 1 Then imgBackward.Picture = imgBackward2.Picture
End Sub

Private Sub imgBackward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgBackward.Picture = imgBackward1.Picture
End Sub

Private Sub imgDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgDel.Picture = imgDel2.Picture
End Sub

Private Sub imgDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    lstPlaylist.RemoveItem lstPlaylist.ListIndex
    imgDel.Picture = imgDel1.Picture
End If
End Sub

Private Sub imgEject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgEject.Picture = imgEject2.Picture
End Sub

Private Sub imgEject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As String
If Button = 1 Then
    msg = PromptFile
    Playfile msg
    imgEject.Picture = imgEject1.Picture
End If
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgExit.Picture = imgExit2.Picture
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    imgExit.Picture = imgExit1.Picture
    Unload Me
    End
End If
End Sub

Private Sub imgForeward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgForeward.Picture = imgForeward2.Picture
End Sub

Private Sub imgForeward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As String
If Button = 1 Then
    If lPlaylist.pCurrent < lPlaylist.pCount Then
        Mp3Player.Stop
        Mp3Player.Close
        lPlaylist.pCurrent = lPlaylist.pCurrent + 1
        msg = lPlaylist.pFiles(lPlaylist.pCurrent).fFilepath & "\" & lPlaylist.pFiles(lPlaylist.pCurrent).fFilename
        
        Playfile msg
    End If
    imgForeward.Picture = imgForeward1.Picture
End If
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgMinimize.Picture = imgMinimize2.Picture
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    WindowState = vbMinimized
    imgMinimize.Picture = imgMinimize1.Picture
End If
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgPause.Picture = imgPause2.Picture
End Sub

Private Sub imgPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Mp3Player.Pause
    imgPause.Picture = imgPause1.Picture
End If
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Mp3Player.Play
    imgPlay.Picture = imgPlay2.Picture
End If
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And imgPlay.Picture = imgPlay2.Picture Then imgPlay.Picture = imgPlay1.Picture
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And imgStop.Picture = imgStop1.Picture Then imgStop.Picture = imgStop2.Picture
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Mp3Player.Stop
    imgStop.Picture = imgStop1.Picture
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lPlaylist.pVolumeButton = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, p As Integer

SliderVolume Int(Y)
'If lPlaylist.pVolumeButton = True Then
'    i = Y / 16 + 122
'    If i > 122 And Y < 1100 Then
'        imgSlider.Top = i
'        p = Y / 100 * 9.1
'        mp3Player.SetVolumeP p, p
'    End If
'End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    lPlaylist.pVolumeButton = False
    SaveSetting App.Title, "frmAudica", "VolumePos", imgSlider.Top
End If
End Sub

Private Sub lstPlaylist_DblClick()
Dim i As Integer
i = FindPlaylistIndex(lstPlaylist.Text)
If i <> 0 Then
    Playfile lPlaylist.pFiles(i).fFilepath & "\" & lPlaylist.pFiles(i).fFilename
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuSingleFile_Click()
imgEject_MouseDown 1, 0, 0, 0
imgEject_MouseUp 1, 0, 0, 0
End Sub

Private Sub mp3player_Failure(ByVal ErrorCode As Long, ByVal ErrStr As String)
If ErrorCode = 61100 Then
    Dim msg As String
    msg = MsgBox("Function failed using " & frmAudica.Mp3Player.GetDevName(lSettings.sOutputDevice) & " failed. Would you like to use following device?" & vbCrLf & vbCrLf & frmAudica.Mp3Player.GetDevName(lSettings.sOutputDevice + 1), vbYesNo + vbQuestion)
    If msg = vbYes Then
        lSettings.sOutputDevice = lSettings.sOutputDevice + 1
        frmAudica.Mp3Player.SetOutDevice lSettings.sOutputDevice
    ElseIf msg = vbNo Then
        frmOutputDevice.Show 1
    End If
Else
    MsgBox "Your player had an error" & vbCrLf & ErrStr, vbInformation
End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub mp3Player_ThreadEnded()
tmrFake.Enabled = False
End Sub

Private Sub tmrFake_Timer()
 Dim f As Single, H As Single
 Dim j As Single

crazy = 15  'change crazy to get a more wild analyzer pattern (higher more energy)

Randomize

Dim stp!

H = 25
f = 25

 With Spectrum1
     w! = .SpectrumWidth
     hg! = .SpectrumHeight


For j = 1 To w
    If (j - 10) >= 0 Then
        f = Int((H + crazy) - (H - crazy) + 1) * Rnd + (H - crazy)
    Else
        f = Int(crazy + 1) * Rnd
    End If
    
    If f <= hg And H <= hg Then
       If f > H Then stp = -1 Else stp = 1
       
       For X = j - 1 To j
          For Y = f To H Step stp
            If X >= 0 And X <= w And Y >= 0 And Y <= hg Then
               .SetLine j, f, j - 1, H
            End If
          Next
       Next
            
    End If
    H = f
Next j

End With

End Sub

'Private Sub Spectrum1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FormDrag Me
'End Sub

Private Sub tmrGFXFlash_Timer()
Select Case lInterface.iCurrentFlash
Case Audica_Logo
    Select Case lInterface.iFlashLoop
    Case 0
        tmrGFXFlash.interval = 40
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 10, 10
    Case 1
        imgAudicaLogo.Picture = imgLogo(0).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 20, 20
    Case 2
        imgAudicaLogo.Picture = imgLogo(1).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 30, 30
    Case 3
        imgAudicaLogo.Picture = imgLogo(2).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 40, 40
    Case 4
        imgAudicaLogo.Picture = imgLogo(3).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 50, 50
    Case 5
        imgAudicaLogo.Picture = imgLogo(3).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 60, 60
    Case 6
        imgAudicaLogo.Picture = imgLogo(4).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 70, 70
    Case 7
        imgAudicaLogo.Picture = imgLogo(5).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
        Spectrum1.SetPoint 80, 80
        Pause 0.2
    Case 8
        imgAudicaLogo.Picture = imgLogo(4).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 9
        imgAudicaLogo.Picture = imgLogo(3).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 10
        imgAudicaLogo.Picture = imgLogo(2).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 11
        imgAudicaLogo.Picture = imgLogo(1).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 12
        imgAudicaLogo.Picture = imgLogo(0).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 13
        imgAudicaLogo.Picture = imgLogo(6).Picture
        lInterface.iFlashLoop = lInterface.iFlashLoop + 1
    Case 14
        lInterface.iFlashLoop = 0
        lInterface.iCurrentFlash = 0
        tmrGFXFlash.Enabled = False
    End Select
End Select
End Sub
