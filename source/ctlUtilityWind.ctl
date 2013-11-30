VERSION 5.00
Begin VB.UserControl ctlUtilityWind 
   BackColor       =   &H00000000&
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3030
   ScaleWidth      =   3855
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Add"
      Height          =   255
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   900
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Del"
      Height          =   255
      Left            =   960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   900
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Load"
      Height          =   255
      Left            =   1920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Clear"
      Height          =   255
      Left            =   2880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   900
   End
   Begin VB.ListBox lstPlaylist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0C0&
      Height          =   2820
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "ctlUtilityWind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub InitPlaylist()
Dim i As Integer, msg As String
lstPlaylist.Clear
For i = 1 To lPlaylist.pCount
    msg = lPlaylist.pFiles(i).fFilepath & lPlaylist.pFiles(i).fFilename
    If DoesFileExist(msg) = True Then lstPlaylist.AddItem lPlaylist.pFiles(i).fFilename
Next i
End Sub

Private Sub Command1_Click()
Dim msg As String, i As Integer
i = FindPlaylistIndex(lstPlaylist.Text)


End Sub

Private Sub lstPlaylist_DblClick()
Dim i As Integer

i = FindPlaylistIndex(lstPlaylist.Text)
OpenFile lPlaylist.pFiles(i).fFilepath & "\" & lPlaylist.pFiles(i).fFilename
Playfile i, lPlaylist.pFiles(i).fFiletype
End Sub
