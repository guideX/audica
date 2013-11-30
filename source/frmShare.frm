VERSION 5.00
Begin VB.Form frmShare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Options"
      Height          =   495
      Index           =   6
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Browse"
      Height          =   495
      Index           =   5
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Uploads"
      Height          =   495
      Index           =   4
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Downloads"
      Height          =   495
      Index           =   3
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Trade"
      Height          =   495
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&Search"
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optToolbar 
      Caption         =   "&News"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraNews 
      Caption         =   "News"
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7575
   End
   Begin VB.Frame fraBrowse 
      Caption         =   "Browse"
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7575
      Begin VB.ListBox lstFileBorwse 
         Height          =   2595
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   7335
      End
      Begin VB.TextBox txtBrowseNickname 
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraUpload 
      Caption         =   "Uploads"
      Height          =   3495
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   7575
      Begin VB.ListBox lstUploads 
         Height          =   3060
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame fraDownloads 
      Caption         =   "Downloads"
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   7575
      Begin VB.ListBox lstDownloads 
         Height          =   3060
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   7575
      Begin VB.CommandButton Command3 
         Caption         =   "Del"
         Height          =   255
         Left            =   4200
         TabIndex        =   39
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox lstMp3 
         Height          =   840
         Left            =   3360
         TabIndex        =   37
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   29
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtLocationOfMedia 
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Hidden Components"
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.Label Label9 
         Caption         =   "Supported Types:"
         Height          =   255
         Left            =   3360
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Password:"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Location of media:"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame fraTrade 
      Caption         =   "Trade"
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7575
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search"
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   7575
      Begin VB.ListBox lstFiles 
         Height          =   1425
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   7335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Matching Files:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "&Title:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Artist:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H8000000D&
      Caption         =   "Status: Idle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   7575
   End
End
Attribute VB_Name = "frmShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim msg As String

End Sub

Private Sub Form_Load()
'WebBrowser1.Visible = True
'WebBrowser1.Navigate "http://www.team-nexgen.com"
FileshareLoadSettings
End Sub

Private Sub optToolbar_Click(Index As Integer)
fraNews.Visible = False
fraSearch.Visible = False
fraTrade.Visible = False
fraDownloads.Visible = False
fraBrowse.Visible = False
fraUpload.Visible = False
fraOptions.Visible = False
Select Case Index

Case 0
    fraNews.Visible = True
Case 1
    fraSearch.Visible = True
Case 2
    fraTrade.Visible = True
Case 3
    fraDownloads.Visible = True
Case 4
    fraUpload.Visible = True
Case 5
    fraBrowse.Visible = True
Case 6
    fraOptions.Visible = True
End Select
End Sub
