VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Folder"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "This form loaded incorrectly"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
lFileshare.fReturn = Dir1.Path
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Icon = frmMain.Icon
End Sub
