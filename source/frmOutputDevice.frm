VERSION 5.00
Begin VB.Form frmOutputDevice 
   Caption         =   "Select Output Device"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutputDevice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   3270
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstOutputDev 
      Height          =   3765
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmOutputDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
SaveSetting "Audica", "Settings", "OutputDevice", lstOutputDev.ListIndex
lSettings.sOutputDevice = lstOutputDev.ListIndex
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer, i As Integer
If lInterface.iOS = Windows_XP Then
    x = InputBox("Select output device (max=2)")
    SaveSetting "Audica", "Settings", "OutputDevice", x
    lSettings.sOutputDevice = x
    Unload Me
    Exit Sub
End If
For i = 0 To 2
    lstOutputDev.AddItem frmMain.ctlMp3Player.GetDevName(i)
Next i
End Sub
