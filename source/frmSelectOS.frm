VERSION 5.00
Begin VB.Form frmSelectOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select OS"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
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
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox lstOschoices 
      Height          =   1230
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "I am using..."
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelectOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
lInterface.iOS = lstOschoices.ListIndex + 1
SaveSetting App.Title, "Settings", "OS", lInterface.iOS
Unload Me
End Sub

Private Sub Form_Load()
Icon = frmMain.Icon
lstOschoices.AddItem "Windows 95"
lstOschoices.AddItem "Windows 98/ME"
lstOschoices.AddItem "Windows NT"
lstOschoices.AddItem "Windows 2000"
lstOschoices.AddItem "Windows XP"
End Sub
