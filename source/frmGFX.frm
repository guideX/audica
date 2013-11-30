VERSION 5.00
Begin VB.Form frmGFX 
   Caption         =   "Hidden"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picBulkGfx 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8295
      ScaleWidth      =   8535
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Image imgUtilityWind 
         Appearance      =   0  'Flat
         Height          =   4050
         Left            =   3360
         Picture         =   "frmGFX.frx":0000
         Top             =   4080
         Visible         =   0   'False
         Width           =   4860
      End
      Begin VB.Image imgSmPause2 
         Height          =   315
         Left            =   120
         Picture         =   "frmGFX.frx":2E36
         Top             =   4920
         Width           =   435
      End
      Begin VB.Image imgSmPause1 
         Height          =   315
         Left            =   120
         Picture         =   "frmGFX.frx":33F6
         Top             =   5280
         Width           =   435
      End
      Begin VB.Image imgSmBack2 
         Height          =   285
         Left            =   600
         Picture         =   "frmGFX.frx":37C2
         Top             =   4920
         Width           =   390
      End
      Begin VB.Image imgSmBack1 
         Height          =   285
         Left            =   600
         Picture         =   "frmGFX.frx":3D1C
         Top             =   5280
         Width           =   390
      End
      Begin VB.Image imgSmPlay1 
         Height          =   315
         Left            =   120
         Picture         =   "frmGFX.frx":408E
         Top             =   4200
         Width           =   435
      End
      Begin VB.Image imgSmPlay2 
         Height          =   315
         Left            =   120
         Picture         =   "frmGFX.frx":446D
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image imgSmStop1 
         Height          =   315
         Left            =   600
         Picture         =   "frmGFX.frx":4A08
         Top             =   4200
         Width           =   435
      End
      Begin VB.Image imgSmStop2 
         Height          =   315
         Left            =   600
         Picture         =   "frmGFX.frx":4DCC
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image imgSmNext1 
         Height          =   315
         Left            =   1080
         Picture         =   "frmGFX.frx":535F
         Top             =   4200
         Width           =   435
      End
      Begin VB.Image imgSmNext2 
         Height          =   315
         Left            =   1080
         Picture         =   "frmGFX.frx":572F
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image imgSmOptions1 
         Height          =   345
         Left            =   1080
         Picture         =   "frmGFX.frx":5CF5
         Top             =   4920
         Width           =   390
      End
      Begin VB.Image imgSmOptions2 
         Height          =   345
         Left            =   1080
         Picture         =   "frmGFX.frx":60B7
         Top             =   5280
         Width           =   390
      End
      Begin VB.Image imgSmEject1 
         Height          =   345
         Left            =   1920
         Picture         =   "frmGFX.frx":6660
         Top             =   4200
         Width           =   390
      End
      Begin VB.Image imgSmEject2 
         Height          =   345
         Left            =   1920
         Picture         =   "frmGFX.frx":6A23
         Top             =   4560
         Width           =   390
      End
      Begin VB.Image imgSmWindow 
         Appearance      =   0  'Flat
         Height          =   4575
         Left            =   0
         Picture         =   "frmGFX.frx":6FB9
         Top             =   120
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Image imgAbout 
         Height          =   3480
         Left            =   120
         Picture         =   "frmGFX.frx":D88E
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Image imgVolume 
         Height          =   960
         Left            =   3000
         Picture         =   "frmGFX.frx":11489
         Top             =   3120
         Visible         =   0   'False
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmGFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

