VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrHideMe 
      Interval        =   2000
      Left            =   1800
      Top             =   1200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrHideMe_Timer()
Unload Me
End Sub
