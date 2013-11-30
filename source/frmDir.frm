VERSION 5.00
Begin VB.Form frmDir 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
