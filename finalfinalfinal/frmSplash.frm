VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5745
   ClientLeft      =   1515
   ClientTop       =   3540
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Johnson Zhu
'Title: RCI's Minesweeper
'Date: June 2, 2016
'Files: Final_ZhuJ.vbp, Final_ZhuJ.frm, Final_ZhuJ.bas, Final_ZhuJ.frx, frmAbout.frm, frmHighScores.frm, frmSplash.frm, frmSplash.frx
'Purpose: The purpose of this application is to play the game Minesweeper.

Option Explicit

Private Sub Form_Load()
    CentreForm frmSplash, Screen.Width, Screen.Height
    SplashScreen frmSplash, 3
    frmMain.Show
End Sub
