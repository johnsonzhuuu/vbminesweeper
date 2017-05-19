VERSION 5.00
Begin VB.Form frmHighScores 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   12405
   Begin VB.ListBox lstExpert 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   8280
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.ListBox lstIntermediate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   4200
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.ListBox lstBeginner 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Beginner                          Intermediate                      Expert"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   12015
   End
End
Attribute VB_Name = "frmHighScores"
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
    
    Dim X As Integer
    
    For X = 1 To 5
        lstBeginner.AddItem Beginner(X).Name & "  " & Beginner(X).Time
        lstIntermediate.AddItem Intermediate(X).Name & "  " & Intermediate(X).Time
        lstExpert.AddItem Expert(X).Name & "  " & Expert(X).Time
    Next X
    
    CentreForm frmHighScores, Screen.Width, Screen.Height
End Sub
