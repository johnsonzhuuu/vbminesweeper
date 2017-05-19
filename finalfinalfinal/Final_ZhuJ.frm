VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RCI's Minesweeper"
   ClientHeight    =   10785
   ClientLeft      =   1245
   ClientTop       =   675
   ClientWidth     =   16905
   Icon            =   "Final_ZhuJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   16905
   Begin MSFlexGridLib.MSFlexGrid grdBoard 
      Height          =   8775
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   15478
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   480
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   3
      ScrollBars      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   7680
   End
   Begin VB.CommandButton cmdNew 
      Height          =   615
      Left            =   2160
      Picture         =   "Final_ZhuJ.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgHappy 
      Height          =   480
      Left            =   1320
      Picture         =   "Final_ZhuJ.frx":1194
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgShocked 
      Height          =   480
      Left            =   1800
      Picture         =   "Final_ZhuJ.frx":1A5E
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWrong 
      Height          =   480
      Left            =   2280
      Picture         =   "Final_ZhuJ.frx":2328
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRedMine 
      Height          =   480
      Left            =   1800
      Picture         =   "Final_ZhuJ.frx":2BF2
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMine 
      Height          =   480
      Left            =   1320
      Picture         =   "Final_ZhuJ.frx":34BC
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   2280
      Picture         =   "Final_ZhuJ.frx":3D86
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   1800
      Picture         =   "Final_ZhuJ.frx":4650
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWin 
      Height          =   480
      Left            =   2760
      Picture         =   "Final_ZhuJ.frx":4F1A
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLose 
      Height          =   480
      Left            =   2280
      Picture         =   "Final_ZhuJ.frx":57E4
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTile2 
      Height          =   480
      Left            =   2760
      Picture         =   "Final_ZhuJ.frx":60AE
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTile1 
      Height          =   480
      Left            =   1320
      Picture         =   "Final_ZhuJ.frx":6978
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblFlags 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginner 
         Caption         =   "Beginner"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIntermediate 
         Caption         =   "Intermediate"
      End
      Begin VB.Menu mnuExpert 
         Caption         =   "Expert"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScores 
         Caption         =   "High Scores"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Const MAXCOLS = 30
Const MAXROWS = 16

Dim Value(0 To MAXCOLS - 1, 0 To MAXROWS - 1) As String
Dim Status(0 To MAXCOLS - 1, 0 To MAXROWS - 1) As String
Dim NumMines As Integer
Dim NumRevealed As Integer
Dim Flags As Integer

Private Sub cmdNew_Click()
    
    Me.Visible = False
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True
    
End Sub

Private Sub Form_Load()
        
    Dim C As Integer, R As Integer
    
    Randomize
    
    LoadScores App.Path & "\beginner.rec", Beginner()
    LoadScores App.Path & "\intermediate.rec", Intermediate()
    LoadScores App.Path & "\expert.rec", Expert()
    Me.Visible = False
    SetLevel grdBoard, mnuBeginner, mnuIntermediate, mnuExpert, NumMines, imgTile1.Width
    CentreControls frmMain, grdBoard, lblFlags, lblTimer, cmdNew, imgTile1.Width
    CentreForm frmMain, Screen.Width, Screen.Height
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True
    
End Sub

Private Sub grdBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If Status(grdBoard.Col, grdBoard.Row) <> "revealed" Then
            cmdNew.Picture = imgShocked
            If Status(grdBoard.Col, grdBoard.Row) <> "flag" Then
                Set grdBoard.CellPicture = imgTile2
            End If
        End If
    End If

End Sub

Private Sub grdBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrTimer.Enabled = True
    cmdNew.Picture = imgHappy
    With grdBoard
        If Button = 1 And Status(.Col, .Row) <> "flag" Then
            If Value(.Col, .Row) = "mine" Then
                tmrTimer.Enabled = False
                Lose grdBoard, .Col, .Row, Value(), Status(), imgMine, imgRedMine, cmdNew, imgLose, imgWrong
            ElseIf Status(.Col, .Row) = "hidden" Or Status(.Col, .Row) = "question" Then
                RevealTile grdBoard, .Col, .Row, Value(), Status(), imgTile2, NumRevealed
                If Value(.Col, .Row) = "0" Then
                    grdBoard.Visible = False
                    RevealBlanks grdBoard, .Col, .Row, Value(), Status(), imgTile2, NumRevealed
                    grdBoard.Visible = True
                End If
            End If
            
            If NumRevealed + NumMines = grdBoard.Cols * grdBoard.Rows Then
                tmrTimer.Enabled = False
                Win grdBoard, cmdNew, imgWin, Value(), imgFlag, lblFlags
                If mnuBeginner.Checked = True Then
                    CheckScore Beginner(), Seconds
                    SaveScores App.Path & "\beginner.rec", Beginner()
                ElseIf mnuIntermediate.Checked = True Then
                    CheckScore Intermediate(), Seconds
                    SaveScores App.Path & "\intermediate.rec", Intermediate()
                ElseIf mnuExpert.Checked = True Then
                    CheckScore Expert(), Seconds
                    SaveScores App.Path & "\expert.rec", Expert()
                End If
            End If
        ElseIf Button = 2 Then
            Mark grdBoard, .MouseCol, .MouseRow, Value(), Status(), imgTile1, imgFlag, imgQuestion, Flags
            lblFlags.Caption = Flags
        End If
    End With
End Sub



Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBeginner_Click()
    
    mnuBeginner.Checked = True
    mnuIntermediate.Checked = False
    mnuExpert.Checked = False
    
    Me.Visible = False
    SetLevel grdBoard, mnuBeginner, mnuIntermediate, mnuExpert, NumMines, imgTile1.Width
    CentreControls frmMain, grdBoard, lblFlags, lblTimer, cmdNew, imgTile1.Width
    CentreForm frmMain, Screen.Width, Screen.Height
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExpert_Click()
    
    mnuBeginner.Checked = False
    mnuIntermediate.Checked = False
    mnuExpert.Checked = True
    
    Me.Visible = False
    SetLevel grdBoard, mnuBeginner, mnuIntermediate, mnuExpert, NumMines, imgTile1.Width
    CentreControls frmMain, grdBoard, lblFlags, lblTimer, cmdNew, imgTile1.Width
    CentreForm frmMain, Screen.Width, Screen.Height
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True
End Sub

Private Sub mnuIntermediate_Click()
    
    mnuBeginner.Checked = False
    mnuIntermediate.Checked = True
    mnuExpert.Checked = False
    
    Me.Visible = False
    SetLevel grdBoard, mnuBeginner, mnuIntermediate, mnuExpert, NumMines, imgTile1.Width
    CentreControls frmMain, grdBoard, lblFlags, lblTimer, cmdNew, imgTile1.Width
    CentreForm frmMain, Screen.Width, Screen.Height
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True
End Sub

Private Sub mnuNew_Click()
    
    Me.Visible = False
    SetLevel grdBoard, mnuBeginner, mnuIntermediate, mnuExpert, NumMines, imgTile1.Width
    Reset grdBoard, Value(), Status(), cmdNew, imgHappy, NumRevealed, Flags, NumMines, lblFlags, Seconds, lblTimer, tmrTimer, imgTile1
    GenerateMines grdBoard, Value(), NumMines
    CalcMines grdBoard, Value(), Status()
    Me.Visible = True
    
End Sub

Private Sub mnuScores_Click()
    frmHighScores.Show vbModal
End Sub

Private Sub tmrTimer_Timer()
    
    Seconds = Seconds + 1
    lblTimer.Caption = Seconds
    lblTimer.Enabled = True
    
End Sub
