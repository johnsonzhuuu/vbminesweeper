Attribute VB_Name = "Module1"
'Author: Johnson Zhu
'Title: RCI's Minesweeper
'Date: June 2, 2016
'Files: Final_ZhuJ.vbp, Final_ZhuJ.frm, Final_ZhuJ.bas, Final_ZhuJ.frx, frmAbout.frm, frmHighScores.frm, frmSplash.frm, frmSplash.frx
'Purpose: The purpose of this application is to play the game Minesweeper.

Option Explicit

Global Const RECLEN = 22
Global Const ENTRIES = 5
Global Seconds As Integer
Global Beginner(1 To ENTRIES) As ScoreRec
Global Intermediate(1 To ENTRIES) As ScoreRec
Global Expert(1 To ENTRIES) As ScoreRec
Type ScoreRec
    Name As String * 20
    Time As Integer
End Type

'This procedure reveals the tile in every direction of a tile and this procedure is repeated for every blank tile revealed.

Sub RevealBlanks(ByRef Grid As Control, ByVal X As Integer, ByVal Y As Integer, ByRef Value() As String, ByRef Status() As String, ByRef Image As Control, ByRef NumRevealed As Integer)
       
    Dim C As Integer, R As Integer

    For C = X - 1 To X + 1
        For R = Y - 1 To Y + 1
            If C >= 0 And R >= 0 And C < Grid.Cols And R < Grid.Rows Then
                If Value(C, R) <> "mine" And Status(C, R) <> "revealed" And Status(C, R) <> "flag" Then
                    RevealTile Grid, C, R, Value(), Status(), Image, NumRevealed
                    If Value(C, R) = "0" Then
                        RevealBlanks Grid, C, R, Value(), Status(), Image, NumRevealed
                    End If
                End If
            End If
        Next R
    Next C
 
End Sub

'This procedure generates mines in random locations of a grid.

Sub GenerateMines(ByRef Grid As Control, ByRef Value() As String, ByVal NumMines As Integer)

    Dim X As Integer, Y As Integer, K As Integer

    Do
        X = Int(Rnd * Grid.Cols)
        Y = Int(Rnd * Grid.Rows)
        If Value(X, Y) <> "mine" Then
            Grid.Col = X
            Grid.Row = Y
            Value(X, Y) = "mine"
            K = K + 1
        End If
    Loop While K < NumMines
End Sub

'This procedure calculates the value of each non-mine tile on the grid.

Sub CalcMines(ByRef Grid As Control, ByRef Value() As String, ByRef Status() As String)
    
    Dim C As Integer, R As Integer, K As Integer, X As Integer, Y As Integer
    
    For C = 0 To Grid.Cols - 1
        Grid.Col = C
        For R = 0 To Grid.Rows - 1
            Grid.Row = R
            If Value(C, R) <> "mine" Then
                K = 0
                For X = C - 1 To C + 1
                    For Y = R - 1 To R + 1
                        If X >= 0 And Y >= 0 And X < Grid.Cols And Y < Grid.Rows Then
                            If Value(X, Y) = "mine" Then
                                K = K + 1
                            End If
                        End If
                    Next Y
                Next X
                Value(C, R) = VBA.Trim$(VBA.Str$(K))
            End If
        Next R
    Next C
End Sub

'This procedure reveals a tile and its value.

Sub RevealTile(ByRef Grid As Control, ByVal X As Integer, ByVal Y As Integer, ByRef Value() As String, ByRef Status() As String, ByRef Image As Control, ByRef NumRevealed As Integer)
    
    Grid.Col = X
    Grid.Row = Y
    Set Grid.CellPicture = Image
    Status(X, Y) = "revealed"
    If Value(X, Y) <> "0" Then
        Select Case Value(X, Y)
            Case 1
                Grid.CellForeColor = RGB(0, 0, 255)
            Case 2
                Grid.CellForeColor = RGB(77, 153, 0)
            Case 3
                Grid.CellForeColor = RGB(255, 0, 0)
            Case 4
                Grid.CellForeColor = RGB(0, 51, 102)
            Case 5
                Grid.CellForeColor = RGB(128, 0, 0)
            Case 6
                Grid.CellForeColor = RGB(0, 153, 153)
            Case 7
                Grid.CellForeColor = RGB(0, 0, 0)
            Case 8
                Grid.CellForeColor = RGB(166, 166, 166)
        End Select
        Grid.Text = Value(X, Y)
    End If
    NumRevealed = NumRevealed + 1

End Sub

'This procedure sets the game to the "lost" state, revealing all mines and wrongly flagged tiles.

Sub Lose(ByRef Grid As Control, ByVal X As Integer, ByVal Y As Integer, ByRef Value() As String, ByRef Status() As String, ByRef MineImg As Control, ByRef RedMineImg As Control, ByRef CmdButton As Control, ByRef LoseImg As Control, ByRef WrongImg As Control)

    Dim C As Integer, R As Integer
    
    Grid.Visible = False
    Grid.Enabled = False
    CmdButton.Picture = LoseImg
    For C = 0 To Grid.Cols - 1
        Grid.Col = C
        For R = 0 To Grid.Rows - 1
            Grid.Row = R
            If Value(C, R) = "mine" And (Status(C, R) = "hidden" Or Status(C, R) = "question") Then
                Set Grid.CellPicture = MineImg
            ElseIf Value(C, R) <> "mine" And Status(C, R) = "flag" Then
                Set Grid.CellPicture = WrongImg
            End If
        Next R
    Next C
    
    Grid.Col = X
    Grid.Row = Y
    Grid.CellPicture = RedMineImg
    Grid.Visible = True
End Sub

'This procedure resets the game to its "default" state, before the user has clicked anything.

Sub Reset(ByRef Grid As Control, ByRef Value() As String, ByRef Status() As String, ByRef Button As Control, ByRef HappyImg As Control, ByRef NumRevealed As Integer, ByRef Flags As Integer, ByVal NumMines As Integer, ByRef FlagLbl As Control, ByRef Seconds As Integer, ByRef TimeLbl As Control, ByRef Timer As Control, ByRef TileImg As Control)
    
    Dim X As Integer, Y As Integer
        
    Button.Picture = HappyImg
    Grid.Visible = False
    With Grid
        For X = 0 To .Cols - 1
            .Col = X
            For Y = 0 To .Rows - 1
                .Row = Y
                .Text = ""
                Set .CellPicture = TileImg.Picture
                Value(X, Y) = ""
                Status(X, Y) = "hidden"
            Next Y
        Next X
    End With
    
    Flags = NumMines
    FlagLbl.Caption = Flags
    Timer.Enabled = False
    Seconds = 0
    TimeLbl.Caption = Seconds
    NumRevealed = 0
    Grid.Enabled = True
    Grid.Visible = True
    
End Sub

'This procedure sets the game to the "win" state, flagging all unflagged mines.

Sub Win(ByRef Grid As Control, ByRef Button As Control, ByRef WinImage As Control, ByRef Value() As String, ByRef FlagImage As Control, ByRef FlagLbl As Control)

    Dim C As Integer, R As Integer
    
    Grid.Visible = False
    Grid.Enabled = False
    Button.Picture = WinImage
    FlagLbl.Caption = "0"
    
    For C = 0 To Grid.Cols - 1
        For R = 0 To Grid.Rows - 1
            If Value(C, R) = "mine" Then
                Grid.Col = C
                Grid.Row = R
                Set Grid.CellPicture = FlagImage
            End If
        Next R
    Next C
    Grid.Visible = True
    
End Sub

'This procedure marks the tile with a blank, flag, or question mark, depending on its current state.

Sub Mark(ByRef Grid As Control, ByVal X As Integer, ByVal Y As Integer, ByRef Value() As String, ByRef Status() As String, ByRef TileImg As Control, ByRef FlagImg As Control, ByRef QImg As Control, ByRef Flags As Integer)
    
    With Grid
        .Col = .MouseCol
        .Row = .MouseRow
        If Status(X, Y) = "hidden" Then
            Status(X, Y) = "flag"
            Set .CellPicture = FlagImg
            Flags = Flags - 1
        ElseIf Status(X, Y) = "flag" Then
            Status(X, Y) = "question"
            Set .CellPicture = QImg
            Flags = Flags + 1
        ElseIf Status(X, Y) = "question" Then
            Status(X, Y) = "hidden"
            Set .CellPicture = TileImg
        End If
    End With
End Sub

'This procedure loads the scores from a record file.

Sub LoadScores(ByVal FilePath As String, ByRef Score() As ScoreRec)
    
    Dim X As Integer
    
    Open FilePath For Random As #1 Len = RECLEN
        Do While Not EOF(1) And X < 5
            X = X + 1
            Get #1, X, Score(X)
        Loop
    Close #1
End Sub

'This procedure reads the scores from a text file.

Sub ReadText(ByVal FilePath As String, ByRef Score() As ScoreRec)

    Dim X As Integer
    
    Open App.Path & "\highscores.txt" For Input As #1
        Do While Not EOF(1)
            X = X + 1
            Input #1, Score(X).Name, Score(X).Time
        Loop
    Close #1

End Sub

'This procedure saves the scores in an array into a record file.

Sub SaveScores(ByVal FilePath As String, ByRef Score() As ScoreRec)
    
    Dim X As Integer
   
    On Error GoTo ErrorHandler
    Kill FilePath
    Open FilePath For Random As #1 Len = RECLEN
        For X = 1 To ENTRIES
            Put #1, X, Score(X)
        Next X
    Close #1
    Exit Sub
ErrorHandler:
    Resume Next
    
End Sub

'This procedure checks to see if the user has beaten a high score and shifts the scores accordingly.

Sub CheckScore(ByRef Score() As ScoreRec, ByVal Seconds As Integer)
    
    Dim X As Integer
    Dim Name As String
    Dim Updated As Boolean
    Dim Y As Integer

    Updated = False
    X = 1
    Do While X <= ENTRIES And Updated = False
        If Seconds < Score(X).Time Then
            For Y = ENTRIES - 1 To X Step -1
                Score(Y + 1).Name = Score(Y).Name
                Score(Y + 1).Time = Score(Y).Time
            Next Y
            Name = InputBox$("Enter your name: ", "New High Score!")
            Score(X).Name = Name
            Score(X).Time = Seconds
            Updated = True
        End If
        X = X + 1
    Loop
End Sub

'This procedure displays the splash screen.

Sub SplashScreen(ByRef FormName As Form, ByVal SS As Single)
    FormName.Show
    Delay SS
    Unload FormName
End Sub

'This procedure creates a delay.

Sub Delay(ByVal Interval As Single)
    
    Dim Start As Single
    Dim Current As Single
    
    Start = Timer()
    Do
        Current = Timer()
        DoEvents
    Loop Until (Current - Start) >= Interval
    
End Sub

'This procedure centres a form.

Sub CentreForm(ByRef FormName As Form, ByVal X As Integer, ByVal Y As Integer)

    FormName.Left = (X - FormName.Width) / 2
    FormName.Top = (Y - FormName.Height) / 2
    
End Sub

'This procedure centres the controls in the Minesweeper game.

Sub CentreControls(ByRef FormName As Form, ByRef Grid As Control, ByRef FlagLbl As Control, ByRef TimeLbl As Control, ByRef CmdButton As Control, ByVal CellSize As Integer)
        
    FormName.Width = CellSize * (Grid.Cols + 2)
    FormName.Height = CellSize * (Grid.Rows + 4)
    Grid.Left = (FormName.Width - CellSize * Grid.Cols) / 2 - 45
    Grid.Top = (FormName.Height - CellSize * Grid.Rows) / 2
    FlagLbl.Left = (FormName.Width - CellSize * Grid.Cols) / 2 - 45
    FlagLbl.Top = (FormName.Height - FlagLbl.Height) / 30
    TimeLbl.Left = FormName.Width - ((FormName.Width - CellSize * Grid.Cols) / 2) - TimeLbl.Width - 45
    TimeLbl.Top = (FormName.Height - TimeLbl.Height) / 30
    CmdButton.Left = (FormName.Width - CmdButton.Width) / 2 - 45
    CmdButton.Top = (FormName.Height - CmdButton.Height) / 40

End Sub

'This procedure sets the columns, rows, and number of mines according to the game level.

Sub SetLevel(ByRef Grid As Control, ByRef Lvl1 As Control, ByRef Lvl2 As Control, ByRef Lvl3 As Control, ByRef NumMines As Integer, ByVal CellSize As Integer)
    
    Dim C As Integer
    
    If Lvl1.Checked = True Then
        Grid.Cols = 9
        Grid.Rows = 9
        NumMines = 10
    ElseIf Lvl2.Checked = True Then
        Grid.Cols = 16
        Grid.Rows = 16
        NumMines = 40
    ElseIf Lvl3.Checked = True Then
        Grid.Cols = 30
        Grid.Rows = 16
        NumMines = 99
    End If
        
    For C = 0 To Grid.Cols - 1
        Grid.ColWidth(C) = CellSize
        Grid.ColAlignment(C) = 4
        Grid.Col = C
    Next C
    
End Sub
