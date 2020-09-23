VERSION 5.00
Begin VB.Form minesweeper 
   AutoRedraw      =   -1  'True
   Caption         =   "Minesweeper"
   ClientHeight    =   2955
   ClientLeft      =   6600
   ClientTop       =   4380
   ClientWidth     =   2385
   Icon            =   "mineForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   2385
   Begin VB.PictureBox flags 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   385
      Left            =   240
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   630
   End
   Begin VB.PictureBox timer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   385
      Left            =   1440
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   630
   End
   Begin VB.PictureBox face 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   385
      Left            =   960
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      Top             =   240
      Width           =   385
   End
   Begin VB.PictureBox gameScreen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1915
      Left            =   180
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   820
      Width           =   1915
   End
   Begin VB.Menu gameopts 
      Caption         =   "&Game"
      Begin VB.Menu newgametime 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu beginneropts 
         Caption         =   "&Beginner"
      End
      Begin VB.Menu intermediateopts 
         Caption         =   "&Intermediate"
      End
      Begin VB.Menu expertopts 
         Caption         =   "&Expert"
      End
      Begin VB.Menu customopts 
         Caption         =   "&Custom..."
      End
      Begin VB.Menu bord2 
         Caption         =   "-"
      End
      Begin VB.Menu besttimes 
         Caption         =   "Best &Times..."
      End
      Begin VB.Menu bord3 
         Caption         =   "-"
      End
      Begin VB.Menu quitout 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu helpmain 
      Caption         =   "&Help"
      Begin VB.Menu helptopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu bord4 
         Caption         =   "-"
      End
      Begin VB.Menu aboutmine 
         Caption         =   "&About Minesweeper"
      End
   End
End
Attribute VB_Name = "minesweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private t As Long
'holds field info
Private field(-1 To 100, -1 To 100) As Integer
'where are the mines
Private mines(-1 To 100, -1 To 100) As Boolean
'numbers in cells of surrounding mines
Private nums(-1 To 100, -1 To 100) As Integer
'records when a cell has been uncovered
Private hit(-1 To 100, -1 To 100) As Boolean
Private flagged(0 To 99, 0 To 99) As Boolean
Private lastX As Integer, lastY As Integer
'dimension of board
Public szWid As Integer, szHi As Integer
'...
Public numMines As Integer
'i use this for double clicking clears
Private button1 As Boolean
Private button2 As Boolean
'...
Private time As Double
Private isDead As Boolean, gameover As Boolean, startTime As Boolean
'this keeps track of what the smiley face is doing
Private currentface As Integer
Private numflags As Integer
Private lastFlagCnt As Integer

Private Sub aboutmine_Click()
    MsgBox "I coded this today. 5/9/04 - John Hollister"
End Sub

'''the following is just setting vars for the different difficulty settings.

Private Sub beginneropts_Click()
    beginneropts.Checked = True
    intermediateopts.Checked = False
    expertopts.Checked = False
    customopts.Checked = False
    numMines = 10
    szWid = 8
    szHi = 8
    lastFlagCnt = 10
    numflags = 10
    Call makeNew
    Call drawStats
End Sub

Private Sub besttimes_Click()
    minesweeper.Enabled = False
    Load best
    best.Show
    While minesweeper.Enabled = False
        DoEvents
    Wend
End Sub

Private Sub customopts_Click()
    minesweeper.Enabled = False
    beginneropts.Checked = False
    intermediateopts.Checked = False
    expertopts.Checked = False
    customopts.Checked = True
    Load customize
    customize.Show
    customize.Enabled = True
    'wait for user to enter dimensions...
    While minesweeper.Enabled = False
        DoEvents
    Wend
    lastFlagCnt = numMines
    numflags = numMines
    Call makeNew
    Call drawStats
End Sub

Private Sub helptopics_Click()
    MsgBox "Who doesn't know how to play minesweeper?"
End Sub

Private Sub intermediateopts_Click()
    beginneropts.Checked = False
    intermediateopts.Checked = True
    expertopts.Checked = False
    customopts.Checked = False
    numMines = 40
    szWid = 16
    szHi = 16
    lastFlagCnt = 40
    numflags = 40
    Call makeNew
    Call drawStats
End Sub

Private Sub expertopts_Click()
    beginneropts.Checked = True
    intermediateopts.Checked = False
    expertopts.Checked = False
    customopts.Checked = False
    numMines = 99
    szWid = 30
    szHi = 16
    lastFlagCnt = 99
    numflags = 99
    Call makeNew
    Call drawStats
End Sub


'smiley face stuff
Private Sub face_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    currentface = 2
    face.Cls
    BitBlt face.hDC, 0, 0, 26, 26, pics.faces.hDC, currentface * 26, 0, vbSrcCopy
    face.Refresh
End Sub

Private Sub face_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    currentface = 1
    Call makeNew
    Call gameLoop
End Sub

Private Sub Form_Load()
'setting some default variables...
    currentface = 1
    Me.Show
    numMines = 10
    beginneropts.Checked = True
    szWid = 8
    szHi = 8
    lastFlagCnt = 10
    numflags = 10
    Call makeNew
    Call drawStats
    Call gameLoop
End Sub

Public Function gameLoop()
'game loop here
Do
    t = GetTickCount + 31
    If startTime Then time = time + 1

    Call drawStats

    Call drawScreen
    
    Call checkEndGame

    While t > GetTickCount
        DoEvents
    Wend
Loop Until isDead Or gameover

End Function

Public Function makeNew()
    'reset vars for a new game
    gameScreen.Cls
    face.Cls
    timer.Cls
    flags.Cls
    minesweeper.Cls
    isDead = False
    startTime = False
    gameover = False
    numflags = lastFlagCnt
    time = 0
    
    BitBlt face.hDC, 0, 0, 26, 26, pics.faces.hDC, 26, 0, vbSrcCopy
    'draw zeros on clock
    BitBlt timer.hDC, 0, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    BitBlt timer.hDC, 13, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    BitBlt timer.hDC, 26, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    'draw zeros on flag counter
    BitBlt flags.hDC, 0, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    BitBlt flags.hDC, 13, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    BitBlt flags.hDC, 26, 0, 11, 21, pics.clock.hDC, 0, 0, vbSrcCopy
    'set up board size
    gameScreen.Width = szWid * 240
    gameScreen.Height = szHi * 240
    'set up form size
    minesweeper.Width = gameScreen.Width + 475
    minesweeper.Height = gameScreen.Height + 1675
    'place items...
    face.Left = minesweeper.Width / 2 - face.Width + 140
    timer.Left = minesweeper.Width - 1000
    
    Dim a As Integer, b As Integer
    'set up background
    'left vert. border
    For a = 0 To szHi - 1
        BitBlt minesweeper.hDC, 0, 55 + (a * 16), 16, 16, pics.back.hDC, 0, 56, vbSrcCopy
    Next a
        BitBlt minesweeper.hDC, 0, 55 + (a * 16) - 4, 16, 16, pics.back.hDC, 0, 71, vbSrcCopy
    'bottom horiz. border
    For b = 1 To szWid
        BitBlt minesweeper.hDC, (b * 16) - 3, 55 + (a * 16), 16, 16, pics.back.hDC, 16, 75, vbSrcCopy
    Next b
        BitBlt minesweeper.hDC, (b * 16) - 7, 55 + (a * 16) - 4, 16, 16, pics.back.hDC, 32, 71, vbSrcCopy
    'top horiz. borders
    For a = 1 To szWid
        BitBlt minesweeper.hDC, (a * 16) - 3, 0, 16, 55, pics.back.hDC, 16, 0, vbSrcCopy
    Next a
    'right vertical border
    For b = 0 To szHi - 1
        BitBlt minesweeper.hDC, (a * 16) - 3, 55 + (b * 16), 16, 16, pics.back.hDC, 0, 56, vbSrcCopy
    Next b
        BitBlt minesweeper.hDC, (a * 16) - 7, 0, 16, 55, pics.back.hDC, 32, 0, vbSrcCopy
    BitBlt minesweeper.hDC, 0, 0, 16, 55, pics.back.hDC, 0, 0, vbSrcCopy
    BitBlt minesweeper.hDC, 0, 0, 16, 55, pics.back.hDC, 0, 0, vbSrcCopy
        
    'draw the cells and reset arrays
    For a = 0 To szWid - 1
        For b = 0 To szHi - 1
            BitBlt gameScreen.hDC, a * 16, b * 16, 16, 16, pics.blocks.hDC, 0, 0, vbSrcCopy
            field(a, b) = 0
            mines(a, b) = False
            hit(a, b) = False
            flagged(a, b) = False
        Next b
    Next a
    
    'reset numbers array i made this an element extra on each side to handle error conditions
    'in recursion easier. (instead of 0 to 99, it's -1 to 100 for elements)
    For a = -1 To szWid
        For b = -1 To szHi
            nums(a, b) = 0
        Next b
    Next a

    'make mines
    Dim tx As Integer, ty As Integer
    Randomize
    For a = 1 To numMines
        tx = Int(Rnd * szWid)
        ty = Int(Rnd * szHi)
        'make sure mines are in a unique location
        If mines(tx, ty) Then
            a = a - 1
        Else
            'add up number of mines surround a cell
            mines(tx, ty) = True
            nums(tx - 1, ty - 1) = nums(tx - 1, ty - 1) + 1
            nums(tx, ty - 1) = nums(tx, ty - 1) + 1
            nums(tx + 1, ty - 1) = nums(tx + 1, ty - 1) + 1
            nums(tx - 1, ty) = nums(tx - 1, ty) + 1
            nums(tx + 1, ty) = nums(tx + 1, ty) + 1
            nums(tx - 1, ty + 1) = nums(tx - 1, ty + 1) + 1
            nums(tx, ty + 1) = nums(tx, ty + 1) + 1
            nums(tx + 1, ty + 1) = nums(tx + 1, ty + 1) + 1
        End If
    Next a
End Function

Private Sub gameScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this piece of code gets the relative location of the mouse click.. you will see it
    'a few more times further down
    Dim tx As Integer, ty As Integer
    tx = Int(X / 16)
    ty = Int(Y / 16)
    If Button = 1 And button1 = False Then
        'uh oh face because player is left clicking
        currentface = 3
        button1 = True
        'clear the cell if it isn't already
        If Not hit(tx, ty) And Not flagged(tx, ty) Then field(tx, ty) = 1
        lastX = tx
        lastY = ty
    ElseIf Button = 2 And button2 = False Then
        'this just catches a double click
        button2 = True
    End If
End Sub

Private Sub gameScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this code just lets you hold the mouse button down and drag over the field
    If Button = 1 Then
        If Not (hit(lastX, lastY)) And Not flagged(lastX, lastY) Then field(lastX, lastY) = 0
        Dim tx As Integer, ty As Integer
        tx = Int(X / 16)
        ty = Int(Y / 16)
        If tx >= 0 And ty >= 0 Then
            If Not hit(tx, ty) Then field(tx, ty) = 1
            lastX = tx
            lastY = ty
        End If
    End If
End Sub

Private Sub gameScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not (isDead Or gameover) Then
    'uncover
    currentface = 1
    Dim tx As Integer, ty As Integer
    tx = Int(X / 16)
    ty = Int(Y / 16)
    
    If Button = 1 Then startTime = True
    If ty >= 0 And tx >= 0 And ty <= szHi And tx <= szWid Then
        If button1 And button2 And hit(tx, ty) And field(tx, ty) = 1 And Not flagged(tx, ty) Then
            If canClear(tx, ty) Then
                'begin the double click clear recursion
                If field(tx - 1, ty) = 0 Then
                    If tx > 0 And Not (hit(tx - 1, ty)) Then Call uncover(tx - 1, ty)
                End If
                If field(tx, ty - 1) = 0 Then
                    If ty > 0 And Not (hit(tx, ty - 1)) Then Call uncover(tx, ty - 1)
                End If
                If field(tx + 1, ty) = 0 Then
                    If tx < szWid - 1 And Not (hit(tx + 1, ty)) Then Call uncover(tx + 1, ty)
                End If
                If field(tx, ty + 1) = 0 Then
                    If ty < szHi - 1 And Not (hit(tx, ty + 1)) Then Call uncover(tx, ty + 1)
                End If
                
                If field(tx - 1, ty - 1) = 0 Then
                    If tx > 0 And ty > 0 And Not (hit(tx - 1, ty - 1)) Then Call uncover(tx - 1, ty - 1)
                End If
                If field(tx + 1, ty - 1) = 0 Then
                    If tx < szWid - 1 And ty > 0 And Not (hit(tx + 1, ty - 1)) Then Call uncover(tx + 1, ty - 1)
                End If
                If field(tx - 1, ty + 1) = 0 Then
                    If tx > 0 And ty < szHi - 1 And Not (hit(tx - 1, ty + 1)) Then Call uncover(tx - 1, ty + 1)
                End If
                If field(tx + 1, ty + 1) = 0 Then
                    If tx < szWid - 1 And ty < szHi - 1 And Not (hit(tx + 1, ty + 1)) Then Call uncover(tx + 1, ty + 1)
                End If
            End If
        ElseIf Button = 1 And Not flagged(tx, ty) Then
            button1 = False
            hit(tx, ty) = True
            If mines(tx, ty) Then
                currentface = 0
                Call endgame(tx, ty)
            Else
                Call uncover(tx, ty)
            End If
        ElseIf Button = 2 And (Not gameover Or Not isDead) Then
            button2 = False
            If field(tx, ty) = 5 Then
                numflags = numflags + 1
                flagged(tx, ty) = False
                field(tx, ty) = 6
            ElseIf field(tx, ty) = 6 Then
                field(tx, ty) = 0
            ElseIf field(tx, ty) = 0 Then
                field(tx, ty) = 5
                flagged(tx, ty) = True
                numflags = numflags - 1
            End If
        End If
    Else
        If Not hit(lastX, lastY) Then
            field(lastX, lastY) = 0
            button1 = False
            button2 = False
        End If
    End If
End If
End Sub


Public Function drawScreen()
'draw the screen...
    gameScreen.Cls
    Dim a As Integer, b As Integer
    For a = 0 To szWid - 1
        For b = 0 To szHi - 1
            BitBlt gameScreen.hDC, a * 16, b * 16, 16, 16, pics.blocks.hDC, field(a, b) * 16, 0, vbSrcCopy
            If field(a, b) = 1 And hit(a, b) And Not (mines(a, b)) And nums(a, b) > 0 Then
                BitBlt gameScreen.hDC, (a * 16), (b * 16), 16, 16, pics.numbers.hDC, ((nums(a, b) - 1) * 16), 16, vbSrcAnd
                BitBlt gameScreen.hDC, (a * 16), (b * 16), 16, 16, pics.numbers.hDC, ((nums(a, b) - 1) * 16), 0, vbSrcPaint
            End If
        Next b
    Next a
    face.Cls
    BitBlt face.hDC, 0, 0, 26, 26, pics.faces.hDC, currentface * 26, 0, vbSrcCopy
    face.Refresh
    gameScreen.Refresh
End Function


Public Function drawStats()
'this just updates stats like timer and flags
timer.Cls

'timer
Dim ttime As Integer
    ttime = Int(time / 33)
        ttime = ttime - ((Int(ttime / 1000)) * 1000)
        'hundreds
        BitBlt timer.hDC, 0, 0, 11, 21, pics.clock.hDC, ((Int(ttime / 100) Mod 10) * 11), 0, vbSrcCopy
        
        ttime = ttime - ((Int(ttime / 100)) * 100)
        'tens
        BitBlt timer.hDC, 13, 0, 11, 21, pics.clock.hDC, ((Int(ttime / 10) Mod 10) * 11), 0, vbSrcCopy
        
        ttime = ttime - ((Int(ttime / 10)) * 10)
        'ones
        BitBlt timer.hDC, 26, 0, 11, 21, pics.clock.hDC, ((ttime Mod 10) * 11), 0, vbSrcCopy

flags.Cls
'flags left

Dim tflag As Integer
    tflag = Abs(numflags)
        'hundreds
        tflag = tflag - ((Int(tflag / 1000) * 1000))
        If numflags >= 0 Then
            BitBlt flags.hDC, 0, 0, 11, 21, pics.clock.hDC, ((Int(tflag / 100) Mod 10) * 11), 0, vbSrcCopy
        Else
            BitBlt flags.hDC, 0, 0, 11, 21, pics.clock.hDC, (11 * 10), 0, vbSrcCopy
        End If
        'tens
        tflag = tflag - ((Int(tflag / 100)) * 100)
        BitBlt flags.hDC, 13, 0, 11, 21, pics.clock.hDC, ((Int(tflag / 10) Mod 10) * 11), 0, vbSrcCopy
        'ones
        tflag = tflag - ((Int(tflag / 10)) * 10)
        BitBlt flags.hDC, 26, 0, 11, 21, pics.clock.hDC, ((Int(tflag / 1) Mod 10) * 11), 0, vbSrcCopy
End Function

Public Function uncover(tx As Integer, ty As Integer)
    'this uses recursion to uncover unmined blocks with 0 mines surrounding
    'here's a catch during the recursion if player double clicks and has the wrong
    'cell flagged
    If mines(tx, ty) Then
        currentface = 0
        Call endgame(tx, ty)
    Else
        hit(tx, ty) = True
        field(tx, ty) = 1
        If nums(tx, ty) = 0 Then
            If field(tx - 1, ty) = 0 Then
                If tx > 0 And Not (hit(tx - 1, ty)) Then Call uncover(tx - 1, ty)
            End If
            If field(tx, ty - 1) = 0 Then
                If ty > 0 And Not (hit(tx, ty - 1)) Then Call uncover(tx, ty - 1)
            End If
            If field(tx + 1, ty) = 0 Then
                If tx < szWid - 1 And Not (hit(tx + 1, ty)) Then Call uncover(tx + 1, ty)
            End If
            If field(tx, ty + 1) = 0 Then
                If ty < szHi - 1 And Not (hit(tx, ty + 1)) Then Call uncover(tx, ty + 1)
            End If
            
            If field(tx - 1, ty - 1) = 0 Then
                If tx > 0 And ty > 0 And Not (hit(tx - 1, ty - 1)) Then Call uncover(tx - 1, ty - 1)
            End If
            If field(tx + 1, ty - 1) = 0 Then
                If tx < szWid - 1 And ty > 0 And Not (hit(tx + 1, ty - 1)) Then Call uncover(tx + 1, ty - 1)
            End If
            If field(tx - 1, ty + 1) = 0 Then
                If tx > 0 And ty < szHi - 1 And Not (hit(tx - 1, ty + 1)) Then Call uncover(tx - 1, ty + 1)
            End If
            If field(tx + 1, ty + 1) = 0 Then
                If tx < szWid - 1 And ty < szHi - 1 And Not (hit(tx + 1, ty + 1)) Then Call uncover(tx + 1, ty + 1)
            End If
        End If
    End If
End Function

Public Function revealAll(tx As Integer, ty As Integer)
    'this shows where all the mines are at the end of the game
    Dim a As Integer, b As Integer
    For a = 0 To szWid - 1
        For b = 0 To szHi - 1
            If mines(a, b) Then field(a, b) = 2
            If Not (mines(a, b)) And field(a, b) = 5 Then field(a, b) = 4
        Next b
    Next a
    'this is the one the player clicked on
    field(tx, ty) = 3
End Function

Public Function canClear(tx As Integer, ty As Integer) As Boolean
    'this sees if the player has the correct amount of flags to be able to
    'do the double click clear
    Dim flagCount As Integer
    flagCount = 0
    'count up flags on surrounding 8 squares to see if it matches the number
    'double clicked on
    If field(tx - 1, ty - 1) = 5 Then flagCount = flagCount + 1
    If field(tx, ty - 1) = 5 Then flagCount = flagCount + 1
    If field(tx + 1, ty - 1) = 5 Then flagCount = flagCount + 1
    If field(tx - 1, ty) = 5 Then flagCount = flagCount + 1
    If field(tx + 1, ty) = 5 Then flagCount = flagCount + 1
    If field(tx - 1, ty + 1) = 5 Then flagCount = flagCount + 1
    If field(tx, ty + 1) = 5 Then flagCount = flagCount + 1
    If field(tx + 1, ty + 1) = 5 Then flagCount = flagCount + 1
    If flagCount = nums(tx, ty) Then
        canClear = True
    Else
        canClear = False
    End If
End Function

Public Function checkEndGame()
    Dim endthegame As Boolean
    Dim a As Integer, b As Integer
    endthegame = True
    For a = 0 To szWid - 1
        For b = 0 To szHi - 1
            'found an uncovered cell.. the game is not over- exit the loop
            If mines(a, b) = False And hit(a, b) = False Then
                endthegame = False
                a = szWid
                b = szHi
            End If
        Next b
    Next a
    
    If endthegame Then
        'check scores...
        'you win
        currentface = 4
        gameover = True
        isDead = True
        Call drawScreen
        Call loadScores
        'see if there's a new high score
        'if there is, get the name and display high scores
        If beginneropts.Checked Then
            If Int(time / 33) < scores(0) Then
                names(0) = InputBox("You have the fastest time for beginner level. Please type your name:")
                scores(0) = Int(time / 33)
                Call saveScores
                minesweeper.Enabled = False
                Load best
                best.Show
                While minesweeper.Enabled = False
                    DoEvents
                Wend
            End If
        ElseIf intermediateopts.Checked Then
            If Int(time / 33) < scores(1) Then
                names(1) = InputBox("You have the fastest time for intermediate level. Please type your name:")
                scores(1) = Int(time / 33)
                Call saveScores
                minesweeper.Enabled = False
                Load best
                best.Show
                While minesweeper.Enabled = False
                    DoEvents
                Wend
            End If
        ElseIf expertopts.Checked Then
            If Int(time / 33) < scores(2) Then
                names(2) = InputBox("You have the fastest time for expert level. Please type your name:")
                scores(2) = Int(time / 33)
                Call saveScores
                minesweeper.Enabled = False
                Load best
                best.Show
                While minesweeper.Enabled = False
                    DoEvents
                Wend
            End If
        End If
    End If
End Function

Public Function endgame(tx As Integer, ty As Integer)
    'simply ends the game
    field(tx, ty) = 3
    isDead = True
    Call revealAll(tx, ty)
    Call drawScreen
End Function

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub newgametime_Click()
    currentface = 1
    Call makeNew
    Call gameLoop
End Sub

Private Sub quitout_Click()
    End
End Sub
