VERSION 5.00
Begin VB.Form best 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Best Times"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "best.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset Scores"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fastest Mine Sweepers"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label nameE 
         Caption         =   "Label4"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1240
         Width           =   1335
      End
      Begin VB.Label nameI 
         Caption         =   "Label4"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   760
         Width           =   1335
      End
      Begin VB.Label nameB 
         Caption         =   "Label4"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   280
         Width           =   1335
      End
      Begin VB.Label timeE 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1240
         Width           =   1215
      End
      Begin VB.Label timeI 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   760
         Width           =   1215
      End
      Begin VB.Label timeB 
         Caption         =   "999 Seconds"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Expert:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Intermediate:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Beginner:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   280
         Width           =   1095
      End
   End
End
Attribute VB_Name = "best"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For a = 0 To 2
        scores(a) = 999
        names(a) = "Anonymous"
    Next a
    Call saveScores
    timeB.Caption = scores(0)
    timeI.Caption = scores(1)
    timeE.Caption = scores(2)
    nameB.Caption = names(0)
    nameI.Caption = names(1)
    nameE.Caption = names(2)
End Sub

Private Sub Command2_Click()
    minesweeper.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call loadScores
    timeB.Caption = scores(0)
    timeI.Caption = scores(1)
    timeE.Caption = scores(2)
    nameB.Caption = names(0)
    nameI.Caption = names(1)
    nameE.Caption = names(2)
End Sub

Private Sub Form_Terminate()
    minesweeper.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    minesweeper.Enabled = True
End Sub
