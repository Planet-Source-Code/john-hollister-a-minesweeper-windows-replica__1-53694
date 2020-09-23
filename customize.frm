VERSION 5.00
Begin VB.Form customize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Field"
   ClientHeight    =   1695
   ClientLeft      =   6270
   ClientTop       =   4170
   ClientWidth     =   3480
   Icon            =   "customize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox widText 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "30"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox hiText 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "16"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox mineText 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "99"
      Top             =   1300
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "&Mines:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1300
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "&Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "&Height:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If hiText.Text > 24 Then hiText.Text = 24
    If hiText.Text < 8 Then hiText.Text = 8
    If widText.Text > 30 Then widText.Text = 30
    If widText.Text < 8 Then widText.Text = 8
    If mineText.Text > (hiText.Text - 1) * (widText.Text - 1) Then
        mineText.Text = (hiText.Text - 1) * (widText.Text - 1)
    End If

    minesweeper.szHi = hiText.Text
    minesweeper.szWid = widText.Text
    minesweeper.numMines = mineText.Text
    minesweeper.Enabled = True
    Unload Me
End Sub

Private Sub Command2_Click()
    minesweeper.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    hiText.Text = minesweeper.szHi
    widText.Text = minesweeper.szWid
    mineText.Text = minesweeper.numMines
End Sub
