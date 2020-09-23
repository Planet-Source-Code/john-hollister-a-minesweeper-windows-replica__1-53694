VERSION 5.00
Begin VB.Form pics 
   Caption         =   "Form2"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   2640
      Picture         =   "pics.frx":0000
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      Top             =   480
      Width           =   720
   End
   Begin VB.PictureBox numbers 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      Picture         =   "pics.frx":3132
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   1800
      Width           =   1920
   End
   Begin VB.PictureBox faces 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   480
      Picture         =   "pics.frx":6174
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   2
      Top             =   1320
      Width           =   1950
   End
   Begin VB.PictureBox clock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   480
      Picture         =   "pics.frx":8986
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.PictureBox blocks 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   480
      Picture         =   "pics.frx":A7A4
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   0
      Top             =   480
      Width           =   1680
   End
End
Attribute VB_Name = "pics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
