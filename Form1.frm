VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   840
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roll Dice"
      Height          =   495
      Left            =   345
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   165
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      Picture         =   "Form1.frx":70C2
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////
'*****************************************
'Bitblt credits goto PSC
'All other source code created by Carlos.
'Graphics created by Carlos (for personnal use ONLY)
'*****************************************
'///////////////////////////////////

Dim Die1, Die2
Dim Die1AnimIndex
Dim D1X, D1Y, OldD1X, OldD1Y
Dim D2X, D2Y, OldD2X, OldD2Y
Dim DiceON  As Boolean
Private Sub Command1_Click()
Die1AnimIndex = 0
TimeAnimDice = 0
Randomize
Die1 = Int((6 * Rnd) + 1)
Die2 = Int((6 * Rnd) + 1)
MakeDiceRoll
End Sub
Private Sub MakeDiceRoll()
If DiceON Then EraseDice
DiceON = True
InitAnimDiceValue
SaveDiceBuffer
RollDice
FirstDouble = False
DoubleDice = False
If Die1 = Die2 Then DoubleDice = True
StopAnimBuck = True
End Sub


Private Sub Form_Load()
Die1AnimIndex = 0
End Sub

Private Sub RollDice()
Do
If Die1AnimIndex = 8 Then
  Die1AnimIndex = 0
  TimeAnimDice = TimeAnimDice + 1
End If
If TimeAnimDice = 2 Then
  EraseDice
  SaveDiceBuffer
  PrintDice
  Exit Sub
End If

D1X = D1X - 4
D1Y = D1Y - 6
D2X = D2X - 2
D2Y = D2Y - 7

Die1AnimIndex = Die1AnimIndex + 1
EraseDice

OldD1X = D1X
OldD1Y = D1Y
OldD2X = D2X
OldD2Y = D2Y

SaveDiceBuffer
AnimateDice
Refresh

DoEvents
Loop
End Sub
Private Sub AnimateDice()
TransparentBlt hdc, D2X, D2Y, 40, 40, Picture1.hdc, Die1AnimIndex * 40, 0, TRANSCOLOR
TransparentBlt hdc, D1X, D1Y, 40, 40, Picture1.hdc, Die1AnimIndex * 40, 0, TRANSCOLOR
End Sub
Private Sub PrintDice()
 TransparentBlt hdc, D1X, D1Y, 40, 40, Picture2.hdc, (Die1 - 1) * 40, 0, TRANSCOLOR
 TransparentBlt hdc, D2X, D2Y, 40, 40, Picture2.hdc, (Die2 - 1) * 40, 0, TRANSCOLOR
 Refresh
End Sub
Private Sub EraseDice()
 BitBlt hdc, OldD2X, OldD2Y, 40, 40, Picture4.hdc, 0, 0, SRCCOPY
 BitBlt hdc, OldD1X, OldD1Y, 40, 40, Picture3.hdc, 0, 0, SRCCOPY
End Sub
Private Sub SaveDiceBuffer()
 BitBlt Picture4.hdc, 0, 0, 40, 40, hdc, D2X, D2Y, SRCCOPY
 BitBlt Picture3.hdc, 0, 0, 40, 40, hdc, D1X, D1Y, SRCCOPY
End Sub


Private Sub InitAnimDiceValue()
TimeAnimDice = 0
D1X = 100: D1Y = 150
OldD1X = 100: OldD1Y = D1Y
D2X = 140: D2Y = D1Y
OldD2X = 140: OldD2Y = D1Y
End Sub
