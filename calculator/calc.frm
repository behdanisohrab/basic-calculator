VERSION 5.00
Begin VB.Form calc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bss calculator"
   ClientHeight    =   4860
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4860
   LinkTopic       =   "calc"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton taghsim 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   3825
      TabIndex        =   17
      Top             =   1170
      Width           =   915
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   15
      Left            =   3840
      TabIndex        =   16
      Top             =   2655
      Width           =   915
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   14
      Left            =   3855
      TabIndex        =   15
      Top             =   2145
      Width           =   915
   End
   Begin VB.CommandButton zarb 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   13
      Left            =   3840
      TabIndex        =   14
      Top             =   1665
      Width           =   915
   End
   Begin VB.CommandButton buttondot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   12
      Left            =   1380
      TabIndex        =   13
      Top             =   3375
      Width           =   915
   End
   Begin VB.CommandButton button8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   11
      Left            =   1395
      TabIndex        =   12
      Top             =   1305
      Width           =   915
   End
   Begin VB.CommandButton button9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   10
      Left            =   2475
      TabIndex        =   11
      Top             =   1305
      Width           =   915
   End
   Begin VB.CommandButton button4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   9
      Left            =   300
      TabIndex        =   10
      Top             =   2025
      Width           =   915
   End
   Begin VB.CommandButton button5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   8
      Left            =   1410
      TabIndex        =   9
      Top             =   2025
      Width           =   915
   End
   Begin VB.CommandButton button6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   7
      Left            =   2505
      TabIndex        =   8
      Top             =   2010
      Width           =   915
   End
   Begin VB.CommandButton button1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   285
      TabIndex        =   7
      Top             =   2700
      Width           =   915
   End
   Begin VB.CommandButton button2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   1380
      TabIndex        =   6
      Top             =   2670
      Width           =   915
   End
   Begin VB.CommandButton button3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   2655
      Width           =   915
   End
   Begin VB.CommandButton clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   270
      TabIndex        =   4
      Top             =   4020
      Width           =   2955
   End
   Begin VB.CommandButton button0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   3345
      Width           =   915
   End
   Begin VB.CommandButton equal 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   3990
      TabIndex        =   2
      Top             =   3195
      Width           =   645
   End
   Begin VB.CommandButton button7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   1305
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   45
      TabIndex        =   0
      Top             =   165
      Width           =   4770
   End
End
Attribute VB_Name = "calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnum As Integer
Dim snum As Integer
Dim sign As String

Private Sub button0_Click(Index As Integer)
Text1.Text = Text1.Text & 0

End Sub

Private Sub button1_Click(Index As Integer)
Text1.Text = Text1.Text & 1
End Sub

Private Sub button2_Click(Index As Integer)
Text1.Text = Text1.Text & 2
End Sub

Private Sub button3_Click(Index As Integer)
Text1.Text = Text1.Text & 3
End Sub

Private Sub button4_Click(Index As Integer)
Text1.Text = Text1.Text & 4
End Sub

Private Sub button5_Click(Index As Integer)
Text1.Text = Text1.Text & 5
End Sub

Private Sub button6_Click(Index As Integer)
Text1.Text = Text1.Text & 6
End Sub

Private Sub button7_Click(Index As Integer)
Text1.Text = Text1.Text & 7
End Sub

Private Sub button8_Click(Index As Integer)
Text1.Text = Text1.Text & 8
End Sub

Private Sub button9_Click(Index As Integer)
Text1.Text = Text1.Text & 9
End Sub

Private Sub buttondot_Click(Index As Integer)
Text1.Text = Text1.Text & "."
End Sub

Private Sub clear_Click(Index As Integer)
Text1.Text = ""
End Sub

Private Sub equal_Click(Index As Integer)
snum = Text1.Text
If sign = "+" Then
Text1.Text = fnum + snum
ElseIf sign = "/" Then
Text1.Text = fnum / fnum
ElseIf sign = "x" Then
Text1.Text = fnum * snum
ElseIf sign = "-" Then
Text1.Text = fnum - snum
End If
End Sub

Private Sub minus_Click(Index As Integer)
fnum = Text1.Text
sign = "-"
Text1.Text = ""
End Sub

Private Sub plus_Click(Index As Integer)
fnum = Text1.Text
sign = "+"
Text1.Text = ""
End Sub

Private Sub taghsim_Click(Index As Integer)
fnum = Text1.Text
sign = "/"
Text1.Text = ""
End Sub

Private Sub zarb_Click(Index As Integer)
fnum = Text1.Text
sign = "x"
Text1.Text = ""
End Sub
