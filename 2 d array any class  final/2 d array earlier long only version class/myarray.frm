VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "2d array class"
   ClientHeight    =   7764
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   10104
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   647
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "new"
      Height          =   372
      Index           =   11
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   840
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "cut range"
      Height          =   372
      Index           =   10
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7200
      Width           =   3372
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   15
      Left            =   3960
      TabIndex        =   28
      Text            =   "8"
      Top             =   5400
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   14
      Left            =   2760
      TabIndex        =   26
      Text            =   "3"
      Top             =   5400
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   13
      Left            =   2040
      TabIndex        =   25
      Text            =   "2"
      Top             =   5400
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Set item"
      Height          =   372
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set Range to Value"
      Height          =   372
      Index           =   8
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   3252
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   12
      Left            =   9240
      TabIndex        =   22
      Text            =   "8"
      Top             =   5520
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new lower bound rows"
      Height          =   372
      Index           =   7
      Left            =   6120
      TabIndex        =   21
      Top             =   3360
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   11
      Left            =   9120
      TabIndex        =   20
      Text            =   "8"
      Top             =   3360
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   10
      Left            =   9120
      TabIndex        =   19
      Text            =   "8"
      Top             =   2160
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new lower bound cols"
      Height          =   372
      Index           =   6
      Left            =   6120
      TabIndex        =   18
      Top             =   2760
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   9
      Left            =   9120
      TabIndex        =   17
      Text            =   "8"
      Top             =   2760
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve rows"
      Height          =   372
      Index           =   5
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   8
      Left            =   9120
      TabIndex        =   15
      Text            =   "8"
      Top             =   1680
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve cols"
      Height          =   372
      Index           =   4
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   7
      Left            =   7680
      TabIndex        =   12
      Text            =   "4"
      Top             =   5040
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   6
      Left            =   7680
      TabIndex        =   11
      Text            =   "1"
      Top             =   4560
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   5
      Left            =   9120
      TabIndex        =   10
      Text            =   "4"
      Top             =   4080
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   4
      Left            =   8400
      TabIndex        =   9
      Text            =   "1"
      Top             =   4080
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range horizontal descending"
      Height          =   372
      Index           =   3
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   4212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range horizontal ascending"
      Height          =   372
      Index           =   2
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   4212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range vertical descending"
      Height          =   372
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   4212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range vertical ascending"
      Height          =   372
      Index           =   0
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   3
      Left            =   7680
      TabIndex        =   3
      Text            =   "8"
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   2
      Left            =   7680
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   1
      Left            =   9240
      TabIndex        =   1
      Text            =   "10"
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   0
      Left            =   8280
      TabIndex        =   0
      Text            =   "2"
      Top             =   120
      Width           =   732
   End
   Begin VB.Shape Shape1 
      Height          =   972
      Left            =   0
      Top             =   5160
      Width           =   5292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "row"
      Height          =   252
      Index           =   1
      Left            =   2040
      TabIndex        =   30
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "col"
      Height          =   252
      Index           =   0
      Left            =   2760
      TabIndex        =   29
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   372
      Left            =   3600
      TabIndex        =   27
      Top             =   5520
      Width           =   372
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dimensions:"
      Height          =   372
      Left            =   6120
      TabIndex        =   13
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "range:"
      Height          =   492
      Left            =   6120
      TabIndex        =   8
      Top             =   4440
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARR2D As New Array2d

Private Sub Form_Load()
With ARR2D
.ZeroBaseDimension 5, 5
.setValue 0, 0, 6
.setValue 3, 3, 10
.setValue 2, 1, 7
.setRow 1, 5
.setcol 2, 3
.DoSwap 1, 1, 3, 3
.doprint Me
End With
End Sub

Private Sub Command1_Click(Index As Integer)
With ARR2D
Select Case Index
Case 0
ARR2D.dosort Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), True, True
Case 1
ARR2D.dosort Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), True, False
Case 2
ARR2D.dosort Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), False, True
Case 3
ARR2D.dosort Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), False, False
Case 4
.redimPreserveCols Val(Text1(8))
Case 5
.redimPreserveROWS Val(Text1(10))
Case 6
.resetColBounds Val(Text1(9))

Case 7
.resetRowBounds Val(Text1(11))
Case 8
.setRange Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), Val(Text1(12))

Case 9
.setValue Val(Text1(13)), Val(Text1(14)), Val(Text1(15))
Case 10
.cutArray Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5))
Case 11
.dimension Val(Text1(2)), Val(Text1(3)), Val(Text1(0)), Val(Text1(1))
End Select
ARR2D.doprint Me
End With
End Sub

