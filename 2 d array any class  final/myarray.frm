VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "2d array class"
   ClientHeight    =   8220
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   10104
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Reset to saved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   21
      Left            =   2760
      TabIndex        =   44
      Top             =   7560
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Array"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   20
      Left            =   600
      TabIndex        =   43
      Top             =   7560
      Width           =   1932
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      ItemData        =   "myarray.frx":0000
      Left            =   2760
      List            =   "myarray.frx":0010
      TabIndex        =   41
      Top             =   6960
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve Ubound Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   15
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3360
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   19
      Left            =   9120
      TabIndex        =   39
      Text            =   "8"
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve Ubound Cols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   14
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3840
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   18
      Left            =   9120
      TabIndex        =   37
      Text            =   "8"
      Top             =   3840
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve Lbound Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   13
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2400
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   17
      Left            =   9120
      TabIndex        =   35
      Text            =   "8"
      Top             =   2400
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "redim preserve Lbound Cols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   12
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2880
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   16
      Left            =   9120
      TabIndex        =   33
      Text            =   "-1"
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   11
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "cut range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   10
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6240
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   15
      Left            =   4200
      TabIndex        =   28
      Text            =   "8"
      Top             =   5160
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   14
      Left            =   3000
      TabIndex        =   26
      Text            =   "3"
      Top             =   5160
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   13
      Left            =   2280
      TabIndex        =   25
      Text            =   "2"
      Top             =   5160
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Set item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set Range to Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   3252
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   840
      TabIndex        =   21
      Top             =   4200
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   11
      Left            =   3720
      TabIndex        =   20
      Text            =   "8"
      Top             =   4200
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   10
      Left            =   9120
      TabIndex        =   19
      Text            =   "8"
      Top             =   1800
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new lower bound cols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   840
      TabIndex        =   18
      Top             =   3720
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   9
      Left            =   3720
      TabIndex        =   17
      Text            =   "8"
      Top             =   3720
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "redim preserve number rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   8
      Left            =   9120
      TabIndex        =   15
      Text            =   "8"
      Top             =   1320
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "redim preserve number cols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   7
      Left            =   7680
      TabIndex        =   12
      Text            =   "4"
      Top             =   5040
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   6
      Left            =   7680
      TabIndex        =   11
      Text            =   "1"
      Top             =   4560
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   5
      Left            =   9120
      TabIndex        =   10
      Text            =   "4"
      Top             =   4440
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   4
      Left            =   8400
      TabIndex        =   9
      Text            =   "1"
      Top             =   4440
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range horizontal descending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   4812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range horizontal ascending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   4812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range vertical descending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   4812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sort range vertical ascending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   4812
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   7560
      TabIndex        =   3
      Text            =   "8"
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   7560
      TabIndex        =   2
      Text            =   "0"
      Top             =   240
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   9000
      TabIndex        =   1
      Text            =   "10"
      Top             =   0
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   8280
      TabIndex        =   0
      Text            =   "2"
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Vartype"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   480
      TabIndex        =   42
      Top             =   6960
      Width           =   2172
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   972
      Left            =   240
      Top             =   4920
      Width           =   4812
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   2280
      TabIndex        =   30
      Top             =   4920
      Width           =   852
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "col"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3120
      TabIndex        =   29
      Top             =   4920
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      TabIndex        =   27
      Top             =   5280
      Width           =   372
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dimensions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6000
      TabIndex        =   13
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "range:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6600
      TabIndex        =   8
      Top             =   4920
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************optional user type
'don't need this but make life easier to store an array
'from the class and its dimensions for later use or for transfer
'to another instance of the class
'could be made public and put in a module
Private Type TwoDArrayType
tarray As Variant
tRLBound As Long
tRUBound As Long
tCLBound As Long
tCUBound As Long
tHASARRAY As Boolean
tVARTYPE As Integer
End Type
Dim savedarray As TwoDArrayType
'************

Dim ARR2D As New Array2d '1) Declare an instance of the class

Private Sub Form_Load()

With ARR2D
.setvartype vbInteger '2) set type of array
.ZeroBaseDimension 5, 5 '3) set some dimensions
'and 4)... that's it, we're off...

.SETALL 5
.setValue 0, 0, 5
.setValue 3, 3, 9
.setValue 2, 1, 9
.setRow 1, 8
.setcol 2, 6
.DoSwap 1, 1, 3, 3

Me.Caption = "vartype " & .getvartype

.doprint Me, 65
End With
End Sub

Private Sub Command1_Click(Index As Integer) 'just for demonstration
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
If .getvartype = vbString Then
.setRange (Text1(6)), (Text1(4)), (Text1(7)), (Text1(5)), (Text1(12))
Else
.setRange Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5)), Val(Text1(12))
End If
Case 9
If .getvartype = vbString Then
.setValue (Text1(13)), (Text1(14)), (Text1(15))

Else
.setValue Val(Text1(13)), Val(Text1(14)), Val(Text1(15))
End If
Case 10
.cutArray Val(Text1(6)), Val(Text1(4)), Val(Text1(7)), Val(Text1(5))
Case 11
.dimension Val(Text1(2)), Val(Text1(3)), Val(Text1(0)), Val(Text1(1))
Case 12
.redimPreserveByLowerColBound Val(Text1(16))
Case 13
.redimPreserveByLowerRowBound Val(Text1(17))
Case 14
.redimPreserveByUpperColBound Val(Text1(18))
Case 15
.redimPreserveByUpperRowBound Val(Text1(19))
Case 20

With savedarray 'save an error
.tarray = ARR2D.FetchArray(.tRLBound, .tCLBound, .tRUBound, .tCUBound, .tVARTYPE)
.tHASARRAY = True
End With
Case 21
With savedarray 'reset class to saved array
If .tHASARRAY Then
ARR2D.SetArray .tRLBound, .tCLBound, .tRUBound, .tCUBound, .tarray
End If
End With

End Select
ARR2D.doprint Me, 65
End With
End Sub

Private Sub Combo1_Click()
'will throw up errors if arrays can't be coerced
With ARR2D
.setvartype Combo1.ItemData(Combo1.ListIndex)
.doprint Me, 65
Me.Caption = "vartype " & .getvartype
End With
End Sub

