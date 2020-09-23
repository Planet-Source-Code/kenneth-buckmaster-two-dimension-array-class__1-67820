VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Array2d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'2-d array class based on 1-d array - more flexible than VB's 2-d

'after a suggestion from La Volpe I've generalised this class
'to work with any type of variable
'array is stored in a variant and values are passed as variant
'the earlier version could perhaps be used as a template for
'specific length variables like longs

'Kenneth Buckmaster
Dim UBarray As Long
Dim arr As Variant, rows As Long, cols As Long
Dim RLBound As Long, RUBound As Long, CLBound As Long, CUBound As Long
'for sort
Dim arrstart As Long, colw As Long, sortvertical As Boolean, sortAscending As Boolean
Dim vartypeset As Boolean
Private Sub Class_Initialize()
UBarray = -1
End Sub
'choice of two versions of setting vartype
' pass as array of the type
Sub setvartype(var)
arr = var
vartypeset = True
End Sub
'or use a select case
Sub setvartype2(ByVal var As Integer)
Select Case var 'calling this will be slightly easier but this sub
'won't look pretty for all the variable types possible
Case vbInteger
Dim ii() As Integer
arr = ii
Case vbLong
Dim ll() As Long
arr = ll
Case vbString
Dim ss() As String
arr = ss
Case vbDate
Dim dd() As Date
arr = dd
'case etc
End Select
vartypeset = True
End Sub

Sub ZeroBaseDimension(ByVal newrows, ByVal newcols)
If newrows <= 0 Or newcols <= 0 Then Exit Sub
rows = newrows
cols = newcols
UBarray = ((rows) * (cols)) - 1
ReDim arr(UBarray)
RLBound = 0
CLBound = 0
RUBound = rows - 1
CUBound = cols - 1
End Sub

Sub dimension(LBRows As Long, UBRows As Long, LBcols As Long, UBCols As Long)
If UBRows < LBRows Or UBCols < LBcols Then Exit Sub

rows = UBRows - LBRows + 1
cols = UBCols - LBcols + 1
UBarray = ((rows) * (cols)) - 1
ReDim arr(UBarray)
RLBound = LBRows
CLBound = LBcols
RUBound = UBRows
CUBound = UBCols
End Sub

Sub resetRowBounds(RowLbound As Long)
RUBound = RowLbound + rows - 1
RLBound = RowLbound
End Sub
Sub resetColBounds(ColLbound As Long)

CUBound = ColLbound + cols - 1
CLBound = ColLbound

End Sub

Sub redimPreserveCols(newcols As Long)

Dim tmpclbound As Long, tmprlbound As Long, oldcols As Long, rowvals As Long, i As Long, postarget As Long, possource As Long
Dim tmp As Variant

If newcols <= 0 Then Exit Sub
tmpclbound = CLBound
tmprlbound = RLBound

tmp = arr

Dim j As Long
oldcols = cols

If newcols < oldcols Then rowvals = newcols Else rowvals = oldcols

ZeroBaseDimension rows, newcols

For i = 0 To rows - 1

For j = 0 To rowvals - 1
arr(postarget + j) = tmp(possource + j)
Next

postarget = postarget + newcols
possource = possource + oldcols
Next
cols = newcols
CLBound = tmpclbound
CUBound = CLBound + cols - 1
RLBound = tmprlbound
RUBound = RLBound + rows - 1


End Sub

Sub redimPreserveROWS(ByVal newrows As Long)
If newrows <= 0 Then Exit Sub
UBarray = ((newrows) * (cols)) - 1
ReDim Preserve arr(UBarray)
rows = newrows
RUBound = RLBound + rows - 1
End Sub
Sub redimPreserve(ByVal newrows As Long, ByVal newcols As Long)
If newrows <= 0 Or newcols <= 0 Then Exit Sub
redimPreserveCols newcols
redimPreserveROWS newrows
End Sub
Sub redimPreserveByBounds(ByVal NewRUBound As Long, ByVal NewCUBound As Long)
redimPreserveCols NewCUBound - CLBound + 1
redimPreserveROWS NewRUBound - RLBound + 1
End Sub

Sub setValue(ByVal row As Long, ByVal col As Long, VALUE As Variant)
If checkBounds(row, col) = False Then Exit Sub
arr(((row - RLBound) * cols) + col - CLBound) = VALUE
End Sub
Function getValue(ByVal row As Long, ByVal col As Long) As Variant
If checkBounds(row, col) = False Then Exit Function
getValue = arr(((row - RLBound) * cols) + col - CLBound)
End Function

Sub setRow(ByVal rowNo As Long, VALUE As Variant)
If checkBounds(rowNo, CLBound) = False Then Exit Sub
Dim i As Long
For i = CLBound To CUBound
setValue rowNo, i, VALUE
Next
End Sub
Sub setcol(ByVal ColNo As Long, VALUE As Variant)
If checkBounds(RLBound, ColNo) = False Then Exit Sub
Dim i As Long
For i = RLBound To RUBound
setValue i, ColNo, VALUE
Next

End Sub
Sub setRange(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long, VALUE As Variant)
If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub
Dim i As Long, j As Long
For i = row1 To row2
For j = col1 To col2
setValue i, j, VALUE
Next
Next
End Sub
Sub SETALL(VALUE As Variant)
Dim i As Long
For i = 0 To UBarray
arr(i) = VALUE
Next
End Sub

Sub DoSwap(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long)
If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub
Dim tmp As Variant
tmp = getValue(row1, col1)
setValue row1, col1, getValue(row2, col2)
setValue row2, col2, tmp
End Sub
Sub cutArray(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long)
If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub
Dim j As Long
Dim tmp, newrows As Long, newcols As Long, rowvals As Long, possource As Long, postarget As Long, i As Long
newrows = row2 - row1 + 1
newcols = col2 - col1 + 1

tmp = arr
ReDim tmp(((newrows) * (newcols)) - 1)
rowvals = newcols

possource = row1 - RLBound + ((col1 - CLBound) * cols)

For i = 0 To newrows - 1

For j = 0 To rowvals - 1
tmp(postarget + j) = arr(possource + j)
Next

postarget = postarget + newcols
possource = possource + cols
Next
cols = newcols
rows = newrows
arr = tmp

resetRowBounds row1
resetColBounds col1
End Sub

Private Function checkBounds(ByVal row As Long, ByVal col As Long) As Boolean
If Not vartypeset Or UBarray = -1 Or row < RLBound Or row > RUBound Or col < CLBound Or col > CUBound Then Exit Function
checkBounds = True
End Function

'** getbounds
Function getRLBound() As Long
getRLBound = RLBound
End Function
Function getRUBound() As Long
getRUBound = RUBound
End Function
Function getCLBound() As Long
getCLBound = CLBound
End Function
Function getCUBound() As Long
getCUBound = CUBound
End Function
'** end getbounds

'** get getvartype
Function getvartype() As Long
getvartype = VarType(arr(0))
End Function


'*** maths

Sub ChangeRange(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long, VALUE As Variant, SwapValue As Variant)

If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub
Dim i As Long, j As Long
For i = row1 To row2
For j = col1 To col2
If getValue(i, j) = VALUE Then setValue i, j, SwapValue
Next
Next

End Sub
Sub IncrementRange(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long, increment)

If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub
Dim i As Long, j As Long
For i = row1 To row2
For j = col1 To col2
setValue i, j, getValue(i, j) + increment
Next
Next

End Sub
Function sumRange(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long) As Variant
If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Function
Dim i As Long, j As Long
For i = row1 To row2
For j = col1 To col2
sumRange = sumRange + getValue(i, j)
Next
Next
End Function
'*** end maths


'*******SORT
Sub dosort(ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long, Optional Sort_Vertical_True_or_Horizontal = False, Optional Sort_Ascending_True_or_Descending = True)

Dim ROWH As Long

If checkBounds(row1, col1) = False Or checkBounds(row2, col2) = False Then Exit Sub

arrstart = row1 - RLBound + ((col1 - CLBound) * cols)
ROWH = row2 - row1 + 1
colw = col2 - col1 + 1

sortvertical = Sort_Vertical_True_or_Horizontal
sortAscending = Sort_Ascending_True_or_Descending

 InsertionSort 0, (ROWH * colw) - 1
End Sub
'modified insertionsort original from Philippe Lord http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
Private Sub InsertionSort(ByVal vmin As Long, ByVal vmax As Long)
   Dim i  As Long, j As Long, tmp, arrj As Long, arrjMinus1 As Long
   
   For i = vmin + 1 To vmax
      tmp = arr(getarr(i))
      j = i
      
      Do While j > vmin
         
         arrjMinus1 = getarr(j - 1)
         
If sortAscending Then '
If arr(arrjMinus1) <= tmp Then Exit Do
Else '
If arr(arrjMinus1) >= tmp Then Exit Do '
End If '

        arr(getarr(j)) = arr(arrjMinus1)
         j = j - 1
      Loop
      
       arr(getarr(j)) = tmp
   Next i
End Sub

Private Function getarr(ByVal pos As Long) As Long
If sortvertical Then
getarr = arrstart + ((pos Mod colw) * cols) + ((pos \ colw))
Else
getarr = arrstart + (pos Mod colw) + ((pos \ colw) * cols)
End If
End Function
'******* END SORT

'******* print
Sub doprint(OBJ As Object)
Dim colwidth As Long, c As Long, i As Long, j As Long
colwidth = 65
OBJ.Cls
OBJ.ForeColor = vbBlue
For i = CLBound To CUBound
OBJ.CurrentX = 30 + (c * colwidth)
c = c + 1

OBJ.Print i;
Next
OBJ.Print ""
c = 0
For i = RLBound To RUBound
OBJ.ForeColor = vbBlue
OBJ.Print i;
OBJ.ForeColor = vbBlack
For j = CLBound To CUBound
OBJ.CurrentX = 30 + (c * colwidth)
c = c + 1
OBJ.Print getValue(i, j);
Next
c = 0
OBJ.Print ""

Next
End Sub
'**end print
