VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Tom Pydeski's Example for Grid Position HighLighting"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Descr 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Descr"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   25
      Cols            =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GridInit As Byte
Dim eTitle$
Dim EMess$
Dim mError As Long
Dim Buttonin As Byte
Dim LastRow As Integer
Dim LastCol As Integer
Dim LastRow1 As Integer
Dim LastCol1 As Integer
Dim LastDRow As Integer
Dim LastDCol As Integer
Dim NewRowIn As Integer
Dim NewColIn As Integer
Dim Rowp As Integer
Dim Colp As Integer
Dim RowIn As Integer
Dim ColIn As Integer
Dim Changingsel As Byte
Dim CellX As Integer
Dim CellY As Integer
Dim GridX As Integer
Dim GridY As Integer
Dim gRow As Integer

Private Sub Form_Load()
Dim r As Integer
Dim c As Integer
For r = 0 To Grid1.Rows - 1
    Grid1.TextMatrix(r, 0) = "Row " & r
Next r
For c = 0 To Grid1.Cols - 1
    Grid1.TextMatrix(0, c) = "Col " & c
Next c
Grid1.TextMatrix(0, 0) = "Row0/Col0"
For r = 1 To Grid1.Rows - 1
    For c = 1 To Grid1.Cols - 1
        Grid1.TextMatrix(r, c) = "R=" & r & ":C=" & c
    Next c
Next r
GridInit = 1
SetWhite
Grid1.Redraw = True
End Sub

Private Sub Grid1_Click()
'Buttonin = 1
'uncomment above if you want to allow selection
Colp = Grid1.Col
Rowp = Grid1.Row
LastCol = Grid1.Col
LastRow = Grid1.Row
ColIn = Grid1.Col
RowIn = Grid1.Row
NewRowIn = RowIn
NewColIn = ColIn
If Grid1.Row = Grid1.RowSel Then MoveDescr
'select row
Grid1.Redraw = False
If GridInit = 0 Then Exit Sub
'
If Grid1.RowSel - Grid1.Row > 1 Then
    GoTo clexit
End If
If Grid1.ColSel - Grid1.Col > 1 Then
    GoTo clexit
End If
SetWhite
Changingsel = 1
Grid1.FillStyle = flexFillRepeat
'
Grid1.Row = RowIn
Grid1.Col = 1
Grid1.RowSel = RowIn
Grid1.ColSel = Grid1.Cols - 1
Grid1.CellBackColor = vbCyan 'Gray
LastCol = Grid1.Col
LastRow = Grid1.Row
Grid1.Row = RowIn
Grid1.Col = ColIn
Changingsel = 0
Me.Refresh
Grid1.Redraw = True
Grid1.FillStyle = flexFillSingle
SameRow:
'
Grid1.SetFocus
clexit:
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Oops
Dim ToolText$
Dim LastDRow As Integer
Dim LastDCol As Integer
CellX = Grid1.MouseCol
CellY = Grid1.MouseRow
If CellY < 1 Then Exit Sub
If CellX < 1 Then Exit Sub
ToolText$ = "Row = " & CellY & " : Col = " & CellX
Grid1.ToolTipText = ToolText$
If Buttonin = 1 Then
    Exit Sub 'grid was clicked
End If
'below moves the description box to the selected cell
If (LastDRow <> CellY Or LastDCol <> CellX) Then ' And Buttonin = 1 Then
    'this would never be executed, so why is it here
    'i remmed out the buttonin so now it can be executed
    'MoveDescr
    Descr.Text = CellY & " , " & CellX
    Descr.Width = Me.TextWidth(Descr.Text) + 100
    Descr.Top = (Grid1.RowPos(CellY - 1) + Grid1.Top) - 100
    Descr.Left = (Grid1.ColPos(CellX) + Grid1.Left)
    If CellY > 0 Then
        Descr.Visible = True
    Else
        Descr.Visible = False
    End If
    Descr.ZOrder 0
    LastDRow = CellY
    LastDCol = CellX
    Refresh
    DoEvents
End If
skipold:
SetRow
mend:
GoTo Exit_Grid1_MouseMove
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_MouseMove "
EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in Grid1_MouseMove"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_MouseMove:
End Sub

Sub SetRow()
If GridInit = 0 Then Exit Sub
If Grid1.MouseRow = LastRow1 And Grid1.MouseCol = LastCol1 Then
    'Caption = "same " & Grid1.MouseRow & " " & LastRow
    GoTo SameRow
End If
'Caption = "change"
Grid1.Redraw = False
Changingsel = 1
SetWhite
Changingsel = 1
'
Grid1.FillStyle = flexFillRepeat
If RowIn < 1 Then RowIn = 1
If CellY < 1 Then CellY = 1
If CellX < 1 Then CellX = 1
Grid1.Row = CellY
Grid1.Col = 1
Grid1.RowSel = CellY
Grid1.ColSel = Grid1.Cols - 1
Grid1.CellBackColor = vbCyan 'Gray
'
Grid1.Row = 1
Grid1.Col = CellX
Grid1.RowSel = CellY
Grid1.ColSel = CellX
Grid1.CellBackColor = vbCyan ' Gray
'
Grid1.Row = CellY
Grid1.Col = CellX
LastCol = Grid1.Col
LastRow = Grid1.Row
LastRow1 = CellY
LastCol1 = CellX
Changingsel = 0
Grid1.Redraw = True
SameRow:
End Sub

Sub SetWhite()
Dim t As Integer
'sets the grid cell backcolors to alternate white and grayish
Changingsel = 1
Grid1.FillStyle = flexFillRepeat
Grid1.Redraw = False
For t = 1 To Grid1.Rows - 1
    Grid1.Row = t
    Grid1.Col = 1
    Grid1.RowSel = t ' Grid1.Rows - 1
    Grid1.ColSel = Grid1.Cols - 1
    If t Mod 2 = 0 Then
        Grid1.CellBackColor = &H9FFFC0
    Else
        Grid1.CellBackColor = vbWhite
    End If
Next t
Grid1.Row = 1
Grid1.Col = 1
Grid1.RowSel = 1
Grid1.ColSel = 1
Grid1.FillStyle = flexFillSingle
Changingsel = 0
End Sub

Sub MoveDescr()
On Error GoTo Oops
If Descr.Text = "" Then Exit Sub
GridX = Grid1.ColPos(Grid1.Col)
gRow = Grid1.Row
If gRow < 0 Then
    gRow = 0
    GoTo mend
End If
GridY = Grid1.Top + Grid1.RowPos(gRow)
Descr.Text = Grid1.Row & " , " & Grid1.Col
Descr.Width = Me.TextWidth(Descr.Text) + 100 ' (Len(Descr.Text) * 100)
Descr.Top = (GridY - Descr.Height) - 25
'Debug.Print Descr.Top
Descr.Left = GridX
Descr.Visible = True
Descr.ZOrder 0
Refresh
DoEvents
LastDRow = CellY
LastDCol = CellX
GoTo mend
Oops:
EMess$ = "Error # " & Str$(Err) & " - " & Error$
mError = MsgBox(EMess$, 2, "movedescr") 'ABORT=3,RETRY=4,IGNORE=5
If mError = 4 Then Resume
If mError = 5 Then Resume Next
mend:
Me.Refresh
Changingsel = 0
End Sub

