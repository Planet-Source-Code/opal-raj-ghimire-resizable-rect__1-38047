VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rectangle Class Demo"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   ForeColor       =   &H00000000&
   Icon            =   "oDraw.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      DrawWidth       =   4
      Height          =   4080
      Left            =   60
      ScaleHeight     =   268
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   0
      Top             =   75
      Width           =   1215
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "E&xit"
         Height          =   780
         Left            =   105
         TabIndex        =   4
         Top             =   3165
         Width           =   960
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&How to"
         Height          =   780
         Left            =   105
         TabIndex        =   5
         Top             =   2400
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete All"
         Height          =   780
         Left            =   105
         Picture         =   "oDraw.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1635
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         Height          =   780
         Left            =   105
         Picture         =   "oDraw.frx":010E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   870
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Draw"
         Height          =   780
         Left            =   105
         Picture         =   "oDraw.frx":0210
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   105
         Width           =   960
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ikey As Integer
Dim SelRect As Integer 'Selected Rectangle's ID
Dim Mode As String
Dim CNEW As Boolean 'Create New Rectangle
Dim Rectangles As New Collection
Dim Bdx As Long, Bdy As Long
Dim oBdx As Long, oBdy As Long 'o=old
Dim oX As Long, oY As Long
Dim iHandle As Integer 'Selectd sizing handle

Private Sub Command1_Click()
Unload Me
End
End Sub


Private Sub Command2_Click()
Dim N As Rectangle
Dim id As Integer
For Each N In Rectangles
    If N.Selected Then
        id = N.Count
        N.DeleteRect
        Rectangles.Remove (Str(id))
        SelRect = 0
Exit For
    End If 'N.Selected
Next
Set N = Nothing
End Sub

Private Sub Command3_Click()
'Deletes all
Dim N As Rectangle
Dim id As Integer
Form1.DrawMode = 10
For Each N In Rectangles
       id = N.Count
       N.DeleteRect
       Rectangles.Remove (Str(id))
Next
Form1.DrawMode = 13
Cls
Ikey = 0
SelRect = 0
Set N = Nothing
End Sub


Private Sub Command4_Click()
Dim Mes As String
Mes = Mes + "Click 'Draw' and drag mouse on the form to draw a new rectangle." + vbCrLf + vbCrLf
Mes = Mes + "Right mouse down inside rectangle to select." + vbCrLf + vbCrLf
Mes = Mes + "Left mouse down inside rectangle and drag to drag." + vbCrLf + vbCrLf
Mes = Mes + "Left mouse down at sizing handles and drag to resize." + vbCrLf + vbCrLf
Mes = Mes + vbCrLf + vbCrLf + "Hope you like this"


MsgBox Mes, , "Info"
End Sub




Private Sub Form_Load()
Mode = ""
Picture1.Line (5, 5)-(Picture1.ScaleWidth - 5, Picture1.ScaleHeight - 5), &HC0C0C0, B

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Bdx = x: Bdy = y: oX = x: oY = y: oBdx = x: oBdy = y
Dim M As Rectangle
Dim id As Integer
Mode = ""
If Button = 2 Then 'Select only one
    DeSelectAll

        For Each M In Rectangles
            If M.IsPtInsideRect(CLng(x), CLng(y)) Then
                SelRect = M.Count
                M.ShowHideHandles
        Exit For
            End If 'M.IsPtInsideRect(CLng(X), CLng(Y))
        Next
Exit Sub

End If 'Button=2

If SelRect = 0 Or CNEW Then Exit Sub
Mode = ""

If Rectangles(Str(SelRect)).IsPtInsideRect(CLng(x), CLng(y)) Then Mode = "Drag": Exit Sub


If Rectangles(Str(SelRect)).FindSelectedCornor(CLng(x), CLng(y)) <> -1 Then
    iHandle = Rectangles(Str(SelRect)).FindSelectedCornor(CLng(x), CLng(y))
    Mode = "Resize"
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DrawMode = 10
Dim id As Integer
'
If CNEW And Button = 1 Then 'Create New Rect
    Line (oBdx, oBdy)-(oX, oY), vbBlack, B
    oBdx = Bdx: oBdy = Bdy
    oX = x: oY = y
    Line (Bdx, Bdy)-(x, y), vbBlack, B
End If 'cnew
'----------------------------------
If Button = 1 Then
    If Mode = "Resize" Then
        DrawMode = 10
        Rectangles(Str(SelRect)).DeleteRect
        Rectangles(Str(SelRect)).ResizeRect iHandle, CLng(x), CLng(y)
        Rectangles(Str(SelRect)).ShowHideHandles
        DrawMode = 13
    End If 'Mode = "Resize"
 
    If Mode = "Drag" Then
        DrawMode = 10
        Rectangles(Str(SelRect)).DeleteRect
        Rectangles(Str(SelRect)).DragRect Bdx, Bdy, x, y
        Bdx = x: Bdy = y 'Delete this line and see what happens!!
        Rectangles(Str(SelRect)).ShowHideHandles
        DrawMode = 13
    End If 'Mode="Drag"

End If 'button=1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DrawMode = 13
If CNEW And Button = 1 Then
    Dim oRect As Rectangle
    Set oRect = New Rectangle

    oRect.DrawRect Bdx, Bdy, x, y
    oRect.Count = Ikey
    Rectangles.Add oRect, Str(Ikey)
    MousePointer = 0
    CNEW = False
    Option1.Value = False
    Set oRect = Nothing
End If 'CNEW And Button = 1

End Sub


Private Sub DeSelectAll()
Dim N As Rectangle
For Each N In Rectangles
    If N.Selected Then N.ShowHideHandles
Next
Set N = Nothing
End Sub

Private Sub Option1_Click()
CNEW = True
Ikey = Ikey + 1
MousePointer = 2
DeSelectAll
SelRect = 0
End Sub
