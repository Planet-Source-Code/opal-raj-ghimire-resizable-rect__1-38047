VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Rectangle class: draw, resize, delete and drag the rectangles
'Opal Raj Ghimire

'buna48@hotmail.com
'http://geocities.com/opalraj/vb

'No matter what they say
'I did this in this way


Option Explicit
Private cCount As Integer
Private cColor As Long
Private bSelected As Boolean
Private iSize As Integer
Private Cornors(3) As POINTS
Private Type POINTS
        x  As Integer
        y  As Integer
End Type
'DrawRect draws rectangle and updates cornors array
Public Sub DrawRect(ByVal x As Long, ByVal y As Long, ByVal Dx As Long, ByVal Dy As Long)
Dim RectHeight As Long
Dim RectWidth As Long
Dim TmpOne As Long, TmpTwo As Long

'x,y should always be smaller
If Dx < x And Dy < y Then
'swaping the values
TmpOne = Dx: TmpTwo = Dy
Dx = x: Dy = y
x = TmpOne: y = TmpTwo
End If

If x < Dx And y > Dy Then
'swaping the values
TmpOne = y
y = Dy
Dy = TmpOne
End If

If x > Dx And y < Dy Then
'swaping the values
TmpOne = x
x = Dx
Dx = TmpOne
End If



RectHeight = Dy - y: RectWidth = Dx - x
'Update cornors array
Cornors(0).x = x: Cornors(0).y = y
Cornors(1).x = x + RectWidth: Cornors(1).y = y
Cornors(2).x = Dx: Cornors(2).y = Dy
Cornors(3).x = x: Cornors(3).y = y + RectHeight
Form1.Line (x, y)-(Dx, Dy), cColor, B
End Sub

Public Function FindSelectedCornor(ByVal x As Long, ByVal y As Long) As Integer
'if given point (x and y) is in any sizing handle
'the return is number of the sizing handle
'Left Top =0 Right Top=1
'Left bottom=2 Right bottom=3
'otherwise -1
Dim iK As Integer
FindSelectedCornor = -1
           For iK = 0 To 3 '*
            If x >= Cornors(iK).x - iSize And x <= Cornors(iK).x + iSize And _
                y >= Cornors(iK).y - iSize And y <= Cornors(iK).y + iSize Then
                FindSelectedCornor = iK
            End If        '*
        Next iK
End Function

Public Function IsPtInsideRect(ByVal x As Long, ByVal y As Long) As Boolean
'Returns True if given point(x and y) is inside the rectangle
If x >= Cornors(0).x And x <= Cornors(2).x And y >= Cornors(0).y And y <= Cornors(2).y Then IsPtInsideRect = True
End Function
Public Sub ShowHideHandles()
'First time it shows sizing handles
'calling it again hides the sizing handles

Dim iTmpFS As Integer    ' Fill Style
Dim ItmpFC As Long        ' Fill Color
Dim iTmpDM As Long      ' Draw Mode
'backing up current properties
iTmpFS = Form1.FillStyle
ItmpFC = Form1.FillColor
iTmpDM = Form1.DrawMode
'Applying new properties
Form1.FillStyle = vbSolid
Form1.FillColor = vbRed
Form1.DrawMode = vbNotXorPen
'Drawing the sizing handles
Form1.Line (Cornors(0).x, Cornors(0).y)-(Cornors(0).x - iSize, Cornors(0).y - iSize), , B
Form1.Line (Cornors(1).x, Cornors(1).y)-(Cornors(1).x + iSize, Cornors(1).y - iSize), , B
Form1.Line (Cornors(2).x, Cornors(2).y)-(Cornors(2).x + iSize, Cornors(2).y + iSize), , B
Form1.Line (Cornors(3).x, Cornors(3).y)-(Cornors(3).x - iSize, Cornors(3).y + iSize), , B
'Reapplying backed up properties
Form1.DrawMode = iTmpDM
Form1.FillStyle = iTmpFS
Form1.FillColor = ItmpFC
'Changing the selected flag
bSelected = Not bSelected

End Sub
Public Property Get Selected() As Boolean
'flag to notify selection (Read Only Property)
Selected = bSelected
End Property

Private Sub Class_Initialize()
iSize = 6 'Sizing handles size
bSelected = False 'Selected
cColor = vbBlack  'Line color
End Sub


'ResizeRect It resizes the rect,
'Index= Numerical value of sizing handle, X,Y = mousedown point on sizing handle
'See Function FindSelectedCornor
Public Sub ResizeRect(Index As Integer, x As Long, y As Long)
Dim iA As Long, iB As Long

Select Case Index
Case 0 'Left top
iA = Cornors(2).x
iB = Cornors(2).y
'Don't let the rectangles flip
If x > iA Then x = Cornors(2).x - iSize
If y > iB Then y = Cornors(2).y - iSize
DrawRect x, y, Cornors(2).x, Cornors(2).y

Case 2 'Right bottom
iA = Cornors(0).x
iB = Cornors(0).y
If x < iA Then x = Cornors(0).x + iSize
If y < iB Then y = Cornors(0).y + iSize
DrawRect Cornors(0).x, Cornors(0).y, x, y

Case 1 'Top Right
iA = Cornors(3).x
iB = Cornors(3).y
If x < iA Then x = Cornors(3).x + iSize
If y > iB Then y = Cornors(3).y - iSize
DrawRect Cornors(3).x, y, x, Cornors(3).y

Case 3 'Left Bottom

iA = Cornors(1).x
iB = Cornors(1).y
If x > iA Then x = Cornors(1).x - iSize
If y < iB Then y = Cornors(1).y + iSize
DrawRect x, Cornors(0).y, Cornors(1).x, y

End Select
End Sub


Public Sub DeleteRect()
'Actually Redraws the rectangles
'forms DrawMode=10 makes the rectangles disappear
DrawRect Cornors(0).x, Cornors(0).y, Cornors(2).x, Cornors(2).y
If Selected Then ShowHideHandles 'Turns off sizing handles

End Sub
Public Sub DragRect(ByVal x As Long, ByVal y As Long, ByVal NewX As Long, ByVal NewY As Long)
'X and Y are point inside rect clicked to drag(Mouse down event)
'NewX and NewY are new points where new drawing be made(MouseUp event)

'X and Y distance to Left and top of Rect
Dim UpX As Long
Dim UpY As Long

'X and Y distance to Right and bottom of rect
Dim DnX As Long
Dim DnY As Long

'calcualting distance
UpX = x - Cornors(0).x: UpY = y - Cornors(0).y
DnX = Cornors(2).x - x: DnY = Cornors(2).y - y

'Drawing the new rect
DrawRect NewX - UpX, NewY - UpY, NewX + DnX, NewY + DnY

End Sub


Public Property Get Count() As Integer
Count = cCount
End Property

Public Property Let Count(ByVal NewValue As Integer)
cCount = NewValue
End Property
