VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Rotation Transformations Around Axies"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll10 
      Height          =   225
      Left            =   6630
      Max             =   100
      Min             =   -100
      TabIndex        =   13
      Top             =   8040
      Value           =   10
      Width           =   3045
   End
   Begin VB.HScrollBar HScroll9 
      Height          =   225
      Left            =   6630
      Max             =   100
      Min             =   -100
      TabIndex        =   12
      Top             =   7755
      Value           =   15
      Width           =   3045
   End
   Begin VB.HScrollBar HScroll8 
      Height          =   225
      Left            =   6630
      Max             =   100
      Min             =   -100
      TabIndex        =   11
      Top             =   7485
      Value           =   25
      Width           =   3045
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   225
      Left            =   3450
      Max             =   100
      Min             =   -100
      TabIndex        =   10
      Top             =   7500
      Width           =   3105
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   225
      Left            =   3450
      Max             =   100
      Min             =   -100
      TabIndex        =   9
      Top             =   7770
      Width           =   3105
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   225
      Left            =   3450
      Max             =   100
      Min             =   -100
      TabIndex        =   8
      Top             =   8055
      Width           =   3105
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   225
      Left            =   270
      Max             =   360
      TabIndex        =   6
      Top             =   8325
      Value           =   45
      Width           =   9405
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   225
      Left            =   270
      Max             =   360
      TabIndex        =   4
      Top             =   8055
      Width           =   3105
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   225
      Left            =   270
      Max             =   360
      TabIndex        =   2
      Top             =   7770
      Width           =   3105
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   270
      Max             =   360
      TabIndex        =   0
      Top             =   7500
      Width           =   3105
   End
   Begin VB.Label Label5 
      Caption         =   "Scale"
      Height          =   195
      Index           =   2
      Left            =   8115
      TabIndex        =   16
      Top             =   7230
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "Aligment"
      Height          =   195
      Index           =   1
      Left            =   4785
      TabIndex        =   15
      Top             =   7245
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Rotation"
      Height          =   195
      Index           =   0
      Left            =   1635
      TabIndex        =   14
      Top             =   7275
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "P"
      Height          =   210
      Left            =   45
      TabIndex        =   7
      Top             =   8325
      Width           =   150
   End
   Begin VB.Label Label3 
      Caption         =   "X"
      Height          =   210
      Left            =   45
      TabIndex        =   5
      Top             =   8055
      Width           =   150
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      Height          =   210
      Left            =   45
      TabIndex        =   3
      Top             =   7770
      Width           =   150
   End
   Begin VB.Label Label1 
      Caption         =   "Z"
      Height          =   210
      Left            =   45
      TabIndex        =   1
      Top             =   7500
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The idea is simple:
'Use the 3 matrices for rotation around the axies and change X,Y,Z loc to rotated ones
'Use the method of adding an amount of Z to x and an amount of Z to X
'See Html For more information
Option Explicit
Const cx = 320
Const cy = 240
Const pi = 3.14159265358979
Dim VP(2) As Integer
Dim Xc&(7), Yc&(7), Zc&(7)
Sub ZRotate(ByRef X() As Long, ByRef Y() As Long, ByVal T#)
   Dim X1&, i&
   For i = LBound(X) To UBound(X)
     X1 = X(i)
     X(i) = X1 * Cos(T) - Y(i) * Sin(T)
     Y(i) = X1 * Sin(T) + Y(i) * Cos(T)
   Next i
End Sub

Sub YRotate(ByRef X() As Long, ByRef Z() As Long, ByVal T#)
   Dim X1&, i&
   For i = LBound(X) To UBound(X)
     X1 = X(i)
     X(i) = X1 * Cos(T) + Z(i) * Sin(T)
     Z(i) = -X1 * Sin(T) + Z(i) * Cos(T)
   Next i
End Sub
Sub XRotate(ByRef Y() As Long, ByRef Z() As Long, ByVal T#)
   Dim Y1&, i&
   For i = LBound(Y) To UBound(Y)
     Y1 = Y(i)
     Y(i) = Y1 * Cos(T) - Z(i) * Sin(T)
     Z(i) = Y1 * Sin(T) + Z(i) * Cos(T)
   Next i
End Sub

Private Sub Form_Load()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Change()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
   Dim T#, i&, P#, C1&, C2&, C3&, T1&, T2&, T3&, W&, H&, D&
   P = pi * HScroll4.Value / 180 'Z axis prespective
   Me.Cls
   Me.DrawWidth = 1
   'Locate which axis was rotated and paint it blue
   If HScroll1.Value <> VP(0) Then C1& = vbBlue
   If HScroll2.Value <> VP(1) Then C2& = vbBlue
   If HScroll3.Value <> VP(2) Then C3& = vbBlue
   VP(0) = HScroll1.Value: VP(1) = HScroll2.Value: VP(2) = HScroll3.Value
   'Draw the 3 axes
   Me.Line (0, cy)-(2 * cx, cy), C3
   Me.Line (cx, 0)-(cx, 2 * cy), C2
   Me.Line (cx - Cos(P) * 200, cy + Sin(P) * 200)-(cx + Cos(P) * 200, cy - Sin(P) * 200), C1
   Me.DrawWidth = 2
   'Get the left/top/bottom transposations
   T1 = HScroll5.Value
   T2 = HScroll6.Value
   T3 = HScroll7.Value
   'Get the Width/Height/Debth
   W = HScroll8.Value
   H = HScroll9.Value
   D = HScroll10.Value
   'The X,Y,Z points of the cube
   Xc(0) = -W + T1: Yc(0) = -H + T2: Zc(0) = D + T3
   Xc(1) = W + T1:  Yc(1) = -H + T2: Zc(1) = D + T3
   Xc(2) = W + T1:  Yc(2) = -H + T2: Zc(2) = -D + T3
   Xc(3) = -W + T1: Yc(3) = -H + T2: Zc(3) = -D + T3
   Xc(4) = -W + T1: Yc(4) = H + T2:  Zc(4) = D + T3
   Xc(5) = W + T1:  Yc(5) = H + T2:  Zc(5) = D + T3
   Xc(6) = W + T1:  Yc(6) = H + T2:  Zc(6) = -D + T3
   Xc(7) = -W + T1: Yc(7) = H + T2:  Zc(7) = -D + T3
   'Call the 3 subs to rotate the cube on the 3 axes
   ZRotate Xc(), Yc(), pi * HScroll1.Value / 180 'rotate the point co-ordinates to angle t
   YRotate Xc(), Zc(), pi * HScroll2.Value / 180 'rotate the point co-ordinates to angle t
   XRotate Yc(), Zc(), pi * HScroll3.Value / 180 'rotate the point co-ordinates to angle t
   'Connect the points with a line
   Lne 0, 1, P
   Lne 1, 2, P
   Lne 2, 3, P
   Lne 3, 0, P
   
   Lne 0, 4, P
   Lne 1, 5, P
   Lne 2, 6, P
   Lne 3, 7, P
   
   Lne 4, 5, P
   Lne 5, 6, P
   Lne 6, 7, P
   Lne 7, 4, P
   
   Me.Refresh
End Sub
Function Lne(P1, P2, P#)
    'Draws a line given X1,Y1,Z1,X2,Y2,Z3 at a Z prespective P
    'Input: Point1 Index,Point2 Index, Prespective
    Me.Line (cx + Xc(P1) - Zc(P1) * Cos(P), cy - Yc(P1) + Zc(P1) * Sin(P))-(cx + Xc(P2) - Zc(P2) * Cos(P), cy - Yc(P2) + Zc(P2) * Sin(P)), vbRed
End Function

Private Sub HScroll2_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll2_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll3_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll3_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll4_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll4_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll5_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll5_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll6_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll6_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll7_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll7_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll8_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll8_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll9_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll9_Scroll(): HScroll1_Scroll: End Sub
Private Sub HScroll10_Change(): HScroll1_Scroll: End Sub
Private Sub HScroll10_Scroll(): HScroll1_Scroll: End Sub
