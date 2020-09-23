VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim cf As CurveFit
   Dim x(1 To 4) As Double
   Dim y(1 To 4) As Double
   Dim c() As Double
   Dim i As Long
   
   
   Set cf = New CurveFit
   
   x(1) = 0: y(1) = 0
   x(2) = 1: y(2) = 2
   x(3) = 2: y(3) = 2
   x(4) = 3: y(4) = 0
   
   Print "Points:"
   For i = 1 To 4
      Print "  x=" & CStr(x(i)) & ", y=" & CStr(y(i))
   Next i
   
   Print
   
   Print "Calculating polynomial fit to: c(1) + c(2)x + c(3)x^2"
   
   Call cf.PolynomialCurveFit(x, y, 2, c)
   
   For i = LBound(c) To UBound(c)
      Print "   c(" & CStr(i) & ")=" & CStr(c(i))
   Next i
End Sub
