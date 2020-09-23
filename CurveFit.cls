VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurveFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CurveFitType
   CurveFitLinear = 1      ' y = Ax + B
   CurveFit1 = 2           ' y = A/x + B
   CurveFit2 = 3           ' y = A/[x + B]
   CurveFit3 = 4           ' y = 1/[Ax + B]
   CurveFit4 = 5           ' y = x/[A + Bx]
   CurveFit5 = 6           ' y = A ln(x) + B
   CurveFit6 = 7           ' y = A*exp(Bx)
   CurveFit7 = 8           ' y = A*x^B
   CurveFit8 = 9           ' y = [Ax + B]^(-2)
End Enum
   
Public Sub LeastSquaresFit(x() As Double, y() As Double, FitType As CurveFitType, a As Double, b As Double)
   Dim tX() As Double
   Dim tY() As Double
   Dim i As Long
   Dim tA As Double
   Dim tB As Double
   
   ' Make a temporary copy of x(), y()
   tX = x
   tY = y
   
   ' Transformations to linearize data
   Select Case FitType
      Case CurveFitLinear:
         'Do Nothing
      Case CurveFit1:
         For i = LBound(tX) To UBound(tX)
            tX(i) = 1 / tX(i)
         Next i
         
      Case CurveFit2:
         For i = LBound(tX) To UBound(tX)
            tX(i) = tX(i) * tY(i)
         Next i
         
      Case CurveFit3:
         For i = LBound(tX) To UBound(tX)
            tY(i) = 1 / tY(i)
         Next i
         
      Case CurveFit4:
         For i = LBound(tX) To UBound(tX)
            tX(i) = 1 / tX(i)
            tY(i) = 1 / tY(i)
         Next i
      
      Case CurveFit5:
         For i = LBound(tX) To UBound(tX)
            tX(i) = Log(tX(i))
         Next i
         
      Case CurveFit6:
         For i = LBound(tX) To UBound(tX)
            tY(i) = Log(tY(i))
         Next i
         
      Case CurveFit7:
         For i = LBound(tX) To UBound(tX)
            tX(i) = Log(tX(i))
            tY(i) = Log(tY(i))
         Next i
         
      Case CurveFit8:
         For i = LBound(tX) To UBound(tX)
            tY(i) = tY(i) ^ -0.5
         Next i
   End Select
   
   
   Call LeastSquaresFitLine(tX, tY, tA, tB)
   
   'Transform coefficents if necessary
   Select Case FitType
      Case CurveFit2:
         a = -tB / tA
         b = -1 / tA
         
      Case CurveFit5:
         a = Exp(tB)
         b = tA
         
      Case CurveFit6:
         a = Exp(tB)
         b = tA
         
      Case Else:
         a = tA
         b = tB
   End Select
End Sub

'Curve fit to y = Ax + B
Public Sub LeastSquaresFitLine(x() As Double, y() As Double, a As Double, b As Double)
   Dim Xmean As Double
   Dim Ymean As Double
   Dim i As Long
   Dim N As Double
   Dim SumX As Double
   Dim SumXY As Double
   
   
   N = UBound(x) - LBound(x) + 1
   
   'Find the mean of x
   Xmean = 0
   For i = LBound(x) To UBound(x)
      Xmean = Xmean + x(i)
   Next i
   Xmean = Xmean / N
   
   'Find the mean of y
   Ymean = 0
   For i = LBound(y) To UBound(y)
      Ymean = Ymean + y(i)
   Next i
   Ymean = Ymean / N
   
   'Find Sum(x(i)-Xmean)^2
   SumX = 0
   For i = LBound(x) To UBound(x)
      SumX = SumX + ((x(i) - Xmean) ^ 2)
   Next i
   
   'Find Sum(x(i)-Xmean)(y(i)-Ymean)
   SumXY = 0
   For i = LBound(x) To UBound(x)
      SumXY = SumXY + ((x(i) - Xmean) * (y(i) - Ymean))
   Next i
   
   'Compute the slope
   a = SumXY / SumX
   'Compute the y-intercept
   b = Ymean - (a * Xmean)
End Sub

'Curve fit to P(x) = c(1) + c(2)x + c(3)x^2 + c(4)x^3 + ... c(M+1)x^(M)
Public Function PolynomialCurveFit(x() As Double, y() As Double, M As Long, c() As Double)
   Dim b() As Double
   Dim p() As Double
   Dim a() As Double
   Dim i As Long
   Dim j As Long
   Dim tY As Double
   Dim tX As Double
   Dim tP As Double
   
   ReDim b(1 To M + 1)
   ReDim c(1 To M + 1)
   ReDim p(0 To (2 * M))
   ReDim a(1 To M + 1, 1 To M + 1)
   
   'Compute the column vector
   For i = LBound(x) To UBound(x)
      tY = y(i)
      tX = x(i)
      tP = 1
      
      For j = 1 To M + 1
         b(j) = b(j) + tY * tP
         tP = tP * tX
      Next j
   Next i
   
   
   'Compute the sum of powers
   For i = LBound(x) To UBound(x)
      tX = x(i)
      tP = x(i)
      
      For j = 1 To 2 * M
         p(j) = p(j) + tP
         tP = tP * tX
      Next j
   Next i
   
   
   'Determine the matrix entries
   For i = 1 To M + 1
      For j = 1 To M + 1
         a(i, j) = p(i + j - 2)
      Next j
   Next i
   
   'Solve A*C = B
   Call SolveEqu(a, c, b)
End Function

'Solves A*C = B
Private Sub SolveEqu(a() As Double, c() As Double, b() As Double)
   Dim i As Long
   Dim j As Long
   Dim k As Long
   Dim t As Double
   Dim f As Double
   
   'Call PrintMatrix(a, b)
   
   For i = LBound(a, 2) To UBound(a, 2)
      'Try to swap invalid rows
      If a(i, i) = 0 Then
         For j = i + 1 To UBound(a, 2)
            If a(i, j) <> 0 Then
               For k = LBound(a, 1) To UBound(a, 1)
                  t = a(k, i)
                  a(k, i) = a(k, j)
                  a(k, j) = t
               Next k
               
               t = b(i)
               b(i) = b(j)
               b(j) = t
               
               'Debug.Print "Swap."
               'Call PrintMatrix(a, b)
               Exit For
            End If
         Next j
      End If
      
      If a(i, i) <> 0 Then
         'Normalize
         f = a(i, i)
         For k = LBound(a, 1) To UBound(a, 1)
            a(k, i) = a(k, i) / f
         Next k
         b(i) = b(i) / f
         
         'Debug.Print "Normalize."
         'Call PrintMatrix(a, b)
         
         For j = i + 1 To UBound(a, 2)
            If a(i, j) <> 0 Then
               f = -a(i, j) / a(i, i)
               
               For k = LBound(a, 1) To UBound(a, 1)
                  a(k, j) = a(k, j) + f * a(k, i)
               Next k
               b(j) = b(j) + f * b(i)
            End If
         Next j
         'Debug.Print "Pivot."
         'Call PrintMatrix(a, b)
      End If
   Next i
   
   'Solve up.
   For i = UBound(a, 2) To LBound(a, 2) Step -1
      c(i) = b(i)
      
      For j = i - 1 To LBound(a, 2) Step -1
         b(j) = b(j) - (a(i, j) * b(i))
         a(i, j) = 0!
      Next j
   Next i
   
   'Debug.Print "Solved."
   'Call PrintMatrix(a, b)
      
   'Fill out the result array
   For i = LBound(b) To UBound(b)
      c(i) = b(i)
   Next i
End Sub


Private Sub PrintMatrix(a() As Double, b() As Double)
   Dim i As Long
   Dim j As Long
   
   For i = LBound(b) To UBound(b)
      For j = LBound(a, 2) To UBound(a, 2)
         Debug.Print Right$("      " & Format(a(j, i), "00.0"), 6);
      Next j
      Debug.Print " | " & Right$("      " & Format(b(i), "00.0"), 6)
   Next i
End Sub
