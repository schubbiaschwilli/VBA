Attribute VB_Name = "mdl_pnorms"
Option Explicit

Function pnorm_Hart_West(x As Double) As Double
    Dim XAbs As Double
    Dim build As Double
    Dim Exponential As Double
    Dim Cumnorm As Double
    
    XAbs = Abs(x)
    If XAbs > 37 Then
            Cumnorm = 0
        Else
            Exponential = Exp(-XAbs ^ 2 / 2)
            
            If XAbs < 7.07106781186547 Then
                    build = 3.52624965998911E-02 * XAbs + 0.700383064443688
                    build = build * XAbs + 6.37396220353165
                    build = build * XAbs + 33.912866078383
                    build = build * XAbs + 112.079291497871
                    build = build * XAbs + 221.213596169931
                    build = build * XAbs + 220.206867912376
                    Cumnorm = Exponential * build
                    build = 8.83883476483184E-02 * XAbs + 1.75566716318264
                    build = build * XAbs + 16.064177579207
                    build = build * XAbs + 86.7807322029461
                    build = build * XAbs + 296.564248779674
                    build = build * XAbs + 637.333633378831
                    build = build * XAbs + 793.826512519948
                    build = build * XAbs + 440.413735824752
                    Cumnorm = Cumnorm / build
               Else
                    build = XAbs + 0.65
                    build = XAbs + 4 / build
                    build = XAbs + 3 / build
                    build = XAbs + 2 / build
                    build = XAbs + 1 / build
                    Cumnorm = Exponential / build / 2.506628274631
            End If
        End If
        
        If x > 0 Then Cumnorm = 1 - Cumnorm
        
        pnorm_Hart_West = Cumnorm

End Function

Function pnorm_Hastings(x As Double) As Double
    ' Paul Glassermann, Monte Carlo Methods in Financial Engineering, 2003, p. 69
    ' Milton Abramowitz and Irene A. Stegun: Handbook of Mathematical Functions (With Formulas, Graphs, and Mathematical Tables); 10th Printing, 1972; Formula 26.2.17 p. 932
    Const b1 As Double = 0.31938153
    Const b2 As Double = -0.356563782
    Const b3 As Double = 1.781477937
    Const b4 As Double = -1.821255978
    Const b5 As Double = 1.330274429
    Const p As Double = 0.2316419
    Dim c As Double
    Dim t As Double
    Dim S As Double
    Dim y As Double
    
    c = Log(Sqr(2 * 3.14159265358979))
  
    t = 1 / (1 + Abs(x) * p)
    S = ((((b5 * t + b4) * t + b3) * t + b2) * t + b1) * t
    y = S * Exp(-0.5 * x * x - c)
    If (x > 0) Then y = 1 - y
    pnorm_Hastings = y
  
End Function

Function pnorm_Marsaglia(x As Double) As Double
    ' Paul Glassermann, Monte Carlo Methods in Financial Engineering, 2003, p. 70
    Dim v(15) As Double
    v(1) = 1.2533141373155          ' 1.253314137315500
    v(2) = 0.655679542418799        ' 0.6556795424187985
    v(3) = 0.421369229288055        ' 0.4213692292880545
    v(4) = 0.304590298710103        ' 0.3045902987101033
    v(5) = 0.236652382913561        ' 0.2366523829135607
    v(6) = 0.192808104715316        ' 0.1928081047153158
    v(7) = 0.162377660896868        ' 0.1623776608968675
    v(8) = 0.14010418345305         ' 0.1401041834530502
    v(9) = 0.123131963257933        ' 0.1231319632579329
    v(10) = 0.109787282578308       ' 0.1097872825783083
    v(11) = 9.90285964717319E-02    ' 0.09902859647173193
    v(12) = 9.01756755010647E-02    ' 0.09017567550106468
    v(13) = 8.27662865013692E-02    ' 0.08276628650136917
    v(14) = 7.64757610162485E-02    ' 0.0764757610162485
    v(15) = 7.10695805388521E-02    ' 0.07106958053885211
    
    Dim c As Double
    c = Log(Sqr(2 * 3.14159265358979))
    
    Dim a As Double
    Dim b As Double
    Dim h As Double
    Dim i As Integer
    Dim j As Double
    Dim q As Double
    Dim S As Double
    Dim y As Double
    Dim z As Double
    
    j = Int(min(Abs(x) + 0.5, 14))
    z = j
    h = Abs(x) - z
    a = v(j + 1)
    b = z * a - 1
    q = 1
    S = a + h * b
  
    For i = 2 To 24 - j Step 2
        a = (a + z * b) / i
        b = (b + z * a) / (i + 1)
        q = q * h * h
        S = S + q * (a + h * b)
    Next
    
    y = S * Exp(-0.5 * x * x - c)
    If (x > 0) Then y = 1 - y
  
    pnorm_Marsaglia = y
    
End Function

Function min(x1, x2)
    min = (x1 + x2 - Abs(x1 - x2)) / 2
End Function
