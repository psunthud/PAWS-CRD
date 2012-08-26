Module Stat
    Public Const MachineEpsilon As Double = 0.0000000000000005
    Public Const MaxRealNumber As Double = 1.0E+300
    Public Const MinRealNumber As Double = 1.0E-300


    Public Function StudentTDistribution(ByVal k As Integer, ByVal t As Double) As Double
        Dim Result As Double
        Dim x As Double
        Dim rk As Double
        Dim z As Double
        Dim f As Double
        Dim tz As Double
        Dim p As Double
        Dim xsqk As Double
        Dim j As Double

        If t = 0.0# Then
            Result = 0.5
            StudentTDistribution = Result
            Exit Function
        End If
        If t < -2.0# Then
            rk = k
            z = rk / (rk + t * t)
            Result = 0.5 * IncompleteBeta(0.5 * rk, 0.5, z)
            StudentTDistribution = Result
            Exit Function
        End If
        If t < 0.0# Then
            x = -t
        Else
            x = t
        End If
        rk = k
        z = 1.0# + x * x / rk
        If k Mod 2.0# <> 0.0# Then
            xsqk = x / Math.Sqrt(rk)
            p = Math.Atan(xsqk)
            If k > 1.0# Then
                f = 1.0#
                tz = 1.0#
                j = 3.0#
                Do While j <= k - 2.0# And tz / f > MachineEpsilon
                    tz = tz * ((j - 1.0#) / (z * j))
                    f = f + tz
                    j = j + 2.0#
                Loop
                p = p + f * xsqk / z
            End If
            p = p * 2.0# / Math.PI
        Else
            f = 1.0#
            tz = 1.0#
            j = 2.0#
            Do While j <= k - 2.0# And tz / f > MachineEpsilon
                tz = tz * ((j - 1.0#) / (z * j))
                f = f + tz
                j = j + 2.0#
            Loop
            p = f * x / Math.Sqrt(z * rk)
        End If
        If t < 0.0# Then
            p = -p
        End If
        Result = 0.5 + 0.5 * p

        StudentTDistribution = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Functional inverse of Student's t distribution
    '
    'Given probability p, finds the argument t such that stdtr(k,t)
    'is equal to p.
    '
    'ACCURACY:
    '
    'Tested at random 1 <= k <= 100.  The "domain" refers to p:
    '                     Relative error:
    'arithmetic   domain     # trials      peak         rms
    '   IEEE    .001,.999     25000       5.7e-15     8.0e-16
    '   IEEE    10^-6,.001    25000       2.0e-12     2.9e-14
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1995, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InvStudentTDistribution(ByVal k As Long, _
             ByVal p As Double) As Double
        Dim Result As Double
        Dim t As Double
        Dim rk As Double
        Dim z As Double
        Dim rflg As Double

        rk = k
        If p > 0.25 And p < 0.75 Then
            If p = 0.5 Then
                Result = 0.0#
                InvStudentTDistribution = Result
                Exit Function
            End If
            z = 1.0# - 2.0# * p
            z = InvIncompleteBeta(0.5, 0.5 * rk, Math.Abs(z))
            t = Math.Sqrt(rk * z / (1.0# - z))
            If p < 0.5 Then
                t = -t
            End If
            Result = t
            InvStudentTDistribution = Result
            Exit Function
        End If
        rflg = -1.0#
        If p >= 0.5 Then
            p = 1.0# - p
            rflg = 1.0#
        End If
        z = InvIncompleteBeta(0.5 * rk, 0.5, 2.0# * p)
        If MaxRealNumber * z < rk Then
            Result = rflg * MaxRealNumber
            InvStudentTDistribution = Result
            Exit Function
        End If
        t = Math.Sqrt(rk / z - rk)
        Result = rflg * t

        InvStudentTDistribution = Result
    End Function

    Public Function IncompleteBeta(ByVal a As Double, _
         ByVal b As Double, _
         ByVal x As Double) As Double
        Dim Result As Double
        Dim t As Double
        Dim xc As Double
        Dim w As Double
        Dim y As Double
        Dim flag As Double
        Dim sg As Double
        Dim big As Double
        Dim biginv As Double
        Dim MAXGAM As Double
        Dim MINLOG As Double
        Dim MAXLOG As Double

        big = 4.5035996273705E+15
        biginv = 0.000000000000000222044604925031
        MAXGAM = 171.624376956303
        MINLOG = Math.Log(MinRealNumber)
        MAXLOG = Math.Log(MaxRealNumber)
        If x = 0.0# Then
            Result = 0.0#
            IncompleteBeta = Result
            Exit Function
        End If
        If x = 1.0# Then
            Result = 1.0#
            IncompleteBeta = Result
            Exit Function
        End If
        flag = 0.0#
        If b * x <= 1.0# And x <= 0.95 Then
            Result = IncompleteBetaPS(a, b, x, MAXGAM)
            IncompleteBeta = Result
            Exit Function
        End If
        w = 1.0# - x
        If x > a / (a + b) Then
            flag = 1.0#
            t = a
            a = b
            b = t
            xc = x
            x = w
        Else
            xc = w
        End If
        If flag = 1.0# And b * x <= 1.0# And x <= 0.95 Then
            t = IncompleteBetaPS(a, b, x, MAXGAM)
            If t <= MACHinEePsilon Then
                Result = 1.0# - MACHinEePsilon
            Else
                Result = 1.0# - t
            End If
            IncompleteBeta = Result
            Exit Function
        End If
        y = x * (a + b - 2.0#) - (a - 1.0#)
        If y < 0.0# Then
            w = IncompleteBetaFE(a, b, x, big, biginv)
        Else
            w = IncompleteBetaFE2(a, b, x, big, biginv) / xc
        End If
        y = a * Math.Log(x)
        t = b * Math.Log(xc)
        If a + b < MAXGAM And Math.Abs(y) < MAXLOG And Math.Abs(t) < MAXLOG Then
            t = Math.Pow(xc, b)
            t = t * Math.Pow(x, a)
            t = t / a
            t = t * w
            t = t * (Gamma(a + b) / (Gamma(a) * Gamma(b)))
            If flag = 1.0# Then
                If t <= MachineEpsilon Then
                    Result = 1.0# - MachineEpsilon
                Else
                    Result = 1.0# - t
                End If
            Else
                Result = t
            End If
            IncompleteBeta = Result
            Exit Function
        End If
        y = y + t + LnGamma(a + b, sg) - LnGamma(a, sg) - LnGamma(b, sg)
        y = y + Math.Log(w / a)
        If y < MINLOG Then
            t = 0.0#
        Else
            t = Math.Exp(y)
        End If
        If flag = 1.0# Then
            If t <= MACHinEePsilon Then
                t = 1.0# - MACHinEePsilon
            Else
                t = 1.0# - t
            End If
        End If
        Result = t

        IncompleteBeta = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Inverse of imcomplete beta integral
    '
    'Given y, the function finds x such that
    '
    ' incbet( a, b, x ) = y .
    '
    'The routine performs interval halving or Newton iterations to find the
    'root of incbet(a,b,x) - y = 0.
    '
    '
    'ACCURACY:
    '
    '                     Relative error:
    '               x     a,b
    'arithmetic   domain  domain  # trials    peak       rms
    '   IEEE      0,1    .5,10000   50000    5.8e-12   1.3e-13
    '   IEEE      0,1   .25,100    100000    1.8e-13   3.9e-15
    '   IEEE      0,1     0,5       50000    1.1e-12   5.5e-15
    'With a and b constrained to half-integer or integer values:
    '   IEEE      0,1    .5,10000   50000    5.8e-12   1.1e-13
    '   IEEE      0,1    .5,100    100000    1.7e-14   7.9e-16
    'With a = .5, b constrained to half-integer or integer values:
    '   IEEE      0,1    .5,10000   10000    8.3e-11   1.0e-11
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1996, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InvIncompleteBeta(ByVal a As Double, _
             ByVal b As Double, _
             ByVal y As Double) As Double
        Dim Result As Double
        Dim aaa As Double
        Dim bbb As Double
        Dim y0 As Double
        Dim d As Double
        Dim yyy As Double
        Dim x As Double
        Dim x0 As Double
        Dim x1 As Double
        Dim lgm As Double
        Dim yp As Double
        Dim di As Double
        Dim dithresh As Double
        Dim yl As Double
        Dim yh As Double
        Dim xt As Double
        Dim i As Double
        Dim rflg As Double
        Dim dir As Double
        Dim nflg As Double
        Dim s As Double
        Dim MainLoopPos As Double
        Dim ihalve As Double
        Dim ihalvecycle As Double
        Dim newt As Double
        Dim newtcycle As Double
        Dim breaknewtcycle As Double
        Dim breakihalvecycle As Double

        i = 0.0#
        If y = 0.0# Then
            Result = 0.0#
            InvIncompleteBeta = Result
            Exit Function
        End If
        If y = 1.0# Then
            Result = 1.0#
            InvIncompleteBeta = Result
            Exit Function
        End If
        x0 = 0.0#
        yl = 0.0#
        x1 = 1.0#
        yh = 1.0#
        nflg = 0.0#
        MainLoopPos = 0.0#
        ihalve = 1.0#
        ihalvecycle = 2.0#
        newt = 3.0#
        newtcycle = 4.0#
        breaknewtcycle = 5.0#
        breakihalvecycle = 6.0#
        Do While True

            '
            ' start
            '
            If MainLoopPos = 0.0# Then
                If a <= 1.0# Or b <= 1.0# Then
                    dithresh = 0.000001
                    rflg = 0.0#
                    aaa = a
                    bbb = b
                    y0 = y
                    x = aaa / (aaa + bbb)
                    yyy = IncompleteBeta(aaa, bbb, x)
                    MainLoopPos = ihalve
                    GoTo Cont_1
                Else
                    dithresh = 0.0001
                End If
                yp = -InvNormalDistribution(y)
                If y > 0.5 Then
                    rflg = 1.0#
                    aaa = b
                    bbb = a
                    y0 = 1.0# - y
                    yp = -yp
                Else
                    rflg = 0.0#
                    aaa = a
                    bbb = b
                    y0 = y
                End If
                lgm = (yp * yp - 3.0#) / 6.0#
                x = 2.0# / (1.0# / (2.0# * aaa - 1.0#) + 1.0# / (2.0# * bbb - 1.0#))
                d = yp * Math.Sqrt(x + lgm) / x - (1.0# / (2.0# * bbb - 1.0#) - 1.0# / (2.0# * aaa - 1.0#)) * (lgm + 5.0# / 6.0# - 2.0# / (3.0# * x))
                d = 2.0# * d
                If d < Math.Log(MinRealNumber) Then
                    x = 0.0#
                    Exit Do
                End If
                x = aaa / (aaa + bbb * Math.Exp(d))
                yyy = IncompleteBeta(aaa, bbb, x)
                yp = (yyy - y0) / y0
                If Math.Abs(yp) < 0.2 Then
                    MainLoopPos = newt
                    GoTo Cont_1
                End If
                MainLoopPos = ihalve
                GoTo Cont_1
            End If

            '
            ' ihalve
            '
            If MainLoopPos = ihalve Then
                dir = 0.0#
                di = 0.5
                i = 0.0#
                MainLoopPos = ihalvecycle
                GoTo Cont_1
            End If

            '
            ' ihalvecycle
            '
            If MainLoopPos = ihalvecycle Then
                If i <= 99.0# Then
                    If i <> 0.0# Then
                        x = x0 + di * (x1 - x0)
                        If x = 1.0# Then
                            x = 1.0# - MACHinEePsilon
                        End If
                        If x = 0.0# Then
                            di = 0.5
                            x = x0 + di * (x1 - x0)
                            If x = 0.0# Then
                                Exit Do
                            End If
                        End If
                        yyy = IncompleteBeta(aaa, bbb, x)
                        yp = (x1 - x0) / (x1 + x0)
                        If Math.Abs(yp) < dithresh Then
                            MainLoopPos = newt
                            GoTo Cont_1
                        End If
                        yp = (yyy - y0) / y0
                        If Math.Abs(yp) < dithresh Then
                            MainLoopPos = newt
                            GoTo Cont_1
                        End If
                    End If
                    If yyy < y0 Then
                        x0 = x
                        yl = yyy
                        If dir < 0.0# Then
                            dir = 0.0#
                            di = 0.5
                        Else
                            If dir > 3.0# Then
                                di = 1.0# - (1.0# - di) * (1.0# - di)
                            Else
                                If dir > 1.0# Then
                                    di = 0.5 * di + 0.5
                                Else
                                    di = (y0 - yyy) / (yh - yl)
                                End If
                            End If
                        End If
                        dir = dir + 1.0#
                        If x0 > 0.75 Then
                            If rflg = 1.0# Then
                                rflg = 0.0#
                                aaa = a
                                bbb = b
                                y0 = y
                            Else
                                rflg = 1.0#
                                aaa = b
                                bbb = a
                                y0 = 1.0# - y
                            End If
                            x = 1.0# - x
                            yyy = IncompleteBeta(aaa, bbb, x)
                            x0 = 0.0#
                            yl = 0.0#
                            x1 = 1.0#
                            yh = 1.0#
                            MainLoopPos = ihalve
                            GoTo Cont_1
                        End If
                    Else
                        x1 = x
                        If rflg = 1.0# And x1 < MACHinEePsilon Then
                            x = 0.0#
                            Exit Do
                        End If
                        yh = yyy
                        If dir > 0.0# Then
                            dir = 0.0#
                            di = 0.5
                        Else
                            If dir < -3.0# Then
                                di = di * di
                            Else
                                If dir < -1.0# Then
                                    di = 0.5 * di
                                Else
                                    di = (yyy - y0) / (yh - yl)
                                End If
                            End If
                        End If
                        dir = dir - 1.0#
                    End If
                    i = i + 1.0#
                    MainLoopPos = ihalvecycle
                    GoTo Cont_1
                Else
                    MainLoopPos = breakihalvecycle
                    GoTo Cont_1
                End If
            End If

            '
            ' breakihalvecycle
            '
            If MainLoopPos = breakihalvecycle Then
                If x0 >= 1.0# Then
                    x = 1.0# - MACHinEePsilon
                    Exit Do
                End If
                If x <= 0.0# Then
                    x = 0.0#
                    Exit Do
                End If
                MainLoopPos = newt
                GoTo Cont_1
            End If

            '
            ' newt
            '
            If MainLoopPos = newt Then
                If nflg <> 0.0# Then
                    Exit Do
                End If
                nflg = 1.0#
                lgm = LnGamma(aaa + bbb, s) - LnGamma(aaa, s) - LnGamma(bbb, s)
                i = 0.0#
                MainLoopPos = newtcycle
                GoTo Cont_1
            End If

            '
            ' newtcycle
            '
            If MainLoopPos = newtcycle Then
                If i <= 7.0# Then
                    If i <> 0.0# Then
                        yyy = IncompleteBeta(aaa, bbb, x)
                    End If
                    If yyy < yl Then
                        x = x0
                        yyy = yl
                    Else
                        If yyy > yh Then
                            x = x1
                            yyy = yh
                        Else
                            If yyy < y0 Then
                                x0 = x
                                yl = yyy
                            Else
                                x1 = x
                                yh = yyy
                            End If
                        End If
                    End If
                    If x = 1.0# Or x = 0.0# Then
                        MainLoopPos = breaknewtcycle
                        GoTo Cont_1
                    End If
                    d = (aaa - 1.0#) * Math.Log(x) + (bbb - 1.0#) * Math.Log(1.0# - x) + lgm
                    If d < Math.Log(MinRealNumber) Then
                        Exit Do
                    End If
                    If d > Math.Log(MaxRealNumber) Then
                        MainLoopPos = breaknewtcycle
                        GoTo Cont_1
                    End If
                    d = Math.Exp(d)
                    d = (yyy - y0) / d
                    xt = x - d
                    If xt <= x0 Then
                        yyy = (x - x0) / (x1 - x0)
                        xt = x0 + 0.5 * yyy * (x - x0)
                        If xt <= 0.0# Then
                            MainLoopPos = breaknewtcycle
                            GoTo Cont_1
                        End If
                    End If
                    If xt >= x1 Then
                        yyy = (x1 - x) / (x1 - x0)
                        xt = x1 - 0.5 * yyy * (x1 - x)
                        If xt >= 1.0# Then
                            MainLoopPos = breaknewtcycle
                            GoTo Cont_1
                        End If
                    End If
                    x = xt
                    If Math.Abs(d / x) < 128.0# * MachineEpsilon Then
                        Exit Do
                    End If
                    i = i + 1.0#
                    MainLoopPos = newtcycle
                    GoTo Cont_1
                Else
                    MainLoopPos = breaknewtcycle
                    GoTo Cont_1
                End If
            End If

            '
            ' breaknewtcycle
            '
            If MainLoopPos = breaknewtcycle Then
                dithresh = 256.0# * MACHinEePsilon
                MainLoopPos = ihalve
                GoTo Cont_1
            End If
Cont_1:
        Loop

        '
        ' done
        '
        If rflg <> 0.0# Then
            If x <= MACHinEePsilon Then
                x = 1.0# - MACHinEePsilon
            Else
                x = 1.0# - x
            End If
        End If
        Result = x

        InvIncompleteBeta = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Continued fraction expansion #1 for incomplete beta integral
    '
    'Cephes Math Library, Release 2.8:  June, 2000
    'Copyright 1984, 1995, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function IncompleteBetaFE(ByVal a As Double, _
             ByVal b As Double, _
             ByVal x As Double, _
             ByVal big As Double, _
             ByVal biginv As Double) As Double
        Dim Result As Double
        Dim xk As Double
        Dim pk As Double
        Dim pkm1 As Double
        Dim pkm2 As Double
        Dim qk As Double
        Dim qkm1 As Double
        Dim qkm2 As Double
        Dim k1 As Double
        Dim k2 As Double
        Dim k3 As Double
        Dim k4 As Double
        Dim k5 As Double
        Dim k6 As Double
        Dim k7 As Double
        Dim k8 As Double
        Dim r As Double
        Dim t As Double
        Dim ans As Double
        Dim thresh As Double
        Dim n As Double

        k1 = a
        k2 = a + b
        k3 = a
        k4 = a + 1.0#
        k5 = 1.0#
        k6 = b - 1.0#
        k7 = k4
        k8 = a + 2.0#
        pkm2 = 0.0#
        qkm2 = 1.0#
        pkm1 = 1.0#
        qkm1 = 1.0#
        ans = 1.0#
        r = 1.0#
        n = 0.0#
        thresh = 3.0# * MACHinEePsilon
        Do
            xk = -(x * k1 * k2 / (k3 * k4))
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk
            xk = x * k5 * k6 / (k7 * k8)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk
            If qk <> 0.0# Then
                r = pk / qk
            End If
            If r <> 0.0# Then
                t = Math.Abs((ans - r) / r)
                ans = r
            Else
                t = 1.0#
            End If
            If t < thresh Then
                Exit Do
            End If
            k1 = k1 + 1.0#
            k2 = k2 + 1.0#
            k3 = k3 + 2.0#
            k4 = k4 + 2.0#
            k5 = k5 + 1.0#
            k6 = k6 - 1.0#
            k7 = k7 + 2.0#
            k8 = k8 + 2.0#
            If Math.Abs(qk) + Math.Abs(pk) > big Then
                pkm2 = pkm2 * biginv
                pkm1 = pkm1 * biginv
                qkm2 = qkm2 * biginv
                qkm1 = qkm1 * biginv
            End If
            If Math.Abs(qk) < biginv Or Math.Abs(pk) < biginv Then
                pkm2 = pkm2 * big
                pkm1 = pkm1 * big
                qkm2 = qkm2 * big
                qkm1 = qkm1 * big
            End If
            n = n + 1.0#
        Loop Until n = 300.0#
        Result = ans

        IncompleteBetaFE = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Continued fraction expansion #2
    'for incomplete beta integral
    '
    'Cephes Math Library, Release 2.8:  June, 2000
    'Copyright 1984, 1995, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function IncompleteBetaFE2(ByVal a As Double, _
             ByVal b As Double, _
             ByVal x As Double, _
             ByVal big As Double, _
             ByVal biginv As Double) As Double
        Dim Result As Double
        Dim xk As Double
        Dim pk As Double
        Dim pkm1 As Double
        Dim pkm2 As Double
        Dim qk As Double
        Dim qkm1 As Double
        Dim qkm2 As Double
        Dim k1 As Double
        Dim k2 As Double
        Dim k3 As Double
        Dim k4 As Double
        Dim k5 As Double
        Dim k6 As Double
        Dim k7 As Double
        Dim k8 As Double
        Dim r As Double
        Dim t As Double
        Dim ans As Double
        Dim z As Double
        Dim thresh As Double
        Dim n As Double

        k1 = a
        k2 = b - 1.0#
        k3 = a
        k4 = a + 1.0#
        k5 = 1.0#
        k6 = a + b
        k7 = a + 1.0#
        k8 = a + 2.0#
        pkm2 = 0.0#
        qkm2 = 1.0#
        pkm1 = 1.0#
        qkm1 = 1.0#
        z = x / (1.0# - x)
        ans = 1.0#
        r = 1.0#
        n = 0.0#
        thresh = 3.0# * MACHinEePsilon
        Do
            xk = -(z * k1 * k2 / (k3 * k4))
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk
            xk = z * k5 * k6 / (k7 * k8)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk
            If qk <> 0.0# Then
                r = pk / qk
            End If
            If r <> 0.0# Then
                t = Math.Abs((ans - r) / r)
                ans = r
            Else
                t = 1.0#
            End If
            If t < thresh Then
                Exit Do
            End If
            k1 = k1 + 1.0#
            k2 = k2 - 1.0#
            k3 = k3 + 2.0#
            k4 = k4 + 2.0#
            k5 = k5 + 1.0#
            k6 = k6 + 1.0#
            k7 = k7 + 2.0#
            k8 = k8 + 2.0#
            If Math.Abs(qk) + Math.Abs(pk) > big Then
                pkm2 = pkm2 * biginv
                pkm1 = pkm1 * biginv
                qkm2 = qkm2 * biginv
                qkm1 = qkm1 * biginv
            End If
            If Math.Abs(qk) < biginv Or Math.Abs(pk) < biginv Then
                pkm2 = pkm2 * big
                pkm1 = pkm1 * big
                qkm2 = qkm2 * big
                qkm1 = qkm1 * big
            End If
            n = n + 1.0#
        Loop Until n = 300.0#
        Result = ans

        IncompleteBetaFE2 = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Power series for incomplete beta integral.
    'Use when b*x is small and x not too close to 1.
    '
    'Cephes Math Library, Release 2.8:  June, 2000
    'Copyright 1984, 1995, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function IncompleteBetaPS(ByVal a As Double, _
             ByVal b As Double, _
             ByVal x As Double, _
             ByVal MAXGAM As Double) As Double
        Dim Result As Double
        Dim s As Double
        Dim t As Double
        Dim u As Double
        Dim v As Double
        Dim n As Double
        Dim t1 As Double
        Dim z As Double
        Dim ai As Double
        Dim sg As Double

        ai = 1.0# / a
        u = (1.0# - b) * x
        v = u / (a + 1.0#)
        t1 = v
        t = u
        n = 2.0#
        s = 0.0#
        z = MACHinEePsilon * ai
        Do While Math.Abs(v) > z
            u = (n - b) * x / n
            t = t * u
            v = t / (a + n)
            s = s + v
            n = n + 1.0#
        Loop
        s = s + t1
        s = s + ai
        u = a * Math.Log(x)
        If a + b < MAXGAM And Math.Abs(u) < Math.Log(MaxRealNumber) Then
            t = Gamma(a + b) / (Gamma(a) * Gamma(b))
            s = s * t * Math.Pow(x, a)
        Else
            t = LnGamma(a + b, sg) - LnGamma(a, sg) - LnGamma(b, sg) + u + Math.Log(s)
            If t < Math.Log(MinRealNumber) Then
                s = 0.0#
            Else
                s = Math.Exp(t)
            End If
        End If
        Result = s

        IncompleteBetaPS = Result
    End Function

    Public Function Gamma(ByVal X As Double) As Double
        Dim Result As Double
        Dim p As Double
        Dim PP As Double
        Dim q As Double
        Dim QQ As Double
        Dim z As Double
        Dim i As Double
        Dim SgnGam As Double

        SgnGam = 1.0#
        q = Math.Abs(X)
        If q > 33.0# Then
            If X < 0.0# Then
                p = Int(q)
                i = Math.Round(p)
                If i Mod 2.0# = 0.0# Then
                    SgnGam = -1.0#
                End If
                z = q - p
                If z > 0.5 Then
                    p = p + 1.0#
                    z = q - p
                End If
                z = q * Math.Sin(Math.PI * z)
                z = Math.Abs(z)
                z = Math.PI / (z * GammaStirF(q))
            Else
                z = GammaStirF(X)
            End If
            Result = SgnGam * z
            Gamma = Result
            Exit Function
        End If
        z = 1.0#
        Do While X >= 3.0#
            X = X - 1.0#
            z = z * X
        Loop
        Do While X < 0.0#
            If X > -0.000000001 Then
                Result = z / ((1.0# + 0.577215664901533 * X) * X)
                Gamma = Result
                Exit Function
            End If
            z = z / X
            X = X + 1.0#
        Loop
        Do While X < 2.0#
            If X < 0.000000001 Then
                Result = z / ((1.0# + 0.577215664901533 * X) * X)
                Gamma = Result
                Exit Function
            End If
            z = z / X
            X = X + 1.0#
        Loop
        If X = 2.0# Then
            Result = z
            Gamma = Result
            Exit Function
        End If
        X = X - 2.0#
        PP = 0.000160119522476752
        PP = 0.00119135147006586 + X * PP
        PP = 0.0104213797561762 + X * PP
        PP = 0.0476367800457137 + X * PP
        PP = 0.207448227648436 + X * PP
        PP = 0.494214826801497 + X * PP
        PP = 1.0# + X * PP
        QQ = -0.000023158187332412
        QQ = 0.000539605580493303 + X * QQ
        QQ = -0.00445641913851797 + X * QQ
        QQ = 0.011813978522206 + X * QQ
        QQ = 0.0358236398605499 + X * QQ
        QQ = -0.234591795718243 + X * QQ
        QQ = 0.0714304917030273 + X * QQ
        QQ = 1.0# + X * QQ
        Result = z * PP / QQ
        Gamma = Result
        Exit Function

        Gamma = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Natural logarithm of gamma function
    '
    'Input parameters:
    '    X       -   argument
    '
    'Result:
    '    logarithm of the absolute value of the Gamma(X).
    '
    'Output parameters:
    '    SgnGam  -   sign(Gamma(X))
    '
    'Domain:
    '    0 < X < 2.55e305
    '    -2.55e305 < X < 0, X is not an integer.
    '
    'ACCURACY:
    'arithmetic      domain        # trials     peak         rms
    '   IEEE    0, 3                 28000     5.4e-16     1.1e-16
    '   IEEE    2.718, 2.556e305     40000     3.5e-16     8.3e-17
    'The error criterion was relative when the function magnitude
    'was greater than one but absolute when it was less than one.
    '
    'The following test used the relative error criterion, though
    'at certain points the relative error could be much higher than
    'indicated.
    '   IEEE    -200, -4             10000     4.8e-16     1.3e-16
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1989, 1992, 2000 by Stephen L. Moshier
    'Translated to AlgoPascal by Bochkanov Sergey (2005, 2006, 2007).
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function LnGamma(ByVal X As Double, ByRef SgnGam As Double) As Double
        Dim Result As Double
        Dim A As Double
        Dim B As Double
        Dim C As Double
        Dim p As Double
        Dim q As Double
        Dim u As Double
        Dim w As Double
        Dim z As Double
        Dim i As Double
        Dim LogPi As Double
        Dim LS2PI As Double
        Dim Tmp As Double

        SgnGam = 1.0#
        LogPi = 1.1447298858494
        LS2PI = 0.918938533204673
        If X < -34.0# Then
            q = -X
            w = LnGamma(q, Tmp)
            p = Int(q)
            i = Math.Round(p)
            If i Mod 2.0# = 0.0# Then
                SgnGam = -1.0#
            Else
                SgnGam = 1.0#
            End If
            z = q - p
            If z > 0.5 Then
                p = p + 1.0#
                z = p - q
            End If
            z = q * Math.Sin(Math.PI * z)
            Result = LogPi - Math.Log(z) - w
            LnGamma = Result
            Exit Function
        End If
        If X < 13.0# Then
            z = 1.0#
            p = 0.0#
            u = X
            Do While u >= 3.0#
                p = p - 1.0#
                u = X + p
                z = z * u
            Loop
            Do While u < 2.0#
                z = z / u
                p = p + 1.0#
                u = X + p
            Loop
            If z < 0.0# Then
                SgnGam = -1.0#
                z = -z
            Else
                SgnGam = 1.0#
            End If
            If u = 2.0# Then
                Result = Math.Log(z)
                LnGamma = Result
                Exit Function
            End If
            p = p - 2.0#
            X = X + p
            B = -1378.25152569121
            B = -38801.6315134638 + X * B
            B = -331612.992738871 + X * B
            B = -1162370.97492762 + X * B
            B = -1721737.0082084 + X * B
            B = -853555.664245765 + X * B
            C = 1.0#
            C = -351.815701436523 + X * C
            C = -17064.2106651881 + X * C
            C = -220528.590553854 + X * C
            C = -1139334.44367983 + X * C
            C = -2532523.07177583 + X * C
            C = -2018891.41433533 + X * C
            p = X * B / C
            Result = Math.Log(z) + p
            LnGamma = Result
            Exit Function
        End If
        q = (X - 0.5) * Math.Log(X) - X + LS2PI
        If X > 100000000.0# Then
            Result = q
            LnGamma = Result
            Exit Function
        End If
        p = 1.0# / (X * X)
        If X >= 1000.0# Then
            q = q + ((7.93650793650794 * 0.0001 * p - 2.77777777777778 * 0.001) * p + 0.0833333333333333) / X
        Else
            A = 8.11614167470508 * 0.0001
            A = -(5.95061904284301 * 0.0001) + p * A
            A = 7.93650340457717 * 0.0001 + p * A
            A = -(2.777777777301 * 0.001) + p * A
            A = 8.33333333333332 * 0.01 + p * A
            q = q + A / X
        End If
        Result = q

        LnGamma = Result
    End Function


    Private Function GammaStirF(ByVal X As Double) As Double
        Dim Result As Double
        Dim y As Double
        Dim w As Double
        Dim v As Double
        Dim Stir As Double

        w = 1.0# / X
        Stir = 0.000787311395793094
        Stir = -0.000229549961613378 + w * Stir
        Stir = -0.00268132617805781 + w * Stir
        Stir = 0.00347222221605459 + w * Stir
        Stir = 0.0833333333333482 + w * Stir
        w = 1.0# + w * Stir
        y = Math.Exp(X)
        If X > 143.01608 Then
            v = Math.Pow(X, 0.5 * X - 0.25)
            y = v * (v / y)
        Else
            y = Math.Pow(X, X - 0.5) / y
        End If
        Result = 2.506628274631 * y * w

        GammaStirF = Result
    End Function
    Public Function Erf(ByVal x As Double) As Double
        Dim Result As Double
        Dim XSq As Double
        Dim S As Double
        Dim P As Double
        Dim Q As Double

        S = Math.Sign(x)
        x = Math.Abs(x)
        If x < 0.5 Then
            XSq = x * x
            P = 0.00754772803341863
            P = 0.288805137207594 + XSq * P
            P = 14.3383842191748 + XSq * P
            P = 38.0140318123903 + XSq * P
            P = 3017.82788536508 + XSq * P
            P = 7404.07142710151 + XSq * P
            P = 80437.363096084 + XSq * P
            Q = 0.0#
            Q = 1.0# + XSq * Q
            Q = 38.0190713951939 + XSq * Q
            Q = 658.07015545924 + XSq * Q
            Q = 6379.60017324428 + XSq * Q
            Q = 34216.5257924629 + XSq * Q
            Q = 80437.363096084 + XSq * Q
            Result = S * 1.12837916709551 * x * P / Q
            Erf = Result
            Exit Function
        End If
        If x >= 10.0# Then
            Result = S
            Erf = Result
            Exit Function
        End If
        Result = S * (1.0# - ErfC(x))

        Erf = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Complementary error function
    '
    ' 1 - erf(x) =
    '
    '                          inf.
    '                            -
    '                 2         | |          2
    '  erfc(x)  =  --------     |    exp( - t  ) dt
    '              sqrt(pi)   | |
    '                          -
    '                           x
    '
    '
    'For small x, erfc(x) = 1 - erf(x); otherwise rational
    'approximations are computed.
    '
    '
    'ACCURACY:
    '
    '                     Relative error:
    'arithmetic   domain     # trials      peak         rms
    '   IEEE      0,26.6417   30000       5.7e-14     1.5e-14
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1988, 1992, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function ErfC(ByVal x As Double) As Double
        Dim Result As Double
        Dim P As Double
        Dim Q As Double

        If x < 0.0# Then
            Result = 2.0# - ErfC(-x)
            ErfC = Result
            Exit Function
        End If
        If x < 0.5 Then
            Result = 1.0# - Erf(x)
            ErfC = Result
            Exit Function
        End If
        If x >= 10.0# Then
            Result = 0.0#
            ErfC = Result
            Exit Function
        End If
        P = 0.0#
        P = 0.56418778255074 + x * P
        P = 9.67580788298727 + x * P
        P = 77.0816173036843 + x * P
        P = 368.519615471001 + x * P
        P = 1143.26207070389 + x * P
        P = 2320.43959025164 + x * P
        P = 2898.02932921677 + x * P
        P = 1826.33488422951 + x * P
        Q = 1.0#
        Q = 17.1498094362761 + x * Q
        Q = 137.125596050062 + x * Q
        Q = 661.736120710765 + x * Q
        Q = 2094.38436778954 + x * Q
        Q = 4429.61280388368 + x * Q
        Q = 6089.54242327244 + x * Q
        Q = 4958.82756472114 + x * Q
        Q = 1826.33488422951 + x * Q
        Result = Math.Exp(-1 * x * x) * P / Q

        ErfC = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Normal distribution function
    '
    'Returns the area under the Gaussian probability density
    'function, integrated from minus infinity to x:
    '
    '                           x
    '                            -
    '                  1        | |          2
    '   ndtr(x)  = ---------    |    exp( - t /2 ) dt
    '              sqrt(2pi)  | |
    '                          -
    '                         -inf.
    '
    '            =  ( 1 + erf(z) ) / 2
    '            =  erfc(z) / 2
    '
    'where z = x/sqrt(2). Computation is via the functions
    'erf and erfc.
    '
    '
    'ACCURACY:
    '
    '                     Relative error:
    'arithmetic   domain     # trials      peak         rms
    '   IEEE     -13,0        30000       3.4e-14     6.7e-15
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1988, 1992, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function NormalDistribution(ByVal x As Double) As Double
        Dim Result As Double

        Result = 0.5 * (Erf(x / 1.4142135623731) + 1.0#)

        NormalDistribution = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Inverse of the error function
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1988, 1992, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InvErf(ByVal E As Double) As Double
        Dim Result As Double

        Result = InvNormalDistribution(0.5 * (E + 1.0#)) / Math.Sqrt(2.0#)

        InvErf = Result
    End Function



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Inverse of Normal distribution function
    '
    'Returns the argument, x, for which the area under the
    'Gaussian probability density function (integrated from
    'minus infinity to x) is equal to y.
    '
    '
    'For small arguments 0 < y < exp(-2), the program computes
    'z = sqrt( -2.0 * log(y) );  then the approximation is
    'x = z - log(z)/z  - (1/z) P(1/z) / Q(1/z).
    'There are two rational functions P/Q, one for 0 < y < exp(-32)
    'and the other for y up to exp(-2).  For larger arguments,
    'w = y - 0.5, and  x/sqrt(2pi) = w + w**3 R(w**2)/S(w**2)).
    '
    'ACCURACY:
    '
    '                     Relative error:
    'arithmetic   domain        # trials      peak         rms
    '   IEEE     0.125, 1        20000       7.2e-16     1.3e-16
    '   IEEE     3e-308, 0.135   50000       4.6e-16     9.8e-17
    '
    'Cephes Math Library Release 2.8:  June, 2000
    'Copyright 1984, 1987, 1988, 1992, 2000 by Stephen L. Moshier
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InvNormalDistribution(ByVal y0 As Double) As Double
        Dim Result As Double
        Dim Expm2 As Double
        Dim S2Pi As Double
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Dim y2 As Double
        Dim x0 As Double
        Dim x1 As Double
        Dim code As Double
        Dim P0 As Double
        Dim Q0 As Double
        Dim P1 As Double
        Dim Q1 As Double
        Dim P2 As Double
        Dim Q2 As Double

        Expm2 = 0.135335283236613
        S2Pi = 2.506628274631
        If y0 <= 0.0# Then
            Result = -MaxRealNumber
            InvNormalDistribution = Result
            Exit Function
        End If
        If y0 >= 1.0# Then
            Result = MaxRealNumber
            InvNormalDistribution = Result
            Exit Function
        End If
        code = 1.0#
        y = y0
        If y > 1.0# - Expm2 Then
            y = 1.0# - y
            code = 0.0#
        End If
        If y > Expm2 Then
            y = y - 0.5
            y2 = y * y
            P0 = -59.9633501014108
            P0 = 98.0010754186 + y2 * P0
            P0 = -56.676285746907 + y2 * P0
            P0 = 13.931260938728 + y2 * P0
            P0 = -1.23916583867381 + y2 * P0
            Q0 = 1.0#
            Q0 = 1.95448858338142 + y2 * Q0
            Q0 = 4.67627912898882 + y2 * Q0
            Q0 = 86.3602421390891 + y2 * Q0
            Q0 = -225.462687854119 + y2 * Q0
            Q0 = 200.260212380061 + y2 * Q0
            Q0 = -82.0372256168333 + y2 * Q0
            Q0 = 15.9056225126212 + y2 * Q0
            Q0 = -1.1833162112133 + y2 * Q0
            x = y + y * y2 * P0 / Q0
            x = x * S2Pi
            Result = x
            InvNormalDistribution = Result
            Exit Function
        End If
        x = Math.Sqrt(-(2.0# * Math.Log(y)))
        x0 = x - Math.Log(x) / x
        z = 1.0# / x
        If x < 8.0# Then
            P1 = 4.05544892305962
            P1 = 31.5251094599894 + z * P1
            P1 = 57.1628192246421 + z * P1
            P1 = 44.0805073893201 + z * P1
            P1 = 14.6849561928858 + z * P1
            P1 = 2.1866330685079 + z * P1
            P1 = -(1.40256079171354 * 0.1) + z * P1
            P1 = -(3.50424626827848 * 0.01) + z * P1
            P1 = -(8.57456785154685 * 0.0001) + z * P1
            Q1 = 1.0#
            Q1 = 15.7799883256467 + z * Q1
            Q1 = 45.3907635128879 + z * Q1
            Q1 = 41.3172038254672 + z * Q1
            Q1 = 15.0425385692908 + z * Q1
            Q1 = 2.50464946208309 + z * Q1
            Q1 = -(1.42182922854788 * 0.1) + z * Q1
            Q1 = -(3.80806407691578 * 0.01) + z * Q1
            Q1 = -(9.33259480895457 * 0.0001) + z * Q1
            x1 = z * P1 / Q1
        Else
            P2 = 3.23774891776946
            P2 = 6.91522889068984 + z * P2
            P2 = 3.93881025292474 + z * P2
            P2 = 1.33303460815808 + z * P2
            P2 = 2.01485389549179 * 0.1 + z * P2
            P2 = 1.2371663481782 * 0.01 + z * P2
            P2 = 3.01581553508235 * 0.0001 + z * P2
            P2 = 2.65806974686738 * 0.000001 + z * P2
            P2 = 6.23974539184983 * 0.000000001 + z * P2
            Q2 = 1.0#
            Q2 = 6.02427039364742 + z * Q2
            Q2 = 3.67983563856161 + z * Q2
            Q2 = 1.37702099489081 + z * Q2
            Q2 = 2.16236993594497 * 0.1 + z * Q2
            Q2 = 1.34204006088543 * 0.01 + z * Q2
            Q2 = 3.28014464682128 * 0.0001 + z * Q2
            Q2 = 2.89247864745381 * 0.000001 + z * Q2
            Q2 = 6.79019408009981 * 0.000000001 + z * Q2
            x1 = z * P2 / Q2
        End If
        x = x0 - x1
        If code <> 0.0# Then
            x = -x
        End If
        Result = x

        InvNormalDistribution = Result
    End Function

End Module
