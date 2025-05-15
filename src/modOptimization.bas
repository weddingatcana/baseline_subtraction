Option Explicit

Private Type tRamanStats
    Average As Double
    Deviation As Double
End Type

Public Enum eValleys
    left_to_right
    both_ways
    both_ways_and_ends
End Enum

Public Function optWhittaker_search(ByRef raw#(), _
                                    ByRef Ax#(), _
                                    ByRef Ay#(), _
                                    ByRef valleys#(), _
                                    ByRef peaks#(), _
                                    Optional ByVal p# = 0.01, _
                                    Optional ByVal lambda# = 0#, _
                                    Optional ByVal delta& = 2, _
                                    Optional ByVal opt_iter& = 3, _
                                    Optional ByVal search_iter# = 20, _
                                    Optional ByVal lambda_mult# = 10) As Double()
                                                 
    Dim rawMaxRowZf&, _
        rawMaxRowAy&, _
        rawMaxColAy&, _
        rawMaxRowP&, _
        rawMaxRow&, _
        i&, j&, k&, _
        poly_function#(), _
        Zf_moving_peaks#(), _
        temp#(), _
        coeff#(), _
        W#(), d#(), _
        Z1#(), Z2#(), _
        Z3#(), Z4#(), _
        Z5#(), Z6#(), _
        Zf#(), Z_L1#(), _
        Z_L2#(), Z_Lf#(), _
        loss_function_v#, _
        loss_function_p#, _
        residual#, dl&, _
        y#, z#, s&, qq&

        rawMaxRowAy = UBound(Ay, 1)
        rawMaxColAy = UBound(Ay, 2)
        rawMaxRowP = UBound(peaks, 1)
        rawMaxRow = UBound(raw, 1)
        
        ReDim Zf(1 To rawMaxRowAy, 1 To rawMaxColAy)
        ReDim Zf_moving_peaks(1 To rawMaxRowP, 1 To 1)
        
        W = modMatrix.matSpy(rawMaxRowAy)
        d = modMatrix.matDiff(W, delta)
        
        s = 0
        qq = 0
        dl = 1
        Do
        
            k = 1
            Do
            
                Z1 = modMatrix.matTra(d)
                Z2 = modMatrix.matMul(Z1, d)
                Z3 = modMatrix.matScl(Z2, lambda)
                Z4 = modMatrix.matAdd(W, Z3)
                Z5 = modMatrix.matInv(Z4)
                Z6 = modMatrix.matMul(Z5, W)
                Zf = modMatrix.matMul(Z6, Ay)
            
                For i = 1 To rawMaxRowAy
                    
                    y = Ay(i, 1)
                    z = Zf(i, 1)
                    residual = y - z
                    
                    If residual > 0 Then
                        W(i, i) = p
                    Else
                        W(i, i) = 1 - p
                    End If
                    
                Next i
            
                k = k + 1
            
                If k > opt_iter Then
                    Exit Do
                End If
            
            Loop
            
            coeff = modOptimization.optPolyCoeff(valleys, 4)
            poly_function = modOptimization.optPolyFit_seperate_coeff(Ax, coeff)
            
            j = 1
            For i = 1 To rawMaxRow
            
                If peaks(j, 1) = raw(i, 1) Then
                
                    For k = 1 To rawMaxRowAy
                
                        If peaks(j, 1) >= Ax(k, 1) Then
                            Zf_moving_peaks(j, 1) = Zf(k, 1)
                            j = j + 1
                            Exit For
                        End If
                    
                    Next k
                    
                    If j > rawMaxRowP Then
                        Exit For
                    End If
                    
                End If
                
            Next i
            
            loss_function_p = modOptimization.optSSR(modMatrix.matVec(peaks, 2), Zf_moving_peaks)
            loss_function_v = modOptimization.optSSR(modMatrix.matVec(poly_function, 2), Zf)
            
            rawMaxRowZf = UBound(Zf, 1)
            Z_L1 = modMatrix.matReduce(Zf)
            
            ReDim Preserve Z_L1(1 To (rawMaxRowZf + 2))
            ReDim Z_L2(1 To (rawMaxRowZf + 2), 1 To 1)
            
            For i = 1 To (rawMaxRowZf + 2)
            
                If i = 1 Then
                    Z_L2(i, 1) = loss_function_p
                ElseIf i = 2 Then
                    Z_L2(i, 1) = loss_function_v
                Else
                    Z_L2(i, 1) = Z_L1(i - 2)
                End If
                
            Next i
            
            If s = 0 Then
                Z_Lf = Z_L2
                s = s + 1
            Else
                Z_Lf = modMatrix.matJoin_Ext(Z_Lf, Z_L2)
            End If
            
            'lambda = lambda + dl
            lambda = lambda + (10 ^ (qq))
            qq = qq + 1
            
            If lambda > search_iter Then
                Exit Do
            End If
        
        Loop
    
        optWhittaker_search = Z_Lf

End Function

Public Function optWhittaker(ByRef A#(), _
                             Optional ByVal p# = 0.01, _
                             Optional ByVal lambda# = 100#, _
                             Optional ByVal delta& = 2, _
                             Optional ByVal iter& = 5) As Double()
                                                 
    Dim rawMaxRowA&, _
        rawMaxColA, _
        i&, j&, k&, _
        W#(), d#(), _
        Z1#(), Z2#(), _
        Z3#(), Z4#(), _
        Z5#(), Z6#(), _
        Zf#(), y#, z#, _
        residual#
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        ReDim Zf(1 To rawMaxRowA, 1 To rawMaxColA)
        
        W = modMatrix.matSpy(rawMaxRowA)
        d = modMatrix.matDiff(W, delta)
        
        k = 1
        Do
        
            Z1 = modMatrix.matTra(d)
            Z2 = modMatrix.matMul(Z1, d)
            Z3 = modMatrix.matScl(Z2, lambda)
            Z4 = modMatrix.matAdd(W, Z3)
            Z5 = modMatrix.matInv(Z4)
            Z6 = modMatrix.matMul(Z5, W)
            Zf = modMatrix.matMul(Z6, A)
        
            For i = 1 To rawMaxRowA
                
                y = A(i, 1)
                z = Zf(i, 1)
                residual = y - z
                
                If residual > 0 Then
                    W(i, i) = p
                Else
                    W(i, i) = 1 - p
                End If
                
            Next i
        
            k = k + 1
        
            If k > iter Then
                Exit Do
            End If
        
        Loop
    
        optWhittaker = Zf

End Function

Public Function optWeightedLeastSquares(ByRef Ax#(), _
                                        ByRef Ay#(), _
                                        Optional ByVal polyOrder# = 5, _
                                        Optional ByVal p# = 0.01, _
                                        Optional ByVal iter& = 5) As Double()
                                                 
    Dim rawMaxRowA&, _
        i&, j&, k&, _
        W#(), B1#(), _
        B2#(), B3#(), _
        B4#(), B5#(), _
        B6#(), Bf#(), _
        Xv#(), WeightedPoly#(), _
        y#, z#, _
        residual#
        
        rawMaxRowA = UBound(Ax, 1)
        ReDim Bf(1 To polyOrder, 1 To 1)
        
        W = modMatrix.matSpy(rawMaxRowA)
        Xv = modMath.mathVandermonde(Ax, polyOrder)
        
        k = 1
        Do
        
            B1 = modMatrix.matTra(Xv)
            B2 = modMatrix.matMul(B1, W)
            B3 = modMatrix.matMul(B2, Xv)
            B4 = modMatrix.matInv(B3)
            B5 = modMatrix.matMul(B4, B1)
            B6 = modMatrix.matMul(B5, W)
            Bf = modMatrix.matMul(B6, Ay)
            
            WeightedPoly = modOptimization.optPolyFit_seperate_coeff(Ax, Bf)
        
            For i = 1 To rawMaxRowA
                
                y = Ay(i, 1)
                z = WeightedPoly(i, 2)
                residual = y - z
                
                If residual > 0 Then
                    W(i, i) = p
                Else
                    W(i, i) = 1 - p
                End If
                
            Next i
        
            k = k + 1
        
            If k > iter Then
                Exit Do
            End If
        
        Loop
    
        optWeightedLeastSquares = Bf

End Function

Public Function optfD(ByRef B#(), _
                      ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, j&, _
        fD#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)

        ReDim fD(1 To rawMaxRowA - 2)
        
        For i = 2 To (rawMaxRowA - 1)
            
            fD(i - 1) = (A(i + 1, 1) - A(i - 1, 1)) / _
                        (B(i + 1, 1) - B(i - 1, 1))
                        
        Next i
        
        optfD = fD

End Function

Public Function optPeaks_SG(ByRef A#(), _
                            ByRef SG#()) As Double()
                            
    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowSG&, _
        rawMaxColSG&, _
        first#, _
        second#, _
        third#, _
        waves#(), _
        peaks#(), _
        final#(), _
        i&, j&, k&
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowSG = UBound(SG, 1)
        rawMaxColSG = UBound(SG, 2)
        
        j = 1
        For i = 2 To (rawMaxRowSG - 3)
        
            first = (SG((i + 1), 2) - SG((i - 1), 2)) / _
                    (SG((i + 1), 1) - SG((i - 1), 1))
            
            second = (SG((i + 2), 2) - SG((i), 2)) / _
                     (SG((i + 2), 1) - SG((i), 1))
        
            third = (SG((i + 3), 2) - SG((i + 1), 2)) / _
                    (SG((i + 3), 1) - SG((i + 1), 1))
            
            If first > 0 And _
               second > 0 And _
               third < 0 Then
               
               ReDim Preserve peaks(1 To j)
               ReDim Preserve waves(1 To j)
               
               waves(j) = SG((i + 2), 1)
               
               For k = 1 To rawMaxRowA
               
                   If waves(j) = A(k, 1) Then
                       peaks(j) = A(k, 2)
                    End If
               
               Next k
               
               j = j + 1
               
            End If
        
        Next i
        
        final = modMatrix.matJoin(modMatrix.matDim(waves), modMatrix.matDim(peaks))
        optPeaks_SG = final

End Function

Public Function optPeaks(ByRef raw#(), _
                         ByRef A#(), _
                         ByRef fD#()) As Double()

        Dim stats As tRamanStats, _
            cutoff#, _
            rawMaxRow&, _
            rawMaxRowA&, _
            rawMaxRowS&, _
            rawMaxRowfD&, _
            i&, j&, k&, _
            metric#, _
            peaks#(), _
            waves#(), _
            sizes#(), _
            final#()
            
            rawMaxRowA = UBound(A, 1)
            rawMaxRowfD = UBound(fD)
            rawMaxRow = UBound(raw, 1)
            
            j = 1
            For i = 1 To (rawMaxRowfD - 2)
            
                If fD(i + 0) > 0 And _
                   fD(i + 1) > 0 And _
                   fD(i + 2) < 0 Then
                    
                   ReDim Preserve waves(1 To j)
                   ReDim Preserve peaks(1 To j)
                   ReDim Preserve sizes(1 To j)
                   
                   For k = 1 To rawMaxRow
                   
                    If raw(k, 1) >= A(i + 1, 1) And _
                       raw(k, 1) <= A((i + 4), 1) Then
                       
                       If raw((k + 1), 2) < raw(k, 2) Then
                           
                           waves(j) = raw(k, 1)
                           peaks(j) = raw(k, 2)
                           sizes(j) = Abs((fD(i + 2) - fD(i + 1)))
                   
                           j = j + 1
                           Exit For
                           
                        End If
                       
                    End If
                   
                   Next k
        
                End If
            
            Next i
            
            stats.Average = optAvg(modMatrix.matDim(sizes))
            stats.Deviation = optStd(modMatrix.matDim(sizes), stats.Average)
            'metric = stats.Deviation / stats.Average
            rawMaxRowS = UBound(sizes)
            
'            If metric > 1.5 Then
'                cutoff = 1 / metric
'            Else
'                cutoff = metric
'            End If
            
            k = 0
            For i = 1 To rawMaxRowS
            
                If sizes(i) < cutoff Then
                    k = k + 1
                End If
            
            Next i
            
            ReDim final(1 To (rawMaxRowS - k), 1 To 2)
            
            j = 1
            For i = 1 To rawMaxRowS
                
                If sizes(i) >= cutoff Then
                    final(j, 1) = waves(i)
                    final(j, 2) = peaks(i)
                    j = j + 1
                End If
            
            Next i
            
            optPeaks = final

End Function

Public Function optValleys(ByRef raw#(), _
                           ByRef A#(), _
                           ByRef fD#()) As Double()

        Dim stats As tRamanStats, _
            cutoff#, _
            rawMaxRow&, _
            rawMaxRowA&, _
            rawMaxRowS&, _
            rawMaxRowfD&, _
            i&, j&, k&, _
            metric#, _
            valleys#(), _
            waves#(), _
            sizes#(), _
            final#()
            
            rawMaxRow = UBound(raw, 1)
            rawMaxRowA = UBound(A, 1)
            rawMaxRowfD = UBound(fD)
            
            j = 1
            For i = 1 To (rawMaxRowfD - 2)
            
                If fD(i + 0) < 0 And _
                   fD(i + 1) < 0 And _
                   fD(i + 2) > 0 Then
                    
                   ReDim Preserve waves(1 To j)
                   ReDim Preserve valleys(1 To j)
                   ReDim Preserve sizes(1 To j)
                   
                   For k = 1 To rawMaxRow
                   
                    If raw(k, 1) >= A(i + 1, 1) And _
                       raw(k, 1) <= A((i + 4), 1) Then
                       
                       If raw((k + 1), 2) > raw(k, 2) Then
                           
                           waves(j) = raw(k, 1)
                           valleys(j) = raw(k, 2)
                           sizes(j) = Abs((fD(i + 2) - fD(i + 1)))
                   
                           j = j + 1
                           Exit For
                           
                        End If
                       
                    End If
                   
                   Next k
                   
                End If
            
            Next i
            
            stats.Average = optAvg(modMatrix.matDim(sizes))
            stats.Deviation = optStd(modMatrix.matDim(sizes), stats.Average)
            'metric = stats.Deviation / stats.Average
            rawMaxRowS = UBound(sizes)
            
'            If metric > 1.5 Then
'                cutoff = 1 / metric
'            Else
'                cutoff = metric
'            End If
            
            k = 0
            For i = 1 To rawMaxRowS
            
                If sizes(i) < cutoff Then
                    k = k + 1
                End If
            
            Next i
            
            ReDim final(1 To (rawMaxRowS - k), 1 To 2)
            
            j = 1
            For i = 1 To rawMaxRowS
                
                If sizes(i) >= cutoff Then
                    final(j, 1) = waves(i)
                    final(j, 2) = valleys(i)
                    j = j + 1
                End If
            
            Next i
            
            optValleys = final

End Function

Public Function optValleys_advanced(ByRef raw#(), _
                           ByRef A#(), _
                           ByRef fD#(), _
                           ByVal special As eValleys) As Double()

        Dim stats As tRamanStats, _
            cutoff#, _
            rawMaxRow&, _
            rawMaxRowA&, _
            rawMaxRowS&, _
            rawMaxRowfD&, _
            i&, j&, n&, _
            k&, p&, _
            metric#, _
            valleys#(), _
            waves#(), _
            sizes#(), _
            final#()
            
            rawMaxRow = UBound(raw, 1)
            rawMaxRowA = UBound(A, 1)
            rawMaxRowfD = UBound(fD)
            
            j = 1
            For i = 1 To (rawMaxRowfD - 2)
            
                If fD(i + 0) < 0 And _
                   fD(i + 1) < 0 And _
                   fD(i + 2) > 0 Then
                    
                   ReDim Preserve waves(1 To j)
                   ReDim Preserve valleys(1 To j)
                   ReDim Preserve sizes(1 To j)
                   
                   For k = 1 To rawMaxRow
                   
                    If raw(k, 1) >= A(i + 1, 1) And _
                       raw(k, 1) <= A((i + 4), 1) Then
                       
                       If raw((k + 1), 2) > raw(k, 2) Then
                           
                           waves(j) = raw(k, 1)
                           valleys(j) = raw(k, 2)
                           sizes(j) = Abs((fD(i + 2) - fD(i + 1)))
                   
                           j = j + 1
                           Exit For
                           
                        End If
                       
                    End If
                   
                   Next k
                   
                End If
            
            Next i
            
            If special = both_ways Or _
               special = both_ways_and_ends Then
            
                p = j
                For n = 1 To (rawMaxRowfD - 2)
                
                    If fD((rawMaxRowfD - 2) - n + 3) < 0 And _
                       fD((rawMaxRowfD - 2) - n + 2) < 0 And _
                       fD((rawMaxRowfD - 2) - n + 1) > 0 Then
                        
                       ReDim Preserve waves(1 To p)
                       ReDim Preserve valleys(1 To p)
                       ReDim Preserve sizes(1 To p)
                       
                       For k = rawMaxRow To 1 Step -1
                       
                        If raw(k, 1) >= A(((rawMaxRowfD - 2) - n + 3) - 1, 1) And _
                           raw(k, 1) <= A((((rawMaxRowfD - 2) - n + 3) - 4), 1) Then
                           
                           If raw((k - 1), 2) > raw(k, 2) Then
                               
                               waves(p) = raw(k, 1)
                               valleys(p) = raw(k, 2)
                               sizes(p) = Abs((fD(((rawMaxRowfD - 2) - n + 3) - 2) - fD(((rawMaxRowfD - 2) - n + 3) - 1)))
                       
                               p = p + 1
                               Exit For
                               
                            End If
                           
                        End If
                       
                       Next k
                       
                    End If
                
                Next n
                
            ElseIf special = left_to_right Then
            End If
            
            stats.Average = optAvg(modMatrix.matDim(sizes))
            stats.Deviation = optStd(modMatrix.matDim(sizes), stats.Average)
            'metric = stats.Deviation / stats.Average
            rawMaxRowS = UBound(sizes)
            
'            If metric > 1.5 Then
'                cutoff = 1 / metric
'            Else
'                cutoff = metric
'            End If
            
            k = 0
            For i = 1 To rawMaxRowS
            
                If sizes(i) < cutoff Then
                    k = k + 1
                End If
            
            Next i
            
            If k = rawMaxRowS Then
                Exit Function
            End If
            
            If special = both_ways_and_ends Then
            
                ReDim final(1 To (rawMaxRowS - k + 1), 1 To 2)
                
                final(1, 1) = raw(1, 1)
                final(1, 2) = raw(1, 2)

                For i = 1 To rawMaxRowS
                
                    If sizes(i) >= cutoff Then
                        final(i + 1, 1) = waves(i)
                        final(i + 1, 2) = valleys(i)
                    End If
                
                Next i
                
                final(rawMaxRowS - k + 1, 1) = raw(rawMaxRow, 1)
                final(rawMaxRowS - k + 1, 2) = raw(rawMaxRow, 2)
            
            Else
            
                ReDim final(1 To (rawMaxRowS - k), 1 To 2)
        
                j = 1
                For i = 1 To rawMaxRowS
                    
                    If sizes(i) >= cutoff Then
                        final(j, 1) = waves(i)
                        final(j, 2) = valleys(i)
                        j = j + 1
                    End If
                
                Next i
                
            End If
            
            optValleys_advanced = final

End Function

Public Function optPolyCoeff(ByRef A#(), _
                             ByRef polyOrder&) As Double()

    Dim rawX#(), _
        rawY#(), _
        Vm#(), _
        i_coeff#(), _
        f_coeff#()

        
        rawX = modMatrix.matVec(A, 1)
        rawY = modMatrix.matVec(A, 2)
        Vm = modMath.mathVandermonde(rawX, polyOrder)
        
        i_coeff = modMatrix.matPin(Vm)
        f_coeff = modMatrix.matMul(i_coeff, rawY)
        
        optPolyCoeff = f_coeff

End Function

Public Function optPolyFit(A#(), _
                           polyOrder&) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        coeff#(), _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        coeff = optPolyCoeff(A, polyOrder)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit = C

End Function

Public Function optPolyFit_seperate_coeff(A#(), _
                                          coeff#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To (rawMaxColA + 1))
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit_seperate_coeff = C

End Function

Public Function optSavGol(ByRef A#(), _
                          Optional ByVal window& = 11, _
                          Optional ByVal polyOrder& = 2) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        moving_mid&, _
        i&, j&, _
        k&, p&, _
        length&, _
        buffer#(), _
        Poly#(), _
        mid&, _
        C#()

        If window Mod 2 <> 0 And _
           window >= polyOrder + 1 Then
        Else
            Exit Function
        End If

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        length = rawMaxRowA - (window - 1)
        mid = window \ 2
        moving_mid = mid
        
        If rawMaxColA <> 2 Then
            Exit Function
        End If
        
        ReDim C(1 To length, 1 To rawMaxColA)
        ReDim buffer(1 To window, 1 To rawMaxColA)
        ReDim Poly(1 To window, 1 To rawMaxColA)
        
        For i = 1 To length
            
            j = i
            k = 1
            Do
            
                If k > window Then
                    Exit Do
                End If
        
                buffer(k, 1) = A(j, 1)
                buffer(k, 2) = A(j, 2)
                
                j = j + 1
                k = k + 1
        
            Loop
            
            Poly = modOptimization.optPolyFit(buffer, polyOrder)
            
            C(i, 1) = A(moving_mid, 1)
            C(i, 2) = Poly(mid, 2)
            moving_mid = moving_mid + 1
        
        Next i

        optSavGol = C
        
End Function

Public Function optSSR#(ByRef A#(), _
                        ByRef B#())
                        
    Dim rawMaxRowA&, _
        rawMaxRowB&, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        rawMaxRowB = UBound(B, 1)
               
        If rawMaxRowA <> rawMaxRowB Then
            Exit Function
        End If
        
        sum = 0
        For i = 1 To rawMaxRowA
        
            sum = sum + (A(i, 1) - B(i, 1)) ^ 2
        
        Next i
        
        optSSR = sum

End Function

Public Function optSSE#(ByRef A#(), _
                        ByVal avg#)

    Dim rawMaxRowA&, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        For i = 1 To rawMaxRowA
        
            sum = sum + (A(i, 1) - avg) ^ 2
        
        Next i
        
        optSSE = sum
        
End Function

Public Function optR2#(ByVal SSR#, _
                       ByVal SSE#)
                       
    optR2 = 1 - (SSR / (SSR + SSE))
                       
End Function

Public Function optAvg#(ByRef A#())

    Dim rawMaxRowA&, _
        avg#, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        For i = 1 To rawMaxRowA
            sum = sum + A(i, 1)
        Next i
        
        avg = sum / rawMaxRowA
        optAvg = avg

End Function

Public Function optStd#(ByRef A#(), _
                        ByVal avg#)

    Dim rawMaxRowA&, _
        sum#, _
        dev#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        dev = 0
        For i = 1 To rawMaxRowA
            sum = (A(i, 1) - avg) ^ 2
            dev = dev + sum
        Next i
        
        dev = Sqr(dev / rawMaxRowA)
        optStd = dev
                        
End Function

Public Function optBubble(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        i&, j&, _
        swap#
        
        rawMaxRowA = UBound(A, 1)
        
        For i = 1 To rawMaxRowA
            For j = i + 1 To rawMaxRowA
        
                If A(i, 1) > A(j, 1) Then
                
                    swap = A(i, 1)
                    A(i, 1) = A(j, 1)
                    A(j, 1) = swap
                    
                End If
        
            Next j
        Next i
        
        optBubble = A

End Function
