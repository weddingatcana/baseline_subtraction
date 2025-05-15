Option Explicit

Public Function mathDownSampling(ByRef A#(), _
                                 ByVal rate&) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, j&, k&, _
        C#()

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)

        If rawMaxRowA <> 1 And rawMaxColA <> 1 Then
            Exit Function
        End If
        
        If rawMaxRowA > rawMaxColA Then
        
            If (rawMaxRowA \ rate) < 1 Then
                Exit Function
            Else
                ReDim C(1 To (rawMaxRowA \ rate), 1 To 1)
                j = 1
            End If
            
        Else
        
            If (rawMaxColA \ rate) < 1 Then
                Exit Function
            Else
                ReDim C(1 To 1, 1 To (rawMaxColA \ rate))
                j = 0
            End If
            
        End If
        
        i = 1
        k = 1
        If j > 0 Then
        
            Do
                If k > (rawMaxRowA \ rate) Then
                    mathDownSampling = C
                    Exit Function
                Else
                    If i > rawMaxRowA Then
                        Exit Do
                    Else
                        C(k, 1) = A(i, 1)
                        i = i + rate
                        k = k + 1
                    End If
                End If
            Loop
        
        Else
        
            Do
                If k > (rawMaxColA \ rate) Then
                    mathDownSampling = C
                    Exit Function
                Else
                    If i > rawMaxColA Then
                        Exit Do
                    Else
                        C(1, k) = A(1, i)
                        i = i + rate
                        k = k + 1
                    End If
                End If
            Loop
            
        End If
        
        mathDownSampling = C
        
End Function

Public Function mathVandermonde(ByRef rawX#(), ByVal polyOrder&) As Double()

    Dim rawMaxRowX&, _
        i&, j&, _
        final#()
        
        rawMaxRowX = UBound(rawX, 1)
        
        ReDim final(1 To rawMaxRowX, 1 To (polyOrder + 1))
        
        For j = 0 To polyOrder
            For i = 1 To rawMaxRowX
            
                final(i, (j + 1)) = rawX(i, 1) ^ (j)
            
            Next i
        Next j

        mathVandermonde = final

End Function
