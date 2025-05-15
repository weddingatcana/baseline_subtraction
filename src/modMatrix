Option Explicit

Public Function matMul(ByRef A#(), _
                       ByRef B#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowB&, _
        rawMaxColB&, _
        i&, j&, k&, _
        C#(), _
        sum#

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowB = UBound(B, 1)
        rawMaxColB = UBound(B, 2)
        
        If rawMaxColA <> rawMaxRowB Then
            Exit Function
        End If
        
        'C(i,k) = A(i,j) * B(j,k)
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColB)
        
        For i = 1 To rawMaxRowA
            For k = 1 To rawMaxColB
                For j = 1 To rawMaxColA
                
                    sum = sum + A(i, j) * B(j, k)
                    
                Next j
                
                C(i, k) = sum
                sum = 0
                
            Next k
        Next i
        
        matMul = C

End Function

Public Function matTra(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, j&, _
        C#()
    
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        ReDim C(1 To rawMaxColA, 1 To rawMaxRowA)
        
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                C(j, i) = A(i, j)
                
            Next j
        Next i

        matTra = C
        
End Function

Public Function matScl(ByRef A#(), _
                       ByVal scalar#) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, j&, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)

        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                C(i, j) = A(i, j) * scalar
            
            Next j
        Next i
        
        matScl = C
        
End Function

Public Function matVec(ByRef A#(), _
                       column&) As Double()

    Dim rawMaxRowA&, _
        i&, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        ReDim C(1 To rawMaxRowA, 1 To 1)
        
        For i = 1 To rawMaxRowA
            C(i, 1) = A(i, column)
        Next i
        
        matVec = C
        
End Function

Public Function matJoin(ByRef A#(), _
                        ByRef B#()) As Double()

    Dim rawMaxRowA, _
        rawMaxColA, _
        rawMaxRowB, _
        rawMaxColB, _
        p&, q&, _
        i&, j&, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowB = UBound(B, 1)
        rawMaxColB = UBound(B, 2)

        If rawMaxRowA = 1 Or rawMaxColA = 1 Then
            If rawMaxRowA > rawMaxColA Then
                p = rawMaxRowA
                j = 0
            Else
                p = rawMaxColA
                j = 1
            End If
        Else
            Exit Function
        End If
                
        If rawMaxRowB = 1 Or rawMaxColB = 1 Then
            If rawMaxRowB > rawMaxColB Then
                q = rawMaxRowB
                j = j + 0
            Else
                q = rawMaxColB
                j = j + 1
            End If
        Else
            Exit Function
        End If
        
        If p <> q Then
            Exit Function
        End If
        
        If j > 0 Then
            Exit Function
        End If
        
        ReDim C(1 To p, 1 To 2)
        
        For i = 1 To p
        
            C(i, 1) = A(i, 1)
            C(i, 2) = B(i, 1)
        
        Next i
        
        matJoin = C

End Function

Public Function matJoin_Ext(ByRef A#(), _
                            ByRef B#()) As Double()
                            
    Dim rawMaxRowA, _
        rawMaxColA, _
        rawMaxRowB, _
        rawMaxColB, _
        p&, q&, _
        i&, j&, _
        max&, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowB = UBound(B, 1)
        rawMaxColB = UBound(B, 2)
        max = (rawMaxColA + rawMaxColB)

        If rawMaxRowA <> rawMaxRowB Then
            Exit Function
        End If
        
        ReDim C(1 To rawMaxRowA, 1 To max)
        
        p = 1
        q = 1
        For j = 1 To max
            For i = 1 To rawMaxRowA
                
                If p > rawMaxColA Then
                    C(i, max) = B(i, q)
                Else
                    C(i, j) = A(i, j)
                End If
            
            Next i
            p = p + 1
        Next j
        
        matJoin_Ext = C
    
End Function

Public Function matInv(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowC&, _
        rawMaxColC&, _
        i&, j&, k&, _
        pivot#, _
        temp#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        If rawMaxRowA <> rawMaxColA Then
            Exit Function
        End If
        
        C = modMatrix.appendIdentity(A)
        rawMaxRowC = UBound(C, 1)
        rawMaxColC = UBound(C, 2)
        
        For i = 1 To rawMaxRowC
            
            pivot = C(i, i)
            For j = 1 To rawMaxColC
                
                If pivot = 0 Then
                    C(i, j) = C(i, j) * pivot
                Else
                    C(i, j) = C(i, j) / pivot
                    
                    'If Abs(C(i, j)) < 1E-50 Then
                        'C(i, j) = 0
                    'End If
                    
                End If
        
            Next j
            
            For k = 1 To rawMaxRowC
                
                temp = C(k, i)
                For j = 1 To rawMaxColC
                    
                    If k <> i Then
                        C(k, j) = C(k, j) - (C(i, j) * temp)
                    End If
            
                Next j
            Next k
            
        Next i
        
        matInv = modMatrix.extractIdentity(C)

End Function

Public Function matReduce(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        If rawMaxColA <> 1 Then
            Exit Function
        End If
        
        ReDim C(1 To rawMaxRowA)
        
        For i = 1 To rawMaxRowA
        
            C(i) = A(i, 1)
        
        Next i
        
        matReduce = C

End Function

Public Function matSpy(ByVal rank&) As Double()

    Dim i&, j&, _
        C#()
        
        ReDim C(1 To rank, 1 To rank)
        
        i = 1
        j = 1
        Do
        
            If (j > rank) And (i > rank) Then
               Exit Do
            End If
            
            C(i, j) = 1
            i = i + 1
            j = j + 1
        
        Loop

        matSpy = C
        
End Function

Public Function matDiff(ByRef A#(), _
                        ByVal delta&) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        copyA#(), _
        i&, j&, k&, _
        C#(), E#(), _
        d&
        
        If delta = 0 Then
            matDiff = A
            Exit Function
        End If
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        ReDim copyA(1 To rawMaxRowA, 1 To rawMaxColA)
        copyA = A
        
        d = 1
        Do
        
            If d > delta Then
                matDiff = C
                Exit Function
            End If
            
            ReDim C(1 To (rawMaxRowA - d), 1 To rawMaxColA)
        
            k = 1
            For i = 2 To rawMaxRowA
                For j = 1 To rawMaxColA
                
                    If k > (rawMaxRowA - d) Then
                        Exit For
                    End If
                    
                    C(k, j) = copyA(i, j) - copyA(i - 1, j)
                
                Next j
                k = k + 1
            Next i

            ReDim copyA(1 To (rawMaxRowA - d), 1 To rawMaxColA)
            copyA = C
            
            d = d + 1
            
        Loop

        matDiff = C

End Function

Public Function matAdd(ByRef A#(), _
                       ByRef B#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowB&, _
        rawMaxColB&, _
        i&, j&, _
        C#()

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowB = UBound(B, 1)
        rawMaxColB = UBound(B, 2)

        If rawMaxRowA <> rawMaxRowB And _
           rawMaxColA <> rawMaxColB Then
           
            Exit Function
            
        End If
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                C(i, j) = A(i, j) + B(i, j)
            
            Next j
        Next i
        
        matAdd = C
        
End Function

Public Function matSub(ByRef A#(), _
                       ByRef B#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowB&, _
        rawMaxColB&, _
        i&, j&, _
        C#()

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowB = UBound(B, 1)
        rawMaxColB = UBound(B, 2)

        If rawMaxRowA <> rawMaxRowB And _
           rawMaxColA <> rawMaxColB Then
           
            Exit Function
            
        End If
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                C(i, j) = A(i, j) - B(i, j)
            
            Next j
        Next i
        
        matSub = C
        
End Function

Public Function matPin(ByRef A#()) As Double()

    Dim T#(), p#(), _
        V#(), C#()
        
        T = modMatrix.matTra(A)
        p = modMatrix.matMul(T, A)
        V = modMatrix.matInv(p)
        C = modMatrix.matMul(V, T)
        
        matPin = C

End Function

Public Function matDim(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        C#(), _
        i&
        
        rawMaxRowA = UBound(A)
        ReDim C(1 To rawMaxRowA, 1 To 1)
        
        For i = 1 To rawMaxRowA
        
            C(i, 1) = A(i)
        
        Next i
        
        matDim = C

End Function

Private Function appendIdentity(ByRef A#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        i&, j&, _
        C#()

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        C = A
        
        ReDim Preserve C(1 To rawMaxRowA, 1 To (2 * rawMaxColA))
        
        j = (rawMaxColA + 1)
        For i = 1 To rawMaxRowA
            
            If j <= (2 * rawMaxColA) Then
                C(i, j) = 1
                j = j + 1
            End If

        Next i
        
        appendIdentity = C
        
End Function

Private Function extractIdentity(ByRef C#()) As Double()

    Dim rawMaxRowC&, _
        rawMaxColC&, _
        i&, j&, _
        d#()
        
        rawMaxRowC = UBound(C, 1)
        rawMaxColC = UBound(C, 2)
        
        ReDim d(1 To rawMaxRowC, 1 To (rawMaxColC \ 2))
        
        For i = 1 To rawMaxRowC
            For j = ((rawMaxColC \ 2) + 1) To rawMaxColC
            
                d(i, j - (rawMaxColC \ 2)) = C(i, j)
            
            Next j
        Next i
        
        extractIdentity = d

End Function
