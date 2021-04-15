Sub Button()
    c = 0
    j = 1
    While Cells(3 + j, 3) <> ""
        c = c + 1
        j = j + 1
    Wend
    a = 0
    b = 0
    Sheet1.Range("J3:T100").ClearContents
    For i = 1 To c
        If Cells(3 + i, 3) < Cells(3 + i, 4) Then
            Cells(3 + a, 10) = Cells(3 + i, 3)
            Cells(3 + a, 11) = i
            a = a + 1
        ElseIf Cells(3 + i, 3) > Cells(3 + i, 4) Then
            Cells(3 + b, 12) = Cells(3 + i, 4)
            Cells(3 + b, 13) = i
            b = b + 1
        Else
            x = Rnd
            If x <= 0.5 Then
                Cells(3 + a, 10) = Cells(3 + i, 3)
                Cells(3 + a, 11) = i
                a = a + 1
            Else
                Cells(3 + b, 12) = Cells(3 + i, 4)
                Cells(3 + b, 13) = i
                b = b + 1
            End If
        End If
    Next i
    For m = 1 To a
        Min = 1000
        For k = 1 To a
            If Min > Cells(2 + k, 10) And Cells(2 + k, 10) <> "" Then
                Min = Cells(2 + k, 10)
                Minw = Cells(2 + k, 11)
                Minp = 2 + k
            ElseIf Min = Cells(2 + k, 10) And Cells(2 + k, 10) <> "" Then
                y = Rnd
                If y <= 0.5 Then
                    Min = Cells(2 + k, 10)
                    Minw = Cells(2 + k, 11)
                    Minp = 2 + k
                End If
            End If
        Next k
        Cells(3 + m, 6) = Minw
        Sheet1.Cells(Minp, 10).ClearContents
        Sheet1.Cells(Minp, 11).ClearContents
    Next m
     For m = 1 To b
        Max = 0
        For k = 1 To b
            If Max < Cells(2 + k, 12) And Cells(2 + k, 12) <> "" Then
                Max = Cells(2 + k, 12)
                Maxw = Cells(2 + k, 13)
                Maxp = 2 + k
            ElseIf Max = Cells(2 + k, 12) And Cells(2 + k, 12) <> "" Then
                y = Rnd
                If y <= 0.5 Then
                    Max = Cells(2 + k, 12)
                    Maxw = Cells(2 + k, 13)
                    Maxp = 2 + k
                End If
            End If
        Next k
        Cells(3 + m + a, 6) = Maxw
        Sheet1.Cells(Maxp, 12).ClearContents
        Sheet1.Cells(Maxp, 13).ClearContents
    Next m
End Sub