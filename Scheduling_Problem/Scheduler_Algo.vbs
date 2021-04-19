Attribute VB_Name = "Module1"
Sub SchedulingProb()
    Dim delta As Double
    delta = 0
    For i = 2 To 11
    delta = delta + Sheet2.Cells(i, 2)
    Next i
    
For x = 0 To 7
    C = 0
    Dim Di As Double
    For i = 2 To 11
    Sheet3.Cells(10 + 17 * x - 3, i).Value = Sheet2.Cells(i, 3) + delta
    Sheet3.Cells(5 + 17 * x, i).Value = Sheet2.Cells(i, 2)
    Sheet3.Cells(6 + 17 * x, i).Value = Sheet2.Cells(i, 3)
    Next i
    Sheet3.Cells(17 * x + 2, 2) = delta

Dim R As Double
    R = 0
    For i = 2 To 11
        R = R + Sheet2.Cells(i, 2)
    Next i
    
    Dim pistar As Double
    Dim distar As Double
        istar = 0
    For m = 1 To 10
        pistar = 0
        distar = 0
        For i = 2 To 11
        If Sheet3.Cells(5 + 17 * x, i) > pistar And Sheet3.Cells(7 + 17 * x, i) >= R Then
                distar = Sheet3.Cells(7 + 17 * x, i)
                pistar = Sheet3.Cells(5 + 17 * x, i)
                istar = Sheet3.Cells(4 + 17 * x, i)
    ElseIf Sheet3.Cells(5 + 17 * x, i) = pistar And Sheet3.Cells(7 + 17 * x, i) >= R Then
        If Sheet3.Cells(7 + 17 * x, i) >= distar Then
            distar = Sheet3.Cells(7 + 17 * x, i)
                    pistar = Sheet3.Cells(5 + 17 * x, i)
                    istar = Sheet3.Cells(4 + 17 * x, i)
                End If
        End If
        Next i
        If istar = 0 Then
            GoTo nochoice
        Else
            Sheet3.Cells(7 + 17 * x, istar + 1) = 0
            Sheet3.Cells(17 * x + 9, m + 1) = R
            Sheet3.Cells(17 * x + 10, m + 1) = istar
            Sheet3.Cells(17 * x + 11, m + 1) = pistar
            R = R - pistar
            Sheet3.Cells(13 + 17 * x, istar + 1) = 11 - m
        End If
    Next m
    
    For i = 1 To 10
        For t = 2 To 11
            If Sheet3.Cells(13 + 17 * x, t) = i Then
                Sheet3.Cells(14 + 17 * x, t) = C + Sheet3.Cells(5 + 17 * x, t)
                C = C + Sheet3.Cells(5 + 17 * x, t)
            End If
        Next t
    Next i
    
    For i = 2 To 11
        Sheet3.Cells(15 + 17 * x, i).Value = Sheet2.Cells(i, 3)
    Next i
    
    For i = 2 To 11
        tsk = Sheet3.Cells(14 + 17 * x, i) - Sheet3.Cells(15 + 17 * x, i)
        If tsk >= 0 Then
            Sheet3.Cells(16 + 17 * x, i) = tsk
        Else
            Sheet3.Cells(16 + 17 * x, i) = 0
        End If
    Next i
    
    Dim hsum As Double
    hsum = 0
    For i = 2 To 11
        hsum = Sheet3.Cells(14 + 17 * x, i) + hsum
    Next i
    Sheet3.Cells(14 + 17 * x, 13) = hsum
    
    Tard = 0
    For i = 2 To 11
        If Sheet3.Cells(16 + 17 * x, i) > Tard Then
            Tard = Sheet3.Cells(16 + 17 * x, i)
        End If
    Next i
    Sheet3.Cells(16 + 17 * x, 13) = Tard
     
    For i = 2 To 11
    Sheet3.Cells(10 + 17 * x - 3, i).Value = Sheet2.Cells(i, 3) + delta
    Sheet3.Cells(5 + 17 * x, i).Value = Sheet2.Cells(i, 2)
    Sheet3.Cells(6 + 17 * x, i).Value = Sheet2.Cells(i, 3)
    Next i
    Sheet3.Cells(17 * x + 9, m + 1) = R
    delta = Tard - 1
    
    Sheet4.Cells(x + 2, 1) = x + 1
    Sheet4.Cells(x + 2, 2) = Sheet3.Cells(14 + 17 * x, 13)
    Sheet4.Cells(x + 2, 3) = Sheet3.Cells(16 + 17 * x, 13)
    
    Dim sequence As String
    sequence = ""
    For i = 10 To 1 Step -1
    If i <> 1 Then
            sequence = sequence & Sheet3.Cells(17 * x + 10, i + 1) & "-"
    Else
        sequence = sequence & Sheet3.Cells(17 * x + 10, i + 1)
    End If
    Next i
    Sheet4.Cells(x + 2, 4) = sequence
    Dim Message As String
    Dim Number As Integer
    Number = x + 1
    Message = "There exists " & Number & " efficient solutions."
    
Next x
nochoice:
    MsgBox (Message)
End Sub







