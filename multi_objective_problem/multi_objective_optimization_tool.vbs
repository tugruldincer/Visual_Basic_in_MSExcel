Attribute VB_Name = "Module2"
Sub Knapsack()

For k = 1 To 25
Sheet4.Cells(9, k + 1) = Sheet3.Cells(3, k + 2)
Sheet4.Cells(6, k + 1) = Sheet3.Cells(5, k + 2)
Sheet4.Cells(7, k + 1) = Sheet3.Cells(6, k + 2)
Sheet4.Cells(8, k + 1) = Sheet3.Cells(7, k + 2)
Sheet4.Cells(4, k + 1) = Sheet3.Cells(2, k + 2)
Next k
Sheet4.Cells(6, 29) = Sheet3.Cells(5, 29)
Sheet4.Cells(7, 29) = Sheet3.Cells(6, 29)
Sheet4.Cells(8, 29) = Sheet3.Cells(7, 29)
Sheet4.Cells(9, 29) = 0

Sheet4.Cells(6, 27) = Sheet3.Cells(5, 28)
Sheet4.Cells(7, 27) = Sheet3.Cells(6, 28)
Sheet4.Cells(8, 27) = Sheet3.Cells(7, 28)
If Sheet3.Cells(3, 2) = "Max" Then
    Sheet4.Cells(9, 27) = ">="
Else: Sheet4.Cells(9, 27) = "<="
End If

For c = 0 To 29

Sheet4.Range("A12") = "=SUMPRODUCT(B3:Z3,B4:Z4)"
Sheet4.Range("AB6") = "=SUMPRODUCT(B3:Z3,B6:Z6)"

Sheet4.Range("AB7") = "=SUMPRODUCT(B3:Z3,B7:Z7)"
Sheet4.Range("AB8") = "=SUMPRODUCT(B3:Z3,B8:Z8)"
Sheet4.Range("AB9") = "=SUMPRODUCT(B3:Z3,B9:Z9)"

SolverReset
    SolverOk SetCell:="$A$12", MaxMinVal:=1, ValueOf:=0, ByChange:="$B$3:$Z$3", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$AB$6", Relation:=1, FormulaText:="$AC$6"
    SolverAdd CellRef:="B$3:Z$3", Relation:=4
    SolverAdd CellRef:="$AB$7", Relation:=1, FormulaText:="$AC$7"
    SolverAdd CellRef:="$AB$8", Relation:=1, FormulaText:="$AC$8"
    SolverAdd CellRef:="$AB$9", Relation:=3, FormulaText:="$AC$9"
    SolverOk SetCell:="$A$12", MaxMinVal:=1, ValueOf:=0, ByChange:="$B$3:$Z$3", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$A$12", MaxMinVal:=1, ValueOf:=0, ByChange:="$B$3:$Z$3", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve (True)
    
    If SolverSolve(UserFinish:=True) = 5 Then
        GoTo out
    End If
    
    Sheet4.Cells(13 + c, 3) = Sheet4.Cells(12, 1)
    Sheet4.Cells(13 + c, 2) = Sheet4.Cells(9, 29)
    Sheet4.Cells(9, 29) = Sheet4.Cells(9, 28) + 1
    

    For i = 1 To 25
        Sheet4.Cells(13 + c, 4 + i) = Sheet4.Cells(3, 1 + i)
    Next i
     
Next c

out:
    
End Sub










 





