Sub Solve()
Sheet1.Range("G4:J7").ClearContents
Application.ScreenUpdating = False
Cells(1, 13) = "w1"
Cells(1, 14) = "w2"
Cells(1, 15) = "x1"
Cells(1, 16) = "x2"
Cells(1, 17) = "f1"
Cells(1, 18) = "f2"
Cells(1, 19) = "wf"
For i = 0 To 50
    w1 = i / 50
    w2 = 1 - w1
    Cells(11, 1) = w1
    Cells(11, 2) = w1
    Cells(11, 3) = w2
    Cells(10, 1) = 1
    Cells(10, 2) = -1
    Cells(10, 3) = 1
    Cells(13, 1) = 1
    Cells(13, 2) = 2
    Cells(14, 1) = 2
    Cells(14, 2) = 1
    Cells(15, 1) = 1
    Cells(15, 2) = 1
    Cells(16, 1) = 1
    Cells(16, 2) = -1
    Cells(17, 1) = -1
    Cells(17, 2) = 1
    Cells(18, 1) = 1
    Cells(18, 2) = 2
    Cells(19, 1) = 1
    Cells(19, 2) = -3
    Cells(20, 1) = 2
    Cells(20, 2) = -1
    Cells(13, 4) = 12
    Cells(14, 4) = 12
    Cells(15, 4) = 7
    Cells(16, 4) = 9
    Cells(17, 4) = 9
    Cells(18, 4) = 0
    Cells(19, 4) = 4
    Cells(20, 4) = 10
    Cells(12, 1) = Cells(11, 1) * Cells(10, 1)
    Cells(12, 2) = Cells(11, 2) * Cells(10, 2) + Cells(11, 3) * Cells(10, 3)
    Range("E10").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
    "=SUMPRODUCT(R[2]C[-4]:R[2]C[-3],R[11]C[-4]:R[11]C[-3])"
    Range("E11").Select
    Range("C13").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUMPRODUCT(RC[-2]:RC[-1],R21C1:R21C2)"
    Range("C13").Select
    Selection.AutoFill Destination:=Range("C13:C20")
    Range("C13:C20").Select
        SolverReset
    SolverAdd CellRef:="$C$13", Relation:=1, FormulaText:="$D$13"
    SolverAdd CellRef:="$C$14", Relation:=1, FormulaText:="$D$14"
    SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$D$15"
    SolverAdd CellRef:="$C$16", Relation:=1, FormulaText:="$D$16"
    SolverAdd CellRef:="$C$17", Relation:=1, FormulaText:="$D$17"
    SolverAdd CellRef:="$C$18", Relation:=3, FormulaText:="$D$18"
    SolverAdd CellRef:="$C$19", Relation:=1, FormulaText:="$D$19"
    SolverAdd CellRef:="$C$20", Relation:=1, FormulaText:="$D$20"
    SolverOk SetCell:="$E$10", MaxMinVal:=1, ValueOf:=0, ByChange:="$A$21:$B$21", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve UserFinish:=True
    Cells(i + 2, 13) = w1
    Cells(i + 2, 14) = w2
    Cells(i + 2, 15) = Cells(21, 1).Value
    Cells(i + 2, 16) = Cells(21, 2).Value
    Cells(i + 2, 17) = Cells(21, 1).Value - Cells(21, 2).Value
    Cells(i + 2, 18) = Cells(21, 2).Value
    Cells(i + 2, 19) = Cells(10, 5).Value
    Sheet1.Range("A10:G30").ClearContents
Next i
rownum = 1
nrow = 0
For i = 2 To 52
    f1 = Cells(i, 17)
    f2 = Cells(i, 18)
    x1 = Cells(i, 15)
    x2 = Cells(i, 16)
    same = 0
    For j = 1 To rownum
        If x1 = Cells(j, 20) And x2 = Cells(j, 21) And f1 = Cells(j, 22) And f2 = Cells(j, 23) Then
            same = 1
            Cells(j, 24) = Cells(j, 24).Value + 1
            Exit For
        End If
    Next j
        If same = 0 Then
            nrow = nrow + 1
            Cells(nrow, 24) = 1
            Cells(nrow, 22) = f1
            Cells(nrow, 23) = f2
            Cells(nrow, 20) = x1
            Cells(nrow, 21) = x2
        End If
    rownum = nrow
Next i
   Columns("T:X").Sort key1:=Range("X1"), _
      order1:=xlDescending
For i = 1 To 4
    For j = 1 To 4
        Cells(i + 3, j + 6) = Cells(i, j + 19)
    Next j
Next i
Sheet1.Range("M1:X30").ClearContents
Application.ScreenUpdating = True
End Sub