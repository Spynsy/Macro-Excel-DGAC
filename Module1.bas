Attribute VB_Name = "Module1"
Sub CompacterColonne()

    Dim Plage As Range
    Dim Cellule As Range
    Dim LigneDest As Long

    Set Plage = Range("F2:F4000")
        
    Dim LigneDestA As Long
    Dim LigneDestG As Long
    
    Range("G:G").ClearContents
        
    LigneDestA = Cells(Rows.Count, "A").End(xlUp).Row + 1
    LigneDestG = 1
    
    For Each Cellule In Plage
        If Trim(Cellule.Value) <> "" Then
    
            Cells(LigneDestG, "G").Value = Cellule.Value
            Cells(LigneDestA, "A").Value = Cellule.Value
    
            LigneDestG = LigneDestG + 1
            LigneDestA = LigneDestA + 1
    
        End If
    Next Cellule
    
    For Each Cellule In Plage
        Cellule.FormulaLocal = "=SI(D" & Cellule.Row & "="""";"""";SI(NB.SI(A$2:A$4000;D" & Cellule.Row & ")=0;D" & Cellule.Row & ";""""))"
    Next Cellule

End Sub
