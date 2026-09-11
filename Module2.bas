Attribute VB_Name = "Module2"
Sub attribuer_FFT_PCR()
    Dim Plage As Range
    Dim Cellule As Range
    
    Dim wsRef As Worksheet
    
    Dim Resultat As Variant
    Dim MyDate
    MyDate = Date
    
    Set wsRef = Worksheets("Nbres analyses2")

    For Each Cellule In Worksheets("Work-Analyses").Range("B2:B4000")

        If Trim(Cellule.Value) = "" Then

            Resultat = Application.VLookup(Cellule.Offset(0, -1).Value, wsRef.Range("B:M"), 12, False)
            
            If Not IsError(Resultat) Then

                Cellule.Value = Resultat
                Cellule.Offset(0, 1) = MyDate
            End If

        End If

    Next Cellule
End Sub

Sub Determiner_Type_Correction()
    Dim Plage As Range
    Dim Cellule As Range
    Dim valeur As Range
    
    Dim wsRef As Worksheet
    
    Dim Resultat As Variant
    Dim Coflight As Variant
    Dim MyDate
    MyDate = Date

    Set wsRef = Worksheets("Nbres analyses2")
    
    For Each Cellule In Worksheets("Corrections PCR-IRF").Range("B2:B4000")
        If Trim(Cellule.Value) = "" Then
            
            If Cellule.Offset(0, -1) <> "" Then 'Regarde si une PCR existe'
                Resultat = Application.VLookup(Cellule.Offset(0, -1).Value, _
                wsRef.Range("M:V"), 6, False)
                
                If Not IsError(Resultat) Then
                
                    If Resultat = "Hardware" Then
                        Cellule.Value = "PCR"
                        Cellule.Offset(0, 2) = "HW"
                        Cellule.Offset(0, 3) = "En cours"
                        Cellule.Offset(0, 5) = MyDate
                        Cellule.Offset(0, 6) = "0"
                        
                    ElseIf Resultat = "Documentaire" Then
                        Cellule.Value = "PCR"
                        Cellule.Offset(0, 2) = "DOC"
                        Cellule.Offset(0, 3) = "En cours"
                        Cellule.Offset(0, 5) = MyDate
                        Cellule.Offset(0, 6) = "0"
                        
                    ElseIf Resultat = "Software" Then
                    
                        Coflight = Application.VLookup(Cellule.Offset(0, -1).Value, _
                        wsRef.Range("M:V"), 10, False)
                        
                        If Trim(Coflight) = "4F" & Chr(10) & "COFLIGHT" Then
                            Cellule.Value = "PCR"
                            Cellule.Offset(0, 2) = "SW"
                            Cellule.Offset(0, 1) = "V2COF"
                            Cellule.Offset(0, 3) = "En cours"
                            Cellule.Offset(0, 5) = MyDate
                            Cellule.Offset(0, 6) = "0"
                        
                        
                        ElseIf Coflight = "4F" Or _
                        Trim(Coflight) = "4F" & Chr(10) & "DART" Or _
                        Trim(Coflight) = "4F" & Chr(10) & "JHMI" Or _
                        Trim(Coflight) = "4F" & Chr(10) & "STS 4F" _
                        Or Trim(Coflight) = "4F" & Chr(10) & "ATLAS" _
                        Or Trim(Coflight) = "4F" & Chr(10) & "RESMS" Or _
                        Trim(Coflight) = "4F" & Chr(10) & "SWIFT" Or _
                        Trim(Coflight) = "4F" & Chr(10) & "STM" Then
            
                            Cellule.Value = "PCR"
                            Cellule.Offset(0, 2) = "SW"
                            Cellule.Offset(0, 3) = "En cours"
                            Cellule.Offset(0, 5) = MyDate
                            Cellule.Offset(0, 6) = "0"
                        
                        Else
                            Cellule.Value = ""
                        
                        End If
                    End If
                End If
            End If
        End If
    Next Cellule
                    
End Sub

