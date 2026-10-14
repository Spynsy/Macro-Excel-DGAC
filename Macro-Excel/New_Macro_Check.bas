Attribute VB_Name = "Module1"
'Fichier source
Public Const c_KeyColNb As Integer = 1
Public Const c_IssueTypeColNb As Integer = 4
Public Const c_StatusColNb As Integer = 5

Public Const c_RefKey_RowNb As Integer = 2
Public Const c_SizeBloc As Integer = 15
Public Const c_DateblocSize As Integer = 13
'Public Const c_ClosureColNb As Integer = 21
'Onglet "Liste PCR"
Public Const c_RefDSNAColNb As Integer = 1
Public Const c_DateEnvoiColNb As Integer = 2
Public Const c_StatutPCRColNb As Integer = 3
Public Const c_DateCorrectionColNb As Integer = 4
Public Const c_PeriodeCorrColNb As Integer = 5
Public Const c_RefKeyColNb As Integer = 6
Public Const c_TypeColNb As Integer = 7
Public Const c_FirstStatutColNb As Integer = 8
'Fichier des PCR
Public Const c_StartRowNb As Integer = 2
Public Const c_FixVersionColNb As Integer = 15
Public Const c_ClosureReasonColNb As Integer = 19

Public Const c_Short_Year_Min_Rejected = 11 'Année à ne pas prendre en compte
Public Const c_Long_Year_Min_Rejected = 2011 'Année à ne pas prendre en compte
Public Const c_Day_Step = 0
Public Const c_Month_Step = 1
Public Const c_Year_Step = 2

Public Const c_FileName As String = "CR-4FDTI"
Public Const c_BargraphWidth As Integer = 200 'Taille max du bargraph



Sub GetStatus()

Dim FileToOpen As Variant
Dim wbSource
Dim wsSource
Dim wsLocal
Dim wsLocal_LastRow As Integer
Dim wsSource_LastRow As Integer
Dim lIndexLocal As Integer
Dim lIndexSrce As Integer
Dim lTabSplit() As String
Dim lFileName As String
Dim lYear, lMonth, lDay, lCurrentDate
Dim lCurrentCol As Integer
Dim lComment
'Dim lDate
Dim lOffset As Integer
Dim lExit As Boolean
Dim lStatut As String
Dim ProgressionStep As Integer
Dim ProgressionValue As Integer
Dim Coefficient As Variant
Dim lIndexBargraph As Integer
Dim Tag As String

    'on travaille sur la feuille active
    Set wsLocal = ThisWorkbook.Sheets(ActiveSheet.Name)
    'On détermine le nombre de références spécifiées
    wsLocal_LastRow = wsLocal.Range("F" & Rows.Count).End(xlUp).Row - c_SizeBloc
    'on se positionne sur la dernière colonne remplie
    lCurrentCol = wsLocal.Cells(c_RefKey_RowNb, Columns.Count).End(xlToLeft).Column + 1

    'On récupère le fichier Excel
    FileToOpen = Application.GetOpenFilename(, , "Importer un fichier PCR Thales", "Excel Files (*.xlsx), *.xlsx")
   
    If FileToOpen <> False Then
        'on fige l'afffichage pour ne pas avoir de basculement visible à l'écran entre les deux fichiers Excel
        Application.ScreenUpdating = False
        
        lTabSplit = Split(FileToOpen, "\", , vbTextCompare)
        lFileName = lTabSplit(UBound(lTabSplit))
        If InStr(1, lFileName, c_FileName, vbTextCompare) = 0 Then
            MsgBox "Le nom du fichier est incorrect" & vbCrLf & "Le nom du fichier doit être CR-4FDTI-AAAAMMJJ" & _
            vbCrLf & "AAAA = année" & vbCrLf & "MM = mois" & vbCrLf & "JJ = jour", vbCritical, "Erreur"
        Else
            lTabSplit = Split(lFileName, "_", , vbTextCompare)
            If UBound(lTabSplit) < 1 Then
                MsgBox "Le nom du fichier est incorrect", vbCritical, "Erreur"
            Else
                lRef = lTabSplit(UBound(lTabSplit))
                'AAAAMMDD.xlsx => 13 characters
                If Len(lRef) <> c_DateblocSize Then
                    MsgBox "Le nom du fichier est incorrect", vbCritical, "Erreur"
                Else
                    lYear = Mid(lRef, 1, 4)
                    lMonth = Mid(lRef, 5, 2)
                    lDay = Mid(lRef, 7, 2)
                    lCurrentDate = CDate(lDay & "/" & lMonth & "/" & lYear)
                    'Ouverture du fichier
                    Set wbSource = Workbooks.Open(FileToOpen)
                    Set wsSource = wbSource.Worksheets(1)
                    
                    'on détermine le nombre de cellules à traiter
                    wsSource_LastRow = wsSource.Range("C" & Rows.Count).End(xlUp).Row
                    
                    'On reporte la date du fichier ouvert
                    wsLocal.Cells(c_RefKey_RowNb, lCurrentCol) = lCurrentDate
                    
        BargraphWindow.Show (0)
        
        'Calcul du pas de progression
        lIndexBargraph = 0
        ProgressionStep = ((wsLocal_LastRow - c_RefKey_RowNb) / c_BargraphWidth)
        Coefficient = (100 / c_BargraphWidth)
                    
                    
                    'on recherche le statut pour chaque PCR
                    lIndexLocal = wsLocal_LastRow
                    Do While lIndexLocal > c_RefKey_RowNb
                        lIndexSrce = wsSource_LastRow
                        Do While lIndexSrce > 1
                            If wsSource.Cells(lIndexSrce, c_KeyColNb) = wsLocal.Cells(lIndexLocal, c_RefKeyColNb) Then
                                'Ajout du 05.03.2020
                                'si le PCR est indiqué "DUPLICATION" on positionne la couleur du texte en rouge sur fond jaune
                                If wsSource.Cells(lIndexSrce, c_ClosureReasonColNb).Value = "Duplication" Then
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Font.Color = -16776961
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Interior.Color = 65535
                                Else
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Font.ColorIndex = xlAutomatic
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Interior.Pattern = xlNone
                                End If
                                
                                'Ajout du 18.12.2020 : contrôle du type PCR ou IRF, si l'info a changé, on le signale
                                'par un changement de couleur
                                Tag = ""
                                Select Case wsSource.Cells(lIndexSrce, c_IssueTypeColNb).Value
                                    Case "Information Request"
                                        Tag = "IRF"
                                    Case "Enhancement Request", "Problem Report"
                                        Tag = "PCR"
                                End Select
                                
                                
                                If Tag <> wsLocal.Cells(lIndexLocal, c_TypeColNb) Then
                                    'on le marque en jaune
                                    wsLocal.Cells(lIndexLocal, c_TypeColNb).Font.Color = -16776961
                                    wsLocal.Cells(lIndexLocal, c_TypeColNb).Interior.Color = 65535
                                Else
                                    'on retire tout marquage éventuel
                                    wsLocal.Cells(lIndexLocal, c_TypeColNb).Font.ColorIndex = xlAutomatic
                                    wsLocal.Cells(lIndexLocal, c_TypeColNb).Interior.Pattern = xlNone
                                End If
                                
                            
                                'Mise à jour du statut du PCR/IRF
                                wsLocal.Cells(lIndexLocal, lCurrentCol) = wsSource.Cells(lIndexSrce, c_StatusColNb)
                                
                                'on détermine le statut de correction
                                If ((wsLocal.Cells(lIndexLocal, lCurrentCol).Value = "Closed" Or _
                                   wsLocal.Cells(lIndexLocal, lCurrentCol).Value = "Verified" Or _
                                   wsLocal.Cells(lIndexLocal, lCurrentCol).Value = "Solved") And _
                                   (wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "En cours" Or _
                                   wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "")) Then
                                   
                                    If wsLocal.Cells(lIndexLocal, lCurrentCol).Value = "Closed" Then
                                        If wsSource.Cells(lIndexSrce, c_ClosureReasonColNb).Value = "Rejection" Then
                                            'il est précisé que la fermeture est sur rejet => PCR abandonné
                                            wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Abandonné"
                                            'Modif du 17.12.2020
                                            wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = lCurrentDate
                                        Else
                                            If wsSource.Cells(lIndexSrce, c_ClosureReasonColNb).Value <> "Realisation" Then
                                                'on se fie à la version de correction
                                                If ((wsSource.Cells(lIndexSrce, c_FixVersionColNb).Value <> "") And _
                                                   (wsSource.Cells(lIndexSrce, c_FixVersionColNb).Value <> "na") And _
                                                   (wsSource.Cells(lIndexSrce, c_FixVersionColNb).Value <> "NOT APPLICABLE")) Then
                                                    'une version de correction est précisée => on considère le PCR corrigé
                                                    wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigé"
                                                Else
                                                    'pas de version de correction de précisée => on considère le PCR abandonné
                                                    wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Abandonné"
                                                End If
                                            Else
                                                'il est précisé que la fermeture est issue d'une réalisation => on considère
                                                'le PCR corrigé
                                                wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigé"
                                            End If
                                        End If
                                        'Modif du 17.12.2020 : la statut a été modifié, la date doit être précisée.
                                        wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = lCurrentDate
                                    Else
                                        'Modif du 03/07/2020 : Simplification du test
                                        If wsSource.Cells(lIndexSrce, c_IssueTypeColNb).Value <> "Information Request" Then
                                            'on entre dans cette séquence car le statut courant est "Solved" ou "Verified"
                                            wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigé"
                                        Else
                                            'Modif du 27/05/2020
                                            'On tient compte des IRF "Verified ou "Closed" => IRF CORRIGEES
                                            If wsLocal.Cells(lIndexLocal, lCurrentCol).Value = "Verified" Then
                                                wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigé"
                                            Else
                                                'Cas d'une IRF pour laquelle on ne peut pas affirmer qu'elle soit corrigée
                                                'pas d'opération de datation de la correction.
                                                wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "En cours"
                                            End If
                                        End If
                                        'Modif du 17.12.2020
                                        If wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigé" Then
                                            'PCR/IRF corrigé, on indique la date
                                            wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = lCurrentDate
                                        Else
                                            'on efface la date, le statut est "En cours"...
                                            wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = ""
                                        End If
                                    End If
                                    
                                    'Modif du 03/07/2020 : Datation de la correction
                                    'If wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value <> "En cours" Then
                                    '
                                    '    lIndex = lCurrentCol
                                    '    lExit = False
                                    '    Do
                                    '        Select Case wsLocal.Cells(lIndexLocal, lIndex).Value
                                    '
                                    '            Case "Submitted", "Analysed", "Accepted", "StartedWork", ""
                                    '                lExit = True
                                    '
                                    '            Case "Solved"
                                    '                'on remonte l'historique jusqu'à un autre état
                                    '                lOffset = 1
                                    '                Do While wsLocal.Cells(lIndexLocal, lIndex).Value = "Solved"
                                    '                    If lIndex > c_FirstStatutColNb Then
                                    '                        lIndex = lIndex - 1
                                    '                    Else
                                    '                        lExit = True
                                    '                        lOffset = 0
                                    '                        Exit Do
                                    '                    End If
                                    '                Loop
                                    '
                                    '            Case "Verified"
                                    '                'on remonte l'historique jusqu'à un autre état
                                    '                lOffset = 1
                                    '                Do While wsLocal.Cells(lIndexLocal, lIndex).Value = "Verified"
                                    '                    If lIndex > c_FirstStatutColNb Then
                                    '                        lIndex = lIndex - 1
                                    '                    Else
                                    '                        lExit = True
                                    '                        lOffset = 0
                                    '                        Exit Do
                                    '                    End If
                                    '                Loop
                                    '
                                    '            Case "Rejected"
                                    '                lExit = True
                                    '
                                    '            Case "Closed"
                                    '                lOffset = 1
                                    '                Do While wsLocal.Cells(lIndexLocal, lIndex).Value = "Closed"
                                    '                    If lIndex > c_FirstStatutColNb Then
                                    '                        lIndex = lIndex - 1
                                    '                    Else
                                    '                        lOffset = 0
                                    '                        lExit = True
                                    '                        Exit Do
                                    '                    End If
                                    '                Loop
                                    '        End Select
                                    '    Loop Until lExit = True
                                    '
                                    '    'on prend la date du fichier PCR correspondant à la position actuelle
                                    '    wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = wsLocal.Cells(2, lIndex + lOffset).Value
                                        
                                        'Ne pas effacer. Récupération de la date dans le commentaire de fermeture dans JIRA
                                        'si nécessaire
                                        'wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "Corrigée"
                                        'on récupère la date
                                        'lDate = GetDate(wsSource.Cells(lIndexSrce, c_ClosureColNb))
                                        'wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = lDate
                                    'End If
                                Else
                                    If wsLocal.Cells(lIndexLocal, lCurrentCol).Value <> "Closed" And _
                                       wsLocal.Cells(lIndexLocal, lCurrentCol).Value <> "Verified" And _
                                       wsLocal.Cells(lIndexLocal, lCurrentCol).Value <> "Solved" Then
                                        'PCR est ouvert ou ré-ouvert
                                        wsLocal.Cells(lIndexLocal, c_StatutPCRColNb).Value = "En cours"
                                        'Modif du 03/07/2020
                                        'On efface la date de correction éventuelle
                                        wsLocal.Cells(lIndexLocal, c_DateCorrectionColNb).Value = ""
                                    End If
                                End If
                                
                                'Gestion des commentaires de cellule
                                Set lComment = wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Comment
                                If lComment Is Nothing Then
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).AddComment
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Comment.Visible = False
                                    wsLocal.Cells(lIndexLocal, c_RefKeyColNb).Comment.Text Text:=wsSource.Cells(lIndexSrce, 3).Value
                                End If
                                lIndexSrce = 0
                            Else
                                lIndexSrce = lIndexSrce - 1
                            End If
                        Loop
                        lIndexLocal = lIndexLocal - 1
                        
            'Affichage de la progression
            lIndexBargraph = lIndexBargraph + 1
            ProgressionValue = lIndexBargraph / ProgressionStep
            If ProgressionValue > c_BargraphWidth Then
                ProgressionValue = c_BargraphWidth
            End If
            BargraphWindow.ProgressionLabel = Int((ProgressionValue * Coefficient)) & "% processed..."
            BargraphWindow.Bargraph.Width = ProgressionValue
            DoEvents
                        
                    Loop
                    Unload BargraphWindow
                    'Traitement terminé
                    Application.DisplayAlerts = False
                    wbSource.Close
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    MsgBox "Traitement terminé."
                End If
            End If
        End If
    End If
End Sub

Private Function GetDate(ByVal Target As Range) As Date
'***************************************
' Nom : GetDate()
' Fonction : Retourne la date au format JJ/MM/AAAA la plus récente trouvée dans une cellule passée en paramètre.
' Date de creation : 29/09/2017
' ------------------------
' Suivi des modifications
' 06/12/2017: Un seuil mini de prise en compte des années est fixé par défaut à 2011. Toutes les dates
'             inferieures de sont pas prises en compte.
' 24/01/2018: Les dates lues dans les cellules qui sont supérieures à la date du jour sont ignorées
'***************************************

Dim MyBuffer() As String
Dim MyBufferElement
Dim NbDates As Integer
Dim lDay(20) As String
Dim lMounth(20) As String
Dim lYear(20) As String
Dim lIndex As Integer
Dim lStep As Integer
Dim lStr As String
Dim lDate1 As Date
Dim lDate2 As Date
Dim i As Integer
Dim Short_Year_Max_Rejected As String
Dim Long_Year_Max_Rejected As String


'Calcul de la date du jour
Short_Year_Max_Rejected = Right(Year(Date), 2)
Long_Year_Max_Rejected = Year(Date)

If Target.Value <> "" Then
    If Target.Value Like "*##/##/##*" Then
        'Date présente
        MyBuffer = Split(Target.Value, "/")
        'Récupération de ou des dates à partir des bribes de chaines
        lStep = c_Day_Step
        lIndex = 0
        For Each MyBufferElement In MyBuffer
            
            MyBufferElement = Replace(MyBufferElement, " ", "_")
            MyBufferElement = Replace(MyBufferElement, Chr(10), "_")
            MyBufferElement = Replace(MyBufferElement, Chr(13), "_")
        
            'Traitement du jour
            If lStep = c_Day_Step Then
                'Traitement du cas DD/
                lStr = Right(MyBufferElement, 2)
                If IsNumeric(lStr) Then
                    If CInt(lStr) <= 31 Then
                        lDay(lIndex) = lStr
                        lStep = c_Month_Step
                    Else
                        lStep = c_Day_Step
                    End If
                Else
                    lStep = c_Day_Step
                End If
            'Traitement du mois
            ElseIf lStep = c_Month_Step Then
                'Traitement du /MM/
                lStr = Right(MyBufferElement, 2)
                If IsNumeric(lStr) Then
                    If CInt(lStr) <= 12 Then
                        lMounth(lIndex) = lStr
                        lStep = c_Year_Step
                    Else
                        lStep = c_Day_Step
                    End If
                Else
                    lStep = c_Day_Step
                End If
            Else
                'Traitement de l'année
                lStep = c_Day_Step
                'Traitement du cas /YYYY
                lStr = Left(MyBufferElement, 4)
                If IsNumeric(lStr) = False Then
                    'Traitement du cas /YY
                    lStr = Left(MyBufferElement, 2)
                    If IsNumeric(lStr) Then
                        'Fitrage des années
                        'Modif 06/12/2017 : ajout d'un seuil c_Short_Year_Min_Rejected
                        'Modif 24/01/2018 : ajout du seuil max
                        If (CInt(lStr) > c_Short_Year_Min_Rejected) And (CInt(lStr) <= Short_Year_Max_Rejected) Then
                            lYear(lIndex) = lStr
                            lIndex = (lIndex + 1) Mod 20
                        End If
                    End If
                Else
                    'Filtrage des années
                    'Modif 06/12/2017 : ajout d'un seuil c_Long_Year_Min_Rejected
                    'Modif 24/01/2018 : ajout du seuil max
                    If CInt(lStr) > c_Long_Year_Min_Rejected And (CInt(lStr) <= Long_Year_Max_Rejected) Then
                        lYear(lIndex) = lStr
                        lIndex = (lIndex + 1) Mod 20
                    End If
                End If
                
                'On s'attaque à la fin de la chaine
                lStr = Right(MyBufferElement, 4)
                If IsNumeric(lStr) = False Then
                    lStr = Right(MyBufferElement, 2)
                    If IsNumeric(lStr) Then
                        If CInt(lStr) <= 31 Then
                            'nouveau cycle de détection DD/MM/YYYY ou /YY
                            lDay(lIndex) = lStr
                            'prochaine étape => le mois
                            lStep = c_Month_Step
                        End If
                    End If
                End If
            End If
        Next
    End If
End If
lDate1 = 0
lDate2 = 0
For i = 0 To lIndex - 1
    lStr = lDay(i) & "/" & lMounth(i) & "/" & lYear(i)
    lDate1 = CDate(lStr)
    If lDate1 > lDate2 Then
        lDate2 = lDate1
    End If
Next i
GetDate = lDate2
End Function

Sub GetType()

Dim FileToOpen As Variant
Dim wbSource
Dim wsSource
Dim wsLocal
Dim wsLocal_LastRow As Integer
Dim wsSource_LastRow As Integer
Dim lIndexLocal As Integer
Dim lIndexSrce As Integer
Dim lTabSplit() As String
Dim lFileName As String
Dim lYear, lMonth, lDay, lCurrentDate
Dim lCurrentCol As Integer
Dim lComment
'Dim lDate
Dim lOffset As Integer
Dim lExit As Boolean
Dim lStatut As String
Dim ProgressionStep
Dim ProgressionValue As Integer
Dim Coefficient As Variant
Dim lIndexBargraph As Integer
Dim lIndex


    'on fige l'affichage pour ne pas avoir de basculement visible à l'écran entre les deux fichiers Excel
    Application.ScreenUpdating = False
    'on travaille sur la feuille active
    Set wsLocal = ThisWorkbook.Sheets(ActiveSheet.Name)
    'On détermine le nombre de références spécifiées
    wsLocal_LastRow = wsLocal.Range("A" & Rows.Count).End(xlUp).Row

    'On récupère le fichier Excel
    FileToOpen = Application.GetOpenFilename(, , "Importer un fichier PCR Thales", "Excel Files (*.xlsx), *.xlsx")
   
    If FileToOpen <> False Then
        lTabSplit = Split(FileToOpen, "\", , vbTextCompare)
        lFileName = lTabSplit(UBound(lTabSplit))
        If InStr(1, lFileName, c_FileName, vbTextCompare) = 0 Then
            MsgBox "Le nom du fichier est incorrect" & vbCrLf & "Le nom du fichier doit être CR-4FDTI-AAAAMMJJ" & _
            vbCrLf & "AAAA = année" & vbCrLf & "MM = mois" & vbCrLf & "JJ = jour", vbCritical, "Erreur"
        Else
            lTabSplit = Split(lFileName, "_", , vbTextCompare)
            If UBound(lTabSplit) < 1 Then
                MsgBox "Le nom du fichier est incorrect", vbCritical, "Erreur"
            Else
                lRef = lTabSplit(UBound(lTabSplit))
                'AAAAMMDD.xlsx => 13 characters
                If Len(lRef) <> c_DateblocSize Then
                    MsgBox "Le nom du fichier est incorrect", vbCritical, "Erreur"
                Else
                    lYear = Mid(lRef, 1, 4)
                    lMonth = Mid(lRef, 5, 2)
                    lDay = Mid(lRef, 7, 2)
                    lCurrentDate = CDate(lDay & "/" & lMonth & "/" & lYear)
                    'Ouverture du fichier
                    Set wbSource = Workbooks.Open(FileToOpen)
                    Set wsSource = wbSource.Worksheets(1)
                    
                    'on détermine le nombre de cellules à traiter
                    wsSource_LastRow = wsSource.Range("C" & Rows.Count).End(xlUp).Row
   
                    'Calcul du pas de progression proportionnel au nombre de lignes que nous devons traiter
                    ProgressionStep = (c_BargraphWidth / wsLocal_LastRow)
                    
                    'on recherche chaque référence dans le fichier des PCR
                    lIndexLocal = wsLocal_LastRow
                    lIndex = 0
                    Do While lIndexLocal > 2
                        lIndexSrce = wsSource_LastRow
                        Do While lIndexSrce > 1
                            If wsSource.Cells(lIndexSrce, 3).Value = wsLocal.Cells(lIndexLocal, 1).Value Then
                                If wsSource.Cells(lIndexSrce, 5).Value = "Problem Report" Then
                                    wsLocal.Cells(lIndexLocal, 2).Value = "PCR"
                                Else
                                    If wsSource.Cells(lIndexSrce, 5).Value = "Information Request" Then
                                        wsLocal.Cells(lIndexLocal, 2).Value = "IRF"
                                        If InStr(1, wsSource.Cells(lIndexSrce, 7), "COFLIGHT", vbTextCompare) > 0 Then
                                            wsLocal.Cells(lIndexLocal, 3).Value = "COFLIGHT"
                                        End If
                                    Else
                                        wsLocal.Cells(lIndexLocal, 2).Value = wsSource.Cells(lIndexSrce, 5).Value
                                    End If
                                End If
                                lIndexSrce = 0
                            Else
                                lIndexSrce = lIndexSrce - 1
                            End If
                            DoEvents
                        Loop
                        lIndexLocal = lIndexLocal - 1
                        'Affichage de la progression
                        lIndex = lIndex + 1
                        ProgressionValue = lIndex * ProgressionStep
                        BargraphWindow.Bargraph.Width = ProgressionValue
                        'Affichage du bargraph de progression
                        BargraphWindow.Show (0)
                    Loop
                End If
            End If
        End If
                
        'Traitement terminé
        Application.DisplayAlerts = False
        wbSource.Close
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "Traitement terminé."
    End If
End Sub


