``` VBA
Sub PayrollHr()
    
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim lastColumnV2 As Long
    Dim lastRowV2 As Long
    Dim lastRowV3 As Long
    Dim lastRowV4 As Long
    Dim rowNum As Integer
    Dim cellAddress As String
    Dim row As Long
    Dim ws As Worksheet
    Dim wsV2 As Worksheet
    Dim cell As Range
    Dim cellV2 As Range
    Dim element As Variant
    Dim elementV2 As Variant
    Dim sousTotalHeures As Long
    Dim salaireBrut As Long
    Dim totalDesRetenues As Long
    Dim matEmployee As String
    Dim cellValue1 As Variant
    Dim cellValue2 As Variant
    Dim cellValue3 As Variant
    Dim firstSheet As String
    
'Récupération du nombre de colonnes et de lignes
    
    firstSheet = Sheets(1).Name
    Set ws = ActiveWorkbook.Sheets(firstSheet)
    
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
'Récupération numéro de ligne  de chaque valeur à récupérer
   
    For Each cell In ws.Range("B:B").Cells
        element = cell.Value
        If cell.row <= lastRow Then
            If element = "Sous-total Heures hors contrat" Then
                cellAddress = cell.Address
                sousTotalHeures = cell.row
            ElseIf element = "Salaire Brut" Then
                cellAddress = cell.Address
                salaireBrut = cell.row
                cellAddress = "C" & salaireBrut
                ws.Range(cellAddress).Delete Shift:=xlToLeft
            ElseIf element = "Total des retenues" Then
                cellAddress = cell.Address
                totalDesRetenues = cell.row
                'suppression de la cellule en C pour que tout soit sur la même ligne
                cellAddress = "C" & totalDesRetenues
                ws.Range(cellAddress).Delete Shift:=xlToLeft
            End If
        End If
    Next cell
'Split des entêtes de colonne
    For Each cell In ws.Range("3:3").Cells
        If cell.column <= lastColumn Then
            ws.Range(cell.Address).UnMerge
        End If
    Next cell
'Récupération valeur de chaque colonne avec matricule et ajout dans un dictionnaire avec collection
    activePath = ActiveWorkbook.Path

    Workbooks.Open (activePath & "\Temps_de_travail_effectif.xlsx")
    firstSheet = Sheets(1).Name
    Set wsV2 = ActiveWorkbook.Sheets(firstSheet)
    
    lastRowV2 = wsV2.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastColumnV2 = wsV2.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column + 1
    lastColumnV3 = wsV2.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column + 2
    lastColumnV4 = wsV2.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column + 3
    
    ActiveWorkbook.Close savechanges:=True
    
    
    For Each cell In ws.Range("3:3").Cells
        element = cell.Value
        If cell.column <= lastColumn Then
            If element <> "" Then
                matEmployee = Left(element, 5) 'matricule
                'valeur pour un employé
                cellValue1 = ws.Cells(sousTotalHeures, cell.column).Value
                cellValue2 = ws.Cells(salaireBrut, cell.column).Value
                cellValue3 = ws.Cells(totalDesRetenues, cell.column).Value
                'Valeur mise dans le fichier de destination
                
                Workbooks.Open (activePath & "\Temps_de_travail_effectif.xlsx")
                firstSheet = Sheets(1).Name
                Set wsV2 = ActiveWorkbook.Sheets(firstSheet)
                For Each cellV2 In wsV2.Range("C2:C" & lastRowV2).Cells
                    elementV2 = cellV2.Value
                        If elementV2 = matEmployee Then
                            If IsNumeric(cellValue1) Then
                                cellValue1 = cellValue1 * 1.42
                                wsV2.Cells(cellV2.row, lastColumnV2).Value = cellValue1
                            Else
                                wsV2.Cells(cellV2.row, lastColumnV2).Value = cellValue1
                            End If
                            wsV2.Cells(cellV2.row, lastColumnV3).Value = cellValue2
                            wsV2.Cells(cellV2.row, lastColumnV4).Value = cellValue3
                        End If
                Next cellV2
            End If
        End If
    Next cell
    'intitulé ligne
    wsV2.Cells(1, lastColumnV2).Value = "HS Chargées"
    wsV2.Cells(1, lastColumnV3).Value = "Salaire Brut"
    wsV2.Cells(1, lastColumnV4).Value = "Charges P"
End Sub

```