``` VBA
Sub DeleteAndSplitIdMatiere()
    Dim row As Long
    Dim cell As Range
    Dim idMatiereList As Object
    Dim element As Variant
    Dim ws As Worksheet
    Dim yWs As Worksheet
    Dim lastRow As Long
    Dim xWs As Worksheet
    Dim wsSource As Worksheet
    Dim wsCible As Worksheet
    Dim graph As ChartObject

    firstSheet = Sheets(1).Name

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In ActiveWorkbook.Sheets
        If xWs.Name <> firstSheet Then
            xWs.Delete
        End If
    Next

    Set wsSource = ActiveWorkbook.Sheets(firstSheet)

    Set idMatiereList = CreateObject("Scripting.Dictionary")

    Set ws = ActiveWorkbook.Sheets(firstSheet)

    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).row

    For Each cell In ws.Range("E2:E" & lastRow)

        element = cell.Value

        If idMatiereList.exists(element) Then
            Set wsCible = ActiveWorkbook.Sheets("id " & cell)
            rowAdress = Split(cell.Address, "$")(2)

            derniereLigne = Cells(Rows.Count, 1).End(xlUp).row + 1

            wsCible.Rows(derniereLigne).Value = wsSource.Rows(rowAdress).Value

        Else
            idMatiereList.Add element, Nothing

            Sheets.Add(After:=Sheets(firstSheet)).Name = "id " & cell
            Sheets(firstSheet).Range("A1:AJ1").Copy Sheets("id " & cell).Range("A1")

            rowAdress = Split(cell.Address, "$")(2)

            derniereLigne = Cells(Rows.Count, 1).End(xlUp).row + 1
            Set wsCible = ActiveWorkbook.Sheets("id " & cell)
            wsCible.Rows(derniereLigne).Value = wsSource.Rows(rowAdress).Value

        End If
    Next cell

    For Each yWs In ActiveWorkbook.Worksheets
            lastRow = yWs.Cells(yWs.Rows.Count, "E").End(xlUp).row
            yWs.Range("AL1").Formula = yWs.Range("AC2").Value
            yWs.Range("AM1").Formula = yWs.Range("AC2").Value
            yWs.Range("AL2").Formula = "=IF(S2=0,-N2*O2*Z2,N2*O2*Z2)"
            yWs.Range("AM2").Formula = "=AM1+AL2"
            yWs.Range("AM2:AM" & lastRow).FillDown
            yWs.Range("AL2:AL" & lastRow).FillDown

        With yWs.Columns("AM:AM").FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .SetFirstPriority
            With .Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        End With

        Set graph = yWs.ChartObjects.Add(Left:=180, Width:=500, Top:=50, Height:=500)
        graph.Chart.SetSourceData Source:=yWs.Range("R:R, AM:AM")
        graph.Chart.ChartType = xlLineMarkers

        graph.Chart.HasTitle = True
        graph.Chart.ChartTitle.Text = "Évolution stock matière"

        yWs.Activate
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 1

        ActiveWindow.FreezePanes = True

    Next yWs

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub DeleteWorksheet()
    Dim xWs As Worksheet

    firstSheet = Sheets(1).Name

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> firstSheet Then
            xWs.Delete
        End If
    Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
```
