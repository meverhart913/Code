I'llOption Explicit

Sub BuildOptimizedPuckLayout()

    Dim ws As Worksheet
    Dim startCell As Range
    Dim targetRange As Range
    Dim clearRange As Range
    Dim shp As Shape
    Dim cell As Range
    Dim i As Long

    Dim partW As Double
    Dim partH As Double
    Dim puckD As Double

    Dim m As Long
    Dim n As Long

    Dim scalePts As Double
    Dim rowHt As Double
    Dim colW As Double

    Dim rx As Double
    Dim ry As Double

    Dim bestCX As Double
    Dim bestCY As Double
    Dim testCX As Double
    Dim testCY As Double

    Dim bestCount As Long
    Dim testCount As Long
    Dim stepSize As Double
    Dim counter As Long

    Set ws = ActiveSheet
    Set startCell = ws.Range("K3")

    partW = ws.Range("H4").Value
    partH = ws.Range("H5").Value
    puckD = ws.Range("H6").Value

    If partW <= 0 Or partH <= 0 Or puckD <= 0 Then
        MsgBox "H4, H5, and H6 must all be greater than zero.", vbExclamation
        Exit Sub
    End If

    'Scale factor: Excel points per mm
    scalePts = 20

    'Calculate grid size large enough to cover puck
    m = WorksheetFunction.Ceiling(puckD / partW, 1) + 2
    n = WorksheetFunction.Ceiling(puckD / partH, 1) + 2

    '----------------------------------------
    ' CLEAR OLD OUTPUT AREA
    '----------------------------------------
    Set clearRange = ws.Range("K3", ws.Cells(ws.Rows.Count, ws.Columns.Count))

    clearRange.ClearContents
    clearRange.ClearFormats
    clearRange.Borders.LineStyle = xlNone

    'Reset old column widths / row heights in output area
    ws.Range("K:XFD").ColumnWidth = 8.43
    ws.Range("3:" & ws.Rows.Count).RowHeight = 15

    'Delete old oval/output shapes
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name Like "RangeCircle*" Then
            ws.Shapes(i).Delete
        End If
    Next i

    '----------------------------------------
    ' CREATE GRID AREA
    '----------------------------------------
    Set targetRange = startCell.Resize(n, m)

    rowHt = partH * scalePts
    targetRange.RowHeight = rowHt

    colW = targetRange.Columns(1).ColumnWidth * _
           ((partW * scalePts) / targetRange.Columns(1).Width)

    targetRange.Columns.ColumnWidth = colW

    Set targetRange = startCell.Resize(n, m)

    'Fixed oval size from puck diameter
    rx = (puckD * scalePts) / 2
    ry = (puckD * scalePts) / 2

    '----------------------------------------
    ' OPTIMIZE OVAL POSITION
    'Constrained so oval never moves left/up over input area
    '----------------------------------------
    stepSize = 1
    bestCount = 0

    For testCX = targetRange.Left + rx To targetRange.Left + targetRange.Width - rx Step stepSize
        For testCY = targetRange.Top + ry To targetRange.Top + targetRange.Height - ry Step stepSize

            testCount = 0

            For Each cell In targetRange.Cells
                If CellFullyInsideOval(cell, testCX, testCY, rx, ry) Then
                    testCount = testCount + 1
                End If
            Next cell

            If testCount > bestCount Then
                bestCount = testCount
                bestCX = testCX
                bestCY = testCY
            End If

        Next testCY
    Next testCX

    'Fallback if constrained loop cannot run
    If bestCX = 0 Or bestCY = 0 Then
        bestCX = targetRange.Left + rx
        bestCY = targetRange.Top + ry
    End If

    '----------------------------------------
    ' DRAW OVAL
    '----------------------------------------
    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeOval, _
        Left:=bestCX - rx, _
        Top:=bestCY - ry, _
        Width:=puckD * scalePts, _
        Height:=puckD * scalePts)

    With shp
        .Name = "RangeCircle"
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 3
        .Placement = xlFreeFloating
    End With

    '----------------------------------------
    ' APPLY RED BORDERS AND NUMBERS
    '----------------------------------------
    counter = 1

    For Each cell In targetRange.Cells

        cell.ClearContents
        cell.Borders.LineStyle = xlNone

        If CellFullyInsideOval(cell, bestCX, bestCY, rx, ry) Then

            With cell.Borders
                .LineStyle = xlContinuous
                .Color = RGB(255, 0, 0)
                .Weight = xlThick
            End With

            cell.Value = counter
            cell.HorizontalAlignment = xlCenter
            cell.VerticalAlignment = xlCenter
            cell.Font.Size = 10
            cell.Font.Bold = True

            counter = counter + 1

        End If

    Next cell

    ws.Range("H10").Value = n
    ws.Range("H11").Value = m
    ws.Range("H12").Value = bestCount

End Sub


Function CellFullyInsideOval(cell As Range, _
                             cx As Double, cy As Double, _
                             rx As Double, ry As Double) As Boolean

    CellFullyInsideOval = _
        PointInsideOval(cell.Left, cell.Top, cx, cy, rx, ry) And _
        PointInsideOval(cell.Left + cell.Width, cell.Top, cx, cy, rx, ry) And _
        PointInsideOval(cell.Left, cell.Top + cell.Height, cx, cy, rx, ry) And _
        PointInsideOval(cell.Left + cell.Width, cell.Top + cell.Height, cx, cy, rx, ry)

End Function


Function PointInsideOval(x As Double, y As Double, _
                         cx As Double, cy As Double, _
                         rx As Double, ry As Double) As Boolean

    Dim v As Double

    v = (((x - cx) * (x - cx)) / (rx * rx)) + _
        (((y - cy) * (y - cy)) / (ry * ry))

    PointInsideOval = v <= 1

End Function



SELECT measurements_log.item_id, measurements_log.param_no, measurements_log.meas_date, measurements_log.meas_value, parameter_master.param_desc
FROM products.dbo.measurements_log measurements_log, products.dbo.parameter_master parameter_master
WHERE parameter_master.param_no = measurements_log.param_no AND ((measurements_log.item_id Like '%+A26%') AND (measurements_log.last_tag_item_param_flag=1) AND (measurements_log.param_no In (905,906,1236,1237)))
ORDER BY measurements_log.item_id


SELECT
    ml.item_id,

    CASE
        WHEN LEFT(ml.item_id, 2) = '1+' THEN 'Upper'
        WHEN LEFT(ml.item_id, 2) = '2+' THEN 'Lower'
        ELSE 'Unknown'
    END AS slice_position,

    SUBSTRING(ml.item_id, CHARINDEX('+', ml.item_id) + 1, LEN(ml.item_id)) AS boule_id,

    ml.param_no,

    pm.param_desc,

    CASE
        WHEN ml.param_no IN (1236, 1237) THEN 'Pre Anneal'
        WHEN ml.param_no IN (905, 906) THEN 'Post Anneal'
        ELSE 'Unknown'
    END AS anneal_stage,

    CASE
        WHEN ml.param_no IN (905, 1236) THEN 'Inner'
        WHEN ml.param_no IN (906, 1237) THEN 'Outer'
        ELSE 'Unknown'
    END AS ring_position,

    ml.meas_date,
    ml.meas_value

FROM products.dbo.measurements_log ml
JOIN products.dbo.parameter_master pm
    ON pm.param_no = ml.param_no

WHERE
    ml.item_id LIKE '%+A26%'
    AND ml.last_tag_item_param_flag = 1
    AND ml.param_no IN (905, 906, 1236, 1237)

ORDER BY
    boule_id,
    slice_position,
    anneal_stage,
    ring_position;


WITH ActiveInventory AS (
    SELECT
        bt.boule_id,
        bt.tier_id,
        bt.tier_length_mm,
        bt.tier_start_mm
    FROM products.dbo.boule_tiers bt
    WHERE
        bt.material_no = 54
        AND bt.tier_status = 1
),

LatestMeasurements AS (
    SELECT
        ml.item_id,

        CASE
            WHEN LEFT(ml.item_id, 2) = '1+' THEN 'Upper'
            WHEN LEFT(ml.item_id, 2) = '2+' THEN 'Lower'
            ELSE 'Unknown'
        END AS slice_position,

        SUBSTRING(ml.item_id, CHARINDEX('+', ml.item_id) + 1, LEN(ml.item_id)) AS boule_id,

        ml.param_no,
        pm.param_desc,

        CASE
            WHEN ml.param_no IN (1236, 1237) THEN 'Pre Anneal'
            WHEN ml.param_no IN (905, 906) THEN 'Post Anneal'
            ELSE 'Unknown'
        END AS anneal_stage,

        CASE
            WHEN ml.param_no IN (905, 1236) THEN 'Inner'
            WHEN ml.param_no IN (906, 1237) THEN 'Outer'
            ELSE 'Unknown'
        END AS ring_position,

        ml.meas_date,
        ml.meas_value,

        ROW_NUMBER() OVER (
            PARTITION BY ml.item_id, ml.param_no
            ORDER BY ml.meas_date DESC
        ) AS rn

    FROM products.dbo.measurements_log ml
    JOIN products.dbo.parameter_master pm
        ON pm.param_no = ml.param_no
    WHERE
        ml.param_no IN (905, 906, 1236, 1237)
)

SELECT
    ai.boule_id,
    ai.tier_id,
    ai.tier_length_mm,
    ai.tier_start_mm,

    lm.item_id,
    lm.slice_position,
    lm.param_no,
    lm.param_desc,
    lm.anneal_stage,
    lm.ring_position,
    lm.meas_date,
    lm.meas_value

FROM ActiveInventory ai
LEFT JOIN LatestMeasurements lm
    ON lm.boule_id = ai.boule_id
    AND lm.rn = 1

ORDER BY
    ai.boule_id,
    lm.slice_position,
    lm.anneal_stage,
    lm.ring_position;


Sub FindInventory()

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing SQL data..."

    'Refresh SQL / query tables
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    Application.StatusBar = "Recalculating helper columns..."

    'Force Data table helper formulas to recalculate
    Worksheets("Data").ListObjects("DataTable1").Range.Calculate

    'Force Inventory formulas to recalculate
    Worksheets("Inventory").Calculate

    'Optional stronger full workbook calc if needed
    Application.CalculateFull

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Inventory updated.", vbInformation
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Inventory update failed: " & Err.Description, vbExclamation

End Sub

Option Explicit

Sub DrawBouleGradient()

    Const WS_INV As String = "Inventory"
    Const WS_DATA As String = "Data"
    Const TBL_DATA As String = "DataTable1"

    Const START_CELL As String = "K4"
    Const SEGMENT_MM As Double = 10

    Dim wsI As Worksheet, wsD As Worksheet
    Dim lo As ListObject
    Dim bouleID As String
    Dim lowerLimit As Double, upperLimit As Double

    Dim tierLength As Double, tierStart As Double
    Dim fullLength As Double
    Dim segmentCount As Long
    Dim r As Long

    Dim preUpperAvg As Double, preLowerAvg As Double
    Dim postUpperAvg As Double, postLowerAvg As Double

    Dim posFromTop As Double
    Dim preAlpha As Double, postAlpha As Double

    Dim startRange As Range, drawRange As Range

    Set wsI = ThisWorkbook.Worksheets(WS_INV)
    Set wsD = ThisWorkbook.Worksheets(WS_DATA)
    Set lo = wsD.ListObjects(TBL_DATA)

    bouleID = Trim(wsI.Range("J2").Value)

    If bouleID = "" Then
        MsgBox "Select a boule ID in J2 first.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsI.Range("B2").Value) Or Not IsNumeric(wsI.Range("C2").Value) Then
        MsgBox "Lower and upper limits must be numeric.", vbExclamation
        Exit Sub
    End If

    lowerLimit = CDbl(wsI.Range("B2").Value)
    upperLimit = CDbl(wsI.Range("C2").Value)

    If lowerLimit > upperLimit Then
        MsgBox "Lower limit cannot be greater than upper limit.", vbExclamation
        Exit Sub
    End If

    tierLength = GetFirstNumeric(lo, bouleID, "tier_length_mm")
    tierStart = GetFirstNumeric(lo, bouleID, "tier_start_mm")

    If tierLength <= 0 Then
        MsgBox "No valid tier_length_mm found for " & bouleID & ".", vbExclamation
        Exit Sub
    End If

    fullLength = tierStart + tierLength

    preUpperAvg = GetMeasurementAvg(lo, bouleID, "Pre Anneal", "Upper")
    preLowerAvg = GetMeasurementAvg(lo, bouleID, "Pre Anneal", "Lower")

    postUpperAvg = GetMeasurementAvg(lo, bouleID, "Post Anneal", "Upper")
    postLowerAvg = GetMeasurementAvg(lo, bouleID, "Post Anneal", "Lower")

    If preUpperAvg = -1 Or preLowerAvg = -1 Or postUpperAvg = -1 Or postLowerAvg = -1 Then
        MsgBox "Missing required upper/lower pre or post anneal measurements for " & bouleID & ".", vbExclamation
        Exit Sub
    End If

    Set startRange = wsI.Range(START_CELL)

    wsI.Range("K4:L200").Clear
    wsI.Range("K4:L200").Interior.Pattern = xlNone
    wsI.Range("K4:L200").Borders.LineStyle = xlNone

    segmentCount = Application.WorksheetFunction.RoundUp(tierLength / SEGMENT_MM, 0)

    For r = 1 To segmentCount

        posFromTop = tierStart + ((r - 0.5) * SEGMENT_MM)

        If posFromTop > fullLength Then
            posFromTop = fullLength - ((fullLength - tierStart) / 2)
        End If

        preAlpha = LinearAlpha(posFromTop, fullLength, preUpperAvg, preLowerAvg)
        postAlpha = LinearAlpha(posFromTop, fullLength, postUpperAvg, postLowerAvg)

        startRange.Offset(r - 1, 0).Value = Round(preAlpha, 3)
        startRange.Offset(r - 1, 1).Value = Round(postAlpha, 3)

        If preAlpha >= lowerLimit And preAlpha <= upperLimit Then
            startRange.Offset(r - 1, 0).Interior.Color = RGB(198, 239, 206)
        End If

        If postAlpha >= lowerLimit And postAlpha <= upperLimit Then
            startRange.Offset(r - 1, 1).Interior.Color = RGB(198, 239, 206)
        End If

    Next r

    Set drawRange = wsI.Range(startRange, startRange.Offset(segmentCount - 1, 1))

    With drawRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    With drawRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    With drawRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    With drawRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    drawRange.HorizontalAlignment = xlCenter
    drawRange.VerticalAlignment = xlCenter
    drawRange.NumberFormat = "0.000"

End Sub

Private Function LinearAlpha(ByVal posFromTop As Double, ByVal fullLength As Double, ByVal upperAvg As Double, ByVal lowerAvg As Double) As Double
    If fullLength = 0 Then
        LinearAlpha = upperAvg
    Else
        LinearAlpha = upperAvg + (posFromTop / fullLength) * (lowerAvg - upperAvg)
    End If
End Function

Private Function GetMeasurementAvg(ByVal lo As ListObject, ByVal bouleID As String, ByVal annealStage As String, ByVal slicePosition As String) As Double

    Dim i As Long
    Dim total As Double
    Dim countVals As Long

    Dim colBoule As Long, colStage As Long, colSlice As Long, colRing As Long, colValue As Long

    colBoule = lo.ListColumns("boule_id").Index
    colStage = lo.ListColumns("anneal_stage").Index
    colSlice = lo.ListColumns("slice_position").Index
    colRing = lo.ListColumns("ring_position").Index
    colValue = lo.ListColumns("meas_value").Index

    For i = 1 To lo.DataBodyRange.Rows.Count

        If Trim(lo.DataBodyRange.Cells(i, colBoule).Value) = bouleID _
           And Trim(lo.DataBodyRange.Cells(i, colStage).Value) = annealStage _
           And Trim(lo.DataBodyRange.Cells(i, colSlice).Value) = slicePosition _
           And Trim(lo.DataBodyRange.Cells(i, colRing).Value) <> "Unknown" _
           And IsNumeric(lo.DataBodyRange.Cells(i, colValue).Value) Then

            total = total + CDbl(lo.DataBodyRange.Cells(i, colValue).Value)
            countVals = countVals + 1

        End If

    Next i

    If countVals = 0 Then
        GetMeasurementAvg = -1
    Else
        GetMeasurementAvg = total / countVals
    End If

End Function

Private Function GetFirstNumeric(ByVal lo As ListObject, ByVal bouleID As String, ByVal columnName As String) As Double

    Dim i As Long
    Dim colBoule As Long, colTarget As Long

    colBoule = lo.ListColumns("boule_id").Index
    colTarget = lo.ListColumns(columnName).Index

    For i = 1 To lo.DataBodyRange.Rows.Count
        If Trim(lo.DataBodyRange.Cells(i, colBoule).Value) = bouleID Then
            If IsNumeric(lo.DataBodyRange.Cells(i, colTarget).Value) Then
                GetFirstNumeric = CDbl(lo.DataBodyRange.Cells(i, colTarget).Value)
                Exit Function
            End If
        End If
    Next i

    GetFirstNumeric = 0

End Function