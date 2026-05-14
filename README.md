Option Explicit

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



WITH RankedMeasurements AS (
    SELECT 
        ap.order_id,
        ap.serial_id,
        ml.param_no,
        ml.meas_value,
        ml.meas_date,
        ROW_NUMBER() OVER (
            PARTITION BY ap.serial_id, ml.param_no
            ORDER BY ml.meas_date DESC
        ) AS rn
    FROM products.dbo.active_parts AS ap
    INNER JOIN products.dbo.measurements_log AS ml
        ON ap.tag_id = ml.tag_id
    WHERE ap.order_id = ?
      AND ml.meas_value_type = 'F'
      AND ml.param_no IN (119, 120, 432, 534)
)
SELECT


WITH RankedMeasurements AS ( SELECT ap.order_id, ap.serial_id, ml.param_no, ml.meas_value, ml.meas_date, ROW_NUMBER() OVER ( PARTITION BY ap.serial_id, ml.param_no ORDER BY ml.meas_date DESC ) AS rn FROM products.dbo.active_parts AS ap INNER JOIN products.dbo.measurements_log AS ml ON ap.tag_id = ml.tag_id WHERE ap.order_id = 198766  AND ml.meas_value_type = 'I' AND ml.param_no IN (119, 120, 432, 534) ) SELECT order_id, serial_id, MAX(CASE WHEN param_no = 119 THEN meas_value END) AS overall_length, MAX(CASE WHEN param_no = 120 THEN meas_value END) AS diameter, MAX(CASE WHEN param_no = 432 THEN meas_value END) AS cr4_length, MAX(CASE WHEN param_no = 534 THEN meas_value END) AS transmission_percent FROM RankedMeasurements WHERE rn = 1 GROUP BY order_id, serial_id ORDER BY serial_id;
