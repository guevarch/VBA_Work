VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CleanPerfectMind()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim temp As Variant
    Dim i As Long
    Dim r As Long
    Dim C As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Step 1: Keep only the specified columns
    keepCols = Array("Name", "Attribute 1", "Available Quantity")
    For i = ws.Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        If IsError(Application.Match(ws.Cells(1, i).Value, keepCols, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i

    ' Find the last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Step 2: Remove "--None--"
    For r = 1 To lastRow
        For C = 1 To lastCol
            If ws.Cells(r, C).Value = "--None--" Then
                ws.Cells(r, C).Value = ""
            End If
        Next C
    Next r

    ' Step 3: Concatenate columns A and B into column A and delete original columns A and B
    For i = 1 To lastRow ' Include the first row
        ws.Cells(i, 1).Value = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
    Next i
    ws.Range("D1").Value = "Concatenated" ' Header for the new column
    ws.Range("D1:D" & lastRow).Value = ws.Range("A1:A" & lastRow).Value
    ws.Columns("A:B").Delete

    ' Step 4: Rename column B to "Item" and switch columns A and B
    ws.Cells(1, 2).Value = "Item"
    For i = 1 To lastRow ' Include the first row
        temp = ws.Cells(i, 1).Value
        ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = temp
    Next i

    MsgBox "Operations completed: Columns kept, '--None--' removed, columns concatenated, renamed, and switched."
End Sub

Sub CleanAndSummarizeSales()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim temp As Variant
    Dim i As Long
    Dim r As Long
    Dim C As Long
    Dim uniqueItems As Collection
    Dim item As Variant
    Dim outputRow As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Step 1: Clean Sales Data
    ' Keep only the specified columns
    keepCols = Array("ProductName", "Attribute1", "Sold")
    For i = ws.Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        If IsError(Application.Match(ws.Cells(1, i).Value, keepCols, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i

    ' Switch columns B and C
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow ' Include the first row
        temp = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = ws.Cells(i, 3).Value
        ws.Cells(i, 3).Value = temp
    Next i

    ' Concatenate columns A and B into column A and delete original columns A and B
    For i = 1 To lastRow ' Include the first row
        ws.Cells(i, 1).Value = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
    Next i
    ws.Range("D1").Value = "Concatenated" ' Header for the new column
    ws.Range("D1:D" & lastRow).Value = ws.Range("A1:A" & lastRow).Value
    ws.Columns("A:B").Delete

    ' Rename column B to "Item" and switch columns A and B
    ws.Cells(1, 2).Value = "Item"
    For i = 1 To lastRow ' Include the first row
        temp = ws.Cells(i, 1).Value
        ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = temp
    Next i

    ' Replace any occurrences of "--None--" with an empty string
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For r = 1 To lastRow
        For C = 1 To lastCol
            If ws.Cells(r, C).Value = "--None--" Then
                ws.Cells(r, C).Value = ""
            End If
        Next C
    Next r

    ' Step 2: Summarize Sales Data
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create a collection to store unique items
    Set uniqueItems = New Collection

    ' Loop through each row to collect unique items
    On Error Resume Next ' Ignore errors when adding duplicate items
    For i = 2 To lastRow ' Start from the second row to skip headers
        uniqueItems.Add ws.Cells(i, 1).Value, CStr(ws.Cells(i, 1).Value)
    Next i
    On Error GoTo 0 ' Turn error handling back on

    ' Output the unique items to column C
    outputRow = 2 ' Start from the second row to skip headers
    ws.Cells(1, 3).Value = "Item"
    ws.Cells(1, 4).Value = "Total Sold"

    For Each item In uniqueItems
        ws.Cells(outputRow, 3).Value = item
        ws.Cells(outputRow, 4).Formula = "=SUMIF(A:A, C" & outputRow & ", B:B)"
        outputRow = outputRow + 1
    Next item

    MsgBox "Operations completed: Data cleaned, summarized, and results placed into columns C and D."
End Sub


Sub Controlreport()
    Dim wsControl As Worksheet
    Dim wsSales As Worksheet
    Dim wsPMQT As Worksheet
    Dim lastRowControl As Long
    Dim lastRowSales As Long
    Dim lastRowPMQT As Long
    Dim i As Long
    Dim matchRow As Variant
    Dim physicalCount As Variant
    Dim perfectMindQty As Variant
    Dim salesQty As Variant
    Dim result As Double

    ' Set the worksheets
    Set wsControl = ThisWorkbook.Sheets("Control Report")
    Set wsSales = ThisWorkbook.Sheets("Sales")
    Set wsPMQT = ThisWorkbook.Sheets("PMQT")

    ' Find the last rows in all sheets
    lastRowControl = wsControl.Cells(wsControl.Rows.Count, "A").End(xlUp).Row
    lastRowSales = wsSales.Cells(wsSales.Rows.Count, "C").End(xlUp).Row
    lastRowPMQT = wsPMQT.Cells(wsPMQT.Rows.Count, "A").End(xlUp).Row

    ' Loop through the Control Report column A
    For i = 2 To lastRowControl ' Assuming row 1 is a header
        ' Match in Sales tab column C and pull data from column D to Control Report column E
        matchRow = Application.Match(wsControl.Cells(i, 1).Value, wsSales.Range("C2:C" & lastRowSales), 0)
        If Not IsError(matchRow) Then
            wsControl.Cells(i, 5).Value = wsSales.Cells(matchRow + 1, 4).Value ' matchRow + 1 adjusts for header
        Else
            wsControl.Cells(i, 5).Value = "No Match"
        End If

        ' Match in PMQT tab column A and pull data from column B to Control Report column D
        matchRow = Application.Match(wsControl.Cells(i, 1).Value, wsPMQT.Range("A2:A" & lastRowPMQT), 0)
        If Not IsError(matchRow) Then
            wsControl.Cells(i, 4).Value = wsPMQT.Cells(matchRow + 1, 2).Value ' matchRow + 1 adjusts for header
        Else
            wsControl.Cells(i, 4).Value = "No Match"
        End If

        ' Calculate the sum of Physical Count and PerfectMind Quantity, then subtract Sales
        physicalCount = wsControl.Cells(i, 3).Value
        perfectMindQty = wsControl.Cells(i, 4).Value
        salesQty = wsControl.Cells(i, 5).Value

        ' Ensure the values are numeric, otherwise set them to 0
        If Not IsNumeric(physicalCount) Then physicalCount = 0
        If Not IsNumeric(perfectMindQty) Then perfectMindQty = 0
        If Not IsNumeric(salesQty) Then salesQty = 0

        ' Perform the calculation
        result = physicalCount + perfectMindQty - salesQty

        ' Place the result in Column F (or next available column)
        wsControl.Cells(i, 6).Value = result
    Next i

    MsgBox "Data transfer and inventory difference calculation complete.", vbInformation

End Sub






