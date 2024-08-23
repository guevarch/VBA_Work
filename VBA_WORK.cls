VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CleanSales()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim temp As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Step 1: Keep only the specified columns
    keepCols = Array("ProductName", "Attribute1", "Sold")
    For i = ws.Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        If IsError(Application.Match(ws.Cells(1, i).Value, keepCols, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i

    ' Step 2: Switch columns B and C
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow ' Include the first row
        temp = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = ws.Cells(i, 3).Value
        ws.Cells(i, 3).Value = temp
    Next i

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

    ' Step 5: Replace any occurrences of "--None--" with an empty string
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For r = 1 To lastRow
        For c = 1 To lastCol
            If ws.Cells(r, c).Value = "--None--" Then
                ws.Cells(r, c).Value = ""
            End If
        Next c
    Next r

    MsgBox "Operations completed: Columns kept, switched, concatenated, renamed, and '--None--' replaced."
End Sub

Sub CleanPerfectMind()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim temp As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Long

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
        For c = 1 To lastCol
            If ws.Cells(r, c).Value = "--None--" Then
                ws.Cells(r, c).Value = ""
            End If
        Next c
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

Sub SummarizeSales()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim uniqueItems As Collection
    Dim item As Variant
    Dim i As Long
    Dim outputRow As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

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

    MsgBox "Summarization of items completed and placed into columns C and D."
End Sub


Sub PullColumnsAndMatchData()
    Dim wsSales As Worksheet
    Dim wsPM As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowSales As Long
    Dim lastRowPM As Long
    Dim lastRowOutput As Long
    Dim i As Long
    Dim matchRow As Variant

    ' Set the worksheets
    Set wsSales = Worksheets("Sales VBA")
    Set wsPM = Worksheets("PMQT")
    Set wsOutput = ActiveSheet

    ' Find the last row with data in Sales VBA sheet
    lastRowSales = wsSales.Cells(wsSales.Rows.Count, "C").End(xlUp).Row

    ' Copy columns C and D from Sales VBA to the active sheet
    wsSales.Range("C1:D" & lastRowSales).Copy
    wsOutput.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Clear the clipboard to remove the copied data
    Application.CutCopyMode = False

    ' Find the last row with data in PMQT sheet
    lastRowPM = wsPM.Cells(wsPM.Rows.Count, "A").End(xlUp).Row

    ' Find the last row with data in the active sheet
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row

    ' Loop through each item in the active sheet and match with PMQT sheet
    For i = 1 To lastRowOutput ' Start from the second row to skip headers
        matchRow = Application.Match(wsOutput.Cells(i, 1).Value, wsPM.Columns(1), 0)
        If Not IsError(matchRow) Then
            wsOutput.Cells(i, 3).Value = wsPM.Cells(matchRow, 2).Value
        End If
    Next i

    MsgBox "Columns C and D from Sales VBA have been copied to the active sheet with values only. Matching data from PMQT has been placed in column C."
End Sub

