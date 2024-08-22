# VBA_Work

Sub sales()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim i As Integer

    ' Specify the columns to keep
    keepCols = Array("ProductName", "Attribute1", "Sold")
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Loop through columns from the end to the beginning
    For i = ws.Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        ' If the column header is not in the keepCols array, delete the column
        If IsError(Application.Match(ws.Cells(1, i).Value, keepCols, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i
End Sub
 
Sub PM()
    Dim ws As Worksheet
    Dim keepCols As Variant
    Dim i As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long

    ' Specify the columns to keep
    keepCols = Array("Name", "Attribute 1", "Available Quantity")
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Loop through columns from the end to the beginning
    For i = ws.Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        ' If the column header is not in the keepCols array, delete the column
        If IsError(Application.Match(ws.Cells(1, i).Value, keepCols, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i

    ' Find the last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each cell and remove "--None--"
    For r = 1 To lastRow
        For c = 1 To lastCol
            If ws.Cells(r, c).Value = "--None--" Then
                ws.Cells(r, c).Value = ""
            End If
        Next c
    Next r

    MsgBox "'--None--' has been removed from the worksheet."
End Sub

Sub SwitchColumns1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim temp As Variant
    Dim i As Long

    ' Set the worksheet
    Set ws = ActiveSheet

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row and switch columns B and C
    For i = 1 To lastRow ' Start from the first row
        temp = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = ws.Cells(i, 3).Value
        ws.Cells(i, 3).Value = temp
    Next i

    MsgBox "Columns B and C have been switched."
End Sub

Sub ConcatenateAndDeleteColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet
    Set ws = ActiveSheet

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row and concatenate columns A and B into column A
    For i = 1 To lastRow ' Include the first row
        ws.Cells(i, 1).Value = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
    Next i

    ' Move the concatenated values to a new column (e.g., column D)
    ws.Range("D1").Value = "Concatenated" ' Header for the new column
    ws.Range("D1:D" & lastRow).Value = ws.Range("A1:A" & lastRow).Value

    ' Delete the original columns A and B
    ws.Columns("A:B").Delete

    MsgBox "Columns A and B have been concatenated and moved to column D. Original columns A and B have been deleted."
End Sub

Sub SwitchColumns2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim temp As Variant
    Dim i As Long

    ' Set the worksheet
    Set ws = ActiveSheet

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Rename column B to "Item"
    ws.Cells(1, 2).Value = "Item"

    ' Loop through each row and switch columns A and B
    For i = 1 To lastRow ' Include the first row
        temp = ws.Cells(i, 1).Value
        ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
        ws.Cells(i, 2).Value = temp
    Next i

    MsgBox "Column B has been renamed to 'Item' and columns A and B have been switched."
End Sub

Sub CombinedOperationsSales()
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
        ws.Cells(i, 1).Value = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
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

Sub CombinedOperationsPM()
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
        ws.Cells(i, 1).Value = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
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
