# Challenge2
Challenge2 submission for the Data bootcamp


For this homework I based myself initially on the recording of the class we had on thursdey feb. 22. (which I couldn't attend).
In that lesson the part of coloring the cells was explained as well as how to target particular cells in a range, ordering the tickers, have the code work in may worksheets and so on.


From there I got together with a couple peers from class, as well as a quick revision from my brother who studied Computer technologies.

Another resource that helped me with this homework was ChatGPT. Here are some snippets of code it gave me (some did help, some others I could implement because of how I initially wrote my code)

1.
what dos rows.count do in visual basic

Dim totalRows As Long
totalRows = ActiveSheet.Rows.Count

2. how would you count the rows that contain the same value and stopping when it changes in visual basic
   Sub CountRowsWithSameValue()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change "Sheet1" to your sheet's name
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Assuming the values are in column A

    Dim currentValue As Variant
    Dim previousValue As Variant
    Dim rowCount As Long
    Dim i As Long

    rowCount = 0
    previousValue = ws.Cells(1, 1).Value ' Assuming the values start from row 1
    For i = 1 To lastRow
        currentValue = ws.Cells(i, 1).Value
        If currentValue = previousValue Then
            rowCount = rowCount + 1
        Else
            Exit For
        End If
    Next i

    MsgBox "Number of rows with the same value: " & rowCount
End Sub

3. what is worksheetfunction. in visual basic

   Dim total As Double
total = WorksheetFunction.Sum(Range("A1:A10"))



Sub FindFirstValueInColumn()
    Dim ws As Worksheet
    Dim firstValueCell As Range
    Dim columnNumber As Integer
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the first non-empty cell in the specified column
    Set firstValueCell = ws.Columns(columnNumber).Find("*", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not firstValueCell Is Nothing Then
        MsgBox "The first value in column " & columnNumber & " is: " & firstValueCell.Value
    Else
        MsgBox "Column " & columnNumber & " is empty."
    End If
End Sub



Sub FindLargestValueInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim columnNumber As Integer
    Dim maxVal As Double
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the last row with data in the specified column
    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    
    ' Find the largest value in the specified column
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(1, columnNumber), ws.Cells(lastRow, columnNumber)))
    
    MsgBox "The largest value in column " & columnNumber & " is: " & maxVal
End Sub



Sub FindFirstValueInColumn()
    Dim ws As Worksheet
    Dim firstValueCell As Range
    Dim columnNumber As Integer
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the first non-empty cell in the specified column
    Set firstValueCell = ws.Columns(columnNumber).Find("*", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not firstValueCell Is Nothing Then
        MsgBox "The first value in column " & columnNumber & " is: " & firstValueCell.Value
    Else
        MsgBox "Column " & columnNumber & " is empty."
    End If
End Sub



Sub FindFirstCellInColumn()
    Dim ws As Worksheet
    Dim firstCell As Range
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the first cell in column 1 (A)
    Set firstCell = ws.Cells(1, 1).End(xlDown)
    
    If Not firstCell Is Nothing Then
        MsgBox "The first cell in column 1 is: " & firstCell.Address
    Else
        MsgBox "Column 1 is empty."
    End If
End Sub



Sub FindFirstAndLastRowInColumnRange()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim columnNumber As Integer
    Dim rng As Range
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the first and last non-empty cells in the specified column
    Set rng = ws.Columns(columnNumber)
    On Error Resume Next
    firstRow = rng.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    lastRow = rng.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    
    If firstRow > 0 And lastRow > 0 Then
        MsgBox "The first row in column " & columnNumber & " is: " & firstRow & vbCrLf & _
               "The last row in column " & columnNumber & " is: " & lastRow
    Else
        MsgBox "Column " & columnNumber & " is empty."
    End If
End Sub




Dim firstRow As Long
firstRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Row


Sub FindLargestValueInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim columnNumber As Integer
    Dim maxVal As Double
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the last row with data in the specified column
    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    
    ' Find the largest value in the specified column
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(1, columnNumber), ws.Cells(lastRow, columnNumber)))
    
    MsgBox "The largest value in column " & columnNumber & " is: " & maxVal
End Sub


Sub FindLargestValueInColumnWithoutApplication()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim columnNumber As Integer
    Dim maxVal As Double
    Dim cell As Range
    
    ' Set the worksheet (change Sheet1 to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the column number (column A = 1, B = 2, etc.)
    columnNumber = 1 ' For column A
    
    ' Find the last row with data in the specified column
    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    
    ' Initialize maxVal to a very small value
    maxVal = -999999999999
    
    ' Iterate over each cell in the column to find the largest value
    For Each cell In ws.Range(ws.Cells(1, columnNumber), ws.Cells(lastRow, columnNumber))
        If IsNumeric(cell.Value) And cell.Value > maxVal Then
            maxVal = cell.Value
        End If
    Next cell
    
    If maxVal <> -999999999999 Then
        MsgBox "The largest value in column " & columnNumber & " is: " & maxVal
    Else
        MsgBox "Column " & columnNumber & " is empty."
    End If
End Sub


Sub CalculatePercentage()
    Dim part As Double
    Dim total As Double
    Dim percentage As Double

    ' Assuming part and total are numeric variables representing the values
    part = 25
    total = 50

    ' Calculate the percentage
    percentage = (part / total) * 100

    MsgBox "The percentage is: " & percentage & "%"
End Sub




Sub CalculateAndDisplayPercentage()
    Dim part As Double
    Dim total As Double
    Dim percentage As Double
    Dim resultCell As Range

    ' Assuming part and total are numeric variables representing the values
    part = 25
    total = 50

    ' Calculate the percentage
    percentage = (part / total) * 100

    ' Set the range where you want to display the result (change to your desired cell)
    Set resultCell = ThisWorkbook.Sheets("Sheet1").Range("A1")

    ' Display the percentage in the cell with the "%" symbol
    resultCell.NumberFormat = "0.00%" ' Sets the cell format to display two decimal places and the "%" symbol
    resultCell.Value = percentage / 100 ' Divides by 100 to convert the percentage to a decimal for display
End Sub




Dim ws As Worksheet
Dim lastRow As Long
Dim columnNumber As Integer
Dim maxVal As Double

' Set the worksheet (change Sheet1 to your sheet's name)
Set ws = ThisWorkbook.Sheets("Sheet1")

' Set the column number (column A = 1, B = 2, etc.)
columnNumber = 1 ' For column A

' Find the last row with data in the specified column
lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

' Find the largest value in the specified column using WorksheetFunction.Max
maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(1, columnNumber), ws.Cells(lastRow, columnNumber)))

MsgBox "The largest value in column " & columnNumber & " is: " & maxVal



Dim ws As Worksheet
Dim lastRow As Long
Dim columnNumber As Integer
Dim maxVal As Double
Dim cell As Range

' Set the worksheet (change Sheet1 to your sheet's name)
Set ws = ThisWorkbook.Sheets("Sheet1")

' Set the column number (column A = 1, B = 2, etc.)
columnNumber = 1 ' For column A

' Find the last row with data in the specified column
lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

' Initialize maxVal to a very small value
maxVal = -999999999999

' Iterate over each cell in the column to find the largest value
For Each cell In ws.Range(ws.Cells(1, columnNumber), ws.Cells(lastRow, columnNumber))
    If IsNumeric(cell.Value) And cell.Value > maxVal Then
        maxVal = cell.Value
    End If
Next cell

If maxVal <> -999999999999 Then
    MsgBox "The largest value in column " & columnNumber & " is: " & maxVal
Else
    MsgBox "Column " & columnNumber & " is empty."
End If



Dim ws As Worksheet
Dim lastRow As Long
Dim columnNumber As Integer
Dim maxVal As Double

' Set the worksheet (change Sheet1 to your sheet's name)
Set ws = ThisWorkbook.Sheets("Sheet1")

' Set the column number (column A = 1, B = 2, etc.)
columnNumber = 1 ' For column A

' Find the last row with data in the specified column
lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

' Find the largest value in the specified column using the Evaluate method
maxVal = ws.Evaluate("MAX(" & ws.Cells(1, columnNumber).Address & ":" & ws.Cells(lastRow, columnNumber).Address & ")")

MsgBox "The largest value in column " & columnNumber & " is: " & maxVal








