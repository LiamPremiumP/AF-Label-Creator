Sub Layout1()

    Dim RowNo As Double     'Number of rows containing data, counted from column A (Specified column if printing specific labels)
    Dim IE As Variant   'The chosen layout
    Dim check As Variant
    Dim check2 As Variant
    Dim check3 As String
    Dim j As Integer
    j = 1
    
    check = 0
    check2 = 0
    check3 = Sheets("Existing").Cells(5, 1).Value
    


    RowNoE = (Sheets("Existing").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) + 4
    RowNoEP = (Sheets("Proposed").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count)

    For i = 5 To RowNoE
        Sheets("E+P").Rows(i - 4).Value = Sheets("Existing").Rows(i).Value
        For k = 1 To RowNoEP
            If Sheets("Proposed").Cells(k, 1).Value = Sheets("Existing").Cells(i, 1).Value And Sheets("Proposed").Cells(k, 2).Value = Sheets("Existing").Cells(i, 2).Value Then
                Sheets("E+P").Rows(i - 4).Value = Sheets("Proposed").Rows(k).Value
                Sheets("E+P").Rows(i - 4).Interior.ColorIndex = 4
            End If
        Next k
    Next i

End Sub


