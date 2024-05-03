Sub Layout2()

    Dim RowNo As Double     'Number of rows containing data, counted from column A (Specified column if printing specific labels)
    Dim IE As Variant   'The chosen layout
    Dim check As Variant
    Dim check2 As Variant
    Dim check3 As String
    Dim j As Integer
    j = 1
    Dim k As Integer
    k = 1
    
    check = Sheets("E+P").Cells(1, 1).Value
    check2 = 0
    check3 = ""
    
    
    
    RowNo = (Sheets("E+P").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) + 3
    
    For i = 1 To RowNo

    
        If check = Sheets("E+P").Cells(i, 1).Value Then
           check = Sheets("E+P").Cells(i, 1).Value
            If check2 <= Sheets("E+P").Cells(i, 19).Value Then
                check2 = Sheets("E+P").Cells(i, 19).Value
                If check3 <= Sheets("E+P").Cells(i, 17).Value Then
                check3 = Sheets("E+P").Cells(i, 17).Value
                IE = i
                End If
            End If
            k = k + 1
        ElseIf Sheets("E+P").Cells(i - 1, 1).Value <> Sheets("E+P").Cells(i, 1).Value And k > 0 Then
            Sheets("WorstCase E+P").Rows(j).Value = Sheets("E+P").Rows(IE).Value
                Sheets("WorstCase E+P").Rows(j).Interior.ColorIndex = Sheets("E+P").Rows(IE).Interior.ColorIndex
                k = 0
                j = j + 1
                check = Sheets("E+P").Cells(i, 1).Value
                check2 = Sheets("E+P").Cells(i, 19).Value
                check3 = Sheets("E+P").Cells(i, 17).Value
        ElseIf k = 0 Then
                Sheets("WorstCase E+P").Rows(j).Value = Sheets("E+P").Rows(i).Value
                Sheets("WorstCase E+P").Rows(j).Interior.ColorIndex = Sheets("E+P").Rows(i).Interior.ColorIndex
                check = Sheets("E+P").Cells(i, 1).Value
                check2 = Sheets("E+P").Cells(i, 19).Value
                check3 = Sheets("E+P").Cells(i, 17).Value
            j = j + 1
        End If
    Next i
    
    End Sub
    
