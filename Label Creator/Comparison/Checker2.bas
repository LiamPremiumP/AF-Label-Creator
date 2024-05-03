Sub Layout4()

    Dim RowNo As Double     'Number of rows containing data, counted from column A (Specified column if printing specific labels)
    Dim IE As Variant   'The chosen layout
    Dim check As Variant
    Dim check2 As Variant
    Dim check3 As String
    Dim j As Integer
    j = 1
    
    check = 0
    check2 = 0
    check3 = Sheets("IEEE 1584 2018").Cells(5, 1).Value
    


    RowNoE = (Sheets("IEEE 1584 2018").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) +4
    RowNoEP = (Sheets("WorstCase E+P").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count)

    For i = 5 To RowNoE
        Sheets("Checker2").Rows(i - 4).Value = Sheets("IEEE 1584 2018").Rows(i).Value
        For k = 1 To RowNoEP
            If Sheets("WorstCase E+P").Cells(k, 1).Value = Sheets("IEEE 1584 2018").Cells(i, 1).Value  Then
                If Sheets("WorstCase E+P").Cells(k, 19).Value = Round(Sheets("IEEE 1584 2018").Cells(i, 19).Value, 2) Or Sheets("WorstCase E+P").Cells(k, 19).Value = Round(Sheets("IEEE 1584 2018").Cells(i, 19).Value, 1)Then
                
                Sheets("Checker2").Rows(i - 4).Interior.ColorIndex = 4
                Else 
                Sheets("Checker2").Rows(i - 4).Interior.ColorIndex = 3
                End If
            End If
        Next k
    Next i

End Sub

