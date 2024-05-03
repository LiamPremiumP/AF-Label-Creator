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
    


    RowNoE = (Sheets("WorstCase E+P").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) 


    For i = 1 To RowNoE
        
            If Sheets("WorstCase E+P").Cells(i, 2).Value = "UTILITY+GEN" Then
                Sheets("WorstCase E+P").Cells(i, 20).Value = "S2"
            Elseif Sheets("WorstCase E+P").Cells(i, 2).Value = "UTILITY ONLY" Then
            Sheets("WorstCase E+P").Cells(i, 20).Value = "S0"
            Elseif Sheets("WorstCase E+P").Cells(i, 2).Value = "GEN ONLY" Then
            Sheets("WorstCase E+P").Cells(i, 20).Value = "S1"
            Elseif Sheets("WorstCase E+P").Cells(i, 2).Value = "UPS ONLY" Then
            Sheets("WorstCase E+P").Cells(i, 20).Value = "S3"
            End If
       
    Next i

End Sub

