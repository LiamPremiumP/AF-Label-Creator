


Sub Layout1()
    'Autolabels V3.2 by Stephen Sheridan 13/08/2019
    'Edited By Liam Mills 25/03/24
    ' Declare the variables
    
    
    Dim i As Integer        'Counter for the outer loop
    Dim k As Integer        'Counter for the inner loop
    Dim num1 As Double      'Used for setting decimal points
    Dim num2 As Double      'Num1 after rounding
    Dim rnd As Integer      'Number of decimal points to round to
    
    Dim RowNo As Double     'Number of rows containing data, counted from column A (Specified column if printing specific labels)
    Dim Layout As Integer   'The chosen layout
    Dim sCol As Integer     'The column containing the relevant text box numbers
    Dim box As Integer      'Textbox number
    Dim Row As Integer      'Used for printing specific labels
    Dim trow As Integer
    Dim w1 As Double        'Declare and set the Cut contour dimensions
    Dim h1 As Double
    Dim sRange As String
    sRange = Sheets("Control").Range("B50").Value    'Labels to print
    Dim listCol As String   'Column for the print specific list
    listCol = Sheets("Control").Range("C8").Value
    trow = 5
    
    
    
           
    
    
        
        'Count the number of rows with data
    If sRange = 3 Then
    RowNo = Sheets("Data").Columns(listCol).Cells.SpecialCells(xlCellTypeConstants).Count + 3 'If print specific range is set
    
    ElseIf sRange = 2 Then
    RowNo = (Sheets("Data").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) + 4 'Print all labels
    
        
    ElseIf sRange = 1 Then 'Check if print just one label is set
    RowNo = 5
    End If
        
        
    
        'Prompt to choose a folder to save the labels in
        
     MsgBox "Please make sure PowerPoint is open with a blank presentation"
    
    Dim fldr As FileDialog
        Dim sItem As String
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .Title = "Select a Folder to save your labels."
            .AllowMultiSelect = False
            .InitialFileName = Application.DefaultFilePath
            If .Show <> -1 Then GoTo NextCode
            sItem = .SelectedItems(1)
        End With
NextCode:
        GetFolder = sItem
        Set fldr = Nothing
    
     
        
        'Check to see if the date is required
    If box = 0 Then GoTo NoDate
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = Sheets("Control").Range("F16").Value 'Date
NoDate:
    
    
        
        
        'Start the outer loop.
    'The outer loop inserts the label number and saves the label as a pdf.
    
       For i = 5 To RowNo
       
    


    If sRange = 3 Then 'Check if print specific range is set
    Row = Sheets("Data").Cells(trow, listCol).Value
    ElseIf sRange = 2 Then ' Otherwise print all labels
    Row = i
    End If
    If sRange = 1 Then 'Check if print just one label is set
    Row = Sheets("Control").Range("C8").Value
    End If
        
            'Check which layout is required and initialise the location variables
            Layout = Sheets("Control").Range("C22").Value 'Sets the layout

LayoutChange:
        
            sCol = 3 + (Layout * 7) 'Sets the column where the textbox numbers are
            box = Sheets("Control").Cells(7, sCol).Value 'The setting value / textbox number
            

            w1 = Sheets("Control").Cells(3, sCol).Value * 2.83465
            h1 = Sheets("Control").Cells(4, sCol).Value * 2.83465
    
    If (Sheets("Data").Cells(Row, 19).Value) > 40 And Layout = 10 Then 'Changes layout

        Layout = 11
        GoTo LayoutChange
    ElseIf (Sheets("Data").Cells(Row, 19).Value) <= 40 And Layout = 11 Then
        Layout = 10
        GoTo LayoutChange
    End If
    
    
        'Start the inner loop.
    'The inner loop places the data into all other textboxes before the outer loop saves the label
    
   ' Sheets("Layout" & Layout).Shapes("TextBox " & 27).TextFrame.Characters.Text = Sheets("Control").Cells(16, 6).Value
    Dim ControlToDatadata(40) As Variant
    ControlToDatadata(9 - 9) = Sheets("Data").Cells(Row, 1).Value
    Dim IdLength As Integer
    IdLength = Len(ControlToDatadata(9 - 9))

    ControlToDatadata(11 - 9) = Round((Sheets("Data").Cells(Row, 3).Value), 1)
    
    Dim voltage As Integer
    voltage = ControlToDatadata(11 - 9) * 1000

    ControlToDatadata(11 - 9) = Round((Sheets("Data").Cells(Row, 3).Value), 1) 'Redefines to get the integer
    ControlToDatadata(25 - 9) = Round((Sheets("Data").Cells(Row, 17).Value), 0)
    ControlToDatadata(26 - 9) = "Worst Case Arc Incident Energy at Nominal Working Distance (" & Round(Sheets("Data").Cells(Row, 18).Value, 0) & " cm from source)"
    
    Dim ppe As Double
    ppe = Round((Sheets("Data").Cells(Row, 19).Value), 1)
    If ppe = 1.2 Then
    ControlToDatadata(27 - 9) = "<1.2" ' ppe (cal/cm)
    ElseIf ppe = 12 Then
    ControlToDatadata(27 - 9) = "<12.0" ' ppe (cal/cm)
    ElseIf ppe = 40 Then
    ControlToDatadata(27 - 9) = "<40.0" ' ppe (cal/cm)
    End If
    

    ControlToDatadata(28 - 9) = Sheets("Data").Cells(Row, 20).Value


    ControlToDatadata(35 - 9) = Sheets("Data").Cells(Row, 25).Value 'Maintenance PPE
    Dim mppe As Double
    mppe = ControlToDatadata(35 - 9)

    ControlToDatadata(36 - 9) = Sheets("Data").Cells(Row, 29).Value
    
   

    Dim placeh As Integer
    For placeh = 0 To 16
                    If voltage > Sheets("Register").Cells(placeh + 26, 1).Value And voltage <= Sheets("Register").Cells(placeh + 27, 1).Value Then
                    
                        ControlToDatadata(33 - 9) = Sheets("Register").Cells(placeh + 26, 3).Value * 100 'limited approach
                        ControlToDatadata(34 - 9) = Sheets("Register").Cells(placeh + 26, 4).Value  'restricted apporach
                    End If
    Next placeh

    For k = 9 To 40
    
        box = Sheets("Control").Cells(k, sCol).Value
        
                    If box = 0 Then GoTo Skip 'Check to see if field is required
    
    If Sheets("Control").Cells(k, sCol + 1).Value = 1 Then 'Checks if its on or not
        
    'Makes equipment id change size and renames if too long
        
        If k = 9 And IdLength <= 40 Then
        Sheets("Layout" & Layout).Shapes("TextBox " & box).TextFrame.Characters.Font.Size = 14
        Sheets("Layout" & Layout).Shapes("TextBox " & box).TextFrame.VerticalAlignment = xlTop
        ElseIf k = 9 And IdLength > 40 Then
        

            
        Dim titlearray As String
        titlearray = ControlToDatadata(9 - 9)
       ControlToDatadata(9 - 9) = ""
        Dim titleloop As Integer
            For titleloop = 1 To IdLength
                If Mid(titlearray, titleloop, 1) = "(" Then
                    ControlToDatadata(9 - 9) = ControlToDatadata(9 - 9) + "(INC LineSide)"
                    GoTo titleskip
                End If
                ControlToDatadata(9 - 9) = ControlToDatadata(9 - 9) + Mid(titlearray, titleloop, 1)
            Next titleloop
        End If
titleskip:

    'Adds or removes maintenance PPE IF the same
    If (mppe = ppe) And (k = 35) And (Layout = 10) Then
        Sheets("Layout" & Layout).Shapes("TextBox " & box).TextFrame.Characters.Text = " "
        GoTo Skip
    End If


   
        'if not just send the data to the textbox
     
        Sheets("Layout" & Layout).Shapes("TextBox " & box).TextFrame.Characters.Text = (Sheets("Control").Cells(k, sCol + 3).Value) & ControlToDatadata(k - 9) & (Sheets("Control").Cells(k, sCol + 2).Value)
        
        'Makes equipment id change size
        
        For placeh = 0 To 9
            If ControlToDatadata(36 - 9) = Sheets("Register").Cells(placeh + 25, 6).Value Then
                Sheets("Layout" & Layout).Shapes("Rectangle 26").Fill.ForeColor.RGB = Sheets("Register").Cells(placeh + 25, 7).Interior.Color

            End If
        Next placeh


        'Makes part of maintenance ppe green
        If k = 35 Then
            Sheets("Layout" & Layout).Shapes("TextBox " & box).TextFrame.Characters(35, 13).Font.ColorIndex = 10
        End If
    
    
    'Sheets("Layout" & Layout).Shapes("TextBox " & 31).TextFrame.Characters.Text = "Document Number: <br />" & Sheets("Control").Cells(25, 3).Value
    
    End If
    
    
    
Skip:
    
    Next k
    
    
        
        'Insert the label number
        
        box = Sheets("Control").Cells(8, sCol).Value
        If box = 0 Then GoTo Skip2 'Skip if not required
    
        If box > 0 Then
        Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = (Sheets("Control").Range("F6").Value) & _
        (((Sheets("Control").Range("F10").Value) + i) - 2) & (Sheets("Control").Range("F8").Value)
        
    End If
Skip2:
            
            
            
        'This section copies the label to PowerPoint and saves it as a pdf
     
      
            Dim newPowerPoint As PowerPoint.Application 'Declare PowerPoint as an app
            Dim activeSlide As PowerPoint.Slide 'Declare the slide
            Dim ActivePresentation As PowerPoint.Presentation 'Declare the presentation
            Dim activeImage As Shape 'Declare the shape
    
       
            Set newPowerPoint = GetObject(, "PowerPoint.Application")
            newPowerPoint.Visible = True
            Set activeSlide = newPowerPoint.ActivePresentation.Slides(1)
            Set ActivePresentation = newPowerPoint.ActivePresentation
            
           
            If activeSlide.Shapes.Count < 3 Then GoTo shape3 'Check to see if the slide has an image already
            ActivePresentation.Slides(1).Shapes(2).Delete 'If yes delete the images
shape3:
            If activeSlide.Shapes.Count < 2 Then GoTo shape2
            ActivePresentation.Slides(1).Shapes(1).Delete
            
            
shape2:
            If activeSlide.Shapes.Count < 1 Then GoTo shape1
            ActivePresentation.Slides(1).Shapes(1).Delete
            
            
shape1:
    
           
    
            
    With ActivePresentation.PageSetup 'Set the dimensions of the slide
        
        .SlideWidth = Sheets("Control").Cells(3, sCol).Value * 2.83465
    
        .SlideHeight = Sheets("Control").Cells(4, sCol).Value * 2.83465
    
    End With
    
        Sheets("Layout" & Layout).Activate
        Sheets("Layout" & Layout).Shapes.SelectAll 'Select the label
        Selection.Copy ' Copy
        activeSlide.Shapes.PasteSpecial DataType:=2        '2 = ppPasteEnhancedMetafile
        With ActivePresentation.Slides(1).Shapes(1) 'Set the size of the image to match the slide
        .Width = Sheets("Control").Cells(3, sCol).Value * 2.83465
        .Height = Sheets("Control").Cells(4, sCol).Value * 2.83465
        .Left = ActivePresentation.PageSetup.SlideWidth / 2 - .Width / 2 'Centre the image
        .Top = ActivePresentation.PageSetup.SlideHeight / 2 - .Height / 2
        
        If Sheets("Control").Range("B46").Value = 1 Then GoTo nosafe
        
        'Add the cut contour
       ActivePresentation.Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _
       Left:=((w1 / 2) - (w1 / 2)) + 8.50394, Top:=((h1 / 2) - (h1 / 2)) + 8.50394, Width:=w1 - 17.0079, Height:=h1 - 17.0079 'Create a shape and set the dimensions
       ActivePresentation.Slides(1).Shapes(2).Fill.Visible = msoFalse 'Set fill as transparent
       ActivePresentation.Slides(1).Shapes(2).Line.ForeColor.RGB = RGB(238, 42, 152) 'Set the outline colour
       ActivePresentation.Slides(1).Shapes(2).Line.Weight = 0.5 'Set the outline weight
       
nocut:
    
        If Sheets("Control").Range("B46").Value = 3 Then GoTo nosafe
       
           'Add the safe area
       ActivePresentation.Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _
       Left:=((w1 / 2) - (w1 / 2)) + 17.0079, Top:=((h1 / 2) - (h1 / 2)) + 17.0079, Width:=w1 - 34.0157, Height:=h1 - 34.0157 'Create a shape and set the dimensions
       ActivePresentation.Slides(1).Shapes(3).Fill.Visible = msoFalse 'Set fill as transparent
       ActivePresentation.Slides(1).Shapes(3).Line.ForeColor.RGB = RGB(0, 255, 0) 'Set the outline colour
       ActivePresentation.Slides(1).Shapes(3).Line.Weight = 0.5 'Set the outline weight
       
nosafe:
       
        End With
        
    
    
    
        'End With
        
       
    'Save As PDF Document
        ActivePresentation.ExportAsFixedFormat GetFolder & "\" & i & Sheets("Control").Range("C24").Value & "-" & Sheets("Data").Cells(Row, 1).Value & ".pdf", _
          ppFixedFormatTypePDF, ppFixedFormatIntentPrint, msoCTrue, ppPrintHandoutHorizontalFirst, _
          ppPrintOutputSlides, msoFalse, , ppPrintAll, , False, False, False, False, False
    
    
        
        'Display remaining labels
        
        Application.StatusBar = RowNo - i & " Labels left to be created"
        
    trow = trow + 1
    
    Next i
    
    Sheets("Control").Activate
    
    GoTo Finish
    
    
Error1:
    MsgBox "There was an error, please check your settings"
    GoTo Exit1
    
    
    
Finish:
    MsgBox "Your labels are ready. Have a nice day :)"
    
    
Exit1:
    End Sub
    
    
    
    





    






