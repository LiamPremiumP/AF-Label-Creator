Attribute VB_Name = "Layout1"

Sub Layout1()
'Autolabels V3.2 by Stephen Sheridan 13/08/2019

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
Dim sRange As String
sRange = Sheets("Control").Range("B50").Value    'Labels to print
Dim listCol As String   'Column for the print specific list
listCol = Sheets("Control").Range("C8").Value
trow = 5



       
    'Check which layout is required and initialise the location variables
Layout = Sheets("Control").Range("C22").Value 'Sets the layout
sCol = 3 + (Layout * 7) 'Sets the column where the textbox numbers are
box = Sheets("Control").Cells(7, sCol).Value 'The setting value / textbox number

Dim w1 As Double        'Declare and set the Cut contour dimensions
Dim h1 As Double
w1 = Sheets("Control").Cells(3, sCol).Value * 2.83465
h1 = Sheets("Control").Cells(4, sCol).Value * 2.83465


    
    'Count the number of rows with data
If sRange = 3 Then
RowNo = Sheets("Data").Columns(listCol).Cells.SpecialCells(xlCellTypeConstants).Count + 3 'If print specific range is set

ElseIf sRange = 2 Then
RowNo = (Sheets("Data").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count) + 3 'Print all labels

    
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
    
    'Start the inner loop.
'The inner loop places the data into all other textboxes before the outer loop saves the label

        For k = 9 To 40

box = Sheets("Control").Cells(k, sCol).Value

            If box = 0 Then GoTo Skip 'Check to see if field is required


If Sheets("Control").Cells(k, sCol + 1).Value = 99 Then 'Check if rounding is required

    'if not just send the data to the textbox
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = Sheets("Data").Cells(Row, k - 8).Value

End If

    'Otherwise get the decimals required
If Sheets("Control").Cells(k, sCol + 1).Value < 50 Then
    rnd = Sheets("Control").Cells(k, sCol + 1).Value
    num1 = Sheets("Data").Cells(Row, k - 8) 'Get the data
    num2 = Round(num1, rnd) 'Round up to the required number
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = (num2) & (Sheets("Control").Cells(k, sCol + 2).Value) 'Send to the appropriate text box
End If
            
If Sheets("Control").Cells(k, sCol + 1).Value = 98 Then 'Check if incident energy
    num1 = Sheets("Data").Cells(Row, k - 8) 'Get the data
    If num1 > 10 Then
    rnd = 1
    num2 = Round(num1, rnd) 'Round up to the required number
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = (num2) & (Sheets("Control").Cells(k, sCol + 2).Value) 'Send to the appropriate text box
    ElseIf num1 < 10 Then
    rnd = 2
    num2 = Round(num1, rnd) 'Round up to the required number
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = (num2) & (Sheets("Control").Cells(k, sCol + 2).Value)
    End If
End If

Skip:

Next k


    
    'Insert the label number
    
    box = Sheets("Control").Cells(8, sCol).Value
    If box = 0 Then GoTo Skip2 'Skip if not required

    If box > 0 Then
    Sheets("Layout" & Layout).TextBoxes("TextBox " & box).Text = (Sheets("Control").Range("F6").Value) & (((Sheets("Control").Range("F10").Value) + i) - 2) & (Sheets("Control").Range("F8").Value)
    
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
    ActivePresentation.ExportAsFixedFormat GetFolder & "\" & Sheets("Control").Range("C24").Value & Row & ".pdf", _
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

