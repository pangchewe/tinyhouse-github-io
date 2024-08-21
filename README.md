ihi<h2 align="center"> Web demo for "Tiny House Project"</h2>



## Author
pangchewe

## Description
<a href="https://pangchewe.github.io/tiny-house/" target="_blank"> THANIIE </a> is a website to showcase Tiny House Project. <!-- Built with love -->

## License
THANIIE is licensed under the **MIT License**.



Sub Update_Input()
'
' Update_Input Macro
'
Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End

    Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7) = "Y" Then
      
            Path = Sheets("Input").Cells(i, 6)
            Filename = Sheets("Input").Cells(i, 3)
            From_File = Path + "\" + Filename
            From_Tab = Sheets("Input").Cells(i, 4)
            From_Range = Sheets("Input").Cells(i, 5)
            To_Tab = Sheets("Input").Cells(i, 1)
            To_Range = Sheets("Input").Cells(i, 2)
                  
            If Dir(From_File) = "" Then
                MsgBox (From_File + " not found")
            Else
                        
                Sheets(To_Tab).Range(To_Range).Cells.ClearContents
                'Sheets(To_Tab).Range(To_Range).Cells.ClearFormats
            
                Workbooks.Open Filename:=From_File, UpdateLinks:=0
                Workbooks(Filename).Worksheets(From_Tab).Range(From_Range).Cells.Copy
                ThisWorkbook.Activate
           
                Sheets(To_Tab).Select
                Range(To_Range).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                'Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           
                Sheets("Input").Select
                Cells(i, 8).Select
                Selection.Value = "Finished at " & DateTime.Date & " " & DateTime.Time
                Selection.Font.ColorIndex = 3
                Application.CutCopyMode = False
           
                Workbooks(Filename).Close (False)
           
           End If
        
        End If

    Application.Calculation = xlAutomatic
        
    Next i


'
End Sub













=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AY$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AX$3))
to
=-(SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$AG:$AG, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$C:$C, "6100000100", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$K:$K, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$A:$A, '[MTD Results Summary]MTD Results Summary'!AY$3, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!AH:AH, "AGENCY") - SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$AG:$AG, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$C:$C, "6100000100", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$K:$K, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$A:$A, '[MTD Results Summary]MTD Results Summary'!AX$3, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!AH:AH, "AGENCY"))
so make this 
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AX$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AW$3))

1.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AW$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AV$3))
2.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AV$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AU$3))

3.=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AU$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AT$3))
4.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AT$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AS$3))
5.=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AS$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AR$3))


Sub FnGetSheetsName()

    Dim mainworkBook As Workbook

    Set mainworkBook = ActiveWorkbook

    For i = 1 To mainworkBook.Sheets.Count

    'Either we can put all names in an array , here we are printing all the names in Sheet 2

    mainworkBook.Sheets("Sheet2").Range("A" & i) = mainworkBook.Sheets(i).Name

    Next i

End Sub

background
your file is now require various of data source. VBA could help us to update those data by 1 click, and that is your work
 
Todo so, please use VBA from 1st vba as your reference to develop the VBA in your file
you need to list all the data source (tab name) needed ,and number of columns in scope of updating. The data source path is not available for now, you can leave it blank first, and fill in the target tab and column to paste the info from data source.
just roughly is fine, we will look on it together later.


Sub Update_Input()
'
' Update_Input Macro
'
Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End

    Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7) = "Y" Then
      
            Path = Sheets("Input").Cells(i, 6)
            Filename = Sheets("Input").Cells(i, 3)
            From_File = Path + "\" + Filename
            From_Tab = Sheets("Input").Cells(i, 4)
            From_Range = Sheets("Input").Cells(i, 5)
            To_Tab = Sheets("Input").Cells(i, 1)
            To_Range = Sheets("Input").Cells(i, 2)
                  
            If Dir(From_File) = "" Then
                MsgBox (From_File + " not found")
            Else
                        
                Sheets(To_Tab).Range(To_Range).Cells.ClearContents
                'Sheets(To_Tab).Range(To_Range).Cells.ClearFormats
            
                Workbooks.Open Filename:=From_File, UpdateLinks:=0
                Workbooks(Filename).Worksheets(From_Tab).Range(From_Range).Cells.Copy
                ThisWorkbook.Activate
           
                Sheets(To_Tab).Select
                Range(To_Range).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                'Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           
                Sheets("Input").Select
                Cells(i, 8).Select
                Selection.Value = "Finished at " & DateTime.Date & " " & DateTime.Time
                Selection.Font.ColorIndex = 3
                Application.CutCopyMode = False
           
                Workbooks(Filename).Close (False)
           
           End If
        
        End If

    Application.Calculation = xlAutomatic
        
    Next i


'
End Sub



Sub UpdateSheetList()

    Dim ws As Worksheet
    Dim mainWorkbook As Workbook
    Dim sheetIndex As Long
    Dim lastRow As Long
    
    Set mainWorkbook = ActiveWorkbook
    
    ' Assuming the list starts at row 7 based on the provided image
    sheetIndex = 7
    
    ' Clear previous data starting from row 7 to avoid duplicates
    mainWorkbook.Sheets("INPUT").Rows("7:" & mainWorkbook.Sheets("INPUT").Rows.Count).ClearContents

    ' Loop through each sheet in the workbook
    For Each ws In mainWorkbook.Sheets
        
        ' Sheet name in column A
        mainWorkbook.Sheets("INPUT").Cells(sheetIndex, 1).Value = ws.Name
        
        ' Total number of columns in column B, using column A as reference for the last column
        mainWorkbook.Sheets("INPUT").Cells(sheetIndex, 2).Value = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Increase sheetIndex to move to the next row
        sheetIndex = sheetIndex + 1
    Next ws
    
    ' Now, add the current date to B2 as per your screenshot
    mainWorkbook.Sheets("INPUT").Range("B2").Value = VBA.Format(Now, "mm/dd/yyyy")
    
End Sub
​
Sub Update_Input()
'
' Update_Input Macro
'

    Dim ws As Worksheet
    Dim sheetIndex As Long

    ' List all sheet names starting from A7
    sheetIndex = 7 ' Starting row for sheet names
    For Each ws In ThisWorkbook.Sheets
        Sheets("Input").Cells(sheetIndex, 1).Value = ws.Name
        sheetIndex = sheetIndex + 1
    Next ws

    Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End

        Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7) = "Y" Then
      
            Path = Sheets("Input").Cells(i, 6)
            Filename = Sheets("Input").Cells(i, 3)
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4)
            From_Range = Sheets("Input").Cells(i, 5)
            To_Tab = Sheets("Input").Cells(i, 1)
            To_Range = Sheets("Input").Cells(i, 2)
                  
            If Dir(From_File) = "" Then
                MsgBox From_File & " not found"
            Else
                        
                Sheets(To_Tab).Range(To_Range).Cells.ClearContents
            
                Workbooks.Open Filename:=From_File, UpdateLinks:=0
                Workbooks(Filename).Worksheets(From_Tab).Range(From_Range).Copy
                ThisWorkbook.Activate
           
                Sheets(To_Tab).Range(To_Range).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           
                Sheets("Input").Cells(i, 8).Value = "Finished at " & VBA.Date & " " & VBA.Time
                Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
                Application.CutCopyMode = False
           
                Workbooks(Filename).Close SaveChanges:=False
           
            End If
        
        End If

        Application.Calculation = xlAutomatic
        
    Next i

End Sub




Sub Update_Input()
'
' Update_Input Macro
'
Dim ws As Worksheet
Dim i As Long, j As Long
Dim Extract_Start As Long, Extract_End As Long
Dim Path, Filename, From_File, From_Tab, From_Range, To_Tab, To_Range As String

Application.Calculation = xlAutomatic

' New Section: Fill A7, A8, A9, etc., with sheet names
j = 7 ' Start from row 7 in column A
For Each ws In ThisWorkbook.Sheets
    Sheets("Input").Cells(j, 1).Value = ws.Name
    j = j + 1
Next ws

' Existing code
Sheets("Input").Select
Extract_Start = Sheets("Input").Range("B3").Value
Extract_End = Sheets("Input").Range("B4").Value

For i = Extract_Start To Extract_End

    Application.Calculation = xlManual

    If Sheets("Input").Cells(i, 7) = "Y" Then

        Path = Sheets("Input").Cells(i, 6)
        Filename = Sheets("Input").Cells(i, 3)
        From_File = Path & "\" & Filename
        From_Tab = Sheets("Input").Cells(i, 4)
        From_Range = Sheets("Input").Cells(i, 5)
        To_Tab = Sheets("Input").Cells(i, 1)
        To_Range = Sheets("Input").Cells(i, 2)

        If Dir(From_File) = "" Then
            MsgBox (From_File & " not found")
        Else

            Sheets(To_Tab).Range(To_Range).Cells.ClearContents
            'Sheets(To_Tab).Range(To_Range).Cells.ClearFormats

            Workbooks.Open Filename:=From_File, UpdateLinks:=0
            Workbooks(Filename).Worksheets(From_Tab).Range(From_Range).Cells.Copy
            ThisWorkbook.Activate

            Sheets(To_Tab).Select
            Range(To_Range).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            'Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

            Sheets("Input").Select
            Cells(i, 8).Select
            Selection.Value = "Finished at " & DateTime.Date & " " & DateTime.Time
            Selection.Font.ColorIndex = 3
            Application.CutCopyMode = False

            Workbooks(Filename).Close SaveChanges:=False

        End If

    End If

    Application.Calculation = xlAutomatic

Next i

'
End Sub


Sub Update_Input()
'
' Update_Input Macro
'
Dim ws As Worksheet
Dim i As Long, j As Long
Dim Extract_Start As Long, Extract_End As Long
Dim Path, Filename, From_File, From_Tab, From_Range, To_Tab, To_Range As String
Dim lastCol As Long, colRange As String

Application.Calculation = xlAutomatic

' Fill A7, A8, A9, etc., with sheet names and B7, B8, B9... with column range
j = 7 ' Start from row 7 in column A
For Each ws In ThisWorkbook.Sheets
    Sheets("Input").Cells(j, 1).Value = ws.Name
    
    ' Determine the last used column and convert to letter(s)
    With ws
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colRange = "A:" & Split(Cells(1, lastCol).Address, "$")(1)
    End With
    
    Sheets("Input").Cells(j, 2).Value = colRange
    j = j + 1
Next ws

' Existing code for updating inputs
Sheets("Input").Select
Extract_Start = Sheets("Input").Range("B3").Value
Extract_End = Sheets("Input").Range("B4").Value

For i = Extract_Start To Extract_End

    Application.Calculation = xlManual
    
    ' Your existing code for processing each row based on conditions
    
Next i

Application.Calculation = xlAutomatic

End Sub








Sub Update_Input()
'
' Update_Input Macro
'
Dim ws As Worksheet
Dim i As Long, j As Long
Dim Extract_Start As Long, Extract_End As Long
Dim Path, Filename, From_File, From_Tab, From_Range, To_Tab, To_Range As String
Dim firstCol As Long, lastCol As Long, colRange As String

Application.Calculation = xlAutomatic

' Fill A7, A8, A9, etc., with sheet names and B7, B8, B9... with column range
j = 7 ' Start from row 7 in column A
For Each ws In ThisWorkbook.Sheets
    Sheets("Input").Cells(j, 1).Value = ws.Name
    
    ' Determine the first and last used columns and convert to letter(s)
    With ws
        firstCol = .UsedRange.Column
        lastCol = firstCol + .UsedRange.Columns.Count - 1
        colRange = ConvertToLetter(firstCol) & ":" & ConvertToLetter(lastCol)
    End With
    
    Sheets("Input").Cells(j, 2).Value = colRange
    j = j + 1
Next ws

' Existing code for updating inputs
Sheets("Input").Select
Extract_Start = Sheets("Input").Range("B3").Value
Extract_End = Sheets("Input").Range("B4").Value

For i = Extract_Start To Extract_End

    Application.Calculation = xlManual
    
    ' Your existing code for processing each row based on conditions
    
Next i

Application.Calculation = xlAutomatic

End Sub

' Function to convert column number to letter
Function ConvertToLetter(iCol As Long) As String
    Dim iAlpha As Integer
    Dim iRemainder As Integer
    iAlpha = Int(iCol / 27)
    iRemainder = iCol - (iAlpha * 26)
    If iAlpha > 0 Then
        ConvertToLetter = Chr(iAlpha + 64)
    End If
    If iRemainder > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
    End If
End Function


Sub Update_Input()
'
' Update_Input Macro
'
Dim ws As Worksheet
Dim i As Long, j As Long
Dim Extract_Start As Long, Extract_End As Long
Dim Path, Filename, From_File, From_Tab, From_Range, To_Tab, To_Range As String
Dim firstCol As Long, lastCol As Long
Dim colRange As String, tempFirstCol As Long, tempLastCol As Long

Application.Calculation = xlAutomatic

j = 7 ' Start from row 7 in column A on the Input sheet
For Each ws In ThisWorkbook.Sheets
    Sheets("Input").Cells(j, 1).Value = ws.Name
    firstCol = ws.Columns.Count
    lastCol = 1
    
    For i = 1 To 20 ' Check from row 1 to row 20
        ' Temporarily find the first and last column with data in the current row
        tempFirstCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
        tempLastCol = ws.Cells(i, 1).End(xlToRight).Column
        
        ' Update firstCol and lastCol based on the found columns
        If tempFirstCol < firstCol And tempFirstCol <> 1 Then firstCol = tempFirstCol
        If tempLastCol > lastCol And tempLastCol <> ws.Columns.Count Then lastCol = tempLastCol
    Next i

    ' Convert column numbers to letters
    If firstCol <= lastCol Then
        colRange = ConvertToLetter(firstCol) & ":" & ConvertToLetter(lastCol)
    Else
        colRange = "N/A" ' In case there's no data within the first 20 rows
    End If

    Sheets("Input").Cells(j, 2).Value = colRange
    j = j + 1
Next ws

' Existing code for updating inputs (unchanged)
Sheets("Input").Select
Extract_Start = Sheets("Input").Range("B3").Value
Extract_End = Sheets("Input").Range("B4").Value

For i = Extract_Start To Extract_End
    ' Your existing processing logic...
Next i

Application.Calculation = xlAutomatic
End Sub

' Function to convert column number to letter
Function ConvertToLetter(iCol As Long) As String
    Dim iAlpha As Integer
    Dim iRemainder As Integer
    iAlpha = Int(iCol / 27)
    iRemainder = iCol - (iAlpha * 26)
    If i



Sub Update_Input()
'
' Update_Input Macro
'
Dim ws As Worksheet
Dim i As Long, j As Long
Dim Extract_Start As Long, Extract_End As Long
Dim Path, Filename, From_File, From_Tab, From_Range, To_Tab, To_Range As String
Dim firstCol As Long, lastCol As Long
Dim colRange As String, tempFirstCol As Long, tempLastCol As Long

Application.Calculation = xlAutomatic

j = 7 ' Start from row 7 in column A on the Input sheet
For Each ws In ThisWorkbook.Sheets
    Sheets("Input").Cells(j, 1).Value = ws.Name
    firstCol = ws.Columns.Count
    lastCol = 1
    
    For i = 1 To 20 ' Check from row 1 to row 20
        ' Temporarily find the first and last column with data in the current row
        tempFirstCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
        tempLastCol = ws.Cells(i, 1).End(xlToRight).Column
        
        ' Update firstCol and lastCol based on the found columns
        If tempFirstCol < firstCol And tempFirstCol <> 1 Then firstCol = tempFirstCol
        If tempLastCol > lastCol And tempLastCol <> ws.Columns.Count Then lastCol = tempLastCol
    Next i

    ' Convert column numbers to letters
    If firstCol <= lastCol Then
        colRange = ConvertToLetter(firstCol) & ":" & ConvertToLetter(lastCol)
    Else
        colRange = "N/A" ' In case there's no data within the first 20 rows
    End If

    Sheets("Input").Cells(j, 2).Value = colRange
    j = j + 1
Next ws

' Existing code for updating inputs (unchanged)
Sheets("Input").Select
Extract_Start = Sheets("Input").Range("B3").Value
Extract_End = Sheets("Input").Range("B4").Value

For i = Extract_Start To Extract_End
    ' Your existing processing logic...
Next i

Application.Calculation = xlAutomatic
End Sub

' Function to convert column number to letter
Function ConvertToLetter(iCol As Long) As String
    Dim iAlpha As Integer
    Dim iRemainder As Integer
    iAlpha = Int(iCol / 27)
    iRemainder = iCol - (iAlpha * 26)
    If iAlpha > 0 Then
        ConvertToLetter = Chr(iAlpha + 64)
    End If
    If iRemainder > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
    End If
End Function





Sub Update_SheetInfo()
'
' Update_SheetInfo Macro
'
Dim ws As Worksheet
Dim i As Integer, j As Integer
Dim firstCol As Integer, lastCol As Integer
Dim colLetterFirst As String, colLetterLast As String
Dim rng As Range
Dim maxFirstCol As Integer, maxLastCol As Integer

j = 7 ' Starting row in the "Input" sheet
For Each ws In ThisWorkbook.Sheets
    maxFirstCol = ws.Columns.Count
    maxLastCol = 1

    ' Determine the max range from row 1 to 20
    For i = 1 To 20
        Set rng = ws.Rows(i).Find(What:="*", After:=ws.Cells(i, 1), LookAt:=xlPart, LookIn:=xlFormulas)
        If Not rng Is Nothing Then
            firstCol = rng.Column
            If firstCol < maxFirstCol Then maxFirstCol = firstCol
        End If

        Set rng = ws.Rows(i).Find(What:="*", After:=ws.Cells(i, ws.Columns.Count), LookAt:=xlPart, _
                      LookIn:=xlFormulas, SearchDirection:=xlPrevious)
        If Not rng Is Nothing Then
            lastCol = rng.Column
            If lastCol > maxLastCol Then maxLastCol = lastCol
        End If
    Next i

    ' Convert column number to letter
    colLetterFirst = Split(ws.Cells(1, maxFirstCol).Address(True, False), "$")(0)
    colLetterLast = Split(ws.Cells(1, maxLastCol).Address(True, False), "$")(0)

    ' Fill in sheet name and max used range
    Sheets("Input").Cells(j, 1).Value = ws.Name
    Sheets("Input").Cells(j, 2).Value = colLetterFirst & ":" & colLetterLast
    j = j + 1
Next ws

End Sub



Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7).Value = "Y" Then
            Path = Sheets("Input").Cells(i, 6).Value
            Filename = Sheets("Input").Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4).Value
            From_Range = Sheets("Input").Cells(i, 5).Value
            To_Tab = Sheets("Input").Cells(i, 1).Value
            
            If Dir(From_File) = "" Then
                MsgBox (From_File & " not found")
            Else
                ' Find the last row in the destination tab
                With Sheets(To_Tab)
                    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
                End With
            
                ' Open the source workbook
                Workbooks.Open Filename:=From_File, UpdateLinks:=0
                Workbooks(Filename).Worksheets(From_Tab).Range(From_Range).Cells.Copy
                
                ' Paste into the destination workbook after the last row
                ThisWorkbook.Activate
                Sheets(To_Tab).Select
                Sheets(To_Tab).Cells(LastRow, 1).PasteSpecial _
                    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                ' Mark the operation as completed with timestamp
                Sheets("Input").Select
                Sheets("Input").Cells(i, 8).Value = "Finished at " & Now
                Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
                Application.CutCopyMode = False
                
                ' Close the source workbook
                Workbooks(Filename).Close (False)
            End If
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub



Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7).Value = "Y" Then
            Path = Sheets("Input").Cells(i, 6).Value
            Filename = Sheets("Input").Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4).Value
            From_Range = Sheets("Input").Cells(i, 5).Value
            To_Tab = Sheets("Input").Cells(i, 1).Value
            
            If Dir(From_File) = "" Then
                MsgBox (From_File & " not found")
            Else
                ' Find the last row in the destination tab
                With Sheets(To_Tab)
                    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
                End With
            
                ' Open the source workbook
                Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
                SourceRange = SourceWorkbook.Worksheets(From_Tab).Range(From_Range)
                
                ' Ensure correct size when pasting
                NumColumns = SourceRange.Columns.Count
                
                ' Select appropriate destination range
                ThisWorkbook.Activate
                With Sheets(To_Tab)
                    DestinationRange = .Cells(LastRow, 1).Resize(SourceRange.Rows.Count, NumColumns)
                End With
                
                ' Paste into the destination range
                SourceRange.Copy
                DestinationRange.PasteSpecial _
                    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                ' Mark the operation as completed with timestamp
                Sheets("Input").Select
                Sheets("Input").Cells(i, 8).Value = "Finished at " & Now
                Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
                Application.CutCopyMode = False
                
                ' Close the source workbook
                SourceWorkbook.Close (False)
            End If
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub

Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7).Value = "Y" Then
            Path = Sheets("Input").Cells(i, 6).Value
            Filename = Sheets("Input").Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4).Value
            From_Range = Sheets("Input").Cells(i, 5).Value
            To_Tab = Sheets("Input").Cells(i, 1).Value
            To_Range = Sheets("Input").Cells(i, 2).Value
            
            If Dir(From_File) = "" Then
                MsgBox From_File & " not found"
            Else
                ' Find the last row in the destination tab
                With Sheets(To_Tab)
                    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
                End With
                
                ' Open the source workbook
                Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
                
                ' Find the range to copy
                Set SourceRange = SourceWorkbook.Worksheets(From_Tab).Range(From_Range)
                
                ' Determine destination range based on the last row
                Set DestinationRange = Sheets(To_Tab).Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
                
                ' Paste the data into the correct location
                SourceRange.Copy
                DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                ' Mark the task as completed with a timestamp
                Sheets("Input").Cells(i, 8).Value = "Finished at " & Now
                Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
                
                ' Reset cut/copy mode and close the source workbook
                Application.CutCopyMode = False
                SourceWorkbook.Close SaveChanges:=False
            End If
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub




Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    Sheets("Input").Select
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    
    ' Ensure valid extract start and end values
    If IsNumeric(Extract_Start) = False Or IsNumeric(Extract_End) = False Then
        MsgBox "Invalid extract start or end values."
        Exit Sub
    End If
    
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
    
        If Sheets("Input").Cells(i, 7).Value = "Y" Then
            Path = Sheets("Input").Cells(i, 6).Value
            Filename = Sheets("Input").Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4).Value
            From_Range = Sheets("Input").Cells(i, 5).Value
            To_Tab = Sheets("Input").Cells(i, 1).Value
            
            ' Ensure the "To_Tab" worksheet exists
            On Error Resume Next
            If Sheets(To_Tab) Is Nothing Then
                MsgBox "Worksheet '" & To_Tab & "' not found."
                Exit Sub
            End If
            On Error GoTo 0
            
            ' Ensure the source file exists
            If Dir(From_File) = "" Then
                MsgBox From_File & " not found"
                Exit Sub
            End If
            
            ' Find the last row in the destination tab
            With Sheets(To_Tab)
                LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            End With
            
            ' Open the source workbook
            Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
            
            ' Ensure the "From_Tab" worksheet exists
            If Not SourceWorkbook.Worksheets(From_Tab) Is Nothing Then
                ' Find the source range to copy
                Set SourceRange = SourceWorkbook.Worksheets(From_Tab).Range(From_Range)
                
                ' Determine destination range based on the last row
                Set DestinationRange = Sheets(To_Tab).Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
                
                ' Paste the data into the correct location
                SourceRange.Copy
                DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                ' Mark the task as completed with a timestamp
                Sheets("Input").Cells(i, 8).Value = "Finished at " & Now
                Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
                
                ' Reset cut/copy mode and close the source workbook
                Application.CutCopyMode = False
                SourceWorkbook.Close SaveChanges:=False
            Else
                MsgBox "Worksheet '" & From_Tab & "' not found in source file."
            End If
            
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub

Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
    
    ' Validate Extract_Start and Extract_End
    On Error Resume Next
    Extract_Start = Sheets("Input").Range("B3").Value
    Extract_End = Sheets("Input").Range("B4").Value
    On Error GoTo 0
    
    ' Validate that they are numeric and within valid bounds
    If IsNumeric(Extract_Start) = False Or IsNumeric(Extract_End) = False Then
        MsgBox "Invalid start or end values in Input sheet."
        Exit Sub
    End If
    
    If Extract_Start < 7 Or Extract_End < Extract_Start Then
        MsgBox "Invalid range in Input sheet."
        Exit Sub
    End If
    
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
        
        ' Check if row index is valid
        If i > Sheets("Input").Rows.Count Then
            MsgBox "Row index out of range."
            Exit Sub
        End If
    
        ' Validate if "Y" is in the correct cell
        If Sheets("Input").Cells(i, 7).Value = "Y" Then
            Path = Sheets("Input").Cells(i, 6).Value
            Filename = Sheets("Input").Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = Sheets("Input").Cells(i, 4).Value
            From_Range = Sheets("Input").Cells(i, 5).Value
            To_Tab = Sheets("Input").Cells(i, 1).Value
            To_Range = Sheets("Input").Cells(i, 2).Value
            
            ' Check if the source file exists
            If Dir(From_File) = "" Then
                MsgBox From_File & " not found."
                Exit Sub
            End If
            
            ' Validate "To_Tab" worksheet
            If Sheets(To_Tab) Is Nothing Then
                MsgBox "Worksheet '" & To_Tab & "' not found."
                Exit Sub
            End If
            
            ' Validate "From_Tab" in the source workbook
            On Error Resume Next
            Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
            If SourceWorkbook.Worksheets(From_Tab) Is Nothing Then
                MsgBox "Worksheet '" & From_Tab & "' not found in source workbook."
                SourceWorkbook.Close
                Exit Sub
            End If
            On Error GoTo 0
            
            ' Find the last row in the "To_Tab"
            With Sheets(To_Tab)
                LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            End With
            
            ' Find the source range to copy
            Set SourceRange = SourceWorkbook.Worksheets(From_Tab).Range(From_Range)
            
            ' Determine destination range based on the last row
            Set DestinationRange = Sheets(To_Tab).Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            
            ' Paste data as values to avoid overwriting
            SourceRange.Copy
            DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            ' Mark task as completed with a timestamp
            Sheets("Input").Cells(i, 8).Value = "Finished at " & Now
            Sheets("Input").Cells(i, 8).Font.ColorIndex = 3
            
            ' Reset cut/copy mode and close source workbook without saving
            Application.CutCopyMode = False
            SourceWorkbook.Close (False)
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub



Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    ' Validate the "Input" sheet
    On Error Resume Next
    Set InputSheet = Sheets("Input")
    On Error GoTo 0
    If InputSheet Is Nothing Then
        MsgBox "Sheet 'Input' not found."
        Exit Sub
    End If
    
    ' Ensure start and end values are numeric and valid
    Extract_Start = InputSheet.Range("B3").Value
    Extract_End = InputSheet.Range("B4").Value
    
    If Not IsNumeric(Extract_Start) Or Not IsNumeric(Extract_End) Or Extract_Start > Extract_End Then
        MsgBox "Invalid range in 'Input'."
        Exit Sub
    End If
    
    ' Loop through the range
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
        
        ' Validate row index
        If i < 7 Or i > InputSheet.Rows.Count Then
            MsgBox "Row index out of range."
            Exit Sub
        End If
        
        ' Check if operation should proceed
        If InputSheet.Cells(i, 7).Value = "Y" Then
            ' Retrieve data
            Path = InputSheet.Cells(i, 6).Value
            Filename = InputSheet.Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = InputSheet.Cells(i, 4).Value
            To_Tab = InputSheet.Cells(i, 1).Value
            To_Range = InputSheet.Cells(i, 2).Value
            
            ' Validate "To_Tab" exists
            On Error Resume Next
            Set DestinationSheet = Sheets(To_Tab)
            On Error GoTo 0
            If DestinationSheet Is Nothing Then
                MsgBox "Sheet '" & To_Tab & "' not found."
                Exit Sub
            End If
            
            ' Validate source file exists
            If Dir(From_File) = "" Then
                MsgBox "File '" & From_File & "' not found."
                Exit Sub
            End If
            
            ' Open source workbook
            Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
            
            ' Validate "From_Tab" exists
            On Error Resume Next
            Set SourceSheet = SourceWorkbook.Worksheets(From_Tab)
            On Error GoTo 0
            If SourceSheet Is Nothing Then
                MsgBox "Worksheet '" & From_Tab & "' not found in the source workbook."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Find the last row in the "To_Tab"
            LastRow = DestinationSheet.Cells(DestinationSheet.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Determine source range to copy
            On Error Resume Next
            Set SourceRange = SourceSheet.Range(From_Range)
            On Error GoTo 0
            If SourceRange Is Nothing Then
                MsgBox "Invalid range '" & From_Range & "' in source worksheet."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Define destination range based on the last row
            Set DestinationRange = DestinationSheet.Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            
            ' Copy and paste special values
            SourceRange.Copy
            DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            ' Mark the task as completed with a timestamp
            InputSheet.Cells(i, 8).Value = "Finished at " & Now
            InputSheet.Cells(i, 8).Font.ColorIndex = 3
            
            ' Reset cut/copy mode and close the source workbook without saving
            Application.CutCopyMode = False
            SourceWorkbook.Close SaveChanges:=False
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub
Sub Update_Input_Append()
    '
    ' Update_Input_Append Macro
    '
    Application.Calculation = xlAutomatic
   
    ' Validate "Input" sheet exists
    On Error Resume Next
    Set InputSheet = Sheets("Input")
    On Error GoTo 0
    If InputSheet Is Nothing Then
        MsgBox "Sheet 'Input' not found."
        Exit Sub
    End If
    
    ' Get extract start and end
    Extract_Start = InputSheet.Range("B3").Value
    Extract_End = InputSheet.Range("B4").Value
    
    ' Validate start and end values
    If Not IsNumeric(Extract_Start) Or Not IsNumeric(Extract_End) Then
        MsgBox "Invalid extract start or end values."
        Exit Sub
    End If
    
    ' Ensure start and end are within valid bounds
    If Extract_Start < 7 Or Extract_End < Extract_Start Then
        MsgBox "Extract start or end out of range."
        Exit Sub
    End If
    
    ' Loop through the range
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
        
        ' Validate row index
        If i < 7 Or i > InputSheet.Rows.Count Then
            MsgBox "Row index out of range."
            Exit Sub
        End If
    
        ' Check if the operation should proceed
        If InputSheet.Cells(i, 7).Value = "Y" Then
            ' Get file details
            Path = InputSheet.Cells(i, 6).Value
            Filename = InputSheet.Cells(i, 3).Value
            From_File = Path & "\" & Filename
            From_Tab = InputSheet.Cells(i, 4).Value
            To_Tab = InputSheet.Cells(i, 1).Value
            To_Range = InputSheet.Cells(i, 2).Value
            
            ' Validate "To_Tab" sheet exists
            On Error Resume Next
            Set DestinationSheet = Sheets(To_Tab)
            On Error GoTo 0
            If DestinationSheet Is Nothing Then
                MsgBox "Sheet '" & To_Tab & "' not found."
                Exit Sub
            End If
            
            ' Validate source file exists
            If Dir(From_File) = "" Then
                MsgBox "File '" & From_File & "' not found."
                Exit Sub
            End If
            
            ' Open the source workbook
            On Error Resume Next
            Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
            On Error GoTo 0
            
            If SourceWorkbook Is Nothing Then
                MsgBox "Could not open workbook '" & From_File & "'."
                Exit Sub
            End If
            
            ' Validate "From_Tab" sheet exists in the source workbook
            On Error Resume Next
            Set SourceSheet = SourceWorkbook.Worksheets(From_Tab)
            On Error GoTo 0
            If SourceSheet Is Nothing Then
                MsgBox "Sheet '" & From_Tab & "' not found in the source workbook."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Find the last row in the "To_Tab"
            LastRow = DestinationSheet.Cells(DestinationSheet.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Validate source range
            On Error Resume Next
            Set SourceRange = SourceSheet.Range(From_Range)
            On Error GoTo 0
            If SourceRange Is Nothing Then
                MsgBox "Invalid range '" & From_Range & "' in source sheet."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Define destination range
            Set DestinationRange = DestinationSheet.Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            
            ' Copy and paste values
            SourceRange.Copy
            DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            ' Mark the task as completed with a timestamp
            InputSheet.Cells(i, 8).Value = "Finished at " & Now
            InputSheet.Cells(i, 8).Font.ColorIndex = 3
            
            ' Reset cut/copy mode and close source workbook without saving
            Application.CutCopyMode = False
            SourceWorkbook.Close SaveChanges:=False
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub




Sub Update_Input_Append()
    ' Set calculation to automatic
    Application.Calculation = xlAutomatic
    
    ' Reference the "Input" sheet and validate
    On Error Resume Next
    Set InputSheet = Sheets("Input")
    On Error GoTo 0
    If InputSheet Is Nothing Then
        MsgBox "Sheet 'Input' not found."
        Exit Sub
    End If
    
    ' Extract start and end and validate
    Extract_Start = InputSheet.Range("B3").Value
    Extract_End = InputSheet.Range("B4").Value
    
    If Not IsNumeric(Extract_Start) Or Not IsNumeric(Extract_End) Then
        MsgBox "Invalid extract start or end values."
        Exit Sub
    End If
    
    If Extract_Start < 7 Or Extract_End < Extract_Start Then
        MsgBox "Extract start or end out of range."
        Exit Sub
    End If
    
    ' Loop through the specified range
    For i = Extract_Start To Extract_End
        Application.Calculation = xlManual
        
        ' Validate row index
        If i < 7 Or i > InputSheet.Rows.Count Then
            MsgBox "Row index out of range."
            Exit Sub
        End If
    
        ' Check if the operation should proceed
        If InputSheet.Cells(i, 7).Value = "Y" Then
            ' Retrieve path and filename details
            Path = InputSheet.Cells(i, 6).Value
            Filename = InputSheet.Cells(i, 3).Value
            From_File = Path & "\" & Filename
            
            ' Validate source file exists
            If Dir(From_File) = "" Then
                MsgBox "File '" & From_File & "' not found."
                Exit Sub
            End If
            
            ' Open source workbook and validate
            On Error Resume Next
            Set SourceWorkbook = Workbooks.Open(Filename:=From_File, UpdateLinks:=0)
            On Error GoTo 0
            If SourceWorkbook Is Nothing Then
                MsgBox "Cannot open workbook '" & From_File & "'."
                Exit Sub
            End If
            
            ' Get "From_Tab" and validate
            From_Tab = InputSheet.Cells(i, 4).Value
            On Error Resume Next
            Set SourceSheet = SourceWorkbook.Worksheets(From_Tab)
            On Error GoTo 0
            If SourceSheet Is Nothing Then
                MsgBox "Sheet '" & From_Tab & "' not found in the source workbook."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Get "To_Tab" and validate
            To_Tab = InputSheet.Cells(i, 1).Value
            Set DestinationSheet = Sheets(To_Tab)
            If DestinationSheet Is Nothing Then
                MsgBox "Sheet '" & To_Tab & "' not found."
                Exit Sub
            End If
            
            ' Determine the last row in the "To_Tab"
            LastRow = DestinationSheet.Cells(DestinationSheet.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Get the source range and validate
            From_Range = InputSheet.Cells(i, 5).Value
            Set SourceRange = SourceSheet.Range(From_Range)
            If SourceRange Is Nothing Then
                MsgBox "Invalid range '" & From_Range & "' in source sheet."
                SourceWorkbook.Close
                Exit Sub
            End If
            
            ' Define the destination range based on the last row
            Set DestinationRange = DestinationSheet.Cells(LastRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            
            ' Copy and paste values
            SourceRange.Copy
            DestinationRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            ' Mark task as completed with timestamp
            InputSheet.Cells(i, 8).Value = "Finished at " & Now
            InputSheet.Cells(i, 8).Font.ColorIndex = 3
            
            ' Reset cut/copy mode and close source workbook without saving
            Application.CutCopyMode = False
            SourceWorkbook.Close (False)
        End If
        
        Application.Calculation = xlAutomatic
    Next i
End Sub
