# Finance_Format_Calculate
Through the use of VBA I helped the Finance Department organize data that had been copied from a website, bringing along multiple ASCII characters.  The worksheet I set up for them allows them to select a button which creates their target worksheet.

![Alt text](relative%20path/to/img.jpg?raw=true "step1")

The next button formats the data as autofit does not work on the copied data.  It also sorts the data to only select the specific school district and replaces ascii character 160 with nothing. Finally it loops through the data summing the dollar amounts for each category and displaying them next to the last row of that category.  A reset button was added to allow for the client to start the process over for the next week. See images for outputs of each button.

```vbscript

    Sub Step1()
    
        Application.ScreenUpdating = False
        Call prepare_new_worksheet
    
    End Sub
    
    Sub Step2()
    
        Application.ScreenUpdating = False
        Call format_wp
        Call sort_FM
        Call convert_amounts
        Call sum_amounts
    
    End Sub
    
    Public Sub prepare_new_worksheet()
    
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
    
        cp.Activate
        
        'creates weekly payment sheet
        Sheets.Add Before:=ActiveSheet
        ActiveSheet.Name = "weekly_payments"
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        
        cp.Activate
    
    End Sub
    
    
    Public Sub format_wp()
    
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        Dim lastrow As Integer
        lastrow = wp.Range("A" & wp.Rows.Count).End(xlUp).Row
        
        wp.Activate
    
        'adjust columns and rows as autofit does not work for this copied data
        Columns("A:A").ColumnWidth = 14
        Columns("B:B").ColumnWidth = 24
        Columns("C:C").ColumnWidth = 10
        Columns("D:D").ColumnWidth = 26.43
        Columns("E:E").ColumnWidth = 10
        Columns("F:F").ColumnWidth = 12.29
        Columns("G:G").ColumnWidth = 10.29
        Columns("H:H").ColumnWidth = 34.29
        Columns("I:I").ColumnWidth = 16.43
        
        Range("A1:I" & lastrow).RowHeight = 22.5
        
    End Sub
    
    
    
    
    Public Sub sort_FM()
    
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        
        lastrow = wp.Range("D" & wp.Rows.Count).End(xlUp).Row
        
        'filter specific school district
        For i = lastrow To 4 Step -1
        
            str1 = Left(wp.Cells(i, 4), 4)
        
            If str1 = "4604" Then
            Else
                Rows(i).EntireRow.Delete
            End If
        
        Next i
        
        
    End Sub
    
    
    Public Sub convert_amounts()
    
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
    
        'replace the ascii 160 character with nothing
        wp.Columns(9).Replace Chr(160), "", xlPart
    
        
    End Sub
    
    
    Public Sub sum_amounts()
    
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        lastrow = wp.Range("A" & wp.Rows.Count).End(xlUp).Row
        Dim amount As Double
        
        wp.Columns.NumberFormat = "0.00"
        Columns(10).ColumnWidth = 11
    
        'sum each category and display summed amount
        For i = 4 To lastrow
        
            If wp.Cells(i, 4) = wp.Cells(i + 1, 4) Then
                amount = amount + wp.Cells(i, 9)
            Else
                amount = amount + wp.Cells(i, 9)
                wp.Cells(i, 10) = amount
                amount = 0
            End If
        
        Next i
    
    End Sub
    
    Public Sub reset()
    
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        
        wp.Delete
        cp.Activate
        
    End Sub
    
    
    Public Sub identify_ascii_chars()
    
        'this macro was only used for exploratory purposes to figure out which ascii character was in the cells within column "I" _
        it was not used in the final product
        
        Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets("weekly_payments")
        Dim mystring As String
        lastrow = wp.Range("D" & wp.Rows.Count).End(xlUp).Row
        mystring = wp.Cells(4, 10).Value
        
        For i = 4 To lastrow
            MsgBox Asc(Mid(mystring, i, 1))
        Next i
    
    End Sub

```
