# transportation_data_cleaning
I am currently data cleaning multiple systems of data from the transportation department to turn them into one system that they can utilize.  They are currently inputting multiple pieces of data into different systems and doubling up their workload.  The current project is to take one column that has addresses, emails, and phone numbers in it and split it into 3 seperate columns. The macro below does this and inputs the data into the same line across for each employee.

```vbscript

Public Sub addr_phone_email()

    Dim ed As Worksheet: Set ed = ThisWorkbook.Worksheets("editable")
    lastrow = ed.Range("C" & ed.Rows.count).End(xlUp).Row
    
    

    For i = 2 To lastrow
        
        
        If ed.Cells(i, 3) <> ed.Cells(i - 1, 3) Then
        
            input_row = i
        
            If ed.Cells(i, 8) = "Address" Then
                ed.Cells(i, 4) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Email" Then
                ed.Cells(i, 7) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Phone" And ed.Cells(i, 9) = "Home" Then
                ed.Cells(i, 5) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Phone" And ed.Cells(i, 9) = "Cell" Then
                ed.Cells(i, 6) = ed.Cells(i, 10)
            Else
            End If
            
        Else
            
            If ed.Cells(i, 8) = "Address" Then
                ed.Cells(input_row, 4) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Email" Then
                ed.Cells(input_row, 7) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Phone" And ed.Cells(i, 9) = "Home" Then
                ed.Cells(input_row, 5) = ed.Cells(i, 10)
            ElseIf ed.Cells(i, 8) = "Phone" And ed.Cells(i, 9) = "Cell" Then
                ed.Cells(input_row, 6) = ed.Cells(i, 10)
            Else
            End If
            
        End If
    
    Next i

End Sub


'need to replace all the i's with the correct row number

```

