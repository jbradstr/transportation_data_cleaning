# transportation_data_cleaning
I am currently data cleaning multiple systems of data from the transportation department to turn them into one system that they can utilize.  They are currently inputting multiple pieces of data into different systems and doubling up their workload.  The current project is to take one column that has addresses, emails, and phone numbers in it and split it into 3 seperate columns.

```vbscript

Public Sub addr_phone_email()

    Dim ed As Worksheet: Set ed = ThisWorkbook.Worksheets("editable")
    lastrow = ed.Range("G" & ed.Rows.Count).End(xlUp).Row

    
    For i = 2 To lastrow
        
                
                
            Select Case ed.Cells(i, 7)
                Case "Address"
                    ed.Cells(i, 4) = ed.Cells(i, 9)
                Case "Email"
                    ed.Cells(i, 6) = ed.Cells(i, 9)
                Case "Phone"
                    ed.Cells(i, 5) = ed.Cells(i, 9)
            End Select
    
    Next i

End Sub


input_row = ActiveCell.Row


'need to replace all the i's with the correct row number

```

