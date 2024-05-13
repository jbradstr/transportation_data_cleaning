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

For this next part I used ChaptGPT to help me seperate names from a fullname column to a firstname and last name column.  The range of this code was edited to fit my own.

```vbscript

Sub SeparateNames()
    
    Dim rng As Range
    Dim cell As Range
    Dim fullName As String
    Dim firstName As String
    Dim lastName As String
    Dim spacePos As Integer
    
    ' Set the range where the names are located
    Set rng = Range("B1:B1000") ' Adjust the range as needed
    
    ' Loop through each cell in the range
    For Each cell In rng
        ' Check if the cell is not empty
        If Not IsEmpty(cell) Then
            fullName = Trim(cell.Value) ' Remove leading and trailing spaces
            ' Find the position of the first space
            spacePos = InStr(fullName, " ")
            If spacePos > 0 Then
                ' Separate first name and last name
                firstName = Left(fullName, spacePos - 1)
                lastName = Mid(fullName, spacePos + 1)
                
                ' Output the separated names to adjacent cells
                cell.Offset(0, 1).Value = firstName
                cell.Offset(0, 2).Value = lastName
            Else
                ' If no space found, assume the whole content as the first name
                cell.Offset(0, 1).Value = fullName
            End If
        End If
    Next cell
End Sub


```
