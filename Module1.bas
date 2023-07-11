Attribute VB_Name = "Module1"
Sub Show_Form()
        
        frmForm.Show
    
End Sub

Sub Reset()

    Dim iRow As Long
    
    iRow = [Counta(Database!A:A] ' Identifying the last row
    
    With frmForm
        .txtID.Value = ""
        .txtName.Value = ""
        .optMale.Value = False
        .optfemale.Value = False
        
        .cmbDepartment.Clear
        .cmbDepartment.AddItem "HR"
        .cmbDepartment.AddItem "Operation"
        .cmbDepartment.AddItem "Training"
        .cmbDepartment.AddItem "Quality"
        
        .txtCity.Value = ""
        .txtCountry.Value = ""
        
        .lstDatabase.ColumnCount = 9
        .lstDatabase.ColumnHeads = True
        
        .lstDatabase.ColumnWidths = "30,60,75,40,60,45,55,70,70"
        
        If iRow > 1 Then
            
            .lstDatabase.RowSource = "Database!A2:I" & iRow
        Else
            .lstDatabase.RowSource = "Database!A2:I2"
        
        End If
            
    
    End With


End Sub

Sub Submit()

    Dim sh As Worksheet
    Dim iRow As Long
    
    Set sh = ThisWorkbook.Sheets("Database")
    
    iRow = [Counta(Database!A:A)] + 1
    With sh
    
        .Cells(iRow, 1) = iRow - 1
        .Cells(iRow, 2) = frmForm.txtID.Value
        .Cells(iRow, 3) = frmForm.txtName.Value
        .Cells(iRow, 4) = IIf(frmForm.optfemale.Value = True, "Female", "Male")
        .Cells(iRow, 5) = frmForm.cmbDepartment.Value
        .Cells(iRow, 6) = frmForm.txtCity.Value
        .Cells(iRow, 7) = frmForm.txtCountry.Value
        .Cells(iRow, 8) = Application.UserName
        .Cells(iRow, 9) = [Text(Now(),"DD-MM-YYY HH:MM:SS")]
        
    End With

End Sub

