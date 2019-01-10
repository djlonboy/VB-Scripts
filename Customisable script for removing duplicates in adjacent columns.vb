Sub RemoveDupsLefttoRight()           
									  '-----------------------------------------------------------------------------------------------------------'
                                      ' Searches for duplicates in adjacent columns and removes, or removes cells with the least data (see below) '
                                      '-----------------------------------------------------------------------------------------------------------'

Application.ScreenUpdating = False

Dim RowNum As Integer
Dim ColNum As Integer
Dim FirstColNum As Integer
Dim BottomRow As Integer
Dim LastCol As Integer
Dim Deleted As Integer
Dim FirstCell As String
Dim ThisCell As String
Dim Character As String
Dim Simplifier As String
Dim i As Integer
Dim n As Integer

    'Find botom row of sheet
    RowNum = 1
    ColNum = 1
    Do
        RowNum = RowNum + 1
    Loop While IsEmpty(Cells(RowNum, ColNum)) = False And IsEmpty(Cells(RowNum + 1, ColNum)) = False
    BottomRow = RowNum
    
    'Find last column of sheet
    RowNum = 1
    ColNum = 1
    Do
        ColNum = ColNum + 1
    Loop While IsEmpty(Cells(RowNum, ColNum)) = False And IsEmpty(Cells(RowNum, ColNum + 1)) = False
    LastCol = ColNum

NextRow:
    RowNum = RowNum + 1
    If RowNum > BottomRow Then Goto Ending
     
	'Update the status box with current progress (requires a userform to be created named "Status").  
    With UserForm_Status
    .StatusText = "Progress: " & RowNum & " of " & BottomRow & vbCrLf & "Deleted Cells: " & Deleted
    .Show vbModeless
    .Repaint
    End With
    
    FirstColNum = 1
    
NextCheck:
    FirstColNum = FirstColNum + 1
    If FirstColNum > LastCol Then Goto NextRow
    
    'Get value from first column
    FirstCell = Cells(RowNum, FirstColNum)
    
    'Simplify (removes spaces and special characters)
    For i = 1 To Len(FirstCell)
        Character = Mid(FirstCell, i, 1)
        If Character Like "[A-Z,a-z,0-9]" Then
            Simplifier = Simplifier & Character
        End If
    Next i
    FirstCell = LCase(Simplifier)
    Character = ""
    Simplifier = ""

    ColNum = FirstColNum
    
NextCell:
    ColNum = ColNum + 1
    
NextCellDeleted:
    If ColNum > LastCol Then Goto NextCheck

    'Get value from next column
    ThisCell = Cells(RowNum, ColNum)
    If ThisCell = "" Then Goto NextCell
    

    'Simplify
    For i = 1 To Len(ThisCell)
        Character = Mid(ThisCell, i, 1)
        If Character Like "[A-Z,a-z,0-9]" Then
            Simplifier = Simplifier & Character
        End If
    Next i
    ThisCell = LCase(Simplifier)
    Character = ""
    Simplifier = ""


'Choose the actions to perform (uncomment)
                                                                              
    'Delete if the same as first cell
    '---------------------------------------------------------
    '    If ThisCell = FirstCell Then
    '        Cells(RowNum, ColNum).Delete Shift:=xlToLeft
    '        Deleted = Deleted + 1
    '        Goto NextCellDeleted
    '    End If
	'---------------------------------------------------------

    'Delete the cell with the least amount of data
    '---------------------------------------------------------
    '    If Len(ThisCell) < Len(FirstCell) Then
    '        Cells(RowNum, ColNum).Delete Shift:=xlToLeft
    '        Deleted = Deleted + 1
    '        GoTo NextCellDeleted
    '    End If
    '    If Len(ThisCell) > Len(FirstCell) Then
    '        Cells(RowNum, FirstColNum).Delete Shift:=xlToLeft
    '        Deleted = Deleted + 1
    '        GoTo NextCellDeleted
    '    End If
	'---------------------------------------------------------

    Goto NextCell

Ending:

Application.ScreenUpdating = True
UserForm_Status.Hide

End Sub