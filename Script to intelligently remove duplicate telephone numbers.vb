Sub IntelligentNumberCleaner()        '-----------------------------------------------------------------'
                                      ' Intelligently checks for duplicate tel numbers and deletes them '
                                      '-----------------------------------------------------------------'
    Dim ColNum As Integer
    Dim RowNum As Integer
    Dim NextColNum As Integer
    Dim ThisCellContents As String
    Dim NextCellContents As String
    Dim ModContents As String
    Dim ThisCellCompare As String
    Dim NextCellCompare As String
    Dim LastCol As Integer
    Dim BottomRow As Integer
    Dim DeleteNum As Integer
    
'----------------------------------------------------------------------------------------'
' This macro takes a sheet contining clients with multiple telephone numbers, and        '
' intelligently checks them for duplicates, removing the duplicate if found.		         '
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------' 
'This script assumes the following:														 '		     								     
' - Contacts are arranged with a contact name in the leftmost column, followed by a      '
'	series of telephone numbers in the columns to the right.								 '
' - A userform has been created named "Status", with a text box called "StatusText"		 '
'----------------------------------------------------------------------------------------'


Application.ScreenUpdating = False 
	
    DeleteNum = 0

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

    RowNum = 1

    'Take the next Row
NextRow:
    ColNum = 1
    RowNum = RowNum + 1
    If RowNum > BottomRow Then Goto Ending

    With UserForm_Status
    .StatusText = "Progress: " & RowNum & " of " & BottomRow & vbCrLf & "Deleted Cells: " & DeleteNum
    .Show vbModeless
    .Repaint
    End With
    
    'Take the next number in a row
NextCol:
    ColNum = ColNum + 1
    NextColNum = ColNum
    ThisCellContents = Cells(RowNum, ColNum)

    'When all numbers have been checked, move on to the next row
    If ThisCellContents = "" Then Goto NextRow
    
    'Take the next number to the right of it
NextCompare:
    NextColNum = NextColNum + 1
    NextCellContents = Cells(RowNum, NextColNum)
    If NextCellContents = "" Then Goto NextCol
    
    'Apply the following rules:
    
    'Remove any zeros from numbers being compared
    ThisCellContents = Replace(ThisCellContents, "0", "")
    NextCellContents = Replace(NextCellContents, "0", "")
    
    'Remove spaces from the numbers being compared
    ThisCellContents = Replace(ThisCellContents, " ", "")
    NextCellContents = Replace(NextCellContents, " ", "")
    
    'Remove dashes from the numbers being compared
    ThisCellContents = Replace(ThisCellContents, "-", "")
    NextCellContents = Replace(NextCellContents, "-", "")
    
    'Remove brackets from numbers being compared
    ThisCellContents = Replace(ThisCellContents, "(", "")
    NextCellContents = Replace(NextCellContents, "(", "")
    ThisCellContents = Replace(ThisCellContents, ")", "")
    NextCellContents = Replace(NextCellContents, ")", "")
    
    'Remove slashes from numbers being compared
    ThisCellContents = Replace(ThisCellContents, "/", "")
    NextCellContents = Replace(NextCellContents, "/", "")
    
    'Remove the last match
    If ThisCellContents = NextCellContents Then
        Cells(RowNum, NextColNum).Delete Shift:=xlToLeft
        DeleteNum = DeleteNum + 1
    End If
    
    Goto NextCompare

Ending:

    Application.StatusBar = False
    UserForm_Status.Hide

Application.ScreenUpdating = True 
End Sub