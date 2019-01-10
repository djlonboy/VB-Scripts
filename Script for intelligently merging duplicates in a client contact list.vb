
Sub RemoveDuplicateContacts()
                             
Dim EmailCol As Integer      
Dim PhoneStartCol As Integer 
Dim PhoneCols As Integer     
Dim SubbedCol As Integer     
Dim JobCol As Integer        
Dim RowNum As Integer
Dim ColNum As Integer
Dim BottomRow As Integer
Dim LastCol As Integer
Dim ThisEmail As String
Dim NextEmail As String
Dim Character As String
Dim Simplifier As String
Dim i As Integer
Dim IsFilled As String
Dim ThisFillCount As Integer
Dim NextFillCount As Integer
Dim FavourTop As Boolean
Dim ThisCellContents As String
Dim NextCellContents As String
Dim DupCount As Integer							


'----------------------------------------------------------------------------------------------------'
' Detects duplicates using simplified email field as a start point, then intelligently decides which '
' contact to delete based on a series of rules:                                                      '
'  - Counts the number of filled fields in both and favours the more complete entry                  '
'  - If the numbers are different, copy to the entry to be kept                                      '
'  - If the more complete entry has no subbed/unsubbed info, copy it to the entry to be kept         '
'  - If one entry has subbed and another has unsubbed, unsubbed wins                                 '
'  - If the job title is missing, copy to the entry to be kept                                       '
' 																								     '
' The script assumes the following:																     '
'  - Contacts are arranged with email addresses arranged from A-Z.								     '
'  - Phone numbers are the rightmost columns of the contact list (so that additional columns can be  '
'    added for non-duplicate numbers																	 '
'  - A userform has been created named "Status", with a text box called "StatusText" 				 '
'----------------------------------------------------------------------------------------------------'
'NOTES:                                                        										 '
'Order contacts by email (A-Z) before running 														 '
'Ensure the following are correct before running: 													 '
'----------------------------------------------------------------------------------------------------'
  EmailCol = 6         'The column number containing the contact's email address                     '                 
  PhoneStartCol = 11   'The leftmost column containing telephone numbers of the contacts             '                 
  SubbedCol = 8        'The column containing the subscribed/unsubcribed information of the contacts '
  JobCol = 5           'The column containing the job title of the contacts                          ' 
'----------------------------------------------------------------------------------------------------'
                                                        


Application.ScreenUpdating = False
        
        
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
    
NextRow:	 
	'Move to next row or end the sub if the bottom has been reached.
    ColNum = EmailCol

    RowNum = RowNum + 1
    If Cells(RowNum, 1) = "" And Cells(RowNum, 4) = "" And Cells(RowNum, 6) = "" Then Goto Ending
    
	'Update the status box with current progress 
    With UserForm_Status 
    .StatusText = "Progress: " & RowNum & " of " & BottomRow & vbCrLf & "Duplicate Count: " & DupCount
    .Show vbModeless
    .Repaint
    End With


	'Check for duplicate rows
    ThisEmail = Cells(RowNum, ColNum)
    NextEmail = Cells(RowNum + 1, ColNum)
    
    'Simplify both emails
    For i = 1 To Len(ThisEmail)
        Character = Mid(ThisEmail, i, 1)
        If Character Like "[A-Z,a-z,0-9]" Then
            Simplifier = Simplifier & Character
        End If
    Next i
    ThisEmail = LCase(Simplifier)
    Character = ""
    Simplifier = ""
    
    For i = 1 To Len(NextEmail)
        Character = Mid(NextEmail, i, 1)
        If Character Like "[A-Z,a-z,0-9]" Then
            Simplifier = Simplifier & Character
        End If
    Next i
    NextEmail = LCase(Simplifier)
    Character = ""
    Simplifier = ""

    'If emails are the same and not empty, check if one of the names matches to confirm it's a duplicate
    If ThisEmail = NextEmail And ThisEmail <> "" Then
        
        If Cells(RowNum, 2) = Cells(RowNum + 1, 2) Or Cells(RowNum, 3) = Cells(RowNum + 1, 3) Then 'it's a Duplicate
            Goto ProcessDuplicate
        Else 'it's not a Duplicate
            Goto NextRow
        End If
    
    'If emails are the same, but empty, check if both of the names match to confirm it's a duplicate
    ElseIf ThisEmail = NextEmail And ThisEmail = "" Then
        If Cells(RowNum, 2) = Cells(RowNum + 1, 2) And Cells(RowNum, 3) = Cells(RowNum + 1, 3) Then 'it's a Duplicate
            Goto ProcessDuplicate
        Else 'it's not a Duplicate
            Goto NextRow
        End If
        
    'If there is no match, check the next row
    Else
        Goto NextRow
    End If
    
ProcessDuplicate:
    DupCount = DupCount + 1

    'Count the number of filled fields in both and favour the more complete entry
    ThisFillCount = 0
    NextFillCount = 0
    
    For i = 1 To LastCol
        IsFilled = Cells(RowNum, i)
        If IsFilled <> "" Then ThisFillCount = ThisFillCount + 1
    Next i
    IsFilled = ""

    For i = 1 To LastCol
        IsFilled = Cells(RowNum + 1, i)
        If IsFilled <> "" Then NextFillCount = NextFillCount + 1
    Next i
    IsFilled = ""

    If ThisFillCount >= NextFillCount Then
        FavourTop = True
    End If
    
    If ThisFillCount < NextFillCount Then
        FavourTop = False
    End If
    
	'Consolidate numbers from both rows and remove duplicate numbers

    'See how many phone number rows are in use
    ThisFillCount = 0
    NextFillCount = 0
    
    For i = PhoneStartCol To LastCol
        IsFilled = Cells(RowNum, i)
        If IsFilled <> "" Then ThisFillCount = ThisFillCount + 1
    Next i
    IsFilled = ""
    For i = PhoneStartCol To LastCol
        IsFilled = Cells(RowNum + 1, i)
        If IsFilled <> "" Then NextFillCount = NextFillCount + 1
    Next i
    IsFilled = ""

    'Copy numbers from lower row into top row
    Range(Cells(RowNum + 1, PhoneStartCol), Cells(RowNum + 1, PhoneStartCol + NextFillCount - 1)).Copy
    Cells(RowNum, PhoneStartCol + ThisFillCount).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ColNum = PhoneStartCol - 1
    
    'Take the next number in a row
NextTelCol:
    ColNum = ColNum + 1
    NextColNum = ColNum

Cells(RowNum, ColNum).Select
      
    ThisCellContents = Cells(RowNum, ColNum)

    'When all numbers have been checked, ensure numbers are in the row to be kept, then move on to the next process
    If ThisCellContents = "" Then
        If FavourTop = False Then
                Range(Cells(RowNum, PhoneStartCol), Cells(RowNum, PhoneStartCol + 7)).Copy
                Cells(RowNum + 1, PhoneStartCol).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
        End If
        Goto ProcessSubs
    
    End If
    
    'Take the next number to the right of it
NextCompare:
    NextColNum = NextColNum + 1
    NextCellContents = Cells(RowNum, NextColNum)
    If NextCellContents = "" Then Goto NextTelCol

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
    
    'Check for duplicate and remove the second match if so
    If ThisCellContents = NextCellContents Then
        Cells(RowNum, NextColNum).Delete Shift:=xlToLeft
        NextColNum = NextColNum - 1
    End If
    
    Goto NextCompare

ProcessSubs:
    
    'If the more complete entry has no subbed/unsubbed info, copy it to the entry to be kept
    If FavourTop = True And IsEmpty(Cells(RowNum, SubbedCol)) = True Then
        Cells(RowNum, SubbedCol) = Cells(RowNum + 1, SubbedCol)
    End If

    'If one entry has subbed and another has unsubbed, unsubbed wins
    If Cells(RowNum + 1, SubbedCol) = "Unsubbed" Then
        Range(Cells(RowNum + 1, SubbedCol), Cells(RowNum + 1, SubbedCol + 2)).Copy
        Cells(RowNum, SubbedCol).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    'If the job title is missing, copy to entry to be kept
    If IsEmpty(Cells(RowNum, JobCol)) = True Then
        Cells(RowNum, JobCol) = Cells(RowNum + 1, JobCol)
    End If
    
    If IsEmpty(Cells(RowNum + 1, JobCol)) = True Then
        Cells(RowNum + 1, JobCol) = Cells(RowNum, JobCol)
    End If


    'Delete the favoured row
    If FavourTop = True Then
        Rows(RowNum + 1).EntireRow.Delete
    End If
    
    If FavourTop = False Then
        Rows(RowNum).EntireRow.Delete
    End If
    
    If (RowNum + 1) >= BottomRow Then Goto Ending
    RowNum = RowNum - 1
    Goto NextRow

Ending:

Application.CutCopyMode = False
    UserForm_Status.Hide
    Application.ScreenUpdating = True

End Sub
