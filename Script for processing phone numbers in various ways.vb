Sub NumberSort()        

'---------------------------------------------------------'
' Processes telephone numbers in various ways (see below) '
'---------------------------------------------------------'

Dim CountNumber As Integer
Dim CellContents As String
Dim ModContents As String
Dim ColNum As Integer
Dim RowNum As Integer
Dim BottomRow As Integer
Dim LastCol As Integer
Dim Character As String
Dim Simplifier As String
Dim i As Integer
Dim RemoveString As String
Dim CharLocation As Integer
Dim RemStringLen As Integer
Dim CellStringLen As Integer
Dim PhoneStartCol As Integer
Dim HeaderRow As Integer
Dim a As Integer
Dim b As Integer

'----------------------------------------------------------------------------------------'
' This macro takes a sheet contining clients with multiple telephone numbers, and        '
' intelligently processes them in various ways to rewrite them in a specified format.	 '
'																						 '
' I used this code to process a list of more than 4,000 contacts from various sources 	 '
' which had a whole range of different number formats, depending on where they had been	 '
' pulled from.																	         '
'----------------------------------------------------------------------------------------'

'NOTES:
'----------------------------------------------------------------------------------------'                                                        
'Choose the processes to apply by uncommenting the required functions below 		         '
'(line 89 onward)																		 '
'																						 '
'Ensure the following are correct before running:                                        '
'----------------------------------------------------------------------------------------'
  PhoneStartCol = 2    'The leftmost column containing telephone numbers of the contacts '
  HeaderRow = 1        'The bottom row of the table header			   				     ' 
'----------------------------------------------------------------------------------------' 
'This script assumes the following:														 '		     								     
' - Contacts are arranged with telephone numbers in the rightmost columns				 '
' - A userform has been created named "Status", with a text box called "StatusText"		 '
'----------------------------------------------------------------------------------------'


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
    
    
    RowNum = HeaderRow  	 	'Last row of header
    
NextRow:
    ColNum = PhoneStartCol  'The first column in the sheet containing telephone numbers
    
    RowNum = RowNum + 1
    If RowNum > BottomRow Then Goto Ending
    
    With UserForm_Status
    .StatusText = "Progress: " & RowNum & " of " & BottomRow
    .Show vbModeless
    .Repaint
    End With


NextCol:
    CellContents = Cells(RowNum, ColNum)
    
    
'Choose the functions to process by uncommenting below:
    
    'If the number starts with +, change it to "00[space]"
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 1) = "+" Then
    '      ModContents = "00 " & Mid(CellContents, 2, 50)
    '      CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
      
      
    'If the number has text in it, remove (also removes punctuation except + and /)
    '-------------------------------------------------------------------------------
    '  For i = 1 To Len(CellContents)
    '      Character = Mid(CellContents, i, 1)
    '      If Character Like "[0-9, ,/,(,)]" Then
    '          Simplifier = Simplifier & Character
    '      End If
    '  Next i
    '  ModContents = LCase(Simplifier)
    '  Character = ""
    '  Simplifier = ""
    ' CellContents = ModContents
    '-------------------------------------------------------------------------------
    
    
    'Remove specified characters from contents (whatever is in RemoveString)
    '-------------------------------------------------------------------------------
    '  RemoveString = "()"
    '  CharLocation = InStr(CellContents, RemoveString)
    '  If CharLocation <> 0 Then
    '    RemStringLen = Len(RemoveString)
    '    CellStringLen = Len(CellContents)
    '    'a = CharLocation - 1
    '    'b = CharLocation + RemStringLen
    '    ModContents = Mid(CellContents, 1, CharLocation - 1) & Mid(CellContents, CharLocation + RemStringLen, 50)
    '    CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
    
    
    'If there is no space after double 00, add one
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 2) = "00" And Mid(CellContents, 3, 1) <> " " Then
    '      ModContents = Left(CellContents, 2) & " " & Mid(CellContents, 3, 50)
    '      CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
      
      
    'If number is 020X XXX... then change to 020 XXXX XXXX
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 3) = "020" And Left(CellContents, 5) <> " " Then
    '      For i = 1 To Len(CellContents)
    '       Character = Mid(CellContents, i, 1)
    '       If Character Like "[0-9,/]" Then
    '          Simplifier = Simplifier & Character
    '      End If
    '  Next i
    '  CellContents = LCase(Simplifier)
    '  Character = ""
    '  Simplifier = ""
    '  ModContents = Left(CellContents, 3) & " " & Mid(CellContents, 4, 4) & " " & Mid(CellContents, 8, 50)
    '  CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------

    
    'If number is 020 then change to 00 44 (0)20
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 3) = "020" Then
    '      ModContents = "00 44 (0)20 " & Mid(CellContents, 4, 50)
    '      CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
           
           
    'If number is 44 (0) 20 then change to 00 44 (0)20
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 9) = "44 (0) 20" Then
    '      ModContents = "00 44 (0)20 " & Mid(CellContents, 10, 50)
    '      CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
    
      
    'If number is " ", delete it.
    '-------------------------------------------------------------------------------
    '  If CellContents = " " Or CellContents = "  " Or CellContents = "  " Then
    '      Cells(RowNum, ColNum).ClearContents
    '  End If
    '-------------------------------------------------------------------------------
    
    
    'If number doesn't start with 0, add one plus a space
    '-------------------------------------------------------------------------------
    '  If Left(CellContents, 1) <> "0" Then
    '      ModContents = "0 " & CellContents
    '      CellContents = ModContents
    '  End If
    '-------------------------------------------------------------------------------
    
    
    'If number is "0", delete it.
    '-------------------------------------------------------------------------------
    '  If CellContents = "0" Or CellContents = "0 " Then
    '      Cells(RowNum, ColNum).ClearContents
    '  End If
    '-------------------------------------------------------------------------------

    Cells(RowNum, ColNum) = CellContents

    ColNum = ColNum + 1
    If ColNum > LastCol Then Goto NextRow
    Goto NextCol

Ending:

Application.ScreenUpdating = True
UserForm_Status.Hide

End Sub