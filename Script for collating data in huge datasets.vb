Sub FindAndPasteV2()

Application.ScreenUpdating = False

Dim RowNumCombined As Integer
Dim ColNumCombined As Integer
Dim RowNumCopy As Integer
Dim ColNumCopy As Integer
Dim BottomRowCopy As Integer
Dim BottomRowCombined As Integer
Dim LastColCopy As Integer
Dim LastColCombined As Integer
Dim SheetNum As Integer
Dim CombinedSheet As Integer
Dim CopySheet As Integer
Dim ClientName As String
Dim ClientLocation As Range
Dim ClientData As Variant
Dim HeaderData As Variant
Dim NotFound As Integer

'----------------------------------------------------------------------------------------------------'
' This macro takes data from a sheet (the "CopySheet") and pastes it into a combined   			     '
' sheet (the "CombinedSheet"). It works row by row and takes the name in the leftmost   			     '
' column of the CopySheet and tries to find a matching name in the CombinedSheet. If it 			     '
' finds one, it pastes the corresponding data from the CopySheet on the CombinedSheet, 			     '
' after any data already on the CombinedSheet. If it can't find the name, a warning is   			 '
' displayed and the data is is pasted into the CombinedSheet on a new row, at the       			     '
' bottom. Finally, it copies the header row from the CopySheet and pastes it in the     			     '
' CombinedSheet, in line with the data.												    			 '
'																									 '
' I used this code to collate data from sets of client lists from different sources, 				 '
' most of which had over 4,000 entries in them and so were impossible to combine						 '
' manually. I then used additional code to intellegently scan the combined list and					 '
' remove duplicate data.																	  			 '
'----------------------------------------------------------------------------------------------------'
' The script assumes the following:																     '
'  - Contacts are arranged with contact names in the leftmost column, arranged from A-Z.			     '
'  - A userform has been created named "Status", with a text box called "StatusText" 				 '
'----------------------------------------------------------------------------------------------------'

'-----------------------------------------------------------------' 
'                                                                 ' 
   CopySheet = 8 '      ! Check this is correct before running !  '
   CombinedSheet = 9 '  ! Check this is correct before running !  '
'                                                                 '
'-----------------------------------------------------------------'

NotFound = 0

Worksheets(CopySheet).Select
    
    'Find botom row of copy sheet
    RowNumCopy = 1
    ColNumCopy = 1
    Do
        RowNumCopy = RowNumCopy + 1
    Loop While IsEmpty(Cells(RowNumCopy, ColNumCopy)) = False And IsEmpty(Cells(RowNumCopy + 1, ColNumCopy)) = False
    BottomRowCopy = RowNumCopy
    
    'Find last column of copy sheet
    RowNumCopy = 1
    ColNumCopy = 1
    Do
        ColNumCopy = ColNumCopy + 1
    Loop While IsEmpty(Cells(RowNumCopy, ColNumCopy)) = False And IsEmpty(Cells(RowNumCopy, ColNumCopy + 1)) = False
    LastColCopy = ColNumCopy
    
    'MsgBox ("Bottom Row is: " & BottomRowCopy & vbCrLf & "Last Column is: " & LastColCopy)
    
    
Worksheets(CombinedSheet).Select

    'Find botom row of combined sheet
    RowNumCombined = 1
    ColNumCombined = 1
    Do
        RowNumCombined = RowNumCombined + 1
    Loop While IsEmpty(Cells(RowNumCombined, ColNumCombined)) = False And IsEmpty(Cells(RowNumCombined + 1, ColNumCombined)) = False
    BottomRowCombined = RowNumCombined
    
    'Find last column of combined sheet
    RowNumCombined = 1
    ColNumCombined = 1
    Do
        ColNumCombined = ColNumCombined + 1
    Loop While IsEmpty(Cells(RowNumCombined, ColNumCombined)) = False And IsEmpty(Cells(RowNumCombined, ColNumCombined + 1)) = False
    LastColCombined = ColNumCombined
    
    'MsgBox ("Bottom Row is: " & BottomRowCombined & vbCrLf & "Last Column is: " & LastColCombined)

    RowNumCopy = 1
    ColNumCopy = 1


CheckNext:

Worksheets(CopySheet).Select

    RowNumCopy = RowNumCopy + 1
    If RowNumCopy > BottomRowCopy Then Goto CopyHeaders
    
    With UserForm_Status
    .StatusText = "Progress: " & RowNumCopy & " of " & BottomRowCopy & vbCrLf & "Not found count: " & NotFound
    .Show vbModeless
    .Repaint
    End With
    
    'Find the next client name on the copy sheet
    ClientName = Cells(RowNumCopy, ColNumCopy)
    
    'MsgBox ("'" & ClientName & "'")
    
    'Copy the associated data on that row
    ClientData = Range(Cells(RowNumCopy, 2), Cells(RowNumCopy, LastColCopy))

Worksheets(CombinedSheet).Select

    'Search for ClientName in combined sheet
    Worksheets(CombinedSheet).Select
    Set ClientLocation = Range(Cells(1, 1), Cells(BottomRowCombined, 1)).Find(ClientName) 'default is not case sensitive
    
    'If there is no result found, goto NotFound, otherwise skip it
    If ClientLocation Is Nothing Then
        Goto NotFound
    Else
        Goto Found
    End If
    
NotFound:
    NotFound = NotFound + 1
    
    'Add the client as an additional row at the bottom of the combined sheet
    BottomRowCombined = BottomRowCombined + 1
    Cells(BottomRowCombined, 1) = ClientName
    Range(Cells(BottomRowCombined, LastColCombined + 1), Cells(BottomRowCombined, LastColCopy + LastColCombined - 1)) = ClientData
    
    'MsgBox ("Check Next?")
    Goto CheckNext
    
Found:
    'Paste the client data at the end of the existing data on the correct row
    'MsgBox ("Found on row: " & ClientLocation.Row)
    Range(Cells(ClientLocation.Row, LastColCombined + 1), Cells(ClientLocation.Row, LastColCopy + LastColCombined - 1)) = ClientData
    
   'MsgBox ("Check Next?")
    Goto CheckNext
        
    End  'added for redundancy in case the program finds itsef here for some reason.
    
    
CopyHeaders: 'Add the headers for the data from the copy sheet

    RowNumCopy = 1
    ColNumCopy = 1

Worksheets(CopySheet).Select
    HeaderData = Range(Cells(1, 2), Cells(1, LastColCopy))
Worksheets(CombinedSheet).Select
    Range(Cells(1, LastColCombined + 1), Cells(1, LastColCopy + LastColCombined - 1)) = HeaderData
    
    
UserForm_Status.Hide
Application.ScreenUpdating = False / True
End Sub


