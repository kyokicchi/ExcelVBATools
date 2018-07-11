
Sub gatherSheets()

Dim rc As Integer
rc = MsgBox("continue to gather sheets?", vbYesNo + vbQuestion, "confirmation")


If rc = vbYes Then

Range("D10:D500").ClearContents

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
    
    
Dim thisBook, thisSheet, openedBook As String
thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name

Dim sh As Worksheet
Dim sheetFound As Boolean

Dim recRow As Integer
recRow = 10

Dim folderPath, filePath, fullPath, sheetName, nameNewSheet As String
folderPath = Cells(2, 4)    'put all target excel books in a folder and specify here
sheetName = Cells(4, 4)     'input name of the sheet to collect from all books. (all books need to have same sheet)

filePath = Dir(folderPath & "/*.*")
fullPath = folderPath & "\" & filePath

Do While filePath <> ""
    
    Application.StatusBar = "working on: " & filePath

    fullPath = folderPath & "\" & filePath
    
    Workbooks.Open fileName:=fullPath, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True
    openedBook = ActiveWorkbook.Name
    
    sheetFound = False

    For Each sh In Workbooks(openedBook).Sheets
        If sh.Name = sheetName Then sheetFound = True
    Next sh
    
    
    If sheetFound = True Then
       
    Workbooks(thisBook).Activate
    Sheets(thisSheet).Activate
    Cells(recRow, 4) = filePath
    nameNewSheet = Cells(recRow, 3)
    
    With Worksheets.Add()
        .Name = nameNewSheet
    End With
    
    Workbooks(openedBook).Sheets(sheetName).Activate
    Cells.Select
    Selection.Copy
    
    Workbooks(thisBook).Activate
    Sheets(nameNewSheet).Activate
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
'   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'       SkipBlanks:=False, Transpose:=False
    
    Sheets(thisSheet).Activate
    recRow = recRow + 1
    
    End If
    
    Workbooks(filePath).Close savechanges:=0
    
    filePath = Dir()
    
Loop



Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = "completed"
MsgBox ("current value-pull completed")

Else
    MsgBox "aborted"
End If



End Sub
