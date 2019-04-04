

Sub showAllSheets()

Dim numSheets As Integer
numSheets = Worksheets.Count

Dim i As Integer

For i = 1 To numSheets

Sheets(i).Visible = True

Next i

MsgBox ("done")

End Sub



Sub sortSheets()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

thisSheetName = ActiveWorkbook.ActiveSheet.Name


Dim i As Integer
i = 1
Dim sheetName As String

If Cells(4, 9) <> "" Then
    sheetName = Cells(4, 9)
    Worksheets(sheetName).Move Before:=Worksheets(1)
End If

Sheets(thisSheetName).Activate
Do While Cells(4 + i, 9) <> ""
       sheetName = Cells(4 + i, 9)
       Worksheets(sheetName).Move After:=Worksheets(i)

i = i + 1
Sheets(thisSheetName).Activate
Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
 
MsgBox ("done")

End Sub



Sub pullSheetNames()
Range(Cells(4, 4), Cells(500, 6)).ClearContents
Range(Cells(4, 9), Cells(500, 9)).ClearContents

Dim i, numSheets As Integer

numSheets = Worksheets.Count

For i = 1 To numSheets

    Cells(i + 3, 4) = Sheets(i).Name
    Cells(i + 3, 6) = Sheets(i).Name
    Cells(i + 3, 9) = Sheets(i).Name


Next i



End Sub



Sub changeSheetNames()

Dim i As Integer
Dim oldName As String
Dim newName As String

i = 1

oldName = Cells(i + 3, 4)

Do While oldName <> ""

newName = Cells(i + 3, 6)
Sheets(oldName).Name = newName

i = i + 1
oldName = Cells(i + 3, 4)

Loop


End Sub


Sub fileNamePull()

    Dim filePath As String
    Dim fileType As String
    Dim fullPath As String
    Dim cnt As Integer
    
    filePath = Cells(2, 4)
    fileType = Cells(3, 4)
    fullPath = Dir(filePath & "/" & fileType)
    cnt = 10
    
    Do While fullPath <> ""
        Cells(cnt, 3) = "file"
        Cells(cnt, 4) = fullPath
        Cells(cnt, 6) = fullPath
        cnt = cnt + 1
        fullPath = Dir()

    Loop


Dim folPath As String
folPath = Dir(filePath & "/", vbDirectory)

    Do While folPath <> ""
    On Error Resume Next
    If GetAttr(filePath & "/" & folPath) And vbDirectory Then
    If folPath <> "." And folPath <> ".." Then
        Cells(cnt, 3) = "folder"
        Cells(cnt, 4) = folPath
        Cells(cnt, 6) = folPath
        cnt = cnt + 1
    End If
    End If
    folPath = Dir()
    Loop


MsgBox ("done!")


End Sub


Sub fileNameChange()
    
    Dim filePath As String
    Dim oldName, oldFullName As String
    Dim newName, newFullName As String
    Dim cnt As Integer

    filePath = Cells(2, 4)
    cnt = Cells(4, 4) + 10

For i = 10 To cnt
    oldName = Cells(i, 4)
    newName = Cells(i, 6)
    If oldName <> "" And newName <> "" Then
      oldFullName = filePath & "/" & oldName
      newFullName = filePath & "/" & newName
      Name oldFullName As newFullName
    End If
Next i

MsgBox ("done!")

End Sub



Sub createFolders()

  Dim createThis As String

  Dim parentFolder As String
  parantFolder = Cells(2, 3)
   
  Dim subFolder As String
  Dim i As Integer
  i = 4
  subFolder = Cells(i, 3)
  
  Do While subFolder <> ""

  createThis = parantFolder & "\" & subFolder

  If Dir(createThis, vbDirectory) <> "" Then
    MsgBox "Folder already exists"
  Else
    MkDir createThis
  End If
  
  i = i + 1
  subFolder = Cells(i, 3)
  Loop
 
  Shell _
    pathName:="explorer" & Chr(32) & parantFolder, _
    WindowStyle:=vbNormalFocus
  
End Sub




Sub pullValue()

    Dim rc As Integer
    rc = MsgBox("continue to value-pull?", vbYesNo + vbQuestion, "confirmation")
    If rc = vbYes Then
        MsgBox "value-pull starting"

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Dim pathName, fName, sheetName(100), cellName(100), curValue(100), chgValue(100), opnThis As String
Dim thisBook, thisSheet, setPass As String

pathName = Cells(3, 2)
setPass = Cells(3, 3)

thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name

Dim startRow, startCol, i, j As Integer

startRow = 3
startCol = 6

i = 0

Do While Cells(startRow, startCol + i) <> ""


fName = Cells(startRow, startCol + i)

j = 0
Do While Cells(startRow + 2 + (5 * j), startCol + i) <> ""
    sheetName(j) = Cells(startRow + 2 + (5 * j), startCol + i)
    cellName(j) = Cells(startRow + 3 + (5 * j), startCol + i)
    j = j + 1
Loop

Application.StatusBar = "processing: " & fName

opnThis = pathName & "\" & fName

If Dir(opnThis) <> "" Then

Workbooks.Open fileName:=opnThis, Password:=setPass, ReadOnly:=True, UpdateLinks:=0, CorruptLoad:=xlRepairFile, IgnoreReadOnlyRecommended:=True


j = 0
Do While sheetName(j) <> ""
        Workbooks(fName).Sheets(sheetName(j)).Activate
'value
        curValue(j) = Cells.Range(cellName(j)).Value
'formula
'        curValue(j) = Cells.Range(cellName(j)).Formula
        
j = j + 1

Loop

Workbooks(fName).Close savechanges:=0
               
End If

Workbooks(thisBook).Activate
Sheets(thisSheet).Activate

j = 0
Do While sheetName(j) <> ""
Cells(startRow + 4 + (5 * j), startCol + i) = curValue(j)
Cells(startRow + 5 + (5 * j), startCol + i) = curValue(j)
j = j + 1
Loop


i = i + 1

Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = "completed"
MsgBox ("current value-pull completed")


    Else
        MsgBox "value-pull aborted"
    End If

End Sub



Sub chgValue()
    Dim rc As Integer
    rc = MsgBox("continue to override value?", vbYesNo + vbQuestion, "confirmation")
    If rc = vbYes Then
        MsgBox "override starting"

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Dim pathName, fileName, sheetName(100), cellName(100), curValue(100), chgValue(100), opnThis As String
Dim thisBook, thisSheet As String

pathName = Cells(3, 2)


thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name

Dim startRow, startCol, i, j As Integer

startRow = 3
startCol = 6
i = 0



Do While Cells(startRow, startCol + i) <> ""


fileName = Cells(startRow, startCol + i)

j = 0
Do While Cells(startRow + 2 + (5 * j), startCol + i) <> ""
    sheetName(j) = Cells(startRow + 2 + (5 * j), startCol + i)
    cellName(j) = Cells(startRow + 3 + (5 * j), startCol + i)
    chgValue(j) = Cells(startRow + 5 + (5 * j), startCol + i).Formula
    j = j + 1
Loop

Application.StatusBar = "processing: " & fileName

opnThis = pathName & "\" & fileName

If Dir(opnThis) <> "" Then

Workbooks.Open fileName:=opnThis, UpdateLinks:=1, IgnoreReadOnlyRecommended:=True

j = 0
Do While sheetName(j) <> ""
        ActiveWorkbook.Sheets(sheetName(j)).Activate
        Cells.Range(cellName(j)).Formula = chgValue(j)
        

j = j + 1
Loop

Workbooks(fileName).Close savechanges:=1
               
End If

Workbooks(thisBook).Activate
Sheets(thisSheet).Activate


i = i + 1

Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = "completed"
MsgBox ("override completed")


    Else
        MsgBox "override aborted"
    End If
End Sub


Sub gatherSheets()
On Error Resume Next
Dim rc As Integer
rc = MsgBox("continue to gather sheets?", vbYesNo + vbQuestion, "confirmation")


If rc = vbYes Then

Range("D10:D500").ClearContents

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
    
    
Dim thisBook, thisSheet, openedBook As String
thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name



Dim recRow As Integer
recRow = 10

Dim folderPath, filePath, fullPath, sheetName, nameNewSheet, filePW As String
folderPath = Cells(2, 4)
filePath = Dir(folderPath & "/*.*")
fullPath = folderPath & "\" & filePath
sheetName = Cells(4, 4)
filePW = Cells(6, 4)


Do While filePath <> ""
    
    Application.StatusBar = "working on: " & filePath

    fullPath = folderPath & "\" & filePath
    
    Workbooks.Open fileName:=fullPath, ReadOnly:=True, UpdateLinks:=False, Password:=filePW
    openedBook = ActiveWorkbook.Name
    

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



Sub pullRmtSheetNames()

    Dim rc As Integer
    rc = MsgBox("continue to pull Sheet Names?", vbYesNo + vbQuestion, "confirmation")
    If rc = vbYes Then
        MsgBox "Name pull starting"

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Dim pathName, fileName, opnThis As String
Dim curName() As String

Dim thisBook, thisSheet As String

pathName = Cells(3, 2)

thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name

Dim startRow, startCol, i, j, x, numSheets As Integer

startRow = 6
startCol = 5

i = 0

Do While Cells(startRow + i, 2) <> ""

fileName = Cells(startRow + i, 2)
Application.StatusBar = "processing: " & fileName
opnThis = pathName & "\" & fileName

If Dir(opnThis) <> "" Then

Workbooks.Open fileName:=opnThis, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True

numSheets = Workbooks(fileName).Sheets.Count
ReDim curName(numSheets)

For x = 1 To numSheets

curName(x) = Sheets(x).Name

Next x

Workbooks(fileName).Close savechanges:=0
               

Workbooks(thisBook).Activate
Sheets(thisSheet).Activate

For j = 1 To numSheets

Cells(startRow + i, startCol + 3 * (j - 1)) = curName(j)
Cells(startRow + i, startCol + 3 * (j - 1) + 1) = curName(j)

Next j


End If

i = i + 1

Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = "completed"
MsgBox ("sheets name pull completed")

    Else
        MsgBox "sheets name pull aborted"
    End If

End Sub


Sub chgRmtSheetNames()

    Dim rc As Integer
    rc = MsgBox("continue to change Sheet Names?", vbYesNo + vbQuestion, "confirmation")
    If rc = vbYes Then
        MsgBox "Name change starting"

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Dim pathName, fileName, curName(100), newName(100), opnThis As String
Dim thisBook, thisSheet As String

Dim linkToDo As Boolean

pathName = Cells(3, 2)

linkToDo = False

If Cells(2, 12) = "yes" Then
linkToDo = True
End If


thisBook = ActiveWorkbook.Name
thisSheet = ActiveSheet.Name

Dim startRow, startCol, i, j As Integer

startRow = 6
startCol = 5

i = 0

Do While Cells(startRow + i, 2) <> ""

fileName = Cells(startRow + i, 2)

j = 1

Do While Cells(startRow + i, 6 + ((j - 1) * 3)) <> ""

curName(j) = Cells(startRow + i, 5 + ((j - 1) * 3))
newName(j) = Cells(startRow + i, 6 + ((j - 1) * 3))

j = j + 1
Loop


Application.StatusBar = "processing: " & fileName
opnThis = pathName & "\" & fileName

If Dir(opnThis) <> "" Then

Workbooks.Open fileName:=opnThis, UpdateLinks:=linkToDo, IgnoreReadOnlyRecommended:=True

For j = 1 To Workbooks(fileName).Sheets.Count

Sheets(curName(j)).Name = newName(j)

Next j

Workbooks(fileName).Close savechanges:=True
               
End If

Workbooks(thisBook).Activate
Sheets(thisSheet).Activate


i = i + 1

Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = "completed"
MsgBox ("sheets name change completed")


    Else
        MsgBox "sheets name change aborted"
    End If

End Sub



