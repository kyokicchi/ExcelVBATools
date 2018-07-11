Sub fileNamePull()

    Dim filePath As String
    Dim fileType As String
    Dim fullPath As String
    Dim cnt As Integer
    
    filePath = Cells(2, 4)  'file path input cell eg. " c:\ "
    fileType = Cells(3, 4)  'file path input cell eg. " *.* "
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


Sub fileNameChange()      'input new names in column F then run.
    
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
