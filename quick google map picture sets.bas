




Sub changeMap()

'set 3 cells for input, name them address_input, size_1, size_2

'insert 2 rectangle shapes and name them map 1, map 2

Dim Target, pict_1, pict_2 As String
Dim zoom_1, zoom_2 As Single

Target = Range("address_input").Value
zoom_1 = Range("size_1").Value
zoom_2 = Range("size_2").Value


pict_1 = "http://maps.googleapis.com/maps/api/staticmap?size=512x512&center=" & Target & "&zoom=" & zoom_1
pict_2 = "http://maps.googleapis.com/maps/api/staticmap?size=512x512&center=" & Target & "&zoom=" & zoom_2


ActiveSheet.Shapes.Range(Array("map 1")).Select
Selection.ShapeRange.Fill.UserPicture pict_1

ActiveSheet.Shapes.Range(Array("map 2")).Select
Selection.ShapeRange.Fill.UserPicture pict_2

ActiveSheet.Cells(4, 8).Select

End Sub
