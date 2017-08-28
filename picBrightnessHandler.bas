 Sub imageBright()

 Dim forma As Shape
 
 
 For Each forma In ActiveDocument.Shapes
     forma.Select
    ' MsgBox Selection.ShapeRange.Name
         
     
        If forma.Type = msoPicture Then
            forma.Select
            forma.PictureFormat.Brightness = 1
        End If
        MsgBox Selection.ShapeRange(1).Name
        If forma.Type = msoTextEffect Then
        MsgBox "aqui"
            forma.Fill.Visible = False
            forma.Fill.Transparency = 1
            forma.Line.Visible = False
        End If
        
 Next

 End Sub
