Attribute VB_Name = "Módulo1"
Sub caixaAuta()

    For Each forma In ActiveDocument.Shapes
    
    forma.Select
    'On Error GoTo Handler2

        If Selection.ShapeRange.TextFrame.HasText Then
    
            Selection.ShapeRange.TextFrame.TextRange.Select
           ' nameOk = Selection.ShapeRange(1).Name
           ' MsgBox nameOk
        End If
        
        
        If forma.Type = msoTextBox Then
            forma.Select
            'MsgBox Selection.ShapeRange(1).Name
            Selection.Font.Size = 12
            Selection.Font.AllCaps = True
            Selection.Range.Case = wdUpperCase
    
        End If
        
    Next

End Sub
