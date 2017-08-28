Sub caixaAuta()
    For Each forma In ActiveDocument.Shapes
         If forma.Type = msoTextBox Then
            forma.Select
            Selection.Font.Size = 12
            Selection.Font.AllCaps = True
            Selection.Range.Case = wdUpperCase
         End If
    Next
End Sub
