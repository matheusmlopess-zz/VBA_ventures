Sub procuraCaixasDeTexto()
    Dim forma As Shape
    Dim aux As String
    Dim achei As Integer
    Dim objInLineShape As InlineShape

   'Exclui bordas do documento
     ActiveDocument.Range.Font.Color = wdColorWhite
    
   'Text color Automatic to White
    ActiveDocument.Sections(1).Borders.Enable = False

    For Each forma In ActiveDocument.Shapes
    forma.Select
    'On Error GoTo Handler2

        If Selection.ShapeRange.TextFrame.HasText Then
        On Error GoTo Handler1
            Selection.ShapeRange.TextFrame.TextRange.Select
           ' nameOk = Selection.ShapeRange(1).Name
           ' MsgBox nameOk
        End If

        If forma.Type = msoTextBox Then
            forma.Select
            Selection.ShapeRange.TextFrame.TextRange.Select
            aux = Selection.Text    ' aux = Left(aux, 20) para 20 caracteres
           achei = MsgBox("[" + Selection.ShapeRange(1).Name + "]:" _
            & vbCrLf _
            & vbCrLf _
            & aux _
            & vbCrLf _
            & "Parar aqui?", _
            vbYesNo, "Caixa de texto encontrada")
                
            If achei = vbYes Then Exit For
                
      '  For Each objInLineShape In InlineShapes
      '  objInLineShape.Select
      '  Selection.ShapeRange(1).Name
      '  Selection.Font.Color = wdColorRed
      '  MsgBox "dasdasdasfasfsafsaasfsaf"
      '  Selection.ShapeRange.TextFrame.TextRange.Select
      '  Next objInLineShape

            End If 
    Next

Handler1:
  ' MsgBox "Tudo Pronto codigo (1) !"
   
'Handler2:
   'MsgBox "Erro inner loop codigo (2) !"
   
End Sub

