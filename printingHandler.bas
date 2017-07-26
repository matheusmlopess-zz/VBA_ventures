Attribute VB_Name = "Módulo1"
Sub procuraCaixasDeTexto()
    Dim forma As Shape
    Dim aux As String
    Dim achei As Integer
    Dim objInLineShape As InlineShape
      

    
   
    'Text color Automatic to White
      'ActiveDocument.Range.Font.Color = wdColorWhite
       ActiveDocument.Range.Font.Color = wdColorRed
    
    'Exclui bordas do documento
      ActiveDocument.Sections(1).Borders.Enable = True

     For Each forma In ActiveDocument.Shapes
     forma.Select

        If Selection.ShapeRange.TextFrame.HasText Then
        On Error GoTo Handler1
            'Selection.Font.Color = wdColorWhite
             Selection.Font.Color = wdColorRed
        End If
        
        'If forma.Type = msoTextBox Then
            'forma.Select
            'forma.PictureFormat.Brightness = 0
            
            
          '  forma.Line.Visible = False
            
            
            
          '  Selection.Borders.Enable = False
          '  Selection.ShapeRange.Fill.BackColor = wdColorWhite
           ' Selection.ShapeRange.Fill.ForeColor = wdColorWhite
           ' Selection.ShapeRange.TextFrame.TextRange.Select
            
        'Not that usefull ... just for the sake o clra
        '    aux = Selection.Text    ' aux = Left(aux, 20) para 20 caracteres
           ' achei = MsgBox("[" + Selection.ShapeRange(1).Name + "]:" _
                & vbCrLf _
                & vbCrLf _
                & aux _
                & vbCrLf _
                & "Parar aqui?", _
                  vbYesNo, "Caixa de texto encontrada")
                
      ' If achei = vbYes Then Exit For
       
        'End If
        
    Next
    
    
Handler1:
   'MsgBox "Tudo Pronto codigo (1) !"
   
'Handler2:
   'MsgBox "Erro inner loop codigo (2) !"
   
End Sub



