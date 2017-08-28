Sub find_wordart()

    Dim eh_WordArt As Shape, eh_WordArt_InlineShpNrm As InlineShape, sText As String
    
    On Error GoTo ErroHandler
        For Each eh_WordArt In ActiveDocument.Shapes
            sText = "no Word Art"
            sText = eh_WordArt.TextEffect.Text
                If sText <> "no Word Art" Then
                    eh_WordArt.Select
                    eh_WordArt.Fill.Visible = False
                    eh_WordArt.Fill.Transparency = 1
                    eh_WordArt.Line.Visible = False
                End If
        Next 
        
        For Each eh_WordArt_InlineShpNrm In ActiveDocument.InlineShapes
            sText = "no Word Art"
            sText = eh_WordArt_InlineShpNrm.TextEffect.Text
                If sText <> "no Word Art" Then
                    eh_WordArt_InlineShpNrm.Select
                    
                     eh_WordArt_InlineShpNrm.Fill.Visible = False
                     eh_WordArt_InlineShpNrm.Fill.Transparency = 1
                     eh_WordArt_InlineShpNrm.Line.Visible = False
                 End If
        Next

ErroHandler:
    Err.Clear
    Resume Next

End Sub
