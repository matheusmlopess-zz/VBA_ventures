Sub procuraCaixasDeTexto()
    Dim forma As Shape
    Dim forma1 As Shape
    Dim aux As String
    Dim achei As Integer
    Dim objInLineShape As InlineShape
    Dim eh_WordArt As Shape
    Dim eh_WordArt_Inline As InlineShape
    Dim sTexto As String
    Dim hdr As HeaderFooter
   
' #####################################################
' Handler de imagens
' #####################################################
 
    Dim forma2 As Shape
    
    For Each forma2 In ActiveDocument.Shapes
       forma2.Select
           If forma2.Type = msoPicture Then
               'forma2.Select
               'MsgBox Selection.ShapeRange.Name
               forma2.PictureFormat.Brightness = 0.5
               If Selection.ShapeRange.Name = "Imagem 3" Then
                   forma2.Select
                   forma2.PictureFormat.Brightness = 0.5
               End If
            End If
    Next
 
' #####################################################
' Handler de wordArt
' #####################################################

   
     On Error GoTo ErroHandler
       For Each eh_WordArt In ActiveDocument.Shapes
       On Error GoTo ErroHandler
            sTexto = "no Word Art"
            sTexto = eh_WordArt.TextEffect.Text
                If sText <> "no Word Art" Then
                    eh_WordArt.Select
                    eh_WordArt.Fill.Visible = False
                    eh_WordArt.Fill.Transparency = 1
                    
                    eh_WordArt.Line.Visible = False
                    eh_WordArt.Line.Transparency = 1
          
                End If
        
    Next
        
    For Each eh_WordArt_Inline In ActiveDocument.InlineShapes
    On Error GoTo ErroHandler
           sTexto = "no Word Art"
           sTexto = eh_WordArt_Inline.TextEffect.Text
                If sText <> "no Word Art" Then
                   eh_WordArt_InlineS.Select
                   eh_WordArt_Inline.Fill.Visible = False
                   eh_WordArt_Inline.Fill.Transparency = 1
                   
                   eh_WordArt.Line.Visible = False
                   eh_WordArt_Inline.Line.Transparency = 1
                  
                 End If
        
    Next
  
' #####################################################
' só parte P&B
' #####################################################
 
    pathOf = CreateObject("WScript.Shell").specialfolders("Desktop")
    
        If Application.Documents.Count >= 1 Then
            nomeDoc = ActiveDocument.Name
        Else
            MsgBox "No documents are open"
        End If
    
    nomeDaPasta = pathOf & "\" & nomeDoc & "_"
    CreateFolder (nomeDaPasta)
    
    ActiveDocument.SaveAs FileName:=nomeDaPasta _
    & strNewFolderName & "\" & "Parte_Preto&Branco", _
    FileFormat:=wdFormatDocument


 '########################################################################
    
' #####################################################
' Handler de cabeçalho
' #####################################################
     
     For Each hdr In ActiveDocument.Sections(1).Headers
        hdr.Range.Text = vbNullString
     Next hdr
     

' #####################################################
' Reverter wordArt
' #####################################################

     On Error GoTo ErroHandler
       For Each eh_WordArt In ActiveDocument.Shapes
       On Error GoTo ErroHandler
            sTexto = "no Word Art"
            sTexto = eh_WordArt.TextEffect.Text
                If sText <> "no Word Art" Then
                    eh_WordArt.Select
                    eh_WordArt.Fill.Visible = True
                    eh_WordArt.Fill.Transparency = 0.5
                    eh_WordArt.Line.Transparency = 0.5
          
                End If
        
    Next
        
    For Each eh_WordArt_Inline In ActiveDocument.InlineShapes
    On Error GoTo ErroHandler
           sTexto = "no Word Art"
           sTexto = eh_WordArt_Inline.TextEffect.Text
                If sText <> "no Word Art" Then
                   eh_WordArt_InlineS.Select
                   eh_WordArt_Inline.Fill.Visible = True
                   eh_WordArt_Inline.Fill.Transparency = 0.1
                   eh_WordArt_Inline.Line.Transparency = 0.1
                  
                 End If
     Next

    Dim forma3 As Shape
    
    For Each forma3 In ActiveDocument.Shapes
       forma3.Select
           If forma2.Type = msoPicture Then
               forma3.Select
               forma3.PictureFormat.Brightness = 0.5
               If Selection.ShapeRange.Name = "Imagem 3" Then
                   forma3.Select
                   forma3.PictureFormat.Brightness = 1
               End If
            End If
    Next

    ActiveDocument.Range.Font.Color = wdColorWhite
    For Each formaText In ActiveDocument.Shapes
        
        If formaText.Type = msoTextBox Then
            formaText.Select
            Selection.ShapeRange.Fill.Transparency = 1
        If Selection.ShapeRange.TextFrame.HasText Then
        On Error GoTo Handler1
            Selection.Font.Color = wdColorWhite
        End If
            
         End If
    Next
    

' #####################################################
' só parte Colorida
' #####################################################
 
    ActiveDocument.SaveAs FileName:=nomeDaPasta & strNewFolderName & "\" & "Parte_PretoBranco", FileFormat:=wdFormatDocument



ErroHandler:
    Err.Clear
    Resume Next
 
 
Handler1:
    Err.Clear
    Resume Next
   
Handler2:
    'MsgBox "Erro inner loop codigo (2) !"
    Resume Next
   
End Sub



Function CreateFolder(ByVal sPath As String) As Boolean
'by Patrick Honorez - www.idevlop.com
    Dim fs As Object
    Dim FolderArray
    Dim Folder As String, i As Integer, sShare As String

    If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
    Set fs = CreateObject("Scripting.FileSystemObject")
    'UNC path ? change 3 "\" into 3 "@"
    If sPath Like "\\*\*" Then
        sPath = Replace(sPath, "\", "@", 1, 3)
    End If
    'now split
    FolderArray = Split(sPath, "\")
    'then set back the @ into \ in item 0 of array
    FolderArray(0) = Replace(FolderArray(0), "@", "\", 1, 3)
    On Error GoTo hell
    'start from root to end, creating what needs to be
    For i = 0 To UBound(FolderArray) Step 1
        Folder = Folder & FolderArray(i) & "\"
        If Not fs.FolderExists(Folder) Then
            fs.CreateFolder (Folder)
        End If
    Next
    CreateFolder = True
hell:
End Function
