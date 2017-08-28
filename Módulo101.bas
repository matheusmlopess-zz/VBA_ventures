Sub blackNwhite()
    Dim forma, forma2 As Shape
    Dim hdr As HeaderFooter
    
    For Each forma In ActiveDocument.Shapes
       forma.Select
           If forma.Type = msoPicture Then
               forma.Select
                If Selection.ShapeRange.Name <> "Imagem 3" Then
                 forma.PictureFormat.Brightness = 1
                End If
                If Selection.ShapeRange.Name = "Imagem 3" Then
                    forma.PictureFormat.Brightness = 0.5
                End If
            End If
    Next
    
' só parte P&B
     pathOf = CreateObject("WScript.Shell").specialfolders("Desktop")
    
        If Application.Documents.Count >= 1 Then
            nomeDoc = ActiveDocument.Name
        Else
            MsgBox " Try again mate! "
        End If
    
    nomeDaPasta = pathOf & "\" & nomeDoc & "_"
    CreateFolder (nomeDaPasta)
    
    ActiveDocument.SaveAs FileName:=nomeDaPasta _
    & strNewFolderName & "\" & "Parte_Preto&Branco", _
    FileFormat:=wdFormatDocument

 ' Handler de cabeçalho
     For Each hdr In ActiveDocument.Sections(1).Headers
        hdr.Range.Text = vbNullString
     Next hdr
    
    
' Handler de imagens
    Dim forma2 As Shape
       For Each forma2 In ActiveDocument.Shapes
       forma2.Select
           If forma2.Type = msoPicture Then
               forma2.Select
                If Selection.ShapeRange.Name <> "Imagem 3" Then
                     forma2.PictureFormat.Brightness = 0.5
                End If
                If Selection.ShapeRange.Name = "Imagem 3" Then
                     forma2.PictureFormat.Brightness = 1
                End If         
            End If
    Next

    ActiveDocument.Range.Font.Color = wdColorWhite
    For Each formaText In ActiveDocument.Shapes
        
        If formaText.Type = msoTextBox Then
            formaText.Select
            Selection.ShapeRange.Fill.Transparency = 1
        If Selection.ShapeRange.TextFrame.HasText Then
        On Error GoTo ErroHandler
            Selection.Font.Color = wdColorWhite
        End If
            
         End If
    Next

' só parte Colorida
    ActiveDocument.SaveAs FileName:=nomeDaPasta & strNewFolderName & "\" & "Parte_Colorida", FileFormat:=wdFormatDocument

ErroHandler:
    Err.Clear
    Resume Next
   
End Sub

Function CreateFolder(ByVal sPath As String) As Boolean
'by Patrick Honorez - www.idevlop.com
    Dim fs As Object
    Dim FolderArray
    Dim Folder As String, i As Integer, sShare As String

    If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
        
    Set fs = CreateObject("Scripting.FileSystemObject")
        If sPath Like "\\*\*" Then
            sPath = Replace(sPath, "\", "@", 1, 3)
        End If

    FolderArray = Split(sPath, "\")
    FolderArray(0) = Replace(FolderArray(0), "@", "\", 1, 3)
     On Error GoTo hell
        For i = 0 To UBound(FolderArray) Step 1
            Folder = Folder & FolderArray(i) & "\"
            If Not fs.FolderExists(Folder) Then
                fs.CreateFolder (Folder)
            End If
        Next
    CreateFolder = True

    hell:
    End Function
