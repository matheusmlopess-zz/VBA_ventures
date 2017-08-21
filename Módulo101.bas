Attribute VB_Name = "Módulo1"
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
   StartUndoSaver
   
    Dim forma2 As Shape
    For Each forma2 In ActiveDocument.Shapes
       forma2.Select
           If forma2.Type = msoPicture Then
               forma2.Select
                If Selection.ShapeRange.Name <> "Imagem 3" Then
                 forma2.PictureFormat.Brightness = 1
                End If
                If Selection.ShapeRange.Name = "Imagem 3" Then
                    forma2.PictureFormat.Brightness = 0.5
                End If
            End If
            
    Next
    
  EndUndoSaver

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
    MsgBox "okay"
' #####################################################
' Handler de cabeçalho
' #####################################################
     
     For Each hdr In ActiveDocument.Sections(1).Headers
        hdr.Range.Text = vbNullString
     Next hdr
     
' #####################################################
' Handler de imagens
' #####################################################

    Dim forma3 As Shape
    
    For Each forma3 In ActiveDocument.Shapes
       forma3.Select
           If forma3.Type = msoPicture Then
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
 
    ActiveDocument.SaveAs FileName:=nomeDaPasta & strNewFolderName & "\" & "Parte_Colorida", FileFormat:=wdFormatDocument



ErroHandler:
    Err.Clear
    Resume Next
  
Handler1:
    Err.Clear
    Resume Next
   
Handler2:
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






Sub StartUndoSaver()
    On Error Resume Next
    ActiveDocument.Bookmarks.Add "_InMacro_"
    On Error GoTo 0
End Sub


Sub EndUndoSaver()
    On Error Resume Next
    ActiveDocument.Bookmarks("_InMacro_").Delete
    On Error GoTo 0
End Sub


Sub EditUndo() ' Catches Ctrl-Z
    If ActiveDocument.Undo = False Then Exit Sub
    While BookMarkExists("_InMacro_")
        If ActiveDocument.Undo = False Then Exit Sub
    Wend
End Sub


Sub EditRedo() ' Catches Ctrl-Y
    If ActiveDocument.Redo = False Then Exit Sub
    While BookMarkExists("_InMacro_")
        If ActiveDocument.Redo = False Then Exit Sub
    Wend
End Sub


Private Function BookMarkExists(Name As String) As Boolean
    On Error Resume Next
    BookMarkExists = Len(ActiveDocument.Bookmarks(Name).Name) > -1
    On Error GoTo 0
End Function
