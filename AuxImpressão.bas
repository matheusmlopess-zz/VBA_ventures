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
    
' #####################################################
' Handler de texto
' #####################################################

    'Text color Automatic to White
     ActiveDocument.Range.Font.Color = wdColorWhite
    
    'Exclui bordas do documento
     ActiveDocument.Sections(1).Borders.Enable = False

     For Each forma In ActiveDocument.Shapes
     forma.Select
        If Selection.ShapeRange.TextFrame.HasText Then
        On Error GoTo Handler1
            Selection.Font.Color = wdColorWhite
        End If
     Next
    
' #####################################################
' Naveando pelas shapes
' #####################################################
'Para documentos com mais shapes

     'If forma.Type = msoTextBox Then
         'forma.Select
         'forma.PictureFormat.Brightness = 0
         'forma.Line.Visible = False
         
       '  Selection.Borders.Enable = False
       '  Selection.ShapeRange.Fill.BackColor = wdColorWhite
        ' Selection.ShapeRange.Fill.ForeColor = wdColorWhite
        ' Selection.ShapeRange.TextFrame.TextRange.Select
         
      
        ' aux = Selection.Text    ' aux = Left(aux, 20) para 20 caracteres
        ' achei = MsgBox("[" + Selection.ShapeRange(1).Name + "]:" _
             & vbCrLf _
             & vbCrLf _
             & aux _
             & vbCrLf _
             & "Parar aqui?", _
               vbYesNo, "Caixa de texto encontrada")
             
         'If achei = vbYes Then Exit For
    
     'End If
    
' #####################################################
' só parte colorida
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
    & strNewFolderName & "\" & "Parte_Colorid", _
    FileFormat:=wdFormatDocument

 '#####################################################

 MsgBox "Revertenduuu ..."
 
    ActiveDocument.Range.Font.Color = wdColorAuto
    ActiveDocument.Sections(1).Borders.Enable = True
    
        For Each forma In ActiveDocument.Shapes
        forma.Select

            If Selection.ShapeRange.TextFrame.HasText Then
            On Error GoTo Handler1
                 Selection.Font.Color = wdColorAuto
            End If
        Next

' #####################################################
' Handler de wordArt
' #####################################################

   
     On Error GoTo ErroHandler
       For Each eh_WordArt In ActiveDocument.Shapes
            sTexto = "no Word Art"
            sTexto = eh_WordArt.TextEffect.Text
                If sText <> "no Word Art" Then
                    eh_WordArt.Select
                    eh_WordArt.Fill.Visible = False
                    eh_WordArt.Fill.Transparency = 1
                    eh_WordArt.Line.Visible = False
 
                End If
        
    Next
        
    For Each eh_WordArt_Inline In ActiveDocument.InlineShapes
           sTexto = "no Word Art"
           sTexto = eh_WordArt_Inline.TextEffect.Text
                If sText <> "no Word Art" Then
                   eh_WordArt_InlineS.Select
                    
                   eh_WordArt_Inline.Fill.Visible = False
                   eh_WordArt_Inline.Fill.Transparency = 1
                   eh_WordArt_Inline.Line.Visible = False
 
                 End If
        
    Next
' #####################################################
' Handler de imagens
' #####################################################
 
 Dim forma2 As Shape
 
 For Each forma2 In ActiveDocument.Shapes
 
      forma2.Select
    ' MsgBox Selection.ShapeRange.Name
     
        If forma2.Type = msoPicture Then
            forma2.Select
            forma2.PictureFormat.Brightness = 1
        End If
        
    ' MsgBox Selection.ShapeRange(1).Name
        
 Next
 
 
     
' #####################################################
' só parte preto e branca
' #####################################################
 
    ActiveDocument.SaveAs FileName:=nomeDaPasta _
    & strNewFolderName & "\" & "Parte_PretoBranco", _
    FileFormat:=wdFormatDocument

 '########################################################################


ErroHandler:
    Err.Clear
    Resume Next
 
 
Handler1:
    Err.Clear
    Resume Next
   
Handler2:
    MsgBox "Erro inner loop codigo (2) !"
    Resume Next
   
End Sub



Function CreateFolder(ByVal sPath As String) As Boolean
'by Patrick Honorez - www.idevlop.com
'create full sPath at once, if required
'returns False if folder does not exist and could NOT be created, True otherwise
'sample usage: If CreateFolder("C:\toto\test\test") Then debug.print "OK"
'updated 20130422 to handle UNC paths correctly ("\\MyServer\MyShare\MyFolder")

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



