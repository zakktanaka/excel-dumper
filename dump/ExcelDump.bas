Attribute VB_Name = "ExcelDump"
' Microsoft Visual Basic for Application Extensibilly 5.3

Private outputDir_ As String

Public Function ExcelDump_Dump(book As Workbook)
    
    Dim outputDir As String
    outputDir = SelectOutputDir_
    
    If outputDir <> "" Then
        
        DumpModules_ book, outputDir
        
    End If
    
End Function


Private Function SelectOutputDir_() As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .Title = "ダンプファイルの出力先フォルダを選択してください。"
        
        If outputDir_ <> "" Then .InitialFileName = outputDir_
        
        If .Show Then
            outputDir_ = .SelectedItems(1)
            SelectOutputDir_ = outputDir_
        End If
        
    End With
    
End Function

Private Function DumpModules_(book As Workbook, outputDir As String)
    
    Dim mdl As VBComponent
    Dim modules As VBComponents
    Set modules = book.VBProject.VBComponents
    
    For Each mdl In modules
    
        mdl.Export (outputDir & "\" & mdl.Name & "." & ResolveExtention_(mdl))
    
    Next mdl
    
    Set modules = Nothing
    
End Function

Private Function ResolveExtention_(mdl As VBComponent) As String
    
    If mdl.Type = vbext_ct_ClassModule Then
        ResolveExtention = "cls"
    ElseIf mdl.Type = vbext_ct_StdModule Or mdl.Type = vbext_ct_Document Then
        ResolveExtention_ = "bas"
    ElseIf mdl.Type = vbext_ct_MSForm Then
        ResolveExtention_ = "frm"
    Else
        Err.Raise Number:=999, Description:="unsupported component type"
    End If
    
End Function
