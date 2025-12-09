Attribute VB_Name = "Mod_ConvertWordToPDF"
Sub BatchConvertWordToPDF()
    Dim targetType As Integer
    Dim folderPath As String
    Dim selectedFiles As Variant
    Dim i As Integer
    Dim doc As Document
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 让用户选择转换类型
    targetType = MsgBox("请选择处理模式：" & vbCrLf & vbCrLf & _
                      "【是】- 转换文件夹（包括所有子文件夹）" & vbCrLf & _
                      "【否】- 转换单个或多个Word文件" & vbCrLf & _
                      "【取消】- 退出宏" & vbCrLf & _
                      "停止运行按Ctrl+Pause，笔记本可尝试Ctrl+Fn+B", _
                      vbYesNoCancel + vbQuestion, "批量Word转PDF")
    
    If targetType = vbCancel Then Exit Sub
    
    ' 根据选择调用不同的处理函数
    If targetType = vbYes Then
        ' 选择文件夹模式（包含子文件夹）
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "请选择包含Word文档的文件夹（将处理所有子文件夹）"
            If .Show <> -1 Then Exit Sub
            folderPath = .SelectedItems(1)
        End With
        
        If folderPath = "" Then Exit Sub
        
        ' 递归处理文件夹及其所有子文件夹
        ProcessFolderWithSubfolders folderPath
    Else
        ' 选择多个文件模式
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "请选择一个或多个Word文档"
            .Filters.Clear
            .Filters.Add "Word文档", "*.doc;*.docx"
            .AllowMultiSelect = True
            If .Show <> -1 Then Exit Sub
            
            ' 获取选中的所有文件
            ReDim selectedFiles(1 To .SelectedItems.count)
            For i = 1 To .SelectedItems.count
                selectedFiles(i) = .SelectedItems(i)
            Next i
        End With
        
        ' 处理选中的文件
        ProcessSelectedFiles selectedFiles
    End If
    
    Set fso = Nothing
    MsgBox "批量转换完成！", vbInformation
End Sub

' 递归处理文件夹及其所有子文件夹
Sub ProcessFolderWithSubfolders(folderPath As String)
    Dim fso As Object
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim filePath As String
    Dim pdfFileName As String
    Dim doc As Document
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fso.GetFolder(folderPath)
    
    ' 处理当前文件夹中的Word文件
    For Each file In mainFolder.Files
        filePath = file.Path
        If IsWordDocument(filePath) Then
            ' 转换当前文件
            Set doc = Documents.Open(filePath)
            pdfFileName = fso.BuildPath(fso.GetParentFolderName(filePath), _
                           fso.GetBaseName(filePath) & ".pdf")
            
            doc.ExportAsFixedFormat _
                OutputFileName:=pdfFileName, _
                ExportFormat:=wdExportFormatPDF, _
                CreateBookmarks:=wdExportCreateHeadingBookmarks
            
            doc.Close False
        End If
    Next
    
    ' 递归处理所有子文件夹
    For Each subFolder In mainFolder.subFolders
        ProcessFolderWithSubfolders subFolder.Path
    Next
    
    Set fso = Nothing
    Set mainFolder = Nothing
End Sub

' 处理选中的多个文件
Sub ProcessSelectedFiles(filePaths As Variant)
    Dim i As Integer
    Dim filePath As String
    Dim pdfFileName As String
    Dim doc As Document
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For i = LBound(filePaths) To UBound(filePaths)
        filePath = filePaths(i)
        If IsWordDocument(filePath) Then
            Set doc = Documents.Open(filePath)
            pdfFileName = fso.BuildPath(fso.GetParentFolderName(filePath), _
                           fso.GetBaseName(filePath) & ".pdf")
            
            doc.ExportAsFixedFormat _
                OutputFileName:=pdfFileName, _
                ExportFormat:=wdExportFormatPDF, _
                CreateBookmarks:=wdExportCreateHeadingBookmarks
            
            doc.Close False
        End If
    Next i
    
    Set fso = Nothing
End Sub

' 检查文件是否为Word文档
Function IsWordDocument(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ext As String
    ext = LCase(fso.GetExtensionName(filePath))
    
    ' 检查是否是Word文档且不是临时文件
    If (ext = "doc" Or ext = "docx") And Left(fso.GetFileName(filePath), 2) <> "~$" Then
        IsWordDocument = True
    Else
        IsWordDocument = False
    End If
    
    Set fso = Nothing
End Function

