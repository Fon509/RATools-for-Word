Attribute VB_Name = "Mod_BatchRenameFiles"
Sub BatchRenameFiles()
    ' ==============================================================================
    ' 功能：批量修改文件名
    '       1. 仅保留汉字、字母、数字、中划线和下划线
    '       2. 空格将被直接删除，其他非法字符替换为中划线 "-"
    '       3. 支持“文件夹模式”（包含所有子文件夹）和“多文件选择模式”
    '       4. 如果文件被占用无法重命名，自动创建改名后的副本
    ' ==============================================================================

    Dim fDialog As FileDialog
    Dim mode As VbMsgBoxResult
    Dim fileList As Collection
    Dim vFile As Variant
    Dim targetFolder As String
    Dim fullPath As String
    Dim fileName As String, newFileName As String
    Dim baseName As String, extName As String
    Dim cleanName As String
    Dim regEx As Object
    Dim fso As Object
    Dim count As Integer
    Dim copyCount As Integer
    Dim i As Integer
    Dim newPath As String
    Dim dupCounter As Integer
    
    ' 1. 询问用户模式
    mode = MsgBox("请选择操作模式：" & vbCrLf & vbCrLf & _
                  "【是 (Yes)】 选择一个文件夹 (递归处理所有子文件夹)" & vbCrLf & _
                  "【否 (No)】  选择具体文件 (支持按住Ctrl或Shift多选)", _
                  vbYesNoCancel + vbQuestion, "选择模式")
    
    If mode = vbCancel Then Exit Sub
    
    Set fileList = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. 根据模式获取文件列表
    If mode = vbYes Then
        ' --- 文件夹模式 (递归) ---
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
        fDialog.Title = "请选择包含待处理文件的文件夹"
        
        If fDialog.Show = -1 Then
            targetFolder = fDialog.SelectedItems(1)
            ' 调用递归过程获取所有子文件夹中的文件
            RecursiveGetFiles fso.GetFolder(targetFolder), fileList
        Else
            Exit Sub
        End If
    Else
        ' --- 文件多选模式 ---
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        fDialog.Title = "请选择需要清洗文件名的文件（可多选）"
        fDialog.AllowMultiSelect = True
        fDialog.Filters.Clear
        fDialog.Filters.Add "所有文件", "*.*"
        
        If fDialog.Show = -1 Then
            For i = 1 To fDialog.SelectedItems.count
                fileList.Add fDialog.SelectedItems(i)
            Next i
        Else
            Exit Sub
        End If
    End If
    
    If fileList.count = 0 Then
        MsgBox "没有找到可处理的文件。", vbExclamation
        Exit Sub
    End If
    
    ' 3. 初始化正则
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .IgnoreCase = True
        ' 允许范围：a-z, A-Z, 0-9, -, _, 汉字
        .Pattern = "[^a-zA-Z0-9\-\_" & ChrW(&H4E00) & "-" & ChrW(&H9FA5) & "]"
    End With
    
    count = 0
    copyCount = 0
    Application.ScreenUpdating = False
    
    ' 4. 统一循环处理
    For Each vFile In fileList
        fullPath = CStr(vFile)
        
        ' 获取当前文件所在的文件夹路径（这对子文件夹中的文件很重要）
        targetFolder = fso.GetParentFolderName(fullPath) & "\"
        fileName = fso.GetFileName(fullPath)
        baseName = fso.GetBaseName(fullPath)
        extName = "." & fso.GetExtensionName(fullPath)
        If extName = "." Then extName = ""
        
        ' 清洗逻辑：先去空格，再替特殊字符
        baseName = Replace(baseName, " ", "")
        cleanName = regEx.Replace(baseName, "-")
        
        If Len(cleanName) = 0 Then cleanName = "RenamedFile"
        
        newFileName = cleanName & extName
        
        If fileName <> newFileName Then
            newPath = targetFolder & newFileName
            
            ' 防重名
            dupCounter = 1
            Do While fso.FileExists(newPath)
                newFileName = cleanName & "_" & dupCounter & extName
                newPath = targetFolder & newFileName
                dupCounter = dupCounter + 1
            Loop
            
            ' 尝试重命名或复制
            On Error Resume Next
            Err.Clear
            Name fullPath As newPath ' 尝试直接重命名
            
            If Err.Number = 0 Then
                ' 重命名成功
                count = count + 1
            Else
                ' 重命名失败（通常因为文件被打开），尝试创建副本
                Err.Clear
                fso.CopyFile fullPath, newPath
                If Err.Number = 0 Then
                    copyCount = copyCount + 1
                End If
            End If
            On Error GoTo 0
        End If
        
    Next vFile
    
    Application.ScreenUpdating = True
    
    ' 5. 结果提示
    MsgBox "处理完成！" & vbCrLf & _
           "直接重命名: " & count & " 个" & vbCrLf & _
           "创建副本(原文件被占用): " & copyCount & " 个", _
           vbInformation, "批量修改文件名"
    
    Set regEx = Nothing
    Set fso = Nothing
    Set fDialog = Nothing
    Set fileList = Nothing

End Sub

' ==========================================
' 辅助过程：递归获取文件夹及子文件夹下的所有文件
' ==========================================
Private Sub RecursiveGetFiles(ByVal oFolder As Object, ByRef colFiles As Collection)
    Dim oFile As Object
    Dim oSubFolder As Object
    
    On Error Resume Next ' 防止遇到无权限文件夹导致中断
    
    ' 1. 添加当前文件夹的文件
    For Each oFile In oFolder.Files
        ' 排除临时文件
        If Left(oFile.Name, 2) <> "~$" Then
            colFiles.Add oFile.Path
        End If
    Next oFile
    
    ' 2. 递归处理子文件夹
    For Each oSubFolder In oFolder.SubFolders
        RecursiveGetFiles oSubFolder, colFiles
    Next oSubFolder
    
    On Error GoTo 0
End Sub

