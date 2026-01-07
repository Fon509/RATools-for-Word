Attribute VB_Name = "Mod_BatchConvertWordToPDF"
Option Explicit

' 定义全局变量用于记录日志
Dim processLog As String
Dim successCount As Integer
Dim failCount As Integer

Sub BatchConvertWordToPDF()
    Dim modeInput As String
    Dim folderPath As String
    Dim selectedFiles As Variant
    Dim i As Integer
    Dim reportDoc As Document
    Dim viewReport As Integer
    
    ' 初始化日志变量
    processLog = "【批量转PDF处理报告】" & vbCrLf & "时间：" & Now & vbCrLf & String(50, "-") & vbCrLf
    successCount = 0
    failCount = 0
    
    ' 优化性能：关闭屏幕更新和警告
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    On Error GoTo ErrorHandler

    ' 让用户输入模式编号
    modeInput = InputBox("请输入模式编号：" & vbCrLf & vbCrLf & _
                         "1 - 【当前文档】处理当前打开的文档" & vbCrLf & _
                         "2 - 【文件模式】选择单个或多个文件" & vbCrLf & _
                         "3 - 【文件夹模式】递归处理文件夹", _
                         "批量Word转PDF", "1")
                         
    If modeInput = "" Then GoTo Cleanup ' 用户点击取消或未输入

    Select Case modeInput
        Case "1"
            ' --- 模式1：当前文档 ---
            If Documents.count > 0 Then
                ConvertActiveDocument
            Else
                MsgBox "当前没有打开的文档！", vbExclamation
                GoTo Cleanup
            End If
            
        Case "2"
            ' --- 模式2：多文件 ---
            With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "请选择一个或多个Word文档"
                .Filters.Clear
                .Filters.Add "Word文档", "*.doc;*.docx;*.docm"
                .AllowMultiSelect = True
                If .Show <> -1 Then GoTo Cleanup
                
                For i = 1 To .SelectedItems.count
                    ConvertOneFile .SelectedItems(i)
                Next i
            End With
            
        Case "3"
            ' --- 模式3：文件夹 ---
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "请选择包含Word文档的文件夹"
                If .Show <> -1 Then GoTo Cleanup
                folderPath = .SelectedItems(1)
            End With
            
            If folderPath <> "" Then
                ProcessFolderWithSubfolders folderPath
            End If
            
        Case Else
            MsgBox "输入无效，请输入 1、2 或 3。", vbExclamation
            GoTo Cleanup
    End Select
    
    ' --- 恢复屏幕更新 ---
    Application.ScreenUpdating = True
    
    ' --- 结果反馈 ---
    If modeInput = "1" Then
        ' 模式1：直接提示完成，不询问报告，符合“直接关闭对话框”的需求
        ' 只有当有成功或失败计数时才弹窗。如果因为未保存文档退出(success=0, fail=0)，则不重复弹窗。
        If successCount > 0 Or failCount > 0 Then
            MsgBox "当前文档处理完成！" & vbCrLf & _
                   IIf(failCount > 0, "注意：转换失败。", "转换成功，PDF已保存在同级目录下。"), vbInformation
        End If
    Else
        ' 模式2和3：询问是否查看报告
        viewReport = MsgBox("处理完成！" & vbCrLf & _
                            "成功: " & successCount & " 个" & vbCrLf & _
                            "失败: " & failCount & " 个" & vbCrLf & vbCrLf & _
                            "是否生成并查看详细处理报告？", vbYesNo + vbQuestion, "批量转换完成")
        
        If viewReport = vbYes Then
            Set reportDoc = Documents.Add
            With reportDoc.Content
                .Text = processLog & vbCrLf & String(50, "=") & vbCrLf & _
                        "处理完成！" & vbCrLf & _
                        "成功：" & successCount & " 个" & vbCrLf & _
                        "失败：" & failCount & " 个"
                .Font.Name = "微软雅黑"
                .Font.Size = 10
            End With
        End If
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Exit Sub

ErrorHandler:
    MsgBox "发生意外错误: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' 新增：专门处理当前活动文档的函数
Sub ConvertActiveDocument()
    Dim doc As Document
    Dim fso As Object
    Dim pdfFileName As String
    Dim toc As TableOfContents
    Dim tof As TableOfFigures
    
    Set doc = ActiveDocument
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ActiveDocError
    
    ' 检查文档是否已保存（需要路径来保存PDF）
    If doc.Path = "" Then
        MsgBox "请先保存当前文档，以便确定PDF输出位置。", vbExclamation
        Exit Sub
    End If
    
    ' 刷新目录
    If doc.TablesOfContents.count > 0 Then
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
    End If
    
    ' 刷新图表目录
    If doc.TablesOfFigures.count > 0 Then
        For Each tof In doc.TablesOfFigures
            tof.Update
        Next tof
    End If
    
    ' 构建PDF路径
    pdfFileName = fso.BuildPath(doc.Path, fso.GetBaseName(doc.Name) & ".pdf")
    
    ' 导出PDF
    doc.ExportAsFixedFormat _
        OutputFileName:=pdfFileName, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True
        
    successCount = successCount + 1
    processLog = processLog & "[成功] " & doc.Name & " (当前文档)" & vbCrLf
    
    Set fso = Nothing
    Exit Sub

ActiveDocError:
    failCount = failCount + 1
    processLog = processLog & "[失败] " & doc.Name & " - 原因: " & Err.Description & vbCrLf
    MsgBox "转换失败：" & Err.Description, vbCritical
    Set fso = Nothing
End Sub

' 递归处理文件夹及其所有子文件夹
Sub ProcessFolderWithSubfolders(folderPath As String)
    Dim fso As Object
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fso.GetFolder(folderPath)
    
    ' 处理当前文件夹中的Word文件
    For Each file In mainFolder.Files
        If IsWordDocument(file.Path) Then
            ConvertOneFile file.Path
        End If
    Next
    
    ' 递归处理所有子文件夹
    For Each subFolder In mainFolder.SubFolders
        ProcessFolderWithSubfolders subFolder.Path
    Next
    
    Set fso = Nothing
    Set mainFolder = Nothing
End Sub

' 【核心功能】处理单个文件：打开 -> 刷新目录 -> 导出 -> 关闭
Sub ConvertOneFile(filePath As String)
    Dim doc As Document
    Dim fso As Object
    Dim pdfFileName As String
    Dim toc As TableOfContents
    Dim tof As TableOfFigures
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 错误捕捉：防止单个文件损坏导致整个宏停止
    On Error GoTo FileError
    
    ' 以不可见方式打开，避免屏幕闪烁，ReadOnly防止修改原文件属性
    Set doc = Documents.Open(fileName:=filePath, Visible:=False, ReadOnly:=True, AddToRecentFiles:=False)
    
    ' --- 刷新目录核心代码 ---
    ' 1. 刷新普通目录
    If doc.TablesOfContents.count > 0 Then
        For Each toc In doc.TablesOfContents
            toc.Update ' 更新整个目录（页码和标题）
        Next toc
    End If
    
    ' 2. 刷新图表目录
    If doc.TablesOfFigures.count > 0 Then
        For Each tof In doc.TablesOfFigures
            tof.Update
        Next tof
    End If
    ' -----------------------
    
    ' 构建PDF路径
    pdfFileName = fso.BuildPath(fso.GetParentFolderName(filePath), _
                               fso.GetBaseName(filePath) & ".pdf")
    
    ' 导出PDF
    doc.ExportAsFixedFormat _
        OutputFileName:=pdfFileName, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True
    
    ' 关闭文档（不保存对原Word的更改，因为我们只是为了生成PDF才更新了目录）
    doc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' 记录成功日志
    successCount = successCount + 1
    processLog = processLog & "[成功] " & fso.GetFileName(filePath) & vbCrLf
    
    GoTo Finally

FileError:
    ' 记录失败日志
    failCount = failCount + 1
    processLog = processLog & "[失败] " & fso.GetFileName(filePath) & " - 原因: " & Err.Description & vbCrLf
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges

Finally:
    Set doc = Nothing
    Set fso = Nothing
End Sub

' 检查文件是否为Word文档
Function IsWordDocument(filePath As String) As Boolean
    Dim fso As Object
    Dim ext As String
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(filePath))
    fileName = fso.GetFileName(filePath)
    
    ' 排除临时文件(~$开头) 并检查扩展名
    If (ext = "doc" Or ext = "docx" Or ext = "docm") And Left(fileName, 2) <> "~$" Then
        IsWordDocument = True
    Else
        IsWordDocument = False
    End If
    
    Set fso = Nothing
End Function

