Attribute VB_Name = "Mod_HyperlinksToBlue"
'=== 智能设置超链接和域为蓝色（避免不必要的格式更改）===
Sub SetHyperlinksAndFieldsToBlue()
    Dim hyperlink As hyperlink
    Dim field As field
    Dim storyRange As Range
    Dim countChanged As Integer
    
    Application.ScreenUpdating = False
    countChanged = 0
    
    ' 处理超链接
    For Each hyperlink In ActiveDocument.Hyperlinks
        If hyperlink.Range.Font.Color <> RGB(0, 0, 255) Then
            hyperlink.Range.Font.Color = RGB(0, 0, 255)
            countChanged = countChanged + 1
        End If
    Next hyperlink
    
    ' 处理域，但排除题注
    For Each storyRange In ActiveDocument.StoryRanges
        countChanged = countChanged + ProcessFieldsExcludeCaptions(storyRange)
    Next storyRange
    
    Application.ScreenUpdating = True
    
    If countChanged > 0 Then
        MsgBox "已将 " & countChanged & " 个超链接和域设置为蓝色", vbInformation
    Else
        MsgBox "所有超链接和域已经是蓝色", vbInformation
    End If
End Sub

Private Function ProcessFieldsExcludeCaptions(rng As Range) As Integer
    Dim field As field
    Dim count As Integer
    Dim fieldCode As String
    
    count = 0
    For Each field In rng.Fields
        ' 获取域代码并检查是否是题注
        fieldCode = LCase(field.Code.Text)
        
        ' 排除题注相关的域（SEQ域和包含"图"、"表"等题注关键词的域）
        If Not IsCaptionField(fieldCode) Then
            If field.Result.Font.Color <> RGB(0, 0, 255) Then
                field.Result.Font.Color = RGB(0, 0, 255)
                count = count + 1
            End If
        End If
    Next field
    
    ' 处理链接的范围
    Do While Not (rng.NextStoryRange Is Nothing)
        Set rng = rng.NextStoryRange
        For Each field In rng.Fields
            fieldCode = LCase(field.Code.Text)
            If Not IsCaptionField(fieldCode) Then
                If field.Result.Font.Color <> RGB(0, 0, 255) Then
                    field.Result.Font.Color = RGB(0, 0, 255)
                    count = count + 1
                End If
            End If
        Next field
    Loop
    
    ProcessFieldsExcludeCaptions = count
End Function

' 判断是否为题注域的函数
Private Function IsCaptionField(fieldCode As String) As Boolean
    ' 常见的题注域标识
    Dim captionIndicators As Variant
    captionIndicators = Array("seq", "图", "表", "chart", "figure", "table", "caption")
    
    Dim indicator As Variant
    For Each indicator In captionIndicators
        If InStr(1, fieldCode, indicator, vbTextCompare) > 0 Then
            IsCaptionField = True
            Exit Function
        End If
    Next indicator
    
    IsCaptionField = False
End Function

Public Sub SetHyperlinksAndFieldsToBlueRibbon(control As IRibbonControl)
    SetHyperlinksAndFieldsToBlue
End Sub
