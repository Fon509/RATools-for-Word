Attribute VB_Name = "Mod_UpdateChecker"
Option Explicit

Private Const APP_VERSION As String = "v0.6.0"
Private Const RELEASES_API_URL As String = "https://api.github.com/repos/PharmaRA/RATools-for-Word/releases/latest"
Private Const GITHUB_RELEASE_URL_PREFIX As String = "https://github.com/PharmaRA/RATools-for-Word/releases/tag/"
Private Const GITEE_RELEASE_URL_PREFIX As String = "https://gitee.com/PharmaRA/RATools-for-Word/releases/tag/"

Public Sub CheckForUpdatesManually()
    CheckForUpdatesCore
End Sub

Private Sub CheckForUpdatesCore()
    Dim responseText As String
    Dim latestVersion As String
    Dim githubReleaseUrl As String
    Dim giteeReleaseUrl As String
    Dim compareResult As Long
    Dim userChoice As VbMsgBoxResult

    On Error GoTo ErrHandler

    If Not GetLatestReleaseJson(responseText) Then GoTo FetchFailed

    latestVersion = ExtractJsonStringValue(responseText, "tag_name")
    githubReleaseUrl = ExtractJsonStringValue(responseText, "html_url")

    If latestVersion = "" Or githubReleaseUrl = "" Then GoTo FetchFailed

    giteeReleaseUrl = BuildReleaseUrl(GITEE_RELEASE_URL_PREFIX, latestVersion)

    compareResult = CompareVersions(APP_VERSION, latestVersion)

    If compareResult <= 0 Then
        MsgBox "当前已是最新版本。", vbInformation, "检查更新"
        Exit Sub
    End If

    userChoice = MsgBox("发现新版本：" & latestVersion & vbCrLf & _
                        "当前版本：" & APP_VERSION & vbCrLf & vbCrLf & _
                        "是：打开 GitHub Release" & vbCrLf & _
                        "否：打开 Gitee Release" & vbCrLf & _
                        "取消：关闭提示", _
                        vbYesNoCancel + vbInformation, "发现新版本")

    Select Case userChoice
        Case vbYes
            OpenReleasePage githubReleaseUrl
        Case vbNo
            OpenReleasePage giteeReleaseUrl
    End Select
    Exit Sub

FetchFailed:
    MsgBox "检查更新失败，请稍后重试。", vbExclamation, "检查更新"
    Exit Sub

ErrHandler:
    MsgBox "检查更新失败，请稍后重试。", vbExclamation, "检查更新"
End Sub

Private Function GetLatestReleaseJson(ByRef responseText As String) As Boolean
    Dim http As Object

    On Error GoTo ErrHandler

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", RELEASES_API_URL, False
    http.setRequestHeader "User-Agent", "RATools-for-Word"
    http.send

    If http.Status = 200 Then
        responseText = CStr(http.responseText)
        GetLatestReleaseJson = True
    End If

    Exit Function

ErrHandler:
    GetLatestReleaseJson = False
End Function

Private Function ExtractJsonStringValue(ByVal jsonText As String, ByVal keyName As String) As String
    Dim keyToken As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim startQuote As Long
    Dim endQuote As Long

    keyToken = Chr$(34) & keyName & Chr$(34)
    keyPos = InStr(1, jsonText, keyToken, vbTextCompare)
    If keyPos = 0 Then Exit Function

    colonPos = InStr(keyPos + Len(keyToken), jsonText, ":")
    If colonPos = 0 Then Exit Function

    startQuote = InStr(colonPos + 1, jsonText, Chr$(34))
    If startQuote = 0 Then Exit Function

    endQuote = FindJsonStringEnd(jsonText, startQuote + 1)
    If endQuote = 0 Then Exit Function

    ExtractJsonStringValue = Mid$(jsonText, startQuote + 1, endQuote - startQuote - 1)
End Function

Private Function FindJsonStringEnd(ByVal jsonText As String, ByVal startPos As Long) As Long
    Dim i As Long

    For i = startPos To Len(jsonText)
        If Mid$(jsonText, i, 1) = Chr$(34) Then
            If i = 1 Or Mid$(jsonText, i - 1, 1) <> "\" Then
                FindJsonStringEnd = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NormalizeVersion(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If Len(versionText) > 0 Then
        If Left$(versionText, 1) = "v" Or Left$(versionText, 1) = "V" Then
            versionText = Mid$(versionText, 2)
        End If
    End If
    NormalizeVersion = versionText
End Function

Private Function GetVersionPart(ByVal versionText As String, ByVal index As Long) As Long
    Dim normalizedVersion As String
    Dim parts() As String

    normalizedVersion = NormalizeVersion(versionText)
    If Len(normalizedVersion) = 0 Then Exit Function

    parts = Split(normalizedVersion, ".")
    If index <= UBound(parts) Then
        If IsNumeric(parts(index)) Then
            GetVersionPart = CLng(parts(index))
        End If
    End If
End Function

Private Function CompareVersions(ByVal currentVersion As String, ByVal latestVersion As String) As Long
    Dim i As Long
    Dim currentPart As Long
    Dim latestPart As Long

    For i = 0 To 2
        currentPart = GetVersionPart(currentVersion, i)
        latestPart = GetVersionPart(latestVersion, i)

        If latestPart > currentPart Then
            CompareVersions = 1
            Exit Function
        ElseIf latestPart < currentPart Then
            CompareVersions = -1
            Exit Function
        End If
    Next i
End Function

Private Sub OpenReleasePage(ByVal releaseUrl As String)
    If Len(Trim$(releaseUrl)) = 0 Then Exit Sub
    ActiveDocument.FollowHyperlink Address:=releaseUrl, NewWindow:=True
End Sub

Private Function BuildReleaseUrl(ByVal releaseUrlPrefix As String, ByVal versionText As String) As String
    BuildReleaseUrl = releaseUrlPrefix & versionText
End Function


