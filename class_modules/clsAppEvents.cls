VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 定义一个带事件的 Application 对象
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

' 初始化类：绑定 Word 应用程序对象
Private Sub Class_Initialize()
    Set App = Word.Application
End Sub

' ==========================================
' 场景1：新建文档时 -> 激活选项卡
' ==========================================
Private Sub App_NewDocument(ByVal Doc As Document)
    ' 新建文档时，调用 mRibbon 中的方法激活选项卡
    mRibbon.ActivateRAToolsTab
End Sub

' ==========================================
' 场景2：打开文档时 -> 激活选项卡
' ==========================================
Private Sub App_DocumentOpen(ByVal Doc As Document)
    ' 打开现有文档时，同样激活选项卡
    mRibbon.ActivateRAToolsTab
End Sub
