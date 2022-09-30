Option Explicit

' プロパティ一覧
Public masterId As Integer
Public AgrId As Integer
Public CoId As Integer
Public AgrName As String
Public CoName As String

' 初期値の設定
Public Sub initialize(ByVal myRange As Range)
    masterId = CInt(myRange(eMasterId).value)
    AgrId = CInt(myRange(eAgrId).value)
    CoId = CInt(myRange(eCoId).value)
    AgrName = myRange(eAgrName).value
    CoName = myRange(eCoName).value
    '引数の定数はシートモジュールに列挙型で記載
End Sub
